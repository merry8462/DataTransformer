#!/usr/bin/env bash
# ===========================================================================
#  DataTransformer Linux / macOS 启动脚本
# ---------------------------------------------------------------------------
#  用法:
#     chmod +x data_transformer.sh     # 首次使用请先赋予执行权限
#     ./data_transformer.sh            # 交互式转换向导(默认)
#     ./data_transformer.sh --demo     # 预填 Demo/Example.txt 的 SSH + PostgreSQL 示例
#     ./data_transformer.sh --gui      # 启动图形界面(需已安装 PySide6)
#     ./data_transformer.sh --selftest # 无界面自检,结果写入 selftest.log
#     ./data_transformer.sh --check    # 只检查运行环境(不进入向导)
#     ./data_transformer.sh --help     # 查看全部参数
#
#  行为说明:
#     * 自动寻找 Python 3.8+(python3 / python3.x / python);
#     * 依赖缺失时可自动在项目目录创建 .venv 并安装 requirements.txt
#       (Debian/Ubuntu 的 pip 受 PEP 668 保护,虚拟环境是最稳妥的方式);
#     * 缺少 pymysql / psycopg2 / paramiko 时只影响对应的数据库或 SSH 功能,
#       Excel / JSON / CSV 互转仍然可用。
# ===========================================================================

set -u

SCRIPT_DIR=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd -P)
cd "$SCRIPT_DIR" || exit 1

VENV_DIR="$SCRIPT_DIR/.venv"
REQUIREMENTS="$SCRIPT_DIR/requirements.txt"
WIZARD="$SCRIPT_DIR/dt_cli.py"
GUI_APP="$SCRIPT_DIR/data_transformer.py"
CORE_PACKAGES="openpyxl pymysql psycopg2-binary pymongo questionary paramiko"

export PYTHONIOENCODING="${PYTHONIOENCODING:-utf-8}"
export PYTHONUTF8=1

if [ -t 1 ] && [ -z "${NO_COLOR:-}" ]; then
    C_RESET=$(printf '\033[0m'); C_OK=$(printf '\033[1;32m')
    C_WARN=$(printf '\033[1;33m'); C_ERR=$(printf '\033[1;31m'); C_INFO=$(printf '\033[1;36m')
else
    C_RESET=""; C_OK=""; C_WARN=""; C_ERR=""; C_INFO=""
fi

info() { printf '%s%s%s\n' "$C_INFO" "$*" "$C_RESET"; }
ok()   { printf '%s%s%s\n' "$C_OK" "$*" "$C_RESET"; }
warn() { printf '%s%s%s\n' "$C_WARN" "$*" "$C_RESET"; }
fail() { printf '%s%s%s\n' "$C_ERR" "$*" "$C_RESET"; }

ask_yes_no() {
    # $1=提示 $2=默认值(y/n)
    if [ ! -t 0 ]; then
        [ "${2:-n}" = "y" ]
        return $?
    fi
    if [ "${2:-n}" = "y" ]; then
        printf '%s [Y/n]: ' "$1"
    else
        printf '%s [y/N]: ' "$1"
    fi
    read -r answer || answer=""
    case "$answer" in
        y|Y|yes|YES|是) return 0 ;;
        n|N|no|NO|否)   return 1 ;;
        *)              [ "${2:-n}" = "y" ] ;;
    esac
}

has_module() {
    # $1=python 可执行文件 $2=模块名
    [ -n "${1:-}" ] || return 1
    "$1" -c "import importlib.util as u, sys; sys.exit(0 if u.find_spec('$2') else 1)" \
        >/dev/null 2>&1
}

find_python() {
    for candidate in "${PYTHON_BIN:-}" python3 python3.13 python3.12 python3.11 \
                     python3.10 python3.9 python3.8 python; do
        [ -n "$candidate" ] || continue
        if command -v "$candidate" >/dev/null 2>&1 &&
           "$candidate" -c 'import sys; raise SystemExit(0 if sys.version_info >= (3, 8) else 1)' \
               >/dev/null 2>&1; then
            command -v "$candidate"
            return 0
        fi
    done
    return 1
}

install_hint() {
    warn "请按发行版安装 Python 3 与 pip,例如:"
    printf '    Debian/Ubuntu : sudo apt update && sudo apt install -y python3 python3-venv python3-pip\n'
    printf '    RHEL/CentOS   : sudo dnf install -y python3 python3-pip\n'
    printf '    Fedora        : sudo dnf install -y python3 python3-pip\n'
    printf '    Arch/Manjaro  : sudo pacman -S --needed python python-pip\n'
    printf '    openSUSE      : sudo zypper install -y python3 python3-pip\n'
}

describe_python() {
    "$1" -c 'import sys; print("Python %d.%d.%d" % sys.version_info[:3])' 2>/dev/null \
        || printf 'Python(未知版本)\n'
}

create_venv() {
    info "正在创建虚拟环境:$VENV_DIR"
    if ! "$1" -m venv "$VENV_DIR" >/dev/null 2>&1; then
        fail "创建虚拟环境失败:$1 -m venv 不可用"
        warn "Debian/Ubuntu 请先安装:sudo apt install -y python3-venv"
        return 1
    fi
    if [ ! -x "$VENV_DIR/bin/python" ]; then
        fail "虚拟环境创建异常,未找到 $VENV_DIR/bin/python"
        return 1
    fi
    info "正在安装依赖(首次约 1~2 分钟)..."
    "$VENV_DIR/bin/python" -m pip install --upgrade pip >/dev/null 2>&1 \
        || warn "pip 升级失败,继续尝试安装依赖"
    install_packages "$VENV_DIR/bin/python" "$CORE_PACKAGES" || return 1
    ok "虚拟环境就绪:$VENV_DIR"
}

install_packages() {
    # $1=python $2...=包名(命令行向导不需要 PySide6,保持精简)
    "$1" -m pip install $2 || {
        fail "依赖安装失败,请检查网络或改用国内镜像:"
        printf '    %s -m pip install %s -i https://pypi.tuna.tsinghua.edu.cn/simple\n' "$1" "$2"
        return 1
    }
}

report_env() {
    printf '\n%s\n' "---------------- 运行环境 ----------------"
    printf '  项目目录 : %s\n' "$SCRIPT_DIR"
    printf '  解释器   : %s (%s)\n' "$PY" "$(describe_python "$PY")"
    for module in openpyxl pymysql psycopg2 pymongo questionary paramiko PySide6; do
        if has_module "$PY" "$module"; then
            printf '  %-9s: %s已安装%s\n' "$module" "$C_OK" "$C_RESET"
        else
            printf '  %-9s: %s未安装%s\n' "$module" "$C_WARN" "$C_RESET"
        fi
    done
    printf '%s\n\n' "------------------------------------------"
}

usage() {
    cat <<'EOF'
用法:
  ./data_transformer.sh            交互式转换向导(questionary 动态选择)
  ./data_transformer.sh --demo     预填 Demo/Example.txt 的 SSH + PostgreSQL 示例参数
  ./data_transformer.sh --plain    强制纯文本问答(不调用 questionary)
  ./data_transformer.sh --gui      启动图形界面(需已安装 PySide6)
  ./data_transformer.sh --selftest 无界面自检(依赖 / 核心算法 / 文件互转 / 数值保真)
  ./data_transformer.sh --check    只检查运行环境,不进入向导
  ./data_transformer.sh --help     显示本帮助

说明:
  首次使用请先赋予执行权限: chmod +x data_transformer.sh
  其余参数会原样传给 dt_cli.py(例如 --version)。
EOF
}

# ---------------------------------------------------------------------------
# 参数解析
# ---------------------------------------------------------------------------
MODE="wizard"
for arg in "$@"; do
    case "$arg" in
        -h|--help)     usage; exit 0 ;;
        --check)       MODE="check" ;;
        --gui)         MODE="gui" ;;
        --wizard)      MODE="wizard" ;;
    esac
done

info "DataTransformer 启动中 ..."

PY=$(find_python) || {
    fail "未找到 Python 3.8 或更高版本。"
    install_hint
    exit 1
}

# 优先复用项目内已建好的虚拟环境
if [ -x "$VENV_DIR/bin/python" ] && has_module "$VENV_DIR/bin/python" openpyxl; then
    PY="$VENV_DIR/bin/python"
fi

# 只检查运行环境:不因缺少依赖而提前退出,缺什么就在这里如实报出来
if [ "$MODE" = "check" ]; then
    report_env
    if has_module "$PY" openpyxl; then
        ok "环境检查完成。运行 ./data_transformer.sh 进入交互式向导。"
    else
        warn "缺少必需依赖 openpyxl:运行 ./data_transformer.sh 会自动创建 .venv 并安装。"
    fi
    exit 0
fi

# 必需依赖:openpyxl(Excel/JSON/CSV 与数据库互转都依赖它)
if ! has_module "$PY" openpyxl; then
    warn "当前解释器缺少必需依赖 openpyxl:$PY"
    if ask_yes_no "是否在项目目录创建虚拟环境 .venv 并自动安装依赖" y; then
        create_venv "$PY" || exit 1
        PY="$VENV_DIR/bin/python"
    else
        fail "缺少 openpyxl,无法运行。手动安装示例:"
        printf '    python3 -m venv .venv && .venv/bin/pip install -r requirements.txt\n'
        printf '    或用系统环境: python3 -m pip install --user openpyxl pymysql psycopg2-binary paramiko\n'
        exit 1
    fi
fi

case "$MODE" in
    gui)
        if ! has_module "$PY" PySide6; then
            warn "未安装 PySide6(图形界面需要,约 100MB+)"
            if ask_yes_no "是否现在安装 PySide6" n; then
                install_packages "$PY" "PySide6" || exit 1
            else
                printf '    安装: %s -m pip install PySide6\n' "$PY"
                printf '    或直接使用命令行向导: ./data_transformer.sh\n'
                exit 1
            fi
        fi
        [ -f "$GUI_APP" ] || { fail "找不到 $GUI_APP"; exit 1; }
        info "启动图形界面 ..."
        exec "$PY" "$GUI_APP"
        ;;
    *)
        [ -f "$WIZARD" ] || { fail "找不到 $WIZARD"; exit 1; }
        exec "$PY" "$WIZARD" "$@"
        ;;
esac
