# -*- coding: utf-8 -*-
"""确保测试可以直接 import 项目根目录下的模块(Windows / Linux 通用)。"""

import os
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)
