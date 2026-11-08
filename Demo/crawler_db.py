from bs4 import BeautifulSoup
from requests import get


def open_url(url):
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0.0.0 Safari/537.36"
    }
    res = get(url=url, headers=headers)
    return res


def find_movies(res):
    soup = BeautifulSoup(res.text, features="html.parser")

    movies = []  # 电影名
    targets = soup.find_all(name="div", class_="hd")
    for each in targets:
        movies.append(each.a.span.text)

    ranks = []  # 评分
    targets = soup.find_all(name="span", class_="rating_num")
    for each in targets:
        ranks.append(f"评分：{each.text}")

    messages = []  # 资料
    targets = soup.find_all(name="div", class_="bd")
    for each in targets:
        try:
            messages.append(f"{each.p.text.split("\n")[1].strip()}{each.p.text.split("\n")[2].strip()}")
        except:
            continue

    results = []
    length = len(movies)
    for i in range(length):
        results.append(f"{movies[i]}{ranks[i]}{messages[i]}\n")
    return results


def find_depth(res):
    soup = BeautifulSoup(res.text, features="html.parser")
    depth = soup.find(name="span", class_="next").previous_sibling.previous_sibling.text
    return int(depth)


def main():
    host = "https://movie.douban.com/top250"
    res = open_url(url=host)
    depth = find_depth(res=res)

    result = []
    for i in range(depth):
        url = f"{host}/?start={25 * i}"
        res = open_url(url=url)
        result.extend(find_movies(res=res))

    with open(file="豆瓣TOP250电影.txt", mode="w", encoding="utf-8") as f:
        for each in result:
            f.write(each)


if __name__ == '__main__':
    main()