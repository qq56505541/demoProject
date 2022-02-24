import requests
from bs4 import BeautifulSoup

f = open('book.txt', 'a', encoding="gbk")

def novelAnt(startId):
    startBkey = 741674
    staticUrl = "http://m.yunketui.cn/app/index.php?i=601&c=entry&op=read&tid=55020&do=my_shouji&m=iweite_xiaoshuo&bkey="
    for sid in range(startId, 224):
        # 拼接url
        print("爬取数据id" + str(sid))
        bkey = startBkey + sid
        url = staticUrl + str(bkey) + "&sid=" + str(sid)
        # 请求网站
        strhtml = requests.get(url)
        # 解析网站响应
        soup = BeautifulSoup(strhtml.text, 'lxml')
        # 写入标题
        title = soup.find("h2", id="title").get_text()
        f.writelines(str(title) + "\n")
        content = soup.select("#fuzhi > p")
        # 写入内容
        for p in content:
            # print(str(p.get_text()))
            # 为了阅读方便每一段加一个换行
            try:
                f.write(str(p.get_text()) + "\n")
            except:
                print(p.get_text())

def novelAntSingle(startId):
    startBkey = 741674
    url = "http://m.yunketui.cn/app/index.php?i=601&c=entry&op=read&tid=55020&do=my_shouji&m=iweite_xiaoshuo&bkey="
    # 拼接url
    print("爬取数据id" + str(startId))
    bkey = startBkey + startId
    url = url + str(bkey) + "&sid=" + str(startId)
    # 请求网站
    strhtml = requests.get(url)
    # 解析网站响应
    soup = BeautifulSoup(strhtml.text, 'lxml')
    # 写入标题
    title = soup.find("h2", id="title").get_text()
    f.writelines(str(title) + "\n")
    content = soup.select("#fuzhi > p")
    # sArticle = ""
    # 写入内容
    for p in content:
        # sArticle += p.get_text()
        # 为了阅读方便每一段加一个换行
        f.writelines(str(p.get_text()) + "\n")
    # print(sArticle)

if __name__=='__main__':
    novelAnt(1)