import urllib.request,urllib.error #制定URL，获取数据
import numpy #打酱油
import pandas #打酱油
from bs4 import  BeautifulSoup #网页解析获取数据
import  re  #正则表达式，进行文字匹配
import  os  #打酱油
import  xlwt  #进行excel操作

#步骤
#1.爬取网页 2.解析数据 3.保存数据
def main():
    baseurl="https://movie.douban.com/top250?start="
    #1.爬取网页
    datalist = getdata(baseurl)
    savepath = "豆瓣电影TOP250.xls"
    saveData(datalist,savepath)

#影片连接
findlink = re.compile(r'<a href="(.*?)"') #创造正则表达式,表示规则，提取超链接
#影片图片
findimg = re.compile(r'img.*src="(.*?)"',re.S) #re.S表示包含换行符进行匹配，图片链接
#影片片名
findtitle = re.compile(r'<span class="title">(.*)</span>')
#影片评分
findpoint = re.compile(r'span class="rating_num" property="v:average">(.*)</span>')
#评价人数
findnumber = re.compile(r'<span>(\d*)人评价</span>')
#概括
finddiscribe = re.compile(r'<span class="inq">(.*)</span>')
#找到相关内容
findbd = re.compile(r'<p class="">(.*?)</p>',re.S)
#如果是需要默认保存0次到多次，需要括号加问号


#爬取网页并解析数据
def getdata(baseurl):
    datalist=[]
    for i in range(0,10):  #调用十次
        url = baseurl + str(i*25)
        html = askURL(url)  #保存获取的html

        soup = BeautifulSoup(html,"html.parser")
        for item in soup.find_all("div",class_="item"): #查找符合要求的字符串
            #print(item) #查看item信息
            data = []   #储存一部电影的全部信息
            item = str(item)
            #这一步比较关键，以后找的所有信息，都可以用这个方法去查找
            link = re.findall(findlink,item) #re库通过正则表达式查找指定的字符串
            data.append(link)   #添加链接
            #print(link)
            img = re.findall(findimg,item)[0]
            data.append(img)   #添加图片
            #print(img)
            titles = re.findall(findtitle,item) #此处需要考虑中文和外文名
            if (len(titles) == 2):
                chinatitle = titles[0] #添加中文名
                data.append(chinatitle)
                foreigntitle = titles[1].replace("/","") #去掉无关符号
                data.append((foreigntitle)) #添加外国名字
            else :
                data.append(titles)
                data.append(" ") #为了对齐留个空
            #print(titles)
            points = re.findall(findpoint,item)[0]
            data.append(points)   #添加评分
            #print(points)
            num = re.findall(findnumber,item)[0]
            data.append(num)  #添加评价人数
            #print(num)
            discribe = re.findall(finddiscribe,item) #概述
            if len(discribe) != 0:
                discribe = discribe[0].replace("。","") #去除句号
                data.append(discribe)
            else:
                data.append(" ")
            #print(discribe)
            bd = re.findall(findbd,item)[0]
            bd = re.sub('<br(\s+)?/>(\s+)?'," ",bd) #去除一些不需要的符号
            bd = re.sub('/'," ",bd)
            data.append(bd.strip()) #这个函数是去除空格
            #print(bd)
            datalist.append(data)
            #print(datalist)
    return datalist  #此时datalist里面已经存储了我们需要的一部电影的全部内容啦

#得到一个指定URL的网页内容
def askURL(url):
    head = {
        "User-Agent": "Mozilla / 5.0(Windows NT 10.0;WOW64) "
                      "AppleWebKit / 537.36(KHTML, likeGecko) "
                      "Chrome / 75.0.3770 .100Safari / 537.36"
    } #伪装成浏览器
    request = urllib.request.Request(url , headers = head)
    html=""
    try:
        response = urllib.request.urlopen(request)
        html = response.read().decode("utf-8")
        #print(html)
    except urllib.error.URLError as e:
        if hasattr(e,"code"):
            print(e.code)
        elif hasattr((e,"reason")):
            print(e.reason)

    return html

#保存数据
def saveData(datalist,savepath):
    #保存数据到EXCEL，用xlwt库
    book = xlwt.Workbook(encoding="utf-8",style_compression=0) #创建workbook对象
    sheet = book.add_sheet("豆瓣电影TOP250",cell_overwrite_ok=True) #创建工作表
    col = ('电影详情链接','图片链接','影片中文名','影片外文名','评分','评价人数','概括','相关信息')
    for i in range(8):
        sheet.write(0,i,col[i]) #写入列名
    for i in range(0,250):
        #print("第%d条"%(i+1))
        data = datalist[i]
        for j in range(8):
            sheet.write(i+1,j,data[j]) #数据写入

    book.save(savepath)


if __name__ == "__main__":
    main()
    print("恭喜您爬取完成！")