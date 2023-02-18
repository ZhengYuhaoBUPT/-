import xlwt
import time
from selenium import webdriver
# 设置driver
chrome_driver='C:/Program Files (x86)/Google/Chrome/Application/chromedriver.exe'
driver = webdriver.Chrome(chrome_driver)

# 让服务器识别不出是selenium机器人
driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument",{
  "source": """
    Object.defineProperty(navigator, 'webdriver', {
      get: () => undefined
    })
  """
})

# 输入关键词
def InputFunc():
    myinput = input("请输入您想查询的关键词：")
    # 得到网页并搜索关键词
    driver.get('https://kns.cnki.net/kns8/defaultresult/index')
    time.sleep(1)
    driver.find_element_by_id("txt_search").send_keys(myinput)   # .send_keys()  向搜索框中输入内容
    driver.find_element_by_class_name("search-btn").click()       # .click() 代表点击动作
    time.sleep(1)
    return myinput

# 制作工作簿
def MakeWb():
    wb = xlwt.Workbook(encoding='utf-8',style_compression=0)  #一个实例
    sheet = wb.add_sheet('论文信息',cell_overwrite_ok=True) #工作簿名称
    return wb, sheet

# 选择对应页面
def ChooseThePage(myinput,sheet):
    # 获得查询类型
    getwhat = input("请输入您想查询的类型（论文/专利/标准/白皮书）：")
    # 根据查询类型调整页面
    if getwhat == '专利':
        datalist, collen = ZhuanLi(sheet)
    elif getwhat == '论文':
        datalist, collen = LunWen(sheet)
    elif getwhat == '标准':
        datalist, collen = BiaoZhun(sheet)
    elif getwhat == '白皮书':
        datalist, collen = WhiteBook(sheet)
    else:
        datalist = [[]]
        collen = 0
        print("不在搜索范围内！")
    return datalist, collen, getwhat

# 白皮书
def WhiteBook(sheet):
    driver.find_element_by_id("txt_search").send_keys('白皮书')
    driver.find_element_by_class_name("search-btn").click()  # .click() 代表点击动作
    time.sleep(1)
    return LunWen(sheet)

# 标准
def BiaoZhun(sheet):
    pages = 5
    col = ('标题', '类型', '发表时间', '摘要', '链接')
    for i in range(0, 5):
        sheet.write(0, i, col[i])
    datalist = [[]]
    driver.find_element_by_xpath("//span[contains(text(),'标准')]").click()
    time.sleep(1)
    try:
        while True:
            trs = driver.find_elements_by_xpath("//table[@class='result-table-list']/tbody/tr")
            for tr in trs:
                # 题目
                title = tr.find_element_by_xpath("./td[@class='name']").text
                print(title)
                # 论文链接
                href = tr.find_element_by_xpath("./td[@class='name']/a").get_attribute('href')
                print(href)
                # 类型
                type = tr.find_element_by_xpath("./td[@class='standard-source']").text
                print(type)
                # 发表时间
                pubtime = tr.find_element_by_xpath("./td[@class='date']").text
                print(pubtime)
                # 摘要
                driver.execute_script("window.open('%s', '_blank');"%href)
                handles = driver.window_handles
                driver.switch_to.window(handles[-1])
                time.sleep(1)
                try:
                    abstract = driver.find_element_by_class_name("abstract-text").text
                except:
                    abstract = title
                print(abstract)
                driver.close()
                driver.switch_to.window(handles[0])
                datalist.append([title, type, pubtime, abstract, href])
            # 下一页
            if pages > 0:
                pages -= 1
            else:
                break
            try:
                next = driver.find_element_by_id("PageNext").click()
                time.sleep(2)
            except:
                print("没有下一页了")
                break
    except:
        print("已经中断！")
    return datalist, len(col)

# 论文
def LunWen(sheet):
    pages = 5
    col = ('标题', '作者', '发表时间', '摘要', '来源', '下载', '链接')
    for i in range(0, 7):
        sheet.write(0, i, col[i])
    datalist = [[]]
    try:
        while True:
            trs = driver.find_elements_by_xpath("//table[@class='result-table-list']/tbody/tr")
            for tr in trs:
                # 题目
                title = tr.find_element_by_xpath("./td[@class='name']").text
                print(title)
                # 论文链接
                href = tr.find_element_by_xpath("./td[@class='name']/a").get_attribute('href')
                print(href)
                # 作者
                try:
                    author = tr.find_element_by_xpath("./td[@class='author']").text
                    print(author)
                except:
                    author = ''
                    print("没有作者")
                # 来源
                source = tr.find_element_by_xpath("./td[@class='source']").text
                print(source)
                # 下载
                # 发表时间
                download = tr.find_element_by_xpath("./td[@class='download']").text
                pubtime = tr.find_element_by_xpath("./td[@class='date']").text
                print(pubtime)
                print(download)
                # 摘要
                driver.execute_script("window.open('%s', '_blank');"%href)
                handles = driver.window_handles
                driver.switch_to.window(handles[-1])
                time.sleep(1)
                try:
                    abstract = driver.find_element_by_class_name("abstract-text").text
                except:
                    abstract = title
                print(abstract)
                driver.close()
                driver.switch_to.window(handles[0])
                # 写入
                datalist.append([title, author, pubtime, abstract, source, download, href])
            if pages > 0:
                pages -= 1
            else:
                break
            # 下一页
            try:
                next = driver.find_element_by_id("PageNext").click()
                time.sleep(2)
            except:
                print("没有下一页了")
                break
    except:
        print("已经中断！")
    return datalist, len(col)

# 专利
def ZhuanLi(sheet):
    pages = 5
    col = ('标题', '作者', '发表时间', '摘要', '来源', '下载', '链接')
    for i in range(0, 7):
        sheet.write(0, i, col[i])
    datalist = [[]]
    driver.find_element_by_xpath("//span[contains(text(),'专利')]").click()
    time.sleep(1)
    try:
        while True:
            trs = driver.find_elements_by_xpath("//table[@class='result-table-list']/tbody/tr")
            for tr in trs:
                # 题目
                title = tr.find_element_by_xpath("./td[@class='name']").text
                print(title)
                # 论文链接
                href = tr.find_element_by_xpath("./td[@class='name']/a").get_attribute('href')
                print(href)
                # 作者
                try:
                    author = tr.find_element_by_xpath("./td[@class='inventor']").text
                    print(author)
                except:
                    author = ''
                    print("没有作者")
                # 来源
                source = tr.find_element_by_xpath("./td[@class='applicant']").text
                print(source)
                # 发表时间
                pubtime = tr.find_element_by_xpath("//*[@id='gridTable']/table/tbody/tr[1]/td[7]").text
                print(pubtime)
                # 摘要
                driver.execute_script("window.open('%s', '_blank');"%href)
                handles = driver.window_handles
                driver.switch_to.window(handles[-1])
                time.sleep(1)
                try:
                    abstract = driver.find_element_by_class_name("abstract-text").text
                except:
                    abstract = title
                print(abstract)
                driver.close()
                driver.switch_to.window(handles[0])

                download =''
                # 写入
                datalist.append([title, author, pubtime, abstract, source, download, href])
            if pages > 0:
                pages -= 1
            else:
                break
            # 下一页
            try:
                next = driver.find_element_by_id("PageNext").click()
                time.sleep(2)
            except:
                print("没有下一页了")
                break
    except:
        print("已经中断！")
    return datalist, len(col)

# 填写工作簿
def WriteWb(myinput, type, datalist, wb, sheet, collen):
    datalen = len(datalist)
    for row in range(1, datalen):  # 行
        for col in range(0, collen):  # 列
            sheet.write(row, col, str(datalist[row][col]))
    wb.save('%s-%s.xls' %(type, myinput))  # 最后一定要保存，否则无效

if __name__ == '__main__':
    myinput = InputFunc()
    wb, sheet = MakeWb()
    datalist, collen, type = ChooseThePage(myinput,sheet)
    WriteWb(myinput, type, datalist, wb, sheet, collen)
    driver.quit()