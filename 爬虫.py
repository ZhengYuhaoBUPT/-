import xlwt
import time

import myverify
from selenium import webdriver

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


myinput = input("请输入您想查询的关键词：")

driver.get('https://kns.cnki.net/kns8/defaultresult/index')
time.sleep(1)
driver.find_element_by_id("txt_search").send_keys(myinput)   # .send_keys()  向搜索框中输入内容
driver.find_element_by_class_name("search-btn").click()       # .click() 代表点击动作
time.sleep(1)

wb = xlwt.Workbook(encoding='utf-8',style_compression=0)  #一个实例
sheet = wb.add_sheet('论文信息',cell_overwrite_ok=True) #工作簿名称
col = ('标题','作者','摘要','来源','下载', '链接')
for i in range(0,6):
        sheet.write(0,i,col[i])
datalist = [[]]

page = 5
# html = driver.page_source
# print(html)
try:
    while True:
        stop = 0
        trs = driver.find_elements_by_xpath("//table[@class='result-table-list']/tbody/tr")
        for tr in trs:
            try:
                # 题目
                title = tr.find_element_by_xpath("./td[@class='name']").text
                print(title)
            except:
                # 出现验证
                number = myverify.get_verify_number(driver)
                time.sleep(2)
                driver.find_element_by_id("result").send_keys(number)
                time.sleep(2)
                driver.find_element_by_css_selector('input[type="button"]').click()
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
                print("没有作者")
            # 来源
            source = tr.find_element_by_xpath("./td[@class='source']").text
            print(source)
            # 下载
            download = tr.find_element_by_xpath("./td[@class='download']").text
            print(download)
            # # 发表时间
            # pubtime = tr.find_element_by_xpath("./td[@class='date']").text
            # print(pubtime)
            # # 发表类别--期刊，科技成果，国家标准······
            # data = tr.find_element_by_xpath("./td[@class='data']").text
            # print(data)
            # 论文摘要
            try:
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
            except:
                # 出现验证
                number = myverify.get_verify_number(driver)
                time.sleep(2)
                driver.find_element_by_id("result").send_keys(number)
                time.sleep(2)
                driver.find_element_by_css_selector('input[type="button"]').click()
                time.sleep(2)
                # stop = 1
                break

                # driver.execute_script("window.open('%s', '_blank');"%href)
                # handles = driver.window_handles
                # driver.switch_to.window(handles[-1])
                # time.sleep(1)
                # abstract = driver.find_element_by_id("ChDivSummary").text
                # print(abstract)
                # driver.close()
                # driver.switch_to.window(handles[0])

            datalist.append([title, author, abstract, source, download, href])

        # if stop == 1:
        #     break
        # 只找五页
        if page > 0:
            page-=1
        else:
            break
        # 下一页
        try:
            next = driver.find_element_by_id("PageNext").click()
            time.sleep(2)
            # html = driver.page_source
            # print(html)
        except:
            print("没有下一页了")
            break

except:
    print('error')
len = len(datalist)
for row in range(1,len):  #行
    for col in range(0,6):#列
            sheet.write(row, col, str(datalist[row][col]))
wb.save('%s.xls'%myinput)  #最后一定要保存，否则无效
driver.quit()