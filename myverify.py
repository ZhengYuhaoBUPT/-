import os
import time

import pytesseract
import pyautogui
from PIL import Image
from selenium import webdriver
from selenium.webdriver import ActionChains

# chrome_driver = 'C:/Program Files (x86)/Google/Chrome/Application/chromedriver.exe'
# driver = webdriver.Chrome(chrome_driver)

# url = 'https://kns.cnki.net/kcms2/access/verification'
# driver.get(url)
# 保存截图
def image_save_as(driver):
    image = driver.find_element_by_xpath('//img[@id="verifyImg"]')
    actions = ActionChains(driver)
    actions.context_click(image)
    actions.perform()
    pyautogui.typewrite(['down','down','enter','enter'])
    time.sleep(2)
    pyautogui.typewrite(['enter'])
    time.sleep(2)

def get_newest_image(image_path):
    list = os.listdir(image_path)
    list.sort(key = lambda fn:os.path.getmtime(image_path + "\\"+fn))
    image_new = os.path.join(image_path,list[-1])
    return image_new

def get_verify_number(driver):
    image_save_as(driver)
    image = get_newest_image('C:\\Users\\10679\\Downloads')
    # 打开
    image = Image.open(image)
    # 设置灰度图
    lim = image.convert('L')
    lim.save('pice.jpg')
    threshold = 185
    table = []
    for j in range(256):
        if j < threshold:
            table.append(0)
        else:
            table.append(1)
    bim = lim.point(table, '1')
    bim.save('newImg.png')
    # 识别
    lastpic = Image.open('newImg.png')
    text = pytesseract.image_to_string(lastpic, lang='eng', config='--psm 6 --oem 3 -c tessedit_char_whitelist=0123456789')

    with open('output.txt','w') as f:
        f.write(text)
    with open('output.txt','r') as f:
        code = f.read()
    f.close()
    print(code)
    return code