#- * - coding: utf - 8 -*-

'''
    author:Kevin
    date:2019/1/24
'''

from ctypes import *
from PIL import Image,ImageEnhance
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from ctypes import *
import pytesseract
import os
import sys
import time
import re


if __name__ == '__main__':

    #从上到下依次为：起止日期、截止日期和作者单位
    publishdate_from = "2018-1"
    publishdate_to = "2018-3"
    author_school = "同济大学"

    nextPage = "//*[@id='ctl00']/table/tbody/tr[3]/td/table/tbody/tr/td/div/a[9]"
    checkCodeImg = "//*[@id='CheckCodeImg']"
    #保存截图
    screen_img = "screen_img.png"

    #注意！！！请使用chrome浏览器！！
    #下面这一行是chrome的绝对路径
    driver = webdriver.Chrome('D:\\chromeDirver\\chromedriver.exe')
    driver.get('http://kns.cnki.net/kns/brief/result.aspx?dbprefix=scdb')
    driver.maximize_window()
    assert "知网" in driver.title

    #找到输入框“作者单位”
    elem_authorSchool = driver.find_element_by_name("au_1_value2")

    #模拟键盘输入
    elem_authorSchool.send_keys(author_school)

    #年份
    elem_year_from = driver.find_element_by_name("publishdate_from")
    elem_year_to = driver.find_element_by_name("publishdate_to")
    elem_year_from.send_keys(publishdate_from)
    elem_year_to.send_keys(publishdate_to)

    #返回页面
    elem_authorSchool.send_keys(Keys.RETURN)
    time.sleep(3)

    #判断元素是否存在
    def is_element_exist(xpath):
        try:
            driver.find_element_by_xpath(xpath)
            return True
        except:
            return False

    def is_element_exist_s(selector):
        try:
            driver.find_element_by_css_selector(selector)
            return True
        except:
            return False

    #点击全选
    def check_all():
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((
            By.XPATH, "//*[@id='selectCheckbox']")))
        elem_checkAll = driver.find_element_by_xpath("//*[@id='selectCheckbox']")
        elem_checkAll.click()

    #点击下一页
    def next_page():
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((
            By.XPATH, "//*[@id='ctl00']/table/tbody/tr[3]/td/table/tbody/tr/td/div/a[9]")))
        elem_next = driver.find_element_by_xpath(
            "//*[@id='ctl00']/table/tbody/tr[3]/td/table/tbody/tr/td/div/a[9]")
        elem_next.send_keys(Keys.RIGHT)

    #截取验证码
    def cut_pic():
        scroll_add_crowd_button = driver.find_element_by_xpath("//*[@id='CheckCode']")
        driver.execute_script("arguments[0].scrollIntoView();", scroll_add_crowd_button)
        driver.get_screenshot_as_file(screen_img)
        img = Image.open(screen_img)
        img = img.convert('RGBA')  # 转换模式：L | RGB
        img = img.convert('L')  # 转换模式：L | RGB
        img = ImageEnhance.Contrast(img)  # 增强对比度
        img = img.enhance(2.0)  # 增加饱和度
        crop_size = (1395,40,1490,78)
        cropped = img.crop(crop_size)
        cropped.save(screen_img)
        return screen_img

    #识别验证码
    #调用api
    #使用前请修改username和password
    def ydm_func():
        print('>>>正在初始化...')
        YDMApi = windll.LoadLibrary('yundamaAPI-x64')
        appId = 6763
        appKey = b'a061cb9728ddfec149de4f956d536c7a'
        #print('软件ＩＤ：%d\r\n软件密钥：%s' % (appId, appKey))
        username = b'a553407655'
        password = b'yxcpink123'
        if username == b'test':
            exit('\r\n>>>请先设置用户名密码')
        ####################### 一键识别函数 YDM_EasyDecodeByPath #######################
        print('\r\n>>>正在一键识别...')
        codetype = 1005
        result = c_char_p(b"                              ")
        timeout = 30
        filename = b'screen_img.png'
        captchaId = YDMApi.YDM_EasyDecodeByPath(username, password, appId, appKey, filename, codetype, timeout, result)
        print("一键识别：验证码ID：%d，识别结果：%s" % (captchaId, result.value))
        code = str(result.value)
        b = ''
        for i in code.strip('b'):
            pattern = re.compile(r'[a-zA-Z0-9]')
            m = pattern.search(i)
            if m != None:
                b += i
        print(b)
        return b

    #模拟鼠标点击
	
    #切换frame
    driver.switch_to.frame(driver.find_element_by_xpath("//*[@id='iframeResult']"))
    # 每页50条
    driver.find_element_by_xpath("//*[@id='id_grid_display_num']/a[3]").click()
    WebDriverWait(driver,10).until(EC.element_to_be_clickable((By.XPATH,"//*[@id='id_grid_display_num']/a[3]")))
    driver.find_element_by_xpath("//*[@id='id_grid_display_num']/a[3]").click()
    #第一次全选

    check_all()

    count = 0
    #当下一页存在时
    while is_element_exist(nextPage):
        next_page()

        while is_element_exist(checkCodeImg):
            time.sleep(2)
            cut_pic()
            driver.find_element_by_id("CheckCode").send_keys(ydm_func())
            time.sleep(1)
            driver.find_element_by_xpath("/html/body/p[1]/input[2]").click()
        check_all()
        #解决alert
        result = EC.alert_is_present()(driver)
        #print(result)
        time.sleep(1)
        if result:
            alert = driver.switch_to.alert
            alert.accept()
            driver.switch_to.default_content()

            # 导出文件
            #WebDriverWait(driver,10).until(EC.element_to_be_clickable((By.NAME,"btnoutput")))
            elem_export = driver.find_element_by_class_name("btnoutput")
            elem_export.click()
            time.sleep(5)

            # 获取当前窗口的句柄
            for handle in driver.window_handles:
                driver.switch_to_window(handle)
            elem_selfDefine = driver.find_element_by_xpath("//*[@id='SaveTypeList']/li[10]/span[1]/a")
            elem_selfDefine.click()
            # WebDriverWait(driver, 10).until(
            # EC.element_to_be_clickable((By.CSS_SELECTOR, "#selfDefine > table > tbody > tr:nth-child(4) > td > input[type='button']:nth-child(1)")))
            time.sleep(3)
            # wait = WebDriverWait(driver,10)
            # wait.until((lambda driver:driver.find_element("<input type=button' value='全选' onclick='checkAllFields(this.form)'>")))
            all_1= "#selfDefine > table > tbody > tr:nth-child(4) > td > input[type='button']:nth-child(1)"
            all_2 = "selfDefine > table > tbody > tr:nth-child(3) > td > input[type='button']:nth-child(1)"
            if is_element_exist_s(all_1):
                elem_all_1 = driver.find_element_by_css_selector(all_1)
                elem_all_1.click()
            # selfDefine > table > tbody > tr:nth-child(3) > td > input[type="button"]:nth-child(1)
            if is_element_exist_s(all_2):
                elem_all_1 = driver.find_element_by_css_selector(all_2)
                elem_all_1.click()
            elem_download = driver.find_element_by_xpath("//*[@id='exportExcel']")
            elem_download.click()

            #跳转主页面
            windows = driver.window_handles
            driver.switch_to.window(windows[0])
            driver.switch_to.frame(driver.find_element_by_xpath("//*[@id='iframeResult']"))
            driver.find_element_by_xpath("//*[@id='J_ORDER']/tbody/tr[2]/td/table/tbody/tr/td[1]/div/span/a").click()

    driver.switch_to.default_content()
