# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.

# Press the green button in the gutter to run the script.
# 第一步，导入selenium模块的webdrivier包
from selenium import webdriver
from selenium.webdriver.common.by import By

if __name__ == '__main__':
    # 第二步，调用webdriver包的Chrome类，返回chrome浏览器对象
    driver = webdriver.Chrome('F:\Softwares\Library\chromedriver.exe')
    # 第三步，如使用浏览器一样开始对网站进行访问
    driver.maximize_window()  # 设置窗口最大化

    driver.implicitly_wait(3)  # 设置等待3秒后打开目标网页
    url = "https://www.baidu.com"
    # 使用get方法访问网站
    driver.get(url)
    # 使用find_element_by_id方法定位，这里定位到输入框
    # element_kw = driver.find_element_by_id('kw')
    element_kw_ = driver.find_element(By.ID, "kw")
    # 使用send_keys方法，给输入框传递参数
    element_kw_.send_keys('曹鉴华')
    # 定位到百度一下按钮，模拟点击操作
    # element_btn = driver.find_element_by_id('su')
    driver.find_element(By.ID, "su").click()
    # 获得点击搜索按钮后网页中id名为1的结果
    result = driver.find_element(By.ID, '1')
    # 打印id名为1的文本内容
    print(result.text)
    # 退出浏览器
    driver.quit()
