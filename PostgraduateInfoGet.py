from selenium import webdriver
# 导入By类
from selenium.webdriver.common.by import By
# 导入Select类
from selenium.webdriver.support.ui import Select
#新建修改execl表格
import openpyxl

def findWindowByUrl(driver,keyWord):
    # 驱动与url的关键字
    for handle in driver.window_handles:
        # 先切换到该窗口
        driver.switch_to.window(handle)
        # 得到该窗口的标题栏字符串，判断是不是我们要操作的那个窗口
        if keyWord in driver.current_url:
            # 如果是，那么这时候WebDriver对象就是对应的该该窗口，正好，跳出循环，
            #print(driver.current_url)
            break

def initExeclSheet(execlSheet):
    #表格初始化
    execlSheet.title = "大学科目表"
    execlSheet['A1'] = '学校'
    execlSheet['B1'] = '院系所'
    execlSheet['C1'] = '专业'
    execlSheet['D1'] = '研究方向'
    execlSheet['E1'] = '考试方式'
    execlSheet['F1'] = '政治'
    execlSheet['G1'] = '外语'
    execlSheet['H1'] = '业务课一'
    execlSheet['I1'] = '业务课二'
    execlSheet['J1'] = '备注'



def addExeclSheet(execlSheet,school,college,subject,direction,method,politics,foreignLan,businessSec1,businessSec2,other = ""):
    global addExeclSheet_counter
    addExeclSheet_counter = addExeclSheet_counter + 1
    execlSheet['A'+addExeclSheet_counter] = school
    execlSheet['B'+addExeclSheet_counter] = college
    execlSheet['C'+addExeclSheet_counter] = subject
    execlSheet['D'+addExeclSheet_counter] = direction
    execlSheet['E'+addExeclSheet_counter] = method
    execlSheet['F'+addExeclSheet_counter] = politics
    execlSheet['G'+addExeclSheet_counter] = foreignLan
    execlSheet['H'+addExeclSheet_counter] = businessSec1
    execlSheet['I'+addExeclSheet_counter] = businessSec2
    execlSheet['J'+addExeclSheet_counter] = other

if __name__ == '__main__':
    #确定driver位置
    # driver = webdriver.Chrome('F:\Softwares\Library\chromedriver.exe')
    addExeclSheet_counter = 1
    driver = webdriver.Edge('F:\Softwares\Library\msedgedriver.exe')
    # driver.implicitly_wait(3)  # 设置等待3秒后打开目标网页
    url = "https://yz.chsi.com.cn/zsml/zyfx_search.jsp"
    #新建表格文件
    execlFile = openpyxl.Workbook()
    #获取sheet
    execlSheet = execlFile.active
    initExeclSheet(execlSheet)
    execlFile.save("test.xlsx")
    #打开网页
    driver.get(url)
    #门类选择
    sel_Class = driver.find_element(By.ID, "mldm")
    sel_ClassChange = Select(sel_Class)
    sel_ClassChange.select_by_value("02")

    #专业领域选择
    sel_Field = driver.find_element(By.ID, "yjxkdm")
    sel_FieldChange = Select(sel_Field)
    sel_FieldChange.select_by_value("0201")

    #点击查询
    btn_Search = driver.find_element(By.NAME, "button")
    btn_Search.click()

    #获取学校列表窗口的句柄方便返回翻页
    schoolWindow = driver.current_window_handle  # 保存课程页面句柄，用于后期返回课程页面

    #获取所有学校的链接
    lay_Table = driver.find_element(By.CLASS_NAME, "ch-table")
    linkList = lay_Table.find_elements(By.TAG_NAME,"a")
    #对链接列表的链接进行处理
    for link in linkList:
        print(link.text+"   "+link.get_attribute("href"))
        url = link.get_attribute("href")
        # 在新标签页中打开网页
        js = "window.open('"+url+"');"
        driver.execute_script(js)
        findWindowByUrl(driver, "dwmc")
        #切换到学院列表的table中
        collegeTable = driver.find_element(By.TAG_NAME,"table")
        #获取学院的查看列表
        collegeLinkList = collegeTable.find_elements(By.TAG_NAME,"a")
        for collegeLink in collegeLinkList:
            #筛选出查看链接不要详细链接
            if("java" in collegeLink.get_attribute("href")):
                continue
            print(collegeLink.get_attribute("href"))
            #打开查看链接
            driver.get(collegeLink.get_attribute("href"))
            #切换到查看的窗口
            findWindowByUrl(driver, "id")
            #获取数据school,college,subject,direction,method,politics,foreignLan,businessSec1,businessSec2
            data

            #获取课程所在table
            subjectItems = driver.find_elements(By.CLASS_NAME,"zsml-res-items")#每一行
            #将table的每一行分别处理
            for subjectList in subjectItems:

                for subject in subjectList.find_elements(By.TAG_NAME,"td"):
                    print(subject.text)
                    addExeclSheet

            print("\n")
        #driver.quit()


