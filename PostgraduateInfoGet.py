from selenium import webdriver
# 导入By类
from selenium.webdriver.common.by import By
# 导入Select类
from selenium.webdriver.support.ui import Select
from selenium.webdriver.edge.service import Service
# 新建修改execl表格
import openpyxl


def findWindowByUrl(driver, keyWord):
    # 驱动与url的关键字
    for handle in driver.window_handles:
        # 先切换到该窗口
        driver.switch_to.window(handle)
        # 得到该窗口的标题栏字符串，判断是不是我们要操作的那个窗口
        if keyWord in driver.current_url:
            # 如果是，那么这时候WebDriver对象就是对应的该该窗口，正好，跳出循环，
            # print(driver.current_url)
            break


def initExeclSheet(execlSheet):
    # 表格初始化
    execlSheet.title = "大学科目表"
    execlSheet['A1'] = '学校'
    execlSheet['B1'] = '院系所'
    execlSheet['C1'] = '专业'
    execlSheet['D1'] = '研究方向'
    execlSheet['E1'] = '学习方式'
    execlSheet['F1'] = '考试方式'
    execlSheet['G1'] = '指导老师'
    execlSheet['H1'] = '拟招人数'
    execlSheet['I1'] = '政治'
    execlSheet['J1'] = '外语'
    execlSheet['K1'] = '业务课一'
    execlSheet['L1'] = '业务课二'
    execlSheet['M1'] = '备注'


def addExeclSheet(execlSheet, baseData, subjectData, other=""):
    global addExeclSheet_counter
    addExeclSheet_counter = addExeclSheet_counter + 1
    strCounter = str(addExeclSheet_counter)
    execlSheet['A' + strCounter] = baseData[0]
    execlSheet['B' + strCounter] = baseData[2]
    execlSheet['C' + strCounter] = baseData[3]
    execlSheet['D' + strCounter] = baseData[5]
    execlSheet['E' + strCounter] = baseData[4]
    execlSheet['F' + strCounter] = baseData[1]
    execlSheet['G' + strCounter] = baseData[6]
    execlSheet['H' + strCounter] = baseData[7]
    execlSheet['I' + strCounter] = subjectData[0]
    execlSheet['J' + strCounter] = subjectData[1]
    execlSheet['K' + strCounter] = subjectData[2]
    execlSheet['L' + strCounter] = subjectData[3]
    execlSheet['M' + strCounter] = other


if __name__ == '__main__':
    # 确定driver位置
    addExeclSheet_counter = 1
    s = Service("D:\pythonabout\pythonProject\PostgraduateInfoGet\msedgedriver.exe")
    driver = webdriver.Edge(service=s)
    url = "https://yz.chsi.com.cn/zsml/zyfx_search.jsp"
    # 新建表格文件
    execlFile = openpyxl.Workbook()
    # 获取sheet
    execlSheet = execlFile.active
    initExeclSheet(execlSheet)
    execlFile.save('D:\Files\Temp\\test.xlsx')
    # 打开网页
    driver.get(url)
    # 门类选择
    sel_Class = driver.find_element(By.ID, "mldm")
    sel_ClassChange = Select(sel_Class)
    sel_ClassChange.select_by_value("13")

    # 专业领域选择
    sel_Field = driver.find_element(By.ID, "yjxkdm")
    sel_FieldChange = Select(sel_Field)
    sel_FieldChange.select_by_value("1303")

    # 点击查询
    btn_Search = driver.find_element(By.NAME, "button")
    btn_Search.click()
    schoolWindow = driver.current_window_handle  # 保存课程页面句柄，用于后期返回课程页面
    while(1):
        # 获取学校列表窗口的句柄方便返回翻页
        schoolWindow = driver.current_window_handle  # 保存课程页面句柄，用于后期返回课程页面
        # 获取当前页所有学校的链接
        lay_Table = driver.find_element(By.CLASS_NAME, "ch-table")
        linkList = lay_Table.find_elements(By.TAG_NAME, "a")
        for link in linkList:
            print(link.text + "   " + link.get_attribute("href"))
        # 对链接列表的链接进行处理，进入大学界面
        for link in linkList:
            print(link.text + "   " + link.get_attribute("href"))
            url = link.get_attribute("href")
            # 在新标签页中打开网页
            js = "window.open('" + url + "');"
            driver.execute_script(js)
            findWindowByUrl(driver, "dwmc")#定位到新打开的网页
            collegeWindow = driver.current_window_handle  # 保存大学页面
            # 切换到学院列表的table中
            collegeTable = driver.find_element(By.TAG_NAME, "table")
            # 获取学院的查看列表
            collegeLinkList = collegeTable.find_elements(By.TAG_NAME, "a")
            strcollegeLinkList = []
            for collegeLink in collegeLinkList:
                # 筛选出查看链接不要详细链接
                if "java" in collegeLink.get_attribute("href"):
                    continue
                strcollegeLinkList.append(collegeLink.get_attribute("href"))
            print(strcollegeLinkList)
            # 进入学院界面
            for collegeLink in strcollegeLinkList:
                js = "window.open('" + collegeLink + "');"
                driver.execute_script(js)
                # 切换到查看的窗口
                findWindowByUrl(driver, "id")
                # 获取数据并存入execl表格中
                baseData = []
                driver.find_elements(By.CLASS_NAME, "zsml-summary")
                for a in driver.find_elements(By.CLASS_NAME, "zsml-summary"):
                    baseData.append(a.text)
                print(baseData)

                # 获取课程所在table，将table的每一行分别处理
                subjectData = []
                for subjectList in driver.find_elements(By.CLASS_NAME, "zsml-res-items"):
                    if subjectData:
                        print(subjectData)
                        addExeclSheet(execlSheet, baseData, subjectData)
                        execlFile.save('D:\Files\Temp\\test.xlsx')
                        subjectData = []
                    for subject in subjectList.find_elements(By.TAG_NAME, "td"):
                        subjectData.append(subject.text)
                if subjectData:
                    addExeclSheet(execlSheet, baseData, subjectData)
                    execlFile.save('D:\Files\Temp\\test.xlsx')
                    subjectData = []
                driver.close()
                driver.switch_to.window(collegeWindow)
            driver.close()
            driver.switch_to.window(schoolWindow)
        try:
            next_page = driver.find_element(By.CSS_SELECTOR,"[class ='lip lip-last']")
            next_page.click()
        except:
            driver.quit()
            break

