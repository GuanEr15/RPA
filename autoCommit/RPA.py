import os
import time


import win32con
import win32gui
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.wait import WebDriverWait
import selenium.webdriver.support.expected_conditions as EC
from selenium.webdriver.common.by import By
from msedge.selenium_tools import EdgeOptions


def upload(filepath, browser_type="chrome"):
    if browser_type == "chrome":
        title = "打开"
    elif browser_type == "firefox":
        title = "文件上传"
    elif browser_type == "ie":
        title = "选择要加载的文件"
    else:
        title = ""  # 这里根据其它不同浏览器类型来修改

    # 找元素
    # 一级窗口"#32770","打开"
    dialog = win32gui.FindWindow("#32770", title)
    # 向下传递
    ComboBoxEx32 = win32gui.FindWindowEx(dialog, 0, "ComboBoxEx32", None)  # 二级
    comboBox = win32gui.FindWindowEx(ComboBoxEx32, 0, "ComboBox", None)   # 三级
    # 编辑按钮
    edit = win32gui.FindWindowEx(comboBox, 0, 'Edit', None)  # 四级
    # 打开按钮
    button = win32gui.FindWindowEx(dialog, 0, 'Button', "打开(&O)")  # 二级

    print("本次上传文件路径 : ", filepath)
    # 输入文件的绝对路径，点击“打开”按钮
    win32gui.SendMessage(edit, win32con.WM_SETTEXT, None, filepath)  # 发送文件路径
    win32gui.SendMessage(dialog, win32con.WM_COMMAND, 1, button)  # 点击打开按钮


def readFilePath():
    file = open("resource/resource.yaml", encoding='UTF-8')
    lines = file.readlines()  # 读取全部内容 ，并以列表方式返回
    filePathList = []
    print(lines)
    for line in lines:
        if line.startswith("FilePath"):
            FilePath = line[line.index('"') + 1: line.rindex('"')]
            for root, dirs, file in os.walk(FilePath + "\\证据材料"):
                for direc in dirs:
                    filePath = root + "\\" + direc + "\\证据材料.zip"
                    filePathList.append(filePath)
            print(FilePath)
            print("filepath", filePathList)
        elif line.startswith("username"):
            username = line[line.index('"') + 1: line.rindex('"')]
            print(username)
        elif line.startswith("password"):
            password = line[line.index('"') + 1: line.rindex('"')]
            print(password)
    return FilePath, filePathList, username, password


def login(self,driver):
    # 打开网页
    driver.get("https://qdsn.yuntrial.com/")
    # 打开登录网页
    driver.get("https://qdsn.yuntrial.com/lassen/party/app/login")
    # 切换frame
    time.sleep(2)
    iframe = driver.find_element_by_css_selector('iframe')
    driver.switch_to.frame(iframe)
    # 定位输入框 输入参数
    driver.find_element_by_class_name("phone-number-input").send_keys(self.username)
    driver.find_element_by_class_name("password").send_keys(self.password)
    # 登录
    driver.find_element_by_css_selector(".login-button-container span").click()


def step(driver):
    # 等待网页加载完成
    time.sleep(2)
    # 打开我要起诉网页
    driver.get("https://qdsn.yuntrial.com/lassen/party/app/suit/prosecute/choosecasecause")
    # 是否有起诉状 点击否
    time.sleep(1)
    driver.find_element_by_css_selector(
        "#cdk-overlay-2 > nz-modal > div > div.ant-modal-wrap > div > div > div > "
        "app-ai-prosecute > div > div > div.select-content > div > div:nth-child(2)").click()
    # 点击起诉案由
    driver.find_element_by_css_selector(
        "#formly_10_selectCascaderGroup_caseCause_3 > nz-input-group > nz-cascader > div > div").click()
    # 选择 信网权
    time.sleep(1)
    driver.find_element_by_xpath('//*[@id="cdk-overlay-4"]/div/ul/li[6]').click()
    # 点击 选择法院原因
    time.sleep(1)
    # 点击 选择法院原因
    driver.find_element_by_css_selector("#formly_10_select_reason_8 > div > div").click()
    # 选择原告所在地
    driver.find_element_by_xpath('//*[@id="cdk-overlay-5"]/div/div/ul/li[5]').click()
    # 点击下一步
    driver.find_element_by_css_selector(
        "body > app-root > app-prosecute-layout > div > app-prosecute-choose-case-cause "
        "> div > div.m-btn > button.ant-btn.ant-btn-primary").click()
    # 等待网页加载完成
    time.sleep(2)
    # 点击下一步
    driver.find_element_by_css_selector(
        "body > app-root > app-prosecute-layout > div > app-accuser-edit > div > "
        "div.submit-bar.ng-tns-c24-12.ng-star-inserted > button.ng-tns-c24-12.ant-btn.ant-btn-primary").click()
    # 选择批量起诉
    time.sleep(2)
    driver.find_element_by_css_selector("#cdk-overlay-6 > nz-modal > div > div.ant-modal-wrap "
                                        "> div > div > div > div > div > div:nth-child(2)").click()


def uploadFile(self, driver):
    # 选择产品
    driver.find_element_by_css_selector("#formly_87_select_product_0 > div > div").click()
    time.sleep(1)
    # 选择要素表
    driver.find_element_by_css_selector("#cdk-overlay-9 > div > div > ul > li").click()
    time.sleep(1)
    # 选择模板
    driver.find_element_by_css_selector("#formly_87_select_template_1 > div > div").click()
    time.sleep(1)
    # 选择 新美术产品模板
    driver.find_element_by_css_selector("#cdk-overlay-10 > div > div > ul > li:nth-child(2)").click()
    time.sleep(2)
    # 点击导入起诉文件 上传按钮
    driver.find_element_by_css_selector(
        "#formly_87_uploadButton_importProsecute_3 > div > div.btn-upload > button").click()
    time.sleep(1)
    upload(self.SueInformationFilePath)
    # 点击导入证据附件
    time.sleep(4)
    driver.find_element_by_css_selector(
        "#formly_87_uploadButton_importEvidence_4 > div > div.btn-upload > button").click()
    time.sleep(1)
    # 上传多个证据附件
    for i in range(len(self.EvidenceFilePath)):
        upload(self.EvidenceFilePath[i])
        if i < len(self.EvidenceFilePath) - 1:
            time.sleep(2)
            driver.find_element_by_css_selector(
                "#formly_87_uploadButton_importEvidence_4 > div > div.btn-upload > button").click()
            time.sleep(1)

    # 点击起诉状上传附件
    time.sleep(2)
    driver.find_element_by_css_selector(
        "#formly_87_uploadButton_importPleadings_5 > div > div.btn-upload > button").click()
    time.sleep(1)
    upload(self.PleadingsFilePath)
    # 提交
    time.sleep(1)


def submit(driver):
    # 提交
    time.sleep(1)
    # 显示等待 判断是否已经全部上传完毕
    WebDriverWait(driver, 3600, 2).\
        until_not(EC.presence_of_all_elements_located
                      ((By.CSS_SELECTOR, "#formly_87_uploadButton_importEvidence_4 > div "
                                         "> div.col-upload-file.ng-star-inserted > "
                                         "nz-upload > nz-upload-list > div.ant-upload-list-item"
                                         ".ant-upload-list-item-uploading.ng-trigger.ng-trigger-itemState."
                                         "ng-star-inserted")))
    # 点击提交
    driver.find_element_by_css_selector(
        "body > app-root > app-prosecute-layout > div > app-prosecute-batch > div > "
        "div.submit-bar > button.ant-btn.ant-btn-primary").click()
    time.sleep(15)
    driver.quit()


class RPA:
    def __init__(self, info):
        self.SueInformationFilePath = info[0] + "\\要素表.xlsx"
        self.PleadingsFilePath = info[0] + "\\起诉状.zip"
        self.EvidenceFilePath = info[1]
        self.username = info[2]
        self.password = info[3]


if __name__ == '__main__':
    # 读取上传文件路径
    result = readFilePath()
    # chrome 浏览器
    # chrome_options = Options()
    # chrome_options.add_experimental_option("excludeSwitches", ['enable-automation', 'enable-logging'])
    # driver = webdriver.edge(options=chrome_options)

    # edge 浏览器
    driver = webdriver.Edge()

    # 隐式等待
    driver.implicitly_wait(10)

    # 实例化RPA对象
    rpa = RPA(result)
    # 登录
    login(rpa, driver)
    # 中间步骤
    step(driver)
    # 上传文件
    uploadFile(rpa, driver)

    # 提交
    # submit(driver)



