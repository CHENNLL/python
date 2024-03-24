from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import pandas as pd
import openpyxl
import json
import os
import time
import sys



def get_xlsx():
    print('Reading Excel file')
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    file_path = os.path.join(base_path, 'rule.xlsx')
    if not os.path.exists(file_path):
        # print('Configuration file does not exist. Please rename the configuration file to rule.xlsx and place it in the project root directory.')
        return None
    df = pd.read_excel(file_path)
    return df  # Return the DataFrame directly

def AIMind_Login(driver):
    # 输入用户名和密码
    username = driver.find_element(By.ID, value='username')
    password = driver.find_element(By.ID, value='password')
    username.send_keys('newbest')
    password.send_keys('pass')
    # 点击登录按钮
    login_button = driver.find_element(By.XPATH, value='//*[@id="root"]/div/div/div[2]/form/div[3]/div/div/span/button')
    # time.sleep(0.5)
    login_button.click()

def AIMind_Function(driver, Function):
    # 点击区域配置按钮
    if Function == '区域配置':
        # 点击区域配置按钮
        wait = WebDriverWait(driver, 10)
        area_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//a[@href="#/area"]')))
        area_button.click()
    elif Function == '全局场景':
        # 点击全局场景按钮
        wait = WebDriverWait(driver, 10)
        area_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//a[@href="#/scene"]')))
        area_button.click()
    elif Function == '规则管理':
        # 点击规则管理按钮
        wait = WebDriverWait(driver, 10)
        area_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//a[@href="#/rule"]')))
        area_button.click()
    else:
        # 如果传入的参数不是'区域'或'设备'，可以添加相应的处理逻辑
        pass

def Rule_Configuration(driver, rule_param, name_param, event_drivr_name, addr_param, value_param, action_drivr_name, data_param):
    
    if rule_param == '按值比较':
        # 规则名称
        rule_name = driver.find_element(By.ID, 'name')
        rule_name.send_keys(name_param)
        Event_Drive(driver, event_drivr_name)
        time.sleep(1)
        # addr_param
        rule_control_add = driver.find_element(By.ID, 'conditionAddr')
        rule_control_add.send_keys(addr_param)
        # value_param
        rule_control_val = driver.find_element(By.ID, 'conditionValue')
        rule_control_val.send_keys(value_param)
        Action_Drive(driver, rule_param, action_drivr_name)
        time.sleep(1)
        # data_param
        rule_control_data = driver.find_element(By.ID, 'actionData')
        rule_control_data.send_keys(data_param)    

        save_button = driver.find_element(By.XPATH, '//*[@id="root"]/div/div/div/div[2]/div[2]/div/div/form/div[12]/button[2]')
        # time.sleep(3)
        save_button.click()
    elif rule_param == '数据透传':
        # 规则名称
        rule_name = driver.find_element(By.ID, 'name')
        rule_name.send_keys(name_param)

        # addr_param
        rule_control_add = driver.find_element(By.ID, 'conditionAddr')
        rule_control_add.send_keys(addr_param)

        # data_param
        rule_control_data = driver.find_element(By.XPATH, '//*[@id="root"]/div/div/div/div[2]/div[2]/div/div/form/div[7]/div[2]/div/span/div/div[2]/div[1]/input')
        rule_control_data.send_keys(data_param)    

        save_button = driver.find_element(By.XPATH, '//*[@id="root"]/div/div/div/div[2]/div[2]/div/div/form/div[8]/button[2]')
        # time.sleep(3)
        save_button.click()
    elif rule_param == '直接触发':
        # 规则名称
        rule_name = driver.find_element(By.ID, 'name')
        rule_name.send_keys(name_param)

        # addr_param
        rule_control_add = driver.find_element(By.ID, 'conditionAddr')
        rule_control_add.send_keys(addr_param)
        # value_param

        Action_Drive(driver, rule_param, action_drivr_name)
        time.sleep(1)
        # data_param
        rule_control_data = driver.find_element(By.ID, 'actionData')
        rule_control_data.send_keys(data_param)    

        save_button = driver.find_element(By.XPATH, '//*[@id="root"]/div/div/div/div[2]/div[2]/div/div/form/div[8]/button[2]')
        # time.sleep(3)
        save_button.click()
    else:
        # 如果传入的参数不是'区域'或'设备'，可以添加相应的处理逻辑
        pass

def Rule_Select(driver, Rule_param):
    element_input = WebDriverWait(driver, 30).until(
        EC.element_to_be_clickable((By.XPATH, "//button[contains(., '新建规则')]"))
    )
    element_input.click()
    # 等待元素加载完成
    element_input = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="type"]/div/div'))
        )
    driver.execute_script("arguments[0].click();", element_input)

    # 根据传入的参数执行不同的操作
    if Rule_param == '按值比较':
        # 等待区域菜单可见
        rule_menu = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, "//li[contains(., '按值比较')]"))
        )
        driver.execute_script("arguments[0].click();", rule_menu)
        # area_menu.click()
    elif Rule_param == '数据透传':
        # 等待设备菜单可见
        rule_menu = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, "//li[contains(., '数据透传')]"))
        )
        driver.execute_script("arguments[0].click();", rule_menu)        
        # device_menu.click()
    elif Rule_param == '直接触发':
        # 等待设备菜单可见
        rule_menu = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, "//li[contains(., '直接触发')]"))
        )
        driver.execute_script("arguments[0].click();", rule_menu)        
        # device_menu.click()
    elif Rule_param == '地址比较':
        # 等待设备菜单可见
        rule_menu = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, "//li[contains(., '地址比较')]"))
        )
        driver.execute_script("arguments[0].click();", rule_menu)        
        # device_menu.click()
    else:
        # 如果传入的参数不是'区域'或'设备'，可以添加相应的处理逻辑
        pass

def Action_Drive(driver, rule_param, Action_drivr_name):
        # 等待元素加载完成
    element_input = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="actionDriver"]'))
        )
    driver.execute_script("arguments[0].click();", element_input)
    if rule_param == '按值比较':
        
        # 根据传入的参数执行不同的操作
        if Action_drivr_name == 'KNX':
            # 等待区域菜单可见
            # drivr_menu1 = WebDriverWait(driver, 30).until(
            #     EC.element_to_be_clickable((By.XPATH, "//li[text()='KNX']"))
            # )
            # drivr_menu = drivr_menu1[1]
            # driver.execute_script("arguments[0].click();", drivr_menu)
            # area_menu.click()
            pass
        elif Action_drivr_name == '485_1':
            # 等待设备菜单可见
            drivr_menu1 = element_input.find_elements_by_xpath("//li[text()='485_1']")
            drivr_menu = drivr_menu1[1]
            driver.execute_script("arguments[0].click();", drivr_menu)        
            # device_menu.click()
        elif Action_drivr_name == '485_2':
            # 等待设备菜单可见
            drivr_menu1 = element_input.find_elements_by_xpath("//li[text()='485_2']")
            drivr_menu = drivr_menu1[1]
            driver.execute_script("arguments[0].click();", drivr_menu)        
            # device_menu.click()
        elif Action_drivr_name == '485_3':
            # 等待设备菜单可见
            drivr_menu1 = element_input.find_elements_by_xpath("//li[text()='485_3']")
            drivr_menu = drivr_menu1[1]
            driver.execute_script("arguments[0].click();", drivr_menu)        
            # device_menu.click()
        elif Action_drivr_name == '232_1':
            # 等待设备菜单可见
            drivr_menu1 = element_input.find_elements_by_xpath("//li[text()='232_1']")
            drivr_menu = drivr_menu1[1]
            driver.execute_script("arguments[0].click();", drivr_menu)        
            # device_menu.click()
        elif Action_drivr_name == '232_2':
            # 等待设备菜单可见
            drivr_menu1 = element_input.find_elements_by_xpath("//li[text()='232_2']")
            drivr_menu = drivr_menu1[1]
            driver.execute_script("arguments[0].click();", drivr_menu)        
            # device_menu.click()
        else:
            # 如果传入的参数不是'区域'或'设备'，可以添加相应的处理逻辑
            pass
    elif rule_param == '直接触发':
        
        # 根据传入的参数执行不同的操作
        if Action_drivr_name == 'KNX':
            # 等待区域菜单可见
            # drivr_menu1 = WebDriverWait(driver, 30).until(
            #     EC.element_to_be_clickable((By.XPATH, "//li[text()='KNX']"))
            # )
            # drivr_menu = drivr_menu1[1]
            # driver.execute_script("arguments[0].click();", drivr_menu)
            # area_menu.click()
            pass
        elif Action_drivr_name == '485_1':
            # 等待设备菜单可见
            drivr_menu1 = element_input.find_elements_by_xpath("//li[text()='485_1']")
            drivr_menu = drivr_menu1[0]
            driver.execute_script("arguments[0].click();", drivr_menu)        
            # device_menu.click()
        elif Action_drivr_name == '485_2':
            # 等待设备菜单可见
            drivr_menu1 = element_input.find_elements_by_xpath("//li[text()='485_2']")
            drivr_menu = drivr_menu1[0]
            driver.execute_script("arguments[0].click();", drivr_menu)        
            # device_menu.click()
        elif Action_drivr_name == '485_3':
            # 等待设备菜单可见
            drivr_menu1 = element_input.find_elements_by_xpath("//li[text()='485_3']")
            drivr_menu = drivr_menu1[0]
            driver.execute_script("arguments[0].click();", drivr_menu)        
            # drivr_menu1.click()
        elif Action_drivr_name == '232_1':
            # 等待设备菜单可见
            drivr_menu1 = element_input.find_elements_by_xpath("//li[text()='232_1']")
            drivr_menu = drivr_menu1[0]
            driver.execute_script("arguments[0].click();", drivr_menu)        
            # device_menu.click()
        elif Action_drivr_name == '232_2':
            # 等待设备菜单可见
            drivr_menu1 = element_input.find_elements_by_xpath("//li[text()='232_2']")
            drivr_menu = drivr_menu1[0]
            driver.execute_script("arguments[0].click();", drivr_menu)        
            # device_menu.click()
        else:
            # 如果传入的参数不是'区域'或'设备'，可以添加相应的处理逻辑
            pass
    elif rule_param == '数据透传':
        pass

def Event_Drive(driver, Event_drivr_name):
        # 等待元素加载完成
    element_input = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="eventDriver"]'))
        )
    driver.execute_script("arguments[0].click();", element_input)
    # 根据传入的参数执行不同的操作
    if Event_drivr_name == 'KNX':
        # 等待区域菜单可见
        # drivr_menu1 = WebDriverWait(driver, 30).until(
        #     EC.element_to_be_clickable((By.XPATH, "//li[text()='KNX']"))
        # )
        # drivr_menu = drivr_menu1[0]
        # driver.execute_script("arguments[0].click();", drivr_menu)
        # area_menu.click()
        pass
    elif Event_drivr_name == '485_1':
        # 等待设备菜单可见
        drivr_menu1 = element_input.find_elements_by_xpath("//li[text()='485_1']")
        # drivr_menu = drivr_menu1[0]
        driver.execute_script("arguments[0].click();", drivr_menu1)        
        # device_menu.click()
    elif Event_drivr_name == '485_2':
        # 等待设备菜单可见
        drivr_menu1 = element_input.find_elements_by_xpath("//li[text()='485_2']")
        # drivr_menu = drivr_menu1[0]
        driver.execute_script("arguments[0].click();", drivr_menu1)        
        # device_menu.click()
    elif Event_drivr_name == '485_3':
        # 等待设备菜单可见
        drivr_menu1 = element_input.find_elements_by_xpath("//li[text()='485_3']")
        # drivr_menu = drivr_menu1[0]
        driver.execute_script("arguments[0].click();", drivr_menu1)        
        # device_menu.click()
    elif Event_drivr_name == '232_1':
        # 等待设备菜单可见
        drivr_menu1 = element_input.find_elements_by_xpath("//li[text()='232_1']")
        # drivr_menu = drivr_menu1[0]
        driver.execute_script("arguments[0].click();", drivr_menu1)        
        # device_menu.click()
    elif Event_drivr_name == '232_2':
        # 等待设备菜单可见
        drivr_menu1 = element_input.find_elements_by_xpath("//li[text()='232_2']")
        # drivr_menu = drivr_menu1[0]
        driver.execute_script("arguments[0].click();", drivr_menu1)        
        # device_menu.click()
    else:
        # 如果传入的参数不是'区域'或'设备'，可以添加相应的处理逻辑
        pass

def main():
    host_xlsx = get_xlsx()

    # 获取主机IP地址，默认IP地址：192.168.1.188
    ip_address = input("请输入主机IP地址(默认为192.168.1.188): ") or "192.168.1.188"

    # 拼接URL
    url = f"http://{ip_address}/#/login" 

    # 打印信息
    # print(f"您输入的IP地址是: {ip_address}")
    print(f"将访问的URL是: {url}")
    print('打开Edge浏览器')
    driver = webdriver.Edge(executable_path='msedgedriver.exe')
    # 最大化浏览器窗口
    driver.maximize_window()
    # 访问登录页面
    driver.get(url)
    print('登录AIMind主机后台')
    AIMind_Login(driver)
    AIMind_Function(driver, "规则管理")
    time.sleep(1)
    if host_xlsx is not None:
        xls_data_dict = {}
        for rule_number, row in host_xlsx.iterrows():
            rule_param = row["rule_param"]
            name_param = row["name_param"]
            event_drivr_name = row["event_drivr_name"]
            addr_param = row["addr_param"]
            value_param = row["value_param"]
            action_drivr_name = row["action_drivr_name"]
            data_param = row["data_param"]
            xls_data_dict[str(rule_number)] = {
                "rule_param": rule_param,
                "name_param": name_param,
                "event_drivr_name": event_drivr_name,
                "addr_param": addr_param,
                "value_param": value_param,
                "action_drivr_name": action_drivr_name,
                "data_param": data_param
            }
            # 新建规则
            Rule_Select(driver, rule_param)
            time.sleep(1)
            Rule_Configuration(driver, rule_param, name_param, event_drivr_name, addr_param, value_param, action_drivr_name, data_param)
            time.sleep(1)
    time.sleep(5)

if __name__ == '__main__':
    main()
