from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import openpyxl
import time

def init_driver():
    service = Service()
    driver = webdriver.Edge(service=service)
    driver.set_page_load_timeout(30)  # 设置页面加载超时
    return driver

def wait_for_tbody(driver, timeout=10):
    """等待 /html/body/div[2]/table/tbody 元素出现"""
    try:
        WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/table/tbody"))
        )
        return True
    except TimeoutException:
        print("等待 tbody 元素超时")
        return False

def check_rate_limit(driver):
    """检查是否出现 '您的操作过于频繁，请五分钟后再试' 的提示"""
    try:
        element = driver.find_element(By.XPATH, "/html/body/div[2]/table/tbody/tr[2]/td")
        if (element.get_attribute("colspan") == "2" and
                "您的操作过于频繁，请五分钟后再试" in element.text.strip()):
            return True
        return False
    except NoSuchElementException:
        return False

def extract_data(driver):
    """提取 /html/body/div[2]/table/tbody 下的信息"""
    data = {}
    try:
        tbody = driver.find_element(By.XPATH, "/html/body/div[2]/table/tbody")
        rows = tbody.find_elements(By.TAG_NAME, "tr")

        for row in rows:
            cells = row.find_elements(By.TAG_NAME, "td")
            if len(cells) == 2:  # 确保是键值对形式的行
                field_name = cells[0].text.strip().rstrip("：")  # 去掉字段名后的冒号
                field_value = cells[1].text.strip()
                data[field_name] = field_value

                # 检查 "变更前机构信息" 部分的字段（id="td0" 到 id="td3"）
                row_id = row.get_attribute("id")
                if row_id in ["td0", "td1", "td2", "td3"]:
                    data[f"变更前_{field_name}"] = field_value

        return data
    except NoSuchElementException:
        print("未找到 tbody 元素，可能页面为空")
        return {}
    except Exception as e:
        print(f"提取数据时出错: {e}")
        return {}

def main():
    driver = init_driver()
    base_url = "https://xkz.nfra.gov.cn/jr/GNWTIw/showLicenceInfo.do?id="
    start_id = 1
    end_id = 266899

    # 定义字段
    headers = [
        'ID', '机构编码', '机构名称', '英文名称', '机构住所', '机构所在地',
        '邮政编码', '发证日期', '批准日期', '发证机关', '流水号', '业务范围',
        '变更前_流水号', '变更前_机构编码', '变更前_机构名称', '变更前_机构地址'
    ]

    # 创建或加载 XLSX 文件
    try:
        workbook = openpyxl.load_workbook('license_data.xlsx')
    except FileNotFoundError:
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "License Data"
        sheet.append(headers)  # 写入表头
    else:
        sheet = workbook.active

    for id_num in range(start_id, end_id + 1):
        url = f"{base_url}{id_num}"
        retry_count = 0
        max_retries = 3
        rate_limit_retries = 0
        max_rate_limit_retries = 100  # 最多重试100次“操作过于频繁”

        while retry_count < max_retries:
            try:
                driver.get(url)

                # 等待 tbody 出现
                if not wait_for_tbody(driver):
                    print(f"ID {id_num} 未加载到 tbody，记为空数据")
                    sheet.append([id_num] + [''] * (len(headers) - 1))
                    workbook.save('license_data.xlsx')
                    break

                # 检查是否被限制
                if check_rate_limit(driver):
                    rate_limit_retries += 1
                    if rate_limit_retries > max_rate_limit_retries:
                        print(f"ID {id_num} 超过最大频率限制重试次数，跳过")
                        sheet.append([id_num] + [''] * (len(headers) - 1))
                        workbook.save('license_data.xlsx')
                        break
                    print(f"ID {id_num} 操作过于频繁，第 {rate_limit_retries} 次等待 5 分钟")
                    time.sleep(1)  # 等待
                    continue

                # 未被限制，提取数据
                data = extract_data(driver)
                if data:
                    row = [
                        id_num,
                        data.get('机构编码', ''),
                        data.get('机构名称', ''),
                        data.get('英文名称', ''),
                        data.get('机构住所', ''),
                        data.get('机构所在地', ''),
                        data.get('邮政编码', ''),
                        data.get('发证日期', ''),
                        data.get('批准日期', ''),
                        data.get('发证机关', ''),
                        data.get('流水号', ''),
                        data.get('业务范围', ''),
                        data.get('变更前_流水号', ''),
                        data.get('变更前_机构编码', ''),
                        data.get('变更前_机构名称', ''),
                        data.get('变更前_机构地址', '')
                    ]
                    sheet.append(row)
                    workbook.save('license_data.xlsx')
                    print(f"ID {id_num} 已爬取")
                    break
                else:
                    print(f"ID {id_num} 无数据")
                    sheet.append([id_num] + [''] * (len(headers) - 1))
                    workbook.save('license_data.xlsx')
                    break

            except TimeoutException:
                retry_count += 1
                print(f"ID {id_num} 超时，重试 {retry_count}/{max_retries}")
            except WebDriverException as e:
                retry_count += 1
                print(f"ID {id_num} 浏览器错误: {e}，重试 {retry_count}/{max_retries}")
            except Exception as e:
                print(f"ID {id_num} 出错: {e}")
                break

            if retry_count == max_retries:
                print(f"ID {id_num} 超过最大重试次数，跳过")
                sheet.append([id_num] + [''] * (len(headers) - 1))
                workbook.save('license_data.xlsx')

    driver.quit()
    workbook.save('license_data.xlsx')
    print("爬取完成，结果已保存至 license_data.xlsx")

if __name__ == "__main__":
    main()
