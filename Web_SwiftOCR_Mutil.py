from selenium import webdriver
from openpyxl import Workbook
from selenium.webdriver.support.ui import Select,WebDriverWait
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import WebDriverException
import time
import io
import base64
from PIL import Image
import sys
import cv2
import numpy as np
import subprocess
import tempfile
from multiprocessing import Pool
from concurrent.futures import ThreadPoolExecutor


chrome_options = Options()
chrome_options.add_argument("--window-size=1320,1055") # 设置窗口大小

service = Service(executable_path='/Users/bing/chromedriver/chromedriver') # 设置chromedriver路径
driver = webdriver.Chrome(service=service, options=chrome_options) # 创建driver对象

def get_college_names(driver):
    college_code_name = {}
    select = Select(driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_DropDownList_yuanxiao"))
    options = select.options
    for index in range(1, len(options)):
        select = Select(driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_DropDownList_yuanxiao"))
        option = select.options[index]
        code_name_pair = option.text.split(".", 1)
        college_code_name[code_name_pair[0]] = code_name_pair[1]
        print(code_name_pair[0], code_name_pair[1])
    return college_code_name

def preprocess_image(img_pil):
    img_pil = img_pil.convert("RGBA")
    border_width = 10
    white_bg = Image.new("RGBA", (img_pil.width + border_width, img_pil.height + border_width), "white")
    white_bg.paste(img_pil, (border_width, border_width), mask=img_pil.split()[3])
    img_with_white_bg = white_bg.convert("RGB")
    img_with_white_bg = img_with_white_bg.resize((img_with_white_bg.width * 5, img_with_white_bg.height * 5), Image.LANCZOS)
    img_cv = cv2.cvtColor(np.array(img_with_white_bg), cv2.COLOR_RGB2BGR)
    img_blurred = cv2.GaussianBlur(img_cv, (5, 5), 0)
    img_processed = Image.fromarray(cv2.cvtColor(img_blurred, cv2.COLOR_BGR2RGB))
    return img_processed

def call_swiftocr(args):
    img_processed, index = args
    with tempfile.NamedTemporaryFile(suffix=".jpg") as temp_file:
        img_processed.save(temp_file.name, "JPEG")
        result = subprocess.check_output(["SwiftOCR", temp_file.name]).decode('utf-8').strip()
        return (result, index)
    
def process_table(driver, college_names, ws, total_rows_processed):
    table = driver.find_element(By.ID, 'ctl00_ContentPlaceHolder1_GridView1')

    rows = table.find_elements(By.CLASS_NAME, 'GridView_RowStyle')
    total_rows = len(rows)

    img_pil_list = []
    for row_num in range(total_rows):
        table = driver.find_element(By.ID, 'ctl00_ContentPlaceHolder1_GridView1')
        row = table.find_elements(By.CLASS_NAME, 'GridView_RowStyle')[row_num]

        cells = row.find_elements(By.TAG_NAME, 'td')
        for col_num, _ in enumerate(cells):
            cell = row.find_elements(By.TAG_NAME, 'td')[col_num]

            if col_num == 1:
                continue

            img = cell.find_elements(By.TAG_NAME, 'img')
            if img:
                img_src = img[0].get_attribute('src')
                img_data = base64.b64decode(img_src.split('base64,')[1])
                img_pil = Image.open(io.BytesIO(img_data))
                img_pil_list.append(img_pil)
    img_processed_list = [(preprocess_image(img_pil), index) for index, img_pil in enumerate(img_pil_list)]

    with ThreadPoolExecutor(max_workers=30) as executor: # 30个线程
        ocr_results = list(executor.map(call_swiftocr, img_processed_list))

    ocr_results.sort(key=lambda x: x[1])
    text_results = [result[0] for result in ocr_results]

    result_index = 0
    for row_num in range(total_rows):
        table = driver.find_element(By.ID, 'ctl00_ContentPlaceHolder1_GridView1')
        row = table.find_elements(By.CLASS_NAME, 'GridView_RowStyle')[row_num]

        current_row_num = total_rows_processed + row_num + 1
        row_data = []

        cells = row.find_elements(By.TAG_NAME, 'td')
        for col_num, _ in enumerate(cells):
            cell = row.find_elements(By.TAG_NAME, 'td')[col_num]

            if col_num == 1:
                college_code = row_data[0]
                if college_code in college_names:
                    text = college_names[college_code]
                else:
                    text = "未知学院"
            else:
                img = cell.find_elements(By.TAG_NAME, 'img')
                if img:
                    text = text_results[result_index]
                    result_index += 1
                else:
                    text = cell.text
            row_data.append(text)

        for col_num, text in enumerate(row_data):
            ws.cell(row=current_row_num, column=col_num + 1, value=text)

        print(f"第{current_row_num}行数据： {row_data}")

        # 如果需要，在第一行数据处理后等待用户确认
        if total_rows_processed == 0 and row_num == 0:
            confirmation = ""
            while confirmation.lower() != 'y':
                confirmation = input("如果识别无误，请输入 Y 确认后继续: ")
                if confirmation.lower() == 'n':  # 如果用户输入 'n'，则退出程序
                    driver.quit()
                    sys.exit("用户取消，程序已停止")

    return total_rows_processed + total_rows




def main():
    wb = Workbook()
    ws = wb.active
    driver.get('https://www.gxzslm.cn/Main/Xinwen/XW_Jihua.aspx')
    time.sleep(5)

    for batch in range(1, 5):  # 遍历四个批次
        select_element = Select(driver.find_element(By.ID, 'ctl00_ContentPlaceHolder1_DropDownList_pici'))
        select_element.select_by_value(str(batch))  # 选择批次
        time.sleep(5)
        college_names = get_college_names(driver)
        total_rows_processed = 0

        while True:
            try:
                # 等待表格元素出现
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'ctl00_ContentPlaceHolder1_GridView1')))
                total_rows_processed = process_table(driver, college_names, ws, total_rows_processed)
                wb.save('data.xlsx')
                next_button = driver.find_element(By.ID, 'ctl00_ContentPlaceHolder1_GridView1_ctl33_btnNext')
                driver.execute_script("arguments[0].scrollIntoView();", next_button)  # 滚动使按钮可见
                time.sleep(1)  # 等待一下，以确保按钮可见
                next_button.click()
                WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, 'ctl00_ContentPlaceHolder1_GridView1')))
                time.sleep(5)  # 等待一下，以确保页面已加载完成
            except NoSuchElementException:
                break
            except WebDriverException:
                # 重新连接浏览器
                driver.quit()
                driver = webdriver.Chrome(chrome_options=options)
                driver.get('https://www.gxzslm.cn/Main/Xinwen/XW_Jihua.aspx')
                time.sleep(5)

    driver.quit()
    pass



if __name__ == '__main__':
    main()