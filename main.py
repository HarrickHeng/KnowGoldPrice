import time
import threading
from datetime import datetime
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_agg import FigureCanvasAgg as FigureCanvas
from io import BytesIO
import csv
import tkinter as tk
from PIL import ImageTk, Image as PILImage
import itertools  # 用于生成颜色循环

# 设置Chrome浏览器的选项
chrome_options = Options()
chrome_options.add_argument('--headless')  # 设置为无头模式
chrome_options.add_argument('--disable-gpu')  # 禁用GPU加速，避免出现一些兼容性问题

# 指定中文字体（示例为SimHei）
plt.rcParams['font.sans-serif'] = ['SimHei']

# 多个URL和对应的文件名
urls = {
    '跨1': 'https://search.7881.com/G10-100001-G10P001-G10P001001-0.html',
    '跨2': 'https://search.7881.com/G10-100001-G10P026-G10P026001-0.html',
    '跨3-A': 'https://search.7881.com/G10-100001-G10P003-G10P003001-0.html',
    '跨3-B': 'https://search.7881.com/G10-100001-G10P011-G10P011001-0.html',
    '跨4': 'https://search.7881.com/G10-100001-G10P010-G10P010004-0.html',
    '跨5': 'https://search.7881.com/G10-100001-G10P008-G10P008001-0.html',
    '跨6': 'https://search.7881.com/G10-100001-G10P008-G10P002001-0.html',
    '跨7': 'https://search.7881.com/G10-100001-G10P008-G10P002003-0.html',
}

# 对应的文件名列表，与urls的key值对应
csv_file_names = {key: f'csv/data_{key}.csv' for key in urls}

excel_file_template = 'excel/output_{}.xlsx'

# 初始化数据字典
data = {key: [] for key in urls}

# 设置颜色
colors = {
    '跨1': 'b',
    '跨2': 'g',
    '跨3-A': 'r',
    '跨3-B': 'pink',
    '跨4': 'c',
    '跨5': 'm',
    '跨6': 'y',
    '跨7': 'k',
}

# 初始化Tkinter窗口
root = tk.Tk()
root.title("金价数据来源：7881   问题反馈：github.com/harrickHeng")

# 初始化图表标签
img_label = tk.Label(root)
img_label.pack()

# 创建一个新的图表
fig, ax = plt.subplots(figsize=(12, 6))
ax.set_title('DNF金币价格')
ax.set_xlabel('日期')
ax.set_ylabel('每元购得万金数')
ax.grid(True)
fig.autofmt_xdate(rotation=45)
fig.tight_layout()

canvas = FigureCanvas(fig)
image_stream = BytesIO()
canvas.print_png(image_stream)
image_stream.seek(0)
img = PILImage.open(image_stream)
tk_img = ImageTk.PhotoImage(img)

img_label.config(image=tk_img)
img_label.image = tk_img


def fetch_data_async(key, url, csv_file_name):
    global data

    if len(data) == 0:
        # 读取已有的CSV数据到data中
        with open(csv_file_name, mode='r', encoding='utf-8') as file:
            reader = csv.DictReader(file)
            data[key] = [dict(row) for row in reader]

    # 启动Chrome浏览器
    driver = webdriver.Chrome(options=chrome_options)
    driver.get(url)

    # 等待页面元素加载完成（例如，等待3秒钟）
    wait = WebDriverWait(driver, 3)
    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'div.shop-item-box')))

    # 获取页面内容
    html_content = driver.page_source

    # 关闭浏览器
    driver.quit()

    # 解析页面内容
    soup = BeautifulSoup(html_content, 'html.parser')

    # 提取数据
    items = soup.find('div', class_='shop-item-box').findChild('div', class_='shop-v part-02')
    price1 = items.contents[0].find('span').text
    price2 = items.contents[1].find('em').text
    # 获取当前时间
    now = datetime.now()
    formatted_time = now.strftime("%Y-%m-%d %H:%M")
    data[key].append({'每元购得万金数': price1, '元/万金': price2, '日期': formatted_time})

    # 写入数据到CSV文件
    with open(csv_file_name, mode='w', newline='', encoding='utf-8') as file:
        fieldnames = ['每元购得万金数', '元/万金', '日期']
        writer = csv.DictWriter(file, fieldnames=fieldnames)

        # 写入表头
        writer.writeheader()

        # 写入数据行
        for item in data[key]:
            writer.writerow(item)

    # 将数据存储为DataFrame并导出到Excel文件
    df = pd.DataFrame(data[key])
    excel_file = excel_file_template.format(key)
    df.to_excel(excel_file, index=False)
    print(f'{key} 数据已刷新')


def start_scheduler(interval_sec):
    def job():
        while True:
            threads = []
            for key in csv_file_names.keys():
                thread = threading.Thread(target=fetch_data_async, args=(key, urls[key], csv_file_names[key]))
                threads.append(thread)
                thread.start()

            for thread in threads:
                thread.join()

            update_image()
            time.sleep(interval_sec)

    thread = threading.Thread(target=job)
    thread.daemon = True
    thread.start()


def update_image():
    global img_label

    # 更新图表并显示在Tkinter窗口中
    ax.clear()  # 清除旧图表内容

    for key in csv_file_names.keys():
        if data[key]:
            df = pd.DataFrame(data[key])
            ax.plot(df['日期'], df['每元购得万金数'], marker='o', linestyle='-', color=colors[key], label=key)

    ax.legend()
    ax.grid(True)
    fig.autofmt_xdate(rotation=45)
    fig.tight_layout()

    # 将更新后的图表转换为图像并显示在Tkinter窗口中
    canvas.draw()
    image_stream = BytesIO()
    canvas.print_png(image_stream)
    image_stream.seek(0)
    img = PILImage.open(image_stream)
    tk_img = ImageTk.PhotoImage(img)

    img_label.config(image=tk_img)
    img_label.image = tk_img
    root.update_idletasks()


# 启动定时器，定时更新数据
start_scheduler(interval_sec=3)


# Tkinter窗口关闭处理
def on_close():
    global running
    running = False
    root.destroy()


root.protocol("WM_DELETE_WINDOW", on_close)

# 启动Tkinter主循环
root.mainloop()
