# -*- coding: utf-8 -*-

import logging
import os
import pathlib
import random
import re
import shutil
import string
import sys
import textwrap
import time
from datetime import datetime, date, timedelta

import colorlog
import uiautomator2 as u2
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, TwoCellAnchor

logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)

console_handler = logging.StreamHandler()
console_handler.setLevel(logging.DEBUG)

formatter = colorlog.ColoredFormatter(
    fmt='%(log_color)s %(asctime)s %(levelname)s:%(message)s%(reset)s',
    datefmt="%Y-%m-%d %H:%M:%S",
    log_colors={
        'DEBUG': 'cyan',
        'INFO': 'green',
        'WARNING': 'yellow',
        'ERROR': 'red',
        'CRITICAL': 'red,bg_white',
    }
)
console_handler.setFormatter(formatter)
logger.addHandler(console_handler)

d = u2.connect("SNU0220A15007866")
dinfo = d.info
d_displayHeight = dinfo['displayHeight']
d_displayWidth = dinfo['displayWidth']

package_name = "com.taobao.idlefish"
activity_name = ".maincontainer.activity.MainActivity"


class TimeUtil:

    @staticmethod
    def random_sleep(random_start=2, random_end=5):
        wait_time = random.randint(random_start, random_end)
        time.sleep(wait_time)

    @staticmethod
    def sleep(secs):
        time.sleep(secs)

    @staticmethod
    def curr_date():
        return datetime.now().strftime("%Y-%m-%d")

    @staticmethod
    def tomorrow_date():
        today = date.today()
        tomorrow = today + timedelta(days=1)
        return tomorrow.strftime('%Y-%m-%d')


def get_desktop_path():
    if sys.platform == 'win32':
        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    elif sys.platform == 'darwin':
        desktop_path = str(pathlib.Path.home() / "Desktop")
    else:
        desktop_path = None
    return desktop_path


def excel_cell_to_index(cell_str):
    """
    Excel单元格转换为行列索引，例如A1转换为1,1  C4转换为3,4
    :param cell_str:
    :return:
    """
    letter = cell_str[0]
    number = int(cell_str[1:])
    column = ord(letter.upper()) - ord('A') + 1
    row = number
    return row, column


def write_img_by_cell(work_book, sheet_name, cell_str, img_path, target_file):
    """
    向Excel某单元格写入图片
    :param work_book: work_book
    :param sheet_name: sheet页名称
    :param cell_str: 单元格名称，例A1，B2，C3
    :param img_path: 图片全路径
    :return:
    """
    sheet = work_book[sheet_name]
    row, column = excel_cell_to_index(cell_str)
    row = row - 1
    column = column - 1
    img = Image(img_path)
    _from = AnchorMarker(column, 40000, row, 40000)
    to = AnchorMarker(column + 1, -40000, row + 1, -40000)
    img.anchor = TwoCellAnchor('twoCell', _from, to)
    sheet.add_image(img)
    work_book.save(target_file)


def to_excel(data_list):
    dt = TimeUtil.curr_date()
    write_path = os.getcwd()
    if not os.path.exists(write_path):
        os.makedirs(write_path)
    output_file = os.path.join(write_path, f"{dt}结果.xlsx")
    wb = Workbook()
    sheet = wb.active
    sheet_name = 'Sheet1'
    sheet.title = sheet_name
    sheet['A1'] = '标题'
    sheet['B1'] = '价格'
    sheet['C1'] = '图片'
    start_row = 2
    for index, data in enumerate(data_list):
        sheet["A" + str(index + start_row)] = data['title']
        sheet["B" + str(index + start_row)] = data['amount']
        write_img_by_cell(work_book=wb, sheet_name=sheet_name, cell_str='C' + str(index + start_row),
                          img_path=data['img'],
                          target_file=output_file)

    wb.save(filename=output_file)
    return output_file


def swipe(startx, starty, endx, endy):
    d.swipe(startx, starty, endx, endy)


def swipe_up():
    fx = random.randint(200, 600)
    fy = random.randint(d_displayHeight - 500, d_displayHeight - 400)
    tx = random.randint(500, 700)
    ty = random.randint(d_displayHeight - 1200, d_displayHeight - 1000)
    swipe(startx=fx, starty=fy, endx=tx, endy=ty)


def del_temp_file():
    if os.path.exists('images'):
        shutil.rmtree('images')


def open_page_by_keyword(keyword):
    TimeUtil.random_sleep()
    d(resourceId="com.taobao.idlefish:id/title").click()
    d.send_keys(keyword, clear=True)
    d.press('enter')


def generate_random_string(length):
    letters_and_digits = string.ascii_letters + string.digits
    return ''.join(random.choice(letters_and_digits) for i in range(length))


def get_amount(s):
    match = re.search(r'¥(\d+)', s)
    if match:
        amount = match.group(1)
        return amount


def save_image(pil_image):
    if not os.path.exists('images'):
        os.makedirs('images')
    img_path = os.path.join('images', generate_random_string(10) + str(int(time.time())) + ".png")
    pil_image.save(img_path)
    return img_path


def remove_unicode(text):
    special_sequences = '\\xef\\xbf\\xbc'
    text = text.replace('\n', '')
    result_str = ''
    for ch in text:
        if special_sequences not in str(ch.encode()):
            result_str += ch
    return result_str


def get_list_data():
    result = []
    TimeUtil.random_sleep()
    # view_list = d.xpath(
    #     '//android.widget.ScrollView/android.view.View[2]/android.view.View[1]/android.widget.ImageView[1]/android.view.View[1]/android.view.View[1]/*').all()
    view_list = d.xpath(
        '//android.widget.ScrollView//android.view.View').all()
    if len(view_list) > 0:
        for el in view_list:
            item_info = el.info
            el_description = remove_unicode(str(item_info['contentDescription']))
            el_text = str(item_info['text']).replace('\n', '')
            if el_description != "" and el_description != "筛选":
                amount = get_amount(el_description)
                if amount is not None and amount != '':
                    img_path = save_image(el.screenshot())
                    result.append({
                        'title': el_description,
                        'amount': amount,
                        'img': img_path
                    })

    return result


def usage():
    message = """
    ######################################################################################################################
                                                   免责声明                                                               
    此工具仅限于学习研究，用户需自己承担因使用此工具而导致的所有法律和相关责任！作者不承担任何法律责任！                 
    ######################################################################################################################
    """
    logger.error(textwrap.dedent(message))
    while True:
        user_input = input("如果您同意本协议, 请输入Y继续: (y/n) ")
        if user_input.lower() == "y":
            break
        elif user_input.lower() == "n":
            sys.exit(0)


def main_exit():
    d.set_fastinput_ime(False)
    d.app_stop(package_name)


def main(keyword, max_page):
    usage()
    try:
        del_temp_file()
        logger.info(d.info)
        d.app_stop(package_name)
        d.app_start(package_name, activity_name, wait=True)
        outputs = []

        logger.info(f"正在获取【{keyword}】关键字信息...")
        open_page_by_keyword(keyword)
        for i in range(max_page):
            logger.info(f"正在滑动[{i}/{max_page}]...")
            list_data = get_list_data()
            if list_data:
                outputs.extend(list_data)
            swipe_up()

        output_file = to_excel(outputs)
        logger.info(f"运行完成，文件路径{output_file}")
    except Exception as e:
        print(e)
        logger.error("程序运行异常:" + str(e.args[0]))
    finally:
        main_exit()


if __name__ == '__main__':
    keyword = '餐饮券'
    max_page = 5  # 向上滑动次数
    main(keyword=keyword, max_page=max_page)
