from pymouse import PyMouse
from pykeyboard import PyKeyboard
from PIL import ImageGrab
import time
import pytesseract
import numpy as np


def click2(m: PyMouse, x: int, y: int, interval: float = 0.1):
    """单击动作实际程序。
    
    Arguments:
        m {PyMouse} -- Mouse对象
        x {int} -- 指定单击x坐标
        y {int} -- 指定单击y坐标
    
    Keyword Arguments:
        interval {float} -- 指定该动作后的等待时间 (default: {0.1})
    """
    m.click(x, y)
    time.sleep(interval)


def fill_line(m: PyMouse, k: PyKeyboard, start_pos, shift_pos, data_list):
    """从指定位置开始按照指定间隔填写数据。
    
    Arguments:
        m {PyMouse} -- Mouse对象
        k {PyKeyboard} -- Keyboard对象
        start_pos {tuple or list} -- 指定起始坐标
        shift_pos {tuple or list} -- 指定每次输入后的移动坐标
        data_list {tuple or list} -- 指定输入的数据
    """
    assert len(start_pos) == 2
    assert len(shift_pos) == 2
    x0, y0 = start_pos
    dx, dy = shift_pos
    for i, data in enumerate(data_list):
        click2(
            m,
            x0 + i * dx,
            y0 + i * dy,
        )
        k.type_string(data)


def solvent_table(csv_filename: str, button_pos, m, *args):
    """用于填写表格
    
    Arguments:
        csv_filename {str} -- 表格文件载入路径
        button_pos {tuple or list} -- 换页按钮位置
        *args应该提供以下参数，否则fill_line()将抛出错误：
        m {PyMouse} -- Mouse对象
        k {PyKeyboard} -- Keyboard对象
        start_pos {tuple or list} -- 指定起始坐标
        shift_pos {tuple or list} -- 指定每次输入后的移动坐标
    """
    assert len(button_pos) == 2
    bpos_x, bpos_y = button_pos
    with open(csv_filename) as f:
        for line in f:
            data_line = line[:-1].split(',')
            fill_line(m, *args, data_line)
            click2(m, bpos_x, bpos_y)


def screenshot_ocr(area):
    """用于截取屏幕中的内容，并OCR之
    
    Arguments:
        area {tuple or list} -- 需要截取的屏幕范围，应该是一个四元组。
    """
    img = ImageGrab.grab(bbox=area)
    #img.show()
    img = np.array(img.getdata(), np.uint8).reshape(img.size[1], img.size[0],
                                                    3)
    ocr_result = pytesseract.image_to_string(img)
    return ocr_result


def drop_down_select(m: PyMouse, pos, bias, order: int):
    """用于在下拉框选中指定内容。
    
    Arguments:
        m {PyMouse} -- Mouse对象
        pos {tuple or list} -- 选择位置
        bias {tuple or list} -- 给定位置每行偏差
        order {int} -- 选择目标的顺位
    """
    assert len(pos) == 2 and len(bias) == 2
    px, py = pos
    bx, by = bias
    click2(m, px, py)
    click2(m, px + bx * order, py + by * order)