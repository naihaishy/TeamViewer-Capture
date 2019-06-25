# -*- coding:utf-8 -*-
# @Time : 2019/6/22 17:05
# @Author : naihai

"""
自动获取teamviewer  界面截图 并发送到指定邮箱
"""
import datetime
import time

import win32gui
import win32ui
import win32con
from PIL import ImageGrab, Image
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header
import smtplib


hwnd_title = dict()
# 设置显示器的文本缩放比例 设置->系统->显示 缩放与布局
screen_scale = 1  # 1.25


def get_all_hwnd(hwnd, mouse):
    if win32gui.IsWindow(hwnd) and win32gui.IsWindowEnabled(hwnd) and win32gui.IsWindowVisible(hwnd):
        hwnd_title.update({hwnd: win32gui.GetWindowText(hwnd)})


def find_team_viewer():
    win32gui.EnumWindows(get_all_hwnd, 0)

    for h, t in hwnd_title.items():
        if t is not "":
            if str(t).lower() == "teamviewer":
                left, top, right, bottom = win32gui.GetWindowRect(h)
                return left, top, right, bottom


def capture_team_viewer_window_pil(save_path):
    """
    使用PIL截图 但是不支持多块显示器
    如果单块显示器可以显示该函数
    :return:
    """
    # 左上角坐标和右下角坐标
    left, top, right, bottom = find_team_viewer()
    full_box = (left * screen_scale, top * screen_scale, right * screen_scale, bottom * screen_scale)
    ImageGrab.grab(bbox=full_box).save(save_path)


def capture_team_viewer_window_win32(save_path):
    """
    使用pywin32截取 支持多多块显示器
    多块显示器使用该函数
    :return:
    """
    # 获取DC
    hwin = win32gui.GetDesktopWindow()
    hwindc = win32gui.GetWindowDC(hwin)
    srcdc = win32ui.CreateDCFromHandle(hwindc)
    memdc = srcdc.CreateCompatibleDC()

    # 左上角坐标和右下角坐标
    left, top, right, bottom = find_team_viewer()
    width = int(screen_scale * (right - left))
    height = int(screen_scale * (bottom - top))

    # 创建位图
    bmp = win32ui.CreateBitmap()
    bmp.CreateCompatibleBitmap(srcdc, width, height)
    memdc.SelectObject(bmp)
    memdc.BitBlt((0, 0), (width, height), srcdc, (left, top), win32con.SRCCOPY)
    bmp.SaveBitmapFile(memdc, "./data/tmcap.bmp")

    # 压缩位图
    Image.open("./data/tmcap.bmp").save(save_path)


def inform_email(receiver, file_path):
    """
    邮箱通知 附件方式
    :param receiver:
    :param file_path: 截图文件路径
    :return:
    """
    smtp_server = "smtp.qq.com"
    smtp_port = "465"
    user_account = "xxx@qq.com"
    user_name = "xxx"
    password = "xxxx"

    subject = 'TeamViewer最新自动截屏'  # 主题

    # 创建一个带附件的实例
    message = MIMEMultipart()
    message['From'] = Header("{0}<{1}>".format(user_name, user_account), 'utf-8')  # 发送者
    message['To'] = Header(receiver, 'utf-8')  # 接收者
    message['Subject'] = Header(subject, 'utf-8')

    # 邮件正文内容
    message.attach(MIMEText('TeamViewer {0} 的自动截屏'.format(datetime.datetime.now()), 'plain', 'utf-8'))

    # 添加附件
    attachment = MIMEText(open(file_path, 'rb').read(), 'base64', 'utf-8')
    attachment["Content-Type"] = 'application/octet-stream'
    attachment["Content-Disposition"] = 'attachment; filename="teamviewer.png"'
    message.attach(attachment)

    smtp = smtplib.SMTP_SSL(host=smtp_server, port=smtp_port)
    try:
        smtp.login(user_account, password)
        print("登录成功")
    except smtplib.SMTPException:
        print("登录失败")
        return
    try:
        smtp.sendmail(from_addr=user_account, to_addrs=receiver, msg=message.as_string())
        print("发送成功")
    except smtplib.SMTPException:
        print("发送失败")
        return


if __name__ == '__main__':
    # hwnd = win32gui.FindWindow("#32770", "TeamViewer")
    save_file = "./data/tmcap.png"
    while 1:
        capture_team_viewer_window_win32(save_file)
        inform_email("xxx@qq.com", save_file)
        time.sleep(3600 * 3)
