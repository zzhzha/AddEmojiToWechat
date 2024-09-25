import time
import uiautomation as auto
from uiautomation.uiautomation import Bitmap
import win32clipboard
from ctypes import *


class DROPFILES(Structure):
    _fields_ = [
        ("pFiles", c_uint),
        ("x", c_long),
        ("y", c_long),
        ("fNC", c_int),
        ("fWide", c_bool),
    ]


pDropFiles = DROPFILES()
pDropFiles.pFiles = sizeof(DROPFILES)
pDropFiles.fWide = True
matedata = bytes(pDropFiles)


def setClipboardFiles(paths):
    files = ("\0".join(paths)).replace("/", "\\")
    data = files.encode("U16")[2:]+b"\0\0"
    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(
            win32clipboard.CF_HDROP, matedata+data)
    finally:
        win32clipboard.CloseClipboard()


def setClipboardFile(file):
    setClipboardFiles([file])


def readClipboardFilePaths():
    win32clipboard.OpenClipboard()
    paths = None
    try:
        return win32clipboard.GetClipboardData(win32clipboard.CF_HDROP)
    finally:
        win32clipboard.CloseClipboard()


wechatWindow = auto.WindowControl(
    searchDepth=1, Name="微信", ClassName='WeChatMainWndForPC')
wechatWindow.SetActive()
search = wechatWindow.EditControl(Name='搜索')
edit = wechatWindow.EditControl(Name='输入')
messages = wechatWindow.ListControl(Name='消息')
sendButton = wechatWindow.ButtonControl(Name='发送(S)')


def selectSessionFromName(name, wait_time=0.1):
    search.Click()
    auto.SetClipboardText(name)
    edit.SendKeys('{Ctrl}v')
    # 等待微信索引搜索跟上
    time.sleep(wait_time)
    search.SendKeys("{Enter}")


def send_msg(content, msg_type=1):
    if msg_type == 1:
        auto.SetClipboardText(content)
    elif msg_type == 2:
        auto.SetClipboardBitmap(Bitmap.FromFile(content))
    elif msg_type == 3:
        setClipboardFile(content)
    edit.SendKeys('{Ctrl}v')
    sendButton.Click()
