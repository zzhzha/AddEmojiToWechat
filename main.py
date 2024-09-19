import uiautomation as auto
import win32clipboard
import win32con
import pynput
import time
import configparser
import os
import win32api
import sys
import threading
import keyboard


# file = 'D:\\tmp\\test.txt\0D:\\tmp\\股票数据.xlsx\0\0'
# data = file.encode("U16")[2:]
#
# win32clipboard.OpenClipboard()
# try:
#     win32clipboard.EmptyClipboard()
#     win32clipboard.SetClipboardData(
#         win32clipboard.CF_HDROP, bytes(pDropFiles)+data)
# finally:
#     win32clipboard.CloseClipboard()


# formats = auto.GetClipboardFormats()
# for k, v in formats.items():
#     if k == auto.ClipboardFormat.CF_UNICODETEXT:
#         print("文本格式：", auto.GetClipboardText())
#     elif k == auto.ClipboardFormat.CF_HTML:
#         htmlText = auto.GetClipboardHtml()
#         print("富文本格式：", htmlText)
#     elif k == auto.ClipboardFormat.CF_BITMAP:
#         bmp = auto.GetClipboardBitmap()
#         print("位图：", bmp)
def thread_it(func, *args, daemon: bool = True):
    t = threading.Thread(target=func, args=args)
    t.daemon = daemon
    t.start()


def getImageFolderPath(initFilePath):
    cf = configparser.ConfigParser()
    cf.read(initFilePath, encoding='utf-8')
    imageFolderPath = cf.get('Path', 'imageFolderPath')
    if not os.path.exists(imageFolderPath):
        thread_it(win32api.MessageBox, 0, "请先在ini填写图片文件夹路径", '错误', win32con.MB_ICONWARNING, daemon=False)
        sys.exit()
    return imageFolderPath


usualImageFormat = ['.jpg', '.png', '.jpeg', '.gif', '.bmp']
initFilePath = '.\\config.ini'
imageFolderPath = getImageFolderPath(initFilePath)
# 图片的完整路径
imagesPathList = [os.path.join(imageFolderPath, i) for i in os.listdir(imageFolderPath) if
                  os.path.isfile(os.path.join(imageFolderPath, i)) if
                  os.path.splitext(i)[1].lower() in usualImageFormat]
control = pynput.keyboard.Controller()

print(imagesPathList)

imagesPathConvertedList=[]
for i in imagesPathList:
    if os.path.splitext(i)[1].lower() != '.gif':
        print(os.path.splitext(i))
        imageConvertedPath = f'{os.path.splitext(i)[0]}_converted.gif'
        os.system(f'ffmpeg -i {i} {imageConvertedPath}')
        os.remove(i)
        imagesPathConvertedList.append(imageConvertedPath)
    else:
        imagesPathConvertedList.append(i)
print(imagesPathConvertedList)
for imageConvertedPath in imagesPathConvertedList:


    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardText(imageConvertedPath, win32con.CF_UNICODETEXT)
    text = win32clipboard.GetClipboardData()
    win32clipboard.CloseClipboard()



    notepadWindow = auto.WindowControl(Depth=1, ClassName='ChatWnd', Name='文件传输助手')
    notepadWindow.SetActive()
    notepadWindow.SetTopmost(True)
    sendFilesButton = notepadWindow.ButtonControl(depth=11, Name='发送文件')
    sendFilesButton.Click(simulateMove=False)
    getFilesDialogWindow = notepadWindow.WindowControl(Depth=1, Name='打开')
    time.sleep(0.5)
    # 打开文件对话框时，输入位置直接定位到下方的输入栏了
    keyboard.press_and_release('ctrl+v')
    time.sleep(0.5)
    control.press(pynput.keyboard.Key.enter)
    time.sleep(0.5)
    control.press(pynput.keyboard.Key.enter)


# cancelButton=getFilesDialogWindow.ButtonControl(Name='取消',ClassName='Button',depth=1)
# # cancelButton.Click(simulateMove=True)
# addressEdit = cancelButton.GetNextSiblingControl().GetNextSiblingControl()
# addressEdit.Click()
