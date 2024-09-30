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


# Python解释器版本3.12


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

imagesPathConvertedList = []
for i in imagesPathList:
    if os.path.splitext(i)[1].lower() != '.gif':
        imageConvertedPath = f'{os.path.splitext(i)[0]}_converted.gif'
        os.system(f'ffmpeg -i {i} {imageConvertedPath}')
        os.remove(i)
        imagesPathConvertedList.append(imageConvertedPath)
    else:
        imagesPathConvertedList.append(i)
for imageConvertedPath in imagesPathConvertedList:
    print(imageConvertedPath)
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
    time.sleep(0.5)
    informationListControl = notepadWindow.ListControl(Depth=9, Name='消息')
    time.sleep(0.5)
    image = informationListControl.GetLastChildControl()
    time.sleep(0.5)
    image.Click(simulateMove=True)
    time.sleep(0.5)

    # Application键是Windows系统的特殊键，为书页键，效果为右键菜单
    keyboard.send('Application')
    time.sleep(0.5)
    MenuItemControl = notepadWindow.MenuItemControl(Name='添加到表情', depth=4)
    time.sleep(0.5)
    MenuItemControl.Click(simulateMove=True)
    # imageButtonControl = informationListControl.ButtonControl(Depth=6)
    # time.sleep(0.5)
    # imageButtonControl.RightClick(simulateMove=True)
    # keyboard.send('down')
    # time.sleep(0.5)
    # keyboard.send('enter')
    # time.sleep(0.5)

    # informationListControl.RightClick(simulateMove=True)
