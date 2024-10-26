import uiautomation as auto
import win32clipboard
import win32con
# import pynput
import time
import configparser
import os
import win32api
import sys
import threading
import keyboard


# Python解释器版本3.12

class AddEmojiToWechat:
    def __init__(self):
        self._usualImageFormat = ['.jpg', '.png', '.jpeg', '.gif', '.bmp']
        self._initFilePath = self._get_init_file_path()
        self._imageFolderPath = self._get_image_folder_path()
        self._imagesPathConvertedList = self._get_images_path_converted_list()
        self._add_emoji_to_wechat()

    @staticmethod
    def thread_it(func, *args, daemon: bool = True):
        t = threading.Thread(target=func, args=args)
        t.daemon = daemon
        t.start()


    def _get_image_folder_path(self):
        cf = configparser.ConfigParser()
        cf.read(self._initFilePath, encoding='utf-8')
        imageFolderPath = cf.get('Path', 'imageFolderPath')
        if not os.path.exists(imageFolderPath):
            self.thread_it(win32api.MessageBox, 0, "请先在ini填写图片文件夹路径", '错误', win32con.MB_ICONWARNING,
                      daemon=False)
            sys.exit()
        return imageFolderPath

    def _get_init_file_path(self):
        if not os.path.exists('.\\config.ini'):
            with open('.\\config.ini', 'w', encoding='utf-8') as f:
                f.write('[Path]\nimageFolderPath=')
            self.thread_it(win32api.MessageBox, 0, "请先在ini填写图片文件夹路径", '错误', win32con.MB_ICONWARNING, daemon=False)
            sys.exit()
        return '.\\config.ini'

    def _get_images_path_converted_list(self):
        l1=[]
        l2= [os.path.join(self._imageFolderPath, i) for i in os.listdir(self._imageFolderPath) if
                  os.path.isfile(os.path.join(self._imageFolderPath, i)) if
                  os.path.splitext(i)[1].lower() in self._usualImageFormat]
        for i in l2:
            if os.path.splitext(i)[1].lower() != '.gif':
                imageConvertedPath = f'{os.path.splitext(i)[0]}_converted.gif'
                os.system(f'ffmpeg -i {i} {imageConvertedPath}')
                os.remove(i)
                l1.append(imageConvertedPath)
            else:
                l1.append(i)
        return l1

    def _add_emoji_to_wechat(self):
        for imageConvertedPath in self._imagesPathConvertedList:
            print(imageConvertedPath)
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardText(imageConvertedPath, win32con.CF_UNICODETEXT)
            text = win32clipboard.GetClipboardData()
            win32clipboard.CloseClipboard()

            notepadWindow = auto.WindowControl(Depth=1, ClassName='ChatWnd', Name='文件传输助手')
            if not auto.WaitForExist(notepadWindow, 3):
                print('文件传输助手窗口未找到，请打开文件传输助手窗口')
                break
            notepadWindow.SetActive()
            notepadWindow.SetTopmost(True)
            time.sleep(1)
            sendFilesButton = notepadWindow.ButtonControl(depth=11, Name='发送文件')
            sendFilesButton.Click(simulateMove=True)
            time.sleep(1)
            getFilesDialogWindow = notepadWindow.WindowControl(Depth=1, Name='打开')
            time.sleep(1)
            # 打开文件对话框时，输入位置直接定位到下方的输入栏了
            keyboard.press_and_release('ctrl+v')
            time.sleep(1)
            keyboard.press_and_release('enter')
            time.sleep(1)
            keyboard.press_and_release('enter')
            time.sleep(1)
            informationListControl = notepadWindow.ListControl(Depth=9, Name='消息')
            time.sleep(1)
            image = informationListControl.GetLastChildControl()
            time.sleep(1)
            image.Click()
            time.sleep(1)
            # Application键是Windows系统的特殊键，为书页键，效果为右键菜单
            keyboard.send('Application')
            time.sleep(1)
            MenuItemControl = notepadWindow.MenuItemControl(Name='添加到表情', depth=4)
            time.sleep(1)

            if not auto.WaitForExist(MenuItemControl, 3):
                keyboard.press_and_release('Esc')
                continue
            MenuItemControl.Click(simulateMove=True)
            time.sleep(1)


AddEmojiToWechat()

