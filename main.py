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
from PIL import Image


# Python解释器版本3.12

class AddEmojiToWechat:
    def __init__(self):
        self._usual_image_format_list = ['.jpg', '.png', '.jpeg', '.gif', '.bmp']

        self._rootPath = os.path.dirname(os.path.abspath(__file__))
        self._config_file_path = self._get_config_file_path()

        # 原始图片的文件夹路径
        self._original_images_folder_path = self._get_images_folder_path(type='original')
        # 转换后的图片的文件夹路径
        self._converted_images_folder_path = self._get_images_folder_path(type='converted')
        # 转换后的图片路径列表
        self._converted_images_path_list = self._get_converted_images_path_list()

        # 转换图片并将转换后的图片复制到相应文件夹中
        self._convert_images()

        self._add_emoji_to_wechat()

    @staticmethod
    def thread_it(func, *args, daemon: bool = True):
        t = threading.Thread(target=func, args=args)
        t.daemon = daemon
        t.start()

    def _get_images_folder_path(self, type) -> str:
        cf = configparser.ConfigParser()
        cf.read(self._config_file_path, encoding='utf-8')
        if type == 'original':

            imageFolderPath = cf.get('Path', 'imageFolderPath')
        elif type == 'converted':
            imageFolderPath = cf.get('Path', 'imageConvertFolderPath')
        else:
            raise ValueError('type参数错误')
        if not os.path.exists(imageFolderPath):
            self.thread_it(win32api.MessageBox, 0, "请先在ini填写图片文件夹路径", '错误', win32con.MB_ICONWARNING,
                           daemon=False)
            sys.exit()
        return imageFolderPath

    def _get_config_file_path(self) -> str:
        config_path = os.path.join(self._rootPath, 'config.ini')
        if not os.path.exists(config_path):
            with open(config_path, 'w', encoding='utf-8-sig') as f:
                f.write('[Path]\nimageFolderPath=')
            self.thread_it(win32api.MessageBox, 0, "请先在ini填写图片文件夹路径", '错误', win32con.MB_ICONWARNING,
                           daemon=False)
            sys.exit()
        return config_path

    def _get_converted_images_path_list(self) -> list:
        converted_images_path_list = [os.path.join(self._converted_images_folder_path, i) for i in
                                      os.listdir(self._converted_images_folder_path)]
        return converted_images_path_list

    def _add_emoji_to_wechat(self):
        for imageConvertedPath in self._converted_images_path_list:
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

    def _convert_images(self):
        """
        # 转换图片并将转换后的图片保存到相应文件夹中
        # 转换后的图片格式为GIF
        :return: None
        """
        # os.path.splitext 返回一个元组，第一个元素是文件名（包括所处的文件夹），第二个元素是文件扩展名
        # os.path.basename 返回文件名
        # os.path.join 用于路径拼接
        original_images_path_list = [os.path.join(self._original_images_folder_path, i) for i in
                                     os.listdir(self._original_images_folder_path)]
        for image_path in original_images_path_list:
            image_type = os.path.splitext(image_path)[1].lower()
            if image_type in self._usual_image_format_list:
                converted_image_name = os.path.splitext(os.path.basename(image_path))[0] + '_converted.gif'
                converted_image_path = os.path.join(self._converted_images_folder_path, converted_image_name)
                with Image.open(image_path) as img:
                    # 转换为GIF格式
                    img.convert('RGB').save(converted_image_path, 'GIF')


AddEmojiToWechat()
