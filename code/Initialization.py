#这里写一些ini文件的相关函数
from os import path
from PyQt5.QtCore import QSettings

class Initi():
    _app_data = QSettings('config.ini', QSettings.IniFormat)
    _app_data.setIniCodec('UTF-8')

    def inspect_ini_exist(self):
        if path.exists('./config.ini'):
            return None
        else:
            #主菜单界面
            self._app_data.setValue("first_use_GUI",None)
            self._app_data.setValue("title_color", None)
            self._app_data.setValue("first_use_GUI_editpane",None)
