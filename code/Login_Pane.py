from PyQt5.QtCore import pyqtSignal, Qt, QTimer
from PyQt5.QtWidgets import QWidget, QApplication, QMessageBox

from login_ui import Ui_Form
class LoginPane(QWidget,Ui_Form):
    login_closeEvent_signal = pyqtSignal()
    auto_login_signal = pyqtSignal()
    goto_menu_signal = pyqtSignal()
    def __init__(self, parent=None, *args, **kwargs):  #参数有多种，这样写
        super().__init__(parent, *args, **kwargs)
        self.setAttribute(Qt.WA_StyledBackground, True)  # 顶层控件时这个会没有
        self.setupUi(self)
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.showTime)
        self.timer.start(1000 *2)

    #定时
    # def auto_login_timer(self):
    def showTime(self):
        self.auto_login_signal.emit()



    def closeEvent(self,event):
        reply = QMessageBox.question(self, "question", "确定退出系统？", QMessageBox.Yes | QMessageBox.No)
        if reply ==QMessageBox.Yes:
            event.accept()
            # self.login_closeEvent_signal.emit()
        else:
            event.ignore()



    def goto_menu(self):
        self.goto_menu_signal.emit()


if __name__ == '__main__':
    import sys
    app = QApplication(sys.argv)
    _login_pane = LoginPane()
    _login_pane.show()
    sys.exit(app.exec_())
