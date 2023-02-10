from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtWidgets import QDialog, QApplication

from path_ui import Ui_Dialog

class PathDialog(QDialog,Ui_Dialog):
    ok_PathDialog_signal = pyqtSignal(int)
    #0表示保存在正式文件中
    #1表示历史记录

    def __init__(self, parent=None, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.setWindowFlags(Qt.WindowCloseButtonHint)
        self.setupUi(self)

    def set_checked_no(self):
        #每次到这个界面都要清除选中的
        self.normal_btn.setChecked(False)
        self.history_btn.setChecked(False)


    def ok_PathDialog(self):
        try:
            if self.normal_btn.isChecked()  and self.history_btn.isChecked() :
                self.ok_PathDialog_signal.emit(2)
                print("2")
            else :
                if self.normal_btn.isChecked():
                    print("helo")
                #去调用edit潘
                    self.ok_PathDialog_signal.emit(0)
                if  self.history_btn.isChecked():
                    print("shjaj")
                    self.ok_PathDialog_signal.emit(1)
        except Exception as e:
            print("path dialog ok_PathDialog",e)


if __name__ == "__main__":
    import sys
    app = QApplication(sys.argv)

    _path_dialog = PathDialog()
    _path_dialog.show()
    sys.exit(app.exec_())


