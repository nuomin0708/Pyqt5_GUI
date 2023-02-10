
from PyQt5.QtCore import pyqtSignal, Qt
from input_new_class_dialog_ui import  Ui_Dialog
from PyQt5.QtWidgets import QDialog, QMessageBox, QApplication


class InputNewClassDialog(QDialog,Ui_Dialog):
    ok_InputNewClassDialog_signal = pyqtSignal(str,str,str)
    goback_InputNewClassDialog_signal = pyqtSignal()

    goback_to_edit_what_dialog_signal = pyqtSignal()
    def __init__(self, parent=None, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.setupUi(self)
        #去掉问号
        self.setWindowFlags(Qt.WindowCloseButtonHint)

    def some_name(self,str1,str2):
        self.book = str1
        self.sheet = str2


    #编辑旧类时 需要文件夹，sheet名，
    #编辑新类 只需要
    def ok_InputNewClassDialog(self):
        try:
            if self.lineEdit.text().strip() == '':
                QMessageBox.information(self, "error", "请先输入内容！")
            else:
                self.ok_InputNewClassDialog_signal.emit(self.lineEdit.text().strip(),self.book,self.sheet)
        except Exception as e:
            print("input_sgg错误",e)
    def goback_InputNewClassDialog(self):
        #编辑名称的返回
        self.goback_to_edit_what_dialog_signal.emit()



if __name__ == "__main__":
    import sys
    app = QApplication(sys.argv)
    _input_new_class_dialog = InputNewClassDialog()
    _input_new_class_dialog.show()
    sys.exit(app.exec_())
