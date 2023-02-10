
from PyQt5.QtCore import pyqtSignal, Qt
from new_ui import  Ui_Dialog
from PyQt5.QtWidgets import QDialog, QMessageBox, QApplication


class NewDialog(QDialog,Ui_Dialog):
    ok_NewDialog_signal = pyqtSignal(str)
    goback_to_edit_what_dialog_signal = pyqtSignal()
    def __init__(self, parent=None, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.setupUi(self)
        self.setWindowFlags(Qt.WindowCloseButtonHint)

    def clear_lineEdit(self):
        #两个linedeirt情节哦那个
        self.lineEdit.clear()

    #编辑新类 只需要
    def ok_NewDialog(self):
        try:
            if self.lineEdit.text().strip() == '':
                QMessageBox.information(self, "error", "请先输入内容！")
            else:
                self.ok_NewDialog_signal.emit(self.lineEdit.text().strip())
        except Exception as e:
            print("new dilaog错误",e)



if __name__ == "__main__":
    import sys
    app = QApplication(sys.argv)
    _new_dialog = NewDialog()
    _new_dialog.show()
    sys.exit(app.exec_())
