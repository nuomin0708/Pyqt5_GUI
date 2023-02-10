
from PyQt5.QtCore import pyqtSignal, Qt
# from PyQt5.QtGui import QKeySequence
from lis_kuang_ui import  Ui_Dialog
from PyQt5.QtWidgets import QDialog, QMessageBox, QApplication\
    # ,QShortcut


class LisKuangDialog(QDialog,Ui_Dialog):
    ok_LisKuangDialog_signal = pyqtSignal(str,str)  #第一个是文件，第二个是sheet
    cancel_LisKuangDialog_signal = pyqtSignal()
    goback_LisKuangDialog_signal = pyqtSignal()

    def __init__(self, parent=None, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.clicked_item = None
        self.setupUi(self)
        self.setWindowFlags(Qt.WindowCloseButtonHint)

    # def keyPressEvent(self, QKeyEvent):
    #     if QKeyEvent.key() == Qt.Key_Return:
    #         print('Space')
        # self.ok_btn_LisKuangDialog.setShortcut('enter')
        # self.cancel_btn_LisKuangDialog.setShortcut('Ctrl+Q')
        # print("wertyuiop")

        # self.ok_btn_LisKuangDialog.returnPressed(self.et)
        # QShortcut(QKeySequence(self.tr("enter")), self, self.et)

    # def et(self):
    #     print("heli")

    # def keyPressEvent(self, event):
    #     if event.key() == QtCore.Qt.Key_Enter:
    #         print("ok")
            # self.slotLogin()

    # ok_btn_LisKuangDialog

    # def keyPressEvent(self, QKeyEvent):
    #     if QKeyEvent.key() == Qt.Key_Return:
    #         print('Space')

    # 在第二篇文章中作者说: 大键盘上的键是Qt.Key_R

    #每次到这个界面时都要把之前的self.clicked_item清除掉
    def clear_clicked_item(self):
        self.clicked_item = None

    def add_items(self, lis,str):
        self.listWidget.clear()
        self.listWidget.addItems(lis)
        self.file = str
        if str == "class_and_sheet.xlsx":
            self.title_lis_kuang.setText("以下是已存在的文件")
        else:
            self.title_lis_kuang.setText("以下是编辑记录")

        # 这个默认无选择
    def clickitem_to_confirm_btn(self, item):
        self.clicked_item = item

    def ok_LisKuangDialog(self):
        try:
            if self.clicked_item != None:  # 有选择
                # 发射信号1到main，把这个关掉
                self.ok_LisKuangDialog_signal.emit(self.file,self.clicked_item.text())
            else:
                QMessageBox.information(self, "warning", "请先选中一个对象!")
        except Exception as e:
            QMessageBox.information(self, "error", "Lis_Kuang错误 %s" % e)

    def goback_LisKuangDialog(self):
        self.goback_LisKuangDialog_signal.emit()

if __name__ == "__main__":
    import sys
    app = QApplication(sys.argv)
    _lis_kuang_dialog = LisKuangDialog()
    _lis_kuang_dialog.show()
    sys.exit(app.exec_())
