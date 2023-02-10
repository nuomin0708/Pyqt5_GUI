from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtWidgets import QDialog, QApplication, QAbstractItemView, QMessageBox
from one_ui import Ui_Dialog

class OneDialog(QDialog,Ui_Dialog):
    ok_OneDialog_signal_edit = pyqtSignal(str,int)  #str 是被点击的文本
    ok_OneDialog_signal_menu = pyqtSignal()
    def __init__(self, parent=None, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.setWindowFlags(Qt.WindowCloseButtonHint)
        self.setupUi(self)

    def add_item(self, lis,int):   #第二个str是一键导入的还是单个数据 int2_表示是直接在0表是直接在edit——pane导入时得信号
        try:
            #按钮未选中
            self.ok_OneDialog_btn.setChecked(False)
            self.cancel_OneDialog_btn.setChecked(False)
            print("okwhyu")
            #int==0 表示一键导入
            #int ==1 表示从对话框中导入数据（新建）时
            #int == 2 表示editpane清空再导入
            #int == 3表示追加导入
            self.how_to_import_flag = int
            # self.when_to_import_data_flag = int2_
            #一键导入设置为0 并且左边的设置为不可以选择
            self.listWidget.clear()
            self.listWidget.addItems(lis)
            self.clicked_item = None  #每次到这个界面都要清空
            # self.one_or_just_a_sheet_flag = int_
            if  int  == 0: #一键导入
                #设置为不可以选择
                self.listWidget.setFocusPolicy(Qt.NoFocus)
            else :
                #只可以单选
                self.listWidget.setSelectionMode(QAbstractItemView.SingleSelection)
        except Exception as e:
            print("one dialog",e)

    def clickitem_to_confirm(self, item):
        self.clicked_item = item

    #ok_OneDialog_btn
    def ok_OneDialog(self):
        print("hello")
        #一键导入的话只需要int
        #而单个sheet需要名字
        #所以现在就是把
        try:
            if self.how_to_import_flag !=0 :
                if self.clicked_item != None: #有选中一个内容
                    self.ok_OneDialog_signal_edit.emit(self.clicked_item.text(),self.how_to_import_flag)
                else:
                    QMessageBox.information(self, "warning","请先选中一个对象！")
            else :
                self.ok_OneDialog_signal_menu.emit()
        except Exception as e:
            print("one_dialoggahah",e)

if __name__ == "__main__":
    import sys
    app = QApplication(sys.argv)

    _one_dialog = OneDialog()
    _one_dialog.show()
    sys.exit(app.exec_())


