
from PyQt5.QtCore import pyqtSignal, Qt
from edit_what_ui import Ui_Dialog
from PyQt5.QtWidgets import QDialog, QApplication


class EditWhatDialog(QDialog,Ui_Dialog):
    edit_name_btn_EditWhatDialog_signal = pyqtSignal(str,str)
    delete_btn_EditWhatDialog_signal = pyqtSignal(str,str)
    edit_table_data_btn_EditWhatDialog_signal = pyqtSignal(str,str,int)  #int用于edit的返回
    view_table_btn_EditWhatDialog_signal = pyqtSignal(str,str)

    goback_in_edit_what_signal = pyqtSignal()

    def __init__(self, parent=None, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.setupUi(self)
        self.setWindowFlags(Qt.WindowCloseButtonHint)

    def first_btn(self):  #每次到达这个界面都要初始化到第一个按钮
        # self.edit_name_btn_EditWhatDialog.setChecked(True)
        #view_table_btn_EditWhatDialog()
        self.view_table_btn_EditWhatDialog.setChecked(True)

    def set_left_name(self,str1,str2):  #给这个
        self.file = str1
        self.left_name = str2  #第二个是sheet名
        if str == "class_and_sheet.xlsx":
            self.label_EditWhatDialog.setText("你已经选中%s，请选择下一步操作" % str2)
        else:
            self.label_EditWhatDialog.setText("你已选中编辑历史中的%s，请选择下一步操作" % str2 )

    def ok_btn_EditWhat(self):
        try:
            if self.edit_name_btn_EditWhatDialog.isChecked() == True:
                self.edit_name_btn_EditWhatDialog_signal.emit(self.file,self.left_name)   #edit_是选择修改名字
            elif self.delete_btn_EditWhatDialog.isChecked() == True:
                self.delete_btn_EditWhatDialog_signal.emit(self.file,self.left_name)
            elif self.edit_table_data_btn_EditWhatDialog.isChecked() == True:
                self.edit_table_data_btn_EditWhatDialog_signal.emit(self.file,self.left_name,0)
            else:
                self.view_table_btn_EditWhatDialog_signal.emit(self.file, self.left_name)
        except Exception as e:
            print("edit wht ok_btn_EditWhat",e)

    def goback_in_edit_what(self):
        self.goback_in_edit_what_signal.emit(self.file, self.left_name)

if __name__ == "__main__":
    import sys
    app = QApplication(sys.argv)
    _edit_what_dialog = EditWhatDialog()
    _edit_what_dialog.show()
    sys.exit(app.exec_())
