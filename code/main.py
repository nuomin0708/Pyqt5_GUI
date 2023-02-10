
from PyQt5.QtWidgets import QApplication, QMessageBox

from Initialization import Initi
from Login_Pane import  LoginPane   #登录
from Menu_Pane import MenuPane     #主菜单

from Lis_Kuang import LisKuangDialog
from Edit_What import EditWhatDialog

from Edit_Pane import EditPane
from Input_New_Class_Dialog import InputNewClassDialog
from Concrete_Pane import ConcretePane
from One_Dialog import OneDialog
from New_Dialog import  NewDialog
from Path_Dialog import  PathDialog
from View_Pane import  ViewPane

if __name__ == '__main__':
    import sys
    app = QApplication(sys.argv)
    ii = Initi()
    ii.inspect_ini_exist()
    _login_pane = LoginPane()  # 登录变量
    _menu_pane = MenuPane()  # 主菜单变

    _lis_kuang_dialog = LisKuangDialog()
    _edit_what_dialog = EditWhatDialog()

    _edit_pane = EditPane()
    _input_new_class_dialog = InputNewClassDialog()
    _concrete_pane = ConcretePane()
    _one_dialog = OneDialog()
    _new_dialog = NewDialog()
    _path_dialog = PathDialog()
    _view_pane = ViewPane()

    #这是新建类
    def ok_NewDialog(new_class):#这个new——class是新检名
        _menu_pane.write_new_class_to_normal_file_sheet0(new_class)
        _new_dialog.hide()


    def emit_goto_concrete(file_name,sheet_name,excel_row,jiaozheng_row,goto_flag):
        print("main  emit_search_row")
        _concrete_pane.some_canshu(file_name,sheet_name,excel_row,jiaozheng_row,goto_flag)
        _concrete_pane.write_data_to_textline()
        if goto_flag == 0:
            _concrete_pane.show()
            _menu_pane.hide()
        else :
            _concrete_pane.show()
            _view_pane.hide()


    def  goback_to_menu_():
        _menu_pane.show()
        _view_pane.hide()


    _view_pane.goback_to_menu_signal.connect( goback_to_menu_)

    #
    def goto_view_pane_signal() :
        _view_pane.show()
        _concrete_pane.hide()

    _concrete_pane.goto_view_pane_signal.connect(goto_view_pane_signal)




    _view_pane.view_conc_signal.connect(emit_goto_concrete)


    #主菜单界面
    def  open_lis_kuang_dialog():
        _lis_kuang_dialog.clear_clicked_item()
        _lis_kuang_dialog.add_items(_menu_pane.lis,"class_and_sheet.xlsx")
        _lis_kuang_dialog.show()
    def open_edit_pane(file_name,after_edit_name):
        _edit_pane.allow_cao()
        #把数据写入到edit_file中
        _edit_pane.when_to_this_edit_pane(file_name,after_edit_name)
        _edit_pane.write_GUI_data_to_ui_table(file_name,after_edit_name)
        _edit_pane.show()
        _menu_pane.hide()
    def hide_input_dialog():
        _input_new_class_dialog.lineEdit.clear()
        _input_new_class_dialog.hide()
    def one_cao(lis,int):   #str是选择一键还是单个标志
        _one_dialog.add_item(lis,int)  #str是
        _one_dialog.show()

    def add_new_cao():
        #
        _new_dialog.clear_lineEdit()  #
        _new_dialog.show()
        # _input_new_class_dialog.show()
    def open_file_for_sheet0(lis,str):
        try:
            _lis_kuang_dialog.clear_clicked_item()
            _lis_kuang_dialog.add_items(lis,str)
            _lis_kuang_dialog.show()
        except Exception as e:
            QMessageBox.information(_menu_pane, "error", "main写入sheet名错误 %s" % e)

    def emit_path_to_main(lis,int):  #str是sheet名称
        try:
            _one_dialog.add_item(lis,int)
            _one_dialog.show()
        except Exception as e:
            print("main emit_path_to_main",e)
            QMessageBox.information(_menu_pane, "error", "2MenuPane错误 %s" % e)

    # def emit_goto_concrete(str1,int1,int2):
    #     try:
    #         # _concrete_pane.sheet_name = lis[0]
    #         # _concrete_pane.con_row = lis[1]
    #         _concrete_pane.some_canshu(str1,int1,int2)
    #         _concrete_pane.write_data_to_textline()  #把数据写入进去
    #         _concrete_pane.show()
    #         _menu_pane.hide()
        except Exception as e:
            QMessageBox.information(_menu_pane, "error", "3MenuPane错误 %s" % e)

    def goto_history_dialog(new_sheet_name):
        #先要获取编辑历史
        try:
            _menu_pane.open_file_for_sheet0("edit_class_and_sheet.xlsx")
            # _lis_kuang_dialog.listWidget.setCurrentItem(1)
        except Exception as e:
            QMessageBox.information(_menu_pane, "error", "4MenuPane错误 %s" % e)

    #menu
    def new_dialog_again():
        print("666")
        _new_dialog.clear_lineEdit()
        _new_dialog.show()


    def see_old_or_history(str):  #str是文件名
        _menu_pane.open_file_for_sheet0(str)
        # _select_new_or_old_class_dialog.hide()
    #LisKuangDialog
    def ok_LisKuangDialog(str1,str2):
        _edit_what_dialog.set_left_name(str1,str2)
        #并且每次到这个界面都要初始化为第一个选中
        _edit_what_dialog.first_btn()
        _edit_what_dialog.show()
        _lis_kuang_dialog.hide()
    def goback_LisKuangDialog():
        # _select_new_or_old_class_dialog.show()
        _lis_kuang_dialog.hide()
    #EditWhatDialog
    def edit_name_btn_EditWhatDialog(str1,str2):
        #输入新名字
        #str1 是文件夹
        #str2 是sheet名
        # str3 是选择新建类还是编辑名字
        #先把数据清空
        _input_new_class_dialog.lineEdit.clear()
        _input_new_class_dialog.some_name(str1,str2)
        _input_new_class_dialog.show()
        _edit_what_dialog.hide()
    def delete_btn_EditWhatDialog(str1,str2):
        _menu_pane.delete_object(str1,str2)
        _edit_what_dialog.hide()
    def edit_table_data_btn_EditWhatDialog(str1,str2,int):
        # _edit_pane.set_name(str1,str2)
        try:
            _edit_pane.allow_cao()
            _edit_pane.when_to_this_edit_pane(str1,str2,int)
            # _edit_pane.book_name = str1
            # _edit_pane.sheet_name = str2
            _edit_pane.write_GUI_data_to_ui_table(str1,str2)  #分別是类和对象
            _edit_pane.show()
            _edit_what_dialog.hide()
            _menu_pane.hide()
        except Exception as e:
            print("meu",e)

    # SelectPathDialog
    #在这之前，已经在菜单界面说了要不要立刻添加表格数据，这是点击yes后出现的对话框里面的信号
    def add_data_from_pc():
        try:
            # _select_path_dialog.hide()
            #这里就要弹出选择电脑文件对话框
            _menu_pane.open_pc_excel_file()
            #把数据写入到_select_sheet_dialog中，用的是信号传数组到mian里面的open_,,,
        except Exception as e:
            QMessageBox.information(_menu_pane, "error", "5MenuPane错误 %s" % e)
    def add_data_in_ui():
        _edit_pane.allow_cao()
        # _edit_pane.when_to_this_edit_pane()
        _edit_pane.show()
        _menu_pane.hide()  #会把select_path对话框也隐藏

    #SelectSheetDialog
    #这个是针对导入电脑数据的，所以在这个之前还有导入数据的途径选择
    def emit_sheet_dialog(str):  #这个是选中了一个sheet点击sheet对话框里的欧克触发的
        try:
            #获得哪个文件
            sheet = _menu_pane.aac.sheet_by_name(str)
            #需要把里面的内容写入到编辑框中
            # _select_sheet_dialog.hide()  # sheet对话框隐藏，打开edit_pane
            _edit_pane.allow_cao()
            # _edit_pane.when_to_this_edit_pane()
            _edit_pane.write_pc_data_to_ui_table(sheet)
            _edit_pane.show()  # 到达编辑面板
            _menu_pane.hide()
            #先显示说
            QMessageBox.information(_menu_pane, "QAQ", "表格信息导入成功！请根据你的需要来选择新窗口提供的一些操作")
        except Exception as e:
            QMessageBox.information(_menu_pane, "error", "6MenuPane错误 %s" % e)


    # edit_what
    def goback_in_edit_what():
        _lis_kuang_dialog.show()
        _edit_what_dialog.hide()

    #个人中心
    # def goback_menu():
    #     _menu_pane.show()
    #     _personal_pane.hide()

    #InputNewClassDialog
    def ok_InputNewClassDialog(new_class,book,sheet):  #str 是心类还是只是修改个名字
        try:
            #
            # if editname_or_buildnew == "new_":#新建
            #     _menu_pane.write_new_class_to_normal_file_sheet0(new_class)
            #     _input_new_class_dialog.hide()
            # else:
            #     #
             _menu_pane.edit_name(book,sheet,new_class)
        except Exception as e:
            print("main 输入心累",e)
    def goback_to_edit_what_dialog():
        _edit_what_dialog.show()
        _input_new_class_dialog.hide()
    def goback_InputNewClassDialog():
        # _select_new_or_old_class_dialog.show()
        _input_new_class_dialog.hide()
    # def goback_menupane():
    #     _menu_pane.show()
    #     _edit_pane.hide()

    #editpane
    def goback_menupane_edit():
        _menu_pane.show()
        _edit_pane.hide()

    #gui_or_history
    def emit_btn(int):#之恶极更新
        try:
            # _edit_pane.when_to_this_edit_pane()
            _edit_pane.save_data_to_normal_file(int)
        except Exception as e:
            QMessageBox.information(_menu_pane, "error", "7MenuPane错误 %s" % e)

    #_conctre
    def goto_menu_con():
        _menu_pane.show()
        _concrete_pane.hide()

    def gobak_viewpane(file_name,sheet_name):
        #要跟新数据
        #之前的清空
        print("ok")
        try:
            _view_pane.sheet_label.clear()
            _view_pane.tableWidget.clear()
            _view_pane.setName_ane_uitable_ViewPane(file_name,sheet_name)
            #更新
            _view_pane.tableWidget.viewport().update()  # 更新数据
            # print(_view_pane.tableWidget.)
            _view_pane.show()
            _edit_pane.hide()
        except Exception as e:
            print("main gobak_viewpane",e)
    _edit_pane.gobak_viewpane_signal.connect(gobak_viewpane)

    def edit_in_viewpane(file_name,sheet_name,goback_flag):
        #view————》啊编辑
        _edit_pane.allow_cao()
        _edit_pane.when_to_this_edit_pane(file_name,sheet_name,goback_flag)
        _edit_pane.write_GUI_data_to_ui_table(file_name,sheet_name)
        _edit_pane.show()
        _view_pane.hide()


    _view_pane.edit_in_viewpane_signal.connect(edit_in_viewpane)

    #one_dialpog
    #这是一键导入的，设置为0
    def ok_OneDialog_menu():
        _one_dialog.hide()
        _menu_pane.one_to_file()

    #hseet——name为外部文件名 选中的
    def ok_OneDialog_edit(sheet_name,how_to_import_flag):
        try:
            #这个是editPane里面的额按钮
            #-----------------
            if how_to_import_flag == 2: #清空导入
                _one_dialog.hide()
                # 选择了这个sheet
                # 需要把
                _edit_pane.when_import_new_data()
                _edit_pane.one_sheet(sheet_name)
                print("0999")
            #---------------------------


            #这个是外部
            #--------------
            elif how_to_import_flag == 1:

                #注意这里的sheet_name是外部文件的名字

                # file叫做啥666
                # 选择了这个sheet
                # 需要把数据写入到正式文件中
                # sheet = _menu_pane.pitch_file
                # 需要把里面的内容写入到编辑框中
                # _select_sheet_dialog.hide()  # sheet对话框隐藏，打开edit_pane

                # _edit_pane.when_to_this_edit_pane()
                # sheet = _menu_pane.book666[sheet_name]
                # _edit_pane.write_pc_data_to_ui_table(sheet)  #这里的sheet是问检

                # 这个可以在
                # 把sheet写入到表格中

                # print("main ok_OneDialog_edit")
                # print(sheet_name)
                # print(type(sheet_name))
                #新建得多了一步保存
                _one_dialog.hide()
                # 把sheet写入到正式文件中
                # 直接调用menupane里面得写入
                _edit_pane.allow_cao()
                _edit_pane.when_import_new_data()
                _menu_pane. write_just_a_sheet_data_to_normal(sheet_name)
                print("main ok_OneDialog_edit 写入陈工！")
                #
                #把数据写入到edit_fikeH
                _edit_pane.when_to_this_edit_pane("class_and_sheet.xlsx",_menu_pane.new_n)
                #跟新到editpane1中
                _edit_pane.write_GUI_data_to_ui_table("class_and_sheet.xlsx",_menu_pane.new_n)
                #跟新lis
                #这个可以在
                #把sheet写入到表格中

                _edit_pane.show()  # 到达编辑面板
                _menu_pane.hide()
                # 先显示说
                QMessageBox.information(_edit_pane, "QAQ", "表格信息导入成功！请根据你的需要来选择新窗口提供的一些操作")
            else:#追加导入
                _one_dialog.hide()
                _edit_pane.when_import_new_data()
                _edit_pane.append_data_to(sheet_name)
                print("main one_cao 追加")

        except Exception as e:
            print("mian ok_OneDialog_edit",e)

    def tips_about_all_save():
        print("main tips_about_all_save")
        _path_dialog.set_checked_no()
        _path_dialog.show()


    # lis_kuang_dialog
    _lis_kuang_dialog.ok_LisKuangDialog_signal.connect(ok_LisKuangDialog)
    _lis_kuang_dialog.goback_LisKuangDialog_signal.connect(goback_LisKuangDialog)

    #主菜单界面
    _menu_pane.new_dialog_again_signal.connect(new_dialog_again)
    _menu_pane.open_file_for_sheet0_signal.connect(open_file_for_sheet0)  #把数据写入lis——kuang

    #新建类里的 从哪里导入数据
    _menu_pane.emit_path_to_main_signal.connect(emit_path_to_main)

    _menu_pane.emit_goto_concrete_signal.connect(emit_goto_concrete)

    _menu_pane.goto_history_dialog_signal.connect(goto_history_dialog)

    # _menu_pane.old_lis_signal.connect(old_lis_cao)
    _menu_pane.add_new_signal.connect(add_new_cao)
    # _menu_pane.view_history_signal.connect(view_history_cao)
    _menu_pane.old_or_history_signal.connect(see_old_or_history)
    _menu_pane.one_signal.connect(one_cao)

    _menu_pane.hide_input_dialog_signal.connect(hide_input_dialog)
    _menu_pane.open_edit_pane_signal.connect(open_edit_pane)

    _menu_pane. open_lis_kuang_dialog_signal.connect( open_lis_kuang_dialog)

    def ok_PathDialog(int):
        #写入数据到对应文件
        _edit_pane.save_about_path(int)
        #并且隐藏这个path' dialog
        _path_dialog.hide()

    #——path_dialog
    _path_dialog.ok_PathDialog_signal.connect(ok_PathDialog)


    # InputNewClassDialogget_sheet_names()
    _input_new_class_dialog.ok_InputNewClassDialog_signal.connect(ok_InputNewClassDialog)
    _input_new_class_dialog.goback_to_edit_what_dialog_signal.connect(goback_to_edit_what_dialog)


    #editpane
    _edit_pane.goback_menupane_signal.connect(goback_menupane_edit)

    _edit_pane.clear_import_signal.connect(one_cao)  #把sheet名写入到显示框中

    _edit_pane.tips_about_all_save_signal.connect(tips_about_all_save)  #去调用路径框

    _concrete_pane.goto_menu_con_siganl.connect(goto_menu_con)

    def auto_login():
        try:
            _menu_pane.show()
            _login_pane.hide()
            _login_pane.hide()
            _login_pane.hide()
            _login_pane.timer.stop()
        except Exception as e:
            print("main auto_login",e)

    _login_pane.auto_login_signal.connect(auto_login)

    def view_table_btn_EditWhatDialog(file_name,sheet_name):
        #把filename 发到viewPane里
        # _view_pane.File_name_and_sheet_name(file_name,sheet_name)
        _view_pane.setName_ane_uitable_ViewPane(file_name,sheet_name)
        _view_pane.show()
        _menu_pane.hide()
        _edit_what_dialog.hide()

    #EditWhat
    _edit_what_dialog.edit_name_btn_EditWhatDialog_signal.connect(edit_name_btn_EditWhatDialog)
    _edit_what_dialog.delete_btn_EditWhatDialog_signal.connect(delete_btn_EditWhatDialog)
    _edit_what_dialog.edit_table_data_btn_EditWhatDialog_signal.connect(edit_table_data_btn_EditWhatDialog)
    _edit_what_dialog.view_table_btn_EditWhatDialog_signal.connect(view_table_btn_EditWhatDialog)
    _edit_what_dialog.goback_in_edit_what_signal.connect(goback_in_edit_what)

    #one_dialog
    _one_dialog.ok_OneDialog_signal_menu.connect(ok_OneDialog_menu)
    _one_dialog.ok_OneDialog_signal_edit.connect(ok_OneDialog_edit)

    #_new_dialog
    _new_dialog.ok_NewDialog_signal.connect(ok_NewDialog)

    _edit_pane.append_import_signal.connect(one_cao)


    _login_pane.show()
    sys.exit(app.exec_())





