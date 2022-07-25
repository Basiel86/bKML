import os
from DF_DBF import *
from tkinter import filedialog as fd
from datetime import datetime, date
import pandas as pd
import tkinter as tk
from tkinter import ttk
from ttkwidgets import CheckboxTreeview


class DB_FORM:

    def __init__(self):

        self.EXP_DAY = '2022-08-06'

        self.lang_list = ["RU", "EN"]
        self.dbf_ext_list = ['dbf', 'DBF']
        self.diam_list = [4, 4.5, 5.563, 6.625, 8.625, 10.75, 12.75, 14, 16, 18, 20, 22, 24, 26, 28, 30, 32, 34, 36, 38,
                          40, 42, 44, 46, 48, 52, 56]

        self.df_dbf_class = df_DBF
        self.db_df = pd.DataFrame

        self.db_process_form = Tk()

        self.path_label = Label(self.db_process_form, text="DBF файл")
        self.path_variable = StringVar(self.db_process_form)
        self.path_textbox = Entry(self.db_process_form, width=70, textvariable=self.path_variable)
        self.columns_label = Label(self.db_process_form, text="Columns")
        self.columns_variable = StringVar(self.db_process_form)
        self.columns_textbox = Entry(self.db_process_form, width=70, textvariable=self.columns_variable)
        self.load_db_button = Button(self.db_process_form, text="Load DB", command=self.db_load)
        self.export_db_button = Button(self.db_process_form, text="Export DB", command=self.db_export)
        self.clipboard_db_button = Button(self.db_process_form, text="to Clip",
                                          command=lambda: self.db_export(to_clipboard=True))
        self.clear_headers_button = Button(self.db_process_form, text="CL", command=self.clear_headers)

        self.lang_list_variable = StringVar(self.db_process_form)
        self.lang_combobox = OptionMenu(self.db_process_form, self.lang_list_variable, *self.lang_list)

        self.file_open_button = Button(self.db_process_form, text="Open file", command=self.openfile)
        self.blankLabel = Label(self.db_process_form, text="          ")
        self.diam_label = Label(self.db_process_form, text="Diam")

        self.diam_list_variable = StringVar(self.db_process_form)
        self.diam_combobox = OptionMenu(self.db_process_form, self.diam_list_variable, *self.diam_list)

        self.stat_tree = CheckboxTreeview(self.db_process_form, show='tree', column="c1")  # hide tree headings
        self.stat_tree.column("# 1", anchor='center', stretch='YES', width=60)

        self.doc_tree = CheckboxTreeview(self.db_process_form, show='tree', column="c1")  # hide tree headings
        self.doc_tree.column("# 1", anchor='center', stretch='NO', width=10)

        self.scroll_pane = ttk.Scrollbar(self.db_process_form, command=self.stat_tree.yview)
        self.stat_tree.configure(yscrollcommand=self.scroll_pane.set)

        style = ttk.Style(self.db_process_form)
        # remove the indicator in the treeview
        style.layout('Checkbox.Treeview.Item',
                     [('Treeitem.padding',
                       {'sticky': 'nswe',
                        'children': [('Treeitem.image', {'side': 'left', 'sticky': ''}),
                                     ('Treeitem.focus', {'side': 'left', 'sticky': '',
                                                         'children': [('Treeitem.text',
                                                                       {'side': 'left', 'sticky': ''})]})]})])
        style.configure('Checkbox.Treeview', borderwidth=1, relief='sunken')

        self.form_init()

    def form_init(self):
        # читаем аргументы на входе
        arg = ''
        for arg in sys.argv[1:]:
            print('Drag&Drop file: ', arg)

        exp_date_formatted = datetime.strptime(self.EXP_DAY, "%Y-%m-%d").date()
        now_date = date.today()
        days_left = exp_date_formatted - now_date

        ico_abs_path = resource_path('KML.ico')
        self.db_process_form.wm_iconbitmap(ico_abs_path)

        if exp_date_formatted >= now_date:

            self.db_process_form.title("DB process")
            self.db_process_form.geometry("1090x220")

            self.diam_label.place(x=25, y=5)
            self.diam_combobox.place(x=10, y=22)
            self.lang_combobox.place(x=80, y=22)
            self.path_label.place(x=150, y=5)
            self.path_textbox.place(x=150, y=28)

            self.file_open_button.place(x=10, y=55)
            self.load_db_button.place(x=80, y=55)

            self.columns_label.place(x=150, y=80)
            self.columns_textbox.place(x=150, y=104)
            self.export_db_button.place(x=10, y=100)
            self.clipboard_db_button.place(x=75, y=100)
            self.clear_headers_button.place(x=122, y=100)

            self.stat_tree.place(x=600, y=10)
            self.stat_tree.bind("<Double-1>", self.OnDoubleClick)
            self.scroll_pane.pack(side=tk.RIGHT, fill=tk.Y)
            self.doc_tree.place(x=870, y=10)

            if arg != '':
                self.path_variable.set(arg)

            self.lang_list_variable.set(self.lang_list[0])
            self.diam_list_variable.set(self.diam_list[4])

            # self.blankLabel.pack(side='left')

            self.db_process_form.protocol("WM_DELETE_WINDOW", self.on_closing)
            self.db_process_form.mainloop()

        else:
            self.db_process_form.geometry("1900x850")
            self.db_process_form.title('EXPIRED')
            self.db_process_form.geometry("300x50")

            exp_label = Label(master=self.db_process_form, text="Your version has expired", fg='red', font=15)
            exp_label.pack()
            exp_label2 = Label(master=self.db_process_form, text="Please update", fg='red', font=15)
            exp_label2.pack()

            def on_closing():
                self.db_process_form.destroy()
                sys.exit()

            self.db_process_form.protocol("WM_DELETE_WINDOW", on_closing)
            self.db_process_form.mainloop()

        self.db_process_form.mainloop()

    def on_closing(self):
        self.db_process_form.destroy()
        sys.exit()

    def clear_headers(self):
        self.columns_textbox.delete(0, END)
        self.columns_textbox.insert(0, '')

    def OnDoubleClick(self, event):

        # активное
        # item = self.stat_tree.selection()[0]
        # print("you clicked on", self.stat_tree.item(item, "text"))

        # пробежка по чекнутым
        for item in self.stat_tree.get_checked():
            item_text = self.stat_tree.item(item, "text")
            # print(item_text)

    @staticmethod
    def write_log(file_path=""):
        log_path = r"\\vasilypc\Vasily Shared (Full Access)\###\DBLog\DBlog.txt"

        try:
            if not os.path.exists(log_path):
                with open(log_path, 'w') as file:
                    file.write("")
                file.close()

            log_file = open(log_path, 'a', encoding='utf-8')

            now_date_time = datetime.today().strftime("%d.%m.%Y %H:%M:%S")
            username = os.getenv('username')
            pc_name = os.environ['COMPUTERNAME']

            log_file.write(f'{now_date_time}\t'
                           f'user: {username}\t\t'
                           f'pc: {pc_name}\t\t'
                           f'{file_path}\n')
            log_file.close()
        except Exception as ex:
            print("LOG Error: ", ex)

    # грузим базу
    def db_load(self):
        db_path = str(self.path_textbox.get())

        lng = str(self.lang_list_variable.get())
        diameter = float(self.diam_list_variable.get())

        if db_path != "" and db_path[-3:] in self.dbf_ext_list:

            try:
                print("\nProcessing....")

                self.df_dbf_class = df_DBF(DBF_path=db_path, lang=lng)
                self.db_df = self.df_dbf_class.convert_dbf(diameter=diameter)

                stat_df = self.db_df['#FEATURE'].value_counts(dropna=False, normalize=False)
                self.stat_tree.delete(*self.stat_tree.get_children())
                for index, value in stat_df.items():
                    self.stat_tree.insert('', 'end', text=index, values=value)

                doc_df = self.db_df['#DOC'].value_counts(dropna=False, normalize=False)
                self.doc_tree.delete(*self.doc_tree.get_children())
                for index, value in doc_df.items():
                    self.doc_tree.insert('', 'end', text=index, values=value)

                try:
                    self.write_log(file_path=db_path)
                except Exception as logex:
                    print("LOG Error")

            except Exception as ex:
                print(ex)
                print("Что-то пошло не так...")
        else:
            print("Путь не верен! Ты сбился с пути!?")

    def openfile(self):
        filename = fd.askopenfilename(
            title='Open a file',
            initialdir='/',
            filetypes=[("DBF files", ".dbf .DBF")])

        self.path_variable.set(filename)

    @staticmethod
    def resource_path(relative_path):
        try:
            # PyInstaller creates a temp folder and stores path in _MEIPASS
            base_path = sys._MEIPASS
        except Exception as ex:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)

    def db_export(self, to_clipboard=False):

        path = self.path_variable.get()

        if path != '':
            lang = self.lang_list_variable.get()

            custom_columns = self.columns_variable.get()[:-1]
            custom_columns = custom_columns.split('\t')

            # возвращаем столбец что нашли и их перевод для Кастом
            custom_columns, custom_columns_names = self.df_dbf_class.parse_columns(columns_list=custom_columns,
                                                                                   lang=lang)

            exp_format = ['#FEA_NUM', '#JN', '#DIST_START', '#US', '#JL', '#DOC', '#FEATURE', '#FEATURE_TYPE',
                          '#DESCR', '#LENGTH', '#WIDTH', '#DEPTH_PRC', '#DEPTH_MM', '#AMPL', '#DEPTH_PREV',
                          '#ORIENT_DEG',
                          '#WT', '#REMAIN_WT', '#LOC', '#DIMM', '#CLUSTER', '#ERF', '#PSAFE', '#LAT', '#LONG', '#ALT',
                          '#FEA_NUM_PREV', '#VTD_NUM', '#YEARS']

            total_columns = self.db_df.columns.values.tolist()

            # возвращаем пересечение от шаблона к имеющимся
            cross_columns = cross_columns_list(total_columns, exp_format)
            # возвращаем столбец что нашли и их перевод для Дэфолтного
            cross_columns_return, column_names_cross = self.df_dbf_class.parse_columns(columns_list=cross_columns,
                                                                                       lang=lang)

            # добавляем ДОК столбцы в кастом столбцы
            # custom_columns.append('#DOC')
            # custom_columns_names.append('#DOC')

            # проверяем на наличие кликнутых в дереве статистики
            checked_items_fea = []
            for item in self.stat_tree.get_checked():
                checked_items_fea.append(self.stat_tree.item(item, "text"))
            # если список не пустой, фильтруем чекнутые
            if len(checked_items_fea) != 0:
                self.db_df = self.db_df[self.db_df['#FEATURE'].isin(checked_items_fea)]
            # проверяем на наличие кликнутых в дереве документированных
            checked_items_doc = []
            for item in self.doc_tree.get_checked():
                checked_items_doc.append(self.doc_tree.item(item, "text"))
            if len(checked_items_doc) != 0:
                self.db_df = self.db_df[self.db_df['#DOC'].isin(checked_items_doc)]

            # возвращаем пересечение от шаблона к имеющимся
            custom_columns = cross_columns_list(total_columns, custom_columns)

            # если есть кастом, то экспортим его, в противном случае - шаблон
            if len(custom_columns) > 2:
                exp1 = self.db_df[custom_columns]
                column_names = custom_columns_names
            else:
                exp1 = self.db_df[cross_columns]
                column_names = column_names_cross

            # with open(file_path, 'w') as f:
            #    f.write('Custom String\n')

            # df.to_csv(file_path, header=False, mode="a")
            # print (cross_columns)

            if to_clipboard:
                exp1.to_clipboard(sep=',', index=False)
            else:
                absbath = os.path.dirname(path)
                basename = os.path.basename(path)
                exportpath = os.path.join(absbath, basename)
                exportpath_csv = f'{exportpath[:-4]}.csv'
                exportpath_xlsx = f'{exportpath[:-4]}.xlsx'

                if lang == "RU":
                    csv_encoding = "cp1251"
                else:
                    csv_encoding = "utf-8"
                try:
                    self.columns_textbox.delete(0, END)
                    self.columns_textbox.insert(0, '')
                    to_csv_custom_header(df=exp1, csv_path=exportpath_csv, column_names=column_names,
                                         csv_encoding=csv_encoding)
                    # exp1.to_csv(exportpath_csv, encoding=csv_encoding, index=False)
                    # xw.Book(exportpath_csv)
                    # xw.view(exportpath_csv, table=False)
                    if lang == "SP":
                        exp1.to_excel(exportpath_xlsx, encoding=csv_encoding, index=False)

                except Exception as PermissionError:
                    self.columns_textbox.delete(0, END)
                    self.columns_textbox.insert(0, '')
                    print(f"'{basename[:-4]}.csv' is opened, saved as '{basename[:-4]}_1.csv'")
                    exportpath = os.path.join(absbath, basename)
                    exportpath_csv = f'{exportpath[:-4]}_1.csv'
                    exportpath_xlsx = f'{exportpath[:-4]}_1.xlsx'
                    to_csv_custom_header(df=exp1, csv_path=exportpath_csv, column_names=column_names,
                                         csv_encoding=csv_encoding)
                    # exp1.to_csv(exportpath_csv, encoding=csv_encoding, index=False)
                    if lang == "SP":
                        exp1.to_excel(exportpath_xlsx, encoding=csv_encoding, index=False)


if __name__ == "__main__":
    DB_FORM()