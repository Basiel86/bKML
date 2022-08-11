import os

import parse_inch
from DF_DBF import *
from tkinter import filedialog as fd
from datetime import datetime, date
import pandas as pd
import tkinter as tk
from tkinter import ttk
from ttkwidgets import CheckboxTreeview
from parse_inch import parse_inch_prj
import traceback
from columnsCaountValSort import columns_count_val_sort, load_columns_stat_dict

from Export_columns import exp_format


class DB_FORM:

    def __init__(self):

        self.EXP_DAY = '2022-08-20'

        self.lang_list = ["RU", "EN"]
        self.dbf_ext_list = ['dbf', 'DBF']
        # https://www.color-hex.com/popular-colors.php
        # цвета тэгов
        self.tags_colors = {"ANOM": '#ffc3a0',
                            "CON": '#c6e2ff',
                            "WELD": '#c0c0c0',
                            "MARKER": '#bada55',
                            "OTHER": '#f5f5f5',
                            "REP": '#ff7373'}

        self.inch_names_list = parse_inch.get_inch_names_list()
        self.inch_list = parse_inch.get_inch_list()
        self.inch_dict = parse_inch.get_inch_dict()
        # self.diam_list = [4, 4.5, 5.563, 6.625, 8.625, 10.75, 12.75, 14, 16, 18, 20, 22, 24, 26, 28, 30, 32, 34, 36, 38,
        #                   40, 42, 44, 46, 48, 52, 56]

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

        self.custom_listbox_raw = Listbox(width=25, height=33, exportselection=False)
        self.custom_listbox_processed = Listbox(width=30, height=33, exportselection=False)
        self.db_columns_listbox = Listbox(width=35, height=33, exportselection=False)

        self.clear_process_list_button = Button(self.db_process_form, text="Clear", command=self.clear_process_listbox,
                                                width=25)

        self.filter_variable = StringVar()
        self.filter_variable.trace("w", self.filter_total_columns_listbox)
        self.filter_total_columns_textbox = Entry(self.db_process_form, width=35,
                                                  textvariable=self.filter_variable)

        self.process_custom_button = Button(self.db_process_form, text="Process Col",
                                            command=self.process_custom_columns)

        self.diam_list_variable = StringVar(self.db_process_form)
        self.diam_combobox = OptionMenu(self.db_process_form, self.diam_list_variable, *self.inch_names_list)

        # высота в строках
        self.stat_tree = CheckboxTreeview(self.db_process_form, show='tree', column="c1", height=20)
        self.stat_tree.column("# 1", anchor='center', stretch='NO', width=60)

        self.doc_tree = CheckboxTreeview(self.db_process_form, show='tree', column="c1", height=5)
        self.doc_tree.column("# 0", anchor='center', stretch='YES', width=30)
        self.doc_tree.column("# 1", anchor='center', stretch='NO', width=60)

        self.marker_tree = CheckboxTreeview(self.db_process_form, show='tree', column="c1", height=3)
        self.marker_tree.column("# 0", anchor='center', stretch='YES', width=30)
        self.marker_tree.column("# 1", anchor='center', stretch='NO', width=60)

        self.scroll_pane = ttk.Scrollbar(self.db_process_form, command=self.db_columns_listbox.yview)
        self.db_columns_listbox.configure(yscrollcommand=self.scroll_pane.set)

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

        self.process_columns_df = pd.DataFrame(columns=['COL_INDEX', 'COL_NAME'])
        self.total_columns_df = pd.DataFrame(columns=['COL_INDEX', 'COL_NAME'])
        self.total_columns_filter_backup_df = pd.DataFrame(columns=['COL_INDEX', 'COL_NAME'])

        # Инициализация Базы
        self.lng = ''
        self.df_dbf_class = df_DBF()

        self.form_init()

    def form_init(self):
        # читаем аргументы на входе
        arg = ''
        for arg in sys.argv[1:]:
            print('Drag&Drop file: ', arg)

        exp_date_formatted = datetime.strptime(self.EXP_DAY, "%Y-%m-%d").date()
        now_date = date.today()
        days_left = exp_date_formatted - now_date

        ico_abs_path = resource_path('DBF_icon.ico')
        self.db_process_form.wm_iconbitmap(ico_abs_path)

        if exp_date_formatted >= now_date:

            self.db_process_form.title("DB process")
            self.db_process_form.geometry("1555x580")
            self.db_process_form.minsize(1555, 550)
            self.db_process_form.maxsize(1700, 1000)

            self.lang_combobox.place(x=10, y=22)
            self.diam_label.place(x=150, y=4)
            self.diam_combobox.place(x=70, y=22)
            self.path_label.place(x=320, y=33)
            self.path_textbox.place(x=150, y=57)

            self.file_open_button.place(x=10, y=55)
            self.load_db_button.place(x=75, y=55)

            self.columns_label.place(x=320, y=80)
            self.process_custom_button.place(x=320, y=130)

            self.columns_textbox.place(x=150, y=104)
            self.export_db_button.place(x=10, y=100)
            self.clipboard_db_button.place(x=75, y=100)
            self.clear_headers_button.place(x=122, y=100)

            self.stat_tree.place(x=600, y=10)
            self.stat_tree.bind("<Double-1>", self.OnDoubleClick)
            self.scroll_pane.pack(side=tk.RIGHT, fill=tk.Y)
            self.doc_tree.place(x=870, y=10)
            self.marker_tree.place(x=870, y=120)

            self.custom_listbox_raw.place(x=970, y=40)
            self.custom_listbox_processed.place(x=1130, y=40)
            self.db_columns_listbox.place(x=1320, y=40)

            self.clear_process_list_button.place(x=1130, y=10)
            self.filter_total_columns_textbox.place(x=1320, y=10)

            # self.db_columns_listbox.insert(1, "Data Structure")
            # self.db_columns_listbox.insert(2, "Algorithm")
            # self.db_columns_listbox.insert(3, "Data Science")
            # self.db_columns_listbox.insert(4, "Machine Learning")
            # self.db_columns_listbox.insert(5, "Blockchain")
            #
            # self.total_columns_df = self.total_columns_df.append({'COL_INDEX': "Data Structure", 'COL_NAME': "Data Structure"}, ignore_index=True)
            # self.total_columns_df = self.total_columns_df.append({'COL_INDEX': "Algorithm", 'COL_NAME': "Algorithm"},ignore_index=True)
            # self.total_columns_df = self.total_columns_df.append({'COL_INDEX': "Data Science", 'COL_NAME': "Data Science"}, ignore_index=True)
            # self.total_columns_df = self.total_columns_df.append({'COL_INDEX': "Machine Learning", 'COL_NAME': "Machine Learning"}, ignore_index=True)
            # self.total_columns_df = self.total_columns_df.append({'COL_INDEX': "Blockchain", 'COL_NAME': "Blockchain"},ignore_index=True)
            #
            # self.total_columns_filter_backup_df = self.total_columns_filter_backup_df.append({'COL_INDEX': "Data Structure", 'COL_NAME': "Data Structure"}, ignore_index=True)
            # self.total_columns_filter_backup_df = self.total_columns_filter_backup_df.append({'COL_INDEX': "Algorithm", 'COL_NAME': "Algorithm"},ignore_index=True)
            # self.total_columns_filter_backup_df = self.total_columns_filter_backup_df.append({'COL_INDEX': "Data Science", 'COL_NAME': "Data Science"}, ignore_index=True)
            # self.total_columns_filter_backup_df = self.total_columns_filter_backup_df.append({'COL_INDEX': "Machine Learning", 'COL_NAME': "Machine Learning"}, ignore_index=True)
            # self.total_columns_filter_backup_df = self.total_columns_filter_backup_df.append({'COL_INDEX': "Blockchain", 'COL_NAME': "Blockchain"},ignore_index=True)
            #
            # self.custom_listbox_processed.insert(1, "1")
            # self.custom_listbox_processed.insert(2, "2")
            # self.custom_listbox_processed.insert(3, "3")
            # self.custom_listbox_processed.insert(4, "4")
            # self.custom_listbox_processed.insert(5, "5")
            #
            # self.process_columns_df = self.process_columns_df.append({'COL_INDEX': 1, 'COL_NAME': 1}, ignore_index=True)
            # self.process_columns_df = self.process_columns_df.append({'COL_INDEX': 2, 'COL_NAME': 2}, ignore_index=True)
            # self.process_columns_df = self.process_columns_df.append({'COL_INDEX': 3, 'COL_NAME': 3}, ignore_index=True)
            # self.process_columns_df = self.process_columns_df.append({'COL_INDEX': 4, 'COL_NAME': 4}, ignore_index=True)
            # self.process_columns_df = self.process_columns_df.append({'COL_INDEX': 5, 'COL_NAME': 5}, ignore_index=True)

            self.db_columns_listbox.bind('<Double-Button>', self.custom_listbox_processed_add)  # <Button-3>
            self.db_columns_listbox.bind('<Button-3>', self.custom_listbox_processed_replace)
            self.custom_listbox_processed.bind('<Double-Button>', self.custom_listbox_processed_delete)

            self.lang_list_variable.set(self.lang_list[0])

            if arg != '':
                # пишем аргумент в строку пути
                self.path_variable.set(arg)
                # парсим инч
                arg_inch = parse_inch_prj(arg)
                # если нашли инч то селектим его в выпадающем списке
                if arg_inch is not None:
                    inch_index = self.inch_list.index(arg_inch)
                    self.diam_list_variable.set(self.inch_names_list[inch_index])
                else:
                    self.diam_list_variable.set(self.inch_names_list[6])
            else:
                self.diam_list_variable.set(self.inch_names_list[6])

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

    def custom_listbox_processed_add(self, event):

        if self.db_columns_listbox.curselection() != ():
            selection_id = self.db_columns_listbox.curselection()[0]
            db_columns_selected_item = self.db_columns_listbox.get(selection_id)

            col_index = self.total_columns_df.iloc[selection_id][0]
            col_name = self.total_columns_df.iloc[selection_id][1]

            if self.custom_listbox_processed.curselection() != ():
                selected_custom_listbox_processed_id = self.custom_listbox_processed.curselection()[0]
                self.custom_listbox_processed.insert(selected_custom_listbox_processed_id + 1,
                                                     db_columns_selected_item)
                self.custom_listbox_processed.selection_clear(0, END)
                self.custom_listbox_processed.selection_set(selected_custom_listbox_processed_id + 1)

                # добавляем строчку на 0.1 выше индекса
                self.process_columns_df.loc[selected_custom_listbox_processed_id + 0.1] = [col_index,
                                                                                           col_name]  # adding a row
                self.process_columns_df.index = self.process_columns_df.index + 1  # shifting index
                # сортируем индексы чтоб строчка встала на место
                self.process_columns_df = self.process_columns_df.sort_index()  # sorting by index
                # перенумеровываем индексы
                self.process_columns_df = self.process_columns_df.reset_index(drop=True)

            else:
                self.custom_listbox_processed.insert(self.custom_listbox_processed.size(), db_columns_selected_item)
                self.custom_listbox_processed.selection_set(0)

                # если пустой Лист, то просто добавляем записи
                self.process_columns_df = self.process_columns_df.append({'COL_INDEX': col_index, 'COL_NAME': col_name},
                                                                         ignore_index=True)

    def clear_process_listbox(self):
        self.process_columns_df = self.process_columns_df.iloc[0:0]
        self.custom_listbox_processed.delete(0, 'end')
        self.custom_listbox_raw.delete(0, 'end')

    def filter_total_columns_listbox(self, *args):

        filter_text = self.filter_variable.get()

        if filter_text == '':
            self.total_columns_df = self.total_columns_filter_backup_df
            total_columns_list = self.total_columns_filter_backup_df['COL_NAME'].tolist()
            self.db_columns_listbox.delete(0, END)
            for i in range(len(total_columns_list)):
                item = total_columns_list[i]
                self.db_columns_listbox.insert(END, item)
        else:
            filter_index_list = []

            total_columns_list = self.total_columns_filter_backup_df['COL_NAME'].tolist()
            total_columns_index_list = self.total_columns_filter_backup_df.index.tolist()

            self.db_columns_listbox.delete(0, END)

            for i in range(len(total_columns_list)):
                item = total_columns_list[i]
                if filter_text.lower() in item.lower():
                    self.db_columns_listbox.insert(END, item)
                    filter_index_list.append(total_columns_index_list[i])

            self.total_columns_df = self.total_columns_filter_backup_df[self.total_columns_filter_backup_df.index.isin(filter_index_list)]


    def custom_listbox_processed_delete(self, even):
        if self.custom_listbox_processed.curselection() != ():
            selection_id = self.custom_listbox_processed.curselection()[0]
            self.custom_listbox_processed.delete(selection_id)
            self.custom_listbox_processed.selection_set(selection_id)

            # грохаем строку
            self.process_columns_df = self.process_columns_df.drop([selection_id])
            # перенумеровываем индексы
            self.process_columns_df = self.process_columns_df.reset_index(drop=True)

    def custom_listbox_processed_replace(self, even):

        if self.db_columns_listbox.curselection() != ():
            db_columns_selection_id = self.db_columns_listbox.curselection()[0]
            db_columns_selected_item = self.db_columns_listbox.get(db_columns_selection_id)

            col_index = self.total_columns_df.iloc[db_columns_selection_id][0]
            col_name = self.total_columns_df.iloc[db_columns_selection_id][1]

            if self.custom_listbox_processed.curselection() != ():
                custom_listbox_selection_id = self.custom_listbox_processed.curselection()[0]
                self.custom_listbox_processed.delete(custom_listbox_selection_id)
                self.custom_listbox_processed.insert(custom_listbox_selection_id, db_columns_selected_item)
                self.custom_listbox_processed.selection_set(custom_listbox_selection_id)

                self.process_columns_df.loc[custom_listbox_selection_id] = [col_index, col_name]

    def on_closing(self):
        self.db_process_form.destroy()
        sys.exit()

    def clear_headers(self):
        self.columns_textbox.delete(0, END)
        self.columns_textbox.insert(0, '')

    def total_columns_to_list(self):

        """
        Записываем обнаруженные столцбы в Листбокс
        Сортируем в соответствии с Json файлом
        """
        total_columns = self.db_df.columns.values.tolist()

        self.db_columns_listbox.delete(0, 'end')

        # делаем список из словаря количества и порядка столбцов
        columns_stat_list = list(load_columns_stat_dict().keys())
        total_columns_exist_sorted = cross_columns_list(total_columns, columns_stat_list)
        # вернули пересечение и далее добавием остатки под сортированный список
        for i in total_columns:
            if i not in total_columns_exist_sorted:
                total_columns_exist_sorted.append(i)
        # парсим названия столбцов
        total_columns_exist_sorted, total_columns_exist_sorted_names = self.df_dbf_class.parse_columns(
            columns_list=total_columns_exist_sorted, ret_blank=False)

        # чистим и заполняем DF общего число столбцов
        self.total_columns_df = self.total_columns_df.iloc[0:0]
        for i in range(len(total_columns_exist_sorted)):
            self.total_columns_df = self.total_columns_df.append(
                {'COL_INDEX': total_columns_exist_sorted[i], 'COL_NAME': total_columns_exist_sorted_names[i]},
                ignore_index=True)
        self.total_columns_filter_backup_df = self.total_columns_df

        for i in total_columns_exist_sorted_names:
            self.db_columns_listbox.insert('end', i)

    def process_custom_columns(self):

        """
        Кнопка Process Columns
        Берем на вход строку, разбиваем по табулятор и находим соответствия
        """

        try:
            custom_columns = []
            custom_columns_textbox = self.columns_variable.get()[:-1]

            total_columns_index = self.db_df.columns.values.tolist()

            if custom_columns_textbox != '':
                custom_columns_names_raw = custom_columns_textbox.split('\t')
                # возвращаем столбец что нашли и их перевод для Кастом
                custom_columns_index, custom_columns_names = self.df_dbf_class.parse_columns(
                    columns_list=custom_columns_names_raw)
                custom_columns = cross_columns_list(total_columns_index, custom_columns_index)

            # возвращаем пересечение от шаблона к дэфолтному
            cross_columns = cross_columns_list(total_columns_index, exp_format)
            # возвращаем столбец что нашли и их перевод для Дэфолтного
            cross_columns_index, cross_columns_names = self.df_dbf_class.parse_columns(columns_list=cross_columns)

            # чистим и заполняем DF кастомного числа столбцов
            self.process_columns_df = self.process_columns_df.iloc[0:0]

            # если кастом не пустой то грузим его
            if len(custom_columns) > 1:
                columns_raw = custom_columns_names_raw
                columns_names = custom_columns_names

                for i in range(len(custom_columns)):
                    self.process_columns_df = self.process_columns_df.append(
                        {'COL_INDEX': custom_columns_index[i], 'COL_NAME': custom_columns_names[i]}, ignore_index=True)

            # если кастом пустой - заполняем дэфолтный
            else:
                columns_raw = []
                columns_names = cross_columns_names
                for i in range(len(columns_names)):
                    self.process_columns_df = self.process_columns_df.append(
                        {'COL_INDEX': cross_columns_index[i], 'COL_NAME': cross_columns_names[i]}, ignore_index=True)

            self.custom_listbox_raw.delete(0, 'end')
            self.custom_listbox_processed.delete(0, 'end')

            for i in columns_raw:
                self.custom_listbox_raw.insert('end', i)

            list_num = 0
            for i in columns_names:
                self.custom_listbox_processed.insert('end', i)
                # красим красным если нашли #BLANK
                self.custom_listbox_processed.itemconfig("end", bg="#ff7373" if i == '#BLANK' else "white")
                if len(custom_columns) > 2 and i == '#BLANK':
                    self.custom_listbox_raw.itemconfig(list_num, bg="#ff7373" if i == '#BLANK' else "white")
                list_num += 1
        except TypeError and AttributeError:
            print('# ERROR: DB not Loaded!')

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

    @staticmethod
    def get_fea_color_type(fea_name, color_df, lng):
        for i, row in color_df.iterrows():
            df_fea_name = str(color_df.loc[i][f"FEA_{lng}"])
            if fea_name == df_fea_name:
                return str(color_df.loc[i][f"COLOR_TYPE"])

    # грузим базу
    def db_load(self):

        db_path = str(self.path_textbox.get())
        self.lng = str(self.lang_list_variable.get())
        diam_list_value = self.diam_list_variable.get()
        diameter = self.inch_dict[diam_list_value]

        if db_path != "" and db_path[-3:] in self.dbf_ext_list:

            try:
                self.db_df = self.df_dbf_class.convert_dbf(diameter=diameter, dbf_path=db_path, lang=self.lng)

                # получаем список Фич
                stat_df = self.db_df['#FEATURE'].value_counts(dropna=False, normalize=False)
                color_df = self.df_dbf_class.get_color_type_df(lng=self.lng)

                self.stat_tree.delete(*self.stat_tree.get_children())

                # проставляем цвета по тэгам
                for index, color in self.tags_colors.items():
                    self.stat_tree.tag_configure(tagname=index, background=color)

                def treeview(self):
                    w = self.widget
                    curItem = w.focus()
                    # print(w.item(curItem))
                    print(w.item(curItem)['tags'])

                # self.stat_tree.bind("<<TreeviewSelect>>", treeview)

                # бежим по списку и записываем
                for index, value in stat_df.items():
                    current_tag = self.get_fea_color_type(fea_name=index, color_df=color_df, lng=self.lng)

                    self.stat_tree.insert('', 'end', text=index, values=value, tags=(current_tag,))
                # если количество больше 20 - расширяем список до количества
                if len(stat_df) > 20:
                    self.stat_tree.config(height=len(stat_df))

                doc_df = self.db_df['#DOC'].value_counts(dropna=False, normalize=False)
                self.doc_tree.delete(*self.doc_tree.get_children())
                for index, value in doc_df.items():
                    self.doc_tree.insert('', 'end', text=index, values=value)

                ref_df = self.db_df['#REF'].value_counts(dropna=False, normalize=False)
                self.marker_tree.delete(*self.marker_tree.get_children())
                for index, value in ref_df.items():
                    self.marker_tree.insert('', 'end', text=index, values=value)

                # пишем обнаруженные столбцы в список
                self.total_columns_to_list()
                # self.process_custom_columns()

                try:
                    self.write_log(file_path=db_path)
                except Exception as logex:
                    print("LOG Error")

            except Exception as ex:
                print(ex)
                print("DB:form: Что-то пошло не так...")
                print(traceback.format_exc())
        else:
            print("Путь не верен! Ты сбился с пути!?")

    def openfile(self):
        filename = fd.askopenfilename(
            title='Open a file',
            initialdir='/',
            filetypes=[("DBF files", ".dbf .DBF")])

        self.path_variable.set(filename)

        path_inch = parse_inch_prj(filename)
        # если нашли инч то селектим его в выпадающем списке
        if path_inch is not None:
            inch_index = self.inch_list.index(path_inch)
            self.diam_list_variable.set(self.inch_names_list[inch_index])

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

            # !!! ВРЕМЕННО фильтрация работает для каждого поля независимо

            # проверяем на наличие кликнутых в дереве статистики
            checked_items_fea = []
            for item in self.stat_tree.get_checked():
                checked_items_fea.append(self.stat_tree.item(item, "text"))
            # если список не пустой, фильтруем чекнутые
            if len(checked_items_fea) != 0:
                self.db_df_from_tree = self.db_df[self.db_df['#FEATURE'].isin(checked_items_fea)]

            # проверяем на наличие кликнутых в дереве документированных
            checked_items_doc = []
            for item in self.doc_tree.get_checked():
                checked_items_doc.append(self.doc_tree.item(item, "text"))
            if len(checked_items_doc) != 0:
                self.db_df_from_tree = self.db_df[self.db_df['#DOC'].isin(checked_items_doc)]

            # проверяем на наличие кликнутых в дереве маркеров
            checked_items_marker = []
            for item in self.marker_tree.get_checked():
                checked_items_marker.append(self.marker_tree.item(item, "text"))
            if len(checked_items_marker) != 0:
                self.db_df_from_tree = self.db_df[self.db_df['#REF'].isin(checked_items_marker)]

            if checked_items_fea == [] and checked_items_doc == [] and checked_items_marker == []:
                df_for_export = self.db_df
            else:
                df_for_export = self.db_df_from_tree

            # если есть кастом, то экспортим его, в противном случае - шаблон

            export_columns_list_index = self.process_columns_df['COL_INDEX'].tolist()
            export_columns_list_names = self.process_columns_df['COL_INDEX'].tolist()

            if len(export_columns_list_index) != 0:
                exp1 = df_for_export[export_columns_list_index]
                column_names = export_columns_list_names
                columns_count_val_sort(export_columns_list_index)

                # with open(file_path, 'w') as f:
                #    f.write('Custom String\n')

                # df.to_csv(file_path, header=False, mode="a")
                # print (cross_columns)

                if to_clipboard:
                    exp1.to_clipboard(sep=',', index=False)
                    # exp1.to_clipboard(index=False)
                else:
                    absbath = os.path.dirname(path)
                    basename = os.path.basename(path)
                    exportpath = os.path.join(absbath, basename)
                    exportpath_csv = f'{exportpath[:-4]}.csv'
                    exportpath_xlsx = f'{exportpath[:-4]}.xlsx'

                    if self.lng == "RU":
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
                        if self.lng == "SP":
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
                        if self.lng == "SP":
                            exp1.to_excel(exportpath_xlsx, encoding=csv_encoding, index=False)
            else:
                print('# Info: Export table is empty')


if __name__ == "__main__":
    DB_FORM()
