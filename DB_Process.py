from tkinter import *
import parse_inch
import parse_templates
from DF_DBF import *
from tkinter import filedialog as fd
from datetime import datetime, date
import pandas as pd
from tkinter import ttk
from ttkwidgets import CheckboxTreeview
from parse_inch import parse_inch_prj
import traceback
from columnsCountValSort import columns_count_val_sort, load_columns_stat_dict
import pyfiglet
from parse_templates import *
from play_sound import PlaySound
from parse_cfg import CFG

pd.options.mode.chained_assignment = None  # default='warn'

sound_path = r'c:\Users\Vasily\OneDrive\Macro\PYTHON\bKML\Sounds\Duck.mp3'


class DB_FORM:

    def __init__(self):

        self.EXP_DAY = '2022-10-01'
        # 'fender' 'kban' 'larry3d'
        print('\n' + pyfiglet.figlet_format("DB Process", justify='center', font='larry3d'))
        self.cfg = CFG('DB Process')
        self.lang_list = ["RU", "EN"]
        self.dbf_ext_list = ['dbf', 'DBF']
        # https://www.color-hex.com/popular-colors.php
        # Цвета Тэгов
        self.tags_colors = {"ANOM": '#ffc3a0',
                            "CON": '#c6e2ff',
                            "WELD": '#c0c0c0',
                            "MARKER": '#bada55',
                            "OTH": '#f5f5f5',
                            "REP": '#ff7373',
                            "GPS_TMP": '#cd23d5'}
        # Словарь Тэгов к Именам
        self.tags_dict_id_name = {"ANOM": 'Anomalies',
                                  "CON": 'Costruct',
                                  "WELD": 'Welds',
                                  "MARKER": 'Markers',
                                  "OTH": 'Other',
                                  "REP": 'Repairs',
                                  "GPS_TMP": 'GPS_TMP'}
        # Словарь Имен к Тэгам
        self.tags_dict_name_id = {'Anomalies': 'ANOM',
                                  'Costruct': "CON",
                                  'Welds': 'WELD',
                                  'Markers': "MARKER",
                                  'Other': "OTH",
                                  'Repairs': "REP",
                                  'GPS_TMP': "GPS_TMP"}

        self.sound = PlaySound()
        if self.cfg.read_cfg(section="SOUND", key="enable") is True:
            self.play_sound()

        # Список Тэгов в текущем списке особенностей
        self.stat_tree_tags_list = []

        self.inch_names_list = parse_inch.get_inch_names_list()
        self.inch_list = parse_inch.get_inch_list()
        self.inch_dict = parse_inch.get_inch_dict()
        self.templates_list = get_templates_list()
        # self.diam_list = [4, 4.5, 5.563, 6.625, 8.625, 10.75, 12.75, 14, 16, 18, 20, 22, 24, 26, 28, 30, 32, 34, 36, 38,
        #                   40, 42, 44, 46, 48, 52, 56]

        self.df_dbf_class = df_DBF
        self.db_df = pd.DataFrame

        self.db_process_form = Tk()

        self.path_label = Label(self.db_process_form, text="DBF файл")
        self.path_variable = StringVar(self.db_process_form)
        self.path_textbox = Entry(self.db_process_form, width=70, textvariable=self.path_variable)
        self.columns_label = Label(self.db_process_form, text="Custom Columns")
        self.columns_variable = StringVar(self.db_process_form)
        self.columns_textbox = Entry(self.db_process_form, width=70, textvariable=self.columns_variable)
        self.load_db_button = Button(self.db_process_form, text="Load DB", command=self.db_load)
        self.export_db_button = Button(self.db_process_form, text="to CSV", width=10, command=self.db_export)
        self.clipboard_db_button = Button(self.db_process_form, text="to Clip", width=10,
                                          command=lambda: self.db_export(to_clipboard=True))
        self.clear_headers_button = Button(self.db_process_form, text="Clear", command=self.clear_headers)

        self.lang_list_variable = StringVar(self.db_process_form)
        self.lang_combobox = OptionMenu(self.db_process_form, self.lang_list_variable, *self.lang_list)

        self.file_open_button = Button(self.db_process_form, text="Open file", command=self.openfile)
        self.blankLabel = Label(self.db_process_form, text="          ")
        self.diam_label = Label(self.db_process_form, text="Diam")

        self.custom_listbox_raw = Listbox(width=30, height=33, exportselection=False)
        self.custom_listbox_processed = Listbox(width=30, height=33, exportselection=False)
        self.db_columns_listbox = Listbox(width=35, height=33, exportselection=False)

        self.clear_process_list_button = Button(self.db_process_form, text="Clear", command=self.clear_process_listbox,
                                                width=25)

        self.filter_variable = StringVar()
        self.filter_variable.trace("w", self.filter_total_columns_listbox)
        self.filter_total_columns_textbox = Entry(self.db_process_form, width=35,
                                                  textvariable=self.filter_variable)

        self.process_custom_button = Button(self.db_process_form, text="Process Custom Columns",
                                            command=self.process_custom_columns)

        self.diam_list_variable = StringVar(self.db_process_form)
        self.diam_combobox = OptionMenu(self.db_process_form, self.diam_list_variable, *self.inch_names_list)

        self.templates_variable = StringVar(self.db_process_form)
        self.templates_combobox = OptionMenu(self.db_process_form, self.templates_variable, *self.templates_list)
        self.templates_combobox.configure(width=16)
        self.templates_variable.set('Default')

        self.template_name_label = Label(self.db_process_form, text="New Template Name")
        self.template_list_label = Label(self.db_process_form, text="Templates List")
        self.template_name_variable = StringVar(self.db_process_form)
        self.template_name_textbox = Entry(self.db_process_form, width=22, textvariable=self.template_name_variable)

        self.save_template_button = Button(self.db_process_form, text="Save template", command=self.save_template)
        self.rewrite_template_button = Button(self.db_process_form, text="Rewrite template",
                                              command=self.rewrite_template)
        self.load_template_button = Button(self.db_process_form, text="Load template", command=self.load_template)
        self.delete_template_button = Button(self.db_process_form, text="Delete template", command=self.delete_template)
        self.open_templates_folder_button = Button(self.db_process_form, text="Open Templates Folder",
                                                   command=parse_templates.open_templates_folder)

        # высота в строках
        self.stat_tree = CheckboxTreeview(self.db_process_form, show='tree', column="c1", height=20)
        self.stat_tree.column("# 1", anchor='center', stretch='NO', width=60)

        self.stat_tree_tags_variable = StringVar(self.db_process_form)
        self.stat_tree_tags_combobox = OptionMenu(self.db_process_form, self.stat_tree_tags_variable,
                                                  self.stat_tree_tags_list)
        self.stat_tree_tags_select_button = Button(self.db_process_form, text="Select group",
                                                   command=self.select_group_checks)

        self.doc_tree = CheckboxTreeview(self.db_process_form, show='tree', column="c1", height=5)
        self.doc_tree.column("# 0", anchor='center', stretch='YES', width=30)
        self.doc_tree.column("# 1", anchor='center', stretch='NO', width=60)

        self.marker_tree = CheckboxTreeview(self.db_process_form, show='tree', column="c1", height=3)
        self.marker_tree.column("# 0", anchor='center', stretch='YES', width=30)
        self.marker_tree.column("# 1", anchor='center', stretch='NO', width=60)

        self.sel_all_checks = Button(self.db_process_form, text="Select All", command=self.select_all_checks)
        self.remove_all_checks = Button(self.db_process_form, text="Clear All", command=self.remove_all_checks)
        self.invert_checks = Button(self.db_process_form, text="Invert", command=self.invert_checks)

        self.scroll_pane = ttk.Scrollbar(self.db_process_form, command=self.db_columns_listbox.yview)
        self.db_columns_listbox.configure(yscrollcommand=self.scroll_pane.set)

        self.play_button = Button(self.db_process_form, text="Play", command=self.play_sound,
                                  relief=GROOVE)
        self.play_button_stop = Button(self.db_process_form, text="Stop", command=self.stop_sound,
                                       relief=GROOVE)

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

        # self.add_tree()
        # читаем аргументы на входе
        arg = ''
        for arg in sys.argv[1:]:
            if os.path.exists(arg):
                print('Drag&Drop file: ', arg)

        exp_date_formatted = datetime.strptime(self.EXP_DAY, "%Y-%m-%d").date()
        now_date = date.today()
        days_left = exp_date_formatted - now_date

        ico_abs_path = resource_path('DBF_icon.ico')
        self.db_process_form.wm_iconbitmap(ico_abs_path)

        if exp_date_formatted >= now_date:

            self.db_process_form.title("DB process")
            self.db_process_form.geometry("1630x610")
            self.db_process_form.minsize(1630, 600)
            self.db_process_form.maxsize(1700, 650)

            col_span = 5
            self.lang_combobox.grid(row=1, column=0, sticky=EW)
            self.diam_label.grid(row=0, column=1, sticky=EW, columnspan=2)
            self.diam_combobox.grid(row=1, column=1, sticky=EW, columnspan=2)
            self.path_label.grid(row=1, column=3, columnspan=col_span)
            self.path_textbox.grid(row=2, column=3, columnspan=col_span, sticky=EW)

            self.file_open_button.grid(row=2, column=0, sticky=EW, padx=2)
            self.load_db_button.grid(row=2, column=1, sticky=EW, padx=2, columnspan=2)

            self.process_custom_button.grid(row=6, column=1, columnspan=col_span + 2)
            self.columns_label.grid(row=3, column=1, columnspan=col_span + 2, sticky=EW)
            self.columns_textbox.grid(row=5, column=1, columnspan=col_span + 2, sticky=EW)
            self.clear_headers_button.grid(row=5, column=0, sticky=EW, padx=2)

            self.stat_tree.grid(row=2, column=9, rowspan=30, columnspan=3, padx=5, pady=5, sticky=N)
            self.stat_tree.bind("<Double-1>", self.OnDoubleClick)
            self.doc_tree.grid(row=2, column=12, rowspan=4, columnspan=2, sticky=N, pady=5)
            self.marker_tree.grid(row=6, column=12, rowspan=4, columnspan=2, sticky=NS)

            self.stat_tree_tags_combobox.grid(row=1, column=9, padx=5, columnspan=2, pady=5, sticky=EW)
            self.stat_tree_tags_select_button.grid(row=1, column=11, columnspan=1, padx=5, pady=5, sticky=EW)

            self.sel_all_checks.grid(row=0, column=9, sticky=E + W + S, pady=2, padx=3)
            self.remove_all_checks.grid(row=0, column=10, sticky=E + W + S, pady=2, padx=3)
            self.invert_checks.grid(row=0, column=11, sticky=E + W + S, pady=2, padx=3)

            self.custom_listbox_raw.grid(row=2, column=20, rowspan=40, columnspan=2, padx=3, sticky=E)
            self.custom_listbox_processed.grid(row=2, column=23, rowspan=40, columnspan=2)
            self.db_columns_listbox.grid(row=2, column=25, rowspan=40, columnspan=2)
            self.scroll_pane.grid(row=2, column=27, rowspan=40, columnspan=1, sticky=NS)

            self.clear_process_list_button.grid(row=1, column=23, columnspan=2)
            self.filter_total_columns_textbox.grid(row=1, column=25, columnspan=2)

            self.template_name_label.grid(row=10, column=0, columnspan=2)
            self.template_name_textbox.grid(row=11, column=0, columnspan=2, padx=3, sticky=EW)
            self.template_list_label.grid(row=12, column=0, columnspan=2)
            self.templates_combobox.grid(row=13, column=0, columnspan=2)
            self.open_templates_folder_button.grid(row=16, column=0, columnspan=2)

            self.save_template_button.grid(row=11, column=2, padx=3)
            self.rewrite_template_button.grid(row=11, column=3, padx=3)
            self.delete_template_button.grid(row=11, column=4, padx=3, sticky=W)

            self.load_template_button.grid(row=13, column=2, columnspan=1)

            self.export_db_button.grid(row=13, column=4, sticky=W)
            self.clipboard_db_button.grid(row=13, column=3)

            self.play_button.grid(row=18, column=0, columnspan=1)
            self.play_button_stop.grid(row=18, column=1, columnspan=1)

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

            if arg != '' and os.path.exists(arg):
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

    def play_sound(self):
        self.sound.play_music()
        self.cfg.store_settings(section="SOUND", key="enable", value=True)

    def stop_sound(self):
        self.sound.stop_music()
        self.cfg.store_settings(section="SOUND", key="enable", value=False)

    def save_template(self):
        template_name = self.template_name_variable.get()
        columns_list = self.process_columns_df['COL_NAME'].tolist()

        if len(columns_list) != 0 and len(template_name) != 0:
            save_template(template_name=template_name, columns_list=columns_list)
            self.template_name_textbox.delete(0, "end")
        else:
            messagebox.showwarning(f'Save template Warning',
                                   "Template name or Columns list is empty...",
                                   icon='warning')
            print("### Error: Template name or Columns list is empty...")
        self.update_option_menu()

    def delete_template(self):

        template_name = self.templates_variable.get()

        delete_msg_box = messagebox.askquestion(f'Delete template Warning', f'Delete <{template_name}> template?',
                                                icon='warning')

        if delete_msg_box == 'yes':
            template_name = self.templates_variable.get()
            delete_template(template_name)
            self.update_option_menu()
            self.templates_variable.set('Default')

    def rewrite_template(self):
        template_name = self.templates_variable.get()
        columns_list = self.process_columns_df['COL_NAME'].tolist()

        rewrite_msb_box = messagebox.askquestion(f'Rewrite template Warning',
                                                 f"Rewrite <{template_name}> template?",
                                                 icon='warning')
        if rewrite_msb_box == 'yes':
            if len(columns_list) != 0:
                rewrite_template(template_name=template_name, columns_list=columns_list)
                messagebox.showwarning(f'Rewrite template Info',
                                       f"<{template_name}> template updated",
                                       icon='info')
                print('### Info: Шаблон обновлен')
            else:
                messagebox.showwarning(f'Rewrite template Waring',
                                       f"Columns list is Empty!",
                                       icon='warning')
                print("### Error: Список столбцов пуст")
            self.update_option_menu()

    def update_option_menu(self):
        """
        Рефреш Компобокса
        """
        self.templates_list = get_templates_list()
        templates_combobox = self.templates_combobox["menu"]
        templates_combobox.delete(0, "end")
        for string in self.templates_list:
            templates_combobox.add_command(label=string,
                                           command=lambda value=string: self.templates_variable.set(value))

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
                # Добавляем строку к ДФ
                self.process_columns_df.loc[len(self.process_columns_df.index)] = [col_index,
                                                                                   col_name]
                # Старя реализация - Append удаляется из Pandas
                # self.process_columns_df = self.process_columns_df.append({'COL_INDEX': col_index, 'COL_NAME': col_name},
                #                                                          ignore_index=True)

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

            self.total_columns_df = self.total_columns_filter_backup_df[
                self.total_columns_filter_backup_df.index.isin(filter_index_list)]

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

    def add_tree(self):
        # добавляем тестовые записи в Статы
        self.stat_tree.insert('', 'end', text=1, values=1, tags=('ANOM',))
        self.stat_tree.insert('', 'end', text=2, values=2, tags=('CON',))
        self.stat_tree.insert('', 'end', text=3, values=3, tags=('WELD',))
        self.stat_tree.insert('', 'end', text=4, values=4, tags=('ANOM',))
        self.stat_tree.insert('', 'end', text=5, values=5, tags=('CON',))

    def select_group_checks(self):

        """
        Выделяем чекбоксы по наличию в списке Тэгов
        """

        self.get_tags_exist()
        self.fill_tags_listbox()

        tag_name = self.stat_tree_tags_variable.get()

        if tag_name in self.tags_dict_name_id.keys():
            tag_id = self.tags_dict_name_id[tag_name]

            stat_tree = self.stat_tree
            tree_items = stat_tree.get_children()

            for tree_id in tree_items:
                self.stat_tree._uncheck_descendant(tree_id)
                self.stat_tree._uncheck_ancestor(tree_id)
                tag = stat_tree.item(tree_id)['tags'][0]
                if tag == tag_id:
                    self.stat_tree._check_ancestor(tree_id)
                    self.stat_tree._check_descendant(tree_id)

    def get_tags_exist(self):
        """
        Очищаем и заполняем Лист имеющихся Тэгов в дереве особенностей
        """
        self.stat_tree_tags_list = []
        stat_tree = self.stat_tree
        tree_items = stat_tree.get_children()
        for each in tree_items:
            tag = stat_tree.item(each)['tags'][0]
            if tag not in self.stat_tree_tags_list:
                self.stat_tree_tags_list.append(tag)

    def fill_tags_listbox(self):

        """
        Заполняем листбокс Тэгов из особенностей
        """

        self.get_tags_exist()
        templates_combobox = self.stat_tree_tags_combobox["menu"]
        templates_combobox.delete(0, "end")
        for tag in self.stat_tree_tags_list:
            if tag in self.tags_dict_id_name.keys():
                tag_name = self.tags_dict_id_name[tag]
                templates_combobox.add_command(label=tag_name,
                                               command=lambda value=tag_name: self.stat_tree_tags_variable.set(value))

    def select_all_checks(self):

        stat_tree = self.stat_tree
        # Список ID
        tree_items = stat_tree.get_children()

        # Бежим по Айдишникам и печатаем характеристики (хранятся в словаре)
        # for each in tree_items:
        #     print(stat_tree.item(each)['values'][0])  # e.g. prints data in clicked cell
        #     print(stat_tree.item(each)['tags'])

        for tree_id in tree_items:
            self.stat_tree._check_ancestor(tree_id)
            self.stat_tree._check_descendant(tree_id)

    def remove_all_checks(self):

        stat_tree = self.stat_tree
        # Список ID
        tree_items = stat_tree.get_children()

        for tree_id in tree_items:
            self.stat_tree._uncheck_descendant(tree_id)
            self.stat_tree._uncheck_ancestor(tree_id)

    def invert_checks(self):

        stat_tree = self.stat_tree
        # Список ID
        tree_items = stat_tree.get_children()
        tree_checked = stat_tree.get_checked()

        for tree_id in tree_items:
            if tree_id in tree_checked:
                self.stat_tree._uncheck_descendant(tree_id)
                self.stat_tree._uncheck_ancestor(tree_id)
            else:
                self.stat_tree._check_ancestor(tree_id)
                self.stat_tree._check_descendant(tree_id)

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

        total_columns_exist_sorted_df = self.df_dbf_class.parse_columns_df(columns_list=total_columns_exist_sorted,
                                                                           cross_list=total_columns_exist_sorted,
                                                                           ret_blank=False)

        total_columns_exist_sorted = total_columns_exist_sorted_df['COL_ID'].tolist()
        total_columns_exist_sorted_names = total_columns_exist_sorted_df['COL_NAME'].tolist()

        # total_columns_exist_sorted, total_columns_exist_sorted_names = self.df_dbf_class.parse_columns(
        #     columns_list=total_columns_exist_sorted, ret_blank=False)

        # чистим и заполняем DF общего число столбцов
        self.total_columns_df = self.total_columns_df.iloc[0:0]
        for i in range(len(total_columns_exist_sorted)):
            # Добавляем строку к ДФ
            self.total_columns_df.loc[len(self.total_columns_df.index)] = [total_columns_exist_sorted[i],
                                                                           total_columns_exist_sorted_names[i]]
            # Старя реализация - Append удаляется из Pandas
            # self.total_columns_df = self.total_columns_df.append(
            #     {'COL_INDEX': total_columns_exist_sorted[i], 'COL_NAME': total_columns_exist_sorted_names[i]},
            #     ignore_index=True)
        self.total_columns_filter_backup_df = self.total_columns_df

        for i in total_columns_exist_sorted_names:
            self.db_columns_listbox.insert('end', i)

    def load_template(self):

        self.update_option_menu()

        template_name = self.templates_variable.get()
        template_columns_json = read_template(template_name=template_name)

        try:
            total_columns_index = self.db_df.columns.values.tolist()

            template_columns_df = self.df_dbf_class.parse_columns_df(columns_list=template_columns_json,
                                                                     cross_list=total_columns_index)

            template_columns = template_columns_df['COL_ID'].tolist()
            template_columns_names = template_columns_df['COL_NAME'].tolist()

            # чистим и заполняем DF кастомного числа столбцов
            self.process_columns_df = self.process_columns_df.iloc[0:0]

            # если кастом не пустой то грузим его
            if len(template_columns) > 1:
                columns_raw = template_columns_json
                columns_names = template_columns_names

                for i in range(len(template_columns)):
                    # Добавляем строку к ДФ
                    self.process_columns_df.loc[len(self.process_columns_df.index)] = [template_columns[i],
                                                                                       template_columns_names[i]]
                    # Старя реализация - Append удаляется из Pandas
                    # self.process_columns_df = self.process_columns_df.append(
                    #     {'COL_INDEX': template_columns[i], 'COL_NAME': template_columns_names[i]},
                    #     ignore_index=True)

                self.custom_listbox_raw.delete(0, 'end')
                self.custom_listbox_processed.delete(0, 'end')

                for i in columns_raw:
                    self.custom_listbox_raw.insert('end', i)

                list_num = 0
                for i in columns_names:
                    self.custom_listbox_processed.insert('end', i)
                    # красим красным если нашли #BLANK
                    self.custom_listbox_processed.itemconfig("end",
                                                             bg="#ff7373" if i == '#BLANK' or i == "#NOT_IN_DB" else "white")
                    if len(template_columns) > 2 and i == '#BLANK':
                        self.custom_listbox_raw.itemconfig(list_num,
                                                           bg="#ff7373" if i == '#BLANK' or i == "#NOT_IN_DB" else "white")
                    list_num += 1

        except TypeError and AttributeError:

            messagebox.showwarning(f'Error', "DB not Loaded",
                                   icon='error')
            print('# ERROR: DB not Loaded!')

    def process_custom_columns(self):

        """
        Кнопка Process Columns
        Берем на вход строку, разбиваем по табулятор и находим соответствия
        """

        try:
            columns_raw = []
            columns_names = []
            custom_columns_textbox = self.columns_variable.get()[:-1]

            total_columns_index = self.db_df.columns.values.tolist()

            if custom_columns_textbox != '':
                custom_columns_names_raw = custom_columns_textbox.split('\t')
                # возвращаем столбец что нашли и их перевод для Кастом

                custom_columns_df = self.df_dbf_class.parse_columns_df(columns_list=custom_columns_names_raw,
                                                                       cross_list=total_columns_index)
                custom_columns = custom_columns_df['COL_ID'].tolist()
                custom_columns_names = custom_columns_df['COL_NAME'].tolist()

                # чистим и заполняем DF кастомного числа столбцов
                self.process_columns_df = self.process_columns_df.iloc[0:0]

                # если кастом не пустой то грузим его

                columns_raw = custom_columns_names_raw
                columns_names = custom_columns_names

                for i in range(len(custom_columns)):
                    # Новая реализация
                    self.process_columns_df.loc[len(self.process_columns_df.index)] = [custom_columns[i],
                                                                                       custom_columns_names[i]]
                    # Старя реализация - Append удаляется из Pandas
                    # self.process_columns_df = self.process_columns_df.append(
                    #     {'COL_INDEX': custom_columns[i], 'COL_NAME': custom_columns_names[i]}, ignore_index=True)

                self.custom_listbox_raw.delete(0, 'end')
                self.custom_listbox_processed.delete(0, 'end')

                for i in columns_raw:
                    self.custom_listbox_raw.insert('end', i)

                list_num = 0
                for i in columns_names:
                    self.custom_listbox_processed.insert('end', i)
                    # красим красным если нашли #BLANK
                    self.custom_listbox_processed.itemconfig("end",
                                                             bg="#ff7373" if i == '#BLANK' or i == "#NOT_IN_DB" else "white")
                    if len(custom_columns) > 2 and i == '#BLANK':
                        self.custom_listbox_raw.itemconfig(list_num,
                                                           bg="#ff7373" if i == '#BLANK' or i == "#NOT_IN_DB" else "white")
                    list_num += 1
            else:
                messagebox.showwarning(f'Custom columns Warning', "Custom columns field is empty",
                                       icon='warning')
                print("### Info: Custom columns field is empty..")
        except TypeError and AttributeError:
            messagebox.showwarning(f'Error', "DB not Loaded",
                                   icon='error')
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
            print("LOG Error")

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

                # --------------------------------
                # получаем статистику по фичам
                stat_ser = self.db_df['#FEATURE'].value_counts(dropna=False, normalize=False)
                # серия с цветами
                color_ser = self.df_dbf_class.get_color_type_df(lng=self.lng)
                # конвертим статистику в ДФ
                stat_df = stat_ser.to_frame()
                # столбец с числами переименовываем в index_stat
                stat_df = stat_df.rename(columns={'#FEATURE': "index_count"})
                # переносим индексы в столбец (index)
                stat_df = stat_df.reset_index()
                # создаем столбец с Тэгами
                stat_df['tag'] = ''

                # бежим по статистике и записываем тэги
                for i, row in stat_df.iterrows():
                    cur_value = stat_df.loc[i]['index']
                    for j, row in color_ser.iterrows():
                        if cur_value == color_ser.loc[j][1]:
                            stat_df.at[i, 'tag'] = color_ser.loc[j][0]

                # сортируем по Тэгам потом по Количество
                stat_df.sort_values(['tag', 'index_count'], ascending=[True, False], inplace=True)

                # --------------------------------

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
                for i, row in stat_df.iterrows():
                    feature = stat_df.loc[i]['index']
                    feature_count = stat_df.loc[i]['index_count']
                    feature_tag = stat_df.loc[i]['tag']
                    self.stat_tree.insert('', 'end', text=feature, values=feature_count, tags=(feature_tag,))

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

                # заполняем список тэгов
                self.fill_tags_listbox()

                messagebox.showwarning(f'DB Load Info', f"DB Loaded Successfully, total records: {len(self.db_df)}",
                                       icon='info')
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
            messagebox.showwarning(f'DB path error', "DB path empty or not correct",
                                   icon='warning')
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

                    # messagebox.showwarning(f'Clipboard Export', "Copied to the clipboard.",
                    #                        icon='info')
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
                        messagebox.showwarning('CSV Export', f"Data saved to the <{basename[:-4]}.csv> file.",
                                               icon='info')
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

                messagebox.showwarning(f'Export Error', "Columns for Export not selected...",
                                       icon='warning')
                print('# Info: Export table is empty')


if __name__ == "__main__":

    arg = ''

    if len(sys.argv[1:]) == 2:
        run_atr = sys.argv[1:][0]
        path = sys.argv[1:][1]

        if run_atr == "-D":
            export_default(dbf_path=path)
            input("~~~ Done~~~")
        else:
            input("### Error: Wrong Attribute")

    else:
        DB_FORM()
