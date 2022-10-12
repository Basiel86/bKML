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
from parse_templates import *
from play_sound import PlaySound
from parse_cfg import CFG
from TkinterDnD2 import *
import DBtoKML
from other_functions import cls_pyfiglet
import logging.handlers
from logger_init import init_logger

import telebot

pd.options.mode.chained_assignment = None  # default='warn'

DEBUG = False

logger = logging.getLogger('app.DB_Process')


class DB_FORM:

    def __init__(self):

        self.EXP_DAY = '2022-10-22'

        cls_pyfiglet('DB Process', 'larry3d')
        self.cfg = CFG('DB Process')
        self.log_path = os.path.join(self.cfg.get_cfg_project_folder(), 'DB_Process.log')
        init_logger('app', log_pth=resource_path(self.log_path), print_console=DEBUG)

        # bot.polling(none_stop=True, interval=0)

        # self.updater_name = 'DB Process'
        # self.updater = BUpdater(self.updater_name)
        # z = self.updater.pre_update_check()
        # print(z)

        self.lang_list = ["RU", "EN"]
        self.db_ext_list = ('dbf', 'DBF')
        self.kml_line_width_lst = [1, 3, 5]
        self.db_path = ''

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
        if self.cfg.read_cfg(section="SOUND", key="enable", default=False) is True:
            self.play_sound()

        # Список Тэгов в текущем списке особенностей
        self.stat_tree_tags_list = []

        self.inch_names_list = parse_inch.get_inch_names_list()
        self.inch_list = parse_inch.get_inch_list()
        self.inch_dict = parse_inch.get_inch_dict()
        self.templates_list = get_templates_list()
        # self.diam_list = [4, 4.5, 5.563, 6.625, 8.625, 10.75, 12.75, 14, 16, 18, 20, 22, 24, 26, 28, 30, 32, 34, 36, 38,
        #                   40, 42, 44, 46, 48, 52, 56]

        index_cfg_path = self.cfg.read_cfg(section="PATHS", key="index_custom", create_if_none=True)
        struct_cfg_path = self.cfg.read_cfg(section="PATHS", key="struct_custom", create_if_none=True)
        index_remote_path = self.cfg.read_cfg(section="PATHS", key="index_share", create_if_none=True,
                                              default=r'\\vasilypc\Vasily Shared (Read Only)\_Templates\PT\IDs\DBF_INDEX.xlsx')
        struct_remote_path = self.cfg.read_cfg(section="PATHS", key="struct_share", create_if_none=True,
                                               default=r'\\vasilypc\Vasily Shared (Read Only)\_Templates\PT\IDs\STRUCT.xlsx')

        local = self.cfg.read_cfg(section="PATHS", key="local", create_if_none=True, default=True)
        self.df_dbf_class = df_DBF(index_remote_path=index_remote_path,
                                   index_path=index_cfg_path,
                                   struct_remote_path=struct_remote_path,
                                   struct_path=struct_cfg_path,
                                   local=local)
        self.db_df = pd.DataFrame

        self.db_process_form = TkinterDnD.Tk()

        self.db_process_form.iconbitmap(resource_path('icons/DBF_icon.ico'))
        self.db_process_form.drop_target_register(DND_FILES)
        self.db_process_form.dnd_bind('<<Drop>>', self.open_with_dnd)

        self.menu_main = Menu(self.db_process_form)
        self.db_process_form.config(menu=self.menu_main)

        file_menu = Menu(self.menu_main, tearoff=0)
        file_menu.add_command(label="Open...", command=self.openfile)
        file_menu.add_separator()
        file_menu.add_command(label="Load DB...", command=lambda: self.db_load(cls=True))
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.on_closing)
        self.menu_main.add_cascade(label="File",
                                   menu=file_menu)

        # groove, raised, ridge, solid, or sunken
        self.db_file_label = Label(self.db_process_form, text="-DB Name-", font=("Verdana", 12),
                                   justify='center', foreground='red', relief='groove', padx=5, pady=3)

        self.template_menu = Menu(self.menu_main, tearoff=0)
        self.templates_radio_variable = StringVar()
        self.update_menu_templates()
        self.menu_main.add_cascade(label="Templates",
                                   menu=self.template_menu)

        others_menu = Menu(self.menu_main, tearoff=0)
        others_menu.add_command(label="Open Config Folder", command=parse_templates.open_templates_folder)
        others_menu.add_command(label="Open Log", command=lambda: os.startfile(self.log_path))
        others_menu.add_command(label="Write Developer MSG", command=self.write_to_developer)
        others_menu.add_command(label="Save Tab Position",
                                command=lambda: self.cfg.store_settings('GENERAL', 'default_tab',
                                                                        self.tabs_main_func.index(
                                                                            self.tabs_main_func.select()) + 1))
        self.menu_main.add_cascade(label="Configs", menu=others_menu)

        lang_menu = Menu(self.menu_main, tearoff=0)
        self.lang_menu_variable = StringVar()
        self.lang_menu_variable.set(self.lang_list[0])
        for lang in self.lang_list:
            lang_menu.add_radiobutton(label=lang, var=self.lang_menu_variable)
        self.menu_main.add_cascade(label="Language",
                                   menu=lang_menu)

        diam_menu = Menu(self.menu_main, tearoff=0)
        self.diam_menu_variable = StringVar()
        self.diam_menu_variable.trace('w', lambda *args: self.diam_menu_change())

        for inch in self.inch_names_list:
            diam_menu.add_radiobutton(label=inch, var=self.diam_menu_variable, command=self.diam_menu_change)
        self.menu_main.add_cascade(label="Diameter", menu=diam_menu)
        self.diam_menu_change()

        self.templates_list = get_templates_list()

        # ///////////////////////////////  TAB FEA

        self.tabs_main_fea = ttk.Notebook(self.db_process_form)
        self.tab1_fea = ttk.Frame(self.tabs_main_fea)
        self.tabs_main_fea.add(self.tab1_fea, text='Fea')

        self.export_frame = ttk.Labelframe(master=self.tab1_fea, text='Export')  # , width=50, height=50
        self.clipboard_db_button = Button(self.export_frame, text="to Clip", width=10,
                                          command=lambda: self.db_export(to_clipboard=True))
        self.csv_export_button = Button(self.export_frame, text="to CSV", width=10,
                                        command=lambda: self.db_export(ext='csv'))
        self.xlsx_export_button = Button(self.export_frame, text="to XLSX", width=10,
                                         command=lambda: self.db_export(ext='xlsx'))
        self.KML_btn_fea_tab = Button(self.export_frame, text="to KML", command=self.DB_to_KML)

        self.and_or_var = StringVar(self.tab1_fea)
        self.and_or_comb = OptionMenu(self.tab1_fea, self.and_or_var, *('AND', 'OR'))
        self.and_or_var.set('AND')

        # высота в строках
        self.stat_tree = CheckboxTreeview(self.tab1_fea, show='tree', column="c1", height=24)
        self.stat_tree.column("# 1", anchor='center', stretch='NO', width=60)

        self.stat_tree_tags_variable = StringVar(self.tab1_fea)
        self.stat_tree_tags_variable.trace(mode='w', callback=self.stat_tree_tags_option_change)
        self.stat_tree_tags_combobox = OptionMenu(self.tab1_fea, self.stat_tree_tags_variable,
                                                  self.stat_tree_tags_list)
        self.sel_all_checks_btn = Button(self.tab1_fea, text="Select All", command=self.select_all_checks)
        self.remove_all_checks_btn = Button(self.tab1_fea, text="Clear All", command=self.remove_all_checks)
        self.invert_checks_btn = Button(self.tab1_fea, text="Invert", command=self.invert_checks)

        self.doc_tree = CheckboxTreeview(self.tab1_fea, show='tree', column="c1", height=5)
        self.doc_tree.column("# 0", anchor='center', stretch='YES', width=30)
        self.doc_tree.column("# 1", anchor='center', stretch='NO', width=60)

        self.marker_tree = CheckboxTreeview(self.tab1_fea, show='tree', column="c1", height=3)
        self.marker_tree.column("# 0", anchor='center', stretch='YES', width=30)
        self.marker_tree.column("# 1", anchor='center', stretch='NO', width=60)

        style = ttk.Style(self.tab1_fea)
        # remove the indicator in the treeview
        style.layout('Checkbox.Treeview.Item',
                     [('Treeitem.padding',
                       {'sticky': 'nswe',
                        'children': [('Treeitem.image', {'side': 'left', 'sticky': ''}),
                                     ('Treeitem.focus', {'side': 'left', 'sticky': '',
                                                         'children': [('Treeitem.text',
                                                                       {'side': 'left', 'sticky': ''})]})]})])
        style.configure('Checkbox.Treeview', borderwidth=1, relief='sunken')

        # ///////////////////////////////  TAB FUNC

        self.tabs_main_func = ttk.Notebook(self.db_process_form)
        self.tab1_func = ttk.Frame(self.tabs_main_func)
        self.tab2_func = ttk.Frame(self.tabs_main_func)
        self.tab3_func = ttk.Frame(self.tabs_main_func)
        self.tab4_func = ttk.Frame(self.tabs_main_func)
        self.tabs_main_func.add(self.tab1_func, text='Columns')
        self.tabs_main_func.add(self.tab2_func, text='Helper')
        self.tabs_main_func.add(self.tab3_func, text='KML')
        self.tabs_main_func.add(self.tab4_func, text='Other')
        self.tabs_main_func.select(self.get_def_tab())

        self.columns_variable = StringVar(self.tab1_func)
        self.columns_textbox = Entry(self.tab1_func, textvariable=self.columns_variable)
        self.columns_textbox.bind("<Return>", self.process_custom_columns)
        self.clear_headers_button = Button(self.tab1_func, text="Clear", command=self.clear_headers)

        self.custom_listbox_raw = Listbox(master=self.tab1_func, width=30, height=33, exportselection=False)
        self.custom_listbox_processed = Listbox(master=self.tab1_func, width=30, height=33, exportselection=False)
        self.db_columns_listbox = Listbox(master=self.tab1_func, width=35, height=33, exportselection=False)
        self.scroll_pane = ttk.Scrollbar(master=self.tab1_func, command=self.db_columns_listbox.yview)
        self.db_columns_listbox.configure(yscrollcommand=self.scroll_pane.set)

        image_up = PhotoImage(file=resource_path('icons/arrrow_up.png'))
        image_dwn = PhotoImage(file=resource_path('icons/arrrow_dwn.png'))
        self.process_lst_row_up_btn = Button(master=self.tab1_func, image=image_up,
                                             command=self.custom_listbox_processed_row_up, width=22)
        self.process_lst_row_dwn_btn = Button(master=self.tab1_func, image=image_dwn,
                                              command=self.custom_listbox_processed_row_dwn, width=22)

        self.clear_process_list_button = Button(master=self.tab1_func, text="Clear", command=self.clear_process_listbox)
        self.filter_variable = StringVar()
        self.filter_variable.trace("w", self.filter_total_columns_listbox)
        self.filter_total_columns_textbox = Entry(master=self.tab1_func, width=35,
                                                  textvariable=self.filter_variable)

        self.play_button = Button(self.tab4_func, text="Play", command=self.play_sound,
                                  relief=GROOVE)
        self.play_button_stop = Button(self.tab4_func, text="Stop", command=self.stop_sound,
                                       relief=GROOVE)

        self.KML_line_width_lbl = Label(self.tab3_func, text="Line Width")
        self.KML_line_width_var = StringVar(self.tab3_func)
        self.KML_line_width_comb = OptionMenu(self.tab3_func, self.KML_line_width_var, *self.kml_line_width_lst)

        # //////////////////////////////////////////////////////////////////////////////////////////

        self.process_columns_df = pd.DataFrame(columns=['COL_INDEX', 'COL_NAME'])
        self.total_columns_df = pd.DataFrame(columns=['COL_INDEX', 'COL_NAME'])
        self.total_columns_filter_backup_df = pd.DataFrame(columns=['COL_INDEX', 'COL_NAME'])

        # Инициализация Базы
        self.lng = ''
        # self.df_dbf_class = df_DBF()

        self.form_init()

    def form_init(self):

        # self.add_tree()
        # читаем аргументы на входе
        arg = ''
        for arg in sys.argv[1:]:
            if os.path.exists(arg):
                print('Drag&Drop file: ', arg)

        if DEBUG is True:
            arg = r'd:\OneDrive\Macro\PYTHON\bKML\Test\2nwfm.DBF'

        exp_date_formatted = datetime.strptime(self.EXP_DAY, "%Y-%m-%d").date()
        now_date = date.today()
        days_left = exp_date_formatted - now_date

        ico_abs_path = resource_path('icons\DBF_icon.ico')
        self.db_process_form.wm_iconbitmap(ico_abs_path)

        if exp_date_formatted >= now_date:

            self.db_process_form.title("DB process")
            width = 990
            height = 600
            self.db_process_form.geometry(f"{width}x{height}")
            self.db_process_form.minsize(width, height)
            self.db_process_form.maxsize(width, height)

            self.db_file_label.grid(column=0, row=39, sticky=EW, padx=5)

            # TABS FEA
            self.tabs_main_fea.grid(row=0, column=0, columnspan=1, rowspan=35)

            # TAB FEA 1

            self.stat_tree.grid(row=2, column=0, rowspan=30, columnspan=3, padx=5, sticky=N)
            self.stat_tree.bind("<Double-1>", self.OnDoubleClick)
            self.doc_tree.grid(row=2, column=3, rowspan=4, sticky=N)
            self.marker_tree.grid(row=6, column=3, rowspan=4, sticky=NS)
            self.stat_tree_tags_combobox.grid(row=1, column=0, padx=5, columnspan=3, sticky=EW)

            self.sel_all_checks_btn.grid(row=0, column=0, sticky=EW, padx=3)
            self.remove_all_checks_btn.grid(row=0, column=1, sticky=EW, padx=3)
            self.invert_checks_btn.grid(row=0, column=2, sticky=EW, padx=3)

            self.and_or_comb.grid(row=1, column=3, sticky=EW, columnspan=1)

            # EXPORT FRAME
            self.export_frame.grid(row=15, column=3, sticky=EW)
            self.clipboard_db_button.grid(row=0, column=0, sticky=EW)
            self.csv_export_button.grid(row=2, column=0, sticky=EW)
            self.xlsx_export_button.grid(row=3, column=0, sticky=EW)
            self.KML_btn_fea_tab.grid(row=5, column=0, sticky=EW, pady=5)
            self.export_frame.grid_rowconfigure(1, minsize=25)
            self.export_frame.grid_rowconfigure(4, minsize=25)

            # TABS RIGHT
            self.tabs_main_func.grid(row=0, column=1, columnspan=1, rowspan=40)
            # TAB 1

            self.clear_headers_button.grid(row=0, column=0, sticky=EW, padx=2, columnspan=1)
            self.columns_textbox.grid(row=0, column=1, columnspan=1, sticky=EW)

            self.custom_listbox_raw.grid(row=1, column=0, rowspan=40, columnspan=2, padx=3, sticky=E)
            self.custom_listbox_processed.grid(row=1, column=3, rowspan=40, columnspan=6)
            self.db_columns_listbox.grid(row=1, column=9, rowspan=40, columnspan=3)

            self.clear_process_list_button.grid(row=0, column=6, columnspan=3, sticky=EW)
            self.process_lst_row_up_btn.grid(row=0, column=3, columnspan=1, sticky=EW, padx=1)
            self.process_lst_row_dwn_btn.grid(row=0, column=4, columnspan=1, sticky=EW, padx=1)

            self.filter_total_columns_textbox.grid(row=0, column=9, columnspan=2, padx=2)
            self.scroll_pane.grid(row=1, column=12, rowspan=40, columnspan=1, sticky=NS)
            # TAB 3
            self.KML_line_width_lbl.grid(row=0, column=0, columnspan=1)
            self.KML_line_width_comb.grid(row=0, column=1, columnspan=1)
            # TAB 4
            self.play_button.grid(row=0, column=0, columnspan=1)
            self.play_button_stop.grid(row=0, column=1, columnspan=1)
            # ----------------

            col_count, row_count = self.db_process_form.grid_size()
            for col in range(col_count):
                self.db_process_form.grid_columnconfigure(col, minsize=20)

            # for row in range(row_count):
            #     self.db_process_form.grid_rowconfigure(row, minsize=20)

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

            cfg_line_width = self.cfg.read_cfg(section="KML", key='line_width')
            if cfg_line_width not in self.kml_line_width_lst:
                self.KML_line_width_var.set(3)
            else:
                self.KML_line_width_var.set(cfg_line_width)

            if arg != '' and os.path.exists(arg):
                # пишем аргумент в строку пути
                self.db_path = arg
                # парсим инч
                arg_inch = parse_inch_prj(arg)

                # если нашли инч то селектим его в выпадающем списке
                if arg_inch is not None:
                    inch_index = self.inch_list.index(arg_inch)
                    self.diam_menu_variable.set(self.inch_names_list[inch_index])
                    self.db_load()
                else:
                    self.diam_menu_variable.set(self.inch_names_list[6])
                    self.db_file_label.config(text=os.path.basename(self.db_path))
            else:
                self.diam_menu_variable.set(self.inch_names_list[6])

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

    def write_to_developer(self):
        msg = self.columns_variable.get()
        if msg != '':
            self.send_telegram_msg(msg=msg)
            with open(self.log_path, mode='r') as logfile:
                is_sent = self.send_telegram_msg(file=logfile)
                self.clear_headers()
                if is_sent is False:
                    messagebox.showerror(f'Send Message Info', f"Message sending error!", icon='error')
                    return

                messagebox.showinfo(f'Send Message Info', f"Message: <{msg}> sent successfully!", icon='info')

        else:
            messagebox.showwarning(f'Send Message Info',
                                   "Write some message on Custom Columns Textbox",
                                   icon='info')

    @staticmethod
    def send_telegram_msg(msg=None, file=None) -> bool:

        baz_id = 26805602
        token = '5650874524:AAFi82ooRhLt1KOjn0GKTXFB6WA-30XN9Zs'
        username = os.getenv('username')
        pc_name = os.environ['COMPUTERNAME']
        try:
            bot = telebot.TeleBot(token)
            if msg is not None:
                bot.send_message(baz_id, f' User: {username} / PC: {pc_name} - {msg}')
            if file is not None:
                # bot.send_message(baz_id, f' User: {username} / PC: {pc_name} - LOG FILE')
                bot.send_document(baz_id, file)

            return True

        except Exception as ex:
            logger.error('TLG msg error')
            return False

    def get_def_tab(self):
        def_tab = self.cfg.read_cfg(section="GENERAL", key='default_tab', create_if_none=True, default=1)

        if def_tab is None:
            return 0

        if def_tab in (1, 2, 3, 4):
            return def_tab - 1

        self.cfg.store_settings(section="GENERAL", key='default_tab', value=1)

    # def updater_run(self, exe_name):
    #     args = self.updater_name
    #     filename = 'B_Updater.exe'
    #     subprocess.run([filename, args])

    def play_sound(self):
        logger.debug("Sound ON")
        self.sound.play_music()
        self.cfg.store_settings(section="SOUND", key="enable", value=True)

    def stop_sound(self):
        logger.debug("Sound OFF")
        self.sound.stop_music()
        self.cfg.store_settings(section="SOUND", key="enable", value=False)

    def save_template(self):
        template_name = self.columns_textbox.get()
        if len(template_name) > 30:
            messagebox.showwarning(f'Save template Warning',
                                   "Template name is too long (max 30 symbols)...",
                                   icon='info')
            logger.warning(f"Template name is too long (max 30 symbols): "
                           f"{template_name} "
                           f"len={len(template_name)}")
            return None

        columns_list = self.process_columns_df['COL_NAME'].tolist()

        if len(columns_list) != 0 and len(template_name) != 0:
            save_template(template_name=template_name, columns_list=columns_list)
            self.columns_textbox.delete(0, "end")
        else:
            messagebox.showwarning(f'Save template Warning',
                                   "Template name or Columns list is empty...",
                                   icon='warning')
            print("### Error: Template name or Columns list is empty...")
            logger.error('Template name or Columns list is empty')
        self.update_menu_templates()

    def delete_template(self):

        template_name = self.templates_radio_variable.get()

        if template_name == '':
            return None

        delete_msg_box = messagebox.askquestion(f'Delete template Warning', f'Delete <{template_name}> template?',
                                                icon='warning')

        if delete_msg_box == 'yes':
            delete_template(template_name)
            self.update_menu_templates()
            self.templates_radio_variable.set('Default')

    def rewrite_template(self):
        template_name = self.templates_radio_variable.get()
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
                print('### Info: Template updated')
                logger.info(f'Rewrite template Info <{template_name}> template updated')
            else:
                messagebox.showwarning(f'Rewrite template Waring',
                                       f"Columns list is Empty!",
                                       icon='warning')
                print("### Error: Список столбцов пуст")
                logger.info('Columns list is empty')
            self.update_menu_templates()

    def diam_menu_change(self):
        diam = self.diam_menu_variable.get()
        self.menu_main.entryconfig(5, label=diam)

    def update_menu_templates(self):

        self.template_menu.delete(0, 'end')

        self.templates_list = get_templates_list()
        for template in self.templates_list:
            self.template_menu.add_radiobutton(label=template, var=self.templates_radio_variable,
                                               command=self.load_template)
        self.template_menu.add_separator()
        self.template_menu.add_command(label="Save Template", command=self.save_template)
        self.template_menu.add_command(label="Delete Template", command=self.delete_template)

    # def update_option_menu(self):
    #     """
    #     Рефреш Компобокса
    #     """
    #     self.templates_list = get_templates_list()
    #     templates_combobox = self.templates_combobox["menu_main"]
    #     templates_combobox.delete(0, "end")
    #     for string in self.templates_list:
    #         templates_combobox.add_command(label=string,
    #                                        command=lambda value=string: self.templates_variable.set(value))

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
            self.custom_listbox_processed_mark_red()

    def custom_listbox_processed_delete(self, event):
        if self.custom_listbox_processed.curselection() != ():
            selection_id = self.custom_listbox_processed.curselection()[0]
            self.custom_listbox_processed.delete(selection_id)
            self.custom_listbox_processed.selection_set(selection_id)

            # грохаем строку
            self.process_columns_df = self.process_columns_df.drop([selection_id])
            # перенумеровываем индексы
            self.process_columns_df = self.process_columns_df.reset_index(drop=True)

            self.custom_listbox_processed_mark_red()

    def custom_listbox_processed_replace(self, event):
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

            self.custom_listbox_processed_mark_red()

    def custom_listbox_processed_row_up(self):
        # Если не пустой список под Экспорт
        if self.custom_listbox_processed.curselection() != ():
            # текущий выбранный индекс
            selection_id = self.custom_listbox_processed.curselection()[0]
            # имя текущего
            tmp_val = self.custom_listbox_processed.get(selection_id)

            # пока индекс большу нуля (чтоб не вылететь за out of bounds)
            if selection_id > 0:
                # удаляем текущую строчку
                self.custom_listbox_processed.delete(selection_id)
                # вставляем на 1ну выше
                self.custom_listbox_processed.insert(selection_id - 1, tmp_val)
                # выделяем на 1ну выше
                self.custom_listbox_processed.selection_set(selection_id - 1)

                # свопаем датафрэйм значений
                b, c = self.process_columns_df.iloc[selection_id].copy(), self.process_columns_df.iloc[
                    selection_id - 1].copy()
                self.process_columns_df.iloc[selection_id], self.process_columns_df.iloc[selection_id - 1] = c, b

                # ресетим индекс
                self.process_columns_df = self.process_columns_df.reset_index(drop=True)

                self.custom_listbox_processed_mark_red()

    def custom_listbox_processed_row_dwn(self):
        if self.custom_listbox_processed.curselection() != ():
            selection_id = self.custom_listbox_processed.curselection()[0]
            tmp_val = self.custom_listbox_processed.get(selection_id)

            if selection_id < len(self.process_columns_df) - 1:
                self.custom_listbox_processed.delete(selection_id)
                self.custom_listbox_processed.insert(selection_id + 1, tmp_val)
                self.custom_listbox_processed.selection_set(selection_id + 1)

                b, c = self.process_columns_df.iloc[selection_id].copy(), self.process_columns_df.iloc[
                    selection_id + 1].copy()
                self.process_columns_df.iloc[selection_id], self.process_columns_df.iloc[selection_id + 1] = c, b

                self.process_columns_df = self.process_columns_df.reset_index(drop=True)
                self.custom_listbox_processed_mark_red()

    def custom_listbox_processed_mark_red(self):
        for i in range(len(self.process_columns_df)):
            list_val = self.custom_listbox_processed.get(i)
            self.custom_listbox_processed.itemconfig(i,
                                                     bg="#ff7373" if list_val == '#BLANK' or list_val == "#NOT_IN_DB" else "white")

    def DB_to_KML(self):

        if self.db_file_label['text'] == "-DB Name-":
            messagebox.showwarning(f'KML Export Info', "Nothing to export to KML \U0001F625",
                                   icon='info')
            logger.info(f'KML Export Info: Nothing to export to KML')
            return None
        try:

            df_for_kml = self.get_checked_df()
            if df_for_kml is None:
                messagebox.showwarning(f'KML Export Info', "Nothing to export to KML \U0001F625",
                                       icon='info')
                logger.info(f'KML Export Info: Nothing to export to KML')
                return None

            line_width = self.KML_line_width_var.get()
            kml_class = DBtoKML.bKML()
            kml_path = kml_class.dbf_to_kml(line_width=line_width, df=df_for_kml, df_dbf_class=self.df_dbf_class,
                                            export_path=self.db_path)
            if kml_path is None:
                messagebox.showwarning(f'KML Info', "No features with coordinates",
                                       icon='info')
                logger.error("KML, No features with coordinates")
            else:
                self.ask_open_file(kml_path)

        except TypeError and AttributeError:

            # messagebox.showwarning(f'Error', "DB not Loaded",
            #                        icon='error')
            print('# ERROR: DB not Loaded!')
            logger.error('DB not Loaded')

    @staticmethod
    def ask_open_file(file_path):
        if os.path.exists(file_path):
            f_name = os.path.basename(file_path)
            of = messagebox.askquestion(
                'Open file?', f'File <{f_name}> is created.\n\t'
                              f'Open file?', icon='info')
            logger.info(f'File saved: {file_path}')
            if of == 'yes':
                os.startfile(file_path)

    def add_tree(self):
        # добавляем тестовые записи в Статы
        self.stat_tree.insert('', 'end', text=1, values=1, tags=('ANOM',))
        self.stat_tree.insert('', 'end', text=2, values=2, tags=('CON',))
        self.stat_tree.insert('', 'end', text=3, values=3, tags=('WELD',))
        self.stat_tree.insert('', 'end', text=4, values=4, tags=('ANOM',))
        self.stat_tree.insert('', 'end', text=5, values=5, tags=('CON',))

    def stat_tree_tags_option_change(self, *args):
        self.select_group_checks()

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
        self.stat_tree_tags_variable.set('')

        # Бежим по Айдишникам и печатаем характеристики (хранятся в словаре)
        # for each in tree_items:
        #     print(stat_tree.item(each)['values'][0])  # e.g. prints data in clicked cell
        #     print(stat_tree.item(each)['tags'])

        for tree_id in tree_items:
            self.stat_tree._check_ancestor(tree_id)
            self.stat_tree._check_descendant(tree_id)

    def remove_all_checks(self):

        self.stat_tree_tags_variable.set('')

        trees = (self.stat_tree, self.doc_tree, self.marker_tree)

        for tree in trees:
            doc_items = tree.get_children()
            for doc_id in doc_items:
                tree._uncheck_descendant(doc_id)
                tree._uncheck_ancestor(doc_id)

    def invert_checks(self):

        stat_tree = self.stat_tree
        # Список ID
        tree_items = stat_tree.get_children()
        tree_checked = stat_tree.get_checked()
        self.stat_tree_tags_variable.set('')

        for tree_id in tree_items:
            if tree_id in tree_checked:
                self.stat_tree._uncheck_descendant(tree_id)
                self.stat_tree._uncheck_ancestor(tree_id)
            else:
                self.stat_tree._check_ancestor(tree_id)
                self.stat_tree._check_descendant(tree_id)

    def on_closing(self):
        self.cfg.store_settings(section="KML", key='line_width', value=self.KML_line_width_var.get())
        self.db_process_form.destroy()

        # with open(self.log_path) as logfile:
        #     self.send_telegram_msg(file=logfile)

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

    def load_template(self, *args):

        self.update_menu_templates()

        template_name = self.templates_radio_variable.get()
        template_columns_json = read_template(template_name=template_name)

        if self.db_file_label['text'] == "-DB Name-":
            return None
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

            # messagebox.showwarning(f'Error', "DB not Loaded",
            #                        icon='error')
            print('# ERROR: DB not Loaded!')
            logger.error('DB not Loaded')

    def process_custom_columns(self, events=None):

        """
        Process Custom Columns
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
                logger.info("Custom columns field is empty")
        except TypeError and AttributeError:
            messagebox.showwarning(f'Error', "DB not Loaded",
                                   icon='error')
            print('# ERROR: DB not Loaded!')
            logger.error('DB not Loaded')

    def OnDoubleClick(self, event):

        # активное
        # item = self.stat_tree.selection()[0]
        # print("you clicked on", self.stat_tree.item(item, "text"))

        # пробежка по чекнутым
        for item in self.stat_tree.get_checked():
            item_text = self.stat_tree.item(item, "text")
            # print(item_text)

    @staticmethod
    def write_log1(file_path=""):
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
            logger.error('LOG Error')

    @staticmethod
    def get_fea_color_type(fea_name, color_df, lng):
        for i, row in color_df.iterrows():
            df_fea_name = str(color_df.loc[i][f"FEA_{lng}"])
            if fea_name == df_fea_name:
                return str(color_df.loc[i][f"COLOR_TYPE"])

    # грузим базу
    def db_load(self, cls=False):

        if cls is True:
            cls_pyfiglet('DB Process', 'larry3d')
        self.lng = str(self.lang_menu_variable.get())
        diam_list_value = self.diam_menu_variable.get()
        diameter = self.inch_dict[diam_list_value]

        # send_bot_msg(f'# db_load: Opened path: {self.db_path}')
        if self.db_path != "" and self.db_path[-3:] in self.db_ext_list:

            self.db_file_label.config(text=os.path.basename(self.db_path))

            try:
                self.db_df = self.df_dbf_class.convert_dbf(diameter=diameter, dbf_path=self.db_path, lang=self.lng)

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
                # if len(stat_df) > 20:
                #     self.stat_tree.config(height=len(stat_df))

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
                logger.info(f'DB Load Info: DB Loaded Successfully, total records: {len(self.db_df)}')
                # self.process_custom_columns()

                try:
                    pass
                    # self.write_log(export_path=db_path)
                except Exception as logex:
                    print("LOG Error")
                    logger.error('LOG Error')

            except Exception as ex:
                print(ex)
                print("DB:form: Что-то пошло не так...")
                print(traceback.format_exc())
                logger.exception('Что-то пошло не так...')
        else:
            messagebox.showwarning(f'DB path error', "DB path is not correct",
                                   icon='warning')
            print("Путь не верен! Ты сбился с пути!?")
            logger.error(f'Path is wrong: {self.db_path}')

    def openfile(self):

        last_path = os.path.dirname(self.db_path)

        file_path = fd.askopenfilename(
            title='Open a file',
            # initialdir='/',
            initialdir=last_path,
            filetypes=[("DBF files", ".dbf .DBF")])

        cls_pyfiglet('DB Process', 'larry3d')

        self.db_path = file_path
        path_inch = parse_inch_prj(file_path)
        logger.info(f'Open file dialog path: {self.db_path}')
        # если нашли инч то селектим его в выпадающем списке
        if path_inch is not None:
            inch_index = self.inch_list.index(path_inch)
            self.diam_menu_variable.set(self.inch_names_list[inch_index])
            self.db_load()

    def open_with_dnd(self, event):

        if event.data[0] == '{':
            dnd_path = event.data[1:-1]
        else:
            dnd_path = event.data

        cls_pyfiglet('DB Process', 'larry3d')
        print("Drag-n-Drop Path: ", dnd_path)
        logger.info(f'Drag-n-Drop Path: {dnd_path}')

        if dnd_path.lower().endswith(self.db_ext_list):

            self.db_path = dnd_path

            path_inch = parse_inch_prj(dnd_path)
            # если нашли инч то селектим его в выпадающем списке
            if path_inch is not None:
                inch_index = self.inch_list.index(path_inch)
                self.diam_menu_variable.set(self.inch_names_list[inch_index])
                self.db_load()
            self.db_file_label.config(text=os.path.basename(self.db_path))

    @staticmethod
    def resource_path(relative_path):
        try:
            # PyInstaller creates a temp folder and stores path in _MEIPASS
            base_path = sys._MEIPASS
        except Exception as ex:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)

    def get_checked_df(self):

        """
        Возвращаем DF чекнутых фич
        :return: None если пусто
        """

        # DF для перефильтровки и на экспорт
        checked_df = None

        AND_OR = self.and_or_var.get()

        # заполняем чекнутые
        checked_items_fea = []
        checked_items_doc = []
        checked_items_marker = []
        for item in self.stat_tree.get_checked():
            checked_items_fea.append(self.stat_tree.item(item, "text"))
        for item in self.doc_tree.get_checked():
            checked_items_doc.append(self.doc_tree.item(item, "text"))
        for item in self.marker_tree.get_checked():
            checked_items_marker.append(self.marker_tree.item(item, "text"))

        # фильтруем по очереди по каждому из чекнутых
        if len(checked_items_fea) != 0:
            checked_df = self.db_df[self.db_df['#FEATURE'].isin(checked_items_fea)]

        if len(checked_items_doc) != 0:
            if checked_df is None:
                checked_df = self.db_df[self.db_df['#DOC'].isin(checked_items_doc)]
            # AND
            elif AND_OR == "AND":
                checked_df = checked_df[checked_df['#DOC'].isin(checked_items_doc)]
            # OR
            else:
                or_df = self.db_df[self.db_df['#DOC'].isin(checked_items_doc)]
                if len(or_df) != 0:
                    checked_df = pd.concat([checked_df, or_df])

        if len(checked_items_marker) != 0:
            if checked_df is None:
                checked_df = self.db_df[self.db_df['#REF'].isin(checked_items_marker)]
            # AND
            elif AND_OR == "AND":
                checked_df = checked_df[checked_df['#REF'].isin(checked_items_marker)]
            # OR
            else:
                or_df = self.db_df[self.db_df['#REF'].isin(checked_items_marker)]
                if len(or_df) != 0:
                    checked_df = pd.concat([checked_df, or_df])

        # если не нашли чекнутых - возвращаем целый DF
        if checked_df is None:
            checked_df = self.db_df

        # если перефильтровали пусто - возращаем None
        if len(checked_df) == 0:
            return None

        checked_df = checked_df.reset_index()

        return checked_df

    def db_export(self, to_clipboard=False, ext='csv'):

        allow_ext = ('csv', 'xlsx')
        if ext not in allow_ext:
            ext = 'csv'

        if self.db_path == '':
            print('# Info: Export table is empty')
            logger.info('Export table is empty')
            return None

        df_for_export = self.get_checked_df()

        if df_for_export is None:
            messagebox.showwarning(f'Export Info', "Nothing to export \U0001F625",
                                   icon='info')
            logger.info(f'Export Info: Nothing to export')
            return None

        # если есть кастом, то экспортим его, в противном случае - шаблон
        export_columns_list_index = self.process_columns_df['COL_INDEX'].tolist()
        export_columns_list_names = self.process_columns_df['COL_NAME'].tolist()

        if len(export_columns_list_index) != 0:
            export_df = df_for_export[export_columns_list_index]
            column_names = export_columns_list_names
            columns_count_val_sort(export_columns_list_index)

            # with open(export_path, 'w') as f:
            #    f.write('Custom String\n')

            # df.to_csv(export_path, header=False, mode="a")
            # print (cross_columns)

            if to_clipboard:

                # messagebox.showwarning(f'Clipboard Export', "Copied to the clipboard.",
                #                        icon='info')
                export_df.columns = column_names
                export_df.to_clipboard(sep='\t', index=False)
                print(f'# Info: Rows in clipboard: {len(export_df)}')
                logger.info(f'# Info: Rows in clipboard: {len(export_df)}')

                # export_df.to_clipboard(index=False)
            else:
                absbath = os.path.dirname(self.db_path)
                basename = os.path.basename(self.db_path)
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
                    if ext == 'csv' and self.lng != "SP":
                        to_csv_custom_header(df=export_df, csv_path=exportpath_csv, column_names=column_names,
                                             csv_encoding=csv_encoding)
                        change_1251_csv_encoding(exportpath_csv)
                        self.ask_open_file(exportpath_csv)

                    if ext == 'xlsx' or self.lng == "SP":
                        to_excel_custom_header(df=export_df, excel_path=exportpath_xlsx, column_names=column_names)
                        self.ask_open_file(exportpath_xlsx)


                except PermissionError:
                    self.columns_textbox.delete(0, END)
                    self.columns_textbox.insert(0, '')
                    logger.error('File save Permission Error')

                    if ext == 'csv' and self.lng != "SP":
                        print(f"'{basename[:-4]}.csv' is opened, saved as '{basename[:-4]}_1.csv'")
                        logger.info(f"'{basename[:-4]}.csv' is opened, saved as '{basename[:-4]}_1.csv'")
                        exportpath = os.path.join(absbath, basename)
                        exportpath_csv = f'{exportpath[:-4]}_1.csv'
                        to_csv_custom_header(df=export_df, csv_path=exportpath_csv, column_names=column_names,
                                             csv_encoding=csv_encoding)
                        change_1251_csv_encoding(exportpath_csv)
                        self.ask_open_file(exportpath_csv)
                    if ext == 'xlsx' or self.lng != "SP":
                        print(f"'{basename[:-4]}.xlsx' is opened, saved as '{basename[:-4]}_1.xlsx'")
                        logger.info(f"'{basename[:-4]}.xlsx' is opened, saved as '{basename[:-4]}_1.xlsx'")
                        exportpath = os.path.join(absbath, basename)
                        exportpath_xlsx = f'{exportpath[:-4]}_1.xlsx'
                        to_excel_custom_header(df=export_df, excel_path=exportpath_xlsx, column_names=column_names)
                        self.ask_open_file(exportpath_xlsx)

                print(f'# Info: Rows Exported: {len(export_df)}')
                logger.info(f'# Info: Rows Exported: {len(export_df)}')

                # if self.lng == "SP":
                #     export_df.columns = column_names
                #     export_df.to_excel(exportpath_xlsx, encoding=csv_encoding, index=False)
        else:
            messagebox.showwarning(f'Export Error', "Columns for Export not selected...",
                                   icon='warning')
            logger.info(f'Export Error: Columns for Export not selected...')


if __name__ == "__main__":

    arg = ''

    try:
        if len(sys.argv[1:]) == 2:
            run_atr = sys.argv[1:][0]
            path = sys.argv[1:][1]

            if run_atr == "-D":
                export_default(dbf_path=path)
                input("")
            else:
                input("### Error: Wrong Attribute")
                logger.error(f'Wrong Attribute: {run_atr}')

        else:
            DB_FORM()
    except Exception as ex:
        logger.exception("RUNTIME ERROR")
        input(f"### ERROR: {ex}")
