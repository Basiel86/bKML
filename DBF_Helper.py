import pathlib
from tkinter import *
from tkinter import ttk
from TkinterDnD2 import *
import dbf
import pandas as pd
import traceback
from datetime import datetime, date
from resource_path import resource_path
import os
from tkinter import filedialog as fd
from DF_DBF import df_DBF
from backup_save import BackupFile
import pyfiglet


class DBFHelper:
    def __init__(self):
        self.EXP_DAY = '2022-09-20'
        self.backup_delay = 15
        print('\n' + pyfiglet.figlet_format("DBf Halper", justify='center', font='larry3d'))
        self.window = TkinterDnD.Tk()
        self.window.resizable(False, False)
        self.window.attributes('-topmost', True)
        # убираем windows стандатрные кнопки
        self.window.attributes('-toolwindow', True)
        # self.window.attributes('-transparentcolor', "white")
        self.dbf_class = df_DBF()
        self.dbf_path = ""
        self.dbf_filename = ""

        self.backup_auto_class = BackupFile

        # -alpha
        # -transparentcolor
        # -disabled
        # -fullscreen
        # -toolwindow
        # -topmost

        self.columns_raw_list = ['No Columns']

        self.file_frame = ttk.Labelframe(master=self.window, text='DBF', width=50, height=50)
        self.open_button = Button(master=self.file_frame, text='Open a File', command=self.select_file)
        self.update_dbf_button = Button(master=self.file_frame, text='Update', command=self.update_dbf)
        self.open_in_excel_button = Button(master=self.file_frame, text='to EXL', command=self.open_in_excel)

        self.backup_frame = ttk.Labelframe(master=self.window, text='DBF backup', width=50, height=50)
        self.save_backup_button = Button(master=self.window, text='Save Backup', command=self.save_backup_manual)

        self.f_doc_button = Button(master=self.window, text='Doc', command=self.filter_doc)
        self.f_clear_button = Button(master=self.window, text='F Clear', command=self.f_clear)

        self.dbf_columns_frame = ttk.Labelframe(master=self.window, text='DBF Columns', width=50, height=50)
        self.dbf_columns_combobox = ttk.Combobox(self.dbf_columns_frame, values=["No Columns"],
                                                 postcommand=self.columns_list_update)

        self.window.drop_target_register(DND_FILES)
        self.window.dnd_bind('<<Drop>>', self.open_with_dnd)

        self.place_elements()

        exp_date_formatted = datetime.strptime(self.EXP_DAY, "%Y-%m-%d").date()
        now_date = date.today()
        days_left = exp_date_formatted - now_date
        if exp_date_formatted >= now_date:
            # self.window.title(f'DBF Helper, expires in (days): {days_left.days}')
            self.set_window_title()
            self.window.geometry("200x500")
            ico_abs_path = resource_path('icons\DBF helper.ico')
            self.window.wm_iconbitmap(ico_abs_path)
        else:
            self.window.title('DBF Helper - EXPIRED')
            self.window.geometry("300x50")
            self.window.resizable(False, False)

            exp_label = Label(master=self.window, text="Your version has expired", fg='red', font=15)
            exp_label.pack()
            exp_label2 = Label(master=self.window, text="Please update", fg='red', font=15)
            exp_label2.pack()

        self.window.protocol("WM_DELETE_WINDOW", self.on_closing)

        self.auto_save()
        self.window.mainloop()

    def save_backup_manual(self):
        backup = BackupFile(self.dbf_path, mode='manual')
        backup.save_backup()

    def save_backup_auto(self):
        self.backup_auto_class.save_backup()

    def auto_save(self):
        delay_ms = self.backup_delay * 60 * 1000
        self.window.after(delay_ms, self.auto_save)
        if self.dbf_filename != '':
            self.save_backup_auto()

    # stay_on_top()
    def set_window_title(self):
        self.window.title(f'DBF Helper: {self.dbf_filename}')

    def columns_list_update(self):
        self.dbf_columns_combobox["values"] = self.columns_raw_list

    def on_closing(self):
        self.window.destroy()
        sys.exit()

    def place_elements(self):

        self.file_frame.grid(row=0, column=0, sticky=EW, columnspan=3)
        self.open_button.grid(row=0, column=0, padx=2, pady=2)
        self.update_dbf_button.grid(row=0, column=1, padx=2, pady=2)
        self.open_in_excel_button.grid(row=0, column=2, padx=2, pady=2)

        self.backup_frame.grid(row=1, column=0, sticky=EW, columnspan=3)
        self.save_backup_button.grid(row=1, column=0, sticky=W + S, padx=2, pady=5)

        self.dbf_columns_frame.grid(row=2, column=0, columnspan=3, padx=2, pady=2, sticky=EW)
        self.dbf_columns_combobox.grid(row=2, column=0, padx=2, pady=2)
        self.dbf_columns_combobox.configure(width=20)

        self.f_clear_button.grid(row=5, column=0, columnspan=3, padx=2, pady=2, sticky=EW)
        self.f_doc_button.grid(row=6, column=0, padx=2, pady=2, sticky=EW)

    def open_with_dnd(self, event):
        if event.data[0] == '{':
            dnd_path = event.data[1:-1]
        else:
            dnd_path = event.data
        if dnd_path.endswith(('.DBF', '.dbf')):
            self.dbf_path = dnd_path
            self.dbf_filename = os.path.basename(dnd_path)
            self.set_window_title()
            headers = self.update_dbf()

            self.columns_raw_list = headers
            self.columns_list_update()
            self.backup_auto_class = BackupFile(self.dbf_path, mode='auto')
            self.save_backup_auto()

    def select_file(self):

        filetypes = (('DBF files', '*.dbf'),)

        file_path = fd.askopenfilename(
            title='Open a file',
            initialdir=os.path.abspath(__file__),
            filetypes=filetypes)

        if pathlib.Path(file_path).suffix.lower() == ".dbf":
            self.dbf_path = file_path
            self.dbf_filename = os.path.basename(file_path)
            self.set_window_title()
            self.update_dbf()
            self.backup_auto_class = BackupFile(self.dbf_path, mode='auto')
            self.save_backup_auto()

    def update_dbf(self):
        if os.path.exists(self.dbf_path):
            dbf_file = dbf.Table(self.dbf_path, codepage='cp1251')
            dbf_file.open()
            # print((*dbf_file.field_names,))
            return *dbf_file.field_names,
        else:
            return None

    def open_in_excel(self):
        headers = self.update_dbf()
        headers_df = pd.DataFrame(headers)

        # wb = Workbook()
        # ws = wb.active
        #
        # for r in dataframe_to_rows(headers_df, index=True, header=True):
        #     ws.append(r)
        # wb.save('test.xlsx')

    def load_dbf(self):
        if os.path.exists(self.dbf_path):
            dbf_path = self.dbf_path
            dbf_file = dbf.Table(dbf_path, codepage='cp1251')
            dbf_file.open()
            return dbf_file
        else:
            return None

    def filter_doc(self):
        dbf_file = self.load_dbf()
        if dbf_file is not None:
            if 'DOC' in dbf_file.field_names:
                self.f_clear()
                index = dbf_file.create_index(
                    key=lambda r: r.DOC if r.DOC is not None else -1)
                match = index.search(match=('*',), partial=True)
                total_changed = 0
                with dbf_file:
                    for record in match:
                        dbf.write(record, F=1)
                        total_changed += 1
                print(f"Total processed: {total_changed}")

    def f_clear(self):
        dbf_file = self.load_dbf()
        if dbf_file is not None:
            index = dbf_file.create_index(key=lambda r: r.F if (r.F is not None) else -1)
            match = index.search(match=(1,), partial=True)
            with dbf_file:
                for record in match:
                    dbf.write(record, F=None)


def vlk_process(dbf_path, excel_path, replace_column):
    total_changed = 0
    col_vars = ["DEPTH_PREV", "DEPTH", "COMMENT", "REMARKS", "LAN/LONG/ALT", "Clear FEA NUM"]

    try:
        # если выбрали не корректный столбец выходим
        if replace_column in col_vars:
            # открываем эксель и DBF
            excel_data = pd.read_excel(excel_path)
            dbf_file = dbf.Table(dbf_path, codepage='cp1251')
            dbf_file.open()

            # создаем индекс на столбец с НОМЕРОМ
            index = dbf_file.create_index(key=lambda r: r.FEA_NUM if (r.FEA_NUM is not None) else -1)

            for i, row in excel_data.iterrows():
                # бегаем по экселю
                num = row[0]
                # поиск в индексе номера
                match = index.search(match=(num,), partial=True)

                with dbf_file:
                    # бегаем найденному и меняем значения в выбранном столбце
                    for record in match:
                        total_changed += 1
                        if replace_column == 'DEPTH_PREV':
                            vlookup = nan_to_blank(row[1])
                            comment = float(vlookup)
                            dbf.write(record, DEPTH_PREV=comment)
                        if replace_column == 'DEPTH':
                            vlookup = nan_to_blank(row[1])
                            comment = float(vlookup)
                            dbf.write(record, FEA_DEPTH=comment)
                        if replace_column == 'COMMENT':
                            vlookup = nan_to_blank(row[1])
                            comment = str(vlookup)
                            dbf.write(record, COMMENT=comment)
                        if replace_column == 'REMARKS':
                            vlookup = nan_to_blank(row[1])
                            comment = str(vlookup)
                            dbf.write(record, REMARKS=comment)
                        if replace_column == 'LAN/LONG/ALT':
                            lat = nan_to_blank(row[1])
                            long = nan_to_blank(row[2])
                            alt = nan_to_blank(row[3])
                            dbf.write(record, LATITUDE=lat)
                            dbf.write(record, LONGITUDE=long)
                            dbf.write(record, HEIGHT=alt)
                        if replace_column == 'Clear FEA NUM':
                            dbf.write(record, FEA_NUM=None)

            dbf_file.close()
        print(f'\n~~~ Done with changed of {total_changed} record(s) ~~~')

    except Exception as ex:
        print(ex)
        print(traceback.format_exc())
        input('')


def nan_to_blank(poss_blank):
    if pd.isna(poss_blank):
        return None
    else:
        return poss_blank


def filter_depth(dbf_path, depth_min, depth_max):
    dbf_file = dbf.Table(dbf_path, codepage='cp1251')
    dbf_file.open()

    if 'FEA_DEPTH' in dbf_file.field_names:
        # создаем индекс на столбец с НОМЕРОМ
        # index = dbf_file.create_index(key=lambda r: r.F)
        # index = dbf_file.create_index(key=lambda r: r.FEA_DEPTH if (r.FEA_DEPTH is not None
        #                                                             and depth_min <= r.FEA_DEPTH < depth_max) else -1)
        clear_F(dbf_file)

        index = dbf_file.create_index(
            key=lambda r: r.FEA_DEPTH if r.FEA_DEPTH is not None else -1)

        total_changed = 0
        with dbf_file:
            for record in index:
                if record.FEA_DEPTH is not None:
                    cur_depth = record.FEA_DEPTH / record.WT * 100
                    if depth_min <= cur_depth < depth_max:
                        dbf.write(record, F=1)
                        total_changed += 1
        print(f"Total selected: {total_changed}")


def clear_F(dbf_file):
    index = dbf_file.create_index(key=lambda r: r.F if (r.F is not None) else -1)
    match = index.search(match=(1,), partial=True)
    with dbf_file:
        for record in match:
            dbf.write(record, F=None)


def main():
    root = Tk()

    def vlk():
        replace_column = VL_type_combo.get()
        dbf_path = DBF_path_variable.get()
        excel_path = Excel_path_variable.get()
        vlk_process(dbf_path, excel_path, replace_column)
        print("~~~Done~~~")

    def process_depth():
        dbf_path = DBF_path_variable.get()
        depth_min = depth_min_variable.get()
        depth_max = depth_max_variable.get()

        depth_min = 0 if depth_min == "" else float(depth_min)
        depth_max = 500 if depth_max == "" else float(depth_max)

        filter_depth(dbf_path, depth_min, depth_max)

    def filter_doc():

        dbf_path = DBF_path_variable.get()
        dbf_file = dbf.Table(dbf_path, codepage='cp1251')
        dbf_file.open()
        if 'DOC' in dbf_file.field_names:
            clear_F(dbf_file)
            index = dbf_file.create_index(
                key=lambda r: r.DOC if r.DOC is not None else -1)
            match = index.search(match=('*',), partial=True)
            total_changed = 0
            with dbf_file:
                for record in match:
                    dbf.write(record, F=1)
                    total_changed += 1
            print(f"Total processed: {total_changed}")

    def clear_f():
        dbf_path = DBF_path_variable.get()
        dbf_file = dbf.Table(dbf_path, codepage='cp1251')
        dbf_file.open()
        if 'F' in dbf_file.field_names:
            index = dbf_file.create_index(key=lambda r: r.F if (r.F is not None) else -1)
            match = index.search(match=(1,), partial=True)
            total_changed = 0
            with dbf_file:
                for record in match:
                    dbf.write(record, F=None)
                    total_changed += 1
            print(f"Total processed: {total_changed}")

    DBF_path_variable = StringVar(root)

    # DBF_path_variable.set(
    #    r"d:\WORK\#Colombia\Ecopetrol S.A\NEC 20 inch Trunk M-CPF, 13.96 km\Reports\PR\Database\123\1necm.DBF")
    DBF_label = Label(text="DBF path").grid(row=0, column=0, columnspan=3)
    DBF_path_textbox = Entry(width=100, textvariable=DBF_path_variable).grid(row=1, column=0, columnspan=10,
                                                                             sticky=W + E,
                                                                             pady=2)

    Excel_path_variable = StringVar(root)
    # Excel_path_variable.set(r"c:\Users\Vasily\OneDrive\Macro\PYTHON\ERF_CLEAR\Test\111.xlsx")
    Excel_label = Label(text="EXCEL path").grid(row=3, column=0, columnspan=3)
    Excel_path_textbox = Entry(width=100, textvariable=Excel_path_variable).grid(row=4, column=0, columnspan=10,
                                                                                 sticky=W + E, pady=2)

    VL_type_combo_variable = StringVar(root)
    VL_type_combo = ttk.Combobox(root,
                                 values=[
                                     "DEPTH_PREV",
                                     "DEPTH", "COMMENT",
                                     "REMARKS", "LAN/LONG/ALT",
                                     "Clear FEA NUM"], )

    VL_type_combo.grid(row=6, column=0, pady=3)

    Vl_run_button = Button(text="VLookUp", command=vlk)
    Vl_run_button.grid(row=6, column=1, pady=3)

    depth_min_variable = StringVar(root)
    depth_min = Entry(width=10, textvariable=depth_min_variable).grid(row=6, column=3)

    depth_max_variable = StringVar(root)
    depth_max = Entry(width=10, textvariable=depth_max_variable).grid(row=6, column=4)

    filter_button = Button(text="Filter D", command=process_depth)
    filter_button.grid(row=6, column=5, pady=3)

    mark_doc_button = Button(text="Doc", command=filter_doc)
    mark_doc_button.grid(row=6, column=6, pady=3)

    clear_f_button = Button(text="CLF", command=clear_f)
    clear_f_button.grid(row=6, column=7, pady=3)

    root.attributes('-topmost', True)

    root.mainloop()


if __name__ == '__main__':
    DBFHelper()
