import os
import pathlib
import datetime
from datetime import datetime, date, timedelta
from tkinter import *
from tkinter import ttk
from TkinterDnD2 import *
from backup_save import BackupFile
from tkinter import filedialog as fd
from resource_path import resource_path
import sys


class BcSaverForm:
    def __init__(self):
        self.backup_delay = 15
        self.window = TkinterDnD.Tk()
        self.window.resizable(False, False)
        self.window.attributes('-topmost', True)
        self.backup_auto_class = BackupFile

        self.dbf_path = ""
        self.dbf_filename = ""

        self.act_file_frame = ttk.Labelframe(master=self.window, text='Active File', width=50, height=50)

        self.act_file_label = Label(self.act_file_frame, text="-DB Name-", font=("Verdana", 12),
                                    justify='center', foreground='red', padx=5, pady=3, width=23)
        self.file_frame = ttk.Labelframe(master=self.window, text='DBF', width=50, height=50)
        self.open_button = Button(master=self.file_frame, text='Open a File', command=self.select_file)

        self.backup_frame = ttk.Labelframe(master=self.window, text='Backup Settings', width=50, height=50)
        self.save_backup_button = Button(master=self.backup_frame, text='Save Backup', command=self.save_backup_manual)
        self.count_down_lbl = ttk.Label(master=self.backup_frame, text='')

        self.on_top_variable = BooleanVar()
        self.on_top_variable.set(1)
        self.on_top_ch = Checkbutton(self.backup_frame, text="on Top",
                                     variable=self.on_top_variable,
                                     onvalue=1, offvalue=0,
                                     command=self.stay_on_top)

        self.delay_variable = StringVar()
        self.delay_variable.trace("w", self.update_delay_time)
        self.delay_combobox = OptionMenu(self.backup_frame, self.delay_variable, *(5, 10, 15, 30))
        self.delay_combobox.config(width=2)
        self.delay_variable.set(15)
        self.last_save_time = datetime(year=2000, month=1, day=1, hour=1, minute=1, second=1)

        self.window.drop_target_register(DND_FILES)
        self.window.dnd_bind('<<Drop>>', self.open_with_dnd)

        self.place_elements()

        self.set_window_title()
        self.window.geometry("258x108")
        ico_abs_path = resource_path('bk_icon\BK.ico')
        self.window.wm_iconbitmap(ico_abs_path)

        self.window.protocol("WM_DELETE_WINDOW", self.on_closing)

        self.auto_save()

        arg = ''
        for arg in sys.argv[1:]:
            if os.path.exists(arg):
                print('Drag&Drop file: ', arg)
                self.check_path(arg)

        self.window.mainloop()

    def place_elements(self):

        self.act_file_frame.grid(row=0, column=0, sticky=EW, columnspan=1, padx=5)
        self.act_file_label.grid(row=0, column=0, sticky=EW, columnspan=1)

        self.backup_frame.grid(row=1, column=0, sticky=EW, columnspan=1, padx=5)
        self.save_backup_button.grid(row=0, column=0, sticky=EW, padx=2, pady=5)
        self.delay_combobox.grid(row=0, column=1, sticky=EW, columnspan=1)
        self.on_top_ch.grid(row=0, column=2, sticky=W, columnspan=1)
        self.count_down_lbl.grid(row=0, column=3, sticky=W, columnspan=1)

        # self.file_frame.grid(row=1, column=0, sticky=EW, columnspan=3)
        # self.open_button.grid(row=1, column=0, padx=2, pady=2)

    def set_window_title(self):
        self.window.title(f'Backup Saver')
        # self.window.title(f'DBF Helper: {self.dbf_filename}')

    def update_delay_time(self, *args):
        self.backup_delay = int(self.delay_variable.get()) * 60 * 1000
        print(f"Delay Time changed to: {self.backup_delay}")
        self.refresh_autosave()
        self.auto_save()

        # fin_date = self.last_save_time + timedelta(0, int(self.delay_variable.get()))
        # left = datetime.now() - fin_date
        # #left_formatted = datetime.strptime(left, "%s").date()
        # self.count_down_lbl.config(text=left.seconds)

    def refresh_autosave(self):
        self.window.after(self.backup_delay, self.auto_save)

    def auto_save(self):
        if self.dbf_filename != '':
            self.refresh_autosave()
            # self.last_save_time = datetime.now()
            self.save_backup_auto()

    def save_backup_manual(self):
        backup = BackupFile(self.dbf_path, mode='manual')
        backup.save_backup()

    def select_file(self):

        filetypes = (('DBF files', '*.dbf'),)

        file_path = fd.askopenfilename(
            title='Open a file',
            initialdir=os.path.abspath(__file__),
            filetypes=filetypes)

        self.check_path(file_path)

    def check_path(self, file_path):
        if pathlib.Path(file_path).suffix.lower() == ".dbf":
            self.dbf_path = file_path
            self.dbf_filename = os.path.basename(file_path)
            self.set_act_file_label()
            self.backup_auto_class = BackupFile(self.dbf_path, mode='auto')
            self.save_backup_auto()

    def stay_on_top(self):
        self.window.attributes('-topmost', self.on_top_variable.get())

    def open_with_dnd(self, event):
        if event.data[0] == '{':
            dnd_path = event.data[1:-1]
        else:
            dnd_path = event.data
        if dnd_path.endswith(('.DBF', '.dbf')):
            self.dbf_path = dnd_path
            self.dbf_filename = os.path.basename(dnd_path)
            self.set_act_file_label()

            self.backup_auto_class = BackupFile(self.dbf_path, mode='auto')
            self.save_backup_auto()

    def set_act_file_label(self):
        self.act_file_label.config(text=self.dbf_filename)

    def save_backup_auto(self):

        # diff = abs(datetime.now() - self.last_save_time.second)
        # start_the_clock + timedelta(seconds=10)
        if datetime.now() > (self.last_save_time + timedelta(seconds=2)):
            now_date_time = datetime.today().strftime("%d.%m.%Y %H:%M:%S")
            print(f'Autosave: {now_date_time}')
            self.backup_auto_class.save_backup()
            self.last_save_time = datetime.now()

    def on_closing(self):
        self.window.destroy()
        sys.exit()


if __name__ == '__main__':
    BcSaverForm()
