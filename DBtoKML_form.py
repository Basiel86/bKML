from tkinter import *
import os
import DBtoKML
from tkinter import filedialog as fd
from datetime import datetime, date
import pandas as pd
import parse_inch
from parse_inch import parse_inch_prj
import traceback

EXP_DAY = '2022-10-01'

lng_list = ["RU", "EN"]
dbf_ext_list = ['dbf', 'DBF']

# diam_list = inch_mm_df['Inch_name'].tolist()
# mm_list = inch_mm_df['Inch'].tolist()

inch_names_list = parse_inch.get_inch_names_list()
inch_list = parse_inch.get_inch_list()
inch_dict = parse_inch.get_inch_dict()


# diam_list = [4, 4.5, 5.563, 6.625, 8.625, 10.75, 12.75, 14, 16, 18, 20, 22, 24, 26, 28, 30, 32, 34, 36, 38,
#              40, 42, 44, 46, 48, 52, 56]
line_width = [1, 3, 5]


def write_log(filepath=""):
    log_path = r"\\vasilypc\Vasily Shared (Full Access)\###\DBtoKMLlog\DBtoKMLlog.txt"

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
                       f'{filepath}\n')
        log_file.close()
    except Exception as ex:
        print("LOG Error: ", ex)


def DBFtoKML():
    dbf_path = str(path_textbox.get())
    line_width = line_width_variable.get()

    lng = str(lng_list_variable.get())

    diam_list_value = diam_list_variable.get()
    diameter = inch_dict[diam_list_value]

    if dbf_path != "" and dbf_path[-3:] in dbf_ext_list:

        try:
            DBtoKML.bKML(dbf_path=dbf_path, lang=lng, diameter=diameter).dbf_to_kml(line_width=line_width)
            try:
                write_log(filepath=dbf_path)
            except Exception as logex:
                print("LOG Error")

        except Exception as ex:
            print(ex)
            print("Что-то пошло не так...")
    else:
        print("Путь не верен! Ты сбился с пути!?")


def openfile():
    filename = fd.askopenfilename(
        title='Open a file',
        initialdir='/',
        filetypes=[("DBF files", ".dbf .DBF")])

    path_variable.set(filename)

    path_inch = parse_inch_prj(filename)
    # если нашли инч то селектим его в выпадающем списке
    if path_inch is not None:
        inch_index = inch_list.index(path_inch)
        diam_list_variable.set(inch_names_list[inch_index])


def resource_path(relative_path):
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception as ex:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def on_closing():
    userForm.destroy()
    sys.exit()


userForm = Tk()

# читаем аргументы на входе
arg = ''
for arg in sys.argv[1:]:
    print('Drag&Drop file: ', arg)

exp_date_formatted = datetime.strptime(EXP_DAY, "%Y-%m-%d").date()
now_date = date.today()
days_left = exp_date_formatted - now_date

ico_abs_path = resource_path('KML.ico')
userForm.wm_iconbitmap(ico_abs_path)

if exp_date_formatted >= now_date:

    userForm.title("DBF to KML")
    userForm.geometry("600x75")

    myLabel = Label(userForm, text="DBF файл")

    path_variable = StringVar(userForm)
    path_textbox = Entry(userForm, width=98, textvariable=path_variable)
    myButon = Button(userForm, text="DBF to KML", command=DBFtoKML)
    fileButon = Button(userForm, text="Open file", command=openfile)

    lng_list_variable = StringVar(userForm)
    lng_combobox = OptionMenu(userForm, lng_list_variable, *lng_list)
    lng_list_variable.set(lng_list[0])

    line_width_label = Label(userForm, text="Line Width")
    line_width_variable = StringVar(userForm)
    line_width_combobox = OptionMenu(userForm, line_width_variable, *line_width)
    line_width_variable.set(3)

    inchLabel = Label(userForm, text="Diam")
    diam_list_variable = StringVar(userForm)
    diam_combobox = OptionMenu(userForm, diam_list_variable, *inch_names_list)


    myLabel.grid(row=0, column=0, columnspan=10)
    path_textbox.grid(row=1, column=0, columnspan=10)

    myButon.grid(row=2, column=0, sticky=W+E, padx=5, pady=5)

    inchLabel.grid(row=2, column=2, sticky=E)
    diam_combobox.grid(row=2, column=3, sticky=W)

    line_width_label.grid(row=2, column=6, sticky=E)
    line_width_combobox.grid(row=2, column=7, sticky=W)

    fileButon.grid(row=2, column=8, sticky=E)
    lng_combobox.grid(row=2, column=9)

    if arg != '':
        path_variable.set(arg)
        # пишем аргумент в строку пути
        path_variable.set(arg)
        # парсим инч
        arg_inch = parse_inch_prj(arg)
        # если нашли инч то селектим его в выпадающем списке
        if arg_inch is not None:
            inch_index = inch_list.index(arg_inch)
            diam_list_variable.set(inch_names_list[inch_index])
        else:
            diam_list_variable.set(inch_names_list[6])
    else:
        diam_list_variable.set(inch_names_list[6])

    userForm.protocol("WM_DELETE_WINDOW", on_closing)
    userForm.mainloop()

else:
    userForm.geometry("1900x850")
    userForm.title('EXPIRED')
    userForm.geometry("300x50")

    exp_label = Label(master=userForm, text="Your version has expired", fg='red', font=15)
    exp_label.pack()
    exp_label2 = Label(master=userForm, text="Please update", fg='red', font=15)
    exp_label2.pack()

    userForm.protocol("WM_DELETE_WINDOW", on_closing)
    userForm.mainloop()

userForm.mainloop()
