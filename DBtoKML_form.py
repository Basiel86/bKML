from tkinter import *
import os
import DBtoKML
from tkinter import filedialog as fd
from datetime import datetime, date
import pandas as pd
import parse_inch
from parse_inch import parse_inch_prj
import traceback



EXP_DAY = '2022-08-20'

lng_list = ["RU", "EN"]
dbf_ext_list = ['dbf', 'DBF']


# diam_list = inch_mm_df['Inch_name'].tolist()
# mm_list = inch_mm_df['Inch'].tolist()

inch_names_list = parse_inch.get_inch_names_list()
inch_list = parse_inch.get_inch_list()
inch_dict = parse_inch.get_inch_dict()

# diam_list = [4, 4.5, 5.563, 6.625, 8.625, 10.75, 12.75, 14, 16, 18, 20, 22, 24, 26, 28, 30, 32, 34, 36, 38,
#              40, 42, 44, 46, 48, 52, 56]


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

    lng = str(lng_list_variable.get())

    diam_list_value = diam_list_variable.get()
    diameter = inch_dict[diam_list_value]

    if dbf_path != "" and dbf_path[-3:] in dbf_ext_list:

        try:
            DBtoKML.bKML(dbf_path=dbf_path, lang=lng, diameter=diameter).dbf_to_kml()
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
    userForm.geometry("600x70")

    myLabel = Label(userForm, text="DBF файл")
    myLabel.pack()

    path_variable = StringVar(userForm)
    path_textbox = Entry(userForm, width=98, textvariable=path_variable)
    path_textbox.pack()


    myButon = Button(userForm, text="DBF to KML", command=DBFtoKML)
    myButon.pack(side='left')

    lng_list_variable = StringVar(userForm)
    MFG_combobox = OptionMenu(userForm, lng_list_variable, *lng_list)
    MFG_combobox.pack(side='right')
    lng_list_variable.set(lng_list[0])

    fileButon = Button(userForm, text="Open file", command=openfile)
    fileButon.pack(side='right')
    # fileButon.place(x=270, y=42)

    blankLabel = Label(userForm, text="          ")
    blankLabel.pack(side='left')

    inchLabel = Label(userForm, text="Diam")
    inchLabel.pack(side='left')

    diam_list_variable = StringVar(userForm)
    diam_combobox = OptionMenu(userForm, diam_list_variable, *inch_names_list)
    diam_combobox.pack(side='left')

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



    def on_closing():
        userForm.destroy()
        sys.exit()


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


    def on_closing():
        userForm.destroy()
        sys.exit()


    userForm.protocol("WM_DELETE_WINDOW", on_closing)
    userForm.mainloop()

userForm.mainloop()


def on_closing():
    userForm.destroy()
    sys.exit()
