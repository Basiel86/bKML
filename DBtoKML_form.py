from tkinter import *
import os
import DBtoKML
from tkinter import filedialog as fd
from datetime import datetime, date

EXP_DAY = '2022-06-30'

lng_list = ["RU", "EN"]
dbf_ext_list = ['dbf', 'DBF']
diam_list = [4, 4.5, 5.563, 6.625, 8.625, 10.75, 12.75, 14, 16, 18, 20, 22, 24, 26, 28, 30, 32, 34, 36, 38,
             40, 42, 44, 46, 48, 52, 56]


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
        print("Log export error: ", ex)


def DBFtoKML():
    dbf_path = str(myTextbox.get())

    lng = str(lng_list_variable.get())
    diameter = float(diam_list_variable.get())

    if dbf_path != "" and dbf_path[-3:] in dbf_ext_list:

        try:
            print("\nProcessing....")
            DBtoKML.bKML(dbf_path=dbf_path, lang=lng, diameter=diameter).dbf_to_kml()
            write_log(filepath=dbf_path)
        except Exception as ex:
            print(ex)
            print("Что-то пошло не так...")
    else:
        print("Путь не верен, ты сбился с пути, друг?")


def openfile():
    filename = fd.askopenfilename(
        title='Open a file',
        initialdir='/',
        filetypes=[("DBF files", ".dbf .DBF")])

    textEntry.set(filename)


userForm = Tk()

exp_date_formatted = datetime.strptime(EXP_DAY, "%Y-%m-%d").date()
now_date = date.today()
days_left = exp_date_formatted - now_date

if exp_date_formatted >= now_date:

    userForm.title("DBF to KML")
    userForm.geometry("600x70")

    myLabel = Label(userForm, text="DBF файл")
    myLabel.pack();

    textEntry = StringVar(userForm)
    myTextbox = Entry(userForm, width=580, textvariable=textEntry)
    myTextbox.pack();

    myButon = Button(userForm, text="DBF to KML", command=DBFtoKML)
    myButon.pack(side='left')

    lng_list_variable = StringVar(userForm)
    MFG_combobox = OptionMenu(userForm, lng_list_variable, *lng_list)
    MFG_combobox.pack(side='right')
    lng_list_variable.set(lng_list[0])

    fileButon = Button(userForm, text="Open file", command=openfile)
    fileButon.pack(side='right')
    # fileButon.place(x=270, y=42)

    diam_list_variable = StringVar(userForm)
    diam_combobox = OptionMenu(userForm, diam_list_variable, *diam_list)
    diam_combobox.pack(side='left')
    diam_list_variable.set(diam_list[4])

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
