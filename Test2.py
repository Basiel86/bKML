import pandas as pd
import os
import xlwings as xw
from pandasgui import show
import tkinter as tk
from tkinter import ttk
from ttkwidgets import CheckboxTreeview


class Tree:

    def OnDoubleClick(self, event):

        # активное
        # item = self.stat_tree.selection()[0]
        # print("you clicked on", self.stat_tree.item(item, "text"))

        # пробежка по чекнутым
        for item in self.stat_tree.get_checked():
            item_text = self.stat_tree.item(item, "text")
            print(item_text)

    def __init__(self):
        df = pd.read_csv(r'c:\Users\Vasily\OneDrive\Macro\PYTHON\bKML\Test\Export_CSV.csv')
        df2 = df['#FEATURE'].value_counts(dropna=False, normalize=False)

        stat = df.groupby('#FEATURE').count()

        root = tk.Tk()
        self.stat_tree = CheckboxTreeview(root, show='tree', column=("c1", "c2"))  # hide tree headings
        self.stat_tree.column("# 1", anchor='center', stretch='NO', width=50)
        self.stat_tree.column("# 2", anchor='center', stretch='NO', width=200)

        #self.stat_tree.column("# 1", anchor=CENTER, stretch=NO, width=100)
        self.stat_tree.pack()
        self.stat_tree.bind("<Double-1>", self.OnDoubleClick)

        style = ttk.Style(root)
        # remove the indicator in the treeview
        style.layout('Checkbox.Treeview.Item',
                     [('Treeitem.padding',
                       {'sticky': 'nswe',
                        'children': [('Treeitem.image', {'side': 'left', 'sticky': ''}),
                                     ('Treeitem.focus', {'side': 'left', 'sticky': '',
                                                         'children': [('Treeitem.text',
                                                                       {'side': 'left', 'sticky': ''})]})]})])
        # make it look more like a listbox
        style.configure('Checkbox.Treeview', borderwidth=1, relief='sunken')

        # get data
        # mydir = (os.getcwd()).replace('\\', '/') + '/'
        # mySiteCode = pd.read_excel(r'' + mydir + 'Governance_Tracker - Copy - Copy.xlsm', usecols=['SiteCode'],
        #                            encoding='latin-1', header=1)
        #
        # a = mySiteCode['SiteCode'].values.tolist()

        for index, value in df2.items():
            print(f"Index : {index}, Value : {value}")
            # z = [1, 22, 33, 44, 55]
            self.stat_tree.insert('', 'end', text=index)

        parent = list(self.stat_tree.get_checked())
        print(parent)

        root.mainloop()


if __name__ == '__main__':
    tree = Tree()
