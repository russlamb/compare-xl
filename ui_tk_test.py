import configparser
import os
from tkinter import Tk, Label, Text, END, IntVar, Checkbutton, W, Button, mainloop

from compare_sql_xl import CompareSqlInXl


def basic_form():
    master = Tk()
    Label(master, text="SQL").grid(row=0)
    e1 = Text(master, height=5, width=20)
    e1.insert(END, "select AssetID,SourceChangeDateTime from Asset")
    e1.grid(row=0, column=1, padx=5, pady=5)
    var2 = IntVar()
    c1 = Checkbutton(master, text="Open after compare", variable=var2).grid(row=1, sticky=W)

    b1 = Button(master, text="Compare", command=(lambda e=[e1.get("1.0", END), var2]: compare_click(e)))
    b1.grid(row=2, padx=5, pady=5)
    mainloop()


def compare_click(args):
    config_parser = configparser.ConfigParser()
    config_parser.read("config.ini")
    config = config_parser["DEFAULT"]
    compare = CompareSqlInXl(config["db_left"], config["db_right"])
    sql = args[0]
    show_file = args[1].get()
    compare.run(config["left"], config["right"], config["diff"], config["sheet"], sql)

    if show_file == 1:
        print("Opening file {}".format(os.path.normpath(config["diff"])))
        os.startfile(os.path.normpath(config["diff"]))
    print("Done")