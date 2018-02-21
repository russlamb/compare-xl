import os
from tkinter import *
from itertools import zip_longest
import configparser
from compare_sql_xl import CompareSqlInXl
from io import StringIO

def make_form(root, fields, defaults):
    entries = []
    for field, default in zip(fields,defaults):
        row = Frame(root)
        lab = Label(row, width=15, text=field, anchor='w')
        ent = Entry(row)
        ent.insert(0,default)
        row.pack(side=TOP, fill=X, padx=5, pady=5)
        lab.pack(side=LEFT)
        ent.pack(side=RIGHT, expand=YES, fill=X)
        entries.append((field, ent))
    return entries

def add_text(root, text_fields, defaults):
    texts = []
    for field,default in zip(text_fields,defaults):
        row = Frame(root)
        lab = Label(row, width=15, text=field, anchor='w')
        txt = Text(row, height=5)
        txt.insert(END,default)
        row.pack(side=TOP, fill=X, padx=5, pady=5)
        lab.pack(side=LEFT)
        txt.pack(side=RIGHT, expand=YES, fill=X)
        texts.append((field,txt))

    return texts

def add_checkbox(root, checkboxes, variables):
    checks = []

    for field, var in zip(checkboxes , variables):
        row = Frame(root)
        lab = Label(row, width=15, text=field, anchor='w')
        chk = Checkbutton(row, text=field, variable=var)
        row.pack(side=TOP, fill=X, padx=5, pady=5)
        lab.pack(side=LEFT)
        chk.pack(side=RIGHT, expand=YES, fill=X)
        checks.append((field,chk))
        # here we create a property on the checkbox called my_variable_reference and reference the variable with it
        chk.my_variable_reference = var

    return checks

def fetch(controls):

    d=dict()
    entries = controls["entry"]
    text_fields = controls["text"]
    checks = controls["check"]

    for entry in entries:
        # control name is first entry in tuple
        field = entry[0]
        # control is second entry in tuple
        text = entry[1].get()
        d[field]=text
        print('{}: "{}"'.format(field, text))

    for text in text_fields:
        field = text[0]
        txt = text[1].get("1.0",END).rstrip()
        d[field]=txt
        print('{}: "{}"'.format(field, txt))

    for chk in checks:
        field = chk[0]
        check = chk[1].my_variable_reference.get()
        d[field] = check
        print('{}: "{}"'.format(field, check))

    return d


def get_defaults():
    """This reads the config file to return two lists of values, which will be the default values for the fields.
    Since we are zipping the defaults and the controls, you *must* add a default value for any new control.
    Otherwise, the control won't be created."""
    config_parser = configparser.ConfigParser()
    config_parser.read("config.ini")
    config = config_parser["DEFAULT"]

    field_vals = [
        config["left"],
        config["right"],
        config["diff"],
        config["sheet"],
        "" # index
    ]
    text_vals = [
        "select assetid,ChangeDateTime from asset",
        config["db_left"],
        config["db_right"],
        "" # console output
    ]
    return field_vals, text_vals


def run_compare(controls):
    """Compare the two files.  Redirect console messages to form."""
    console_out_control = [ctrl[1] for ctrl in controls["text"] if ctrl[0]=="Console Output"][0]

    print(console_out_control.get("1.0",END).rstrip())
    console_out_control.delete("1.0",END)
    old_stdout = sys.stdout
    old_stderr = sys.stderr
    try:
        #redirect standard streams to string in-memory streams
        sys.stdout = console_output = StringIO()
        sys.stderr = console_error = StringIO()
        d=fetch(controls)

        compare = CompareSqlInXl(d["DB Left"], d["DB Right"])

        if d["Compare Only"]==1:
            compare.compare_only(d["Left File"], d["Right File"], d["Diff File"], d["Sheet"])
        else:
            sql = d["SQL"]
            compare.run(d["Left File"], d["Right File"], d["Diff File"], d["Sheet"], sql)

        if d["Open Differences"] == 1:
            print("Opening file {}".format(os.path.normpath(d["Diff File"])))
            os.startfile(os.path.normpath(d["Diff File"]))

        console_out_control.insert(END,console_output.getvalue())
        console_out_control.insert(END, console_error.getvalue())
    except Exception as ex:
        # print error to output control and to the real standard error stream
        console_out_control.insert(END, console_output.getvalue())
        console_out_control.insert(END, console_error.getvalue())
        console_out_control.insert(END, ex)
        #print to standard error stream
        print(ex, file=sys.__stderr__)
    finally:
        # set old output streams back in place
        sys.stderr=old_stderr
        sys.stdout=old_stdout

    print("Done")

def print_control_values(controls):
    print(controls)
    d=fetch(controls)
    print(d)
    print(d["Compare Only"])
    print(d["Open Differences"])

def run_tk():
    """build a form using lists of fields.  Controls are passed to button event in a dictionary"""
    root=Tk()
    text_fields = ["SQL", "DB Left", "DB Right","Console Output"]

    fields = [
        "Left File",
        "Right File",
        "Diff File",
        "Sheet",
        "Index"
    ]
    checkboxes = [
        "Compare Only",
        "Open Differences"
    ]

    check_vars = [
        IntVar(),
        IntVar()
    ]
    # set open differences to 1 by default
    check_vars[1].set(1)

    # defaults are lists that MUST be of the same lentgh as the respective field / text list
    (field_default,text_default) = get_defaults()

    # populate the form
    txts = add_text(root, text_fields, text_default)
    ents = make_form(root, fields,field_default)
    checks = add_checkbox(root,checkboxes,check_vars)

    #controls dictionary splits the fields into two types : entry type and text type
    #each list is a list of tuples.  (field name, control)
    controls = {"entry":ents,"text":txts, "check":checks}

    # Add button
    b1 = Button(root, text="Compare", command=(lambda e=controls: run_compare(e)))
    # b1 = Button(root, text="Compare", command=(lambda e=controls: print_control_values(e)))
    b1.pack(side=LEFT,padx=5,pady=5)

    root.mainloop()

if __name__ == "__main__":
    run_tk()
