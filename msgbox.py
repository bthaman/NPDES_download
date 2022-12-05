"""
Contains functions to display a tkinter or Tkinter message box.

Functions:
    show_error(): Displays error messagebox
    show_message(): Displays informational messagebox
"""
try:
    import Tkinter as tk
    import tkMessageBox as tkmb
except ImportError:
    import tkinter as tk
    from tkinter import messagebox as tkmb


def show_error(title, message):
    """ Displays tkMessageBox window with 'Error' icon. """
    root = tk.Tk()
    root.withdraw()
    tkmb.showerror(title=title, message=message)


def show_message(title, message):
    """ Displays tkMessageBox window with information message icon. """
    root = tk.Tk()
    root.withdraw()
    tkmb.showinfo(title=title, message=message)

if __name__ == '__main__':
    import math
    show_message('Informative Message', 'You are about to see several\ndemonstrations '
                                        'of tkinter\'s messagebox\nshow_error and show_message methods.')
    try:
        x = 1
        y = 0
        print(x/y)
    except ZeroDivisionError as e:
        show_error('Error', e)
    try:
        print(math.log(y))
    except ValueError as e:
        show_error("Error", e)
    try:
        percent_effort_given = 100.001
        if percent_effort_given < 50:
            show_message('Come on!', 'Try harder!')
        elif percent_effort_given > 100:
            raise ValueError('You can\'t give more than 100%')
    except ValueError as e:
        show_error('Error', e)
