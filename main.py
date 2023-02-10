from Loginpage import *
from tkinter import messagebox
# import tkinter as tk


root = Tk()
root.title('换标工具')


def on_closing():
    if messagebox.askokcancel("Quit", "你确定要退出吗?"):
        root.destroy()


root.protocol("WM_DELETE_WINDOW", on_closing)
LoginPage(root)
root.mainloop()

