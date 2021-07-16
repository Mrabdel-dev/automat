# just try to learn well
# from tkinter import *
# from tkinter import filedialog

# top = Tk()
# top.title("Maneo File Generator")
# top.minsize(800, 400)
#
#
# def browsefunc():
#     filename = filedialog.askopenfilename()
#     pathlabel.config(text=filename)
#
#
# browsebutton = Button(top, text="Browse", command=browsefunc)
# browsebutton.pack()
#
# pathlabel = Label(top)
# pathlabel.pack()
# top.mainloop()
import datetime

date = datetime.datetime.now()
now = date.strftime("%m/%Y")
print(now)


def aroundTo(x: int, num):
    y = x % num
    if y != 0:
        k = x + num - y
        return int((k / num) - 1)
    else:
        return int((x / num) - 1)


diameter = {12: 6, 24: 8.5, 36: 8.5, 48: 9.5, 72: 10.5, 96: 11.5, 144: 11.5, 288: 14.5}
diameter.update({24:2.5})
print(diameter[24])
