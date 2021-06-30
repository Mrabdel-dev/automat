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