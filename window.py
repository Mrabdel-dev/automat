from tkinter import *


def btn_clicked():
    print("Button Clicked")


window = Tk()

window.geometry("1565x1143")
window.configure(bg = "#ffffff")
canvas = Canvas(
    window,
    bg = "#ffffff",
    height = 1143,
    width = 1565,
    bd = 0,
    highlightthickness = 0,
    relief = "ridge")
canvas.place(x = 0, y = 0)

entry0_img = PhotoImage(file = f"img_textBox0.png")
entry0_bg = canvas.create_image(
    211.5, -292.0,
    image = entry0_img)

entry0 = Entry(
    bd = 0,
    bg = "#d2e0d4",
    highlightthickness = 0)

entry0.place(
    x = -56.0, y = -321,
    width = 535.0,
    height = 56)

entry1_img = PhotoImage(file = f"img_textBox1.png")
entry1_bg = canvas.create_image(
    211.5, -165.0,
    image = entry1_img)

entry1 = Entry(
    bd = 0,
    bg = "#d2e0d4",
    highlightthickness = 0)

entry1.place(
    x = -56.0, y = -194,
    width = 535.0,
    height = 56)

entry2_img = PhotoImage(file = f"img_textBox2.png")
entry2_bg = canvas.create_image(
    211.5, -48.0,
    image = entry2_img)

entry2 = Entry(
    bd = 0,
    bg = "#d2e0d4",
    highlightthickness = 0)

entry2.place(
    x = -56.0, y = -77,
    width = 535.0,
    height = 56)

entry3_img = PhotoImage(file = f"img_textBox3.png")
entry3_bg = canvas.create_image(
    353.5, 113.0,
    image = entry3_img)

entry3 = Entry(
    bd = 0,
    bg = "#d2e0d4",
    highlightthickness = 0)

entry3.place(
    x = 128.0, y = 84,
    width = 451.0,
    height = 56)

entry4_img = PhotoImage(file = f"img_textBox4.png")
entry4_bg = canvas.create_image(
    565.5, 230.0,
    image = entry4_img)

entry4 = Entry(
    bd = 0,
    bg = "#d2e0d4",
    highlightthickness = 0)

entry4.place(
    x = 411.0, y = 201,
    width = 309.0,
    height = 56)

entry5_img = PhotoImage(file = f"img_textBox5.png")
entry5_bg = canvas.create_image(
    211.5, -421.0,
    image = entry5_img)

entry5 = Entry(
    bd = 0,
    bg = "#d2e0d4",
    highlightthickness = 0)

entry5.place(
    x = -56.0, y = -450,
    width = 535.0,
    height = 56)

img0 = PhotoImage(file = f"img0.png")
b0 = Button(
    image = img0,
    borderwidth = 0,
    highlightthickness = 0,
    command = btn_clicked,
    relief = "flat")

b0.place(
    x = -85, y = 84,
    width = 158,
    height = 58)

img1 = PhotoImage(file = f"img1.png")
b1 = Button(
    image = img1,
    borderwidth = 0,
    highlightthickness = 0,
    command = btn_clicked,
    relief = "flat")

b1.place(
    x = 546, y = -318,
    width = 162,
    height = 58)

img2 = PhotoImage(file = f"img2.png")
b2 = Button(
    image = img2,
    borderwidth = 0,
    highlightthickness = 0,
    command = btn_clicked,
    relief = "flat")

b2.place(
    x = 546, y = -66,
    width = 158,
    height = 58)

img3 = PhotoImage(file = f"img3.png")
b3 = Button(
    image = img3,
    borderwidth = 0,
    highlightthickness = 0,
    command = btn_clicked,
    relief = "flat")

b3.place(
    x = 546, y = -194,
    width = 169,
    height = 58)

img4 = PhotoImage(file = f"img4.png")
b4 = Button(
    image = img4,
    borderwidth = 0,
    highlightthickness = 0,
    command = btn_clicked,
    relief = "flat")

b4.place(
    x = 60, y = 361,
    width = 152,
    height = 83)

img5 = PhotoImage(file = f"img5.png")
b5 = Button(
    image = img5,
    borderwidth = 0,
    highlightthickness = 0,
    command = btn_clicked,
    relief = "flat")

b5.place(
    x = 534, y = -449,
    width = 181,
    height = 58)

img6 = PhotoImage(file = f"img6.png")
b6 = Button(
    image = img6,
    borderwidth = 0,
    highlightthickness = 0,
    command = btn_clicked,
    relief = "flat")

b6.place(
    x = 275, y = 558,
    width = 158,
    height = 58)

background_img = PhotoImage(file = f"background.png")
background = canvas.create_image(
    58.0, 59.5,
    image=background_img)

window.resizable(False, False)
window.mainloop()
