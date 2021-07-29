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
# import datetime
#
# date = datetime.datetime.now()
# now = date.strftime("%m/%Y")
# print(now)
#
#
# def aroundTo(x: int, num):
#     y = x % num
#     if y != 0:
#         k = x + num - y
#         return int((k / num) - 1)
#     else:
#         return int((x / num) - 1)
#
#
# diameter = {12: 6, 24: 8.5, 36: 8.5, 48: 9.5, 72: 10.5, 96: 11.5, 144: 11.5, 288: 14.5}
# diameter.update({24:2.5})
# print(diameter[24])
# def is_palindrome(input_string):
#     # We'll create two strings, to compare them
#     new_string = ""
#     reverse_string = ""
#     # Traverse through each letter of the input string
#     for i in input_string:
#         # Add any non-blank letters to the
#         # end of one string, and to the front
#         # of the other string.
#         if i != "":
#             new_string = new_string + i
#             reverse_string = i + reverse_string
#     # Compare the strings
#     if new_string == reverse_string:
#         return True
#     return False
#
#
# print(is_palindrome("abc"))  # Should be False
# print(is_palindrome("kayak"))  # Should be True
# import openpyxl
#
# from os import walk
#
# # the folder source
# monRepertoire = r'C:/Users/etudes20/Desktop/tesstr/FOA SRO 262/'
# wb = openpyxl.load_workbook('C:/Users/etudes20/Desktop/tesstr/85_018_262_POINT_TECHNIQUE_C.xlsx')
# ws = wb.active
# Names = []
# max = ws.max_row
# for n in range(2, max + 1):
#     k = str(ws.cell(n, 1).value)
#     Names.append(k)
# listeFichiers = []
# for (repertoire, sousRepertoires, fichiers) in walk(monRepertoire):
#     listeFichiers.extend(fichiers)
# i = 0
# listName = []
# while i < len(listeFichiers):
#     x = str(listeFichiers[i])
#     x = x[0:-5].strip()
#     if len(x)>11:
#         x = x[0:10].strip()
#     listName.append(x)
#     i += 1
# test = {}
# f = 0
# N = 0
# print(listName)
# for c in Names:
#     found = "NF"
#     for j in listName:
#         if c == j:
#             found = "F"
#     if found.startswith("F"):
#         f += 1
#     else:
#         N += 1
#     print("le point tech " + c + f" is {found} on folder names ")
# print(f"the number of element found is {f} and the number of element not found is {N}")


    # nbPrise.append(zaPboDbl.records[k]['nb_prise'])
    # tECHNO.append(zaPboDbl.records[k]['techno'])
    # typeBat.append(zaPboDbl.records[k]['type_bat'])
    # statut.append(zaPboDbl.records[k]['statut'])

print(26%12)