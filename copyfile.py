from os import walk

from openpyxl import load_workbook
from openpyxl.drawing.image import Image

# load the C6 book
rep = 'C:/Users/etudes20/Desktop/photos/'
monRepertoire = r'C:/Users/etudes20/Desktop/Etude 2 f/'
listeFichiers = []
for (repertoire, sousRepertoires, fichiers) in walk(monRepertoire):
    listeFichiers.extend(fichiers)
i = 0
ext = '.JPG'
while i < len(listeFichiers):
    workbook = load_workbook(monRepertoire + listeFichiers[i])
    x = str(listeFichiers[i])[11:-5]
    print(x)
    BookC6 = workbook[x]
    try:
        img = Image(rep + x+"_1" + ext)
        img.height = 308
        img.width = 360
        BookC6.add_image(img, "A" + str(62))
        img2 = Image(rep + x+"_2" + ext)
        img2.height = 308
        img2.width = 360
        BookC6.add_image(img2, "M" + str(62))
    except FileNotFoundError:
        print(x)
    workbook.save(listeFichiers[i])
    i += 1
