from PIL import Image
from os import walk
import os

# the folder source that you want take resize inside it
monRepertoire ='C:/Users/etudes20/Desktop/Nouveau dossier/'
# the folder out that u want to put new images in it
sortRepetoire = 'C:/Users/etudes20/Desktop/photo/'

listeFichiers = []
dectName = {"vue_ensemble":"1","vue_tete":"2","complementaire_1":"3","vue_portee":"4","complementaire_2":"5","complementaire_3":"6","complementaire_4":"7","vue_etiquette":"8"}
for (repertoire, sousRepertoires, fichiers) in walk(monRepertoire):
    listeFichiers.extend(sousRepertoires)
for i in listeFichiers:
    print(i)
    listeFichier1 =[]
    reper = monRepertoire +str(i)+'/'
    for (repertoire, sousRepertoires,fichiers) in walk(reper):
        listeFichier1.extend(fichiers)
    for j in listeFichier1:
        ext = str(j)[-4:]
        x= str(j)[:-4]
        os.rename(reper+j,sortRepetoire+str(i)+"_"+str(dectName[x])+ext)




