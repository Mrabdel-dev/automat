import os
from os import walk
from dbfread import DBF

pointTechTable = DBF('C:/Users/etudes20/Desktop/RELEVES_ENEDIS/85_076_812_APPUIS_ENEDIS_A.dbf', load=True,
                     encoding='iso-8859-1')
pointNom = []
pointlen = len(pointTechTable)
for p in range(0, pointlen):
    x = str(pointTechTable.records[p]['NOM'])[6:8] + "00" + str(pointTechTable.records[p]['NOM'])[8:]
    pointNom.append(x)

# the folder source that you want take resize inside it
monRepertoire = r'C:/Users/etudes20/Desktop/test/'
# the folder out that u want to put new images in it
sortRepetoire = r'C:/Users/etudes20/Desktop/test2/'
print(pointNom)
listeFichiers = []
for (repertoire, sousRepertoires, fichiers) in walk(monRepertoire):
    listeFichiers.extend(fichiers)
i = 0
while i < len(listeFichiers):
    x = str(listeFichiers[i])[0:7]
    print(x)
    for k in pointNom:
        if x == k:
            os.rename(monRepertoire + listeFichiers[i], sortRepetoire + listeFichiers[i])
        else:
            continue

    i += 1