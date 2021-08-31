from click import edit
from qgis.core import QgsVectorLayer

path_to_my_layer = "C:/Users/etudes20/Desktop/89_013_022_EXE_CCE_SIO_MAN_001_A/89_013_022_SUPPORT_A.shx"
supptable = QgsVectorLayer(path_to_my_layer, "SUPPORT", "ogr")
path_to_my_layer1 = "C:/Users/etudes20/Desktop/89_013_022_EXE_CCE_SIO_MAN_001_A/89_013_022_POINT_TECHNIQUE_A.shx"
pointTable = QgsVectorLayer(path_to_my_layer1, "POINT TECHNIQUE", "ogr")

path_to_my_layer2 = "C:/Users/etudes20/Desktop/89_013_022_EXE_CCE_SIO_MAN_001_A/89_013_022_SITES_A.shx"
sroTable = QgsVectorLayer(path_to_my_layer2, "SITES", "ogr")
pointName = []
pointCord = []


def getPointIndex(cord):
    return pointCord.index(cord)


for p in pointTable.getFeatures():
    nom = p["NOM"]
    pointName.append(nom)
    cord_X = p["COORDONE_X"]
    pointCord.append(cord_X)
sroCordone = 0.0
for s in sroTable.getFeatures():
    sroCordone = s["X_SRO"]

sro = '89-013-022'

with edit(supptable):
    for f in supptable.getFeatures():
        start = f["X_START"]
        end = f["X_END"]
        try:
            if start > sroCordone:
                i = getPointIndex(start)
                f["AMONT"] = pointName[i]
                j = getPointIndex(end)
                f["AVAL"] = pointName[j]
            else:
                i = getPointIndex(start)
                f["AVAL"] = pointName[i]
                j = getPointIndex(end)
                f["AMONT"] = pointName[j]
        except ValueError:
            f["AVAL"] = '/'
            f["AMONT"] = '/'
        strt = str(f["STRUCTURE"])
        amont = f["AMONT"]
        aval = f["AVAL"]
        f['TYPE_STRUC'] = "TELECOM"
        f['DISPONIBLE'] = "OK"
        if amont.startswith(("C", "P", "F")) and aval.startswith(("C", "P", "F")):
            f["NOM"] = strt + '-' + sro + '-' + amont[10:] + '-' + aval[10:]
        elif amont.startswith(("C", "P", "F")):
            f["NOM"] = strt + '-' + sro + '-' + amont[10:] + '-' + aval[6:]
        elif aval.startswith(("C", "P", "F")):
            f["NOM"] = strt + '-' + sro + '-' + amont[6:] + '-' + aval[10:]
        else:
            f["NOM"] = strt + '-' + sro + '-' + amont[6:] + '-' + aval[6:]
        nom = f["NOM"]
        f['CODE'] = nom
        supptable.updateFeature(f)
