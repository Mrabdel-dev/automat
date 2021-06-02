from dbfread import DBF
# load the both file boite and cable in DBF format
cableTable = DBF('fileGenerated/069_CABLE_OPTIQUE.dbf', load=True)
boiteTable = DBF('fileGenerated/21_011_069_BOITE_OPTIQUE_B.dbf')
# charge the name of all filed in tables
filedCableNam = cableTable.field_names
filedBoiteNam = boiteTable.field_names
boiteLen = len(boiteTable)
cableLen = len(cableTable)
boiteCode = [] ; boiteCable = [] ; boiteCableState = [] ; boiteReference = [] ; nbf = []
cableName = [] ; cableOrigin = [] ; cableExtremity = [] ; cableCapacity = []

for i in range(0, cableLen):
       x = cableTable.records[i]['NOM']
       cableName.append(x)
       cableOrigin.append(cableTable.records[i]['ORIGINE'])
       cableExtremity.append(cableTable.records[i]['EXTREMITE'])
       cableCapacity.append(cableTable.records[i]['CAPACITE'])
# for j in range(0, boiteLen):
#        y = boiteTable.records[j]['NOM']
#        boiteCode.append(y)
#        boiteCable.append(boiteTable.records[i]['AMONT'])
#        boiteCableState.append(boiteTable.records[i]['INTERCO'])
#        boiteReference.append(boiteTable.records[i]['REFERENCE'])

print(cableOrigin)