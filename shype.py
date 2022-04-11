import shapefile

sf = shapefile.Reader('89_042_031_EXE_CCE_SIO_MAN_001_A/89_042_031_BOITE_OPTIQUE_A.shx ')


records = sf.shapeRecords()
print(records[1].record[0])
for i in range (0,40):
    print(records[i].record['NOM'])
# w = shapefile.Writer('shapefiles/test/dtype')
# w.field('TEXT', 'C')
# w.field('SHORT_TEXT', 'C', size=5)
# w.field('LONG_TEXT', 'C', size=250)
# w.null()
# w.record('Hello', 'World', 'World')
# w.record('abdel','abdel','ttttt')
# w.close()

