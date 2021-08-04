from openpyxl import load_workbook

# load the THE ROUTE FILE
workbook = load_workbook('fileGenerated/SRO-51-001-334-ROUTAGE OPTIQUE-DOE-15-12-2020 MODIFIER.xlsx')
rout = workbook.active
# load the ROP FILE
workbook1 = load_workbook('fileGenerated/ROP-SRO-51-001-334-03_08_21.xlsx')
rop = workbook1.active

maxRowRoute = rout.max_row + 1
maxColRoute = rout.max_column + 1
result = []
k=2
for r in range(2, maxRowRoute):
    if r==182:
        k=182
    teroir = str(rop.cell(k, 3).value)[-1]
    boiter = str(rop.cell(k, 2).value).strip()
    fibrer = str(rop.cell(k, 7).value)
    tuber = str(rop.cell(k, 6).value)
    for c in range(13, maxColRoute):
        val = str(rout.cell(r, c).value)
        teroi = str(rout.cell(r, 5).value)[-1]
        if teroi == "E":
            break
        if val == "A STOCKER":
            boite = str(rout.cell(r, c - 1).value).strip()
            fibre = str(rout.cell(r, c - 3).value)
            tube = str(rout.cell(r, c - 4).value)
            k += 1
            if boite == boiter and tube == tuber and fibre == fibrer and teroi == teroir:

                break
            else:
                result.append(str(f"la line {r,k-1} et teroire  {teroi, teroir} et la fibre {fibre, fibrer} et tube {tube, tuber}  dans la boite => {boite, boiter}"))

                break

for k in range(0, len(result)):
    print(result[k])
