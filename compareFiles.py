from openpyxl import load_workbook

# load the THE ROUTE FILE
workbook = load_workbook('fileGenerated/RO-SRO-51-001-335 (22-06-2021).xlsx')
rout = workbook.active
# load the ROP FILE
workbook1 = load_workbook('fileGenerated/ROP-SRO-51-001-335-06_08_21.xlsx')
rop = workbook1.active

maxRowRoute = rout.max_row + 1
maxColRoute = rout.max_column + 1
result = []
k=2
for r in range(2, maxRowRoute):
    # if r==182:
    #     k=182
    teroir = str(rop.cell(r, 3).value)[-1]
    boiter = str(rop.cell(r, 2).value).strip()
    fibrer = str(rop.cell(r, 7).value)
    tuber = str(rop.cell(r, 6).value)
    for c in range(13, maxColRoute):
        val = str(rout.cell(r, c).value)
        teroi = str(rout.cell(r, 5).value)[-1]
        if teroi == "E":
            break
        if val == "A STOCKER":
            boite = str(rout.cell(r, c - 1).value).strip()
            fibre = str(rout.cell(r, c - 3).value)
            tube = str(rout.cell(r, c - 4).value)

            if boite == boiter and tube == tuber and fibre == fibrer and teroi == teroir:

                break
            else:
                result.append(str(f"la line {r} et teroire  {teroi, teroir} et la fibre {fibre, fibrer} et tube {tube, tuber}  dans la boite => {boite, boiter}"))

                break

for k in range(0, len(result)):
    print(result[k])
