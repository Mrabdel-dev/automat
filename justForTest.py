# from openpyxl import load_workbook
# #looad your pds file here
# pds =load_workbook('PDS.xlsx')
# sheet =pds.active
# w= pds.sheetnames


#create the new distination xcel file
# import xlsxwriter
#
# workbook = xlsxwriter.Workbook('hard.xlsx')
# w = workbook.add_worksheet()
# ############# save the same formating text that exsicte in pds file #####################
# bold = workbook.add_format({'bold': True,"border":1})
# border = workbook.add_format({"border":1})
# cassette = workbook.add_format({"bg_color": '#A9A9A9' ,"border":1})
# cell_format0 = workbook.add_format({"bg_color": '#E6E6FA',"border":1})
# cell_format = workbook.add_format({"bg_color": 'red',"border":1})
# cell_format1 = workbook.add_format({"bg_color": 'blue',"border":1})
# cell_format2 = workbook.add_format({"bg_color": '#00FF00',"border":1})
# cell_format3 = workbook.add_format({"bg_color": 'yellow',"border":1})
# cell_format4 = workbook.add_format({"bg_color": '#BF00FF',"border":1})
# cell_format5 = workbook.add_format({"bg_color": 'white',"border":1})
# cell_format6 = workbook.add_format({"bg_color": '#FFBF00',"border":1})
# cell_format7 = workbook.add_format({"bg_color": '#828282',"border":1})
# cell_format8 = workbook.add_format({"bg_color": '#816B56',"border":1})
# cell_format9 = workbook.add_format({"bg_color": '#333333',"border":1})
# cell_format10 = workbook.add_format({"bg_color": '#00FFBF',"border":1})
# cell_format11 = workbook.add_format({"bg_color": '#FFAAD4',"border":1})
# ############### the heder for all infromation should include down of it ##########################
# w.write('A1', 'Entrée', bold)
# w.write('B1', 'Id', bold)
# w.write('C1', 'Capacité', bold)
# w.write('D1', 'N°         ', bold)
# w.write('E1', 'N° Tube', bold)
# w.write('F1', 'N° Fibre', bold)
# w.write('G1', 'Cassette', bold)
# w.write('H1', 'Etat fibre', bold)
# w.write('I1', 'N° Fibre', bold)
# w.write('J1', 'N° Tube', bold)
# w.write('K1', 'N°       ', bold)
# w.write('L1', 'Capacité', bold)
# w.write('M1', '', bold)
# w.write('N1', 'Sortie', bold)
# w.write('O1', 'Statut', bold)
# w.write('P1', 'Client', bold)
#
# workbook.close()
import numpy as np
tab = np.array([[1,2,6], [4,5,6]])
print(tab)