#downoald and modfier a web content
import urllib.request, urllib.error, urllib.parse

url = 'https://www.sap.com/belgique/index.html'

response = urllib.request.urlopen(url)
webContent = response.read()
x = str(webContent)
file = open('fileGenerated/abdel.html', 'w')
file.write(x)
#print(webContent)
fin = open("fileGenerated/abdel.html", "rt")
#output file to write the result to
fout = open("fileGenerated/sami.html", "wt")
#for each line in the input file
oldtext=("SAP","Sap","sap")

for line in fin:
    for k in oldtext:
        fout.write(line.replace(k, 'Odoo'))
#close input and output files
fin.close()
fout.close()