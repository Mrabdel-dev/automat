from os import listdir
from os.path import isfile, join
import os
mypath="C:/Users/etudes20/Desktop/bean"
onlyfiles = [f for f in listdir(mypath) if isfile(join(mypath, f))]
file = open('fileGenerated/names.text', 'w')

for i in range(len(onlyfiles)):
    x = onlyfiles[i].replace(".java", "")
    onlyfiles[i] = x
    file.write(x+', ')

print(onlyfiles)
