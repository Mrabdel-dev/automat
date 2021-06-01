# ########################################csv file dealing with csv module #########################################################
import  csv
import json
with open('RR.csv','rt')as f:
  data = csv.DictReader(f)
  print(data.fieldnames)

  # newData = json.dumps(data)
  # print(newData)


# with open('data.csv', mode='a') as file:
#     writer = csv.writer(file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
#
#     #way to write to csv file
#     writer.writerow(['Programming language', 'Designed by', 'Appeared', 'Extension'])
#     writer.writerow(['Python', 'Guido van Rossum', '1991', '.py'])
#     writer.writerow(['Java', 'James Gosling', '1995', '.java'])
#     writer.writerow(['C++', 'Bjarne Stroustrup', '1985', '.cpp'])
#########################################################csv file dealing with pandas ###################################################
import pandas

#open the csv file
