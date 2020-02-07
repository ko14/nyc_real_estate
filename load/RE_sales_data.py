import subprocess
import os


#download Excel files
fileList = ["rollingsales_manhattan","rollingsales_brooklyn","rollingsales_queens","rollingsales_bronx","rollingsales_statenisland"]
for thisFile in fileList:
  subprocess.run("curl https://www1.nyc.gov/assets/finance/downloads/pdf/rolling_sales/" + thisFile + ".xls --output " + thisFile + ".xls")


#convert to csv
for thisFile in fileList:
  if os.path.exists(thisFile + ".csv"):
    os.remove(thisFile + ".csv")
subprocess.run("wscript XlsToCsv.vbs")


#combine into one file, ignore top verbiage and empty rows at bottom and format date
if os.path.exists("rollingsales.csv"):
  os.remove("rollingsales.csv")

f_final = open("rollingsales.csv","w")
for thisFile in fileList:
  f = open(thisFile + ".csv","r")
  lines = f.readlines()
  f.close()
  keep_going = 0
  for line in lines:                          
    if (line[:8] == "BOROUGH," and thisFile != "rollingsales_manhattan"):  #this prevents the header row from getting added with each file
      keep_going = 1
    elif (line[:8] == "BOROUGH," or keep_going == 1) and line[:1] != ",":
      if (line[:8] != "BOROUGH,"): 
        last_comma = line.rfind(',')
        new_date = line[last_comma+1:len(line)-1].split('/')
        if (len(new_date[0]) == 1):
          new_date[0] = "0" + new_date[0]
        if (len(new_date[1]) == 1):
          new_date[1] = "0" + new_date[1]         
        new_date_final = "," + new_date[2] + new_date[0] + new_date[1] + " \n"
        line = line[0:last_comma] + new_date_final
      f_final.write(line)
      keep_going = 1
f_final.close()  
