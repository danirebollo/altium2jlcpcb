# python altium2jlcpcb.py

# instructions:
# - Use altium outputjob
# - Copy script into "project outputs" folder. Where gerber files are
# execute py script: "python altium2jlcpcb.py" or run altium2pcb.bat in windows

# requisites:
# pip install pyexcel pyexcel-xlsx

# # Manual work without altium outputjob
#-(with PCB open) File, fabrication, gerber
#-(with PCB open) File, fabrication, nc drill
#-(with PCB open) File, assembly, generates p&p

# automated work:
# # p&p 
#-remove first lines
#-change "," for ;
#-change " for nothing
#-change , for ;
#-change TopLayer for Top
#-change BottomLayer for Bottom
#-change Center-X(mm) for Mid X
#-change Center-Y(mm) for Mid Y
# # bom
#-change "," for ;
#-change " for nothing
#-change LCSC for JLCPCB Part #
# # save to XLS

from zipfile import ZipFile
import pandas as pd
import glob
import os
import shutil

cwd = os.getcwd()

def addtozip(extensionarray):
    for extension in extensionarray:
        print('reading extension: ',extension)
        myextension='*.'+extension
        namearray=glob.glob(myextension)
        for name in namearray:
            zipObj.write(name)


print('running altium2jlcpcb script')
print('found files: ')
print(glob.glob('*.GBL'))

zipObj = ZipFile('gerber.zip', 'w')
addtozip(["GBL","GBO","GBP","GBS","GKO","GTL","GTO","GTP","GTS"])


zipObj.close()

#################################################
# replace in BOM

# Read in the file
with open('BOM.csv', 'r') as file :
  filedata = file.read()

# Replace the target string
#filedata = filedata.replace('","', ';')
#filedata = filedata.replace('"', '')
filedata = filedata.replace('LCSC', 'JLCPCB Part #')

# Write the file out again
with open('BOM.csv', 'w') as file:
  file.write(filedata)

#################################################
# replace in PP

# Read in the file
with open('PP.csv', 'r') as file :
  filedata = file.read()

# Replace the target string
#filedata = filedata.replace('","', ';')
#filedata = filedata.replace('"', '')
filedata = filedata.replace('TopLayer', 'Top')
filedata = filedata.replace('BottomLayer', 'Bottom')
filedata = filedata.replace('Center-X(mm)', 'Mid X')
filedata = filedata.replace('Center-Y(mm)', 'Mid Y')

# Write the file out again
with open('PP.csv', 'w') as file:
  file.write(filedata)


with open("PP.csv", "r+") as file:
    first_line = file.readline().rstrip()

if first_line=="Altium Designer Pick and Place Locations":
    print("CAUTION, BAD CSV. Trying to fix")
    with open("PP.csv", "r+") as file:
        all_lines = file.readlines()
        # move file pointer to the beginning of a file
        file.seek(0)
        # truncate the file
        file.truncate()
        # start writing lines except the first line
        file.writelines(all_lines[12:])
else:

    print("CSV is OK. First line: '",first_line,"'")

read_file = pd.read_csv ('BOM.csv')
read_file.to_excel ('BOM.xlsx', index = None, header=True)

#remove manually first non csv text...
read_file = pd.read_csv ('PP.csv')
read_file.to_excel ('PP.xlsx', index = None, header=True)

path = cwd+"\export"

if os.path.isdir('./export'):
    for f in os.listdir(path):
        os.remove(os.path.join(path, f))
    os.rmdir(path)

os.makedirs(path, exist_ok=False)

os.rename(cwd+"\gerber.zip", cwd+"\export\gerber.zip")
os.rename(cwd+"\BOM.xlsx", cwd+"\export\BOM.xlsx")
os.rename(cwd+"\PP.xlsx", cwd+"\export\PP.xlsx")