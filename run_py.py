import subprocess
import os


#Count the xlsx files to run loop x times
all_xlsx_files = []
for x in os.listdir():
    if x.endswith(".xlsx"):
        if not x[:2] == "~$":
            all_xlsx_files.append(x)

for x in range(len(all_xlsx_files)):
    subprocess.call("py test.py", shell=True)
    print("Done")
