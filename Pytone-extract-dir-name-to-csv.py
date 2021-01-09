import os
import sys
import xlwt
from xlwt import Workbook

### -------------------------- CheckFileNumSequence ----------------------------
### Check all file numbers are in sequence
def CheckFileNumSequence(dir_list):
    numbersNotInSequence = False
    prevNum = -2 ### a control value
    for objNum in range(len(dir_list)):
        currNum = MatchStringPattern(dir_list[objNum])
        if currNum == -1 or dir_list[objNum].endswith('full'):
            continue

        if prevNum != -2 and (prevNum != (currNum - 1)):
            numbersNotInSequence = True
            print("File numbers are not in sequence: prevNum=" + str(prevNum) +
                ", currNum=" + str(currNum))

        prevNum = currNum

    if numbersNotInSequence == False:
        print("All file numbers are in sequence\n\n")


### ---------------------------- MatchStringPattern ----------------------------
### does the directory name match the pattern?
### 1) a number followed by "-"; it may have one or more spaces before or after "-"
### 2) a number[a,b,c,...] followed by "-"
### return either the number or -1 if it does not match the pattern
def MatchStringPattern(path):
    rc = -1
    if "-" in path:
        arr = path.split("-", 1)
        arr0 = arr[0].strip()
        ### it is an numeric string or last char is not (such as 100a)
        if arr0.isnumeric():
            rc = int(arr0)
        elif arr[0][:-1].isnumeric():
            rc = int(arr0[:-1])
    return rc


### ---------------------------- HandleOneDirectory ----------------------------
def HandleOneDirectory(path, dir_list):
    '''Handle one direcotry and checks for the pattern'''
    objName = path.split("\\")[-1] ### get the file name from the full path
    if MatchStringPattern(objName) > 0:
       ###print(objName)
       dir_list.append(objName)

### ----------------------------- ProcessDirectories ---------------------------
### Process the whole directory and all it's subdirectories
def ProcessDirectories(path, dir_list):
    arr = os.listdir(path)
    for obj in arr:
        entry_name=path+"\\"+obj
        if os.path.isdir(entry_name):
            print(entry_name)
            if ("RECYCLE.BIN" in obj) or ("System Volume Information" in obj):
                continue
            HandleOneDirectory(entry_name, dir_list)
            ProcessDirectories(entry_name, dir_list)
        else:
            continue ### don't need to process file


### -------------------------------- OutputToCsv -------------------------------
### Output the list to a csv file and it should able to handle multi-bytes characterset
def OutputToCsv(path, dir_list):

    dir_list.sort() ### sort it in ascending order
    with open(path + "\\output.csv", "w", encoding='utf-8') as fh:
        for i in range(len(dir_list)):
            arr = dir_list[i].split("-", 1)
            ### because of csv, replace ',' with ';'
            fh.write(arr[0].strip() + ", " + arr[1].strip().replace(',', ';') + "\n")


### ------------------------------- OutputToExcel ------------------------------
### Output the list to a excel file
def OutputToExcel(path, dir_list):

    wb = Workbook()
    sheet1 = wb.add_sheet('Sheet 1')

    dir_list.sort() ### sort it in ascending order
    for i in range(len(dir_list)):
        arr = dir_list[i].split("-", 1)
        arr0 = arr[0].strip()
        if arr0.isnumeric():
            sheet1.write(i, 0, int(arr0))
        else:
            sheet1.write(i, 0, arr0)
        sheet1.write(i, 1, arr[1].strip())

    wb.save(path + "\\output.xls")


### ---------------------------------- NumDirs ---------------------------------
### Enumerate all directories and subdirectories that match the pattern
def NumDirs():
    print("Welcome to Python NumDirs Utility")
    print("This utility will enumerate all dirtories and sub-directories")
    print("and outputs to a csv and/or Excel file.\n")

    if (len(sys.argv)-1):
       path = sys.argv[1]
    else: ### get the working FOLDER path
        print("Please enter the path name : .\\ ? ", end="")
        path=input()

    if len(path) == 0:
        path = os.getcwd()
    if os.path.isdir(path) == False:
        print("\n\n(" + path + ") is an invalid path name!!!\n\n")
        exit()

    dir_list = []
    HandleOneDirectory(path, dir_list) ### take care the root directory first
    ProcessDirectories(path, dir_list)

    dir_list.sort()
    ### OutputToCsv(path, dir_list)
    OutputToExcel(path, dir_list)
    CheckFileNumSequence(dir_list)

### main
NumDirs()



