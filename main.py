import subprocess
import os
import filecmp
import openpyxl as xl
from openpyxl.styles import Alignment,Font

#excel part
wb = xl.load_workbook("program_analysis.xlsx")
sheet = wb['Sheet1']
comp_count=0
out_count=0

def path_file():
    program_file = subprocess.check_output("realpath program_folder", shell=True, text=True)
    length3 = len(program_file)
    path=program_file[:length3-1]
    return [os.path.join(path, x) for x in os.listdir(path)]

def coordinates(val):
    for row1 in range(5,sheet.max_row+1):
        cell=sheet.cell(row1,2)
        if cell.value == val:
           return row1

lst = path_file()
for i in lst:
    #print(i)
    string2 = i[-15:-4]
    row = coordinates(int(string2))
    if row > sheet.max_row-2:
        break
    p1 = subprocess.run(f"g++ {i}", shell=True, capture_output=True)
    data, temp = os.pipe()
    if (p1.returncode == 0):
        #print("no compilation error")
        cell=sheet.cell(row,3)
        cell.value="✔"
        cell.alignment = Alignment(horizontal='center')

        with open("output.txt", 'w') as f:
            input_file = subprocess.check_output("realpath input.txt", shell=True, text=True)
            length2 = len(input_file)
            file = open(input_file[:length2-1], 'r')
            message = file.read()
            file.close()
            os.write(temp, bytes(message, "utf-8"))
            os.close(temp)
            process = subprocess.run("./a.out", check=True, stdin=data, shell=True, stdout=f, universal_newlines=True)
            out = subprocess.check_output("realpath output.txt", shell=True, text=True)
            length = len(out)
            file1 = out[:length - 1]
            output_file = subprocess.check_output("realpath output_original.txt", shell=True, text=True)
            length1= len(output_file)
            file2 = output_file[:length1-1]
            comp = filecmp.cmp(file1, file2, shallow=False)
            if (comp == True):
                cell=sheet.cell(row,4)
                cell.value="✔"
                cell.alignment = Alignment(horizontal='center')
                #print("Program is correct")
            else:
                cell=sheet.cell(row,4)
                cell.value="❌"
                cell.alignment = Alignment(horizontal='center')
                #print("Program is incorrect")
                out_count+=1
    else:
        cell=sheet.cell(row,3)
        cell.value="❌"
        cell.alignment = Alignment(horizontal='center')
        cell = sheet.cell(row, 4)
        cell.value = "❌"
        cell.alignment = Alignment(horizontal='center')
        #print("compilation error")
        comp_count+=1

cell = sheet.cell(sheet.max_row,3)
cell.value = comp_count
cell.alignment=Alignment(horizontal='center')
cell.font=Font(name='Calibri',size=10.5,bold=True)
cell = sheet.cell(sheet.max_row,4)
cell.value = out_count
cell.alignment=Alignment(horizontal='center')
cell.font=Font(name='Calibri',size=10.5,bold=True)
wb.save('program_analysis.xlsx')
print("Process Completed 🙂️")


