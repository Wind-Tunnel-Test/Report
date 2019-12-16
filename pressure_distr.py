import xlrd
from matplotlib import pyplot as plt
import numpy as np

loc = "cp_test.xlsx"

wb = xlrd.open_workbook(loc)

sheet = wb.sheet_by_index(0)

AoA = []
data = np.zeros((44,48))
print(data)

for p in range(1,43):

    AoA.append(sheet.cell_value(1,p))

print(AoA)

#Retrieve data from Excel file

for i in range(0,43):

    for j in range(0,48):

        data[i,j] = sheet.cell_value(j+2,i)

for k in range(1,43):

    if AoA[k-1] in AoA[0:34]:

        Alpha = str(round((AoA[k-1]), 1))

        Title = 'Pressure Distribution for AoA {} degrees'.format(Alpha)
        name = 'pressuregraphAoA{}.png'.format(Alpha)

    else:

        Alpha = str(round((AoA[k-1]), 1))

        Title = 'Pressure Distribution for AoA {} degrees stalled' .format(Alpha)
        name = 'pressuregraphAoA{}stalled.png' .format(Alpha)

    #print(name)

    plt.plot(data[0], data[k], marker='.')
    plt.xlabel('Chord [%]')
    plt.ylabel('Cp [-]')
    plt.grid()
    plt.title(Title)
    plt.gca().invert_yaxis()
    plt.savefig(name)
    plt.show()
