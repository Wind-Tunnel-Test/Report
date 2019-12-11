import xlrd
from matplotlib import pyplot as plt

loc = "corr_test.xlsx"

wb = xlrd.open_workbook(loc)

sheet = wb.sheet_by_index(0)

# define data sets
Alpha = []
Cl = []
Cd = []


for i in range(2,47): # importing data from excel file

    Alpha.append(sheet.cell_value(i,1))
    Cl.append(sheet.cell_value(i, 3))
    Cd.append(sheet.cell_value(i, 4))


print(Alpha, Cl, Cd)


plt.plot(Alpha, Cl)
plt.xlabel("Alpha [deg]")
plt.ylabel("Cl [-]")
plt.title("Lift Curve")
plt.savefig("liftcurve.png")
plt.show()