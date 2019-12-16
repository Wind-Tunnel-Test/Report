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

4
print(Alpha, Cl, Cd)


plt.plot(Alpha, Cl, marker='.')
plt.xlabel("Alpha [deg]")
plt.ylabel("Cl [-]")
plt.title("Lift Curve")
plt.grid()

plt.vlines(16, -0.25, 0.9, colors='r')
plt.text(16.1,-0.18,"stall", rotation=90)

plt.vlines(15.5, -0.25, 0.9, colors='g')
plt.text(14.9, -0.18, "Clmax", rotation=90)

plt.vlines(14, -0.25, 0.9, colors='b')
plt.text(14.1, -0.18, "Recovery", rotation=90)
plt.savefig("liftcurve.png")
plt.show()