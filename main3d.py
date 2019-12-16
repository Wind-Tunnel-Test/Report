import xlrd
from matplotlib import pyplot as plt

loc1 = "corr_test3d.xlsx"
loc2 = "corr_test2d.xlsx"


wb1 = xlrd.open_workbook(loc1)
wb2 = xlrd.open_workbook(loc2)


sheet1 = wb1.sheet_by_index(0)
sheet2 = wb2.sheet_by_index(0)


# define data sets
Alpha = []
CL= []
CD = []

alpha = []
Cl = []
Cd = []


for i in range(2,47): # importing data from excel file

    Alpha.append(sheet1.cell_value(i,1))
    CL.append(sheet1.cell_value(i, 3))
    CD.append(sheet1.cell_value(i, 4))


for j in range(2,45):
    alpha.append(sheet2.cell_value(j,1))
    Cl.append(sheet2.cell_value(j,3))
    Cd.append(sheet2.cell_value(j,2))

print(Alpha, CL, CD)
print(alpha, Cl, Cd)

CLcut = CL[0:36]
CDcut = CD[0:36]

Clcut = Cl[0:30]
Cdcut = Cd[0:30]


# Lift Polar of a finite wing

plt.plot(Alpha, CL, marker='.')
plt.xlabel("Alpha [deg]")
plt.ylabel("Cl [-]")
plt.title("Lift Curve for the finite wing")
plt.grid()

plt.vlines(16, -0.25, 0.9, colors='r')
plt.text(16.1,-0.18,"stall", rotation=90)

plt.vlines(15.5, -0.25, 0.9, colors='g', linestyles='dashed')
plt.text(14.85, -0.18, "Clmax", rotation=90)

plt.vlines(14, -0.25, 0.9, colors='b', linestyles='dotted')
plt.text(14.1, -0.18, "recovery", rotation=90)
plt.savefig("liftcurve3d.png")
plt.show()

# Drag polar for a finite wing

plt.plot(CDcut, CLcut, marker='.')
plt.xlabel("Cd [-]")
plt.ylabel("Cl [-]")
plt.title("Drag polar for the finite wing")
plt.grid()

plt.savefig("dragpolar3d.png")
plt.show()

# Lift Polar of a infinite wing

plt.plot(alpha, Cl, marker='.')
plt.xlabel("Alpha [deg]")
plt.ylabel("Cl [-]")
plt.title("Lift Curve for the infinite wing")
plt.grid()

plt.vlines(15.5, -0.3, 0.95, colors='r')
plt.text(15.6,-0.18,"stall", rotation=90)

plt.vlines(12.75, -0.3, 0.95, colors='g', linestyles='dashed')
plt.text(12.85, -0.18, "Clmax", rotation=90)

plt.vlines(13.5, -0.3, 0.95, colors='b', linestyles='dotted')
plt.text(13.6, -0.18, "recovery", rotation=90)
plt.savefig("liftcurve2d.png")
plt.show()

# Drag polar for a finite wing

plt.plot(Cdcut, Clcut, marker='.')
plt.xlabel("Cd [-]")
plt.ylabel("Cl [-]")
plt.title("Drag polar for the infinite wing")
plt.grid()

plt.savefig("dragpolar2d.png")
plt.show()