import xlrd
from matplotlib import pyplot as plt

loc1 = "corr_test3d.xlsx"
loc2 = "corr_test2d.xlsx"
loc3 = "xfoil_results.xlsx"
loc4 = "xflr_results.xlsx"


wb1 = xlrd.open_workbook(loc1)
wb2 = xlrd.open_workbook(loc2)
wb3 = xlrd.open_workbook(loc3)
wb4 = xlrd.open_workbook(loc4)


sheet1 = wb1.sheet_by_index(0)
sheet2 = wb2.sheet_by_index(0)
sheet3 = wb3.sheet_by_index(0)
sheet4 = wb4.sheet_by_index(0)

# define data sets
Alpha1 = []
CL1= []
CD1 = []
CM1 = []

Alpha2 = []
CL2= []
CD2 = []
CM2 = []

alpha1 = []
Cl1 = []
Cd1 = []
Cm1 = []

alpha2 = []
Cl2 = []
Cd2 = []
Cm2 = []



for i in range(2,47): # importing data from excel file

    Alpha1.append(sheet1.cell_value(i,1))
    CL1.append(sheet1.cell_value(i, 3))
    CD1.append(sheet1.cell_value(i, 4))
    CM1.append(sheet1.cell_value(i,11))


for j in range(2,45):

    alpha1.append(sheet2.cell_value(j,1))
    Cl1.append(sheet2.cell_value(j,3))
    Cd1.append(sheet2.cell_value(j,2))
    Cm1.append(sheet2.cell_value(j,4))

for k in range(8,44):

    Alpha2.append(sheet4.cell_value(k, 0))
    CL2.append(sheet4.cell_value(k, 1))
    CD2.append(sheet4.cell_value(k, 4))
    CM2.append(sheet4.cell_value(k, 6))

for p in range(11,49):

    alpha2.append(sheet3.cell_value(p, 0))
    Cl2.append(sheet3.cell_value(p, 1))
    Cd2.append(sheet3.cell_value(p, 2))
    Cm2.append(sheet3.cell_value(p, 4))

print(Alpha1, CL1, CD1)
print(alpha1, Cl1, Cd1)
print(Alpha2, CL2, CD2)
print(alpha2, Cl2, Cd2)

CLcut1 = CL1[0:36]
CDcut1 = CD1[0:36]

Clcut1 = Cl1[0:29]
Cdcut1 = Cd1[0:29]


# Lift Polar of a finite wing

plt.plot(Alpha1, CL1, marker='.', color='b')
plt.plot(Alpha2, CL2, marker='^', color='r')
plt.xlabel("Alpha [deg]")
plt.ylabel("Cl [-]")
plt.title("Lift Polar for the 3D case")
plt.grid()

plt.vlines(16, -0.25, 0.9, colors='r')
plt.text(16.1,-0.18,"stall", rotation=90)

plt.vlines(15.5, -0.25, 0.9, colors='g', linestyles='dashed')
plt.text(14.85, -0.18, "Clmax", rotation=90)

plt.vlines(14, -0.25, 0.9, colors='b', linestyles='dotted')
plt.text(14.1, -0.18, "recovery", rotation=90)
plt.savefig("liftcurve3d.png")
plt.show()

# Moment polar for a finite wing

plt.plot(Alpha1, CM1, marker='.', color='b')
plt.plot(Alpha2, CM2, marker='^', color='r')
plt.xlabel("Alpha [deg]")
plt.ylabel("Cm [-]")
plt.title("Moment Polar for the 3D case")
plt.grid()

plt.savefig("momentpolar3d.png")
plt.show()

# Drag polar for a finite wing

plt.plot(CDcut1, CLcut1, marker='.', color='b')
plt.plot(CD2, CL2, marker='^', color='r')
plt.xlabel("Cd [-]")
plt.ylabel("Cl [-]")
plt.title("Drag Polar for the 3D case")
plt.grid()

plt.savefig("dragpolar3d.png")
plt.show()

# Drag curve for a finite wing

plt.plot(Alpha1, CD1, marker='.', color='b')
plt.plot(Alpha2, CD2, marker='^', color='r')
plt.xlabel("Alpha [deg]")
plt.ylabel("CD [-]")
plt.title("Drag curve for the 3D case")
plt.grid()

plt.vlines(16, 0, 0.25, colors='r')
plt.text(16.1, 0.01,"stall", rotation=90)

plt.vlines(15.5, 0, 0.25, colors='g', linestyles='dashed')
plt.text(14.85, 0.01, "Clmax", rotation=90)

plt.vlines(14, 0, 0.25, colors='b', linestyles='dotted')
plt.text(14.1, 0.01, "recovery", rotation=90)

plt.savefig("dragcurve3d.png")
plt.show()



# Lift Polar of a infinite wing

plt.plot(alpha1, Cl1, marker='.', color='b')
plt.plot(alpha2, Cl2, marker='^', color='r')
plt.xlabel("Alpha [deg]")
plt.ylabel("Cl [-]")
plt.title("Lift Polar for the 2D case")
plt.grid()

plt.vlines(15.5, -0.3, 0.95, colors='r')
plt.text(15.6,-0.18,"stall", rotation=90)

plt.vlines(12.75, -0.3, 0.95, colors='g', linestyles='dashed')
plt.text(12.85, -0.18, "Clmax", rotation=90)

plt.vlines(13.5, -0.3, 0.95, colors='b', linestyles='dotted')
plt.text(13.6, -0.18, "recovery", rotation=90)
plt.savefig("liftcurve2d.png")
plt.show()

# Drag curve for the 2D case

plt.plot(alpha1, Cd1, marker='.', color='b')
plt.plot(alpha2, Cd2, marker='^', color='b')
plt.xlabel("Alpha [deg]")
plt.ylabel("Cd [-]")
plt.title("Drag Curve for the 2D case")
plt.grid()

plt.vlines(15.5, -0.02, 0.35, colors='r')
plt.text(15.6, -0.02,"stall", rotation=90)

plt.vlines(12.75, -0.02, 0.35, colors='g', linestyles='dashed')
plt.text(12.85, -0.02, "Clmax", rotation=90)

plt.vlines(13.5, -0.02, 0.35, colors='b', linestyles='dotted')
plt.text(13.6, -0.02, "recovery", rotation=90)
plt.savefig("dragcurve2d.png")
plt.show()

# Drag polar for a finite wing

plt.plot(Cdcut1, Clcut1, marker='.', color='b')
plt.plot(Cd2, Cl2, marker='^', color='r')
plt.xlabel("Cd [-]")
plt.ylabel("Cl [-]")
plt.title("Drag Polar for the 2D case")
plt.grid()

plt.savefig("dragpolar2d.png")
plt.show()

# Moment polar for a finite wing

plt.plot(alpha1, Cm1, marker='.', color='b')
plt.plot(alpha2, Cm2, marker='^', color='r')
plt.xlabel("alpha [deg]")
plt.ylabel("Cm [-]")
plt.title("Moment Polar for the 2D case")
plt.grid()

plt.savefig("momentpolar2d.png")
plt.show()