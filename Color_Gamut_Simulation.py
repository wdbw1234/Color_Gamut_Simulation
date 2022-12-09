# colour gamut calculation, Apr 05 2022

import numpy as np
import xlrd
import matplotlib.pyplot as plt


# data input
form = xlrd.open_workbook('./CIE1931.xls')
sheet_absorbance = form.sheet_by_index(0)
sheet_matching = form.sheet_by_index(1)
wavelength_0 = np.array(sheet_absorbance.col_values(0))
absorbance_ha = np.array(sheet_absorbance.col_values(1))
absorbance_ph = np.array(sheet_absorbance.col_values(3))
absorbance_na = np.array(sheet_absorbance.col_values(5))
matching_x = np.array(sheet_matching.col_values(1))
matching_y = np.array(sheet_matching.col_values(2))
matching_z = np.array(sheet_matching.col_values(3))
incident = np.ones(len(wavelength_0))

# normalizing absorbance
n_ha = 1.0 / np.max(absorbance_ha)
absorbance_ha = absorbance_ha * n_ha
n_ph = 1.0 / np.max(absorbance_ph)
absorbance_ph = absorbance_ph * n_ph
n_na = 1.0 / np.max(absorbance_na)
absorbance_na = absorbance_na * n_na


# function calculating total OD function
def od(a1, a2, a3, c1, c2, c3):
    return c1 * a1 + c2 * a2 + c3 * a3


# function calculating XYZ
def XYZ(wvl, i, ts, mx, my, mz):
    d_lambda = wvl[1] - wvl[0]
    N = np.sum(my * i * d_lambda)
    X = (1.0 / N) * np.sum(mx * ts * i * d_lambda)
    Y = (1.0 / N) * np.sum(my * ts * i * d_lambda)
    Z = (1.0 / N) * np.sum(mz * ts * i * d_lambda)
    return X, Y, Z


# function calculating cie xyz
def cie_xyz(X, Y, Z):
    cx = X / (X + Y + Z)
    cy = Y / (X + Y + Z)
    cz = Z / (X + Y + Z)
    return cx, cy, cz


# main function
if __name__ == '__main__':
    point_x = []
    point_y = []
    list_i = []
    list_j = []
    list_k = []
    for i in range(2):
        for j in range(2):
            for k in range(2):
                c_1 = i * 1
                c_2 = j * 1
                c_3 = k * 1
                A = od(absorbance_ha, absorbance_ph, absorbance_na, c_1, c_2, c_3)
                T = 10 ** (-3 * A)
                X, Y, Z = XYZ(wavelength_0, incident, T, matching_x, matching_y, matching_z)
                x, y, z = cie_xyz(X, Y, Z)
                point_x.append(x)
                point_y.append(y)
                list_i.append(i)
                list_j.append(j)
                list_k.append(k)
    px = np.array(point_x)
    py = np.array(point_y)
    print("HA:" + str(list_i))
    print("PH:" + str(list_j))
    print("NA:" + str(list_k))
    print(point_x)
    print(point_y)

    plt.figure()
    plt.plot(px, py, 'o', color='crimson')
    plt.show()
