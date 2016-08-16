# -*- coding: utf-8 -*-
"""

"""
#import openpyxl
import numpy as np
import numpy.polynomial.polynomial as poly
import scipy
from scipy import integrate



#wb = openpyxl.load_workbook('example.xlsx')
#s1=wb.get_sheet_by_name('Sheet1')

time=[0, 10, 20, 30, 40, 50, 60, 80, 100, 120, 150, 180, 240]

rat1=[458,	380,	317,	316,	251,	262,	241,	239,	281,	296,	334,	342, 298]

coefs=poly.polyfit(time,rat1,4)
ffit=poly.polyval(time, coefs)
ins_func=lambda x: coefs[0]+coefs[1]*x+ coefs[2]*x**2+coefs[3]*x**3+coefs[4]*x**4

d_coefs=poly.polyder(coefs)
d_coefs_roots=np.roots(d_coefs[ : : -1])

for m in range(0,3):
    print '%d'%(m)
    if 0 < d_coefs_roots[m] < 150:
        tnadir=d_coefs_roots[m]
        break

BGnadir=np.polyval(coefs[ : : -1], tnadir)
BGdrop=coefs[0]-BGnadir

tnadj=93.668
#tnadj=BGdrop*0.374672

AUC0totnadj=integrate.quad(ins_func,0,tnadj)
AUCtnadjto240=integrate.quad(ins_func,tnadj,240)

AOC0totnadj=tnadj*coefs[0]-AUC0totnadj[0]
AOCtnadjto240=(240-tnadj)*coefs[0]-AUCtnadjto240[0]


