import openpyxl
import numpy as np
import numpy.polynomial.polynomial as poly
import scipy
from scipy import integrate



wb = openpyxl.load_workbook('example.xlsx')
s1=wb.get_sheet_by_name('Sheet1')

rats_RawData=np.array([[cell.value for cell in col] for col in s1['K1':'W7']])

#np.average()
numRats=len(rats_RawData)


time=rats_RawData[0]
rats=rats_RawData[1:numRats]

mean_rat=np.average(rats,axis=0)

# find the mean of the time series data

#for m in range 

#for m in range(0, s1.max_column):
#rat1=[458,	380,	317,	316,	251,	262,	241,	239,	281,	296,	334,	342, 298]

mean_coefs=poly.polyfit(time,mean_rat,4)
mean_fit=poly.polyval(time, mean_coefs)
ins_func=lambda x: mean_coefs[0]+mean_coefs[1]*x+ mean_coefs[2]*x**2+mean_coefs[3]*x**3+mean_coefs[4]*x**4

d_coefs=poly.polyder(mean_coefs)
d_coefs_roots=np.roots(d_coefs[ : : -1])

for m in range(0,3):
    print '%d'%(m)
    if 0 < d_coefs_roots[m] < 150:
        tnadir=np.real(d_coefs_roots[m])
        break

BGnadir=np.polyval(mean_coefs[ : : -1], tnadir)
BGdrop=mean_coefs[0]-BGnadir

tnadj=BGdrop*0.374672

AUC0totnadj=integrate.quad(ins_func,0,tnadj)
AUCtnadjto240=integrate.quad(ins_func,tnadj,240)

AOC0totnadj=tnadj*mean_coefs[0]-AUC0totnadj[0]
AOCtnadjto240=(240-tnadj)*mean_coefs[0]-AUCtnadjto240[0]
