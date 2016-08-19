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
ind_coefs=np.zeros((numRats-1,5))
ind_d_coefs=np.zeros((numRats-1,4))
ind_d_roots=np.zeros((numRats-1,3), dtype=np.complex_)
#ind_ffit=np.zeros((numRats-1, len(time)))
#ind_ins_func=np.empty((numRats-1,1))

mean_rat=np.average(rats,axis=0)


mean_coefs=poly.polyfit(time,mean_rat,4)
mean_fit=poly.polyval(time, mean_coefs)
mean_ins_func=lambda x: mean_coefs[0]+mean_coefs[1]*x+ mean_coefs[2]*x**2+mean_coefs[3]*x**3+mean_coefs[4]*x**4

for m in range (0, numRats-1):
    print '%d'%(m)
    ind_coefs[m]=poly.polyfit(time,rats[m],4)
    ind_d_coefs[m]=poly.polyder(ind_coefs[m])
    dummycoefs=ind_d_coefs[m]
    dummycoefs=dummycoefs[ : : -1]
    ind_d_roots[m]=np.roots(dummycoefs)
    #dummycoefs=np.zeros((4,1))
    #ind_ffit[m]=poly.polyval(time, ind_coefs[m])
    # ind_ins_func[m]=lambda x: ind_coefs[0]+ind_coefs[1]*x+ ind_coefs[2]*x**2+ind_coefs[3]*x**3+ind_coefs[4]*x**4





d_coefs=poly.polyder(mean_coefs)
d_coefs_roots=np.roots(d_coefs[ : : -1])

for m in range(0,3):
    print '%d'%(m)
    if 0 < d_coefs_roots[m] < 150:
        tnadir=np.real(d_coefs_roots[m])
        break

mean_BGnadir=np.polyval(mean_coefs[ : : -1], tnadir)
mean_BGdrop=mean_coefs[0]-mean_BGnadir

tnadj=mean_BGdrop*0.374672

mean_AUC0totnadj=integrate.quad(mean_ins_func,0,tnadj)
mean_AUCtnadjto240=integrate.quad(mean_ins_func,tnadj,240)

mean_AOC0totnadj=tnadj*mean_coefs[0]-mean_AUC0totnadj
mean_AOCtnadjto240=(240-tnadj)*mean_coefs[0]-mean_AUCtnadjto240
