import openpyxl
import numpy as np
import numpy.polynomial.polynomial as poly
import scipy
from scipy import integrate
import string


wb1 = openpyxl.load_workbook('example.xlsx')
s1=wb1.get_sheet_by_name('Sheet1')
s2=wb1.get_sheet_by_name('Sheet2')
s3=wb1.get_sheet_by_name('Sheet3')
s4=wb1.get_sheet_by_name('Sheet4')
analog='T-0737'

rats_RawData=np.array([[cell.value for cell in col] for col in s1['D1':'P5']])

#np.average()
numRats=len(rats_RawData)


time=rats_RawData[0]
rats=rats_RawData[1:numRats]
ind_coefs=np.zeros((numRats-1,5))
ind_d_coefs=np.zeros((numRats-1,4))
ind_d_roots=np.zeros((numRats-1,3), dtype=np.complex_)
ind_AUC0totnadj=np.zeros((numRats-1,2))
ind_AUCtnadjto240=np.zeros((numRats-1,2))
ind_AOC0totnadj=np.zeros((numRats-1,1))
ind_AOCtnadjto240=np.zeros((numRats-1,1))
ind_tnadir=np.zeros((numRats-1,1))
ind_BGN=np.zeros((numRats-1,1))
ind_IBG=np.zeros((numRats-1,1))
#ind_ffit=np.zeros((numRats-1, len(time)))
#ind_ins_func=np.empty((numRats-1,1))

mean_rat=np.average(rats,axis=0)


mean_coefs=poly.polyfit(time,mean_rat,4)
mean_IBG=mean_coefs[0]
mean_fit=poly.polyval(time, mean_coefs)
mean_ins_func=lambda x: mean_coefs[0]+mean_coefs[1]*x+ mean_coefs[2]*x**2+mean_coefs[3]*x**3+mean_coefs[4]*x**4

#finding the mean nadir
d_coefs=poly.polyder(mean_coefs)
d_coefs_roots=np.roots(d_coefs[ : : -1])

for m in range(0,3):
    #print '%d'%(m)
    if 0 < d_coefs_roots[m] < 132: #upper limit corresponds to tn of average 10ug KP if BG drop is 350
        tnadir=np.real(d_coefs_roots[m])
        break
mean_BGnadir=np.polyval(mean_coefs[ : : -1], tnadir)
mean_BGdrop=mean_coefs[0]-mean_BGnadir

tnadj=mean_BGdrop*0.374672

#finding tnadir for each individual rat
for m in range (0, numRats-1):
    #print '%d'%(m)
    ind_coefs[m]=poly.polyfit(time,rats[m],4)
    ind_IBG[m]=ind_coefs[m,0]
    ind_d_coefs[m]=poly.polyder(ind_coefs[m])
    dummycoefs=ind_d_coefs[m] #dummycoefs gets re-written every time...there may be a better way to do this
    dummycoefs=dummycoefs[ : : -1]
    ind_d_roots[m]=np.roots(dummycoefs)
    

#finding the individual nadirs will go here if need be
for m in range(0,numRats-1):
    for n in range(0,3):        
        #print '%d'%(m)
        if 0 < ind_d_roots[m,n] < 132: #upper limit corresponds to tn of average 10ug KP if BG drop is 350
            ind_tnadir[m]=np.real(ind_d_roots[m,n])
            ind_BGN[m]=np.polyval(ind_coefs[m][::-1],ind_tnadir[m])
            break




mean_AUC0totnadj=integrate.quad(mean_ins_func,0,tnadj)
mean_AUCtnadjto240=integrate.quad(mean_ins_func,tnadj,240)

mean_AOC0totnadj=tnadj*mean_coefs[0]-mean_AUC0totnadj
mean_AOCtnadjto240=(240-tnadj)*mean_coefs[0]-mean_AUCtnadjto240

for m in range (0, numRats-1):
    ind_ins_func=lambda x: ind_coefs[m,0]+ind_coefs[m,1]*x+ ind_coefs[m,2]*x**2+ind_coefs[m,3]*x**3+ind_coefs[m,4]*x**4
    ind_AUC0totnadj[m]=integrate.quad(ind_ins_func,0,tnadj)
    ind_AUCtnadjto240[m]=integrate.quad(ind_ins_func,tnadj,240)
    ind_AOC0totnadj[m]=tnadj*ind_coefs[m,0]-ind_AUC0totnadj[m,0]
    ind_AOCtnadjto240[m]=(240-tnadj)*ind_coefs[m,0]-ind_AUCtnadjto240[m,0]
    
ind_data={'IBG':ind_IBG,'BGN':ind_BGN, 'Tnadir':ind_tnadir,'AOC0totnadj':ind_AOC0totnadj,'AOCtnadjto240':ind_AOCtnadjto240}

#wb=openpyxl.Workbook()
#ws=wb.active
#ws.title='%s Analysis' %(analog) 
#mean_data=np.transpose(np.hstack((mean_IBG, mean_BGnadir,tnadir,mean_AOC0totnadj[0],mean_AOCtnadjto240[0])))
totIBG=np.append(ind_IBG,mean_IBG)
totBGN=np.append(ind_BGN,mean_BGnadir)
tottnadir=np.append(ind_tnadir,tnadir)
totAOC1=np.append(ind_AOC0totnadj,mean_AOC0totnadj[0])
totAOC2=np.append(ind_AOCtnadjto240,mean_AOCtnadjto240[0])
mean_data=np.array((mean_IBG, mean_BGnadir,tnadir,mean_AOC0totnadj[0],mean_AOCtnadjto240[0]))
ind_data_table=np.transpose(np.hstack((ind_IBG,ind_BGN, ind_tnadir,ind_AOC0totnadj,ind_AOCtnadjto240)))

tot_data_table=np.vstack((totIBG,totBGN, tottnadir,totAOC1,totAOC2))
tot_tableshape=np.shape(tot_data_table)
alph=list(string.ascii_uppercase)

for i in range(tot_tableshape[0]):
    for j in range(tot_tableshape[1]):
        s4[alph[i]+str(j+1)] = tot_data_table[i, j]

wb1.save('example.xlsx')
