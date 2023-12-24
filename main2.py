import numpy as np
import os
import pandas as pd
import matplotlib
import matplotlib.pyplot as plt
import datetime

from multiprocessing.pool import ThreadPool
pool = ThreadPool(8) #not sure if this is making a difference...
now = datetime.datetime.now()
print('start code time %s'%now)
#import XLS files
path = os.getcwd()
files = os.listdir(path)
files_xls = [f for f in files if f[-3:] == 'xls']
numbr=len(files_xls)

#initialize data frames needed
dd = []
de = []
df = []
dg = []
dh = []
di = []
names = []
ones = pd.DataFrame(range(1,43,1))


#loop through list
for f in files_xls: #read through excel files in the folder
    fname = f.split('.')[0]
    names.append(fname)
     # sort the oscillation strain (column D)
    dfd = pd.read_excel(f, 'Time Sweep - 1',usecols='D',skiprows=8, nrows=42) 
    dfd.columns = [f.split('.')[0]]
    dd.append(dfd) #list here
    # sort the oscillation stress (column E)
    dfe = pd.read_excel(f, 'Time Sweep - 1',usecols='E',skiprows=8, nrows=42) 
    dfe.columns = [f.split('.')[0]]
    de.append(dfe) #list here
    # sort the tan(delta) (column F)
    dff = pd.read_excel(f, 'Time Sweep - 1',usecols='F',skiprows=8, nrows=42) 
    dff.columns = [f.split('.')[0]]
    df.append(dff) #list here
    #sort the storage modulus (column G)
    dfg = pd.read_excel(f, 'Time Sweep - 1',usecols='G',skiprows=8, nrows=42) 
    dfg.columns = [f.split('.')[0]]
    dg.append(dfg) #list here
    #sort the loss modulus (column H)
    dfh = pd.read_excel(f, 'Time Sweep - 1',usecols='H',skiprows=8, nrows=42) 
    dfh.columns = [f.split('.')[0]]
    dh.append(dfh) #list here
    #sort the position (mm) (columnI)
    dfi = pd.read_excel(f, 'Time Sweep - 1',usecols='I',skiprows=8, nrows=42) 
    dfi.columns = [f.split('.')[0]]
    di.append(dfi) #list here
    

   
## manupulate data from dataframe
# oscillation strain
ddata = pd.concat(dd, axis=1)
davg = ddata.mean()
dstd = ddata.std()
#calculating the max and minimum values
dmax = ddata.max()
dmaxid = dmax.idxmax()
daxismax = dmax[dmaxid]
dmin = ddata.min()
dminid = dmin.idxmin()
daxismin = dmin[dminid]

# Oscillation stress
edata = pd.concat(de, axis=1)
eavg = edata.mean()
estd = edata.std()
#calculating the max and minimum values
emax = edata.max()
emaxid = emax.idxmax()
eaxismax = emax[emaxid]
emin = edata.min()
eminid = emin.idxmin()
eaxismin = emin[eminid]

#tan(delta)
fdata = pd.concat(df, axis=1)
favg = fdata.mean()
fstd = fdata.std()
#calculating the max and minimum values
fmax = fdata.max()
fmaxid = fmax.idxmax()
faxismax = fmax[fmaxid]
fmin = fdata.min()
fminid = fmin.idxmin()
faxismin = fmin[fminid]

#Storage Modulus
gdata = pd.concat(dg, axis=1)
gavg = gdata.mean()
gstd = gdata.std()
#calculating the max and minimum values
gmax = gdata.max()
gmaxid = gmax.idxmax()
gaxismax = gmax[gmaxid]
gmin = gdata.min()
gminid = gmin.idxmin()
gaxismin = gmin[gminid]

#Loss Modulus
hdata = pd.concat(dh, axis=1)
havg = hdata.mean()
hstd = hdata.std()
#calculating the max and minimum values
hmax = hdata.max()
hmaxid = hmax.idxmax()
haxismax = hmax[hmaxid]
hmin = hdata.min()
hminid = hmin.idxmin()
haxismin = hmin[hminid]
                
#position
idata = pd.concat(di, axis=1)
iavg = idata.mean()
istd = idata.std()
#calculating the max and minimum values
imax = idata.max()
imaxid = imax.idxmax()
iaxismax = imax[imaxid]
imin = idata.min()
iminid = imin.idxmin()
iaxismin = imin[iminid]
    
# ## plotting each sample in it's own figure - panel for each column of data
# for i in names:
#     # plt.figure()
#     fig, axs = plt.subplots(3,2, sharex = True, figsize = (12,12))
#     #make plot look nice
#     fig.suptitle(i)
#     # Plot oscillation strain
#     axs[0,0].plot(ones,ddata[i])
#     axs[0,0].set_title('oscillation strain', fontsize = 16)
#     axs[0,0].set_ylabel('%', fontsize = 12)
#     axs[0,0].set_ylim([daxismin, daxismax])
#     #plot oscillation stress
#     axs[0,1].plot(ones,edata[i])
#     axs[0,1].set_title('oscillation stress', fontsize = 16)
#     axs[0,1].set_ylabel('kpa', fontsize = 12)
#     axs[0,1].set_ylim([eaxismin, eaxismax])
#     #plot tan(delta)
#     axs[1,0].plot(ones,fdata[i])
#     axs[1,0].set_title('tan(delta)', fontsize = 16)
#     axs[1,0].set_ylabel(' ', fontsize = 12)
#     axs[1,0].set_ylim([faxismin, faxismax])
#     # Plot storage modulus
#     axs[1,1].plot(ones,gdata[i])
#     axs[1,1].set_title('storage modulus', fontsize = 16)
#     axs[1,1].set_ylabel('kpa', fontsize = 12)
#     axs[1,1].set_ylim([gaxismin, gaxismax])
#     #plot loss modulus
#     axs[2,0].plot(ones,hdata[i])
#     axs[2,0].set_title('loss modulus', fontsize = 16)
#     axs[2,0].set_ylabel('kpa', fontsize = 12)
#     axs[2,0].set_xlabel('acquisition', fontsize = 12)
#     axs[2,0].set_ylim([haxismin, haxismax])
#     #plot position
#     axs[2,1].plot(ones,idata[i])
#     axs[2,1].set_title('z-position', fontsize = 16)
#     axs[2,1].set_ylabel('position (mm)', fontsize = 12)
#     axs[2,1].set_xlabel('acquisition', fontsize = 12)
#     axs[2,1].set_ylim([iaxismin, iaxismax])
#     plt.tight_layout()
#     plt.savefig('plot_%s.png'%i, dpi=600)  
#     plt.close() #comment this out if you want to see the plots. Closing plots helps code run faster
    

    
## Save data that was manipulated in previous operation to a single excel file with multiple folders
with pd.ExcelWriter('day-data.xlsx', engine='openpyxl') as writer:
    #save oscillation strain data to excel sheet
    ddata.to_excel(writer, sheet_name='oscillation strain')
    davg.to_excel(writer, sheet_name='oscillation strain avg')    
    dstd.to_excel(writer, sheet_name='oscillation strain std')
    #save oscillation stress data to excel sheet
    edata.to_excel(writer, sheet_name='oscillation stress')
    eavg.to_excel(writer, sheet_name='oscillation stress avg')
    estd.to_excel(writer, sheet_name='oscillation stress std')
    #save tan(delta) data to excel sheet
    fdata.to_excel(writer, sheet_name='tan(delta)')
    favg.to_excel(writer, sheet_name='tan(delta) avg')    
    fstd.to_excel(writer, sheet_name='tan(delta) std')
    #save storage modulus data to excel sheet
    gdata.to_excel(writer, sheet_name='storage modulus')
    gavg.to_excel(writer, sheet_name='storage modulus avg')    
    gstd.to_excel(writer, sheet_name='storage modulus std')
    #save loss modulus data to excel sheet
    hdata.to_excel(writer, sheet_name='loss modulus')
    havg.to_excel(writer, sheet_name='loss modulus avg')    
    hstd.to_excel(writer, sheet_name='loss modulus std')
now = datetime.datetime.now()
print('code run without errors time: %s'%now)