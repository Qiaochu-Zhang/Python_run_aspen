# -*- coding: utf-8 -*-
"""
Created on Mon Jun 10 22:48:10 2019

@author: Y Zhang

3RCLHG_4variable.py
"""

#cd C:/Qiaochu/Research/Project/"UOP H2 shared file"/"UOP H2 shared file"


# alias python='winpty python'


import os
import win32com.client as win32
import pandas as pd
#import numpy as np
import time as time

aspen = win32.Dispatch('Apwn.Document')

file_name = 'CLHG_UOP_noPSA_P1P2_noUtCalc.bkp'

aspen.InitFromArchive2(os.path.abspath(file_name))

#block, stream, component name
solids=['FE2O3','FE3O4','FE0947O','FE'] #reducer bot molar flow
wequips=['2001Y01','2002Y01','1008Y01','8401P01','1202P01'] #electricity
gmfs=['S10-01','H2PROD','S10-10','S81-62','S81-62O'] #Mass flow
vfs=['S12-54','S12-16','S12-20','S12-02','CGC1'] #Volume flow
nfs=['S10-01']
sts=['S11-07','S11-02','S11-03'] #steam temp
gts=['S10-06','STGEN','S10-12','S12-20','S12-02','CGC1'] # stream temp
mfrs=['FE2O3','TIO2','AL2O3'] #Mass Frac
hrs=['S11-10','S11-20','S11-50']
cis=['S12-20','S12-52','S12-04','S12-06','CGC1']
cos=['S12-51','S12-53','S12-05','S12-07','S11-55']
cus=['1202E02','1202E03','1203E03','1203E04','1201E02'] #heat duty
###temp1=['re_bot_temp','ox_bot_temp','combustor_temp']

df_rb=pd.DataFrame(columns=solids) #reducer bottom iron oxide molar flow rate
df_ob=pd.DataFrame(columns=solids) #oxidizer bottom iron oxide molar flow rate
df_e=pd.DataFrame(columns=wequips)  #electricity
df_mf=pd.DataFrame(columns=gmfs+['S11-07','PureCO2']) #mass flow rate
df_vf=pd.DataFrame(columns=vfs) #volume flow rate
df_nf=pd.DataFrame(columns=nfs+['Fe2O3In']+['PureH2']+['ReducerCarbon']) #molar flow rate
df_err=pd.DataFrame(columns=['error']) #error number
df_t=pd.DataFrame(columns=sts+gts)
df_mfr=pd.DataFrame(columns=mfrs)
df_hr=pd.DataFrame(columns=hrs)
df_ci=pd.DataFrame(columns=cis)
df_co=pd.DataFrame(columns=cos)
######df_co=pd.DataFrame(columns=temp1)

#Sensitivity points
df_tc=pd.DataFrame(columns=['VarRange'])

#Air preheat temperature variation
m_lb=300
dm=50
#m=11
m=5
df_tc.at[1,'VarRange']='Air preheat temperature varies from '+str(m_lb)+'C to '+str(m_lb+dm*(m-1))+'C, increment '+str(dm)+'C'

#Reducer top temperature/combustor operating temperature variation
n_lb=1000
dn=5
#n=16
n=1
df_tc.at[2,'VarRange']='Reducer top/Combustor temperature varies from '+str(n_lb)+'C to '+str(n_lb+dn*(n-1))+'C, increment '+str(dn)+'C'

#Reducer bottom solid conversion variation
p_lb=0.4
dp=0.04
#p=11
p=4
df_tc.at[3,'VarRange']='Reducer bottom solid conversion varies from '+str(int(p_lb*100))+'% to '+str(int(100*(p_lb+dp*(p-1))))+'%, increment '+str(int(100*dp))+'%'

#Fe2O3 wt%
q_lb=0.2
dq=0.03
q=4
df_tc.at[4,'VarRange']='Fe2O3 wt% varies from '+str(int(q_lb*100))+'% to '+str(int(100*(q_lb+dq*(q-1))))+'%, increment '+str(int(100*dq))+'%'

t0=time.time()

##Case count for memory dump
case_count = 0

for l in range(q):
    #oxygen carrier composition setting
    fewt=q_lb+l*dq
    tiwt=0.21
    alwt=1.0-fewt-tiwt
    #？？？ how to findnode: customize>> variable explorer
    # findnode 放在变量前面是赋值，放在后面是取值。
    aspen.Tree.FindNode("\Data\Streams\S11-07\Input\FLOW\CISOLID\FE2O3").Value=fewt
    aspen.Tree.FindNode("\Data\Streams\S11-07\Input\FLOW\CISOLID\TIO2").Value=tiwt
    aspen.Tree.FindNode("\Data\Streams\S11-07\Input\FLOW\CISOLID\AL2O3").Value=alwt
    #Application.Tree.FindNode("\Data\Blocks\CBOIL1\Input\TEMP")
    for k in range(p):
        #reducer bottom solid conversion
        Inp_rbsc=p_lb+k*dp
        aspen.Tree.FindNode("\Data\Flowsheeting Options\Design-Spec\REDSCONV\Input\EXPR2").Value=str(Inp_rbsc)

        for j in range(n):
            #Reducer top T and combustor outlet T, in degree C
            Red_T=n_lb+j*dn
            aspen.Tree.FindNode("\Data\Streams\S11-07\Input\TEMP\CISOLID").Value=Red_T

            for i in range(m):
                print("i = ",i,";j = ",j,"k = ",k,"l = ",l,";second:",time.time()-t0)
                #Air preheat temperature, in degree C
                Air_T=m_lb+i*dm
                aspen.Tree.FindNode("\Data\Blocks\HRSG\Input\VALUE\S10-11").Value=Air_T

                #Model run
                aspen.Reinit()
                aspen.Engine.Run2()

                errall=''
                errormessage=[]
                for e in aspen.Tree.FindNode("\Data\Results Summary\Run-Status\Output\PER_ERROR").Elements:   
                    print (e.Value)
                    errormessage += e.value
                    if '=' in errormessage:
                        break

                errall=errall.join(errormessage)
                if 'error' in errall:
                    errall='error'
                else:
                    errall = 'OK'
# l in q; k in p; j in n; i in m;
#Error message
#mer=aspen.Tree.FindNode("\Data\Results Summary\Run-Status\Output\PER_ERROR").Value
#if mer=1:
#    print

#Data record

#df_out=pd.DataFrame(columns = solids)

# These variable have values only after running
# {} symbol can be used to substitute a variable
                for solid in solids:
                    df_rb.at[l*m*n*p+k*m*n+j*m+i+1,'{}'.format(solid)]=aspen.Tree.FindNode("\Data\Streams\S11-02\Output\MOLEFLOW\CISOLID\{}".format(solid)).Value
                    df_ob.at[l*m*n*p+k*m*n+j*m+i+1,'{}'.format(solid)]=aspen.Tree.FindNode("\Data\Streams\S11-03\Output\MOLEFLOW\CISOLID\{}".format(solid)).Value
# l in q; k in p; j in n; i in m; 
                for equip in wequips:
                    df_e.at[l*m*n*p+k*m*n+j*m+i+1,'{}'.format(equip)]=aspen.Tree.FindNode("\Data\Blocks\{}\Output\WNET".format(equip)).Value
            
                for mf in gmfs:
                    df_mf.at[l*m*n*p+k*m*n+j*m+i+1,'{}'.format(mf)]=aspen.Tree.FindNode("\Data\Streams\{}\Output\MASSFLMX\MIXED".format(mf)).Value

                df_mf.at[l*m*n*p+k*m*n+j*m+i+1,'S11-07']=aspen.Tree.FindNode("\Data\Streams\S11-07\Output\MASSFLMX\CISOLID").Value
                df_mf.at[l*m*n*p+k*m*n+j*m+i+1,'PureCO2']=aspen.Tree.FindNode("\Data\Streams\S20-10\Output\MASSFLOW\MIXED\CO2").Value
    
                for vf in vfs:
                    df_vf.at[l*m*n*p+k*m*n+j*m+i+1,'{}'.format(vf)]=aspen.Tree.FindNode("\Data\Streams\{}\Output\VOLFLMX\MIXED".format(vf)).Value
    
                for st in sts:
                    df_t.at[l*m*n*p+k*m*n+j*m+i+1,'{}'.format(st)]=aspen.Tree.FindNode("\Data\Streams\{}\Output\TEMP_OUT\CISOLID".format(st)).Value
    
                for gt in gts:
                    df_t.at[l*m*n*p+k*m*n+j*m+i+1,'{}'.format(gt)]=aspen.Tree.FindNode("\Data\Streams\{}\Output\TEMP_OUT\MIXED".format(gt)).Value
        
                for mfr in mfrs:
                    df_mfr.at[l*m*n*p+k*m*n+j*m+i+1,'{}'.format(mfr)]=aspen.Tree.FindNode("\Data\Streams\S11-07\Output\MASSFRAC\CISOLID\{}".format(mfr)).Value
        
                for hr in hrs:
                    df_hr.at[l*m*n*p+k*m*n+j*m+i+1,'{}'.format(hr)]=-aspen.Tree.FindNode("\Data\Blocks\HRSG\Output\QCALC\{}".format(hr)).Value
        
                    # HMX_FLOW == ENTHALPY FLOW GCAL/hr
                for ci in cis:
                    df_ci.at[l*m*n*p+k*m*n+j*m+i+1,'{}'.format(ci)]=aspen.Tree.FindNode("\Data\Streams\{}\Output\HMX_FLOW\$TOTAL".format(ci)).Value
        
                for co in cos:
                    df_co.at[l*m*n*p+k*m*n+j*m+i+1,'{}'.format(co)]=aspen.Tree.FindNode("\Data\Streams\{}\Output\HMX_FLOW\$TOTAL".format(co)).Value
        
                df_nf.at[l*m*n*p+k*m*n+j*m+i+1,'S10-01']=aspen.Tree.FindNode("\Data\Streams\S10-01\Output\MOLEFLMX\MIXED").Value
                df_nf.at[l*m*n*p+k*m*n+j*m+i+1,'Fe2O3In']=aspen.Tree.FindNode("\Data\Streams\S11-07\Output\MOLEFLOW\CISOLID\FE2O3").Value
                df_nf.at[l*m*n*p+k*m*n+j*m+i+1,'PureH2']=aspen.Tree.FindNode("\Data\Streams\H2PROD\Output\MOLEFLOW\MIXED\H2").Value
                df_nf.at[l*m*n*p+k*m*n+j*m+i+1,'ReducerCarbon']=aspen.Tree.FindNode("\Data\Streams\S11-02\Output\MOLEFLOW\CISOLID\CARBON").Value
                df_err.at[l*m*n*p+k*m*n+j*m+i+1,'error']=errall
                
                case_count += 1
                print ('Case ', case_count, 'Complete')
                
                ##Restart 
                if case_count %100 == 0:
                    print('RESTART ASPEN')
                    aspen.Close()
                    aspen = win32.Dispatch('Apwn.Document')
                    file_name = 'CLHG_UOP_noPSA_P1P2_noUtCalc.bkp'
                    aspen.InitFromArchive2(os.path.abspath(file_name))
                    #value input for current loop
                    aspen.Tree.FindNode("\Data\Streams\S11-07\Input\FLOW\CISOLID\FE2O3").Value=fewt
                    aspen.Tree.FindNode("\Data\Streams\S11-07\Input\FLOW\CISOLID\TIO2").Value=tiwt
                    aspen.Tree.FindNode("\Data\Streams\S11-07\Input\FLOW\CISOLID\AL2O3").Value=alwt
                    aspen.Tree.FindNode("\Data\Flowsheeting Options\Design-Spec\REDSCONV\Input\EXPR2").Value=str(Inp_rbsc)
                    aspen.Tree.FindNode("\Data\Streams\S11-07\Input\TEMP\CISOLID").Value=Red_T
                    #aspen.Tree.FindNode("\Data\Streams\S11-07\Input\TEMP\CISOLID").Value=Red_T
                    #aspen.Reinit()
                    #aspen.Engine.Run2()

t=time.time()-t0
print("Running time:",t,"seconds")
    
df_prop=pd.DataFrame()
df_scale=pd.DataFrame()

df_qi=df_ci
df_qi.columns=cus
df_qo=df_co
df_qo.columns=cus
df_Q=df_qi-df_qo

#unscalable variables
df_prop['Fe2O3 wt']=df_mfr['FE2O3']
df_prop['RBSC']=1-(df_rb['FE0947O']+df_rb['FE3O4']*4).divide(df_nf['Fe2O3In']*3)
df_prop['OBSC']=1-(df_ob['FE0947O']+df_ob['FE3O4']*4).divide(df_nf['Fe2O3In']*3)
df_prop['H2NGR']=df_nf['PureH2'].divide(df_nf['S10-01'])
df_prop['NGInT(C)']=df_t['S10-06']
df_prop['STInT(C)']=df_t['STGEN']
df_prop['AirInT(C)']=df_t['S10-12']
df_prop['RedBT(C)']=df_t['S11-02']
df_prop['OxBT(C)']=df_t['S11-03']
df_prop['CombT(C)']=df_t['S11-07']
df_prop['error']=df_err['error']

#scalable variables
df_scale['NGMole(kmol/hr)']=df_nf['S10-01']
df_scale['SolidFlow(kg/hr)']=df_mf['S11-07']
df_scale['H2Mass(kg/hr)']=df_mf['H2PROD']
df_scale['CLHRSG(Gcal/hr)']=df_hr['S11-10']+df_hr['S11-20']+df_hr['S11-50']
df_scale['RC1(Gcal/hr)']=df_Q['1202E02']
df_scale['RC2(Gcal/hr)']=df_Q['1202E03']
df_scale['OC1(Gcal/hr)']=df_Q['1203E03']
df_scale['OC2(Gcal/hr)']=df_Q['1203E04']
df_scale['CC(Gcal/hr)']=df_Q['1201E02']
df_scale['AirMass(kg/hr)']=df_mf['S10-10']
df_scale['H2vf(m3/hr)']=df_vf['S12-16']
df_scale['CO2vf(m3/hr)']=df_vf['S12-54']
df_scale['RGvf(m3/hr)']=df_vf['S12-20']
df_scale['OGvf(m3/hr)']=df_vf['S12-02']
df_scale['CGvf(m3/hr)']=df_vf['CGC1']
df_scale['Electricity(kW)']=df_e.sum(axis=1)
df_scale['NGMass(kg/hr)']=df_mf['S10-01']
df_scale['SteamOut(kg/hr)']=df_mf['S81-62O']-df_mf['S81-62']
df_scale['CO2Mass(kg/hr)']=df_mf['PureCO2']

#scale to 4632 kg/hr H2 production
df_all=pd.DataFrame(columns=list(df_prop)+list(df_scale))

for pr in list(df_prop):
    df_all['{}'.format(pr)]=df_prop['{}'.format(pr)]

for sc in list(df_scale):
    df_all['{}'.format(sc)]=df_scale['{}'.format(sc)].divide(df_mf['H2PROD'])*4632.96

aspen.Close()

with pd.ExcelWriter('Fe32_SC_AirT_CT.xlsx') as writer:
    df_tc.to_excel(writer, sheet_name='Info')
    df_all.to_excel(writer, sheet_name='Summary')
    df_e.to_excel(writer, sheet_name='Electricity')
    df_mf.to_excel(writer, sheet_name='Mass Flow')
    df_mfr.to_excel(writer, sheet_name='Mass Frac')
    df_nf.to_excel(writer, sheet_name='Mole Flow')
    df_rb.to_excel(writer, sheet_name='Reducer Bottom Mole Flow')
    df_ob.to_excel(writer, sheet_name='Oxidizer Bottom Mole Flow')
    df_t.to_excel(writer, sheet_name='Stream temp')
    df_vf.to_excel(writer, sheet_name='Volume flow')
    df_Q.to_excel(writer, sheet_name='Heat Duty')