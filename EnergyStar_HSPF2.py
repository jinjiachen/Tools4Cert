#/bin/python #-*-coding:utf-8-*-

'''
Author: Michael Jin
Date: 2023-12

'''

import numpy as np
import matplotlib.pyplot as plt


###全局变量
HLH={
        'I':493,
        'II':857,
        'III':1247,
        'IV':1701,
        'V':2202,
        'VI':1842,
        }
Tod={
        'I':37,
        'II':27,
        'III':17,
        'IV':5,
        'V':-10,
        'VI':30,
        }
C={
        'I':1.10,
        'II':1.06,
        'III':1.30,
        'IV':1.15,
        'V':1.16,
        'VI':1.11,
        }
Cvs={
        'I':1.03,
        'II':0.99,
        'III':1.21,
        'IV':1.07,
        'V':1.08,
        'VI':1.03,
        }
Tzl={
        'I':58,
        'II':57,
        'III':56,
        'IV':55,
        'V':55,
        'VI':57,
        }
nj_N={
        'I':{'62':0,
             '57':0.239,
             '52':0.194,
             '47':0.129,
             '42':0.081,
             '37':0.041,
             '32':0.019,
             '27':0.005,
             '22':0.001,
             '17':0,
             '12':0,
             '7':0,
             '2':0,
             '-3':0,
             '-8':0,
             '-13':0,
             '-18':0,
             '-23':0,
             },
        'II':{
             '62':0,
             '57':0,
             '52':0.163,
             '47':0.143,
             '42':0.112,
             '37':0.088,
             '32':0.056,
             '27':0.024,
             '22':0.008,
             '17':0.002,
             '12':0,
             '7':0,
             '2':0,
             '-3':0,
             '-8':0,
             '-13':0,
             '-18':0,
             '-23':0,
            },
        'III':{
             '62':0,
             '57':0,
             '52':0.138,
             '47':0.137,
             '42':0.135,
             '37':0.118,
             '32':0.092,
             '27':0.047,
             '22':0.021,
             '17':0.009,
             '12':0.005,
             '7':0.002,
             '2':0.001,
             '-3':0,
             '-8':0,
             '-13':0,
             '-18':0,
             '-23':0,
            },
        'IV':{
             '62':0,
             '57':0,
             '52':0.103,
             '47':0.093,
             '42':0.100,
             '37':0.109,
             '32':0.126,
             '27':0.087,
             '22':0.055,
             '17':0.036,
             '12':0.026,
             '7':0.013,
             '2':0.006,
             '-3':0.002,
             '-8':0.001,
             '-13':0,
             '-18':0,
             '-23':0,
            },
        'V':{
             '62':0,
             '57':0,
             '52':0.086,
             '47':0.076,
             '42':0.078,
             '37':0.087,
             '32':0.102,
             '27':0.094,
             '22':0.074,
             '17':0.055,
             '12':0.047,
             '7':0.038,
             '2':0.029,
             '-3':0.018,
             '-8':0.010,
             '-13':0.005,
             '-18':0.002,
             '-23':0.001,
            },
        'VI':{
             '62':0,
             '57':0,
             '52':0.215,
             '47':0.204,
             '42':0.141,
             '37':0.076,
             '32':0.034,
             '27':0.008,
             '22':0.003,
             '17':0,
             '12':0,
             '7':0,
             '2':0,
             '-3':0,
             '-8':0,
             '-13':0,
             '-18':0,
             '-23':0,
            },
        }
Tj_bin=range(62,-28,-5)
Tj_bin=list(Tj_bin)
Tj_bin.append(35)#单独增加35的计算

###计算building load, Equation 4.2-2
def build_load(zone,Q):
    '''
    zone(str):对应的区域，如I,IV等
    Q(float):如冷热型，则为A或A2时的制冷量（Qc_95),如单热，则用Qh_47替代
    '''
    res={}
    for Tj in Tj_bin:#遍历对应的Tj
        BL=(Tzl[zone]-Tj)/(Tzl[zone]-5)*C[zone]*Q
        res[str(Tj)]=BL
    return res


###通过线性插值计算一组Tj对应的值
def Values_Tj(x,y):
    '''
    x(list):自变量，此处对应两个温度值
    y(list):应变量，两个温度对应的制冷量或功率
    '''
    y[0]=float(y[0])
    y[1]=float(y[1])
    if len(x)==2 and len(y)==2:
        res={}
        for Tj in Tj_bin:#遍历对应的Tj
            Q=y[0]+(y[1]-y[0])/(x[1]-x[0])*(Tj-x[0])#计算并记录结果
            res[str(Tj)]=Q
        return res
    else:
        return 'error'


###线性函数，主要服务于Equation 4.2.4-7&8
def liner(QE35_k1,QE47_k1,QE62_k1,QE35_kv,QE35_k2):
    '''
    QE35_k1(float):Equation 4.2.4-1&-2算的Qh,Eh
    QE47_k1(float):H1_1工况下的Qh,Eh
    QE62_k1(float):H0_1工况下的Qh,Eh
    QE35_kv(float):H2_v工况下的Qh,Eh
    QE35_k2(float):H2_2工况下的Qh,Eh或3.6.4(c)估算
    '''
    #计算N,M
    N=(QE35_kv-QE35_k1)/(QE35_k2-QE35_k1)
    M=(QE62_k1-QE47_k1)/(62-47)*(1-N)

    #计算k=v时对应制冷量Q或功率E
    res={}
    for Tj in Tj_bin:#遍历对应的Tj
        QE_kv=QE35_kv+M*(Tj-35)
        res[str(Tj)]=QE_kv
    return res

###计算delta, Equation 4.2.3-3
def cal_delta(Toff,Ton):
    '''
    Toff(float):低于此温度，停机
    Ton(float):高于此温度，开机
    '''
    res={}
    for Tj in Tj_bin:#遍历对应的Tj
        if Tj<=Toff:
            delta=0
        elif Tj>Toff and Tj <=Ton:
            delta=0.5
        elif Tj>Ton:
            delta=1
        res[str(Tj)]=delta
    return res


###根据4.2.3.4章节计算delta'
def cal_delta_(Qh_k2,Eh_k2,Toff,Ton):
    '''
    Qh_k2(dict):根据4.2.4(c,d)计算k=2时所有Tj的Qh
    Eh_k2(dict):根据4.2.4(c,d)计算k=2时所有Tj的Eh
    Toff(float):低于此温度，停机
    Ton(float):高于此温度，开机
    '''
    res={}
    for Tj in Tj_bin:#遍历对应的Tj
        const=Qh_k2[str(Tj)]/(3.413*Eh_k2[str(Tj)])
        if Tj<=Toff or const<1:
            delta=0
        elif Tj>Toff and Tj<=Ton and const>=1:
            delta=0.5
        elif Tj>Ton and const>=1:
            delta=1
        res[str(Tj)]=delta
    return res

###根据3.6.4(c)估算Qh35_k2
def estimate_Qh35_k2(Qh17_k2,Qhcalc47_k2):
    '''
    Qh17_k2(str):H3_2工况下的制冷量
    Qhcalc17_k2(str):H1_2或H1_N下的制冷量
    '''
    Qh17_k2=float(Qh17_k2)
    Qhcalc47_k2=float(Qhcalc47_k2)
    Qh35_k2=0.90*(Qh17_k2+0.6*(Qhcalc47_k2-Qh17_k2))
    return Qh35_k2


###根据3.6.4(c)估算Eh35_k2
def estimate_Eh35_k2(Eh17_k2,Ehcalc47_k2):
    '''
    Eh17_k2(str):H3_2工况下的功率
    Ehcalc17_k2(str):H1_2或H1_N下的功率
    '''
    Eh17_k2=float(Eh17_k2)
    Ehcalc47_k2=float(Ehcalc47_k2)
    Eh35_k2=0.985*(Eh17_k2+0.6*(Ehcalc47_k2-Eh17_k2))
    return Eh35_k2


###根据4.2.4计算Qh,Eh
def cal_QE(QE35_kv,QE47_k1,QE62_k1,QE_kv):
    '''
    QE35_kv(float):k=v时的制冷量,功率
    QE47_k1(float):k=1,47度时的制冷量,功率
    QE62_k1(float):k=1,62度时的制冷量,功率
    QE_kv(list):Equation 4.2.4-7计算的制冷量,功率
    '''
    #Equation 4.2.4-5&6
    res={}
    for Tj in Tj_bin:#遍历对应的Tj
        if Tj>=47:
            QE_k1=QE47_k1+(QE62_k1-QE47_k1)*(Tj-47)/(62-47)
        elif Tj>=35 and Tj<47:
            QE_k1=QE35_kv+(QE47_k1-QE35_kv)*(Tj-35)/(47-35)
        elif Tj<35:
            QE_k1=QE_kv[str(Tj)]
        res[str(Tj)]=QE_k1
    return res


###根据4.2.4(c,d)计算k=2时的Qh，Eh
def cal_QE_k2(QE17_k2,QE47hcalc_k2,QE47_kN,QE35_k2,H4_k2='NO',QE5_k2=''):
    '''
    QE17_k2(float):H3_2工况下的制热量和功率
    QE47hcalc_k2(float):H1_2或H1_N下的制热量或功率
    QE47_kN(float):H1_N下的制热量或功率
    QE35_k2(float):根据3.6.4(c)估算的制热量或功率
    H4_k2(str):H4_2工况是否有做
    QE5_k2(float):H4_2工况下的制热量和功率
    '''
    res={}
    for Tj in Tj_bin:#遍历对应的Tj
        if Tj>=45:
            QE_k2=QE17_k2+(QE47hcalc_k2-QE17_k2)*(Tj-17)/(47-17)*(QE47_kN/QE47hcalc_k2)
        elif Tj>=17 and Tj<45:
            QE_k2=QE17_k2+(QE35_k2-QE17_k2)*(Tj-17)/(35-17)
        elif Tj<17:
            if H4_k2=='NO':
                QE_k2=QE17_k2+(QE47hcalc_k2-QE17_k2)*(Tj-17)/(47-17)
            elif H4_k2=='YES':
                if Tj>=5:
                    QE_k2=QE5_k2+(QE17_k2-QE5_k2)*(Tj-5)/(17-5)
                elif Tj<5:
                    QE_k2=QE5_k2-(QE47hcalc_k2-QE17_k2)*(5-Tj)/(47-17)
        res[str(Tj)]=QE_k2
    return res


###字典转数组
def dict_array(mydict):
    '''
    mydict(dict):想要转换的字典
    '''
    if isinstance(mydict,dict)==True:
        myarray=np.array(list(mydict.values()))
    else:
        print(f'{mydict}不是字典')
    return myarray

###计算HSPF2
def HSPF2():
    Cd=0.25
    zone='IV'
#    Q=input('请输入A或A2工况下的制冷量')
#    Toff=input('请输入Toff')
#    Ton=input('请输入Ton')
#    Qh47_k1=input('请输入H1_1下的制热量')
#    Eh47_k1=input('请输入H1_1下的功率')
#    Qh62_k1=input('请输入H0_1下的制热量')
#    Eh62_k1=input('请输入H0_1下的功率')
#    Qh47_kN=input('请输入H1_N下的制热量')
#    Eh47_kN=input('请输入H1_N下的功率')
#    Qh17_k2=input('请输入H3_2下的制热量')
#    Eh17_k2=input('请输入H3_2下的功率')
#    Qh35_kv=input('请输入H2_v下的制热量')
#    Eh35_kv=input('请输入H2_v下的功率')
    #以下测试用
    Q='11315.2'
    Toff='-25'
    Ton='-25'
    Qh47_k1=3285
    Eh47_k1=184.6
    Qh62_k1=4714
    Eh62_k1=177.9
    Qh47_kN=14344
    Eh47_kN=1108
    Qh17_k2=8296
    Eh17_k2=865.3
    Qh35_kv=6088
    Eh35_kv=414.9
    BL=build_load(zone,float(Q))#计算房间负荷
#    Qh_k1
#    Q_kv=liner()
#    Qh_k1=cal_Q(Q35_k1,Q47_k1,Q62_k1,Q_kv)#用Equation 4.2.4-5计算Qh_k=1
    cigma_eh_N={}
    cigma_RH_N={}
    Qh_k1=Values_Tj([47,62],[Qh47_k1,Qh62_k1])#Equation 4.2.4-1
    Eh_k1=Values_Tj([47,62],[Eh47_k1,Eh62_k1])#Equation 4.2.4-2
    Qhcalc47_k2=Qh47_kN
    Ehcalc47_k2=Eh47_kN
    Qh35_k2=estimate_Qh35_k2(Qh17_k2,Qhcalc47_k2)#估算Qh35_k2
    Eh35_k2=estimate_Eh35_k2(Eh17_k2,Ehcalc47_k2)#估算Qh35_k2
    Qh_kv=liner(Qh_k1['35'],Qh_k1['47'],Qh_k1['62'],Qh35_kv,Qh35_k2)#Equation 4.2.4-7
    Eh_kv=liner(Eh_k1['35'],Eh_k1['47'],Eh_k1['62'],Eh35_kv,Eh35_k2)#Equation 4.2.4-8
    Qh_k1_4245=cal_QE(Qh35_kv,Qh47_k1,Qh62_k1,Qh_kv)
    Eh_k1_4245=cal_QE(Eh35_kv,Eh47_k1,Eh62_k1,Eh_kv)
    Qh_k2=cal_QE_k2(Qh17_k2,Qhcalc47_k2,Qh47_kN,Qh35_k2)
    Eh_k2=cal_QE_k2(Eh17_k2,Ehcalc47_k2,Eh47_kN,Eh35_k2)
#    print(Qh_k1_4245)
    for Tj in range(62,-28,-5):#遍历对应的Tj
        Tj=str(Tj)
        if BL[Tj]<=Qh_k1_4245[Tj]:
            delta=cal_delta(float(Toff),float(Ton))
            X_k1=BL[Tj]/Qh_k1_4245[Tj]#某个Tj下的参数
            PLF=1-Cd*(1-X_k1)#某个Tj下的参数
            eh_N=X_k1*Eh_k1_4245[Tj]*delta[Tj]/PLF*nj_N[zone][Tj]
            RH_N=BL[Tj]*(1-delta[Tj])/3.413*nj_N[zone][Tj]
        elif BL[Tj]>Qh_k1_4245[Tj] and BL[Tj]<Qh_k2[Tj]:
            if BL[Tj]>Qh_k1_4245[Tj] and BL[Tj]<Qh_kv[Tj]:
                COP_k1=Qh_k1[Tj]/Eh_k1[Tj]#计算某个Tj下的COP
                COP_kv=Qh_kv[Tj]/Eh_kv[Tj]#计算某个Tj下的COP
                COP_ki=COP_k1+(COP_kv-COP_k1)/(Qh_kv[Tj]-Qh_k1[Tj])*(BL[Tj]-Qh_k1[Tj])
            elif BL[Tj]>=Qh_kv[Tj] and BL[Tj]<Qh_k2[Tj]:
                COP_kv=Qh_kv[Tj]/Eh_kv[Tj]#计算某个Tj下的COP
                COP_k2=Qh_k2[Tj]/Eh_k2[Tj]#计算某个Tj下的COP
                COP_ki=COP_kv+(COP_k2-COP_kv)/(Qh_k2[Tj]-Qh_kv[Tj])*(BL[Tj]-Qh_kv[Tj])
            Eh_ki=BL[Tj]/(3.413*COP_ki)#某个Tj下的功率
            delta=cal_delta(float(Toff),float(Ton))
            eh_N=Eh_ki*delta[Tj]*nj_N[zone][Tj]
            RH_N=BL[Tj]*(1-delta[Tj])/3.413*nj_N[zone][Tj]
        elif BL[Tj]>=Qh_k2[Tj]:
            delta_=cal_delta_(Qh_k2,Eh_k2,float(Toff),float(Ton))
            eh_N=Eh_k2[Tj]*delta_[Tj]*nj_N[zone][Tj]
            RH_N=(BL[Tj]-Qh_k2[Tj]*delta_[Tj])/3.413*nj_N[zone][Tj]

        cigma_eh_N[Tj]=eh_N
        cigma_RH_N[Tj]=RH_N
        HSPF2=(dict_array(nj_N[zone])*dict_array(BL.pop('35'))).sum()/(dict_array(cigama_eh_N).sum()+dict_array(cigma_RH_N).sum())
#    print(Qh_k1)
    print(cigma_eh_N)
    print(cigma_RH_N)

if __name__=='__main__':
    HSPF2()






