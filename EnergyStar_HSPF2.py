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


###计算building load, Equation 4.2-2
def BL(zone,Q):
    '''
    zone(str):对应的区域，如I,IV等
    Q(float):如冷热型，则为A或A2时的制冷量（Qc_95),如单热，则用Qh_47替代
    '''
    res={}
    for Tj in range(62,-28,-5):#遍历对应的Tj
        BL=(Tzl[zone]-Tj)/(Tzl[zone]-5)*C[zone]*Q
        res[str(Tj)]=BL
    return res


###通过线性插值计算一组Tj对应的值
def Values_Tj(x,y):
    '''
    x(list):自变量，此处对应两个温度值
    y(list):应变量，两个温度对应的制冷量或功率
    '''
    if len(x)==2 and len(y)==2:
        res={}
        for Tj in range(62,-28,-5):#遍历对应的Tj
            Q=y[0]+(y[1]-y[0])/(x[1]-x[0])*(Tj-x[0])#计算并记录结果
            res[str(Tj)]=Q
        return res
    else:
        return 'error'


###线性函数，主要服务于Equation 4.2.4-7&8
def liner(k1,k2,kv):
    '''
    k1(dict):k=1时的Qc,Ec
    k2(dict):k=2时的Qc,Ec
    kv(float):k=v时(H2v工况下)的Qc,Ec
    '''
    #计算N
    N=(kv-k1['35'])/(k2['35']-k1['35'])
    M=(k1['62']-k1['47'])/(62-47)*(1-N)+(k2['35']-k2['17'])/(35-17)*N

    #计算k=v时对应制冷量Q或功率E
    res={}
    for Tj in range(62,-28,-5):#遍历对应的Tj
        Q_kv=kv+M*(Tj-35)
        res[str(Tj)]=Q_kv
    return res

###计算delta, Equation 4.2.3-3
def delta(Toff,Ton):
    '''
    Toff(float):低于此温度，停机
    Ton(float):高于此温度，开机
    '''
    res={}
    for Tj in range(62,-28,-5):#遍历对应的Tj
        if Tj<=Toff:
            delta=0
        elif Tj>Toff and Tj <=Ton:
            delta=0.5
        elif Tj>Ton:
            delta=1
        res[str(Tj)]=delta
    return res


###根据4.2.4计算Qh,Eh
def cal_Q(Q35_kv,Q47_k1,Q62_k1,Q_kv):
    '''
    Q35_kv(float):k=v时的制冷量
    Q47_k1(float):k=1,47度时的制冷量
    Q62_k1(float):k=1,62度时的制冷量
    Q_kv(list):Equation 4.2.4-7计算的制冷量
    '''
    #Equation 4.2.4-5
    res={}
    for Tj in range(62,-28,-5):#遍历对应的Tj
        if Tj>=47:
            Qh_k1=Values_Tj([47,62],[Q47_k1,Q62_k1])
        elif Tj>=35 and Tj<47:
            Qh_k1=Values_Tj([35,47],[Q35_kv,Q47_k1])
        elif Tj<35:
            Qh_k1=Q_kv[str(Tj)]
        res[str(Tj)]=Qh_k1
    return res


###计算HSPF2
def HSPF2():
    Qh_k1
    Q_kv=liner()
    Qh_k1=cal_Q(Q35_k1,Q47_k1,Q62_k1,Q_kv)#用Equation 4.2.4-5计算Qh_k=1
    for Tj in range(62,-28,-5):#遍历对应的Tj
        Tj=str(Tj)
        if BL[Tj]<=Qh_k1[Tj]:
            pass
        elif BL[Tj]>Qh_k1[Tj] and BL[Tj]<Qh_k2[Tj]:
            pass
        elif BL[Tj]>=Qh_k2[Tj]:
            pass

#===========================================================================






###计算EER
def EER(Qc,Ec):
    '''
    Qc(dict):制冷量
    Ec(dict):功率
    '''
    res={}
    for Tj in range(67,103,5):#遍历对应的Tj
        Tj=str(Tj)
        EER=Qc[Tj]/Ec[Tj]
        res[str(Tj)]=EER
    return res


###计算SEER, 适用clause 4.1.4.1
def SEER2(Qc_k1,Ec_k1,Qc_k2,Ec_k2,Qc_kv,Ec_kv,BL,Cd=0.25,vs='YES'):
    '''
    Qc_k1(dict):k=1时的Qc
    Ec_k1(dict):k=1时的Ec
    Qc_k2(dict):k=2时的Qc
    Ec_k2(dict):k=2时的Ec
    Qc_kv(dict):k=v时的Qc
    Ec_kv(dict):k=v时的Ec
    BL(dict):building cooling load
    vs(str):varible speed compressor
    '''
#    BL_np=np.array(list(BL.values()))#转化为numpy的array
    Qc_k1.pop('95')#去掉95这个温度值的信息
    Qc_k1_np=np.array(list(Qc_k1.values()))#转化为numpy的array
    Ec_k1.pop('95')#去掉95这个温度值的信息
    Ec_k1_np=np.array(list(Ec_k1.values()))#转化为numpy的array

    Qc_k2.pop('95')#去掉95这个温度值的信息
    Qc_k2_np=np.array(list(Qc_k2.values()))#转化为numpy的array
    Ec_k2.pop('95')#去掉95这个温度值的信息
    Ec_k2_np=np.array(list(Ec_k2.values()))#转化为numpy的array

#    Qc_kv.pop('95')#去掉95这个温度值的信息
#    Qc_kv_np=np.array(list(Qc_kv.values()))#转化为numpy的array
#    Ec_kv.pop('95')#去掉95这个温度值的信息
#    Ec_kv_np=np.array(list(Ec_kv.values()))#转化为numpy的array

    T=[str(i) for i in range(67,103,5)]
    nj_N=dict(zip(T,[0.214,0.231,0.216,0.161,0.104,0.052,0.018,0.004]))#构建table 19的权重因子字典
    SEER={}
    cigma_qc={}
    cigma_ec={}
    for Tj in range(67,103,5):#遍历对应的Tj
        Tj=str(Tj)
        if BL[Tj]<=Qc_k1[Tj]:
            X_k1=BL[Tj]/Qc_k1[Tj]
            PLF=1-Cd*(1-X_k1)
            qc_N=X_k1*Qc_k1[Tj]*nj_N[Tj]
            ec_N=X_k1*Ec_k1[Tj]/PLF*nj_N[Tj]
            SEER[Tj]=qc_N/ec_N
        elif BL[Tj]>Qc_k1[Tj] and BL[Tj]<Qc_k2[Tj]:
            if vs=='YES':
                X_k1=(Qc_k2[Tj]-BL[Tj])/(Qc_k2[Tj]-Qc_k1[Tj])
                X_k2=1-X_k1
                qc_N=(X_k1*Qc_k1[Tj]+X_k2*Qc_k2[Tj])*nj_N[Tj]
                ec_N=(X_k1*Ec_k1[Tj]+X_k2*Ec_k2[Tj])*nj_N[Tj]
                SEER[Tj]=qc_N/ec_N
            else:
                Qc_ki=BL[Tj]
                if BL[Tj]<Qc_kv[Tj]:
                    EER_ki=Qc_k1[Tj]/Ec_k1[Tj]+(Qc_kv[Tj]/Ec_kv[Tj]-Qc_k1[Tj]/Ec_k1[Tj])/(Qc_kv[Tj]-Qc_k1[Tj])*(BL[Tj]-Qc_k1[Tj])
                    Ec_ki=Qc_ki/EER_ki
                elif BL[Tj]>=Qc_kv[Tj]:
                    EER_ki=Qc_kv[Tj]/Ec_kv[Tj]+(Qc_k2[Tj]/Ec_k2[Tj]-Qc_kv[Tj]/Ec_kv[Tj])/(Qc_k2[Tj]-Qc_kv[Tj])*(BL[Tj]-Qc_kv[Tj])
                    Ec_ki=Qc_ki/EER_ki
                qc_N=Qc_ki*nj_N[Tj]
                ec_N=Ec_ki*nj_N[Tj]
                SEER[Tj]=qc_N/ec_N
        elif BL[Tj]>Qc_k2[Tj]:
            qc_N=Qc_k2[Tj]*nj_N[Tj]
            ec_N=Ec_k2[Tj]*nj_N[Tj]
            SEER[Tj]=qc_N/ec_N
        cigma_qc[Tj]=qc_N
        cigma_ec[Tj]=ec_N
        res=sum(list(cigma_qc.values()))/sum(list(cigma_ec.values()))
    return {'SEER2':res,
            'qc_N':cigma_qc,
            'ec_N':cigma_ec,
            }


###把相关数据感性的体现出来
def myplot(Qc_k1,Qc_k2,Qc_kv,Ec_k1,Ec_k2,Ec_kv,EER_k1,EER_k2,EER_kv,ptf='NO'):
    if ptf=='YES':
        print('='*20)
        print('BL:',BL)
        print('Qc_k1:',Qc_k1)
        print('Ec_k1:',Ec_k1)
        print('Qc_k2:',Qc_k2)
        print('Ec_k2:',Ec_k2)
        print('Qc_kv:',Qc_kv)
        print('Ec_kv:',Ec_kv)
#        print('EER_F1:',Qc_k1['67']/Ec_k1['67'])
#        print('EER_B1:',Qc_k1['82']/Ec_k1['82'])
#        print('EER_A2:',Qc_k2['95']/Ec_k2['95'])
#        print('EER_B2:',Qc_k2['82']/Ec_k2['82'])
        print('EER_EV:',Qc_kv['87']/Ec_kv['87'])
        print('EER_k1:',EER_k1)
        print('EER_k2:',EER_k2)
        print('EER_kv:',EER_kv)
        print('qc_N',res['qc_N'])
        print('ec_N',res['ec_N'])
    fig,ax=plt.subplots(figsize=(5, 2.7), layout='constrained')
    T=[str(i) for i in range(67,103,5)]
    ax.plot(T,Qc_k1.values(),label='Qc,k=1')
    ax.plot(T,Qc_k2.values(),label='Qc,k=2')
    ax.plot(T,Qc_kv.values(),label='Qc,k=v')
    ax.plot(T,BL.values(),label='BL')
    ax.plot(T,EER_k1.values(),label='EER,k=1')
    ax.plot(T,EER_k2.values(),label='EER,k=2')
    ax.plot(T,EER_kv.values(),label='EER,k=v')
    ax.legend()
    plt.show()



def single_plt(data,name):
    fig,ax=plt.subplots(figsize=(5, 2.7), layout='constrained')
    T=[str(i) for i in range(67,103,5)]
    ax.plot(T,data.values(),label=name)
    ax.legend()
    plt.show()
    


###主程序入口
if __name__=='__main__':
    data_init={}
    for condition in ['A2','B2','EV','B1','F1']:
        Cap=input(f'请输入{condition}工况下的制冷量：')
        Pow=input(f'请输入{condition}工况下的功率：')
        data_init[condition]=[float(Cap),float(Pow)]#第一个元素为制冷量，第二个为功率，使用时需要注意顺序
    Qc_k1=Values_Tj([67,82],[data_init['F1'][0],data_init['B1'][0]])#计算k=1时的制冷量
    Ec_k1=Values_Tj([67,82],[data_init['F1'][1],data_init['B1'][1]])#计算k=1时的功率
    EER_k1=EER(Qc_k1,Ec_k1)#计算k=1时的能效
    Qc_k2=Values_Tj([82,95],[data_init['B2'][0],data_init['A2'][0]])#计算k=2时的制冷量
    Ec_k2=Values_Tj([82,95],[data_init['B2'][1],data_init['A2'][1]])#计算k=2时的功率
    EER_k2=EER(Qc_k2,Ec_k2)#计算k=2时的能效
    BL=BL(Qc_k2['95'],0.93)#变频为0.93，其他为1
    Qc_kv=liner(Qc_k1,Qc_k2,data_init['EV'][0])#计算k=v时的制冷量
    Ec_kv=liner(Ec_k1,Ec_k2,data_init['EV'][1])#计算k=v时的功率
    EER_kv=EER(Qc_kv,Ec_kv)#计算k=v时的能效
    #计算qc_N, ec_N和SEER2
    res=SEER2(Qc_k1,Ec_k1,Qc_k2,Ec_k2,Qc_kv,Ec_kv,BL,vs='NO')
    print('SEER',res['SEER2'])
    #图标输出,不需要图表就注释掉
    single_plt(res['qc_N'],'qc_N')
    single_plt(res['ec_N'],'ec_N')
    myplot(Qc_k1,Qc_k2,Qc_kv,Ec_k1,Ec_k2,Ec_kv,EER_k1,EER_k2,EER_kv,ptf='NO')#默认不输出详细参数
