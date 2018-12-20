
# coding: utf-8

# In[1]:


from datetime import datetime,timedelta
import numpy as np
import pandas as pd
from scipy.stats import norm
import openpyxl
import time
import sys
global Config


# In[2]:


def is_third_friday(d):
    return d.weekday() == 4 and 15 <= d.day <= 21

def last_day_of_month(any_day):
    next_month = any_day.replace(day=28) + timedelta(days=4)  # this will never fail
    return next_month - timedelta(days=next_month.day)


# In[3]:


def getKZero(Forward_Price,srtikesList):
    upList=([x for x in srtikesList if x>=Forward_Price])
    downList=([x for x in srtikesList if x<Forward_Price])
    if len(upList)>0:
        RoundUpPX = min(upList)
    else:
        RoundUpPX = np.nan
        print ("Error")

    if len(downList)>0:
        RoundDownPX = max(downList)
    else:
        RoundDownPX = np.nan
        print ("Error")
    RoundIncrement = RoundUpPX - RoundDownPX
    KZero =  int(Forward_Price / RoundIncrement)*RoundIncrement
    return KZero

def test_getKZero():
        Forward_Price = 18.55
        srtikesList = [10,10.5,11,11.5,12,12.5,13,13.5,14,14.5,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,32.5,35,37.5,40,42.5,45,47.5,50,55,60,65,70]
        res = getKZero(Forward_Price,srtikesList)
        assert res==8, "Fail"


# In[4]:


def getVariance(f,CURRENT_TIME,contract,VIX):
    Forward_Price = VIX[VIX.SYMBOL==contract]['FWD PRICE'].values[0]
    PUT_GAP_ALLOW = VIX[VIX.SYMBOL==contract]['PUT GAP ALLOW'].values[0]
    CALL_GAP_ALLOW = VIX[VIX.SYMBOL==contract]['CALL GAP ALLOW'].values[0]
    INT_RATE = VIX[VIX.SYMBOL==contract]['INT RATE'].values[0]

    FWD_START_TIME = getExpirationDatetime(VIX,contract)
    #FWD_END_TIME = pd.to_datetime(VIX[VIX.SYMBOL==contract]['FWD_END_TIME'].values[0])

    options = pd.read_excel(f,sheet_name = contract)
    KZero = getKZero(Forward_Price,options['STRIKE'].values)
    T=getTime2Expiration(CURRENT_TIME,FWD_START_TIME)
    #print T
    
    vixPuts=options[['STRIKE','PUT BID','PUT ASK']]
    vixPuts=vixPuts.sort_values('STRIKE',ascending=False)
    vixPuts['ValidStrike']=1
    vixPuts.loc[vixPuts['STRIKE']>KZero,'ValidStrike']=0
    vixPuts.loc[vixPuts['PUT BID']<=0,'ValidStrike']=0
    vixPuts.loc[vixPuts['PUT ASK']<=0,'ValidStrike']=0
    vixPuts.loc[(vixPuts['ValidStrike'].rolling(PUT_GAP_ALLOW).sum()-vixPuts['ValidStrike']==0) & (vixPuts['STRIKE']!=KZero),'ValidStrike']=0   
    vixPuts=vixPuts.sort_values('STRIKE')
    
    vixCalls=options[['STRIKE','CALL BID','CALL ASK']]
    vixCalls=vixCalls.sort_values('STRIKE',ascending=True)
    vixCalls['ValidStrike']=1
    vixCalls.loc[vixCalls['STRIKE']<KZero,'ValidStrike']=0
    vixCalls.loc[vixCalls['CALL BID']<=0,'ValidStrike']=0
    vixCalls.loc[vixCalls['CALL ASK']<=0,'ValidStrike']=0
    vixCalls.loc[(vixCalls['ValidStrike'].rolling(CALL_GAP_ALLOW).sum()-vixCalls['ValidStrike']==0) & (vixCalls['STRIKE']!=KZero),'ValidStrike']=0        

    optionsData = vixPuts.merge(vixCalls, on = "STRIKE", how='inner')
    optionsData=optionsData.sort_values('STRIKE')
    optionsData['ValidStrike'] = (optionsData['ValidStrike_x']+optionsData['ValidStrike_y'])>0
    optionsData['StrikeInterval'] = 0.5*(optionsData[optionsData['ValidStrike']]['STRIKE'].shift(-1)-optionsData[optionsData['ValidStrike']]['STRIKE'].shift(1))

    tails=optionsData.loc[optionsData['ValidStrike'],'StrikeInterval'].index
    optionsData.loc[tails[0],'StrikeInterval'] = 0.5*(optionsData.loc[tails[0]+1,'STRIKE'] - optionsData.loc[tails[0],'STRIKE'])
    optionsData.loc[tails[-1],'StrikeInterval'] = 0.5*(optionsData.loc[tails[-1],'STRIKE'] - optionsData.loc[tails[-1]-1,'STRIKE'] )
    
    optionsData['BID CONTRIB'] = optionsData['PUT BID']
    optionsData.loc[optionsData['STRIKE']>KZero,'BID CONTRIB'] = optionsData['CALL BID']
    optionsData.loc[optionsData['STRIKE']==KZero,'BID CONTRIB'] = 0.5*(optionsData['PUT BID']+optionsData['CALL BID'])
    
    optionsData['ASK CONTRIB'] = optionsData['PUT ASK']
    optionsData.loc[optionsData['STRIKE']>KZero,'ASK CONTRIB'] = optionsData['CALL ASK']
    optionsData.loc[optionsData['STRIKE']==KZero,'ASK CONTRIB'] = 0.5*(optionsData['PUT ASK']+optionsData['CALL ASK'])

    optionsData['BID CONTRIB']=optionsData['BID CONTRIB']*np.exp(INT_RATE*T)*optionsData['StrikeInterval']/pow(optionsData['STRIKE'],2)
    optionsData['ASK CONTRIB']=optionsData['ASK CONTRIB']*np.exp(INT_RATE*T)*optionsData['StrikeInterval']/pow(optionsData['STRIKE'],2)

    result = 0.5*(2/T*optionsData.loc[optionsData['ValidStrike'],'BID CONTRIB'].sum()*100+2/T*optionsData.loc[optionsData['ValidStrike'],'ASK CONTRIB'].sum()*100)
    
    return result


# In[173]:


def getSPXFrontContract(FWD_START_TIME,SPX,OVERRIDE_FORWARD_DATE,ind):
    if ind==1:
        if OVERRIDE_FORWARD_DATE==1:
            FT=SPX.loc[SPX.FWD_START_TIME>FWD_START_TIME,'FWD_START_TIME'].min()
        else:
            FT=SPX.loc[SPX.FWD_START_TIME<FWD_START_TIME,'FWD_START_TIME'].max()
    else:
        FT=SPX.loc[SPX.FWD_START_TIME>FWD_START_TIME,'FWD_START_TIME'].min()
        
    return SPX.loc[SPX.FWD_START_TIME==FT,'SYMBOL'].values[0]

def getSPXBackContract(FWD_END_TIME,SPX,ind):
    FT=SPX.loc[SPX.FWD_START_TIME==FWD_END_TIME]
    
    if FT.shape[0]>0:
        return FT['SYMBOL'].values[0]
    
    if ind==1:
        FT=SPX.loc[SPX.FWD_START_TIME<FWD_END_TIME,'FWD_START_TIME'].max()
    else:
        FT=SPX.loc[SPX.FWD_START_TIME>FWD_END_TIME,'FWD_START_TIME'].min()
        
    return SPX.loc[SPX.FWD_START_TIME==FT,'SYMBOL'].values[0]

def getExpirationDatetime(VIX,contract):
    return pd.to_datetime(VIX[VIX.SYMBOL==contract]['FWD_START_TIME'].values[0])

def getVIXFrontContract(RANK,FWD_START_TIME,CURRENT_TIME):
    if RANK!='FRONT':
        return Config[RANK].SYMBOL
    if (FWD_START_TIME-CURRENT_TIME).days<7:
        return Config['SECOND'].Symbol
    return Config[RANK].Symbol
    
def getVIXBackContract(RANK,FWD_START_TIME,CURRENT_TIME):
    if RANK!='FRONT':
        return Config[RANK].SYMBOL
    if (FWD_START_TIME-CURRENT_TIME).days<7:
        return Config['THIRD'].Symbol
    return Config['SECOND'].Symbol

def getTime2Expiration(CURRENT_TIME,T):
    return (T - CURRENT_TIME).total_seconds()/60/525600


# In[279]:


RANK = 'FRONT'
def GetTV(f,RANK):
    #reading data
    global Config

    Config=pd.read_excel(f,sheet_name = 'Config')
    CURRENT_TIME = Config['FRONT']['CURRENT TIME']# datetime.strptime('11/21/2018  15:21:30', '%m/%d/%Y %H:%M:%S')#datetime.now()
    TVExcel = Config['FRONT']['TV OUTRIGHT']
    contract=Config[RANK].Symbol
    OVERRIDE_FORWARD_DATE = Config[RANK]['OVERRIDE FORWARD DATE']
    USE_OVERRIDE_VVIX = Config[RANK]['USE OVERRIDE VVIX']
    VVIX_OVERNIGHT_HARD_CODE  = Config[RANK]['VVIX OVERNIGHT HARD CODE']
    OUTRIGHT_MIDPOINT= Config[RANK]['OUTRIGHT MIDPOINT']
    HARD_CODE_CUM_PROB= Config[RANK]['HARD CODE CUM PROB']
    OBSERVED_CUM_PROB_RECENT = Config[RANK]['OBSERVED CUM PROB: RECENT DATA']
    USE_OVERRIDE = ((CURRENT_TIME.time()<datetime.strptime("03:00:00",'%H:%M:%S').time()) |
                  ((CURRENT_TIME.time()>datetime.strptime("03:10:00",'%H:%M:%S').time()) &
                      (CURRENT_TIME.time()<datetime.strptime("03:20:00",'%H:%M:%S').time())) |
                  ((CURRENT_TIME.time()>datetime.strptime("09:14:45",'%H:%M:%S').time()) &
                      (CURRENT_TIME.time()<datetime.strptime("09:30:45",'%H:%M:%S').time())) |
                  (CURRENT_TIME.time()>datetime.strptime("16:15:00",'%H:%M:%S').time()) |
                  (USE_OVERRIDE_VVIX==1))
    

    VIX=pd.read_excel(f,sheet_name = 'VIX')
    VIX=VIX[VIX.SYMBOL.apply(lambda x: type(x)()!=0)]
    SPX=pd.read_excel(f,sheet_name = 'SPX')
    SPX=SPX[SPX.SYMBOL.apply(lambda x: type(x)()!=0)]

    VIX['FWD_START_TIME'] = VIX['EXPIRATION'] + timedelta(minutes=9.5*60)
    VIX['FWD_END_TIME'] = VIX['FWD_START_TIME'] + timedelta(days = 30)

    #filtering SPX expirations
    SPX['IsThirdFriday'] = SPX.EXPIRATION.apply(lambda x: is_third_friday(x))
    SPX['WeekDay'] = SPX.EXPIRATION.apply(lambda x: x.weekday())
    SPX['IsLastMonthDay'] = SPX.EXPIRATION.apply(lambda x: last_day_of_month(x)==x)
    SPX=SPX[~((SPX.WeekDay==0) | (SPX.WeekDay==2)) | SPX.IsLastMonthDay]

    SPX['FWD_START_TIME'] = SPX['EXPIRATION'] + timedelta(minutes=16*60+15)
    SPX.loc[SPX['IsThirdFriday'],'FWD_START_TIME'] = SPX['EXPIRATION'] + timedelta(minutes=9.5*60)

    FWD_START_TIME = getExpirationDatetime(VIX,contract)
    FWD_END_TIME = pd.to_datetime(VIX[VIX.SYMBOL==contract]['FWD_END_TIME'].values[0])
    
    SPXFrontContract1 = getSPXFrontContract(FWD_START_TIME,SPX,OVERRIDE_FORWARD_DATE,1)
    SPXFrontContract2 = getSPXFrontContract(FWD_START_TIME,SPX,OVERRIDE_FORWARD_DATE,2)
    SPXBackContract1 = getSPXBackContract(FWD_END_TIME,SPX,1)
    SPXBackContract2 = getSPXBackContract(FWD_END_TIME,SPX,2)

    FRONT_FORWARD1_EXPIRATION = getExpirationDatetime(SPX,SPXFrontContract1)
    FRONT_FORWARD2_EXPIRATION = getExpirationDatetime(SPX,SPXFrontContract2)
    BACK_FORWARD1_EXPIRATION = getExpirationDatetime(SPX,SPXBackContract1)
    BACK_FORWARD2_EXPIRATION = getExpirationDatetime(SPX,SPXBackContract2)

    FRONT_FORWARD1_TIME = getTime2Expiration(CURRENT_TIME,FRONT_FORWARD1_EXPIRATION) 
    FRONT_FORWARD2_TIME = getTime2Expiration(CURRENT_TIME,FRONT_FORWARD2_EXPIRATION) 
    BACK_FORWARD1_TIME = getTime2Expiration(CURRENT_TIME,BACK_FORWARD1_EXPIRATION) 
    BACK_FORWARD2_TIME = getTime2Expiration(CURRENT_TIME,BACK_FORWARD2_EXPIRATION) 
    
    FRONT_FORWARD1_VARIANCE = getVariance(f,CURRENT_TIME,SPXFrontContract1,SPX)
    FRONT_FORWARD2_VARIANCE = getVariance(f,CURRENT_TIME,SPXFrontContract2,SPX)
    BACK_FORWARD1_VARIANCE = getVariance(f,CURRENT_TIME,SPXBackContract1,SPX)
    BACK_FORWARD2_VARIANCE = getVariance(f,CURRENT_TIME,SPXBackContract2,SPX)
    
    VIXFrontContract = getVIXFrontContract(RANK,FWD_START_TIME,CURRENT_TIME)
    VIXBackContract = getVIXBackContract(RANK,FWD_START_TIME,CURRENT_TIME)
    
    FORWARD_START_TIME = getTime2Expiration(CURRENT_TIME,FWD_START_TIME)
    FORWARD_END_TIME = getTime2Expiration(CURRENT_TIME,FWD_END_TIME)
    
    FRONT_FORWARD_INTERP_VARIANCE = FRONT_FORWARD1_VARIANCE if SPXFrontContract1==SPXFrontContract2 else     (FRONT_FORWARD1_TIME*FRONT_FORWARD1_VARIANCE+((FRONT_FORWARD2_TIME*FRONT_FORWARD2_VARIANCE-FRONT_FORWARD1_TIME*FRONT_FORWARD1_VARIANCE)/(FRONT_FORWARD2_TIME-FRONT_FORWARD1_TIME))*(FORWARD_START_TIME-FRONT_FORWARD1_TIME))/FORWARD_START_TIME

    BACK_FORWARD_INTERP_VARIANCE = BACK_FORWARD1_VARIANCE if SPXBackContract1==SPXBackContract2 else     (BACK_FORWARD1_TIME*BACK_FORWARD1_VARIANCE+((BACK_FORWARD2_TIME*BACK_FORWARD2_VARIANCE-BACK_FORWARD1_TIME*BACK_FORWARD1_VARIANCE)/(BACK_FORWARD2_TIME-BACK_FORWARD1_TIME))*(FORWARD_END_TIME-BACK_FORWARD1_TIME))/FORWARD_END_TIME
    
    FORWARD_VOLATILITY = np.sqrt(((BACK_FORWARD_INTERP_VARIANCE*FORWARD_END_TIME-FRONT_FORWARD_INTERP_VARIANCE*FORWARD_START_TIME)/(FORWARD_END_TIME-FORWARD_START_TIME))/100)*100
    
    FRONT_VARIANCE = getVariance(f,CURRENT_TIME,VIXFrontContract,VIX)
    BACK_VARIANCE = getVariance(f,CURRENT_TIME,VIXBackContract,VIX)

    FRONT_FORWARD_PRICE = VIX[VIX.SYMBOL==VIXFrontContract]['FWD PRICE'].values[0]
    BACK_FORWARD_PRICE = VIX[VIX.SYMBOL==VIXBackContract]['FWD PRICE'].values[0]

    options = pd.read_excel(f,sheet_name = VIXFrontContract)
    KZERO_FRONT = getKZero(FRONT_FORWARD_PRICE,options['STRIKE'].values)
    options = pd.read_excel(f,sheet_name = VIXBackContract)
    KZERO_BACK = getKZero(BACK_FORWARD_PRICE,options['STRIKE'].values)

    FRONT_TIME = getTime2Expiration(CURRENT_TIME,getExpirationDatetime(VIX,VIXFrontContract)) 
    BACK_TIME = getTime2Expiration(CURRENT_TIME,getExpirationDatetime(VIX,VIXBackContract)) 
    FRONT_TIME_HOURS = FRONT_TIME*525600
    BACK_TIME_HOURS = BACK_TIME*525600

    FRONT_ADJUSTED_VARIANCE = FRONT_VARIANCE/100-1/FRONT_TIME*pow(FRONT_FORWARD_PRICE/KZERO_FRONT-1,2)
    BACK_ADJUSTED_VARIANCE = BACK_VARIANCE/100-1/BACK_TIME*pow(BACK_FORWARD_PRICE/KZERO_BACK-1,2)

    FRONT_WEIGHT = 1 if VIXFrontContract==VIXBackContract else (BACK_TIME_HOURS-43200)/(BACK_TIME_HOURS-FRONT_TIME_HOURS)
    BACK_WEIGHT = 1 if VIXFrontContract==VIXBackContract else (43200 - FRONT_TIME_HOURS)/(BACK_TIME_HOURS-FRONT_TIME_HOURS)

    VAR = FRONT_VARIANCE if VIXFrontContract==VIXBackContract else ((FRONT_TIME*FRONT_ADJUSTED_VARIANCE*FRONT_WEIGHT)+(BACK_TIME*BACK_ADJUSTED_VARIANCE*BACK_WEIGHT))*1200

    VVIX_ADJUSTMENT = FORWARD_VOLATILITY*((1-np.sqrt(1-(VVIX_OVERNIGHT_HARD_CODE/100*FORWARD_START_TIME))) if USE_OVERRIDE==1 else (1-np.sqrt(1-(VAR/100*FORWARD_START_TIME))))

    IMPLIED_CUMULATIVE_PROBABILITY = min([norm.cdf((FORWARD_VOLATILITY-OUTRIGHT_MIDPOINT)/VVIX_ADJUSTMENT),0.999999])

    CUMULATIVE_PROBABILITY_USED = HARD_CODE_CUM_PROB if USE_OVERRIDE else OBSERVED_CUM_PROB_RECENT

    TV=FORWARD_VOLATILITY+VVIX_ADJUSTMENT*norm.isf(CUMULATIVE_PROBABILITY_USED)

    return CURRENT_TIME, TVExcel, TV


# In[ ]:


if __name__ == '__main__':
    if len(sys.argv) >1 :
        print (sys.argv)
        f=sys.argv[1]
        RANK = 'FRONT'
        df = pd.DataFrame(columns=['CurTime','TV Excel','TV Python'])

        while (True):
            CURRENT_TIME = 'NAN'
            TVExcel = 'NAN'
            TV = 'NAN'

            #try:
            CURRENT_TIME, TVExcel, TV = GetTV(f,RANK)
            #Config=pd.read_excel(f,sheet_name = 'Config')
            #CURRENT_TIME = Config['FRONT']['CALENDAR MIDPOINT']# datetime.strptime('11/21/2018  15:21:30', '%m/%d/%Y %H:%M:%S')#datetime.now()
    
            #except:
            #    pass

            print (CURRENT_TIME, TVExcel, TV)  
            df.loc[df.shape[0]]={'CurTime':CURRENT_TIME,'TV Excel':TVExcel,'TV Python':TV}
            df.to_csv('TV Results.csv',index=False)
            time.sleep(2)

