# -*- coding: utf-8 -*-
"""
@author: 3900594
"""

from pulp import *
import pandas as pd
import numpy
from matplotlib.backends.backend_pdf import PdfPages 
from PIL import Image
import matplotlib.gridspec as gridspec
import os
import subprocess
import datetime
#%matplotlib inline
import matplotlib.pyplot as plt
import tkinter as tk
import tkinter.messagebox as tmsg

FILEPATH_PM_FILE = r'C:\AAA\PY_DSRC\PM.csv'

PM= pd.read_csv(FILEPATH_PM_FILE,index_col=0)


def ButtonClick():
    ARC1_Rate_val=ARC1_Rate_EB.get()
    PM.at['ARC1_Rate','val']=ARC1_Rate_val
    ARC2_Rate_val=ARC2_Rate_EB.get()
    
    PM.at['ARC2_Rate','val']=ARC2_Rate_val
 
    ARCD_Rate_val=ARCD_Rate_EB.get()
    PM.at['ARCD_Rate','val']=ARCD_Rate_val

    WHS_HANDLE_val=WHS_HANDLE_EB.get()
    PM.at['WHS_HANDLE','val']=WHS_HANDLE_val
    
    ARC2MAX_km_val=ARC2MAX_km_EB.get()
    PM.at['ARC2MAX_km','val']=ARC2MAX_km_val
    
    ARCDMAX_km_val=ARCDMAX_km_EB.get()
    PM.at['ARCDMAX_km','val']=ARCDMAX_km_val
    
    print(PM)
    PM.to_csv(r'C:\AAA\PY_DSRC\PM.csv')
    root.destroy()
    
root = tk.Tk()
root.geometry("530x440")
root.title("Parameters input screen")

Explain= tk.PhotoImage(file="C:\AAA\PY_DSRC\Background\explain.png")
 
canvas = tk.Canvas(bg="white", width=506, height=280)
canvas.place(x=10, y=150)
canvas.create_image(0, 0, image=Explain, anchor=tk.NW)


ARC1_Rate_L=tk.Label(root,text="ARC1 Rate 0.00894",font=("Helvetica",10))
ARC1_Rate_L.place(x=20, y=20)
ARC1_Rate_EB=tk.Entry(width=19, font=("Helvetiva",10))
ARC1_Rate_EB.insert(tk.END,PM.at['ARC1_Rate','val'])
ARC1_Rate_EB.place(x=190,y=22)

ARC2_Rate_L=tk.Label(root,text="ARC2 Rate 0.0192",font=("Helvetica",10))
ARC2_Rate_L.place(x=20, y=40)
ARC2_Rate_EB=tk.Entry(width=19, font=("Helvetiva",10))
ARC2_Rate_EB.insert(tk.END,PM.at['ARC2_Rate','val'])
ARC2_Rate_EB.place(x=190,y=42)

ARCD_Rate_L=tk.Label(root,text="ARCD Rate 0.0192",font=("Helvetica",10))
ARCD_Rate_L.place(x=20, y=60)
ARCD_Rate_EB=tk.Entry(width=19, font=("Helvetiva",10))
ARCD_Rate_EB.insert(tk.END,PM.at['ARCD_Rate','val'])
ARCD_Rate_EB.place(x=190,y=62)

WHS_HANDLE_L=tk.Label(root,text="WHS_IN_OUT Rate 3.7567",font=("Helvetica",10))
WHS_HANDLE_L.place(x=20, y=80)
WHS_HANDLE_EB=tk.Entry(width=19, font=("Helvetiva",10))
WHS_HANDLE_EB.insert(tk.END,PM.at['WHS_HANDLE','val'])
WHS_HANDLE_EB.place(x=190,y=82)

ARC2MAX_km_L=tk.Label(root,text="ARC2MAX_km 350",font=("Helvetica",10))
ARC2MAX_km_L.place(x=20, y=100)
ARC2MAX_km_EB=tk.Entry(width=19, font=("Helvetiva",10))
ARC2MAX_km_EB.insert(tk.END,PM.at['ARC2MAX_km','val'])
ARC2MAX_km_EB.place(x=190,y=102)

ARCDMAX_km_L=tk.Label(root,text="ARCDMAX_km 350",font=("Helvetica",10))
ARCDMAX_km_L.place(x=20, y=120)
ARCDMAX_km_EB=tk.Entry(width=19, font=("Helvetiva",10))
ARCDMAX_km_EB.insert(tk.END,PM.at['ARCDMAX_km','val'])
ARCDMAX_km_EB.place(x=190,y=122)

button1=tk.Button(root,text="Execute",font=("Helvetica",10),command=ButtonClick)
button1.place(x=380,y=70)

root.mainloop()



##TT## ファイルディレクトリ
FILEPATH_TRANSPORT_INPUT = r"C:\AAA\PY_DSRC\transport_5_3.csv"
FILEPATH_FACDATA_INPUT = r"C:\AAA\PY_DSRC\facdata_5_3.csv"
FILEPATH_CDCDATA_INPUT = r"C:\AAA\PY_DSRC\CDC.csv"
FILEPATH_RESULT_OUTPUT = r'C:\AAA\PY_DSRC\PDF\RESULT_'
FILEPATH_LP_FILE_OUTPUT = r'C:\AAA\PY_DSRC\LP.TXT'
FILEPATH_BG_FILE = r'C:\AAA\PY_DSRC\Background\yellow.png'


##TT## 共通 PDFファイル名POSTFIX （日付時刻、拡張子）
FILEPATH_RESULT_OUTPUT += "{0:%Y%m%d_%H%M%S}".format(datetime.datetime.now())+".pdf"

im = Image.open(FILEPATH_BG_FILE)
im_list = numpy.asarray(im)


#======================================================================
#辞書、リストのコーナー
# pandasでCSVからデータフレームを作る
data = pd.read_csv(FILEPATH_TRANSPORT_INPUT)
facdata = pd.read_csv(FILEPATH_FACDATA_INPUT)
cdcdata = pd.read_csv(FILEPATH_CDCDATA_INPUT)

#客先のリスト作成
CSTMR=(data['CSTMR'])
CSTMR_L=CSTMR.values.tolist()#Numpyの機能　リスト作成
CUSTOMERS=CSTMR_L

#客先＝デマンドの辞書作成
dmnd=(data[['CSTMR','CBM']])
dmnd_i=dmnd.set_index('CSTMR')
dem=dmnd_i.to_dict()
#辞書からキーで呼べる
demand=dem["CBM"]

#顧客名＝緯度＝経度＝都市辞書作成
CST=(data[['CSTMR','DLA','DLO','CBM','ADDR']])
CST_i=CST.set_index('CSTMR')
CST_i_dict=CST_i.to_dict(orient='index')
CSTLOC=CST_i_dict
#---------------------------------------------


#倉庫のリスト作成
FAC=(facdata['NAME'])
FAC_L=FAC.values.tolist()
FACILITY=FAC_L

#倉庫名＝都市辞書作成
FACCITY=(facdata[['NAME','CITY']])
FACCITY_i=FACCITY.set_index('NAME')
FACCITY_i_dict=FACCITY_i.to_dict(orient='index')
FACCITY_L=FACCITY_i_dict

#倉庫名＝色辞書作成
FACC=(facdata[['NAME','COLOR']])
FACC_i=FACC.set_index('NAME')
FACC_d=FACC_i.to_dict(orient='index')
#辞書からキーを呼べる
FACCOLOR=FACC_d

#倉庫名＝コストの辞書作成
COST=(facdata[['NAME','COST']])
COST_i=COST.set_index('NAME')
COST_d=COST_i.to_dict()
actcost=COST_d['COST']

#倉庫名＝キャパシティ辞書作成

MAX=(facdata[['NAME','CAPA']])
MAX_i=MAX.set_index('NAME')
MAX_d=MAX_i.to_dict()
maxam=MAX_d['CAPA']

#倉庫名＝緯度＝経度＝都市名辞書作成
FACL=(facdata[['NAME','FLAT','FLON','CITY']])  ##PANDA型をFACLにSET
FACL_i=FACL.set_index('NAME')                 ##Indexを自動連番（0~）→ LOCATIONn のファイル内のINDEXに変更
FACL_i_dict=FACL_i.to_dict(orient='index')         ##さらにDictionary型に変換
FACLOC=FACL_i_dict                             ## 冗長

#後方倉庫のリスト作成
CDC=(cdcdata['CDNAME'])
CDC_L=CDC.values.tolist()
CDCLIST=CDC_L

#後方倉庫名＝都市辞書作成
CDCCITY=(cdcdata[['CDNAME','CDCITY']])
CDCCITY_i=CDCCITY.set_index('CDNAME')
CDCCITY_i_dict=CDCCITY_i.to_dict(orient='index')
CDCCITY_L=FACCITY_i_dict

#後方倉庫名＝色辞書作成
CDCC=(cdcdata[['CDNAME','CDCOLOR']])
CDCC_i=CDCC.set_index('CDNAME')
CDCC_d=CDCC_i.to_dict(orient='index')
#辞書からキーを呼べる
CDCCOLOR=CDCC_d

#後方倉庫名＝コストの辞書作成
CDCCOST=(cdcdata[['CDNAME','CDCOST']])
CDCCOST_i=CDCCOST.set_index('CDNAME')
CDCCOST_d=CDCCOST_i.to_dict()
CDCactcost=CDCCOST_d['CDCOST']

#後方倉庫名＝キャパシティ辞書作成

CDCMAX=(cdcdata[['CDNAME','CDCAPA']])
CDCMAX_i=CDCMAX.set_index('CDNAME')
CDCMAX_d=CDCMAX_i.to_dict()
CDCmaxam=CDCMAX_d['CDCAPA']

#後方倉庫名＝緯度＝経度＝都市名辞書作成
CDCL=(cdcdata[['CDNAME','CDLAT','CDLON','CDCITY']])
CDCL_i=CDCL.set_index('CDNAME')
CDCL_i_dict=CDCL_i.to_dict(orient='index')
CDCLOC=CDCL_i_dict
#=================================================================================
'''
   ＜倉庫から客先への距離の自動計算＞  単純な三平方の定理での計算　sqare_root( (緯度1-緯度2/0.011)^2 + (経度1-経度2/0.0091)^2)
     平面２点間距離から曲面２点間距離への係数　1.25倍 とする　（簡易計算）
'''
#ARC1 後方から前線の距離　ARC2 前線から客先の距離  ARCDIRECT 直送の距離 CDC:Central Distribution Center
ARC2MAX_km=PM.at['ARC2MAX_km','val']
ARCDMAX_km=PM.at['ARCDMAX_km','val']
ARC2 = dict()
for j in FACILITY:
   rr= {i:numpy.sqrt(((FACLOC[j]['FLAT']-CSTLOC[i]['DLA'])/0.0111)**2 
                     +((FACLOC[j]['FLON']-CSTLOC[i]['DLO'])/0.0091)**2)*1.25 for i in CUSTOMERS }
   rrr={i:v for (i,v) in rr.items() if v<999000}   ## 4000km（）以下でミニDictionaryを作成
   ARC2[j]=rrr

for j in FACILITY:                       ##TT##  
   for i in CUSTOMERS:                 ##TT##  
      if ARC2[j][i] > ARC2MAX_km:                ##TT##  
         ARC2[j][i] += 1000000000000     ##TT##  

ARC1 = dict()
for z in CDCLIST:
   pp= {j:numpy.sqrt(((FACLOC[j]['FLAT']-CDCLOC[z]['CDLAT'])/0.0111)**2 
                     +((FACLOC[j]['FLON']-CDCLOC[z]['CDLON'])/0.0091)**2)*1.25 for j in FACILITY }
   ppp={j:u for (j,u) in pp.items() if u<999000}
   ARC1[z]=ppp

ARCDIRECT = dict()
for z in CDCLIST:
   qq= {i:numpy.sqrt(((CSTLOC[i]['DLA']-CDCLOC[z]['CDLAT'])/0.0111)**2 
                     +((CSTLOC[i]['DLO']-CDCLOC[z]['CDLON'])/0.0091)**2)*1.25 for i in CUSTOMERS }
   qqq={z:w for (z,w) in qq.items() if w<999000}
   ARCDIRECT[z]=qqq

for z in CDCLIST:                       ##TT##  
   for i in CUSTOMERS:                 ##TT##  
      if ARCDIRECT[z][i] > ARCDMAX_km:                ##TT##  
         ARCDIRECT[z][i] += 1000000000000    ##TT##     
   
   
   
#=====================================================================================================   
#SET PROBLEM VARIABLE (2)問題の定義
prob=LpProblem("FacilityLocation",LpMinimize)

#=========================================================================================

#DECISION VARIABLES (3)変数の設定

#https://qiita.com/mzmttks/items/82ea3a51e4dbea8fbc17 データ取り出し方　

sss=[(i,j,z)for i in CUSTOMERS for j in FACILITY for z in CDCLIST]+[(i,z)for i in CUSTOMERS for z in CDCLIST]

serv_vars=LpVariable.dicts("Service",sss,0) # 非負変数
use_vars=LpVariable.dicts("UserLocation",FACILITY,0,1,LpBinary)
use_D_vars=LpVariable.dicts("User_D_Location",CDCLIST,0,1,LpBinary)
'''
    倉庫設立コストx倉庫　＋　後方ｚ＝倉庫j＝客先iまでの輸送コスト　　actcost＝倉庫代、use_vars＝倉庫を使うかどうか？(0 or 1)
    0.0225 = 後方または工場から前線倉庫への幹線輸送コスト係数　　## 0.0625 =前線倉庫から顧客への輸送コスト係数
    serv_vars[(i,j,z)] ＝ i,j,z ３か所を通過するm3　（Z:後方or工場、J:前線、i:顧客、arc1:z to j、arc2:j to i）
    ARC１ 54M3を　ハノイーホチミン間　1450km 船で　運ぶのに必要な運賃を　700USDとした場合 1m3 1km 運ぶのに必要なコストは　0.00894　USDとする
    ARC2 54M3を　ハノイーホチミン間　1450km トラックで　運ぶのに必要な運賃を　1500USDとした場合 1m3 1km 運ぶのに必要なコストは　0.0192 USDとする
    ARCD 少し高い
    倉庫ハンドリング代（In-Out)  #286usd / 54 m3= 5.2
'''
ARC1_Rate=PM.at['ARC1_Rate','val']
ARC2_Rate=PM.at['ARC2_Rate','val']
ARCD_Rate=PM.at['ARCD_Rate','val']
WHS_HANDLE=PM.at['WHS_HANDLE','val']
# =============================================================================
#  OBJECTIVE FUNCTION (4)評価関数の生成

prob +=lpSum(actcost[j]*use_vars[j] for j in FACILITY)+lpSum(CDCactcost[z]*use_D_vars[z] for z in CDCLIST)+lpSum(((ARC2[j][i]*ARC2_Rate+ARC1[z][j]*ARC1_Rate)*serv_vars[(i,j,z)] for j in FACILITY for i in CUSTOMERS for z in CDCLIST))+lpSum(ARCDIRECT[z][i]*ARCD_Rate*serv_vars[(i,z)] for z in CDCLIST for i in CUSTOMERS)+lpSum((WHS_HANDLE*serv_vars[(i,j,z)] for j in FACILITY for i in CUSTOMERS for z in CDCLIST))

# =============================================================================
#CONSTRAINS (5)制約条件の生成

for i in CUSTOMERS: 
    for j in FACILITY:
#      prob += lpSum((serv_vars[(i,j,z)] for j in FACILITY for z in CDCLIST for i in CUSTOMERS))==demand[i] 
      prob += lpSum((serv_vars[(i,j,z)] for j in FACILITY for z in CDCLIST))+lpSum((serv_vars[(i,z)] for z in CDCLIST))==demand[i] 

for j in FACILITY:
      prob += lpSum(serv_vars[(i,j,z)] for i in CUSTOMERS for z in CDCLIST)<=maxam[j]*use_vars[j]

for z in CDCLIST:
      prob += lpSum(serv_vars[(i,z)] for i in CUSTOMERS for z in CDCLIST)<=CDCmaxam[z]*use_D_vars[z]
      
      
for i in CUSTOMERS:
     for j in FACILITY:
         for z in CDCLIST:
             prob +=serv_vars[(i,j,z)] <= demand[i]*use_vars[j]

for i in CUSTOMERS:
     for j in FACILITY:
         for z in CDCLIST:
             prob +=serv_vars[(i,z)]<= demand[i]*use_D_vars[z]

              
# =============================================================================
        #SOLVE (7)求解
prob.solve()
print("Status:", LpStatus[prob.status])

#PRINT OPTIMAL SOLUTION (8)結果の確認
aaa=print("The cost for one year =","{:,.2f}".format(value(prob.objective)))

# =============================================================================

#グラフのサイズ、解像度の指定       
#http://ailaby.com/matplotlib_fig/
plt.figure(figsize=(16, 9), dpi=80)
fig = plt.figure(figsize=(14, 9), dpi=80)
gs = gridspec.GridSpec(1,3)
plt.subplot(gs[0, 0])
plt.imshow(im_list,extent=[102, 109.9,8.5, 23.4])

#求解値集合(PROB.VARIABLES)に対して
dFacM3 = dict() ##TT## 前線倉庫別　M3　計算用のDictionary
dCDCM3 = dict() ##TT## 後方倉庫別　M3　計算用のDictionary
for v in prob.variables():
    if v.varValue > 0 :
     print(v.name,"=",v.varValue)
#もしServiceという字が求解値の名前に含まれるなら
     if "Service" in v.name:
        #Service_('Location1',_'FAC1')
        kyaku=v.name[v.name.find("S")+10:v.name.find(",")-1]#Location1
        souko=v.name[v.name.find("FAC"):v.name.find(")")-9]#FAC1
        kouhou=v.name[v.name.find("CD"):v.name.find(")")-1]#CD01
#=ARC1ARC2 Graph======================================================================

    #もし求解値がゼロ以上であるなら
        if v.varValue > 0 :
#倉庫がｋである間
            for k in FACILITY:
#もしkがsoukoリストにあたるなら
              if k == souko :
#              緯度、経度で座標をとりデマンドでサイズをとる散布図を書く　色は色辞書を参照する
               CST_LA=CSTLOC[kyaku]["DLA"]
               CST_LO=CSTLOC[kyaku]["DLO"]
               CST_DMD=CSTLOC[kyaku]["CBM"]/100
               CST_ADD=CSTLOC[kyaku]["ADDR"]
               plt.scatter(CST_LO,CST_LA,c=FACCOLOR[k]['COLOR'],s=CST_DMD, alpha=0.5,edgecolors=FACCOLOR[k]['COLOR'], linewidths="2")
               plt.grid(True)
           #   タイトルを書く     
               plt.title("Vietnam network COST=   "+"{:,.2f}".format(value(prob.objective)),fontsize='8',wrap=True)
           #    都市名のアノテーション入れる
               plt.annotate(CST_ADD, xy=(CST_LO, CST_LA),fontsize='4',fontname='monospace')
# 前線倉庫矢印（ベクトル）の始点
               X = (FACLOC[souko]['FLON'])#105.8546219 
               Y = (FACLOC[souko]['FLAT'])#21.02891922
# 矢印（ベクトル）の成分
               U = CST_LO-X
               V = CST_LA-Y
               
# 矢印（ベクトル）を書く
#http://seesaawiki.jp/met-python/d/matplotlib/plot #https://anaconda.org/conda-forge/basemap #conda install -c conda-forge basemap
               plt.quiver(X,Y,U,V,angles='xy',scale_units='xy',width=0.001,color='b',scale=1,headlength=30,headwidth=30)

#　後方倉庫矢印（ベクトルの始点）               
               XK=CDCLOC[kouhou]['CDLON']
               YK=CDCLOC[kouhou]['CDLAT']
# 後方倉庫矢印（ベクトル）の成分
               UK=X-XK
               VK=Y-YK

               plt.quiver(XK,YK,UK,VK,angles='xy',scale_units='xy',width=0.01,color='r',scale=1,headlength=15,headwidth=15)

#=========#TT## 倉庫別　M3　計算　

               if k in dFacM3:                ##TT##  Dictionary dFacCostのキーにｋがあるなら
                   dFacM3[k]+=v.varValue ##TT##  Dictionary dFacCostの値にはv.varValueを加算
#                   print(k,v.varValue)
               else:                            ##TT##  なければ
                   dFacM3[k]=v.varValue        ##TT##  Dictionary dFacCostに新規にキーとv.varValueを追加
#=DIRECT  Graph======================================================================
#もし求解値がゼロ以上であるなら
        if v.varValue > 0 :
#倉庫がｋである間
            for c in CDCLIST:
#もしkがsoukoリストにあたるなら
              if v.name.count(',')==1 :
#              緯度、経度で座標をとりデマンドでサイズをとる散布図を書く　色は色辞書を参照する
               zCST_LA=CSTLOC[kyaku]["DLA"]
               zCST_LO=CSTLOC[kyaku]["DLO"]
               zCST_DMD=CSTLOC[kyaku]["CBM"]/100
               zCST_ADD=CSTLOC[kyaku]["ADDR"]
               plt.scatter(zCST_LO,zCST_LA,c=CDCCOLOR[c]['CDCOLOR'],s=zCST_DMD, alpha=0.5,edgecolors=CDCCOLOR[c]['CDCOLOR'], linewidths="2")
               plt.grid(True)
               plt.annotate(zCST_ADD, xy=(zCST_LO, zCST_LA),fontsize='4',fontname='monospace')
# 前線倉庫矢印（ベクトル）の始点
               XZ = (CDCLOC[kouhou]['CDLON'])#105.8546219 
               YZ = (CDCLOC[kouhou]['CDLAT'])#21.02891922
# 矢印（ベクトル）の成分
               UZ = zCST_LO-XZ
               VZ = zCST_LA-YZ
               
# 矢印（ベクトル）を書く
#http://seesaawiki.jp/met-python/d/matplotlib/plot #https://anaconda.org/conda-forge/basemap #conda install -c conda-forge basemap
               plt.quiver(XZ,YZ,UZ,VZ,angles='xy',scale_units='xy',width=0.001,color='y',scale=1,headlength=30,headwidth=30)



               if c in dCDCM3:                ##TT##  Dictionary dFacCostのキーにｋがあるなら
                   dCDCM3[c]+=v.varValue ##TT##  Dictionary dFacCostの値にはv.varValueを加算
#                   print(k,v.varValue)
               else:                            ##TT##  なければ
                   dCDCM3[c]=v.varValue        ##TT##  Dictionary dFacCostに新規にキーとv.varValueを追加

#
TOL= .00001
m=0.98
h=0.98
plt.subplot(gs[0,1:])
#plt.subplot(gs[0,1:]).tick_params(labelbottom="off",bottom="off")
#plt.text(0.01,0.01,sum(demand.values())
# =============================================================================
plt.text(0.03,h,'------------------------------------< Condition Parameters>------------------------------------------------------------',fontsize=6,fontname='monospace')
h=h-0.014
plt.text(0.03,h,'Total Demand M3 ==',fontsize=6,fontname='monospace')
plt.text(0.17,h,sum(demand.values()),fontsize=6,fontname='monospace')
plt.text(0.26,h,"Routing simulated==",fontsize=6,fontname='monospace')
plt.text(0.40,h,len(sss),fontsize=6,fontname='monospace')


h=h-0.014

plt.text(0.03,h,"ARC1 frtRate($/m3/km) :"+str(ARC1_Rate),fontsize=6,fontname='monospace')
plt.text(0.26,h,"ARC2 frtRate($/m3/km:"+str(ARC2_Rate),fontsize=6,fontname='monospace')
plt.text(0.47,h,"ARCD frtRate($/m3/km) :"+str(ARCD_Rate),fontsize=6,fontname='monospace')
h=h-0.014

plt.text(0.03,h,"ARC2 Max distance(km) :"+str(ARC2MAX_km),fontsize=6,fontname='monospace')
plt.text(0.26,h,"ARCD Max distance(km) :"+str(ARCDMAX_km),fontsize=6,fontname='monospace')
plt.text(0.47,h,"WHS IN-OUT Rate($/m3) :"+str(WHS_HANDLE),fontsize=6,fontname='monospace')
h=h-0.015
# =============================================================================
plt.text(0.03,h,'------------------------------------< Simulation Result >--------------------------------------------------------------',fontsize=6,fontname='monospace')
h=h-0.02
ttt="{:,.2f}".format(value(prob.objective))
plt.text(0.03,h,"Yearly cost (USD)==:"+ttt,fontsize=8,fontname='monospace')
h=h-0.02
for f in FACILITY:
     if use_vars[f].varValue>TOL:
         plt.text(0.03,h,("{:""<6}".format(f)+"{:""<10}".format(FACLOC[f]['CITY'])+"{:,.0f}".format(dFacM3[f]).rjust(9)+" CBM"+
               " LAT="+"{:,.3f}".format(FACLOC[f]['FLAT'])+
               " LON="+"{:,.3f}".format(FACLOC[f]['FLON'])+" Whs_rent cost $ ",actcost[f],"Whs Capacity m3  "+str(maxam[f])),wrap=True,fontsize=6,fontname='monospace')
         h=h-0.014

# =============================================================================
for c in CDCLIST:
     if use_D_vars[c].varValue>TOL:
         plt.text(0.03,h,("{:""<6}".format(c)+"{:""<10}".format(CDCLOC[c]['CDCITY'])+"{:,.0f}".format(dCDCM3[c]).rjust(9)+" CBM"+
               " LAT="+"{:,.3f}".format(CDCLOC[c]['CDLAT'])+
               " LON="+"{:,.3f}".format(CDCLOC[c]['CDLON'])+" CDC_SET_COST ",CDCactcost[c]),wrap=True,fontsize=6,fontname='monospace')
         h=h-0.016

plt.text(0.03,h,'------------------------------------< Simulation Detail by each arc >--------------------------------------------------',fontsize=6,fontname='monospace')
h=h-0.01
plt.text(0.03,h,'Simulation ID',fontsize=4,fontname='monospace')
plt.text(0.11,h,'Kouhou DC',fontsize=4,fontname='monospace')
plt.text(0.18,h,'Zensen DC',fontsize=4,fontname='monospace')
plt.text(0.25,h,'Customer',fontsize=4,fontname='monospace')
plt.text(0.33,h,'Distance1',fontsize=4,fontname='monospace')
plt.text(0.39,h,'Distance2',fontsize=4,fontname='monospace')
plt.text(0.45,h,'Direct',fontsize=4,fontname='monospace')   
plt.text(0.51,h,'Demand M3',fontsize=4,fontname='monospace')
plt.text(0.58,h,'FRT($)',fontsize=4,fontname='monospace')
plt.text(0.63,h,'FRT/M3($)',fontsize=4,fontname='monospace')
plt.text(0.69,h,'Whs HDL($)',fontsize=4,fontname='monospace')

         
h=h-0.014            
for v in prob.variables():
        if v.varValue > 0 :
          if "Service" in v.name:
            kyaku2=v.name[v.name.find("S")+10:v.name.find(",")-1]
            souko=v.name[v.name.find("FAC"):v.name.find(")")-9]#FAC1
            kouhou2=v.name[v.name.find("CD"):v.name.find(")")-1]#CD01
            if souko=='':
                 WHS_or_DIR='  ------'
                 xxx='------'
                 yyy='------'
                 zzz="{:,.0f}".format(ARCDIRECT[kouhou2][kyaku2])+"km"
                 FRT="{:,.0f}".format(ARCDIRECT[kouhou2][kyaku2]*v.varValue*ARCD_Rate).rjust(9)
                 FRTM3="{:,.0f}".format(ARCDIRECT[kouhou2][kyaku2]*ARCD_Rate).rjust(9)
                 WHDL="{:,.0f}".format(v.varValue*0).rjust(9)
            else:
                 WHS_or_DIR=FACLOC[souko]['CITY']
                 xxx="{:,.0f}".format(ARC2[souko][kyaku2])+"km"
                 yyy="{:,.0f}".format(ARC1[kouhou2][souko])+"km"
                 zzz='------'
                 FRT="{:,.0f}".format((ARC2[souko][kyaku2]*ARC2_Rate+ARC1[kouhou2][souko]*ARC1_Rate)*v.varValue).rjust(9)
                 FRTM3="{:,.1f}".format((ARC2[souko][kyaku2]*ARC2_Rate+ARC1[kouhou2][souko]*ARC1_Rate)).rjust(9)
                 WHDL="{:,.0f}".format(v.varValue*WHS_HANDLE).rjust(9)


            
            plt.text(0.03,h,kyaku2,fontsize=4,fontname='monospace')
            plt.text(0.09,h,r"|",fontsize=4,fontname='monospace')
            plt.text(0.11,h,CDCLOC[kouhou2]['CDCITY'],fontsize=4,fontname='monospace')
            plt.text(0.18,h,WHS_or_DIR,fontsize=4,fontname='monospace')
            plt.text(0.25,h,CSTLOC[kyaku2]["ADDR"],fontsize=4,fontname='monospace')
            plt.text(0.31,h,r"|",fontsize=4,fontname='monospace')
            plt.text(0.33,h,yyy,fontsize=4,fontname='monospace')
            plt.text(0.39,h,xxx,fontsize=4,fontname='monospace')
            plt.text(0.45,h,zzz,fontsize=4,fontname='monospace')
            plt.text(0.49,h,r"|",fontsize=4,fontname='monospace')
            plt.text(0.51,h,"{:,.0f}".format(v.varValue).rjust(9),fontsize=4,fontname='monospace')
            plt.text(0.58,h,FRT,fontsize=4,fontname='monospace')
            plt.text(0.63,h,FRTM3,fontsize=4,fontname='monospace')
            plt.text(0.69,h,WHDL,fontsize=4,fontname='monospace')
            
            h=h-0.01
 
# =============================================================================
'''
   PULP の挙動（モデル）をみるには　モデルを小さくして　下記のLPファイルにかかれたモデルをチェックしてみる
   prob.writeLP(FILEPATH_LP_FILE_OUTPUT)

'''

# PDFとして指定されたディレクトリに図を保存する
pp = PdfPages(FILEPATH_RESULT_OUTPUT)
pp.savefig(fig)
pp.close()
#[(i,j)for i in CUSTOMERS for j in FACILITY]
#print(prob.variables)
#print(prob)



print("Total demand",sum(demand.values()))
print(dFacM3)

NAME = FILEPATH_RESULT_OUTPUT ##TT#

#PDF を　クロームもしくはアクロバットで自動で開く（あとでかならず閉じること）
# ======= PDF アプリの　起動
#returncode = subprocess.call(['C:\Program Files (x86)\Google\Chrome\Application\chrome.exe', NAME])
returncode = subprocess.call(['C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe', NAME])

