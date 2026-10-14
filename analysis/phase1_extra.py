import pandas as pd, numpy as np, sys
pd.set_option('display.width',250)
df=pd.read_csv(sys.argv[1]); df.columns=['month','fte','leads','demos','qo','logos','arr']
df['arr']=df['arr'].str.replace(r'[$,]','',regex=True).astype(float); df['month']=pd.to_datetime(df['month'])
order=['1-10','11-50','51-100','101-500','501+']
piv=lambda c: df.pivot(index='month',columns='fte',values=c)[order]
L,D,Q,W,A=[piv(c) for c in ['leads','demos','qo','logos','arr']]
print('L2D by seg monthly\n',(D/L).round(3)); print('Q2W\n',(W/Q).round(3)); print('ASP\n',(A/W).round(0))
print('lead mix\n',L.div(L.sum(1),axis=0).round(3))
# corr leads vs L2D within seg excl Nov/Dec
m=~L.index.isin(pd.to_datetime(['2021-11-01','2021-12-01']))
for s in order: print(s,'corr(leads,L2D)',round(np.corrcoef(L[s][m],(D/L)[s][m])[0,1],2),'corr(leads,demos)',round(np.corrcoef(L[s][m],D[s][m])[0,1],2))
# mix-adjusted L2W: Jan22 seg rates at Feb21 mix
f,j=pd.Timestamp('2021-02-01'),pd.Timestamp('2022-01-01')
mixF=L.loc[f]/L.loc[f].sum()
print('blended L2W Feb',W.loc[f].sum()/L.loc[f].sum(),'Jan',W.loc[j].sum()/L.loc[j].sum(),'Jan at Feb mix',(mixF*(W.loc[j]/L.loc[j])).sum())
print('blended L2D Feb',D.loc[f].sum()/L.loc[f].sum(),'Jan',D.loc[j].sum()/L.loc[j].sum(),'Jan at Feb mix',(mixF*(D.loc[j]/L.loc[j])).sum())
# Jan vs Feb decomposition (log-based, segment-level): ARR = leads*L2W*ASP; and mix
def decomp(a,b):
    out={}
    for s in order:
        la,lb=L.loc[a,s],L.loc[b,s]; ca,cb=W.loc[a,s]/la,W.loc[b,s]/lb; pa,pb=A.loc[a,s]/W.loc[a,s],A.loc[b,s]/W.loc[b,s]
        dA=A.loc[b,s]-A.loc[a,s]; lr=np.log(A.loc[b,s]/A.loc[a,s]); k=dA/lr
        out[s]=dict(dARR=dA,volume=k*np.log(lb/la),conversion=k*np.log(cb/ca),deal_size=k*np.log(pb/pa))
    o=pd.DataFrame(out).T; o.loc['Total']=o.sum(); return o.round(0)
print('\nLMDI decomposition Feb21->Jan22\n',decomp(f,j))
# window decomposition Feb-Apr vs Aug-Oct (3-mo sums)
w1=pd.to_datetime(['2021-02-01','2021-03-01','2021-04-01']); w2=pd.to_datetime(['2021-08-01','2021-09-01','2021-10-01'])
L.loc['W1']=L.loc[w1].sum(); L.loc['W2']=L.loc[w2].sum()
for X,src in [(W,W),(A,A),(D,D)]: X.loc['W1']=X.loc[w1].sum(); X.loc['W2']=X.loc[w2].sum()
print('\nLMDI Feb-Apr21 -> Aug-Oct21\n',decomp('W1','W2'))
# Lost demos: H2 (Aug-Oct, Jan) at Feb-Jul L2D
h1=pd.date_range('2021-02-01','2021-07-01',freq='MS'); h2=pd.to_datetime(['2021-08-01','2021-09-01','2021-10-01','2022-01-01'])
l2d1=D.loc[h1].sum()/L.loc[h1].sum(); d2w=W.loc[h1].sum()/D.loc[h1].sum(); asp=A.loc[h1].sum()/W.loc[h1].sum()
exp=L.loc[h2].sum()*l2d1; act=D.loc[h2].sum(); gap=exp-act
print('\nL2D H1',l2d1.round(3).to_dict()); print('demo gap (4 non-holiday H2 months)',gap.round(0).to_dict(),'total',gap.sum().round(0))
print('implied ARR at H1 D2W & H2 ASP', (gap*d2w*(A.loc[h2].sum()/W.loc[h2].sum())).round(0).to_dict(), (gap*d2w*(A.loc[h2].sum()/W.loc[h2].sum())).sum().round(0))
# total demos by month
print(D.iloc[:12].sum(1).to_dict())
# workload proxies
s=df.groupby('fte')[['leads','demos','qo','logos','arr']].sum().loc[order]
print(pd.DataFrame({'leads/logo':s.leads/s.logos,'demos/logo':s.demos/s.logos,'qo/logo':s.qo/s.logos,'ARR/demo':s.arr/s.demos,'ARR/QO':s.arr/s.qo}).round(1))
