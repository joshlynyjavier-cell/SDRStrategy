import pandas as pd, numpy as np, sys
pd.set_option('display.width',250); pd.set_option('display.max_columns',30)
df = pd.read_csv(sys.argv[1])
df.columns=['month','fte','leads','demos','qo','logos','arr']
df['arr']=df['arr'].str.replace(r'[$,]','',regex=True).astype(float)
df['month']=pd.to_datetime(df['month'])
order=['1-10','11-50','51-100','101-500','501+']
df['fte']=pd.Categorical(df['fte'],order,ordered=True)
print(df.dtypes); print(df.shape, df.isna().sum().sum(), 'dups', df.duplicated(['month','fte']).sum())
print(df.month.min(), df.month.max(), df.month.nunique(), df.groupby('month').size().unique())
# monotonic funnel check
print('funnel violations', ((df.demos>df.leads)|(df.qo>df.demos)|(df.logos>df.qo)).sum(), 'zero logos', (df.logos==0).sum())
def summ(g):
    s=g[['leads','demos','qo','logos','arr']].sum()
    s['L2D']=s.demos/s.leads; s['D2Q']=s.qo/s.demos; s['Q2W']=s.logos/s.qo; s['L2W']=s.logos/s.leads
    s['ASP']=s.arr/s.logos; s['ARR/lead']=s.arr/s.leads
    return s
tot=summ(df); print('\nTOTAL\n',tot.round(3))
seg=df.groupby('fte',observed=True).apply(summ); seg['arr_share']=seg.arr/seg.arr.sum(); seg['lead_share']=seg.leads/seg.leads.sum(); seg['logo_share']=seg.logos/seg.logos.sum()
print('\nSEGMENT\n',seg.round(3))
mo=df.groupby('month').apply(summ); print('\nMONTHLY\n',mo.round(3))
# Jan22 vs Feb21 and H1 vs H2
for a,b in [('2021-02-01','2022-01-01')]:
    x=df[df.month==a].set_index('fte')[['leads','demos','qo','logos','arr']]; y=df[df.month==b].set_index('fte')[['leads','demos','qo','logos','arr']]
    print('\nJan22 vs Feb21 growth\n',(y/x-1).round(3)); print('total', (y.sum()/x.sum()-1).round(3))
df['period']=np.where(df.month<='2021-07-01','Feb-Jul21','Aug21-Jan22')
p=df.groupby(['fte','period'],observed=True).apply(summ); print('\nHALVES\n',p.round(3))
# Q comparisons
q=df.copy(); q['q']=q.month.dt.to_period('Q'); print(q.groupby(['q']).apply(summ).round(3))
# per-segment monthly ASP and Q2W stdev
g=df.copy(); g['asp']=g.arr/g.logos; g['q2w']=g.logos/g.qo; g['l2d']=g.demos/g.leads
print(g.groupby('fte',observed=True)[['asp','q2w','l2d']].agg(['mean','std','min','max']).round(3))
# Mix decomposition: total ARR growth Feb-Jul vs Aug-Jan into volume, conversion, ASP per segment
h=df.groupby(['period','fte'],observed=True)[['leads','logos','arr']].sum()
A=h.loc['Feb-Jul21']; B=h.loc['Aug21-Jan22']
for d in (A,B): d['c']=d.logos/d.leads; d['asp']=d.arr/d.logos
vol=(B.leads-A.leads)*A.c*A.asp; conv=B.leads*(B.c-A.c)*A.asp; asp=B.leads*B.c*(B.asp-A.asp)
print('\nDECOMP H1->H2\n',pd.DataFrame({'dARR':B.arr-A.arr,'vol':vol,'conv':conv,'asp':asp}).round(0)); print(pd.DataFrame({'dARR':B.arr-A.arr,'vol':vol,'conv':conv,'asp':asp}).sum().round(0))
# Excluding Nov/Dec
ex=df[~df.month.isin(pd.to_datetime(['2021-11-01','2021-12-01']))]
# Lead growth trend (slope) by segment excluding Nov-Dec
for s in order:
    t=ex[ex.fte==s].sort_values('month'); x=np.arange(len(t))
    for col in ['leads','arr']:
        sl=np.polyfit((t.month-t.month.min()).dt.days/30.4, t[col],1)[0]
        print(s,col,'slope/mo',round(sl,1),'as % of mean',round(sl/t[col].mean()*100,1))
