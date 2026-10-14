import openpyxl,sys,json
wb=openpyxl.load_workbook(sys.argv[1],data_only=True)
def dump(name,r0,r1,c0,c1):
    ws=wb[name]; print('==',name)
    for r in range(r0,r1+1):
        vals=[ws.cell(r,c).value for c in range(c0,c1+1)]
        if any(v is not None for v in vals):
            out=[]
            for v in vals:
                if isinstance(v,float): out.append(f"{v:,.3f}" if abs(v)<10 else f"{v:,.0f}")
                else: out.append(str(v) if v is not None else '')
            print(r,' | '.join(out))
for a in sys.argv[2:]:
    n,r0,r1,c0,c1=a.split(':'); dump(n,int(r0),int(r1),int(c0),int(c1))
