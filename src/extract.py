import openpyxl, json, re, datetime, sys
src=sys.argv[1]
wb=openpyxl.load_workbook(src,data_only=True)
def num(v):
    try: return round(float(v),3)
    except: return 0
def mask(n):
    n=str(n or '').strip()
    if not n or n=='0': return ''
    if len(n)<=1: return n
    if len(n)==2: return n[0]+'*'
    if len(n)>4: return n[:2]+'*'*(len(n)-3)+n[-1]   # 법인명 등
    return n[0]+'*'*(len(n)-2)+n[-1]
def xd(v):
    if isinstance(v,(int,float)) and v>40000:
        d=datetime.date(1899,12,30)+datetime.timedelta(days=int(v)); return d.strftime('%m/%d')
    if isinstance(v,datetime.datetime): return v.strftime('%m/%d')
    return ''
D={}
k=wb['KPI']
D['meta']={'date':k['B3'].value.strftime('%Y-%m-%d'),'days':k['C3'].value,'elapsed':k['D3'].value,'baseRate':num(k['E3'].value)}
D['kpi']=[{'name':k.cell(r,2).value,'target':num(k.cell(r,3).value),'actual':num(k.cell(r,4).value),'rate':num(k.cell(r,5).value)} for r in range(8,13)]
D['d1']={}
D['d1Excel']={}
prevExcel={k.cell(r,2).value:num(k.cell(r,4).value) for r in range(24,29) if k.cell(r,2).value}
for r in range(16,22):
    if k.cell(r,2).value: D['d1Excel'][k.cell(r,2).value]=num(k.cell(r,3).value)
hist={}
months=[k.cell(7,c).value for c in range(10,20)]
for r in range(8,24):
    a,b=k.cell(r,8).value,k.cell(r,9).value
    if a and b: hist.setdefault(a,{})[b]=[num(k.cell(r,c).value) if k.cell(r,c).value not in ('-',None) else None for c in range(10,20)]
D['hist']={'months':months,'data':hist}
# managers
ws=wb['설계지원체결현황_매니저']
mm=[ws.cell(17,c).value for c in range(5,17)]
mgr=[]
for r in range(21,40):
    nm=ws.cell(r,2).value
    if not nm: continue
    nm=nm.replace('송헤정','송혜정')
    mgr.append({'code':ws.cell(r,1).value,'name':nm,'status':ws.cell(r,3).value,'tenure':ws.cell(r,4).value,
      'perf':[num(ws.cell(r,c).value) if ws.cell(r,c).value not in (None,'') else None for c in range(5,17)],
      'hicap':num(ws.cell(r,17).value),'designs':num(ws.cell(r,18).value)})
D['mgrMonths']=mm
D['mgrAvg']=[num(ws.cell(19,c).value) for c in range(5,17)]
D['managers']=[m for m in mgr if m['status']=='재직' and m['name']!='정지혜']
# manager phones
ws=wb['인사정보']
mph={}
for r in ws.iter_rows(min_row=6,values_only=True):
    if r[3]=='매니저': mph[r[10]]=re.sub(r'\D','',str(r[12] or ''))
for m in D['managers']: m['phone']=mph.get(m['name'].replace('송헤정','송혜정'),'') or mph.get(m['name'],'')
# planner phones
ws=wb['설계사연락처']
ph={}
for r in ws.iter_rows(min_row=8,values_only=True):
    if r[2] and r[3]: ph[str(r[2]).strip()]=re.sub(r'\D','',str(r[3]))
# agencies
ws=wb['지사_인보험']
am=[ws.cell(8,c).value for c in range(7,20)]
ag={}
for r in ws.iter_rows(min_row=10,values_only=True):
    if not r[5]: continue
    ag[r[5]]={'type':r[1],'mgr':r[2],'hq':r[3],'grp':r[4],'name':r[5],'ins':[num(v) for v in r[6:19]],'wd':num(r[19])}
ws=wb['지사_조직']
om=[ws.cell(8,c).value for c in range(7,15)]
for r in ws.iter_rows(min_row=10,values_only=True):
    if r[5] in ag: ag[r[5]].update({'part':[num(v) for v in r[6:14]],'m5':[num(v) for v in r[14:22]],'m20':[num(v) for v in r[22:30]]})
ws=wb['설계지원체결현황_대리점']
for r in ws.iter_rows(min_row=10,values_only=True):
    if r[5] in ag: ag[r[5]]['dz']=[num(v) for v in r[6:14]]
D['agMonths']=am; D['orgMonths']=om
D['agencies']=[a for a in ag.values() if a['mgr']!='정지혜']
# 환산
ws=wb['환산인원현황_매니저']
hw={}
for r in ws.iter_rows(min_row=5,values_only=True):
    if not r[3]: continue
    hw[r[3]]={'mgr':r[0],'ag':r[1],'name':r[2],'code':r[3],'dz':num(r[4]),'cn':num(r[5]),'amt':num(r[6]),'hw':num(r[7]),'near':r[8] or '',
      'br':r[9] or '','brType':r[10] or '','brGoal':r[11] or '','brOk':r[12] or '','brLeft':num(r[13])}
# planners
ws=wb['설계사_인보험']
pl={}
for r in ws.iter_rows(min_row=4,values_only=True):
    if not r[3]: continue
    vals=[num(v) for v in r[5:13]]
    code=str(r[3]).strip()
    if not any(vals) and code not in hw: continue
    pl[code]={'mgr':r[0],'ag':r[1],'name':r[2],'code':code,'new':r[4]=='신규','m':vals[:7],'cur':vals[7]}
D['plMonths']=['4월','5월','6월','7월','8월','9월','10월']
# 지점 하이캡배수: raw_당월건별BD, 상품명 재물·저축 제외, Σ하이캡실적 ÷ Σ월납보험료
try:
    wsb=wb['raw_당월건별BD'];_P=_H=0
    for r in wsb.iter_rows(min_row=9,values_only=True):
        if not r[4] or r[4] in ('재물','저축'): continue
        _P+=num(r[6]);_H+=num(r[7])
    D['branchHicap']=round(_H/_P,2) if _P else None
except KeyError:
    D['branchHicap']=None
# designs
ws=wb['raw_설계지원건별']
des={}
for r in ws.iter_rows(min_row=10,values_only=True):
    if not r[0] or not r[12]: continue
    code=str(r[12]).strip()
    des.setdefault(code,[]).append({'k':r[14],'p':re.sub(r'^무배당','',str(r[15] or '')).replace('현대해상',''),'no':r[17],
      'c':mask(r[18]),'i':mask(r[19]),'d':xd(r[20]),'prem':num(r[21]),'app':xd(r[22]),'st':r[23],'sig':r[24] if r[24] else '','y':r[25]=='Y','cp':num(r[26])})
    if code not in hw and code not in pl:
        pl[code]={'mgr':r[2],'ag':r[1],'name':r[11],'code':code,'new':False,'m':[0]*7,'cur':0}
# bridge
ws=wb['브릿지달성관리']
br=[]
for r in ws.iter_rows(min_row=13,values_only=True):
    if not r[4]: continue
    br.append({'mgr':r[1],'ag':r[2],'name':r[3],'code':str(r[4]).strip(),'type':r[5],'ok':r[6],'dz':num(r[7]),
      'in':[num(r[8]),num(r[9]),num(r[10]),num(r[11]),r[12]],'ju':[num(r[13]),num(r[14]),num(r[15]),num(r[16]),r[17]]})
D['bridge']=[b for b in br if b['mgr']!='정지혜']
for code,p in pl.items():
    p['tel']=ph.get(code,'')
    if code in hw: p['hwx']=hw[code]
agm={a['name']:a['mgr'] for a in ag.values()}
for p in pl.values():
    if not p['mgr']: p['mgr']=agm.get(p['ag'])
D['planners']=[p for p in pl.values() if p['mgr'] not in ('정지혜',)]
D['designs']=des
print(len(D['planners']),len(D['agencies']),len(D['designs']),len(D['bridge']),sum(1 for p in D['planners'] if p['tel']),file=sys.stderr)

# ---------- 전일유입: 앱이 직접 계산 (일자별 스냅샷 누적) ----------
import os
HIST=sys.argv[2] if len(sys.argv)>2 else 'history.json'
hist=json.load(open(HIST)) if os.path.exists(HIST) else {}
today=sys.argv[3] if len(sys.argv)>3 else D['meta']['date']  # 3번째 인자: 마감일 수동 지정(엑셀 날짜가 틀렸을 때)
snap={'kpi':{x['name']:x['actual'] for x in D['kpi']},
      'mgr':{m['name']:{'perf':m['perf'][-1] or 0} for m in D['managers']},
      'hw':{m['name']:sum(p['hwx']['hw'] for p in D['planners'] if p.get('hwx') and p['hwx']['mgr']==m['name']) for m in D['managers']},
      'ag':{a['name']:a['ins'][-1] for a in D['agencies']},
      'pl':{p['code']:[p['cur'],(p.get('hwx') or {}).get('amt',0)] for p in D['planners'] if p['cur'] or p.get('hwx')}}
# 확장 스냅샷 (2026-11 추가): 설계사 설계·체결·환산·근접·브릿지 / 매니저 설계량·하이캡 / 지사 당월 가동·5만·20만
_L=lambda a:(a or [0])[-1] or 0
snap['plx']={p['code']:{'dz':h.get('dz') or 0,'cn':h.get('cn') or 0,'hw':h.get('hw') or 0,'near':h.get('near') or '','br':h.get('br') or '','brOk':h.get('brOk') or '','brLeft':h.get('brLeft') or 0}
             for p in D['planners'] for h in [p.get('hwx') or {}] if h}
for m in D['managers']:
    snap['mgr'][m['name']].update({'designs':m.get('designs') or 0,'hicap':m.get('hicap')})
snap['agx']={a['name']:{'part':_L(a.get('part')),'m5':_L(a.get('m5')),'m20':_L(a.get('m20')),'dz':_L(a.get('dz'))} for a in D['agencies']}
# 최초 1회: 엑셀 '전일기준 당월KPI' 표로 전날 KPI 스냅샷 시드
if not any(d<today for d in hist):
    import datetime as _dt
    pd=(_dt.date.fromisoformat(today)-_dt.timedelta(days=1)).isoformat()
    hist[pd]={'kpi':prevExcel,'seed':'excel-전일기준표'}
hist[today]=snap
prev=max([d for d in hist if d<today],default=None)
D['prevDate']=prev
P=hist[prev] if prev else {}
D['prev']=P
D['d1']={n:round(v-(P.get('kpi',{}).get(n,v)),3) for n,v in snap['kpi'].items()} if prev else D['d1Excel']
json.dump(dict(sorted(hist.items())),open(HIST,'w'),ensure_ascii=False,separators=(',',':'))
print('prev snapshot:',prev,'d1:',D['d1'],file=sys.stderr)
json.dump(D,open(os.path.join(os.path.dirname(os.path.abspath(__file__)),'data.json'),'w'),ensure_ascii=False,separators=(',',':'))
