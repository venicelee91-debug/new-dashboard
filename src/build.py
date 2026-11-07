"""index.html 빌드: python3 src/build.py  (repo 루트에서 실행)
- 데이터: src/data.json (src/extract.py 결과)
- 폰트: 잘난고딕 TTF는 기존 index.html 안의 @font-face를 재사용 (별도 폰트 파일 없음)"""
import re, datetime, os
root=os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
t=open(os.path.join(root,'src/template.html'),encoding='utf-8').read()
old=open(os.path.join(root,'index.html'),encoding='utf-8').read()
m=re.search(r'@font-face\{font-family:"JalnanTitle";[^}]+\}',old)
assert m, '기존 index.html에서 잘난고딕 폰트를 찾지 못함'
data=open(os.path.join(root,'src/data.json'),encoding='utf-8').read().replace('</','<\\/')
t=t.replace('/*FONTS*/',m.group(0)).replace('/*DATA*/',data)
ts=datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=9))).strftime('%Y-%m-%d %H:%M')
t=t.replace('<meta charset="utf-8">','<meta charset="utf-8">\n<meta name="build-time" content="'+ts+' KST">',1)
open(os.path.join(root,'index.html'),'w',encoding='utf-8').write(t); print('index.html',len(t))
