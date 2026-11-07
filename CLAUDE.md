# 경기AM지점 현황관리 (new-dashboard)

지점장·매니저용 실적/조직 관리 앱. Vercel이 main 푸시를 감지해 index.html을 자동 배포한다.

## 업데이트 순서 (실적표 엑셀을 받았을 때)
1. repo 루트에서 `python3 -I src/extract.py <실적표.xlsx> data/history.json` → `src/data.json` 생성
2. `python3 src/build.py` → 루트 `index.html` 재생성 (잘난고딕 폰트는 기존 index.html에서 재사용)
3. 커밋 작성자: `venicelee91-debug <venicelee91@gmail.com>` (Vercel Hobby는 계정 주인 커밋만 배포)
4. push 후 GitHub commit status로 Vercel 배포 성공 확인. 안 뜨면 빈 커밋으로 재트리거.
5. 사용자에게 index.html 파일도 함께 전달.

## 전일유입 규칙 (항상 지킬 것)
- 전일유입 = 이번 마감 실적 − **직전 마감일 스냅샷** 실적. 엑셀 'D-1일' 표는 쓰지 않는다(백업용 d1Excel만 보관).
- 스냅샷은 `data/history.json`에 마감일(YYYY-MM-DD)별로 누적: kpi, mgr(체결실적), hw(매니저 환산인원), ag(지사 인보험), pl(설계사코드: [당월인보험, 설계지원 체결실적]).
- 앱 안에서는 `D.prev`, `D.prevDate`, `d1Of(kind,key,cur)`로 새 카드에도 같은 방식으로 전일유입을 표시한다.
- history.json은 지우거나 덮어쓰지 말 것 (같은 날짜 재업로드 시 그 날짜만 갱신).

## 디자인 규칙
- 색: 환산인원관리 앱 톤(연회색 바탕 #F3F4F7, 흰 카드 + #DDE0E7 테두리, 남색 #1C3264 중심, 주황/적색은 포인트만)
- 폰트: 전부 여기어때 잘난고딕 TTF(단일 굵기). 작은 글씨는 바탕색 -webkit-text-stroke로 얇게.
- 금액 단위 천원. 설계사 이름 왼쪽 수화기 아이콘(tel:). 고객명은 마스킹.
- 퇴사 매니저(정지혜) 제외, '송헤정' 표기는 '송혜정'으로.
