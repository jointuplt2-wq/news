# 일간뉴스 자동화 루틴

당신은 매일 아침 한국어 뉴스 브리핑 DOCX를 자동 생성하는 에이전트입니다.

## 중요: 날짜 계산
이 루틴은 UTC 23:00에 실행되며, 이는 한국시간(KST) 기준 다음날 오전 08:00입니다.
반드시 KST 날짜를 사용하세요:

```bash
KST_DATE=$(TZ='Asia/Seoul' date +%Y-%m-%d)
KST_DATE8=$(TZ='Asia/Seoul' date +%Y%m%d)
```

## 작업 순서

### 1. KST 날짜 확인
```bash
KST_DATE=$(TZ='Asia/Seoul' date +%Y-%m-%d)
KST_DATE8=$(TZ='Asia/Seoul' date +%Y%m%d)
echo "오늘 날짜(KST): $KST_DATE"
```

### 2. 의존성 설치
```bash
npm install
```

### 3. 웹 검색으로 $KST_DATE 기준 한국 주요 뉴스 수집 (각 5개씩)
- AI/인공지능 뉴스
- 사회 뉴스
- 경제 뉴스
- 정치 뉴스
- 문화/연예 뉴스

### 4. 스크립트 파일 생성
기존 파일 `make_news_20260511.js`를 참고하여 `make_news_${KST_DATE8}.js` 파일 생성
- 날짜, 뉴스 내용만 오늘 것으로 교체
- 출력 파일명: `일간뉴스_${KST_DATE}.docx`

### 5. 스크립트 실행
```bash
node make_news_${KST_DATE8}.js
```
생성 확인: `ls -la 일간뉴스_${KST_DATE}.docx`

### 6. GitHub 푸시 — git push 사용 금지, 반드시 MCP 사용
> ⚠️ `git push`는 403 오류로 항상 실패합니다. 반드시 `mcp__github__push_files`를 사용하세요.

DOCX base64 인코딩:
```bash
base64 -w 0 일간뉴스_${KST_DATE}.docx > /tmp/docx_b64.txt
```

MCP로 JS + DOCX 동시 푸시:
```
mcp__github__push_files(
  owner="jointuplt2-wq", repo="news", branch="main",
  message="chore: 일간뉴스 ${KST_DATE}",
  files=[
    {path: "make_news_${KST_DATE8}.js",          content: <파일 전체 내용>},
    {path: "일간뉴스_${KST_DATE}.docx", content: <base64 인코딩된 내용>}
  ]
)
```

푸시 성공 후 반드시 로컈 동기화 (stop hook 경고 방지):
```bash
git fetch origin main && git reset --hard origin/main
```

### 7. 뉴스 요약 채팅 출력 — 이메일 대신 여기에 직접 출력
> ⚠️ Gmail MCP 토큰이 만료 상태입니다. 이메일 대신 채팅 화면에 뉴스 요약을 출력하세요.

아래 형식으로 출력:
```
## 📰 일간뉴스 ${KST_DATE} (요일)

### 🤖 AI · 기술
1. **제목** — 2~3줄 요약
   출처: 매체명

(이하 사회 / 경제 / 정치 / 문화·연예 동일 형식)
```

## 주의사항
- 뉴스는 반드시 KST 오늘 날짜 기준 최신 기사로 작성
- 각 뉴스 항목: 제목(headline) + 2~3문장 요약(body) + 출처
- DOCX 포맷은 기존 파일과 동일하게 유지
- `git commit` / `git push` 절대 사용 금지
- MCP 푸시 후 항상 `git fetch origin main && git reset --hard origin/main` 실행
