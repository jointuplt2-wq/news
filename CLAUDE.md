# 일간뉴스 자동화 루틴

## 실행 조건
- UTC 23:00 (KST 08:00) 자동 실행
- 오늘 날짜 = KST 기준 (UTC+9), 예: UTC 2026-05-12 23:xx → KST 2026-05-13

## 루틴 순서

### 1. 날짜 확인
```
KST_DATE = 오늘 KST 날짜 (YYYY-MM-DD)
YYYYMMDD  = 날짜 숫자만 (예: 20260513)
```

### 2. npm install
```bash
cd /home/user/news && npm install
```

### 3. 뉴스 검색 (WebSearch)
각 카테고리별 5개 기사 검색:
- **AI·기술**: `AI 인공지능 뉴스 ${KST_DATE}`
- **사회**: `한국 사회 뉴스 ${KST_DATE}`
- **경제**: `한국 경제 뉴스 ${KST_DATE}`
- **정치**: `한국 정치 뉴스 ${KST_DATE}`
- **문화·연예**: `한국 문화 연예 뉴스 ${KST_DATE}`

### 4. JS 파일 생성
- 파일명: `make_news_${YYYYMMDD}.js`
- 기존 `make_news_20260511.js` 구조 동일하게 사용
- 출력파일: `일간뉴스_${KST_DATE}.docx`

### 5. DOCX 생성
```bash
node make_news_${YYYYMMDD}.js
```
생성 확인: `ls -la 일간뉴스_${KST_DATE}.docx`

### 6. GitHub 푸시 (MCP 사용 — git push 사용 금지)

**반드시 `mcp__github__push_files`로 푸시** (git push는 403 오류로 항상 실패)

```bash
# DOCX base64 인코딩
base64 -w 0 일간뉴스_${KST_DATE}.docx > /tmp/docx_b64.txt
```

MCP로 JS + DOCX 동시 푸시:
```
mcp__github__push_files(
  owner="jointuplt2-wq", repo="news", branch="main",
  message="chore: 일간뉴스 ${KST_DATE}",
  files=[
    {path: "make_news_${YYYYMMDD}.js", content: <파일내용>},
    {path: "일간뉴스_${KST_DATE}.docx", content: <base64내용>}
  ]
)
```

**MCP 푸시 성공 후 반드시 로컬 동기화 (stop hook 방지):**
```bash
git fetch origin main && git reset --hard origin/main
```

### 7. 뉴스 요약 채팅 출력 (이메일 대신)

이메일 발송 대신 **채팅 화면에 직접 출력**:

형식:
```
## 📰 일간뉴스 ${KST_DATE} (요일)

### 🤖 AI · 기술
1. **제목** — 2~3줄 요약 / 출처

### 🏙️ 사회
...

### 💰 경제
...

### 🏛️ 정치
...

### 🎭 문화 · 연예
...
```

각 카테고리 5개 기사, 기사당 헤드라인 + 2~3줄 요약 + 출처

## 주의사항
- `git commit` / `git push` 절대 사용 금지 → stop hook 경고 및 403 오류 발생
- MCP 푸시 후 항상 `git fetch origin main && git reset --hard origin/main` 실행
- DOCX는 반드시 base64 인코딩하여 MCP로 전송
