const {
  Document, Packer, Paragraph, TextRun, ExternalHyperlink,
  AlignmentType, BorderStyle, Header, Footer, PageNumber, UnderlineType
} = require('docx');
const fs = require('fs');

const BLUE      = "1F4E79";
const COLORS = { ai: "1565C0", social: "2E7D32", economy: "BF360C", politics: "4A148C", culture: "00695C" };

function sectionHeader(text, color) {
  return new Paragraph({
    spacing: { before: 400, after: 160 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color, space: 4 } },
    children: [new TextRun({ text, bold: true, size: 32, color, font: "Malgun Gothic" })]
  });
}

function newsItem(num, title, body, source, url) {
  return [
    new Paragraph({
      spacing: { before: 240, after: 60 },
      children: [new TextRun({ text: `${num}. ${title}`, bold: true, size: 22, font: "Malgun Gothic" })]
    }),
    new Paragraph({
      spacing: { before: 0, after: 60 },
      indent: { left: 360 },
      children: [new TextRun({ text: body, size: 20, font: "Malgun Gothic", color: "333333" })]
    }),
    new Paragraph({
      spacing: { before: 0, after: 160 },
      indent: { left: 360 },
      children: [
        new TextRun({ text: "출처: ", size: 18, color: "666666", font: "Malgun Gothic" }),
        new ExternalHyperlink({
          link: url,
          children: [new TextRun({ text: source, size: 18, color: "1565C0", underline: { type: UnderlineType.SINGLE }, font: "Malgun Gothic" })]
        })
      ]
    })
  ];
}

function divider() {
  return new Paragraph({
    spacing: { before: 80, after: 80 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: "DDDDDD", space: 1 } },
    children: []
  });
}

const children = [
  // ── 제목 ──
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 100 },
    children: [new TextRun({ text: "2026년 5월 13일", bold: true, size: 44, color: BLUE, font: "Malgun Gothic" })]
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 80 },
    children: [new TextRun({ text: "일일 뉴스 브리핑", bold: true, size: 52, color: BLUE, font: "Malgun Gothic" })]
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 80 },
    children: [new TextRun({ text: "AI  ·  사회  ·  경제  ·  정치  ·  문화", size: 24, color: "555555", font: "Malgun Gothic" })]
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 400 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: BLUE, space: 4 } },
    children: [new TextRun({ text: "작성일: 2026년 5월 13일 (수)  |  자동 생성 뉴스 다이제스트", size: 18, color: "777777", font: "Malgun Gothic" })]
  }),

  // ── AI 뉴스 ──
  sectionHeader("AI / 인공지능 뉴스", COLORS.ai),
  ...newsItem(1,
    "미국 의회, 대중국 AI 반도체 수출 통제 강화 법안 본격 심의",
    "미국 하원 외교위원회가 첨단 AI 반도체에 대한 중국의 접근을 차단하기 위한 'MATCH Act'와 'Chip Security Act'를 본격 심의했다. 법안은 AI 칩 설계·제조 관련 기술의 중국 이전 차단과 제3국 우회 수출 추적·제재 강화를 골자로 한다. 국내 반도체 기업들은 대미 수출 허가 요건 변화에 대한 신속한 대응 전략 마련이 시급한 상황이다.",
    "전국인력신문", "https://www.kjob.news/news/487805"
  ),
  ...newsItem(2,
    "정부, AI 반도체 실증·상용화 테스트베드 지원 대폭 확대",
    "과학기술정보통신부가 국내 AI 반도체 기업의 실증·상용화를 촉진하기 위해 테스트베드 지원 범위를 대폭 확대한다고 발표했다. 공공기관 데이터센터와 연계한 실증 프로그램을 신설하고 스타트업부터 중견기업까지 단계별 맞춤 지원을 제공한다. AI 반도체 생태계 강화를 통한 미국·중국에 대한 기술 자립도 제고가 핵심 목표다.",
    "대한민국 정책브리핑", "https://m.korea.kr/news/pressReleaseView.do?newsId=156744252"
  ),
  ...newsItem(3,
    "개인정보위, 그록(Grok) 등 AI 오남용 대응 국제 프라이버시 공동선언 채택",
    "개인정보보호위원회가 그록(Grok) 등 생성형 AI 서비스의 콘텐츠 오남용 문제에 대응하기 위해 국제 개인정보 감독기관들과 함께 AI 프라이버시 공동선언을 채택했다. 선언에는 AI 학습 데이터 투명성 확보, 딥페이크 규제 강화, 개인정보 침해 구제 절차 마련 등이 포함됐다. 향후 각국 법제화 논의에서 이 공동선언이 기준점 역할을 할 전망이다.",
    "대한민국 정책브리핑", "https://www.korea.kr/briefing/pressReleaseView.do?newsId=156745494"
  ),
  ...newsItem(4,
    "AI·로봇이 신약을 만드는 시대 — 정부, AI 신약개발 인재 연간 500명 양성",
    "보건복지부와 과학기술정보통신부가 협력해 'AI·로봇 신약개발 인재 양성' 사업을 확대 추진한다. AI 기반 신약 후보 물질 탐색부터 임상 설계까지 전 주기를 아우르는 전문 인력을 연간 500명 규모로 키울 계획이다. 글로벌 제약사들의 AI 신약 경쟁이 심화되는 가운데 한국의 바이오·AI 융합 역량 강화를 서두르려는 조치다.",
    "대한민국 정책브리핑", "https://www.korea.kr/briefing/pressReleaseView.do?newsId=156744622"
  ),
  ...newsItem(5,
    "코스피 8천 선 임박 — AI·반도체 주도 랠리, 시총 7,000조 돌파",
    "AI·반도체 대형주 중심의 랠리가 이어지면서 코스피 시가총액이 7,000조원을 돌파하고 8,000선을 목전에 뒀다. 삼성전자·SK하이닉스·한미반도체가 동반 강세를 보이며 지수를 끌어올리는 가운데, 일부 전문가들은 개별 종목의 쏠림 현상과 단기 변동성 확대에 주의를 당부했다.",
    "Investing.com", "https://kr.investing.com/analysis/article-200457262"
  ),

  divider(),

  // ── 사회 뉴스 ──
  sectionHeader("사회 주요 뉴스", COLORS.social),
  ...newsItem(1,
    "이상민 전 행안부 장관, 항소심서 징역 9년 선고 — 형량 2년 늘어",
    "서울고법 형사1부는 12일 내란 중요임무 종사·직권남용·위증 혐의로 기소된 이상민 전 행정안전부 장관의 항소심에서 징역 9년을 선고했다. 재판부는 '위헌·위법 행위에 적극 관여하고 사후 사법절차에서도 허위 진술을 반복했다'며 1심(징역 7년)보다 형량을 2년 늘렸다. 12·3 비상계엄 당시 언론사 단전·단수를 지시한 혐의가 핵심 사안이었다.",
    "오마이뉴스", "https://www.ohmynews.com/NWS_Web/View/at_pg.aspx?CNTN_CD=A0003233160"
  ),
  ...newsItem(2,
    "미중 무역 대표단 서울 전격 담판 — 베센트·허리펑 13일 회동",
    "스콧 베센트 미국 재무장관과 허리펑 중국 국무원 부총리가 13일 서울에서 무역 협상 테이블에 마주 앉았다. 14~15일 베이징 미중 정상회담을 앞두고 사전 조율 성격으로 이뤄진 이번 만남에서 반도체·배터리·희토류 공급망, 관세 조정, 대중 수출 통제 등이 주요 의제로 다뤄졌다. 한국이 미중 협상의 중립적 개최지로 선택된 것은 외교적 위상 제고로 평가된다.",
    "뉴스핌", "https://www.newspim.com/news/view/20260512001327"
  ),
  ...newsItem(3,
    "빚투 36조 돌파 — 코스피 급등에 레버리지 투자 열풍 경고",
    "코스피가 연일 사상 최고치를 경신하면서 신용융자(빚투) 잔액이 36조원을 돌파했다. 금융당국은 과도한 레버리지 투자가 급락 시 연쇄 손실을 유발할 수 있다며 투자자 주의를 촉구했다. 예적금을 해지하고 주식시장에 뛰어드는 사례가 급증해 '묻지마 투자' 우려가 높아지고 있다.",
    "다음뉴스", "https://v.daum.net/v/20260506180913502"
  ),
  ...newsItem(4,
    "서울 낮 최고기온 27.4°C — 초여름 더위 시작, 자외선 주의보",
    "13일 서울 낮 최고기온이 27.4°C까지 오르며 초여름 더위가 시작됐다. 기상청은 자외선 지수가 '매우 높음' 단계에 이를 것으로 예보하며 야외 활동 시 선크림과 모자 착용을 권고했다. 이번 주 내내 맑고 건조한 날씨가 이어지면서 산불 위험도 높아질 전망이다.",
    "기상청", "https://www.weather.go.kr"
  ),
  ...newsItem(5,
    "국회, 가임기 여성 가임력 검진 도입 토론회 — 저출생 대책 집중 논의",
    "국회에서 가임기 여성을 대상으로 한 가임력 검진 도입 필요성을 논의하는 토론회가 열렸다. 저출생 문제 해결을 위해 건강보험을 활용한 가임력 검진 지원 방안이 제안됐으며, 여야 의원과 산부인과학회 전문가들이 구체적인 정책 로드맵을 논의했다. 가임력 보존 지원을 국가 책임으로 명시하는 법 개정도 검토되고 있다.",
    "뉴스핌", "https://www.newspim.com/news/view/20260512001326"
  ),

  divider(),

  // ── 경제 뉴스 ──
  sectionHeader("경제 뉴스", COLORS.economy),
  ...newsItem(1,
    "코스피 시총 7,000조 돌파 — 8,000선 목전, AI·반도체 랠리 지속",
    "AI·반도체 대형주 중심 랠리에 힘입어 코스피 시가총액이 역대 처음으로 7,000조원을 넘어섰다. 코스피 지수는 5월 12일 7,822.24로 마감하며 사상 최고치를 경신했고, 13일 장 중에도 상승세가 이어지며 8,000선 돌파 기대감이 고조되고 있다. 외국인 투자자의 순매수 전환이 추가 상승의 트리거로 작용하고 있다.",
    "Investing.com", "https://kr.investing.com/analysis/article-200457262"
  ),
  ...newsItem(2,
    "미중 서울 관세 협상 — 반도체·희토류 공급망·환율 폭넓게 논의",
    "서울에서 열린 베센트·허리펑 회동에서 반도체·배터리·희토류 공급망 재편, 관세 조정 로드맵, 위안·달러 환율 등이 폭넓게 논의됐다. 한국 정부도 이번 회동과 연계해 대미 투자 패키지 확대를 제안한 것으로 전해졌다. 협상 결과에 따라 글로벌 공급망 재편 속도와 방향이 크게 달라질 전망이다.",
    "뉴스핌", "https://www.newspim.com/news/view/20260511000025"
  ),
  ...newsItem(3,
    "WTI 100달러 재차 근접 — 휘발유 2,200원대, 물가 재점화 우려",
    "중동 지정학적 불안이 지속되면서 서부텍사스산원유(WTI) 가격이 배럴당 100달러에 다시 바짝 다가섰다. 국내 휘발유 가격이 리터당 2,200원대로 올라선 가운데 정부는 에너지 요금 유예와 고유가 피해지원금 지급으로 서민 부담 완화에 나서고 있다. 인플레이션 재점화 우려로 한국은행의 금리 정책 운용 폭도 좁아지고 있다.",
    "오늘의 클릭", "https://www.cliktoday.com/2026/05/20260512-today-news.html"
  ),
  ...newsItem(4,
    "컬리, 창사 이래 첫 분기 흑자 달성 — IPO 재추진 가시화",
    "마켓컬리를 운영하는 컬리가 창사 이래 처음으로 분기 흑자를 기록했다. 새벽배송 물류 자동화와 원가 절감, 뷰티컬리 등 카테고리 다각화가 실적 개선을 이끌었다. 흑자 전환에 힘입어 기업공개(IPO) 재추진이 가시화되며 기업 가치 재평가 기대감이 높아지고 있다.",
    "오늘의 클릭", "https://www.cliktoday.com/2026/05/20260512-today-news.html"
  ),
  ...newsItem(5,
    "2026년 5월 경제 현황 — 수출 호조 속 내수 회복 속도가 최대 변수",
    "반도체·AI 수출 호조로 한국 경제의 외형 성장은 견조하지만, 내수 소비 회복이 더딘 점이 숙제로 꼽힌다. 고금리 장기화와 가계부채 부담이 소비 심리를 억누르고 있어 정책 당국의 균형 잡힌 대응이 필요하다는 지적이 이어진다. 전문가들은 2분기 실적 시즌 결과가 하반기 전망을 가를 핵심 변수라고 분석했다.",
    "KTN 코리아타운뉴스", "https://koreatownnews.com/%EA%B2%BD%EC%A0%9C%EC%A0%84%EB%A7%9D-2026%EB%85%84-5%EC%9B%94-%EA%B2%BD%EC%A0%9C-%ED%98%84%ED%99%A9-%EB%B6%84%EC%84%9D%EA%B3%BC-%EC%A0%84%EB%A7%9D/"
  ),

  divider(),

  // ── 정치 뉴스 ──
  sectionHeader("정치 뉴스", COLORS.politics),
  ...newsItem(1,
    "이재명 대통령, 베센트 미 재무장관·허리펑 中 부총리 연쇄 접견",
    "이재명 대통령이 13일 청와대에서 스콧 베센트 미국 재무장관과 허리펑 중국 국무원 부총리를 연이어 접견했다. 베센트 장관과의 면담에서는 한미 관세 협상 조기 타결과 대미 투자 확대 협력을 강조했으며, 허리펑 부총리와는 한중 공급망 협력 강화와 무역 균형 개선을 논의했다. 미중 사이에서 균형 외교를 펼치는 이재명 정부의 전략적 행보로 평가된다.",
    "헤럴드경제", "https://biz.heraldcorp.com/article/10736479"
  ),
  ...newsItem(2,
    "이상민 전 장관 항소심 징역 9년 — 내란 관련 사법 절차 계속",
    "12·3 비상계엄 당시 언론사 단전·단수를 지시한 혐의로 재판에 넘겨진 이상민 전 행정안전부 장관이 항소심에서 징역 9년을 선고받았다. 1심보다 2년 늘어난 형량으로, 재판부는 '헌법 질서를 무너뜨리는 데 적극 관여하고 허위 진술을 반복했다'고 판시했다. 내란 관련 재판이 이어지면서 12·3 이후 책임 규명 절차가 계속되고 있다.",
    "경향신문", "https://www.khan.co.kr/article/202605121554011"
  ),
  ...newsItem(3,
    "국민의힘, 6·3 지방선거 '국민선거대책위원회' 공식 출범",
    "국민의힘이 6월 3일 제9회 전국동시지방선거를 앞두고 '국민선거대책위원회'를 여의도 중앙당사에서 출범시켰다. 장동혁 대표가 총괄하는 선대위는 전국 17개 광역단체장 후보 지원에 총력을 기울일 방침이다. 지방선거 D-21을 맞아 여야 선거 전략이 본격 가동되고 있다.",
    "뉴스핌", "https://www.newspim.com/news/view/20260513000001"
  ),
  ...newsItem(4,
    "6·3 지방선거 D-21 — 7개 선거 동시, 후보 등록 막바지",
    "제9회 전국동시지방선거가 6월 3일로 D-21을 맞았다. 광역·기초단체장과 지방의원, 교육감 등 7가지 선거를 동시에 치르는 이번 선거는 이재명 정부 출범 후 첫 전국 단위 선거로 여야 모두 중간 평가 성격으로 바라보고 있다. 후보 최종 등록 마감일이 다가오면서 각 정당의 전략 공천과 후보 단일화 논의가 막바지에 이르렀다.",
    "vote2026", "https://vote2026.kr/"
  ),
  ...newsItem(5,
    "국회 개헌특위, 권력구조 개편 여야 공방 격화 — 대통령제 vs 이원집정부제",
    "국회 개헌특별위원회에서 권력 구조 개편을 둘러싼 여야 공방이 격화되고 있다. 민주당은 4년 중임 대통령제와 결선투표제 도입을 주장하는 반면, 국민의힘은 분권형 대통령제 또는 의원내각제 전환을 선호해 이견이 좁혀지지 않고 있다. 지방선거 이후 개헌 논의에 탄력이 붙을지 주목된다.",
    "오늘의 클릭", "https://www.cliktoday.com/2026/05/20260508-today-news.html"
  ),

  divider(),

  // ── 문화 뉴스 ──
  sectionHeader("문화 / 엔터테인먼트 뉴스", COLORS.culture),
  ...newsItem(1,
    "SBS '멋진 신세계' 임지연, 넷플릭스 비영어권 1위 — 시청률 6.9% 상승세",
    "임지연·허남준 주연의 SBS 금토드라마 '멋진 신세계'가 2화 만에 최고 시청률 6.9%를 기록하며 상승세를 보이고 있다. 조선 악녀가 현대에 빙의한다는 독특한 설정이 글로벌 시청자를 사로잡아 넷플릭스 비영어권 TV쇼 부문 전 세계 1위, TV쇼 전체 2위에 올랐다. '임지연이 곧 경쟁력'이라는 평가와 함께 20% 시청률 돌파 기대감도 고조되고 있다.",
    "스타뉴스", "https://www.starnewskorea.com/broadcast-drama/2026/05/12/2026051208044134580"
  ),
  ...newsItem(2,
    "BTS 월드투어 한창 — 34개 도시 79회 글로벌 공연 순항",
    "방탄소년단(BTS)이 정규 5집 발매와 함께 진행 중인 2026~2027 월드투어가 고양 개막 이후 순항 중이다. 로스앤젤레스·런던·도쿄 등 전 세계 34개 도시 79회 공연이 예정된 역대 최대 규모이며, 각 도시 공연 티켓이 수 분 만에 매진되는 열기를 이어가고 있다. 완전체 컴백에 팬덤 '아미'의 글로벌 결집이 더욱 강화됐다.",
    "한터 뉴스", "https://www.hanteonews.com/ko/article/91206"
  ),
  ...newsItem(3,
    "블랙핑크, 국립중앙박물관 협업 — K팝 최초 전통·현대 문화 융합 기획",
    "블랙핑크가 완전체 컴백을 기념해 국립중앙박물관과 K팝 최초의 협업 프로젝트를 선보였다. 전통 유물을 모티브로 한 앨범 아트워크와 팝업 전시를 통해 한국 문화유산을 세계에 알리는 새로운 시도로 평가받고 있다. 해외 팬들의 한국 관광 수요까지 유발하며 문화·관광 융합 효과가 기대된다.",
    "21세기 이슈", "https://issue.cyberbabarian.com/kpop-tour-comeback-2026/"
  ),
  ...newsItem(4,
    "코르티스 'GREENGREEN' 초동 231만 장 — 5월 음반 차트 정상 석권",
    "신예 K팝 그룹 코르티스(CORTIS)의 두 번째 미니앨범 'GREENGREEN'이 초동 231만 장을 돌파하며 5월 음반 차트 정상을 굳혔다. 한터차트 집계 기준 308만 점을 넘어서며 데뷔 1주년을 앞두고 폭발적인 성장세를 입증했다. 글로벌 팬 커뮤니티의 바이럴 마케팅과 높은 음악성이 시너지를 낸 결과로 분석된다.",
    "한터 뉴스", "https://www.hanteonews.com/ko/article/91206"
  ),
  ...newsItem(5,
    "5월 중순 OTT·극장 신작 대거 공개 — 여름 성수기 앞둔 K콘텐츠 대전",
    "넷플릭스·웨이브·티빙 등 주요 OTT와 극장가에서 5월 중순 이후 여름 성수기를 노린 대작 드라마·영화가 잇달아 공개된다. 한국 오리지널 스릴러부터 로맨틱 판타지, 다큐멘터리까지 장르 다양성이 크게 늘었다는 평가다. 글로벌 OTT의 K콘텐츠 투자 확대로 제작 규모와 퀄리티 모두 사상 최고 수준에 이르고 있다.",
    "나무위키 가요계", "https://namu.wiki/w/2026%EB%85%84%20%EA%B0%80%EC%9A%94%EA%B3%84"
  ),

  // ── 푸터 ──
  new Paragraph({
    spacing: { before: 400, after: 120 },
    border: { top: { style: BorderStyle.SINGLE, size: 6, color: BLUE, space: 4 } },
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "※ 본 브리핑은 Claude Code 자동화 시스템이 2026년 5월 13일 공개된 정보를 기반으로 작성했습니다.", size: 16, color: "888888", font: "Malgun Gothic" })]
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "각 기사의 상세 내용은 출처 링크를 통해 확인하시기 바랍니다.", size: 16, color: "888888", font: "Malgun Gothic" })]
  })
];

const doc = new Document({
  styles: {
    default: { document: { run: { font: "Malgun Gothic", size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Malgun Gothic" },
        paragraph: { spacing: { before: 400, after: 160 }, outlineLevel: 0 } }
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      }
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC", space: 2 } },
          children: [new TextRun({ text: "일일 뉴스 브리핑 | 2026.05.13", font: "Malgun Gothic", size: 18, color: "888888" })]
        })]
      })
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          border: { top: { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC", space: 2 } },
          children: [
            new TextRun({ text: "- ", font: "Malgun Gothic", size: 18, color: "888888" }),
            new TextRun({ children: [PageNumber.CURRENT], font: "Malgun Gothic", size: 18, color: "888888" }),
            new TextRun({ text: " -", font: "Malgun Gothic", size: 18, color: "888888" })
          ]
        })]
      })
    },
    children
  }]
});

const path = require('path');
const OUT = path.join(__dirname, '일간뉴스_2026-05-13.docx');

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(OUT, buf);
  console.log("생성 완료:", OUT);
}).catch(err => { console.error(err); process.exit(1); });
