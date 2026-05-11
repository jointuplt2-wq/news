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
    children: [new TextRun({ text: "2026년 5월 8일", bold: true, size: 44, color: BLUE, font: "Malgun Gothic" })]
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
    children: [new TextRun({ text: "작성일: 2026년 5월 8일 (금)  |  자동 생성 뉴스 다이제스트", size: 18, color: "777777", font: "Malgun Gothic" })]
  }),

  // ── AI 뉴스 ──
  sectionHeader("AI / 인공지능 뉴스", COLORS.ai),
  ...newsItem(1,
    "업스테이지, 포털 '다음' 인수 최종 확정",
    "업스테이지가 2026년 5월 7일 포털 '다음' 운영사 AXZ 인수를 최종 확정했다. 자체 개발 거대언어모델 '솔라'를 다음에 적용해 차세대 AI 포털로 재탄생시킨다는 계획으로, 네이버·구글에 맞서는 한국형 AI 검색 생태계의 서막을 알렸다.",
    "Aboda", "https://aboda.kr/entry/%EC%97%85%EC%8A%A4%ED%85%8C%EC%9D%B4%EC%A7%80-%EB%8B%A4%EC%9D%8C-%EC%9D%B8%EC%88%98"
  ),
  ...newsItem(2,
    "제9회 국제인공지능대전(AI EXPO KOREA 2026) 코엑스 개최",
    "아시아 최대 규모 AI 전시회인 제9회 국제인공지능산업대전이 5월 6일~8일 서울 삼성동 코엑스 A홀 전관에서 열렸다. 국내외 AI 기업 수백 곳이 참가해 최신 기술을 선보였다.",
    "AI타임스", "https://www.aitimes.com/"
  ),
  ...newsItem(3,
    "삼성전자, AI 반도체 수요 확대로 사상 최대 실적 달성",
    "삼성전자는 AI 반도체 수요 급증을 기반으로 사상 최대 실적을 달성하며 'AI 중심 반도체·디바이스 기업으로의 전환'을 본격화하고 있다. HBM(고대역폭 메모리) 공급 확대가 실적을 견인했다.",
    "21세기 이슈", "https://issue.cyberbabarian.com/ai-semiconductor-competition-2026/"
  ),
  ...newsItem(4,
    "AI 반도체 특허 패권, 미국에서 중국으로 이동 분석 발표",
    "AI 반도체 특허 패권이 미국에서 중국으로 이동하고 있다는 업계 분석이 나왔다. 중국의 공격적인 특허 출원 전략이 AI 핵심 기술 주도권 경쟁에서 새로운 변수로 부상했다.",
    "AI타임스", "https://www.aitimes.com/"
  ),
  ...newsItem(5,
    "이재명 대통령, 딥마인드 CEO 하사비스 면담 AI 가드레일 필요",
    "이재명 대통령이 구글 딥마인드 CEO 데미스 하사비스(알파고 아버지)를 만나 AI를 통제할 가드레일이 필요하다고 강조했다. 글로벌 AI 거버넌스 협력 방안도 논의됐다.",
    "MBC 뉴스", "https://imnews.imbc.com/replay/2026/nwdesk/article/6818353_37004.html"
  ),

  divider(),

  // ── 사회 뉴스 ──
  sectionHeader("사회 주요 뉴스", COLORS.social),
  ...newsItem(1,
    "어버이날 — 전국 각지서 경로 행사 열려",
    "오늘(8일)은 제64회 어버이날이다. 전국 복지관·경로당에서 기념행사가 열렸으며, 고물가 속에서도 카네이션과 선물을 준비하는 시민들의 발걸음이 이어졌다.",
    "네이트 뉴스", "https://news.nate.com/view/20260507n38470"
  ),
  ...newsItem(2,
    "고유가 피해지원금 60만원 오늘 자정 신청 마감",
    "정부가 기초생활수급자·차상위계층·한부모가족 등 취약계층을 대상으로 지급하는 고유가 피해지원금(60만원) 1차 신청이 오늘(8일) 자정으로 마감된다. 아직 신청하지 않은 대상자는 서둘러 온라인 신청해야 한다.",
    "유용 정보 블로그", "https://patross0303.com/2026-%EA%B3%A0%EC%9C%A0%EA%B0%80-%ED%94%BC%ED%95%B4%EC%A7%80%EC%9B%90%EA%B8%88-60%EB%A7%8C%EC%9B%90-%EB%B0%9B%EB%8A%94-%EB%B2%95-%EC%98%A4%EB%8A%98-%EC%8B%A0%EC%B2%AD-%EC%95%88-%ED%95%98%EB%A9%B4/"
  ),
  ...newsItem(3,
    "코스피, 7490 돌파하며 사상 최고치 재경신",
    "코스피가 7일 전 거래일 대비 105.49p(1.43%) 오른 7,490.05에 마감, 사상 최고치를 다시 썼다. 반도체·AI 관련주가 지수를 견인했으나 개인 투자자들의 포모(FOMO) 심리가 짙어지고 있다는 경고도 나왔다.",
    "SBS 나이트라인", "https://news.sbs.co.kr/news/endPage.do?news_id=N1008551330"
  ),
  ...newsItem(4,
    "개헌안 국회 본회의 투표 불성립 — 8일 재시도 예정",
    "계엄 요건 강화·5·18·부마항쟁 헌법 전문 수록 등을 골자로 한 개헌안이 7일 국회 본회의에 상정됐으나 국민의힘 전원 불참으로 투표 불성립이 선언됐다. 우원식 국회의장은 8일 재상정을 예고했지만 가결 가능성은 낮다.",
    "경향신문", "https://www.khan.co.kr/"
  ),
  ...newsItem(5,
    "5월 8일 전국 날씨 — 낮 최고 22도, 맑음",
    "오늘 전국은 대체로 맑겠으며, 낮 최고기온은 19~22도로 포근한 날씨가 예상된다. 어버이날 나들이객이 많을 것으로 보이며, 다음 주 초 일부 지역에 비 소식이 있다.",
    "국제신문", "https://www.kookje.co.kr/news2011/asp/newsbody.asp?code=0300&key=20260508.22016001868"
  ),

  divider(),

  // ── 경제 뉴스 ──
  sectionHeader("경제 뉴스", COLORS.economy),
  ...newsItem(1,
    "코스피 7490 돌파, 시총 세계 7위권 진입",
    "코스피가 7,490선을 돌파하며 사상 최고치를 경신했다. 한국 증시 시가총액은 약 4.04조 달러를 기록해 프랑스·독일·영국을 앞지르며 글로벌 주요 증시 7위권에 진입했다는 분석이다.",
    "21세기 이슈", "https://issue.cyberbabarian.com/kospi-april-surge-may-strategy-2026/"
  ),
  ...newsItem(2,
    "전기차 보조금 지원 확대 — 역행하는 한국의 선택",
    "주요국이 전기차 보조금을 축소하는 추세 속에서 한국 정부는 오히려 지원 확대를 추진 중이다. 국내 전기차 시장 성장세를 유지하고 소비자 부담을 완화하려는 의도로, 하반기 예산안에 관련 항목이 반영될 전망이다.",
    "21세기 이슈", "https://issue.cyberbabarian.com/ev-subsidy-reduction-2026/"
  ),
  ...newsItem(3,
    "KDI, 2026년 경제성장률 1.9% 전망 — 반도체가 견인",
    "한국개발연구원(KDI)은 2026년 한국 경제가 반도체 수출 호조와 소비 회복세에 힘입어 1.9% 성장할 것으로 전망했다. 이는 2025년 성장률 1.0%를 크게 웃도는 수치다.",
    "KDI", "https://www.kdi.re.kr/research/economy"
  ),
  ...newsItem(4,
    "고유가 피해지원금 60만원 — 오늘 마감, 서둘러야",
    "기초생활수급자·차상위계층·한부모가족 대상 고유가 피해지원금(1인당 60만원)의 1차 신청이 오늘(8일) 자정 종료된다. 누락 시 2차 신청 기간을 기다려야 하며, 정부24 앱 또는 주민센터를 통해 신청할 수 있다.",
    "유용 정보 블로그", "https://patross0303.com/2026-%EA%B3%A0%EC%9C%A0%EA%B0%80-%ED%94%BC%ED%95%B4%EC%A7%80%EC%9B%90%EA%B8%88-60%EB%A7%8C%EC%9B%90-%EB%B0%9B%EB%8A%94-%EB%B2%95-%EC%98%A4%EB%8A%98-%EC%8B%A0%EC%B2%AD-%EC%95%88-%ED%95%98%EB%A9%B4/"
  ),
  ...newsItem(5,
    "5월 주식시장 전망 — 할인율 부담 속 성장 가능성 유지",
    "증권가는 5월 국내 증시가 금리 인하 기대 약화(할인율 부담)에도 반도체·AI 섹터의 견조한 이익 성장에 힘입어 상승 흐름을 이어갈 것으로 전망했다. 5월 중순 이후 실적 발표 시즌이 변곡점이 될 전망이다.",
    "얼리어답터뉴스", "https://eanews.kr/news/950180"
  ),

  divider(),

  // ── 정치 뉴스 ──
  sectionHeader("정치 뉴스", COLORS.politics),
  ...newsItem(1,
    "개헌안 투표 불성립 — 39년 만의 개헌 시도 사실상 무산",
    "5·18·부마항쟁 헌법 전문 수록 및 계엄 요건 강화를 담은 개헌안이 7일 국회 본회의에서 국민의힘 전원 불참으로 투표 불성립됐다. 국민의힘은 이재명 대통령 연임 빌드업이라며 반발했고, 39년 만의 개헌은 사실상 무산됐다.",
    "경향신문", "https://www.khan.co.kr/politics"
  ),
  ...newsItem(2,
    "우원식 의장, 8일 개헌안 재상정 — 가결 가능성 희박",
    "우원식 국회의장은 8일 오전 개헌안을 재상정하겠다고 밝혔다. 그러나 국민의힘의 당론 변경 가능성이 없어 재투표 역시 의결정족수(재적 2/3) 미달이 불가피하다는 관측이 지배적이다.",
    "뉴시스", "https://www.newsis.com/view/NISX20260506_0003619186"
  ),
  ...newsItem(3,
    "한덕수 전 총리, 항소심에서 징역 15년 선고",
    "12·3 비상계엄 내란 중요임무종사 혐의를 받는 한덕수 전 국무총리에게 항소심 재판부가 징역 15년을 선고했다. 1심보다 형량이 높아지면서 항소심 결과에 정치권의 이목이 집중됐다.",
    "MBC 뉴스", "https://imnews.imbc.com/replay/2026/nwdesk/"
  ),
  ...newsItem(4,
    "정진석 전 비서실장, 보궐선거 출마 철회",
    "윤석열 전 대통령의 비서실장인 정진석 전 실장이 충남 공주·부여·청양 국회의원 보궐선거 불출마를 선언했다. 국민의힘 공관위 컷오프 예상이 선제 불출마를 이끌었다는 분석이 나온다.",
    "경향신문", "https://www.khan.co.kr/"
  ),
  ...newsItem(5,
    "이재명 대통령, 하사비스와 AI 거버넌스 논의",
    "이재명 대통령이 구글 딥마인드 CEO 데미스 하사비스와 만나 AI 통제 거버넌스와 한·영 기술 협력 방안을 논의했다. 대통령은 AI가 인류를 위협하지 않도록 국제적 가드레일이 필요하다고 강조했다.",
    "MBC 뉴스", "https://imnews.imbc.com/replay/2026/nwdesk/article/6818353_37004.html"
  ),

  divider(),

  // ── 문화 뉴스 ──
  sectionHeader("문화 / 엔터테인먼트 뉴스", COLORS.culture),
  ...newsItem(1,
    "드라마 '멋진 신세계' 오늘 첫 방송 — 로맨틱 판타지 기대작",
    "로맨틱 판타지 드라마 '멋진 신세계'가 오늘(8일) 첫 선을 보인다. 독특한 세계관과 화려한 캐스팅으로 방영 전부터 화제를 모았으며, OTT와 지상파 동시 편성으로 폭넓은 시청층을 공략한다.",
    "jjinflix", "https://view.jjinflix.com/entry/2026%EB%85%84-5%EC%9B%94-%EC%8B%A0%EC%9E%91-%EB%93%9C%EB%9D%BC%EB%A7%88-%EB%9D%BC%EC%9D%B8%EC%97%85-%EC%B4%9D%EC%A0%95%EB%A6%AC"
  ),
  ...newsItem(2,
    "넷플릭스 2026 한국 라인업 — K-콘텐츠 글로벌 공세 지속",
    "넷플릭스가 2026년 한국 오리지널 라인업을 공개했다. '이 사랑 통역 되나요?', '월간남친', '남편들', '동궁', '스캔들' 등 다양한 장르의 작품이 예정돼 있으며, K-드라마의 글로벌 흥행 기세를 이어갈 전망이다.",
    "Netflix Korea", "https://about.netflix.com/ko/news/next-on-netflix-korea-2026"
  ),
  ...newsItem(3,
    "K-영화 'The World of Love', 프랑스서 주목",
    "윤가은 감독의 한국 드라마틱 영화 'The World of Love'가 5월 6일 글로벌 개봉해 프랑스 현지 매체의 주목을 받았다. 한국 독립영화의 유럽 진출이 확대되는 흐름 속에서 개봉됐다.",
    "Sortiraparis", "https://www.sortiraparis.com/en/what-to-do-in-paris/cinema-series/articles/345017-the-world-of-love-yoon-ga-eun-s-korean-drama"
  ),
  ...newsItem(4,
    "2026년 5월 뮤지컬·연극 라인업 풍성 — 대학로 활기",
    "더뮤지컬이 발표한 2026년 5월 공연 라인업에는 대형 뮤지컬과 소극장 연극이 대거 포함됐다. 어린이날·어버이날 연휴 수요를 겨냥한 가족 공연도 흥행 중이며, 대학로가 오랜만에 활기를 띠고 있다.",
    "더뮤지컬", "https://www.themusical.co.kr/Magazine/Detail?num=5553"
  ),
  ...newsItem(5,
    "인기 게임 '원신', 신규 버전 특별 방송 오늘 밤 공개",
    "글로벌 인기 게임 '원신(Genshin Impact)'의 신규 버전 특별 방송이 오늘(8일) 오후 9시 공식 YouTube 채널에서 진행된다. 국내 팬들 사이에서도 신규 캐릭터·지역 발표에 대한 기대감이 높아지고 있다.",
    "루리웹", "https://bbs.ruliweb.com/news/board/300001/read/2353049"
  ),

  // ── 푸터 ──
  new Paragraph({
    spacing: { before: 400, after: 120 },
    border: { top: { style: BorderStyle.SINGLE, size: 6, color: BLUE, space: 4 } },
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "※ 본 브리핑은 Claude Code 자동화 시스템이 2026년 5월 8일 공개된 정보를 기반으로 작성했습니다.", size: 16, color: "888888", font: "Malgun Gothic" })]
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
          children: [new TextRun({ text: "일일 뉴스 브리핑 | 2026.05.08", font: "Malgun Gothic", size: 18, color: "888888" })]
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
const OUT = path.join(__dirname, 'daily-news-2026-05-08.docx');

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(OUT, buf);
  console.log("생성 완료:", OUT);
}).catch(err => { console.error(err); process.exit(1); });
