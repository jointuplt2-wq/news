const {
  Document, Packer, Paragraph, TextRun,
  AlignmentType, BorderStyle, Header, Footer, PageNumber,
} = require('docx');
const fs = require('fs');

const BLUE      = "1F4E79";
const DARK_BLUE = "2E75B6";
const GRAY      = "595959";
const GREEN     = "1A6B3A";
const RED       = "7B2C2C";
const ORANGE    = "C55A11";
const PURPLE    = "4A235A";

function sectionHeader(text, color) {
  return new Paragraph({
    spacing: { before: 360, after: 120 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: color || DARK_BLUE, space: 4 } },
    children: [new TextRun({ text, bold: true, size: 32, color: color || BLUE, font: "맑은 고딕" })]
  });
}

function newsItem(num, title, body, source) {
  return [
    new Paragraph({
      spacing: { before: 200, after: 60 },
      children: [new TextRun({ text: `${num}. ${title}`, bold: true, size: 24, color: "000000", font: "맑은 고딕" })]
    }),
    new Paragraph({
      spacing: { before: 0, after: 60 },
      indent: { left: 360 },
      children: [new TextRun({ text: body, size: 20, color: GRAY, font: "맑은 고딕" })]
    }),
    new Paragraph({
      spacing: { before: 0, after: 120 },
      indent: { left: 360 },
      children: [new TextRun({ text: `[출처: ${source}]`, size: 18, color: "888888", italics: true, font: "맑은 고딕" })]
    }),
  ];
}

function divider() {
  return new Paragraph({
    children: [],
    spacing: { before: 80, after: 80 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 1, color: "DDDDDD", space: 2 } }
  });
}

// ─── AI 뉴스 ─────────────────────────────────────────────────────────────────
const aiNews = [
  [
    "OpenAI, GPT-5.5 공식 출시 — 에이전틱 코딩·과학 연구 역량 대폭 강화",
    "OpenAI가 4월 23일 최신 모델 GPT-5.5를 공식 출시했다. 자율적 코드 작성·디버깅, 온라인 리서치, 데이터 분석, 소프트웨어 조작 능력이 대폭 향상됐으며 API에서도 동시 제공된다. 샘 올트먼 CEO는 'GPT-5.5가 스스로 출시 파티를 계획했다'고 밝혀 화제를 모았다.",
    "TechCrunch, CNBC"
  ],
  [
    "Google DeepMind AI 임상 보조 시스템, 의사 평가서 GPT-5.4 능가",
    "Google DeepMind의 AI 임상 진료 보조 시스템이 98개 쿼리 기반 맹검 평가에서 OpenAI GPT-5.4-thinking-with-search를 63대 30으로 앞질렀다. 실제 의사들이 평가자로 참여해 신뢰도를 높였으며, 의료 AI 경쟁이 갈수록 심화되고 있다.",
    "WinBuzzer"
  ],
  [
    "국민성장펀드, AI 기업 업스테이지에 5,600억 원 투자 — '소버린 AI' 육성",
    "한국 정부 주도의 국민성장펀드가 국내 AI 스타트업 업스테이지에 5,600억 원 규모의 투자를 결정했다. '소버린 AI(자주적 AI) 개발' 지원을 핵심 목표로 설정하며 국가 AI 경쟁력 강화 전략을 구체화했다.",
    "AI타임스"
  ],
  [
    "휴머노이드 로봇 대량 생산 본격화 — 연내 수천 대 가정 보급 목표",
    "글로벌 로봇 기업들이 가정용 휴머노이드 로봇 대량 생산에 착수했다. 2026년 안에 수천 대 규모 보급을 목표로 하며 가사 보조·노인 돌봄 등 실생활 적용을 앞당기고 있다. 로봇이 일상으로 진입하는 시점이 빠르게 가까워지고 있다.",
    "AI타임스"
  ],
  [
    "중국, GPU 없이 자국 CPU만으로 엑사스케일 슈퍼컴퓨터 개발 추진",
    "미국의 GPU 수출 규제에 대응해 중국이 자국산 CPU만을 활용한 엑사스케일 슈퍼컴퓨터 개발에 나섰다. AI 연산 자립화 전략의 일환으로, 반도체 기술 독립을 향한 중국의 의지가 한층 강해지고 있다.",
    "인공지능신문"
  ]
];

// ─── 사회 뉴스 ───────────────────────────────────────────────────────────────
const socialNews = [
  [
    "호르무즈 해협 한국 선사 선박 폭발 — 피격 여부 정부 긴급 확인",
    "호르무즈 해협에서 한국 선사가 운용하는 선박에서 폭발 사고가 발생했다. 정부는 피격 여부를 긴급 확인 중이며 현재까지 인명 피해는 없는 것으로 파악하고 있다. 중동 리스크가 고조되는 상황에서 국내 해운 업계에 긴장감이 높아지고 있다.",
    "YTN, 뉴스1"
  ],
  [
    "어린이날 연휴 보령 해안서 파도에 휩쓸리는 사고 — 소방 구조 출동",
    "어린이날 황금연휴를 맞아 나들이객이 급증한 가운데 충남 보령 해안에서 파도에 휩쓸리는 사고가 발생했다. 소방당국이 현장에 출동해 구조 작업을 벌였으며 연휴 기간 해안가 안전사고에 대한 각별한 주의가 요구된다.",
    "KBS뉴스"
  ],
  [
    "한국인 의사 태국 호텔서 무허가 성형 상담 혐의로 현지 체포",
    "한국인 의사가 태국 방콕 호텔에서 무허가 성형 상담을 진행한 혐의로 현지 당국에 체포됐다. 해외 원정 의료 시장에 대한 각국의 단속이 강화되는 추세 속에 발생한 사건으로, 의료 관광의 법적 경계에 대한 주의가 촉구된다.",
    "조선일보"
  ],
  [
    "울릉도 오징어 한 마리 17만 원 — '온라인 가격과 너무 달라' 관광객 불만",
    "어린이날 울릉도를 찾은 관광객들이 오징어 한 마리 가격이 17만 원에 달한다며 불만을 쏟아냈다. 온라인 예약 시 확인한 가격과 현장 가격 간의 큰 차이가 문제로 지적됐으며, 지역 상인과 관광객 간 갈등이 불거지고 있다.",
    "중앙일보"
  ],
  [
    "어린이날 황금연휴 전국 나들이 인파 폭증 — 놀이공원·관광지 초만원",
    "5일 어린이날과 6일 대체공휴일로 이어지는 황금연휴를 맞아 전국 주요 관광지와 놀이공원에 나들이 인파가 몰렸다. 고속도로 정체가 이어지는 가운데 가족 단위 나들이 수요가 폭발적으로 증가했다.",
    "뉴스1"
  ]
];

// ─── 경제 뉴스 ───────────────────────────────────────────────────────────────
const econNews = [
  [
    "코스피 6,900선 돌파 사상 최고치 — 5% 급등, 반도체 초호황이 이끌었다",
    "5월 4일 코스피 지수가 5% 급등하며 사상 처음으로 6,900선을 돌파했다. AI 수요 확대에 따른 반도체 업황 회복이 주요 동인으로 분석되며, 삼성전자·SK하이닉스 등 반도체 대형주가 강세를 주도했다.",
    "오늘의클릭, 한국경제"
  ],
  [
    "삼성전자 총파업 임박 — 성과급 반도체 편중 불만 폭발",
    "삼성전자 노동조합이 총파업을 예고한 가운데, 핵심 갈등 요인은 성과급 산정 방식이 반도체 부문에만 유리하다는 내부 반발이다. 전 세계 공급망에 미칠 파급 효과를 우려한 시선이 쏠리고 있다.",
    "21세기 이슈"
  ],
  [
    "농식품부, 라면·식용유 등 9대 가공식품 가격 동결 요청",
    "농림축산식품부가 라면, 식용유, 설탕 등 9대 핵심 가공식품의 가격 동결을 식품업계에 공식 요청했다. 중동발 유가 불안과 원달러 환율 급등이 맞물려 물가 상승 압박이 이어지는 상황에서 나온 선제적 조치다.",
    "오늘의클릭"
  ],
  [
    "홈플러스 회생 계획안 가결 기한 7월로 연장 — 청산 우려 여전",
    "기업 회생 절차를 밟고 있는 홈플러스의 회생 계획안 가결 기한이 7월로 연장됐다. 채권단과 인수 후보자 간 합의가 지연되면서 청산 가능성에 대한 우려가 여전히 가시지 않고 있다.",
    "오늘의클릭"
  ],
  [
    "1분기 임대차 거래 월세 비중 68.6% — 전세에서 월세 전환 가속화",
    "2026년 1분기 임대차 거래에서 월세 비중이 68.6%까지 치솟아 역대 최고 수준을 기록했다. 전세 사기 여파와 금리 부담으로 전세에서 월세로의 전환이 빠르게 진행되며, 서민 주거비 부담이 가중되고 있다.",
    "한국경제"
  ]
];

// ─── 정치 뉴스 ───────────────────────────────────────────────────────────────
const politicsNews = [
  [
    "6.3 재보궐 D-30 — 한동훈·조국·송영길 정치생명 건 총력전",
    "6월 3일 재보궐선거가 한 달 앞으로 다가온 가운데 한동훈 전 국민의힘 대표(부산 북갑), 조국 전 조국혁신당 대표, 송영길 전 민주당 대표(인천 연수갑) 등 거물급 정치인들이 정치적 승부수를 던졌다. 역대 두 번째 규모인 14석이 걸린 '미니 총선'이다.",
    "머니투데이"
  ],
  [
    "정진석 전 비서실장, 6.3 재보궐 출마 선언 — 하남갑 전략공천",
    "윤석열 전 대통령의 핵심 측근인 정진석 전 대통령비서실장이 6.3 재보궐선거 출마 의사를 공식화했다. 하남갑 국민의힘 전략공천 후보로, GTX-D 착공과 지하철 3·9호선 연장 등 지역 현안 해결을 공약으로 내걸었다.",
    "경향신문"
  ],
  [
    "6.3 재보궐 14곳 확정 — 청와대 출신 대거 출마 '미니 총선' 격화",
    "6.3 재보궐선거 실시 지역이 총 14곳으로 최종 확정됐다. 이 중 12곳은 22대 총선에서 더불어민주당이 승리한 지역이며, 청와대 참모 출신들이 대거 도전장을 내밀며 선거전이 한층 달아오르고 있다.",
    "경향신문"
  ],
  [
    "이재명 정부 허니문 선거 — 여당 지지율 고공행진 속 6.3 판세 분석",
    "이재명 정부 출범 이후 정부여당 지지율이 고공행진을 이어가고 있다. 다수 여론조사에서 여당 우세 판세가 확인됐으나 야권의 인물론 변수와 지역 민심이 변수로 꼽힌다. 6월 3일 결과가 이재명 정부 향후 국정 운영의 시험대가 될 전망이다.",
    "21세기 이슈"
  ],
  [
    "정부, 호르무즈 한국 선박 피격 첩보 확인 — 외교·안보 대응 가동",
    "호르무즈 해협 한국 선박 피격 첩보 입수와 관련해 정부가 진상 파악에 나섰다. 외교부와 국방부는 현지 상황을 긴급 모니터링하며 관련국과의 외교 채널을 통해 대응 방안을 협의 중이다.",
    "SPN 서울평양뉴스"
  ]
];

// ─── 문서 구성 ───────────────────────────────────────────────────────────────
const children = [
  new Paragraph({
    children: [new TextRun({ text: "일일 뉴스 브리핑", font: "맑은 고딕", size: 56, bold: true, color: BLUE })],
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 120 }
  }),
  new Paragraph({
    children: [new TextRun({ text: "2026년 5월 5일 (화요일) · 어린이날", font: "맑은 고딕", size: 28, color: GRAY })],
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 80 }
  }),
  new Paragraph({
    children: [new TextRun({ text: "AI 뉴스 5 | 사회면 탑 5 | 경제 뉴스 5 | 정치 뉴스 5", font: "맑은 고딕", size: 22, color: "7F7F7F", italics: true })],
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 400 },
    border: { bottom: { style: BorderStyle.DOUBLE, size: 6, color: BLUE, space: 8 } }
  }),

  sectionHeader("🤖  AI · 인공지능 뉴스  TOP 5", DARK_BLUE),
  ...aiNews.flatMap((item, i) => newsItem(i + 1, item[0], item[1], item[2])),
  divider(),

  sectionHeader("📰  사회면  TOP 5", GREEN),
  ...socialNews.flatMap((item, i) => newsItem(i + 1, item[0], item[1], item[2])),
  divider(),

  sectionHeader("💰  경제 뉴스  TOP 5", ORANGE),
  ...econNews.flatMap((item, i) => newsItem(i + 1, item[0], item[1], item[2])),
  divider(),

  sectionHeader("🏛  정치 뉴스  TOP 5", PURPLE),
  ...politicsNews.flatMap((item, i) => newsItem(i + 1, item[0], item[1], item[2])),
  divider(),

  new Paragraph({
    children: [new TextRun({ text: "본 브리핑은 2026년 5월 5일 기준으로 AI가 자동 수집·작성하였습니다.", font: "맑은 고딕", size: 18, color: "7F7F7F", italics: true })],
    alignment: AlignmentType.CENTER,
    spacing: { before: 200, after: 40 }
  }),
  new Paragraph({
    children: [new TextRun({ text: "출처: AI타임스, TechCrunch, CNBC, 오늘의클릭, YTN, 경향신문, 머니투데이, 한국경제 등", font: "맑은 고딕", size: 18, color: "7F7F7F", italics: true })],
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 0 }
  }),
];

const doc = new Document({
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 }
      }
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          children: [
            new TextRun({ text: "일일 뉴스 브리핑", font: "맑은 고딕", size: 18, color: "7F7F7F" }),
            new TextRun({ text: "\t2026.05.05", font: "맑은 고딕", size: 18, color: "7F7F7F" })
          ],
          tabStops: [{ type: "right", position: 9360 }],
          border: { bottom: { style: BorderStyle.SINGLE, size: 3, color: "CCCCCC", space: 4 } }
        })]
      })
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          children: [
            new TextRun({ text: "자동 생성 뉴스 브리핑   |   ", font: "맑은 고딕", size: 16, color: "AAAAAA" }),
            new TextRun({ children: [PageNumber.CURRENT], font: "맑은 고딕", size: 16, color: "AAAAAA" }),
            new TextRun({ text: " / ", font: "맑은 고딕", size: 16, color: "AAAAAA" }),
            new TextRun({ children: [PageNumber.TOTAL_PAGES], font: "맑은 고딕", size: 16, color: "AAAAAA" })
          ],
          alignment: AlignmentType.CENTER,
          border: { top: { style: BorderStyle.SINGLE, size: 3, color: "CCCCCC", space: 4 } }
        })]
      })
    },
    children
  }]
});

const outPath = "D:\\AI 실습\\클로드 코드\\news\\뉴스브리핑 모음\\daily-news-2026-05-05.docx";
Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync(outPath, buffer);
  console.log("Created:", outPath);
}).catch(err => {
  console.error("Error:", err);
  process.exit(1);
});
