const {
  Document, Packer, Paragraph, TextRun,
  AlignmentType, BorderStyle, Header, Footer, PageNumber,
} = require('docx');
const fs = require('fs');

const BLUE      = "1F4E79";
const DARK_BLUE = "2E75B6";
const GRAY      = "595959";
const GREEN     = "1A6B3A";
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
    "구글 딥마인드, 한국에 첫 해외 AI 캠퍼스 공식 설립 발표",
    "구글 딥마인드가 한국에 첫 번째 해외 AI 캠퍼스를 설립한다고 공식 발표했다. 한국의 높은 AI 기술 수준과 우수한 개발자 역량을 세계가 인정한 결과로 평가받고 있으며, K-스타트업과의 협력 지원 프로그램도 함께 운영될 예정이다.",
    "AI타임스, 전자신문"
  ],
  [
    "AI 반도체 HBM4 경쟁 심화 — 삼성·SK하이닉스 양산 시점 앞당겨",
    "삼성전자와 SK하이닉스가 차세대 고대역폭메모리(HBM4) 시장 선점을 위해 양산 시점을 앞다투며 경쟁을 벌이고 있다. 삼성전자는 2026년 1분기 매출 133조 원, 영업이익 57.2조 원이라는 역대급 실적을 기록하며 AI 수요 급증의 수혜를 누리고 있다.",
    "한국경제, 전자신문"
  ],
  [
    "챗GPT·제미나이·클로드, AI 챗봇 3강 구도 재편 — 제미나이 방문 수 647% 급증",
    "2026년 봄 AI 챗봇 시장 판도가 크게 바뀌고 있다. 챗GPT 점유율이 87%에서 68%로 하락한 반면, 구글 제미나이는 안드로이드·크롬 기본 탑재 효과로 방문 수가 647% 급증했다. 앤트로픽 클로드는 긴 문서 처리와 코딩 정확성에서 높은 평가를 받으며 기업 시장을 공략 중이다.",
    "디지털투데이, KMJ"
  ],
  [
    "이재명 대통령, 데미스 하사비스 딥마인드 CEO 면담 — \"AI 가드레일 필요\"",
    "이재명 대통령이 구글 딥마인드 최고경영자 데미스 하사비스와 만나 AI 기술의 빠른 발전에 대한 우려를 공유하고 'AI를 통제할 가드레일이 필요하다'고 강조했다. 양측은 한국의 AI 거버넌스 체계 구축에 협력하기로 합의했다.",
    "MBC뉴스"
  ],
  [
    "연세대·고려대 등 7개 'AI 중심대학' 선정 — 전문 인재 양성 가속화",
    "정부가 연세대, 고려대를 포함한 7개 대학을 'AI 중심대학'으로 선정해 AI 전문·융합 인재 양성을 가속화한다고 발표했다. 선정 대학에는 특별 예산이 지원되며, 산학 협력 프로그램과 연계해 실무형 AI 인재를 육성할 계획이다.",
    "AI타임스, 정책브리핑"
  ]
];

// ─── 사회 뉴스 ───────────────────────────────────────────────────────────────
const socialNews = [
  [
    "5세대 실손의료보험 6일부터 판매 — 보험료 절반, 보장도 축소",
    "도수치료·체외충격파 등을 보장에서 제외하고 비중증 질환의 자기부담금을 높인 '5세대 실손의료보험'이 6일부터 판매된다. 보험료는 기존 대비 최대 절반 이상 낮아지지만 보장 범위도 함께 축소돼 소비자의 꼼꼼한 비교가 필요하다.",
    "보험업계, 경향신문"
  ],
  [
    "비대면 진료 플랫폼, 약국별 조제 가능 여부 정보 오늘부터 제공",
    "보건복지부와 건강보험심사평가원이 6일부터 비대면 진료 중개 플랫폼에 약국별 처방 의약품 구매·조제 가능 여부 정보를 제공한다. 환자들이 처방전 수령 후 조제 가능한 약국을 사전에 확인할 수 있게 돼 편의성이 크게 높아질 전망이다.",
    "복지부, 뉴스1"
  ],
  [
    "호르무즈 해협 한국 선박 폭발 — 미국, 이란 공격 규정하며 군사 합류 압박",
    "호르무즈 해협에서 한국 선사 선박이 폭발하는 사고가 발생했다. 미국 대통령은 이를 이란의 공격으로 규정하고 한국에 군사 작전 합류를 촉구하고 있어 외교적 파장이 예상된다. 청와대는 대책회의를 열고 전문가를 현지에 급파해 원인 규명에 나섰다.",
    "YTN, 연합뉴스"
  ],
  [
    "'나무호' 화재 실종 선원 수색 이어져 — 해경, 함정·항공기 투입",
    "어린이날인 5월 5일 발생한 선박 나무호 화재로 실종된 선원에 대한 수색 작업이 6일에도 계속됐다. 해양경찰청은 인근 해역에 함정과 항공기를 투입해 수색 범위를 확대하고 있으며, 화재 원인에 대한 조사도 병행 중이다.",
    "KBS뉴스, 해양경찰청"
  ],
  [
    "우크라이나 휴전 협상 소강 국면 — 트럼프 중재 속 양측 입장 차 여전",
    "우크라이나·러시아 전쟁의 휴전 협상이 트럼프 미국 대통령 중재 아래 진행 중이나 양측 입장 차가 좁혀지지 않아 소강 국면을 보이고 있다. 국제사회는 조속한 협상 타결을 촉구하며 외교적 해법 모색에 집중하고 있다.",
    "연합뉴스, 오늘의클릭"
  ]
];

// ─── 경제 뉴스 ───────────────────────────────────────────────────────────────
const econNews = [
  [
    "중동 긴장으로 국제 유가 급등 — 에너지 공급망 불안 가중",
    "호르무즈 해협 사고 여파로 국제 유가가 급등하며 글로벌 에너지 공급망 불안과 인플레이션 압박이 가중되고 있다. 국내 정유업계는 가격 인상 압력에 직면했으며, 수출 기업들의 원가 부담도 증가할 전망이다.",
    "오늘의클릭, 한국경제"
  ],
  [
    "코스피 사상 최고치 경신 — AI·반도체 랠리 주도, 증권가 상단 8,400 전망",
    "코스피 지수가 사상 최고치를 경신했다. AI 관련 종목과 반도체 섹터가 랠리를 주도했으며, 삼성전자·SK하이닉스의 실적 호조가 지수 상승을 이끌었다. 증권가에서는 코스피 상단을 8,400으로 제시하는 낙관적 전망도 나오고 있다.",
    "라이프타임뉴스, 한국경제TV"
  ],
  [
    "삼성전자 1분기 역대 최대 실적 — 매출 133조·영업이익 57.2조",
    "삼성전자가 2026년 1분기에 매출 133조 원, 영업이익 57.2조 원이라는 역대 최대 실적을 달성했다. AI 수요 확대에 따른 HBM 판매 급증과 파운드리 부문 회복이 실적 개선을 견인한 것으로 분석된다.",
    "오늘의클릭, 전자신문"
  ],
  [
    "부동산 양극화 심화 — 서울 상승 둔화, 지방 미분양 속출",
    "고금리 장기화와 대출 규제 강화 영향으로 서울 아파트값 상승폭은 둔화된 반면, 지방에서는 미분양 물량이 쌓이며 하락세가 지속되고 있다. 전문가들은 지역별 양극화가 당분간 이어질 것으로 전망했다.",
    "한국경제, 연합뉴스"
  ],
  [
    "원·달러 환율 1430원대 유지 — 고환율 장기화, 기준금리 연 2.5% 동결",
    "원·달러 환율이 1430원대를 유지하며 고환율 상황이 지속되고 있다. 한국은행 기준금리는 연 2.5%로 동결 상태이며, 시장에서는 하반기 1~2회 추가 금리 인하 가능성을 예상하고 있으나 가계부채 부담이 인하 폭을 제한할 것으로 분석된다.",
    "토스뱅크, 한국은행"
  ]
];

// ─── 정치 뉴스 ───────────────────────────────────────────────────────────────
const politicsNews = [
  [
    "6·3 부산 북구 보궐선거 — 민주 하정우 vs 국힘 박민식 '격전'",
    "6월 3일 지방선거와 함께 실시되는 부산 북구 국회의원 보궐선거에서 더불어민주당 하정우(전 청와대 AI미래기획수석)와 국민의힘 박민식(재선 의원) 후보가 맞붙는다. 선거운동 초반부터 양 진영이 치열한 유세 경쟁을 벌이고 있다.",
    "경향신문, 오늘의클릭"
  ],
  [
    "정청래 민주당 대표, 구포시장 '오빠' 발언 논란 공식 사과",
    "정청래 민주당 대표가 부산 구포시장 방문 중 초등학교 1학년 여아에게 하정우 후보를 '오빠'라고 부르게 한 발언이 논란이 되자 공식 사과했다. 야당과 시민단체는 아동을 선거에 동원한 것은 부적절하다고 비판하고 있다.",
    "경향신문, 연합뉴스"
  ],
  [
    "이재명 대통령 국정수행 긍정 59.5% — 직전 대비 2.7%p 하락",
    "리얼미터 최신 여론조사에서 이재명 대통령의 국정수행 긍정 평가가 59.5%로 직전 조사 대비 2.7%포인트 하락했다. 중동 사태 대응과 경기 불안이 지지율에 영향을 준 것으로 분석된다.",
    "리얼미터, 21세기이슈"
  ],
  [
    "트럼프 '프로젝트 프리덤' — 한국에 이란 대응 군사 작전 합류 요구",
    "도널드 트럼프 미국 대통령이 이란을 겨냥한 '프로젝트 프리덤'에 한국의 군사 협력 참여를 요구하며 압박 수위를 높이고 있다. 외교부는 동맹 의무와 독자적 외교 노선 사이에서 신중한 입장을 유지하고 있다.",
    "오늘의클릭, 연합뉴스"
  ],
  [
    "개혁신당, 6·3 지방선거 본격 참전 — 울산 8명 후보 확정",
    "개혁신당이 6·3 지방선거에 본격적으로 뛰어들며 울산 지역에 8명의 후보를 내기로 확정했다. 여야 양당 구도에 균열을 가져올 변수로 주목받고 있으며, 수도권 및 주요 도시에서도 후보를 추가 발굴 중이다.",
    "21세기이슈, 나무위키"
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
    children: [new TextRun({ text: "2026년 5월 6일 (수요일) · 대체공휴일", font: "맑은 고딕", size: 28, color: GRAY })],
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
    children: [new TextRun({ text: "본 브리핑은 2026년 5월 6일 기준으로 AI가 자동 수집·작성하였습니다.", font: "맑은 고딕", size: 18, color: "7F7F7F", italics: true })],
    alignment: AlignmentType.CENTER,
    spacing: { before: 200, after: 40 }
  }),
  new Paragraph({
    children: [new TextRun({ text: "출처: AI타임스, 전자신문, MBC뉴스, 오늘의클릭, YTN, 경향신문, 한국경제, 연합뉴스 등", font: "맑은 고딕", size: 18, color: "7F7F7F", italics: true })],
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
            new TextRun({ text: "\t2026.05.06", font: "맑은 고딕", size: 18, color: "7F7F7F" })
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

const outPath = "D:\\AI 실습\\클로드 코드\\news\\뉴스브리핑 모음\\일일뉴스브리핑_20260506.docx";
Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync(outPath, buffer);
  console.log("Created:", outPath);
}).catch(err => {
  console.error("Error:", err);
  process.exit(1);
});
