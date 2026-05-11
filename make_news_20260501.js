const {
  Document, Packer, Paragraph, TextRun,
  AlignmentType, BorderStyle, Header, Footer, PageNumber,
} = require('docx');
const fs = require('fs');

const BLUE = "1F4E79";
const DARK_BLUE = "2E75B6";
const GRAY = "595959";
const DARK_GREEN = "375623";
const RED = "7B2C2C";
const PURPLE = "4A235A";

function sectionHeader(text, color) {
  return new Paragraph({
    spacing: { before: 360, after: 120 },
    border: {
      bottom: { style: BorderStyle.SINGLE, size: 6, color: color || DARK_BLUE, space: 4 }
    },
    children: [
      new TextRun({ text, bold: true, size: 32, color: color || BLUE, font: "맑은 고딕" })
    ]
  });
}

function newsItem(num, title, body, source) {
  return [
    new Paragraph({
      spacing: { before: 200, after: 60 },
      children: [
        new TextRun({ text: `${num}. ${title}`, bold: true, size: 24, color: "000000", font: "맑은 고딕" })
      ]
    }),
    new Paragraph({
      spacing: { before: 0, after: 60 },
      indent: { left: 360 },
      children: [
        new TextRun({ text: body, size: 20, color: GRAY, font: "맑은 고딕" })
      ]
    }),
    new Paragraph({
      spacing: { before: 0, after: 120 },
      indent: { left: 360 },
      children: [
        new TextRun({ text: `[출처: ${source}]`, size: 18, color: "888888", italics: true, font: "맑은 고딕" })
      ]
    }),
  ];
}

const sections_data = [
  {
    title: "1부. AI & 인공지능 뉴스",
    color: BLUE,
    items: [
      [
        "Anthropic, 글로벌 AI 모델 성능 랭킹 1위 — xAI·Google·OpenAI 제쳐",
        "2026년 3월 기준 Anthropic이 글로벌 AI 모델 성능 랭킹에서 1위를 기록했다. 커뮤니티 기반 평가 플랫폼 'Arena'에서 Anthropic이 xAI·Google·OpenAI를 뒤로 제쳤으며, 미국과 중국이 AI 모델 성능 면에서 거의 대등한 수준에 올랐다는 평가도 나왔다. 프런티어 모델은 이제 박사급 과학 문제, 수학 경시, 멀티모달 추론 분야에서 인간과 동등하거나 이를 능가하는 성능을 보이고 있다.",
        "MIT Technology Review / Stanford HAI 2026 AI Index Report (2026.05.01)"
      ],
      [
        "생성형 AI 대중 보급률 53% 돌파 — PC·인터넷보다 빠른 역대 최고 확산 속도",
        "생성형 AI가 출시 3년 만에 미국 인구의 53%에게 보급되며 역사상 가장 빠른 기술 확산 속도를 기록했다. 미국 소비자 대상 생성형 AI 도구의 연간 경제적 가치는 2026년 초 1,720억 달러(약 237조 원)에 달하며, 사용자 1인당 체감 가치는 2025년 대비 3배로 증가했다.",
        "Stanford HAI 2026 AI Index Report (2026.05.01)"
      ],
      [
        "AI 에이전트 실전 성공률 77.3%로 급등 — 사이버보안 분야 93% 달성",
        "2025년 20%에 불과했던 AI 에이전트의 실세계 과제 처리 성공률이 2026년 77.3%로 대폭 상승했다. 사이버보안 분야에서는 AI가 문제 해결률 93%를 기록해 2024년 15% 대비 6배 이상 개선됐다. 전문가들은 복잡한 목표를 협업으로 달성하는 '에이전트 팀' 시대가 임박했다고 전망한다.",
        "MIT Technology Review 2026 AI Trends (2026.05.01)"
      ],
      [
        "AI 데이터센터 전력 29.6 GW — 뉴욕주 전체 전력과 맞먹는 수준",
        "전 세계 AI 데이터센터 전력 용량이 29.6 GW에 달해 미국 뉴욕주 최대 수요량과 맞먹는 수준에 이르렀다. GPT-4o 추론 과정에서만 연간 소비되는 냉각수가 1,200만 명의 연간 음용수 수요를 초과할 수 있다는 환경 우려도 제기됐다. AI 인프라 확장에 따른 에너지·수자원 문제가 글로벌 과제로 부상하고 있다.",
        "Stanford HAI 2026 AI Index Report (2026.05.01)"
      ],
      [
        "한국 중기부, AI 예산 약 8,000억 원 투입 — 스타트업·제조 AX 집중 지원",
        "2026년 중소벤처기업부는 AI 관련 예산 약 7,992억 원을 편성했다. 이 중 6,718억 원은 비R&D 사업에 투입되며, AI·딥테크 투자 펀드 조성, 제조 현장 AI 전환(AX), AI 스타트업 육성 등에 활용된다. 국내 '인공지능 윤리 및 안전 기본법'이 2026년 1월부터 시행 중이며, 구글 딥마인드와 과기정통부의 AI 협력 파트너십도 활성화되고 있다.",
        "대한민국 정책브리핑 / korea.kr (2026.05.01)"
      ],
    ]
  },
  {
    title: "2부. 사회면 탑기사",
    color: RED,
    items: [
      [
        "노동절 63년 만에 법정공휴일 지정 — 오늘(5/1)부터 공식 적용",
        "2026년부터 5월 1일 '노동절'(구 근로자의 날)이 63년 만에 법정공휴일로 지정됐다. 명칭도 '근로자의 날'에서 '노동절'로 공식 변경됐다. 올해 노동절은 금요일로, 5월 2~3일 주말과 합쳐 3일 연속 휴일이 이어진다. 공무원·교사를 포함한 모든 근로자가 노동절 혜택을 받게 됐다.",
        "톱스타뉴스 (2026.05.01)"
      ],
      [
        "민주노총·한국노총, 서울 도심 대규모 노동절 집회 — 교통 혼잡 예상",
        "5월 1일 오후 광화문광장 일대에서 민주노총 주최 '2026 세계노동절 대회'가, 여의도 일대에서는 한국노총의 전국 노동자 대회가 동시에 열린다. 경찰은 집회 안전 관리를 위해 대규모 인력을 배치했으며, 서울 도심 주요 구간에서 교통 통제가 예상된다.",
        "YTN / MBC 뉴스 (2026.05.01)"
      ],
      [
        "이주노동자 메이데이 집회 — '강제노동·임금체납 철폐' 행진",
        "노동절을 앞둔 4월 26일, 이주노동자조합·이주노동자평등연대 등이 서울 서울고용노동청 앞에서 '2026 이주노동자 메이데이 집회'를 열었다. 참가자들은 강제노동 철폐, 임금체납 근절, 사업장 변경의 자유 보장 등을 요구하며 청와대까지 행진했다.",
        "경향신문 / MBC 뉴스 (2026.05.01)"
      ],
      [
        "5월 4일 임시공휴일 결국 무산 — 대통령실 '검토 없다' 공식 입장",
        "일부에서 기대했던 5월 4일(월) 징검다리 임시공휴일 지정이 무산됐다. 대통령실은 '5월 4일 임시공휴일 지정을 검토한 바 없다'고 공식 입장을 밝혔다. 이에 따라 완전한 황금연휴 기대는 사라졌으며, 다음 공휴일은 5월 5일 어린이날이다.",
        "톱스타뉴스 (2026.05.01)"
      ],
      [
        "2026년 5월 공휴일 총정리 — 노동절·어린이날·부처님오신날 대체휴일",
        "5월 공휴일은 ▲1일 노동절(금·법정공휴일) ▲5일 어린이날(화) ▲24일 부처님 오신 날이 일요일이라 25일(월)이 대체휴일로 지정됐다. 5월 4일 임시공휴일 지정이 무산돼 연속 황금연휴는 성사되지 않았지만, 5월 한 달간 총 5일의 공휴일이 분산 배치됐다.",
        "서울시 / 톱스타뉴스 (2026.05.01)"
      ],
    ]
  },
  {
    title: "3부. 경제 뉴스",
    color: DARK_GREEN,
    items: [
      [
        "종합소득세 신고 기간 오늘(5/1) 시작 — 마감 6월 1일, 미신고 시 가산세",
        "2025년 귀속 종합소득세 신고·납부 기간이 5월 1일부터 시작됐다. 올해는 5월 31일이 일요일이어서 마감이 6월 1일(월)로 하루 연장된다. 기한 내 미신고 시 무신고 가산세(20%) 및 납부지연 가산세가 부과된다. 국세청은 홈택스 전자신고를 권장하고 있다.",
        "모두의 소식 / unips.co.kr (2026.05.01)"
      ],
      [
        "한국 1분기 GDP 1.7% 성장 — 예상치 2배 '깜짝 성장', 전년비 3.6%↑",
        "한국은행 발표에 따르면 2026년 1분기 실질 GDP가 전기 대비 1.7%, 전년 동기 대비 3.6% 성장했다. 반도체 수출 호조, 건설 투자 회복, 민간소비 개선이 성장을 이끌었으며, 당초 예측치(약 0.9%)를 약 2배 상회하는 깜짝 성장으로 평가받는다.",
        "MBC 뉴스 (2026.05.01)"
      ],
      [
        "JP모건·글로벌 IB, 한국 성장률 전망 3%로 일제히 상향 조정",
        "JP모건이 2026년 한국 경제성장률 전망을 기존 2.2%에서 3.0%로 대폭 올렸다. 골드만삭스, 모건스탠리 등 주요 글로벌 투자은행들도 한국의 반도체 수출 호조와 AI 관련 수요를 근거로 성장률 전망을 일제히 상향 조정하고 있다.",
        "MBC 뉴스 / 한국경제 (2026.05.01)"
      ],
      [
        "PwC 보고서: AI 경제 이익의 75%, 상위 20% 기업이 독식",
        "PwC가 발표한 '2026 AI 성과 연구(AI Performance Study)'에 따르면 AI의 경제적 이익 중 75%가 전체 기업의 20%에 집중되고 있다. 선도 기업들은 단순 생산성 향상을 넘어 매출 성장에 AI를 집중 활용하는 전략을 취하고 있으며, AI 격차 심화가 기업 경쟁력 양극화를 가속화하고 있다.",
        "PwC 글로벌 보도자료 (2026.05.01)"
      ],
      [
        "모건스탠리 경고: 2026년 AI 대전환 임박 — 세계 대부분 준비 안 돼",
        "모건스탠리는 최신 보고서에서 2026년 AI 기술의 획기적 발전이 임박했지만 전 세계 기업·정부 대부분이 준비 부족 상태라고 경고했다. 특히 중소기업과 신흥 시장의 AI 격차 심화를 우려하며 선제적 대응을 촉구했다. AI 투자 쏠림이 지속될 경우 버블 위험도 병존한다는 분석도 나온다.",
        "Yahoo Finance / 모건스탠리 보고서 (2026.05.01)"
      ],
    ]
  },
  {
    title: "4부. 정치 뉴스",
    color: PURPLE,
    items: [
      [
        "정진석 전 의원, 6·3 충남 공주·부여·청양 재보선 출마 선언",
        "정진석 전 국민의힘 의원이 4월 30일 6·3 국회의원 재보궐선거에 충남 공주·부여·청양 지역구 출마를 공식 선언했다. 더불어민주당은 '윤석열 전 대통령의 옥중 출마와 다름없다'며 강하게 비판했다. 해당 지역은 전통적인 보수 강세 지역으로 공천 후보 간 경쟁이 주목된다.",
        "MBC 뉴스 (2026.05.01)"
      ],
      [
        "6·3 재보선 총 14개 지역구 확정 — 수도권·부산경남 표심이 관건",
        "6월 3일 실시되는 국회의원 재보궐선거의 선거구가 최종 14곳으로 확정됐다. 인천 계양 등 5개 지역구에서 출발했으나 시도지사 후보 출마를 위해 사퇴한 현역 의원 9명의 지역구가 추가됐다. 수도권과 부산·경남 지역의 표심이 선거의 핵심 변수로 꼽힌다.",
        "MBC 뉴스 / 정치 뉴스 종합 (2026.05.01)"
      ],
      [
        "윤석열 전 대통령 2심, 1심보다 무거운 형량 선고",
        "내란 전담 재판부가 처음 내놓은 판결로, 윤석열 피고인의 항소심(2심)에서 체포영장 집행 방해 혐의 등을 포함해 1심보다 무거운 형량이 선고됐다. 전날 김건희 씨의 2심 판결에 이어 연속으로 나온 판결로 사회적 관심이 집중됐으며, 향후 대법원 상고 여부가 주목된다.",
        "MBC 뉴스 정치면 (2026.05.01)"
      ],
      [
        "김건희 씨 2심 선고 — 1심 대비 가중 처벌, 여야 공방 격화",
        "4월 30일 김건희 씨의 항소심(2심) 판결이 내려졌으며 1심보다 무거운 형량이 선고됐다. 정치권에서는 여야 간 격렬한 공방이 이어졌으며, 6·3 재보선을 앞두고 정치적 파장이 예상된다. 대법원 상고 여부에 이목이 쏠린다.",
        "MBC 뉴스 (2026.05.01)"
      ],
      [
        "6·3 재보선 앞두고 여야 공천 경쟁 본격화",
        "6월 3일 재보궐선거를 한 달여 앞두고 국민의힘과 더불어민주당 모두 후보 공천 작업에 속도를 내고 있다. 각 당은 수도권·경쟁 지역구에 전략 공천 카드를 검토 중이며, 선거 결과가 향후 당내 권력 구도에도 영향을 미칠 것으로 전망된다. 노동절 집회 현장도 각 당의 정치 메시지 경쟁 무대가 됐다.",
        "한국경제 정치면 / 다음뉴스 정치 (2026.05.01)"
      ],
    ]
  },
];

const children = [
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 240, after: 120 },
    children: [
      new TextRun({ text: "2026년 5월 1일 (금) 뉴스 브리핑", bold: true, size: 48, color: BLUE, font: "맑은 고딕" })
    ]
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 60 },
    children: [
      new TextRun({ text: "노동절 특집 | AI, 사회, 경제, 정치 주요 뉴스", size: 28, color: GRAY, font: "맑은 고딕" })
    ]
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 360 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: DARK_BLUE, space: 6 } },
    children: [
      new TextRun({ text: "Claude Code 자동화 스케줄러 생성", size: 20, color: "AAAAAA", italics: true, font: "맑은 고딕" })
    ]
  })
];

sections_data.forEach((section, sIdx) => {
  if (sIdx > 0) children.push(new Paragraph({ children: [], spacing: { before: 200, after: 0 } }));
  children.push(sectionHeader(section.title, section.color));
  section.items.forEach((item, i) => children.push(...newsItem(i + 1, item[0], item[1], item[2])));
});

children.push(
  new Paragraph({
    spacing: { before: 480, after: 120 },
    alignment: AlignmentType.CENTER,
    border: {
      top: { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC", space: 4 }
    },
    children: [
      new TextRun({ text: "본 브리핑은 Claude Code 자동화 스케줄러가 2026년 5월 1일 수집·생성한 일일 뉴스 요약입니다.", size: 18, color: "AAAAAA", font: "맑은 고딕" })
    ]
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 0 },
    children: [
      new TextRun({ text: "각 기사 원문은 해당 언론사 홈페이지에서 확인하시기 바랍니다.", size: 18, color: "AAAAAA", font: "맑은 고딕" })
    ]
  })
);

const doc = new Document({
  styles: {
    default: { document: { run: { font: "맑은 고딕", size: 22 } } }
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
        children: [
          new Paragraph({
            border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC", space: 4 } },
            children: [
              new TextRun({ text: "일일 뉴스 브리핑   |   2026년 5월 1일 (금) 노동절", font: "맑은 고딕", size: 18, color: "888888" }),
            ]
          })
        ]
      })
    },
    footers: {
      default: new Footer({
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            border: { top: { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC", space: 4 } },
            children: [
              new TextRun({ text: "Page ", font: "맑은 고딕", size: 18, color: "888888" }),
              new TextRun({ children: [PageNumber.CURRENT], font: "맑은 고딕", size: 18, color: "888888" }),
            ]
          })
        ]
      })
    },
    children,
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("일일뉴스브리핑_20260501.docx", buffer);
  console.log("성공: 일일뉴스브리핑_20260501.docx 생성완료");
}).catch(err => {
  console.error("오류:", err.message);
  process.exit(1);
});
