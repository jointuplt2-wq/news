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
const ORANGE = "C55A11";

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

const aiNews = [
  [
    "이재명 대통령, 구글 딥마인드 하사비스 CEO 접견... \"AGI 2030년 가시화\"",
    "이재명 대통령이 '알파고의 아버지' 데미스 하사비스 구글 딥마인드 CEO를 접견했다. 하사비스 대표는 5년 안에 AGI(범용 인공지능)가 가시화될 것이라며 악용을 막을 안전장치의 필요성을 강조했다. 양측은 AI 공동 연구와 인재 양성을 위한 양해각서를 체결하고, 구글 AI 캠퍼스를 한국에 설립하기로 합의했다.",
    "MBC 뉴스데스크 (imnews.imbc.com)"
  ],
  [
    "OpenAI, GPT-5.5 공개... 앱 대체하는 AI 에이전트 폰 개발 중",
    "OpenAI가 GPT-5.5를 발표하며 기존의 수동형 언어모델에서 복잡한 작업을 스스로 수행하는 에이전트 기반 시스템으로 전환을 선언했다. 또한 AI 에이전트가 기존 앱을 대체하는 스마트폰을 개발 중인 것으로 알려져 모바일 시장에 큰 파장을 예고하고 있다.",
    "LLM Stats / AI and News (llm-stats.com)"
  ],
  [
    "NVIDIA, 차세대 AI 플랫폼 'Vera Rubin' 공개... 처리 성능 대폭 향상",
    "NVIDIA가 차세대 플래그십 AI 플랫폼 코드명 'Vera Rubin'을 공개하며 처리 성능과 메모리 용량에서 획기적인 향상을 이뤘다. Cadence Design Systems와의 파트너십 확대를 통해 로봇공학 분야의 시뮬레이션-실제 격차(sim-to-real gap)를 해소하는 기술도 선보였다.",
    "Crescendo AI News (crescendo.ai)"
  ],
  [
    "딥시크, AI 모델 가격 대폭 인하... 글로벌 가격 경쟁 재점화",
    "중국 AI 스타트업 딥시크(DeepSeek)가 자사 AI 모델의 가격을 대폭 낮추면서 OpenAI·구글 등 글로벌 AI 기업들의 가격 정책에 연쇄적인 영향을 미치고 있다. 업계에서는 AI 서비스의 범용화가 가속화될 것으로 전망하고 있다.",
    "Investing.com Korea (kr.investing.com)"
  ],
  [
    "중기부, 2026년 AI 예산 약 8천억 원 투입... 22개 사업 지원",
    "중소벤처기업부가 2026년 AI 관련 예산으로 약 7,992억 원을 배정하고, 비R&D 17개·R&D 5개 등 총 22개 사업에 지원에 나선다. 국내 AI 스타트업 및 중소기업의 AI 도입과 기술 경쟁력 강화를 집중 지원할 방침이다.",
    "대한민국 정책브리핑 (korea.kr)"
  ]
];

const socNews = [
  [
    "서울 성수동 포켓몬 행사에 16만 명 운집... 경찰·소방 긴급 출동",
    "노동절 연휴인 지난 1일 서울 성수동에서 열린 '포켓몬스터' 팝업 행사에 약 16만 명이 한꺼번에 몰려 경찰과 소방이 긴급 출동하고 행사가 중단되는 사태가 발생했다. 당국은 인파 관리 미흡을 지적하며 향후 대규모 행사 안전 기준을 강화할 방침이다.",
    "뉴스핌 (newspim.com)"
  ],
  [
    "삼성가 이건희 상속세 12조 원 완납... 역대 최대 규모 납세 마무리",
    "삼성 오너 일가가 고 이건희 회장의 유산 26조 원에 대한 상속세 약 12조 원을 전액 납부하며 역대 최대 규모 납세를 완료했다. 이로써 법적·재산 정리 절차가 대부분 마무리된 것으로 알려졌다.",
    "오늘의 클릭 (cliktoday.com)"
  ],
  [
    "2026년 첫 공식 노동절... '근로'에서 '노동'으로 용어 역사적 전환",
    "2026년 5월 1일, 대한민국이 역사상 첫 공식 노동절을 맞았다. 기존 '근로자의 날'이 '노동절'로 명칭이 바뀌며 노동 권익 존중의 사회적 분위기가 확산되고 있다. 노동계는 용어 변경을 환영하면서도 실질적인 처우 개선이 뒤따라야 한다고 강조했다.",
    "오늘의 클릭 (cliktoday.com)"
  ],
  [
    "5월 4일 임시공휴일 지정 불발... 황금연휴 기대 무산",
    "정부가 어린이날(5일)과 이어지는 5월 4일을 임시공휴일로 지정하지 않기로 결정하면서 직장인들의 황금연휴 기대가 무산됐다. 온라인에서는 아쉬움의 목소리가 이어졌으며, 일각에서는 탄력적 근로 문화 정착이 선행돼야 한다는 의견도 나왔다.",
    "톱스타뉴스 (topstarnews.net)"
  ],
  [
    "쿠팡 개인정보 대규모 유출 사건... 수백만 명 피해 우려",
    "국내 최대 이커머스 플랫폼 쿠팡에서 대규모 개인정보 유출 사건이 발생해 수백만 명의 회원 정보가 외부로 노출된 것으로 확인됐다. 개인정보보호위원회가 즉각 조사에 착수했으며, 피해 규모와 원인 규명에 나서고 있다.",
    "나무위키 / 개인정보보호위원회 (namu.wiki)"
  ]
];

const ecoNews = [
  [
    "코스피, 6,700선 돌파 사상 최고치 경신",
    "코스피 지수가 6,700선을 돌파하며 사상 최고치를 새로 썼다. AI 반도체 관련 종목과 수출 대형주가 상승을 이끌었다는 분석이다. 다만 중동 리스크와 삼성전자 파업 우려가 변수로 남아 있어 단기 변동성에 대한 주의가 필요하다는 전문가 의견이 나오고 있다.",
    "오늘의 클릭 (cliktoday.com)"
  ],
  [
    "삼성전자 노조 총파업 선언... 5월 21일~6월 7일, 30조 손실 우려",
    "삼성전자 전국삼성전자노동조합(전삼노)이 5월 21일부터 6월 7일까지 18일간 총파업을 선언했다. 조합원 약 3만 7천 명이 참여 예정이며, 업계에서는 파업이 현실화될 경우 최소 20조~30조 원 규모의 손실이 발생할 수 있다고 분석했다. 반도체·디스플레이 글로벌 공급망에도 타격이 우려된다.",
    "머니투데이 / 뉴스웨이 (mt.co.kr)"
  ],
  [
    "미·이란 협상 결렬로 국제 유가 급등... 브렌트유 배럴당 110달러 육박",
    "미국과 이란 간 핵 협상이 양측의 입장 차이로 결렬되며 국제 유가가 급등하고 있다. 브렌트유는 배럴당 110달러에 육박하며 원자재 가격 전반이 상승 압력을 받고 있다. 국내 기업들의 에너지 비용 증가와 물가 상승에 대한 우려도 커지고 있다.",
    "21세기 이슈 (issue.cyberbabarian.com)"
  ],
  [
    "삼성전자 1분기 역대급 실적... 매출 133조·영업이익 57조",
    "삼성전자가 2026년 1분기 매출 133조 원, 영업이익 57.2조 원을 달성하며 창사 이래 최대 실적을 기록했다. AI 반도체 수요 급증에 따른 HBM(고대역폭 메모리) 판매 호조가 주요 원인으로 분석된다. 그러나 노조 파업 리스크로 인해 2분기 전망은 불확실하다는 평가다.",
    "오늘의 클릭 (cliktoday.com)"
  ],
  [
    "4월 수출입 동향 발표... 수출 증가세 지속, 대미 무역 불확실성 주목",
    "산업통상자원부가 2026년 4월 수출입 동향을 발표했다. 반도체·이차전지 등 주력 품목 수출이 증가세를 이어가고 있으나, 미국의 관세 정책 변화에 따른 대미 수출 불확실성이 하반기 리스크 요인으로 꼽히고 있다.",
    "보험AI뉴스 다자비 (dazabi.com)"
  ]
];

const polNews = [
  [
    "6·3 지방선거 D-30... 국민의힘 경기지사 양향자 확정, 여야 대진표 완성",
    "6월 3일 전국동시지방선거가 30일 앞으로 다가온 가운데, 국민의힘이 양향자 전 의원을 경기도지사 후보로 확정하며 전국 광역단체장 선거의 여야 대진표가 마무리됐다. 이재명 대통령의 높은 국정 지지율 속에 여권이 우세하나, 부산·울산·경남 등 일부 지역에서는 격차가 좁혀지고 있다.",
    "뉴스핌 (newspim.com)"
  ],
  [
    "국회의원 재보궐 선거 D-30... 전국 14곳 '미니 총선' 대진표 확정",
    "6월 3일 지방선거와 동시에 치러지는 국회의원 재보궐 선거의 대진표가 전국 14개 지역구에서 확정됐다. '미니 총선'으로 불리는 이번 재보궐에서 여야 모두 의석 확보에 총력을 기울이고 있으며, 지방선거 결과와 함께 하반기 정국의 향배를 가를 변수로 주목받고 있다.",
    "나무위키 / 경향신문 (khan.co.kr)"
  ],
  [
    "이재명 대통령, 하사비스와 AI 정책 협력 강화... AI 가드레일 필요성 공감",
    "이재명 대통령이 데미스 하사비스 구글 딥마인드 CEO와 면담에서 AI 시대의 기본소득 필요성과 AI 안전 규범 마련을 논의했다. 과학기술정보통신부와 딥마인드는 AI 공동연구·인재양성 MOU를 체결하고 한국 내 구글 AI 캠퍼스 설립을 추진하기로 합의했다.",
    "뉴시스 / 파이낸셜뉴스 (fnnews.com)"
  ],
  [
    "AI 민관 정책협의회 출범... 정부, AI 규범 체계 구축 본격화",
    "정부가 '2026 인공지능(AI) 프라이버시 민관 정책협의회'를 공식 출범시키며 AI 시대 개인정보 보호와 규범 정립에 나섰다. 민간 전문가·기업·정부가 함께 참여하는 이 협의회는 AI 활용 윤리 기준과 개인정보 보호 정책을 정기적으로 논의할 예정이다.",
    "대한민국 정책브리핑 (korea.kr)"
  ],
  [
    "한국정치학회, '북한 vs 조선' 명칭 학술회의 개최",
    "한국정치학회가 서울 중구 한국프레스센터에서 '평화 공존을 위한 이름 부르기: 북한인가 조선인가'를 주제로 특별 학술회의를 열었다. 남북관계 개선을 위한 명칭 사용의 상징성과 정치적 함의를 두고 학계의 다양한 의견이 쏟아졌다.",
    "SPN 서울평양뉴스 (spnews.co.kr)"
  ]
];

function buildSection(items, color) {
  const children = [];
  items.forEach(([title, body, source], i) => {
    newsItem(i + 1, title, body, source).forEach(p => children.push(p));
  });
  return children;
}

const today = "2026년 5월 4일 (월)";

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
        children: [
          new Paragraph({
            alignment: AlignmentType.RIGHT,
            border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: DARK_BLUE, space: 4 } },
            spacing: { after: 80 },
            children: [
              new TextRun({ text: `일일 뉴스 브리핑  |  ${today}`, size: 18, color: GRAY, font: "맑은 고딕" })
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
            border: { top: { style: BorderStyle.SINGLE, size: 4, color: DARK_BLUE, space: 4 } },
            spacing: { before: 80 },
            children: [
              new TextRun({ text: "- ", size: 18, color: GRAY, font: "맑은 고딕" }),
              new TextRun({ children: [PageNumber.CURRENT], size: 18, color: GRAY, font: "맑은 고딕" }),
              new TextRun({ text: " -", size: 18, color: GRAY, font: "맑은 고딕" })
            ]
          })
        ]
      })
    },
    children: [
      // Title
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 240, after: 80 },
        children: [
          new TextRun({ text: "일일 뉴스 브리핑", bold: true, size: 52, color: BLUE, font: "맑은 고딕" })
        ]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 360 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: DARK_BLUE, space: 6 } },
        children: [
          new TextRun({ text: today, size: 28, color: GRAY, font: "맑은 고딕" })
        ]
      }),

      // AI 뉴스
      sectionHeader("🤖  AI · 인공지능 뉴스", BLUE),
      ...buildSection(aiNews, BLUE),

      // 사회 뉴스
      sectionHeader("📰  사회면 주요 뉴스", DARK_GREEN),
      ...buildSection(socNews, DARK_GREEN),

      // 경제 뉴스
      sectionHeader("💹  경제 뉴스", ORANGE),
      ...buildSection(ecoNews, ORANGE),

      // 정치 뉴스
      sectionHeader("🏛️  정치 뉴스", RED),
      ...buildSection(polNews, RED),

      // Footer note
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 480, after: 0 },
        border: { top: { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC", space: 4 } },
        children: [
          new TextRun({ text: "본 브리핑은 AI가 자동 수집·요약한 뉴스입니다. 정확한 내용은 원문 출처를 확인하시기 바랍니다.", size: 16, color: "AAAAAA", italics: true, font: "맑은 고딕" })
        ]
      }),
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  const filename = "뉴스브리핑 모음/일일뉴스브리핑_20260504.docx";
  fs.writeFileSync(filename, buffer);
  console.log(`생성 완료: ${filename}`);
});
