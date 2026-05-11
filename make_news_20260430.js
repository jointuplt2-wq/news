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
        "DeepSeek V4-Pro 오픈소스 공개, GPT-5.4와 Claude Opus 4.6 능가",
        "4월 24일 중국 AI 스타트업 DeepSeek이 V4-Pro와 V4-Flash를 동시에 오픈소스로 공개했습니다. V4-Pro는 주요 코딩 벤치마크에서 GPT-5.4, Claude Opus 4.6, Gemini 3.1-Pro를 제치는 성능을 기록했습니다. 중국이 AI 모델 경쟁에서 글로벌 선두에 도전장을 내밀고 있다는 평가가 나옵니다.",
        "AI 타임스, Investing.com (2026.04.30)"
      ],
      [
        "DeepSeek, AI 모델 API 가격 10분의 1로 대폭 인하, 업계 가격 경쟁 재점화",
        "DeepSeek이 자사 AI 모델 API 제품군의 입력 캐시 히트 가격을 기존 대비 10분의 1 수준으로 즉시 인하했습니다. 이는 OpenAI, Anthropic, Google 등 주요 AI 기업들의 가격 인하 압력으로 작용하며 AI API 시장 전반의 가격 경쟁을 재점화시키고 있습니다. 기업들의 AI 도입 비용 감소 효과로 이어질 전망입니다.",
        "Investing.com, AI 타임스 (2026.04.30)"
      ],
      [
        "테슬라 AI 칩 AI4.1, 삼성전자 파운드리 생산 위탁 발표",
        "일론 머스크가 테슬라 AI 칩 'AI4'의 업그레이드 버전인 'AI4.1' 생산을 삼성전자 파운드리에 맡기겠다고 발표했습니다. 삼성전자 주가는 이 소식에 급등했으며, SK하이닉스 등 국내 AI 반도체 관련주도 강세를 보였습니다. AI 서버 수요 확대에 따른 국내 반도체 수주 기대감이 코스피 상승세를 뒷받침하고 있습니다.",
        "오늘의 클릭, AI 타임스 (2026.04.30)"
      ],
      [
        "과기정통부, 딥마인드와 AI 협력 MOU 체결, 구글 AI 캠퍼스 한국 설립",
        "과학기술정보통신부가 구글 딥마인드와 AI 공동 연구와 AI 인재 양성을 위한 양해각서(MOU)를 체결하고 구글 AI 캠퍼스를 한국에 설립하기로 합의했습니다. 이 대통령은 알파고 아버지로 알려진 데미스 하사비스 딥마인드 CEO와 만나 AI 통제를 위한 가드레일 필요성을 강조했습니다.",
        "MBC뉴스, 대한민국 정책브리핑 (2026.04.30)"
      ],
      [
        "에이전틱 AI 시대 본격화, 중기부 AI 예산 약 8천억 원 투입",
        "2026년 4월 넷째 주 AI 트렌드 조사에 따르면, AI는 스스로 판단하고 업무를 수행하는 에이전틱 AI로 빠르게 진화하고 있습니다. 중소벤처기업부는 2026년 AI 관련 예산으로 약 8,000억 원을 편성하고 AI 서비스의 고도화와 내실화에 집중하는 전략으로 전환했습니다.",
        "정책브리핑, carrotglobalblog.com (2026.04.30)"
      ],
    ]
  },
  {
    title: "2부. 사회면 탑기사",
    color: RED,
    items: [
      [
        "공동주택 공시가격 9.13% 상승 확정, 서울 공시지가 4.9% 올라",
        "2026년 전국 공동주택 공시가격이 전년 대비 9.13% 상승으로 최종 확정되어 4월 30일 공시되었습니다. 서울 개별공시지가도 전년 대비 4.9% 올라 용산구의 상승폭이 가장 컸습니다. 공시가격 상승으로 주택 보유세 부담이 현실화된다는 우려가 학계와 시장에서 높아지고 있습니다.",
        "에콘믹글, 대한민국 정책브리핑 (2026.04.30)"
      ],
      [
        "전삼노, 이재용 의장 자택 인근 천막농성 지속, 면담 요구",
        "전국삼성전자노동조합이 이재용 삼성전자 의장의 서울 용산구 자택 인근에서 천막농성을 계속하며 의장과의 면담을 요구하고 있습니다. 노조는 4월 27일부터 천막을 세우고 임금 교섭과 노사 합의를 요구 중입니다. 삼성전자 측은 노동 관련 사항은 노사 협상을 통해 풀어야 한다는 입장을 고수하고 있습니다.",
        "경향신문, 서울신문 (2026.04.30)"
      ],
      [
        "4년제 대학 등록금 평균 727만원, 130개교 인상, 시각장애 학생 점자 교과서 늦게 지급 문제",
        "2026학년도 4년제 일반 및 교육대학의 연간 1인당 평균 등록금이 727만 300원으로 전년보다 14만 7,100원 올랐습니다. 192개교 중 130개교가 등록금을 인상했습니다. 이와 함께 시각장애 학생에게 점자 교과서가 한 달 이상 늦게 지급되는 학습권 침해 문제도 지속되고 있습니다.",
        "서울신문, 경향신문 (2026.04.30)"
      ],
      [
        "김범석 쿠팡Inc 의장, 쿠팡 총수(동일인)로 공식 지정",
        "공정거래위원회가 2026년 공시대상기업집단 발표와 함께 김범석 쿠팡Inc 의장을 쿠팡 그룹의 동일인(총수)으로 공식 지정했습니다. 2021년 대기업집단에 포함된 지 5년 만에 법인에서 자연인인 김 의장으로 동일인이 바뀌었습니다. 동생 김유석씨의 쿠팡 부사장 재직 관련 보수 등이 지정에 영향을 미쳤습니다.",
        "오늘의 클릭, 다음뉴스 (2026.04.30)"
      ],
      [
        "3월 산업활동동향 광공업, 서비스업 생산 증가, 전월대비 0.3% 성장",
        "국가데이터처가 30일 발표한 3월 산업활동동향에 따르면 광공업(0.3%)과 서비스업(1.4%) 생산이 늘면서 전월대비 0.3% 증가했습니다. AI 및 반도체 수요 확대에 힘입은 수출 호조가 광공업 생산을 견인했으며, 서비스업은 소비 회복세와 함께 편의점, 헬스케어 분야가 수요를 끌었습니다.",
        "대한민국 정책브리핑 (2026.04.30)"
      ],
    ]
  },
  {
    title: "3부. 경제 뉴스",
    color: DARK_GREEN,
    items: [
      [
        "코스피 사상 처음 6,600선 돌파, 삼성전자와 SK하이닉스 실적 호조",
        "4월 29일 코스피지수가 사상 처음으로 6,600선을 돌파하는 역사적인 순간을 맞이했습니다. 삼성전자와 SK하이닉스의 눈부신 1분기 실적과 AI 반도체 시장에서의 집중적 투자가 빛을 발했습니다. 외국인 순매도가 지속되는 가운데 시장 상승 모멘텀이 지속될지 관심이 집중되고 있습니다.",
        "오늘의 클릭, Investing.com (2026.04.30)"
      ],
      [
        "공동주택 공시가격 9.13% 상승 확정, 보유세 부담 현실화 우려",
        "정부가 전국 1,585만 가구를 대상으로 산정한 2026년 공동주택 공시가격을 전년 대비 9.13% 상승된 수준으로 확정 발표했습니다. 공시가격 상승으로 보유세를 포함한 세금 부담이 현실화된다는 우려가 높아지고 있습니다. 전문가들은 서울 아파트 보유세가 크게 오를 것으로 분석하고 있습니다.",
        "에콘믹글, 대한민국 정책브리핑 (2026.04.30)"
      ],
      [
        "공시대상기업집단 102곳으로 늘어, 10곳 새로 지정, 쿠팡 동일인 변경",
        "공정거래위원회가 2026년 공시대상기업집단을 102곳(소속회사 3,538개)으로 지정했습니다. 지난해 92곳에서 10곳 늘어난 수치로, 자산총액 5조 원 이상의 대기업집단이 이에 해당합니다. AI 및 산업재편 효과로 성장한 기업들이 다수 신규 진입했습니다.",
        "오늘의 클릭, 다음뉴스 (2026.04.30)"
      ],
      [
        "한국 경제 1분기 1.7% 깜짝 성장, 수출과 투자 호조가 중동전쟁 상쇄",
        "2026년 1분기 한국 경제는 실질 GDP 기준으로 1.7% 성장하며 시장 예상치를 웃도는 호실적을 기록했습니다. 반도체 수출 호조와 AI 서버 투자 확대가 성장을 이끌었습니다. 다만 2분기는 중동 분쟁에 따른 유가 급등과 글로벌 공급망 불안이 성장세 둔화 요인으로 작용할 수 있다는 전망이 나옵니다.",
        "MBC뉴스, 다음뉴스 (2026.04.30)"
      ],
      [
        "원달러 환율 1,470원대 안정, 미이란 휴전 이후 안정세 유지",
        "원달러 환율은 미국과 이란의 2주 휴전 합의 이후 1,470원대 초중반에서 안정세를 보이고 있습니다. 다만 호르무즈 해협 봉쇄 우려로 브렌트유가 배럴당 105달러를 상회하는 등 에너지 가격 변동성이 남아 있습니다. AI 반도체 수출 호조로 국내 수출 지표는 양호한 흐름을 유지하고 있습니다.",
        "나무위키, 한국무역협회 (2026.04.30)"
      ],
    ]
  },
  {
    title: "4부. 정치 뉴스",
    color: PURPLE,
    items: [
      [
        "내란재판 변론 종결, 판결 임박, 평양무인기 의혹 재판도 종결",
        "한국 내란 관련 재판이 4월 중 변론을 종결하여 판결을 앞두고 있습니다. 평양무인기 의혹 관련 1심 재판도 4월 23일 변론이 종결되었습니다. 언론사 단전 의혹 등 핵심 이슈를 둘러싼 법정 공방의 결론이 곧 나올 것으로 주목받고 있습니다.",
        "21세기 이슈, 다음뉴스 (2026.04.30)"
      ],
      [
        "국민의힘, 개헌안 반대 입장 고수, 6.3 지방선거 앞두고 정치권 긴장",
        "국민의힘이 여당의 개헌안에 반대 입장을 견지하며 개헌이 무산 위기에 처했습니다. 6.3 전국동시지방선거를 앞두고 여야 정당이 표심 경쟁에 집중하면서 제도 개혁 이슈는 다음 국회로 넘어갈 가능성이 높습니다.",
        "뉴스핌, 연합뉴스 (2026.04.30)"
      ],
      [
        "6.3 지방선거 재보궐 지역 신청 마감, 선거 행정 준비 본격화",
        "4월 30일은 임기만료 선거와 보궐선거를 동시에 실시할 수 있는 기한의 마감일입니다. 각 지역 선관위는 보궐선거 실시 여부를 결정하고 6.3 지방선거와 동시 실시 단위를 확정하는 행정 절차에 들어갔습니다. 주요 정당들의 후보 공천 전망도 점차 윤곽을 드러내고 있습니다.",
        "뉴스핌, 나무위키 (2026.04.30)"
      ],
      [
        "이 대통령, 알파고 아버지 하사비스 만나, AI 통제할 가드레일 필요",
        "이재명 대통령이 구글 딥마인드의 데미스 하사비스 CEO를 접견하고 AI를 통제할 가드레일이 필요하다고 강조했습니다. 하사비스 CEO는 AI 안전성 연구의 중요성을 강조하며 적절한 규제 패러다임 필요성에 동의했습니다. 양측은 한국에 구글 AI 캠퍼스 설립과 AI 공동 연구에 합의하며 실질적 성과를 냈습니다.",
        "MBC뉴스, 대한민국 정책브리핑 (2026.04.30)"
      ],
      [
        "오늘 국회 일정, 한국포럼, 국회기록원 개원식, 언론자유 토론회",
        "4월 30일 국회 일정으로 오전 9시 2026 한국포럼, 오전 10시 국회기록원 개원식, 오후 3시 신문방송편집인협회 토론회, 저녁 7시 국회정각회 부처님오신날 봉축점등식이 예정되어 있습니다. 신문방송편집인협회에서는 AI 시대 독립언론 보호를 주제로 토론회를 열 예정입니다.",
        "뉴스핌 (2026.04.30)"
      ],
    ]
  },
];

const children = [
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 240, after: 120 },
    children: [
      new TextRun({ text: "2026년 4월 30일 (목) 뉴스 브리핑", bold: true, size: 48, color: BLUE, font: "맑은 고딕" })
    ]
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 360 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: DARK_BLUE, space: 6 } },
    children: [
      new TextRun({ text: "AI, 사회, 경제, 정치 주요 뉴스", size: 28, color: GRAY, font: "맑은 고딕" })
    ]
  })
];

sections_data.forEach((section, sIdx) => {
  if (sIdx > 0) children.push(new Paragraph({ children: [], spacing: { before: 200, after: 0 } }));
  children.push(sectionHeader(section.title, section.color));
  section.items.forEach((item, i) => children.push(...newsItem(i + 1, item[0], item[1], item[2])));
});

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
              new TextRun({ text: "일일 뉴스 브리핑   |   2026년 4월 30일 (목)", font: "맑은 고딕", size: 18, color: "888888" }),
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
  fs.writeFileSync("일일뉴스브리핑_20260430.docx", buffer);
  console.log("성공: 일일뉴스브리핑_20260430.docx 생성완료");
}).catch(err => {
  console.error("오류:", err.message);
  process.exit(1);
});
