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
    "삼성전자 AI 반도체 수요 급증으로 역대급 실적",
    "삼성전자가 2026년 1분기 매출 133조원, 영업이익 57.2조원이라는 창사 이래 최대 실적을 달성했다. 인공지능(AI) 기술 발전에 따른 고대역폭 메모리(HBM) 수요 급증이 실적을 이끌었다는 분석이다.",
    "오늘의 클릭 (www.cliktoday.com)"
  ],
  [
    "한국 연구진, 세계 최초 초저전력 AI 칩 개발 성공",
    "국내 연구진이 세계 최초로 초저전력 AI 칩 개발에 성공했다. 이는 온디바이스 AI 시장에서 한국이 선도적인 역할을 할 수 있는 기반을 마련한 것으로, 향후 스마트폰·웨어러블 기기에 탑재될 전망이다.",
    "오늘의 클릭 (www.cliktoday.com)"
  ],
  [
    "AI 성과 검증 시대 진입... '도입'서 '실적'으로 전환",
    "AI 도입 경쟁이 마무리 단계에 접어들면서 기업들은 이제 구체적인 성과 제출을 요구받고 있다. 빅테크 기업의 대규모 투자 이후 투자자들이 수익성 검증을 요구하며 AI 산업이 성장 기대 단계에서 실적 검증 단계로 이동 중이다.",
    "자본시장뉴스 (www.jabon.co.kr)"
  ],
  [
    "딥시크, AI 모델 가격 대폭 인하... 업계 가격 경쟁 재점화",
    "중국 AI 스타트업 딥시크(DeepSeek)가 자사 AI 모델의 가격을 대폭 낮추면서 글로벌 AI 서비스 업계의 가격 경쟁이 다시 불붙었다. 오픈AI·구글 등 주요 기업들의 가격 정책에도 영향을 미칠 것으로 보인다.",
    "Investing.com Korea (kr.investing.com)"
  ],
  [
    "한국인공지능협회, 아시아 최대 AI 행사 5월 6~8일 코엑스 개최",
    "한국인공지능협회가 오는 5월 6일부터 8일까지 사흘간 서울 삼성동 코엑스에서 제9회 AI 컨퍼런스를 개최한다. 이번 행사는 단일 인공지능 행사로 아시아 최대 규모이며, 글로벌 AI 리더들이 대거 참가할 예정이다.",
    "AI매터스 (aimatters.co.kr)"
  ]
];

const socNews = [
  [
    "'노동절' 63년 만에 공식 부활... 법정 공휴일 첫날",
    "63년 만에 '근로자의 날'이 '노동절'로 공식 복원되어 법정 공휴일 첫날을 맞았다. 전국 각지에서 노동절 기념 집회와 행사가 열렸으며, 이재명 대통령은 노사 상생 모델을 강조하는 메시지를 발표했다.",
    "오늘의 클릭 (www.cliktoday.com)"
  ],
  [
    "삼성바이오로직스 노조, 창사 이래 첫 전면 파업 돌입",
    "삼성바이오로직스 노동조합이 창사 이래 처음으로 전면 파업에 돌입했다. 노조는 영업이익의 20%를 성과급으로 지급할 것을 요구하고 있으며, 사측과 교섭이 결렬되면서 생산 차질이 우려된다.",
    "오늘의 클릭 (www.cliktoday.com)"
  ],
  [
    "황금연휴 프리미엄 식재료 판매 폭발... 랍스터 229%·장어 169% 급증",
    "5월 장기 연휴를 맞아 랍스터, 장어, 게 등 프리미엄 식재료 판매가 폭발적으로 증가했다. 업계에 따르면 랍스터는 전년 동기 대비 229%, 장어는 169% 매출이 급증하며 외식 대신 홈파티 트렌드가 뚜렷하게 나타났다.",
    "서울신문 (www.seoul.co.kr)"
  ],
  [
    "서울 강남권 부동산 급매물 등장... 양도세 중과 유예 종료 임박",
    "양도세 중과 유예 종료가 다가오면서 서울 강남권과 마포 등 주요 지역에서 급매물이 추가로 등록되는 움직임이 나타나고 있다. 전문가들은 유예 기간 내 처분을 서두르는 다주택자들의 매물이 단기적으로 시장에 영향을 줄 것으로 분석했다.",
    "경향신문 (www.khan.co.kr)"
  ],
  [
    "정부, 고유가 피해지원금 모든 주유소로 확대... 매출 제한 폐지",
    "정부가 5월 1일부터 고유가 피해지원금 사용처를 연 매출액과 관계없이 모든 주유소로 확대했다. 기존에는 연 매출 30억원 이하 주유소에만 허용됐으나, 이번 조치로 대형 주유소도 지원 대상에 포함됐다.",
    "대한민국 정책브리핑 (www.korea.kr)"
  ]
];

const ecoNews = [
  [
    "코스피 6,600선 유지... 국제유가 급등에 상승 제한",
    "코스피 지수가 국제유가 급등의 영향으로 6,600선 부근에서 등락을 거듭했다. 글로벌 빅테크 기업의 실적 호조로 기술주가 강세를 보였으나, 에너지 관련 인플레이션 우려가 지수 상승을 제한했다.",
    "한국경제 (www.hankyung.com)"
  ],
  [
    "달러-원 환율 1,471원... 원화 소폭 강세",
    "달러 대비 원화 환율이 1,471원으로 전일 대비 소폭 하락(원화 강세)했다. 달러 약세와 유가 하락, 외국인 투자심리 개선이 맞물리며 원화가 최근 강세 흐름을 이어갔다.",
    "Investing.com Korea (kr.investing.com)"
  ],
  [
    "카카오페이·농심 등 1분기 깜짝 실적... 컨센서스 대폭 상회",
    "카카오페이의 2026년 1분기 영업이익이 최대 320억원, 농심은 최대 650억원으로 시장 컨센서스를 크게 웃돌 것으로 추정됐다. 내수 소비 회복과 플랫폼 광고 수익 개선이 실적 호전의 주요 요인으로 꼽혔다.",
    "한국경제 (www.hankyung.com)"
  ],
  [
    "샤오미, AI 전략·전기차 확장 계획 발표... 글로벌 시장 공략 본격화",
    "중국 샤오미가 2026 투자자의 날(Investor Day) 행사에서 AI 전략 및 전기차(EV) 사업 확장 계획을 발표했다. 자체 AI 모델 개발과 함께 전기차 신모델 3종을 추가 출시할 계획이며, 한국 등 아시아 시장 진출도 검토 중인 것으로 알려졌다.",
    "Investing.com Korea (kr.investing.com)"
  ],
  [
    "한국은행, 2026년 경제성장률 1.8% 전망... 상반기 강세·하반기 둔화",
    "한국은행은 2026년 한국 경제성장률을 1.8%로 전망하며, 상반기 성장세 강화 후 하반기 둔화 흐름을 예상했다. 소비 중심의 내수 회복세가 견조하게 유지되고, 정부의 재정 확대가 경기 하방을 완충할 것으로 분석됐다.",
    "PwC 삼일회계법인 (www.pwc.com/kr)"
  ]
];

const polNews = [
  [
    "6.3 지방선거 D-30... 거물급 출마로 '미니 총선급' 판 커져",
    "6·3 지방선거를 한 달 앞두고 한동훈·조국·송영길 등 거물급 정치인들이 대거 출마 선언을 하며 이번 선거가 단순 의석 경쟁을 넘어 여야 권력 구조 재편의 분수령이 될 전망이다. 전문가들은 일부 접전지에서 막판 변동성이 클 것으로 예상했다.",
    "머니투데이 (www.mt.co.kr)"
  ],
  [
    "민주당, 국민의힘 공천 결과 \"쇄신 없는 윤어게인\" 맹비판",
    "더불어민주당이 국민의힘의 6·3 지방선거 및 국회의원 재·보궐선거 공천 결과를 두고 쇄신은 없고 윤어게인으로 귀결됐다고 강하게 비판했다. 민주당은 공천 과정에서의 친윤계 중심 구성을 문제 삼으며 정권 심판론을 부각시켰다.",
    "경향신문 (www.khan.co.kr)"
  ],
  [
    "조국, \"이재명 정부의 레드팀 역할 수행하겠다\"",
    "조국 조국혁신당 대표가 거대 양당이 오랜 기간 과점한 대한민국 정치에는 객토가 필요하다며 이재명 정부의 레드팀 역할을 수행하겠다고 선언했다. 조 대표는 6.3 지방선거에서 제3지대 세력 확장을 목표로 당의 전략을 가다듬고 있다.",
    "경향신문 (www.khan.co.kr)"
  ],
  [
    "국민의힘 장동혁 대표, 대구 방문... 추경호 시장 후보 개소식 참석",
    "국민의힘 장동혁 대표가 6.3 지방선거를 한 달 앞두고 대구를 방문해 추경호 대구시장 후보 선거사무소 개소식에 참석했다. 장 대표는 TK(대구·경북) 민심 다독이기와 함께 선거 승리를 위한 지원 사격에 나섰다.",
    "YTN (www.ytn.co.kr)"
  ],
  [
    "조정식 민주당 의원, 후반기 국회의장 도전 선언... 보좌관직 사퇴",
    "더불어민주당 조정식 의원(6선, 경기 시흥을)이 22대 후반기 국회의장직에 도전한다고 밝히며 대통령 정무특별보좌관직을 내려놓았다. 조 의원은 풍부한 의정 경험과 여야 소통 능력을 강점으로 내세웠다.",
    "다음뉴스 (news.daum.net)"
  ]
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
        children: [
          new Paragraph({
            alignment: AlignmentType.RIGHT,
            border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: "CCCCCC", space: 4 } },
            children: [
              new TextRun({ text: "일간 뉴스 브리핑  |  2026년 5월 3일 (일요일)", size: 17, color: "999999", font: "맑은 고딕" })
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
            border: { top: { style: BorderStyle.SINGLE, size: 2, color: "CCCCCC", space: 4 } },
            children: [
              new TextRun({ text: "본 브리핑은 AI가 자동 수집·정리한 뉴스입니다.  |  ", size: 17, color: "999999", font: "맑은 고딕" }),
              new TextRun({ children: [PageNumber.CURRENT], size: 17, color: "999999", font: "맑은 고딕" }),
            ]
          })
        ]
      })
    },
    children: [
      // Title block
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 100 },
        border: { bottom: { style: BorderStyle.THICK, size: 10, color: BLUE, space: 4 } },
        children: [
          new TextRun({ text: "일간 뉴스 브리핑", bold: true, size: 52, color: BLUE, font: "맑은 고딕" })
        ]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 100, after: 80 },
        children: [
          new TextRun({ text: "2026년 5월 3일 (일요일)", size: 26, color: GRAY, font: "맑은 고딕" })
        ]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 300 },
        children: [
          new TextRun({ text: "AI·인공지능  |  사회  |  경제  |  정치", size: 20, color: "AAAAAA", font: "맑은 고딕" })
        ]
      }),

      // Intro
      new Paragraph({
        spacing: { before: 0, after: 360 },
        border: { left: { style: BorderStyle.THICK, size: 12, color: DARK_BLUE, space: 8 } },
        indent: { left: 240 },
        children: [
          new TextRun({
            text: "오늘의 주요 뉴스를 AI·인공지능, 사회, 경제, 정치 4개 분야로 나누어 각 5건씩 정리했습니다. 빠르게 변화하는 국내외 이슈를 한눈에 파악하시기 바랍니다.",
            size: 20, color: GRAY, font: "맑은 고딕"
          })
        ]
      }),

      // AI Section
      sectionHeader("[AI·인공지능 뉴스]", DARK_BLUE),
      ...aiNews.flatMap((item, i) => newsItem(i + 1, ...item)),

      // Society Section
      sectionHeader("[사회면 주요 뉴스]", DARK_GREEN),
      ...socNews.flatMap((item, i) => newsItem(i + 1, ...item)),

      // Economy Section
      sectionHeader("[경제 뉴스]", ORANGE),
      ...ecoNews.flatMap((item, i) => newsItem(i + 1, ...item)),

      // Politics Section
      sectionHeader("[정치 뉴스]", PURPLE),
      ...polNews.flatMap((item, i) => newsItem(i + 1, ...item)),
    ]
  }]
});

const outPath = "D:/AI 실습/클로드 코드/news/뉴스브리핑_2026-05-03.docx";
Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync(outPath, buffer);
  console.log("SUCCESS: " + outPath + " (" + buffer.length + " bytes)");
}).catch(err => {
  console.error("Error:", err.message);
  process.exit(1);
});
