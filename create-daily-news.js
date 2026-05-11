const {
  Document, Packer, Paragraph, TextRun, Header, Footer,
  AlignmentType, HeadingLevel, BorderStyle, PageNumber,
  LevelFormat, SectionType
} = require('docx');
const fs = require('fs');

const CONTENT_WIDTH = 9026; // A4 with 1-inch margins

function sectionHeader(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 160 },
    border: {
      bottom: { style: BorderStyle.SINGLE, size: 6, color: "2E4D7B", space: 4 }
    },
    children: [new TextRun({ text, font: "Malgun Gothic", size: 32, bold: true, color: "2E4D7B" })]
  });
}

function newsItem(number, headline, body) {
  return [
    new Paragraph({
      spacing: { before: 200, after: 60 },
      children: [
        new TextRun({ text: `${number}. `, font: "Malgun Gothic", size: 22, bold: true, color: "1F497D" }),
        new TextRun({ text: headline, font: "Malgun Gothic", size: 22, bold: true, color: "1F497D" })
      ]
    }),
    new Paragraph({
      spacing: { before: 0, after: 160 },
      indent: { left: 320 },
      children: [new TextRun({ text: body, font: "Malgun Gothic", size: 20, color: "333333" })]
    })
  ];
}

const children = [
  // Title
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 120 },
    children: [new TextRun({ text: "데일리 뉴스 브리핑", font: "Malgun Gothic", size: 52, bold: true, color: "1A1A2E" })]
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 80 },
    children: [new TextRun({ text: "2026년 5월 6일 (수요일)", font: "Malgun Gothic", size: 28, color: "555555" })]
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 400 },
    border: { bottom: { style: BorderStyle.DOUBLE, size: 6, color: "2E4D7B", space: 4 } },
    children: [new TextRun({ text: "AI · 사회 · 경제 · 정치 · 문화", font: "Malgun Gothic", size: 22, color: "666666", italics: true })]
  }),

  // Section 1: AI
  sectionHeader("AI / 인공지능 뉴스"),
  ...newsItem(1, "AI EXPO KOREA 2026 개막 — 제9회 국제인공지능대전",
    "제9회 국제인공지능산업대전(AI EXPO KOREA 2026)이 5월 6일부터 8일까지 서울 코엑스 A홀에서 개최된다. 약 350개사 600부스 규모로, 피지컬 AI와 AI 에이전트 등 최신 트렌드를 총망라하며 참관객 5만 명이 목표다."),
  ...newsItem(2, "딥시크(DeepSeek) V4-Pro 가격 75% 인하, 가격 경쟁 재점화",
    "딥시크가 최신 V4-Pro 모델을 5월 31일까지 75% 할인된 가격으로 제공한다고 밝혔다. 입력 비용 100만 토큰당 0.0036달러로, GPT-5.5 대비 97% 저렴한 파격가로 글로벌 AI 모델 가격 경쟁이 재점화됐다."),
  ...newsItem(3, "OpenAI, 매출 목표 미달 및 수익성 위기 내부 경고",
    "OpenAI CFO가 매출 성장이 더딜 경우 AI 데이터센터 비용을 감당하지 못할 수 있다는 내부 경고를 이사회에 전달한 것으로 알려졌다. AI 에이전트 수요 급증에 대비한 데이터센터 투자 전략을 두고 내부 이견이 불거졌다."),
  ...newsItem(4, "구글 딥마인드, 한국 첫 해외 AI 캠퍼스 설립 공식 발표",
    "구글 딥마인드가 한국에 첫 해외 AI 캠퍼스를 설립한다고 5월 공식 발표했다. K-스타트업을 대상으로 한 지원 프로그램을 운영하며, 국내 AI 생태계 육성에 기여할 전망이다."),
  ...newsItem(5, "AI 에이전트·피지컬 AI 시대 본격화",
    "2026년이 AI 에이전트의 도약 원년으로 자리매김하는 가운데, 스스로 계획·실행·협업하는 에이전틱 AI가 업무 현장에 깊숙이 침투하고 있다. 디지털 공간을 넘어 기계를 매개로 인간과 직접 상호작용하는 피지컬 AI 시대도 본격 개막 중이다."),

  // Section 2: Society
  sectionHeader("사회 뉴스"),
  ...newsItem(1, "서울고법 신종오 판사 변사, 경찰 수사 착수",
    "김건희 여사의 도이치모터스 주가조작 혐의 항소심에서 유죄를 선고한 신종오 서울고법 판사가 6일 서울고법에서 숨진 채 발견됐다. 경찰은 타살 정황은 없다고 밝히며 정확한 사인을 조사 중이다."),
  ...newsItem(2, "어린이날 계기로 '노키즈존 철폐' 목소리 확산",
    "어린이날을 '어린이 차별 철폐의 날'로 지정하자는 제안이 제기됐다. 전국 노키즈존의 60% 이상이 서울·경기·제주에 집중되어 있으며, 업종별로는 카페·식당이 90% 이상을 차지하는 것으로 나타나 개선 요구가 높아지고 있다."),
  ...newsItem(3, "호르무즈 해협 한국 선박 피격·화재, 청와대 긴급 대응",
    "호르무즈 해협에서 발생한 한국 선박 화재 사고와 관련해 청와대가 긴급 대책회의를 열고 현지에 전문가를 급파했다. 중동 리스크가 에너지 공급망을 위협하는 가운데 정부의 신속한 원인 규명이 주목된다."),
  ...newsItem(4, "6일 전국 맑음, 낮 최고 27도 초여름 날씨",
    "5월 6일은 전국이 맑은 날씨를 보이는 가운데 낮 최고기온이 27도까지 오르는 초여름 날씨가 예보됐다. 낮밤 기온차가 15도 안팎으로 크게 벌어져 일교차 건강 관리가 요구된다."),
  ...newsItem(5, "초등생 70% 이상 생성형 AI 일상 활용, 교육적 대책 논의",
    "초등학생 10명 중 7명 이상이 숙제·놀이·정보검색 등에 생성형 AI를 일상적으로 활용하는 것으로 나타났다. 교육계에서는 AI 리터러시 교육 강화와 올바른 활용 지도 방안 마련을 촉구하고 있다."),

  // Section 3: Economy
  sectionHeader("경제 뉴스"),
  ...newsItem(1, "한국 1분기 GDP 1.7% 깜짝 성장, 5년 반 만의 최고치",
    "한국의 2026년 1분기 GDP가 전분기 대비 1.7% 성장해 한국은행 예상치(0.9%)를 크게 웃돌았다. 2020년 3분기 이후 최고 기록으로, 반도체·IT 수출 5% 이상 증가, 건설·설비투자 플러스 전환이 성장을 이끌었다."),
  ...newsItem(2, "코스피 6,900선 돌파, 외국인 3조 원 대규모 매수",
    "코스피 지수가 6,936.99로 5.12% 급등하며 6,900선을 돌파했다. 외국인 투자자들의 3조 원에 달하는 대규모 자금 투입이 주도했으며, 반도체 관련주를 중심으로 시장 강세가 이어지고 있다."),
  ...newsItem(3, "4월 소비자물가 1년 9개월 만에 최대 상승, 금리 인상론 부상",
    "4월 소비자물가지수가 119.37로 1년 9개월 만에 가장 큰 폭으로 올랐다. 유상대 한국은행 부총재가 '금리 인상을 고민할 때가 됐다'고 공개 발언하면서 기준금리 인상 논의가 본격화될 전망이다."),
  ...newsItem(4, "해수부, 5월 수산물 특별 할인전 개최…장바구니 물가 안정",
    "해양수산부는 6일부터 24일까지 19일간 '5월 수산물 특별 할인전'을 개최한다. 중동 리스크에 따른 국제 유가 급등과 물가 불안 속에 서민 장바구니 부담을 덜기 위한 정부 대응책으로 풀이된다."),
  ...newsItem(5, "JP모건 등 해외 IB, 한국 성장률 전망 상향 조정",
    "JP모건을 비롯한 해외 주요 투자은행들이 한국의 2026년 성장률 전망치를 상향 조정했다. 다만 전문가들은 중동 전쟁의 본격적인 영향이 2분기부터 나타날 수 있어 정부의 선제적 재정 운용이 중요하다고 강조했다."),

  // Section 4: Politics
  sectionHeader("정치 뉴스"),
  ...newsItem(1, "대통령, 국무회의 겸 비상경제점검회의 주재",
    "대통령이 6일 오전 청와대 본관에서 '2026년도 제20회 국무회의 겸 제7차 비상경제점검회의'를 주재했다. 중동 리스크 대응과 에너지·물가 안정 방안이 주요 의제에 포함된 것으로 알려졌다."),
  ...newsItem(2, "안철수, 하남갑 보궐선거 출마 공식 선언",
    "안철수 의원이 6.3 지방선거 하남갑 보궐선거 출마를 공식 선언했다. 기자회견을 통해 출마 의지를 밝히며 지역 현안 해결을 출마의 이유로 제시했다."),
  ...newsItem(3, "신종오 판사 변사에 정치권 파장 확산",
    "김건희 여사 도이치모터스 주가조작 항소심 유죄를 선고한 신종오 판사가 사망한 사건이 정치적 파장을 불러일으키고 있다. 여야는 상반된 반응을 보이며 진상 규명을 촉구하고 있다."),
  ...newsItem(4, "6.3 지방선거 앞두고 여야 복지 정책 경쟁 치열",
    "6.3 지방선거를 앞두고 여야 각 당이 복지·교육·출산 정책을 앞다퉈 내놓으며 민심 잡기 경쟁을 벌이고 있다. 국회에서는 범사회복지정책연합 정책 제안 기자회견이 열리는 등 복지 입법 논의가 가열됐다."),
  ...newsItem(5, "국회 정치 일정 — 기자회견·상임위 활발",
    "국회 상임위와 각 정당의 기자회견이 잇따라 예정됐다. 임미애 의원의 나영민 전 의장 추대 기자회견, 남인순 의원의 6.3 지방선거 정책 제안 등이 주목받고 있다."),

  // Section 5: Culture
  sectionHeader("문화 뉴스"),
  ...newsItem(1, "아이유, 어린이날 맞아 1억 원 기부",
    "가수 겸 배우 아이유가 어린이날을 맞아 소외 어린이를 위한 기부금 1억 원을 쾌척했다. 꾸준한 사회 공헌 활동으로 귀감이 되고 있는 아이유는 이번에도 팬들에게 훈훈한 감동을 안겼다."),
  ...newsItem(2, "한지민, 19년 연속 어린이날 기부 이어가",
    "배우 한지민이 19년 연속으로 어린이날 기부 활동을 이어갔다. 올해는 5천만 원을 기부해 사회적 나눔 문화 선도자로서 입지를 굳히며 팬들의 뜨거운 박수를 받고 있다."),
  ...newsItem(3, "지드래곤, 해외 무대 인종차별 의상 논란 휩싸여",
    "가수 지드래곤이 해외 공연 무대 의상이 인종차별 논란을 불러일으키며 '국제 망신'이라는 비판에 직면했다. 현재 소속사 측은 경위 파악 중이며 공식 입장을 준비하고 있다."),
  ...newsItem(4, "MBC '라디오스타', 최다니엘·남규리·안지영 출연",
    "MBC 예능 프로그램 '라디오스타'가 6일 '만나서 안광입니다' 특집을 방영한다. 배우 최다니엘, 가수 겸 배우 남규리, 가수 안지영이 출연해 솔직한 이야기를 나눌 예정이다."),
  ...newsItem(5, "K-팝 컴백 대전·드라마 '윰세3' 화제",
    "5월 K-팝 대형 아티스트들의 컴백 러시로 팬덤 열기가 고조되고 있다. 드라마 '윰세3'에서 주인공 유미 역을 맡은 배우 김고은은 '5년 동안 유미로 살아 행복했다'는 소감을 밝혀 화제를 모았다."),

  // Footer separator
  new Paragraph({
    spacing: { before: 400, after: 120 },
    border: { top: { style: BorderStyle.SINGLE, size: 4, color: "999999", space: 4 } },
    alignment: AlignmentType.CENTER,
    children: [new TextRun({
      text: "본 뉴스 브리핑은 2026년 5월 6일(수) 기준으로 자동 수집·정리된 자료입니다. 각 기사의 세부 내용은 원문 출처를 확인하시기 바랍니다.",
      font: "Malgun Gothic", size: 18, color: "888888", italics: true
    })]
  }),
];

const doc = new Document({
  styles: {
    default: {
      document: { run: { font: "Malgun Gothic", size: 22 } }
    },
    paragraphStyles: [
      {
        id: "Heading1",
        name: "Heading 1",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: { size: 32, bold: true, font: "Malgun Gothic", color: "2E4D7B" },
        paragraph: { spacing: { before: 360, after: 160 }, outlineLevel: 0 }
      }
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
          children: [new TextRun({ text: "데일리 뉴스 브리핑 | 2026.05.06", font: "Malgun Gothic", size: 18, color: "888888" })]
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
            new TextRun({ text: " -", font: "Malgun Gothic", size: 18, color: "888888" }),
          ]
        })]
      })
    },
    children
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("D:\\AI \xEC\x8B\xA4\xEC\x8A\xB5\\클로드 코드\\news\\daily-news-2026-05-06.docx", buffer);
  console.log("DOCX created successfully!");
}).catch(err => {
  console.error("Error:", err);
  process.exit(1);
});
