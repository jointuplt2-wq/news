const {
  Document, Packer, Paragraph, TextRun, HeadingLevel,
  AlignmentType, BorderStyle, LevelFormat, Header, Footer,
  PageNumber, ShadingType, Table, TableRow, TableCell, WidthType
} = require('docx');
const fs = require('fs');

const today = '2026년 4월 30일 (목요일)';

const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };

const aiColor = "1565C0";
const societyColor = "2E7D32";
const econColor = "6A1B9A";
const politicsColor = "B71C1C";

function sectionTitle(text, color) {
  return new Paragraph({
    spacing: { before: 360, after: 120 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: color, space: 4 } },
    children: [
      new TextRun({
        text,
        bold: true,
        size: 32,
        color: color,
        font: "맑은 고딕",
      })
    ]
  });
}

function newsItem(num, title, summary, source) {
  return [
    new Paragraph({
      spacing: { before: 200, after: 60 },
      children: [
        new TextRun({ text: `${num}. `, bold: true, size: 24, font: "맑은 고딕", color: "333333" }),
        new TextRun({ text: title, bold: true, size: 24, font: "맑은 고딕", color: "111111" }),
      ]
    }),
    new Paragraph({
      spacing: { before: 0, after: 40 },
      indent: { left: 360 },
      children: [
        new TextRun({ text: summary, size: 20, font: "맑은 고딕", color: "444444" }),
      ]
    }),
    new Paragraph({
      spacing: { before: 0, after: 140 },
      indent: { left: 360 },
      children: [
        new TextRun({ text: `[출처: ${source}]`, size: 18, font: "맑은 고딕", color: "888888", italics: true }),
      ]
    }),
  ];
}

// --- AI 뉴스 ---
const aiNews = [
  {
    num: 1,
    title: "OpenAI, ChatGPT Images 2.0 출시 — 추론 결합 이미지 생성 모델",
    summary: "OpenAI가 4월 21일 추론 기능이 통합된 이미지 생성 모델 'ChatGPT Images 2.0'을 공개했다. Thinking 모드는 프롬프트를 분석, 검색, 검증한 후 이미지를 생성하며 최대 8장의 이미지를 일관성 있게 만들 수 있다. 출시 12시간 만에 Image Arena 리더보드 1위를 차지했다.",
    source: "AI타임스 / Carrot Global Blog"
  },
  {
    num: 2,
    title: "OpenAI GPT-5.5 Thinking·Pro 출시 — 워크스페이스 에이전트 기능 포함",
    summary: "4월 24일 OpenAI가 GPT-5.5 Thinking과 GPT-5.5 Pro를 공개했다. 팀 단위 작업을 처리하는 워크스페이스 에이전트 기능도 함께 선보였으며, 빠른 업데이트 사이클로 AI 경쟁이 더욱 가속화되고 있다.",
    source: "Carrot Global Blog"
  },
  {
    num: 3,
    title: "DeepSeek V4-Pro·V4-Flash 오픈소스 공개 — GPT·Claude 코딩 벤치마크서 제쳐",
    summary: "중국 AI 스타트업 DeepSeek이 V4-Pro와 V4-Flash를 오픈소스로 공개했다. V4-Pro는 주요 코딩 벤치마크에서 GPT-5.4, Claude Opus 4.6, Gemini 3.1-Pro를 앞질렀다. DeepSeek은 API 가격도 기존 대비 최대 10분의 1 수준으로 대폭 인하했다.",
    source: "Investing.com / AI타임스"
  },
  {
    num: 4,
    title: "엔비디아, 한국형 AI 합성 데이터셋 '네모트론-페르소나-코리아' 허깅페이스 공개",
    summary: "엔비디아가 한국어 특화 AI 학습용 합성 데이터셋 'Nemotron-Personas-Korea'를 허깅페이스(Hugging Face)에 공개했다. 한국어 LLM 모델 개발에 필요한 고품질 데이터 확보가 한층 용이해질 전망이다.",
    source: "전자신문 / AI타임스"
  },
  {
    num: 5,
    title: "AI 주도 사이버 공격 급증 — 한국 정보기관 주요 안보 위협 공식 지목",
    summary: "2026년 들어 AI가 독립적으로 사이버 공격을 수행하는 단계에 이르렀으며, 한국의 정보기관이 AI 기반 사이버 공격을 주요 국가 안보 위협으로 공식 지목했다. 과기정통부는 AI 스타트업 육성을 위해 1조 9,800억 원 규모 민관 합동 투자를 추진 중이다.",
    source: "전국인력신문 / 전자신문"
  },
];

// --- 사회 뉴스 ---
const societyNews = [
  {
    num: 1,
    title: "전국삼성전자노조, 이재용 회장 자택 옆 천막 설치 — 총파업 앞두고 면담 요구",
    summary: "전국삼성전자노동조합이 총파업을 앞두고 이재용 회장 자택 인근에 천막을 설치하며 직접 면담을 촉구했다. 노조 측은 임금 인상과 처우 개선을 핵심 요구로 내세우고 있으며, 협상 결렬 시 전면 파업 확대가 예상된다.",
    source: "머니투데이 / 경향신문"
  },
  {
    num: 2,
    title: "특검팀, 김건희 여사 관저 설계도 전달 — 공무상 비밀누설죄 적용 검토",
    summary: "특검팀은 김건희 여사가 청와대 관저 설계도를 외부에 전달한 행위에 공무상 비밀누설죄를 적용할 수 있는지 수사 중이다. 해당 문서는 보안 등급이 부여된 자료로 관련 수사가 진행 중이다.",
    source: "경향신문"
  },
  {
    num: 3,
    title: "공정위, 쿠팡 총수 법인에서 김범석 의장으로 변경 지정 — 5년 만에 자연인 총수",
    summary: "공정거래위원회가 쿠팡의 동일인(총수)을 법인 '쿠팡(주)'에서 자연인 김범석 쿠팡Inc 의장으로 변경 지정했다. 대기업집단 포함 5년 만의 변경으로, 김 의장 동생 김유석 부사장의 경영 참여로 사익편취 우려가 있다고 판단했다.",
    source: "머니투데이 / 연합뉴스"
  },
  {
    num: 4,
    title: "서울 개별공시지가 전년 대비 4.9% 상승 — 용산, 성동, 강남 순으로 올라",
    summary: "올해 서울 개별공시지가가 전년 대비 평균 4.9% 상승했다. 구별로는 용산구 상승폭이 가장 크고, 이어 성동구와 강남구 순으로 높은 상승률을 기록했다. 재개발, 재건축 기대감이 반영된 결과로 분석된다.",
    source: "연합뉴스"
  },
  {
    num: 5,
    title: "3월 산업활동 생산, 소비, 투자 '트리플 상승' — 지난해 9월 이후 첫 동반 증가",
    summary: "통계청이 발표한 3월 산업활동동향에 따르면 생산, 소비, 투자가 모두 증가하는 '트리플 상승'이 지난해 9월 이후 처음으로 나타났다. 전 산업 생산은 광공업, 서비스업이 견인하며 전월 대비 0.3% 증가했다.",
    source: "통계청 / 연합뉴스"
  },
];

// --- 경제 뉴스 ---
const econNews = [
  {
    num: 1,
    title: "코스피 사상 최고치 갱신, 6,600선 돌파 — AI 반도체 주도 상승",
    summary: "코스피가 6,600선을 돌파하며 사상 최고치를 경신했다. 삼성전자, SK하이닉스 등 AI 반도체 기업들의 1분기 실적 기대감으로 외국인 매수세가 집중됐으며, 한국 증시 시가총액 세계 8위권 진입이 가시화되고 있다.",
    source: "오늘의 클릭 / 한국경제"
  },
  {
    num: 2,
    title: "외환시장 1분기 일평균 거래액 1,026억 달러 — 1분기 기준 역대 최대",
    summary: "국내 외환시장 1분기 일평균 거래 규모가 1,026억 5,000만 달러를 기록하며 1분기 최초로 1,000억 달러를 돌파했다. WGBI 편입 이후 외국인 국고채 순매수가 10조 원에 달하는 등 자본 유입이 지속되고 있다.",
    source: "머니투데이 / 한국은행"
  },
  {
    num: 3,
    title: "국제 유가 중동 긴장으로 전쟁 이전 대비 39% 급등 — 물가 압박 우려",
    summary: "중동 지정학적 긴장 장기화로 국제 유가가 전쟁 이전 수준 대비 약 39% 급등했다. 에너지 가격 상승에 따른 제조업, 물류 비용 증가로 금리 동결 기조가 이어질 전망이며, 소비자물가 상승 압력도 지속되고 있다.",
    source: "오늘의 클릭 / Investing.com"
  },
  {
    num: 4,
    title: "삼성전자, SK하이닉스 AI 서버 HBM 수요 기대감에 외국인 순매수 집중",
    summary: "삼성전자와 SK하이닉스가 1분기 실적 발표를 앞두고 AI 서버용 고대역폭메모리(HBM) 수요 확대 기대감에 외국인 순매수세를 이어가고 있다. 두 기업 합산 시가총액 상승이 코스피 전체 지수를 주도하는 양상이다.",
    source: "한국경제 / 오늘의 클릭"
  },
  {
    num: 5,
    title: "전문가 \"코스피 단기 과열 경고\" — 개인 투자자 차익 실현 매물 주의",
    summary: "증권 전문가들이 코스피의 급격한 상승에 따른 단기 과열을 경고하고 있다. 외국인 매수세와 달리 개인 투자자들의 차익 실현 매물이 늘어날 가능성이 높아, 단기 조정 가능성에 대비해야 한다는 의견이 나오고 있다.",
    source: "한국경제 / 머니투데이"
  },
];

// --- 정치 뉴스 ---
const politicsNews = [
  {
    num: 1,
    title: "종합특검, 윤석열 전 대통령에 4월 30일 출석 통보 — 변호인 측 불응 예정",
    summary: "2차 종합특검이 윤석열 전 대통령에게 피의자 신분으로 4월 30일 출석할 것을 통보했으나, 변호인 측은 출석하지 않겠다는 입장을 밝혔다. 특검은 내란 우두머리 혐의와 별개로 군형법상 반란 혐의 추가 적용도 검토 중이다.",
    source: "연합뉴스 / 경향신문"
  },
  {
    num: 2,
    title: "내란 우두머리 항소심(2심) 개시 — 1심 무기징역 선고 67일 만에",
    summary: "12, 3 비상계엄 사태와 관련해 1심에서 무기징역을 선고받은 윤석열 전 대통령의 항소심이 개시됐다. 서울고법 내란전담재판부 구성 이후 첫 항소심으로, 심리 속도와 결과에 각계의 관심이 집중되고 있다.",
    source: "21세기 이슈 / 경향신문"
  },
  {
    num: 3,
    title: "내란 재판 변론 종결 — 평양무인기, 언론사 단전 의혹 판결 단계 진입",
    summary: "평양 무인기 의혹 재판과 언론사 단전, 단수 지시 의혹 재판이 4월 23일 변론을 종결하며 판결 선고 단계에 들어갔다. 두 사건 모두 국가 정치 상황과 법치주의에 영향을 미칠 중요한 분수령으로 평가된다.",
    source: "21세기 이슈"
  },
  {
    num: 4,
    title: "내란전담재판부법 위헌 헌법소원 정식심판 회부 — 헌재 본격 심리 개시",
    summary: "헌법재판소가 4월 22일 내란전담재판부법의 위헌성을 다투는 헌법소원을 정식심판에 회부했다. 헌재가 해당 법률의 위헌 여부를 본격 심리하겠다는 의미로, 내란 관련 재판 전체에 파급 효과가 예상된다.",
    source: "21세기 이슈 / 법률신문"
  },
  {
    num: 5,
    title: "한국정치학회, '북한이냐 조선이냐' 특별학술회의 개최 — 통일부 후원",
    summary: "한국정치학회가 4월 30일 서울 프레스센터에서 '평화 공존을 위한 이름 부르기: 북한인가 조선인가' 주제의 특별학술회의를 개최했다. 남북 관계와 호칭의 정치적 함의를 학문적으로 재검토하는 자리로 통일부가 후원했다.",
    source: "뉴스핌"
  },
];

const doc = new Document({
  styles: {
    default: {
      document: { run: { font: "맑은 고딕", size: 22 } }
    }
  },
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
            border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "2E75B6", space: 4 } },
            children: [
              new TextRun({ text: "일일 뉴스 브리핑  |  Daily News Briefing", font: "맑은 고딕", size: 18, color: "555555" }),
              new TextRun({ text: "\t" + today, font: "맑은 고딕", size: 18, color: "888888" }),
            ],
            tabStops: [{ type: "right", position: 9360 }],
          })
        ]
      })
    },
    footers: {
      default: new Footer({
        children: [
          new Paragraph({
            border: { top: { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC", space: 4 } },
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({ text: "본 브리핑은 공개된 뉴스 정보를 바탕으로 자동 생성되었습니다.  |  Page ", font: "맑은 고딕", size: 16, color: "999999" }),
              new TextRun({ children: [PageNumber.CURRENT], font: "맑은 고딕", size: 16, color: "999999" }),
            ]
          })
        ]
      })
    },
    children: [
      // Cover Title
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 480, after: 120 },
        children: [
          new TextRun({ text: "일일 뉴스 브리핑", bold: true, size: 52, font: "맑은 고딕", color: "1F3864" })
        ]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 80 },
        children: [
          new TextRun({ text: "Daily News Briefing", size: 28, font: "맑은 고딕", color: "666666" })
        ]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 480 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: "2E75B6", space: 8 } },
        children: [
          new TextRun({ text: today, bold: true, size: 28, font: "맑은 고딕", color: "2E75B6" })
        ]
      }),

      // AI Section
      sectionTitle("AI / 인공지능 뉴스", aiColor),
      ...aiNews.flatMap(n => newsItem(n.num, n.title, n.summary, n.source)),

      // Society Section
      sectionTitle("사회 뉴스", societyColor),
      ...societyNews.flatMap(n => newsItem(n.num, n.title, n.summary, n.source)),

      // Economy Section
      sectionTitle("경제 뉴스", econColor),
      ...econNews.flatMap(n => newsItem(n.num, n.title, n.summary, n.source)),

      // Politics Section
      sectionTitle("정치 뉴스", politicsColor),
      ...politicsNews.flatMap(n => newsItem(n.num, n.title, n.summary, n.source)),
    ]
  }]
});

const outputPath = "D:/AI 실습/클로드 코드/news/일일뉴스브리핑_20260430.docx";
Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync(outputPath, buffer);
  console.log("Done: " + outputPath);
}).catch(err => {
  console.error('Error:', err);
  process.exit(1);
});
