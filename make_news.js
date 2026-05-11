const {
  Document, Packer, Paragraph, TextRun, HeadingLevel,
  AlignmentType, BorderStyle, Header, Footer, PageNumber,
  LevelFormat, UnderlineType
} = require('docx');
const fs = require('fs');

const BLUE = "1F4E79";
const DARK_BLUE = "2E75B6";
const GRAY = "595959";
const LIGHT_GRAY = "F2F2F2";
const GREEN = "375623";
const DARK_GREEN = "538135";

function sectionHeader(text, color) {
  return new Paragraph({
    spacing: { before: 360, after: 120 },
    border: {
      bottom: { style: BorderStyle.SINGLE, size: 6, color: color || DARK_BLUE, space: 4 }
    },
    children: [
      new TextRun({
        text,
        bold: true,
        size: 32,
        color: color || BLUE,
        font: "맑은 고딕",
      })
    ]
  });
}

function newsItem(num, title, body, source) {
  return [
    new Paragraph({
      spacing: { before: 200, after: 60 },
      children: [
        new TextRun({
          text: `${num}. ${title}`,
          bold: true,
          size: 24,
          color: "000000",
          font: "맑은 고딕",
        })
      ]
    }),
    new Paragraph({
      spacing: { before: 0, after: 60 },
      indent: { left: 360 },
      children: [
        new TextRun({
          text: body,
          size: 20,
          color: GRAY,
          font: "맑은 고딕",
        })
      ]
    }),
    new Paragraph({
      spacing: { before: 0, after: 120 },
      indent: { left: 360 },
      children: [
        new TextRun({
          text: `[출처: ${source}]`,
          size: 18,
          color: "888888",
          italics: true,
          font: "맑은 고딕",
        })
      ]
    }),
  ];
}

const doc = new Document({
  styles: {
    default: {
      document: {
        run: { font: "맑은 고딕", size: 22 }
      }
    }
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
              new TextRun({ text: "일일 뉴스 브리핑", font: "맑은 고딕", size: 18, color: "888888" }),
              new TextRun({ text: "   |   2026년 4월 27일 (월)", font: "맑은 고딕", size: 18, color: "888888" }),
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
    children: [
      // Title
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 240, after: 120 },
        children: [
          new TextRun({
            text: "2026년 4월 27일 (월) 뉴스 브리핑",
            bold: true,
            size: 48,
            color: BLUE,
            font: "맑은 고딕",
          })
        ]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 360 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: DARK_BLUE, space: 6 } },
        children: [
          new TextRun({
            text: "AI · 사회 · 경제 주요 뉴스",
            size: 28,
            color: GRAY,
            font: "맑은 고딕",
          })
        ]
      }),

      // === AI 뉴스 ===
      sectionHeader("🤖  1부. AI & 인공지능 뉴스", BLUE),
      ...newsItem(
        1,
        "OpenAI, GPT-5.5 출시… 에이전틱 능력 대폭 강화",
        "OpenAI가 4월 23일부터 유료 ChatGPT 및 Codex 사용자를 대상으로 GPT-5.5 배포를 시작했습니다. GPT-5.5는 반쯤 완성된 작업을 이어받아 스스로 단계를 계획하고, 필요한 도구를 활용하며, 결과물 제출 전 자체 오류를 수정하는 에이전틱(Agentic) 능력이 크게 향상되었습니다. OpenAI는 연간 반복 매출(ARR) 250억 달러를 돌파하며 2026년 말 기업공개(IPO)를 검토 중인 것으로 알려졌습니다.",
        "OpenAI 공식 발표, CNBC (2026.04.23~27)"
      ),
      ...newsItem(
        2,
        "Google, 차세대 AI 전용 칩 TPU 8t·8i 공개",
        "구글이 대규모 사전 학습에 최적화된 TPU 8t와 추론·고동시성 작업에 특화된 TPU 8i를 발표했습니다. Google Cloud Next 2026 컨퍼런스에서 공개된 이 칩들은 엔비디아 GPU 의존도를 낮추고 자체 AI 인프라 경쟁력을 높이기 위한 전략의 일환입니다. 또한 Google Workspace Intelligence가 Gmail, Docs, Sheets, Drive 등을 통합 분석하는 기능도 함께 발표됐습니다.",
        "Google Cloud Next 2026, Fortune (2026.04.23~27)"
      ),
      ...newsItem(
        3,
        "Meta, 첫 플래그십 LLM 'Muse Spark' 공개",
        "Meta가 최고 AI 책임자 알렉산드르 왕(Alexandr Wang)의 Superintelligence Labs에서 개발한 첫 번째 주력 대형언어모델 'Muse Spark'를 선보였습니다. 멀티모달 인식, 추론, 헬스케어, 에이전틱 작업에서 경쟁력 있는 성능을 보이면서도 기존 Llama 4 대비 훨씬 낮은 컴퓨팅 비용으로 구동됩니다. Meta는 2026년 AI 설비 투자를 전년 대비 두 배 수준인 1,150~1,350억 달러로 확대할 계획입니다.",
        "CNBC (2026.04.08~27)"
      ),
      ...newsItem(
        4,
        "알리바바, AI 동영상 생성 모델 'HappyHorse-1.0' API 테스트 개시",
        "알리바바 그룹이 자체 개발한 AI 동영상 생성 모델 'HappyHorse-1.0'의 API 테스트를 알리클라우드 바이롄(百煉) 플랫폼을 통해 순차적으로 개방했습니다. 4월 말 테스트 개방에 이어 5월 중 정식 상용 출시가 예정되어 있습니다. 중국 AI 기업들의 동영상 생성 모델 경쟁이 더욱 치열해지고 있습니다.",
        "AI타임스, 인공지능신문 (2026.04.27)"
      ),
      ...newsItem(
        5,
        "Meta AI 광고 어시스턴트, 미국·유럽·아시아 글로벌 확장",
        "Meta의 AI 광고 어시스턴트가 미국(EMEA), 유럽, 아시아태평양(APAC), 중남미(LATAM) 광고주들에게 현지 언어 지원과 함께 서비스를 확장했습니다. 초기 테스터들은 광고 비용 대비 성과(CPA)가 평균 12% 개선됐다고 보고했습니다. 한편 WPP 등 글로벌 광고 기업들도 Google Earth AI를 마케팅 플랫폼에 통합하며 AI 광고 생태계가 빠르게 확산되고 있습니다.",
        "CNBC, AI Marketers (2026.04.25~27)"
      ),

      // === 사회 뉴스 ===
      new Paragraph({ children: [], spacing: { before: 200, after: 0 } }),
      sectionHeader("🏛  2부. 사회면 탑기사", "7B2C2C"),
      ...newsItem(
        1,
        "화물연대 집회 차량 돌진 사건… 운전자 살인 혐의로 구속영장",
        "경남 진주시 CU 진주물류센터 인근에서 화물연대 파업 집회 현장에 차량이 돌진해 사상자가 발생한 사건과 관련, 경남경찰청 광역수사대가 운전자 A씨에 대해 살인 및 특수상해 혐의로 구속영장을 신청했습니다. 화물연대는 올 1월부터 4개월간 7차례 교섭을 요청했으나 거부당하자 4월 5일부터 파업에 돌입했으며, 이번 사건으로 노사 갈등이 더욱 첨예해질 것으로 우려됩니다.",
        "나무위키, 다음뉴스 (2026.04.20~27)"
      ),
      ...newsItem(
        2,
        "롯데 자이언츠 선수단, 원정 도박 사건 파문",
        "프로야구 롯데 자이언츠 선수 일부가 원정 원정 경기 중 도박에 관여한 사실이 드러나 구단과 한국야구위원회(KBO)가 조사에 착수했습니다. 스포츠 도박 관련 규정 위반 선수들에 대한 징계 수위에 관심이 쏠리고 있으며, 구단 측은 사실 관계 파악 후 엄정 조치하겠다고 밝혔습니다.",
        "나무위키, 다음뉴스 (2026.04.27)"
      ),
      ...newsItem(
        3,
        "배현진 의원, 국민의힘서 당원권 1년 정지 중징계",
        "서울 송파구 국회의원 배현진이 소속 정당인 국민의힘으로부터 당원권 1년 정지 중징계를 받았습니다. 징계 이유는 공개되지 않았으나, 당 내부 갈등과 관련된 행동이 문제가 된 것으로 전해졌습니다. 배 의원 측은 불복 절차를 검토 중인 것으로 알려졌습니다.",
        "나무위키, 다음뉴스 (2026.04.27)"
      ),
      ...newsItem(
        4,
        "'충주맨' 김선태 주무관 사직서 제출… SNS 스타 공무원 퇴장",
        "유머러스한 SNS 영상으로 전국적 인기를 끈 충주시 홍보 담당 '충주맨' 김선태 주무관이 사직서를 제출한 것으로 알려졌습니다. 수백만 팔로어를 보유한 공무원 인플루언서의 퇴직 소식에 온라인에서 아쉬움의 목소리가 이어지고 있습니다. 사직 사유는 공식적으로 밝혀지지 않았습니다.",
        "나무위키, 다음뉴스 (2026.04.27)"
      ),
      ...newsItem(
        5,
        "중동 군사 긴장 고조… 미 항모 3척 호르무즈 집결",
        "미국과 이란 간 갈등이 다시 최고조로 치닫고 있습니다. 트럼프 대통령은 호르무즈 해협에서 이란 선박이 기뢰를 설치할 경우 즉각 격침하라고 명령했으며, 현재 미 항공모함 3척이 중동 해역에 집결해 있습니다. 이란 핵 협상이 교착 상태에 빠진 가운데 전면 충돌 가능성에 대한 국제사회의 우려가 커지고 있습니다.",
        "오늘의 클릭 (2026.04.25~27)"
      ),

      // === 경제 뉴스 ===
      new Paragraph({ children: [], spacing: { before: 200, after: 0 } }),
      sectionHeader("📈  3부. 경제 뉴스", DARK_GREEN),
      ...newsItem(
        1,
        "코스피 6,500선 유지… 반도체·AI 테마 강세 지속",
        "코스피가 전주에 이어 6,500선을 유지하며 강보합 흐름을 이어가고 있습니다. 반도체와 AI 데이터센터 관련 종목이 상승을 견인하는 가운데, 외국인 순매수세도 지속되고 있습니다. 증권가에서는 1분기 실적 시즌이 마무리되면서 2분기 가이던스에 시장의 이목이 쏠리고 있다고 분석했습니다.",
        "Investing.com, 한국경제 (2026.04.27)"
      ),
      ...newsItem(
        2,
        "이번 주 금융 일정: 금융위 포용금융 회의·한은 경기조사 결과 발표",
        "이억원 금융위원장 주재 '포용적 금융 대전환 제4차 회의'가 27일 오후 2시에 열립니다. 28일(화)에는 한국은행이 2026년 4월 기업경기조사(BSI) 결과와 경제심리지수(ESI)를 발표하며, 3월 금융기관 가중평균금리와 제7차 금통위 의사록도 공개될 예정입니다. 통화정책 방향에 대한 시장의 관심이 집중될 전망입니다.",
        "뉴스핌 (2026.04.24~27)"
      ),
      ...newsItem(
        3,
        "반도체 수출 호조에 GDP 성장 견조… 하반기 중동 리스크 변수",
        "1분기 반도체 수출이 예상치를 웃돌며 GDP 성장률을 전분기 대비 1.7%로 끌어올렸습니다. 다만 중동 지정학적 긴장에 따른 고유가와 공급망 불안이 2분기 이후 성장세 둔화 요인으로 작용할 수 있다는 전망이 나오고 있습니다. 한국은행은 연간 성장률 전망치 유지 여부를 이번 주 발표에서 밝힐 예정입니다.",
        "오늘의 클릭, 뉴스핌 (2026.04.24~27)"
      ),
      ...newsItem(
        4,
        "호텔신라 '매수' 상향·효성중공업 목표주가 430만 원 제시",
        "한국투자증권이 호텔신라에 대한 투자의견을 '중립'에서 '매수'로 상향 조정하고 목표주가를 올렸습니다. 면세·호텔 업황 회복과 중국 관광객 증가가 실적 개선 기대감을 높이고 있습니다. 삼성증권은 HD현대일렉트릭 계열사인 효성중공업에 대해 AI 데이터센터 전력 수요 급증에 따른 수주 증가를 근거로 목표주가를 430만 원으로 제시했습니다.",
        "뉴스핌, 한국경제 (2026.04.27)"
      ),
      ...newsItem(
        5,
        "국제유가 배럴당 105달러 상회… 에너지·방산주 강세",
        "중동 호르무즈 해협 봉쇄 우려로 브렌트유 가격이 배럴당 105달러를 웃돌며 국내 휘발유 가격에도 상방 압력을 가하고 있습니다. 이 같은 지정학적 리스크 속에서 한국 에너지·방산 관련 주식은 강세를 이어가고 있습니다. 전문가들은 유가 고공행진이 지속될 경우 하반기 물가 상승 재연 가능성을 경고했습니다.",
        "오늘의 클릭, Investing.com (2026.04.25~27)"
      ),
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("뉴스브리핑_2026-04-27.docx", buffer);
  console.log("문서 생성 완료: 뉴스브리핑_2026-04-27.docx");
});
