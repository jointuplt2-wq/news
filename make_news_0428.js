const {
  Document, Packer, Paragraph, TextRun,
  AlignmentType, BorderStyle, Header, Footer, PageNumber,
  LevelFormat
} = require('docx');
const fs = require('fs');

const BLUE = "1F4E79";
const DARK_BLUE = "2E75B6";
const GRAY = "595959";
const DARK_GREEN = "375623";
const RED_DARK = "7B2C2C";

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
              new TextRun({ text: "   |   2026년 4월 28일 (화)", font: "맑은 고딕", size: 18, color: "888888" }),
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
            text: "2026년 4월 28일 (화) 뉴스 브리핑",
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
      sectionHeader("1부. AI & 인공지능 뉴스", BLUE),
      ...newsItem(
        1,
        "DeepSeek V4 출시 — 최첨단 AI와 격차 좁혀, 가격은 1/6 수준",
        "중국 AI 스타트업 DeepSeek이 4월 24일 최신 모델 V4의 두 가지 프리뷰 버전(V4 Flash, V4 Pro)을 공개했다. 각각 100만 토큰 컨텍스트 창을 지원하는 혼합전문가(MoE) 아키텍처로, 추론 벤치마크에서 GPT-5.4, Gemini 3.1 Pro 등 최첨단 모델과의 격차를 거의 좁혔다고 밝혔다. 가격 면에서는 V4 Pro가 입력 토큰 100만 개당 $0.145로 Claude Opus 4.7, GPT-5.5 대비 약 1/6 수준이다. 다만 지식 테스트에서는 프론티어 모델 대비 3~6개월의 발전 차이가 있는 것으로 평가된다.",
        "TechCrunch, VentureBeat (2026.04.24)"
      ),
      ...newsItem(
        2,
        "GPT-5.5 정식 출시 — OpenAI, 통합 AI 슈퍼앱 전략 가속",
        "OpenAI는 4월 23일 GPT-5.5를 정식 출시하며, ChatGPT·코딩 도구·브라우저 기능을 단일 인터페이스로 통합하는 AI 슈퍼앱 전략을 공식화했다. 앞서 Anthropic은 4월 16일 Claude Opus 4.7을 출시한 바 있어, 4월 한 달간 주요 AI 모델 3종이 연이어 발표되는 이례적인 경쟁 구도가 펼쳐졌다. DeepSeek은 GPT-5.5 출시 하루 뒤 V4를 공개해 더욱 주목받았다.",
        "aithority.com, llm-stats.com (2026.04.23~24)"
      ),
      ...newsItem(
        3,
        "Snap, AI 코드 생성 65% 달성 — 직원 1,000명(16%) 감원",
        "소셜미디어 기업 Snap이 4월 15일 전체 직원의 16%에 해당하는 약 1,000명을 해고하고 300개 이상의 채용 포지션을 폐지했다. CEO 에반 스피겔은 AI가 현재 Snap 신규 코드의 65% 이상을 생성한다며 AI 효율화를 이유로 들었다. 연간 비용 절감액은 2026년 중반까지 5억 달러 이상이 될 것으로 예상된다. AI로 인한 대규모 인력 감축의 대표적인 사례로 주목받고 있다.",
        "crescendo.ai (2026.04.15)"
      ),
      ...newsItem(
        4,
        "미 공군, AI 워게임 시스템 WarMatrix 실전 배치",
        "미 공군이 150명 이상이 참가하는 대규모 워게임에 AI 기반 작전 시뮬레이션 시스템 WarMatrix를 처음 실전 배치했다. 이 시스템은 실시간보다 최대 1만 배 빠른 속도로 시뮬레이션을 수행하면서도 모든 최종 결정은 인간이 내리는 구조를 유지한다. AI가 군사 전략 분야에도 깊숙이 침투하고 있음을 보여주는 사례로, 국방 AI 활용의 새로운 이정표로 평가된다.",
        "crescendo.ai (2026.04.15)"
      ),
      ...newsItem(
        5,
        "AI 환각으로 허위 판례 제출한 변호사 자격 정지 — 법조계 경고 확산",
        "미국 네브래스카주 대법원은 AI가 만들어낸 환각(hallucination) 판례 20건을 항소 서면에 포함한 오마하 변호사 그레그 레이크에게 자격 정지 처분을 내렸다. 법원은 그가 AI 사용을 부인한 것도 신뢰성이 없다고 판단했다. 한편 뉴욕 연방판사 제드 라코프는 AI 챗봇 대화 내용에 변호사-의뢰인 특권이 적용되지 않는다고 판결해 법조계에 AI 사용 주의보가 확산되고 있다.",
        "crescendo.ai (2026.04.16)"
      ),

      // === 사회 뉴스 ===
      new Paragraph({ children: [], spacing: { before: 200, after: 0 } }),
      sectionHeader("2부. 사회면 탑기사", RED_DARK),
      ...newsItem(
        1,
        "김건희 여사 항소심 선고 — 도이치모터스 주가조작 혐의 결과 주목",
        "4월 28일 오후 3시 서울고법 형사15-2부는 김건희 여사의 자본시장법 위반(도이치모터스 주가조작), 특정경제범죄 가중처벌법상 알선수재(통일교 금품 수수) 혐의에 대한 항소심 선고 공판을 열었다. 이재명 정부 출범 이후 주목받는 대형 정치 사법 사안으로, 결과에 따라 정국에 적지 않은 파장이 예상됐다.",
        "뉴시스, 연합뉴스 (2026.04.28)"
      ),
      ...newsItem(
        2,
        "산업재해노동자의 날 — 이재명 대통령 삶의 터전이 죽음의 현장 되지 않도록",
        "4월 28일 세계 산업안전보건의 날이자 한국의 산업재해노동자의 날을 맞아 이재명 대통령이 추모 메시지를 발표했다. 대통령은 일터에서 돌아오지 못한 모든 노동자를 마음 깊이 추모하며, 가능한 모든 수단을 동원해 산재를 막겠다고 강조했다. 충무공 이순신 탄신일(양력 4월 28일)이기도 한 이날, 산업재해 통계와 대책이 집중 조명됐다.",
        "캐어유뉴스, 정책브리핑 (2026.04.28)"
      ),
      ...newsItem(
        3,
        "고유가 피해 지원금 1차 신청 — 취약계층 대상 출생연도별 접수",
        "정부가 유가 급등 여파를 겪는 기초생활수급자 등 취약계층을 대상으로 고유가 피해 지원금 1차 신청을 4월 28일 시작했다. 이날은 출생 연도 끝자리가 2나 7인 국민이 신청 대상이다. 중동 분쟁으로 국제 원유가격이 배럴당 100달러를 넘어선 가운데, 서민 생활 지원을 위한 긴급 조치의 하나로 추진됐다.",
        "정부 발표, 연합뉴스 (2026.04.28)"
      ),
      ...newsItem(
        4,
        "MBC-국민의힘 갈등 — 방송 클로징 멘트 피고인 표현 논란",
        "MBC 방송에서 국민의힘 추경호 대구시장 후보를 내란 중요임무 종사 혐의로 기소된 피고인이라고 언급한 클로징 멘트를 두고 국민의힘이 강하게 반발했다. 당은 MBC에 공식 사과를 요구하며 응하지 않을 경우 취재 거부도 불사하겠다고 밝혀 언론-정치권 갈등이 재점화됐다.",
        "각 언론사 보도 (2026.04.28)"
      ),
      ...newsItem(
        5,
        "환경미화원 열악한 장비 실태 — 손수레로 쓰레기 수거하는 현실",
        "일부 지방자치단체에서 위탁 업체 소속 환경미화원들이 현대화된 장비 없이 손수레에 의존해 쓰레기를 수거하는 열악한 현실이 보도됐다. 위탁 업체들은 비용 및 계약 조건을 이유로 장비 개선에 소극적인 것으로 드러났으며, 노동 환경 개선을 촉구하는 목소리가 높아지고 있다.",
        "각 언론사 보도 (2026.04.28)"
      ),

      // === 경제 뉴스 ===
      new Paragraph({ children: [], spacing: { before: 200, after: 0 } }),
      sectionHeader("3부. 경제 뉴스", "375623"),
      ...newsItem(
        1,
        "4월 기업심리지수(CBSI) 소폭 반등 — 중동전쟁에도 제조업 회복세",
        "한국은행이 4월 28일 발표한 기업경기조사에서 전산업 기업심리지수(CBSI)는 94.9로 전월 대비 0.8포인트 상승했다. 제조업 CBSI는 99.1(+2.0p)로 기준선(100)에 근접했으나, 비제조업(92.1, +0.1p)과 내수기업(96.4)은 여전히 부진하다. 경제심리지수(ESI)는 91.7로 2.3p 하락해 가계 소비심리는 오히려 악화됐으며, 원자재 가격 상승이 최대 경영 애로 요인으로 지목됐다.",
        "머니투데이, 한국은행 (2026.04.28)"
      ),
      ...newsItem(
        2,
        "전력기기 3사 52주 신고가 — AI 데이터센터 전력 수요 기대감 반영",
        "국내 전력기기 3사(LS일렉트릭, 효성중공업, HD현대일렉트릭)가 28일 모두 52주 신고가를 경신했다. LS일렉트릭은 전일 대비 12.8% 급등했고, 효성중공업(+10.95%), HD현대일렉트릭(+4.65%)도 강세를 보였다. AI 데이터센터 증설에 따른 전력 수요 급증 기대가 투자 심리를 자극한 결과로 분석된다.",
        "한국거래소, 증권사 리포트 (2026.04.28)"
      ),
      ...newsItem(
        3,
        "한국 1분기 GDP 1.7% 성장 — 반도체 수출 호조가 견인",
        "올해 1분기 한국 경제는 반도체 수출 호조에 힘입어 전년 동기 대비 1.7% 성장한 것으로 집계됐다. 삼성전자가 HBM3E(고대역폭메모리) 공급을 엔비디아에 본격 시작했고, SK하이닉스는 HBM4 대량 생산 일정을 앞당기며 AI 반도체 수요에 적극 대응 중이다. 현대자동차도 1분기 매출 46조 원, 메타플랜트 누적 생산 20만 대를 돌파했다.",
        "한국은행, 각 기업 IR (2026.04.28)"
      ),
      ...newsItem(
        4,
        "중동 분쟁발 유가 리스크 — 국제 원유 배럴당 100달러 돌파",
        "미국과 이란 간 호르무즈 해협 긴장이 고조되면서 국제 유가가 배럴당 100달러 선을 돌파했다. 이에 따라 원자재 수입 비용 상승, 물가 압력 재확대, 소비심리 위축이 우려된다. 정부는 취약계층 고유가 지원금 지급에 나섰으나, 지속될 경우 경상수지와 환율에도 부담이 될 수 있다는 전망이 나온다.",
        "오늘의 클릭, Investing.com (2026.04.28)"
      ),
      ...newsItem(
        5,
        "AI 칩 공급 부족 — 스마트폰 6.9%, PC 8% 가격 인상 예고",
        "AI 인프라 확장에 따른 메모리 칩 공급 부족이 일반 소비자 전자기기 가격 상승으로 이어질 전망이다. 시장조사기관에 따르면 2026년 스마트폰 평균 판매가격은 전년 대비 6.9%, PC는 약 8% 인상될 것으로 예측된다. 저가 브랜드가 직격탄을 맞을 가능성이 높아 소비자 부담이 커질 것으로 우려된다.",
        "newsspace.kr (2026.04.28)"
      ),
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("뉴스브리핑_2026-04-28.docx", buffer);
  console.log("문서 생성 완료: 뉴스브리핑_2026-04-28.docx");
});
