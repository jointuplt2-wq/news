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
      sectionHeader("🤖  1부. AI & 인공지능 뉴스", BLUE),
      ...newsItem(
        1,
        "OpenAI, GPT-5.5 공식 출시… 코딩·데이터 분석·에이전틱 작업 대폭 향상",
        "OpenAI가 4월 24일 API에서 GPT-5.5와 GPT-5.5 Pro를 공개하고 ChatGPT Plus, Pro, Business, Enterprise 사용자에게 순차 배포를 시작했습니다. 새 모델은 코드 작성·디버깅, 온라인 리서치, 데이터 분석, 문서 작성, 소프트웨어 조작 등 복합 작업을 도구를 스스로 활용하며 완수하는 에이전틱 능력이 이전 세대보다 크게 향상되었습니다. OpenAI는 GPT-5.5를 '역대 가장 직관적인 모델'로 소개했습니다.",
        "OpenAI 공식 발표, TechCrunch (2026.04.23~24)"
      ),
      ...newsItem(
        2,
        "ChatGPT Images 2.0 출시… 사고(Thinking) 기반 이미지 생성·2K 해상도 지원",
        "OpenAI가 4월 21일 gpt-image-2 모델 기반 ChatGPT Images 2.0을 공개했습니다. 이번 업데이트에서는 추론(thinking) 능력을 활용한 이미지 생성, 다국어 텍스트 완벽 렌더링, API를 통한 최대 2K 해상도 출력, 단일 프롬프트로 최대 8장의 일관된 이미지 생성이 가능해졌습니다. 이는 OpenAI의 역대 가장 큰 이미지 생성 업그레이드로 평가받고 있습니다.",
        "OpenAI 공식 발표, VidMuse (2026.04.21)"
      ),
      ...newsItem(
        3,
        "ChatGPT for Clinicians 출시… 미국 의료진 대상 무료 AI 서비스",
        "OpenAI가 미국 내 인증된 임상의를 대상으로 실제 진료 현장에서 활용할 수 있는 'ChatGPT for Clinicians' 무료 서비스를 출시했습니다. 의료 기록 검토, 임상 의사결정 지원, 환자 커뮤니케이션 등 실질적인 의료 업무를 지원하는 것이 목표입니다. 의료 AI 시장 공략을 본격화하는 OpenAI의 전략적 행보로 주목받고 있습니다.",
        "OpenAI 공식 발표, AI Dispatch (2026.04.27)"
      ),
      ...newsItem(
        4,
        "Google Gemini 3.1 Pro, 추론 벤치마크 선두… GPQA Diamond 94.3% 달성",
        "Google의 Gemini 3.1 Pro가 과학·논리 추론 표준 평가인 GPQA Diamond에서 94.3%를 기록하며 현재 출시된 모델 중 최고 성능을 달성했습니다. 출력 가격은 100만 토큰당 2달러로 경쟁사 대비 최저 수준을 유지하며 비용 효율도 두드러집니다. AI 에이전트 기능 면에서도 Anthropic과 Google이 치열한 경쟁을 벌이고 있다는 분석이 나왔습니다.",
        "LLM Stats, AI Model Releases (2026.04)"
      ),
      ...newsItem(
        5,
        "중기부, 2026년 AI 예산 약 8천억 원 투입… 고도화·내실화 중심으로 전환",
        "중소벤처기업부가 2026년 인공지능 관련 예산으로 약 8천억 원을 편성하고, 단순 보급 확산을 넘어 AI 서비스의 고도화와 내실화에 집중하는 방향으로 전략을 전환했습니다. 또한 '2026 AI 프라이버시 민관 정책협의회'가 공식 출범해 AI 시대 개인정보 보호 체계를 강화하기 위한 민관 협력을 본격화했습니다.",
        "대한민국 정책브리핑, 인공지능신문 (2026.04)"
      ),

      // === 사회 뉴스 ===
      new Paragraph({ children: [], spacing: { before: 200, after: 0 } }),
      sectionHeader("🏛  2부. 사회면 탑기사", "7B2C2C"),
      ...newsItem(
        1,
        "완도 냉동창고 화재 시공업체 대표 구속… 소방관 2명 순직",
        "소방관 2명이 순직한 전남 완도 냉동창고 화재를 유발한 시공업체 대표가 경찰에 구속됐습니다. 경찰은 불량 시공과 안전 규정 위반이 화재 원인으로 판단하고 있으며, 업무상 과실치사 혐의를 적용했습니다. 순직 소방관들의 처우 개선과 안전 규정 강화를 촉구하는 여론이 높아지고 있습니다.",
        "다음뉴스, 경향신문 (2026.04.28)"
      ),
      ...newsItem(
        2,
        "20대 여성, 수면제 먹여 남성 4명 4890만 원 강취… 경찰 검거",
        "지난해 12월부터 올해 4월까지 수도권 일대에서 남성 4명에게 수면제를 먹여 잠들게 한 뒤 계좌이체 등으로 약 4890만 원을 빼앗은 20대 여성이 경찰에 검거됐습니다. 피의자는 데이팅 앱을 통해 피해자를 유인한 것으로 알려졌으며, 경찰은 여죄를 추가 수사 중입니다.",
        "영남경제TV, 다음뉴스 (2026.04.28)"
      ),
      ...newsItem(
        3,
        "고유가 피해 지원금 신청 이틀째… 오늘은 출생연도 끝자리 '2·7' 대상",
        "기초생활수급자 등 취약계층을 위한 '고유가 피해 지원금' 1차 신청이 이틀째 진행 중입니다. 28일은 출생연도 끝자리가 '2' 또는 '7'인 국민이 신청 대상이며, 정부는 국제유가 상승으로 어려움을 겪는 저소득층의 에너지 부담을 줄이기 위해 해당 지원금을 마련했습니다.",
        "대한민국 정책브리핑, 영남경제TV (2026.04.28)"
      ),
      ...newsItem(
        4,
        "국민의힘, MBC '클로징 멘트' 선거개입 주장… 법적 조치 예고",
        "국민의힘이 MBC의 뉴스 클로징 멘트를 '선거개입 행위'로 규정하고 법적·행정적 모든 수단을 동원해 책임을 묻겠다고 밝혔습니다. 방송심의위원회 제소와 함께 민형사상 고소도 검토 중인 것으로 알려졌습니다. MBC 측은 편집권과 표현의 자유를 침해하는 정치적 압박이라고 반박하며 맞섰습니다.",
        "서울신문, 경향신문 (2026.04.28)"
      ),
      ...newsItem(
        5,
        "온라인 그루밍 성착취 피해 아동·청소년 287명… 전수 수사 착수",
        "2025년 1월부터 2026년 2월까지 온라인 그루밍을 통해 성착취 피해를 입은 아동·청소년이 287명에 달하는 것으로 집계됐습니다. 경찰청은 디지털 성범죄 특별수사팀을 가동해 전수 수사에 착수했으며, 피해자 지원과 재발 방지를 위한 법 개정도 추진 중입니다.",
        "다음뉴스, 대한민국 정책브리핑 (2026.04.28)"
      ),

      // === 경제 뉴스 ===
      new Paragraph({ children: [], spacing: { before: 200, after: 0 } }),
      sectionHeader("📈  3부. 경제 뉴스", DARK_GREEN),
      ...newsItem(
        1,
        "기업심리지수(CBSI) 94.9… 21개월 만에 최고치, 제조업 2포인트 상승",
        "한국은행이 28일 발표한 4월 기업경기조사에 따르면 전산업 기업심리지수(CBSI)는 94.9로 전월 대비 0.8p 상승하며 2024년 7월 이후 21개월 만에 최고 수준을 기록했습니다. 제조업 CBSI는 99.1로 2.0p 올랐으며, 재고 감소와 업황 개선이 상승을 이끌었습니다. 다만 지수가 여전히 장기 평균(100)을 밑돌아 체감 경기 회복은 아직 미흡하다는 평가입니다.",
        "한국은행, 머니투데이, 파이낸셜뉴스 (2026.04.28)"
      ),
      ...newsItem(
        2,
        "한국은행 기준금리 2.5% 동결… 중동 리스크·유가 상방압력 주시",
        "한국은행은 4월 금융통화위원회에서 기준금리를 연 2.50%로 동결했습니다. 미이란 갈등에 따른 국제유가 상승이 물가 상방압력을 높이는 가운데, 통화당국은 내수 회복세와 물가 안정을 동시에 고려해 금리를 현 수준에서 유지했습니다. 시장에서는 하반기 인하 재개 가능성을 열어두고 있습니다.",
        "한국은행, 토스뱅크 (2026.04.28)"
      ),
      ...newsItem(
        3,
        "원·달러 환율 1,470원대 유지… 미이란 휴전 이후 안정, 유가 변수 경계",
        "원·달러 환율은 미국과 이란의 2주 휴전 합의 이후 1,470원대 초중반에서 안정세를 보이고 있습니다. 다만 호르무즈 해협 봉쇄 우려로 브렌트유가 배럴당 105달러를 상회하는 등 에너지 가격 변동성이 환율에 추가 상방 압력을 줄 수 있다는 전망이 이어지고 있습니다.",
        "무역협회 환율종합, 나무위키 (2026.04.28)"
      ),
      ...newsItem(
        4,
        "소비자심리지수(CCSI) 99.2… 전월 대비 7.8p 하락, 내수 냉각 우려",
        "4월 소비자심리지수(CCSI)는 99.2로 전월보다 7.8포인트 하락했습니다. 중동 전쟁에 따른 고유가, 물가 불안, 고금리 장기화 우려 등이 복합적으로 작용해 소비 심리를 위축시킨 것으로 분석됩니다. 기업 체감경기 개선과 소비자 심리 악화가 엇갈리는 온도 차가 뚜렷한 상황입니다.",
        "한국은행, 머니투데이 (2026.04.28)"
      ),
      ...newsItem(
        5,
        "1분기 GDP 1.7% 성장… 반도체·건설·소비 '3중 견인', 2분기는 중동 리스크 변수",
        "2026년 1분기 실질 국내총생산(GDP)이 전분기 대비 1.7% 성장하며 시장 예상치를 웃돌았습니다. 반도체 수출 호조, 주택 건설 회복, 민간 소비 증가가 동시에 성장을 뒷받침했습니다. 그러나 2분기에는 중동 분쟁에 따른 유가 급등과 글로벌 공급망 불안이 성장세 둔화 요인으로 작용할 수 있다는 전망이 제기되고 있습니다.",
        "21세기이슈, 뉴스핌 (2026.04.28)"
      ),
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("뉴스브리핑_2026-04-28.docx", buffer);
  console.log("문서 생성 완료: 뉴스브리핑_2026-04-28.docx");
});
