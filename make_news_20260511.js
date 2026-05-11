const {
  Document, Packer, Paragraph, TextRun, ExternalHyperlink,
  AlignmentType, BorderStyle, Header, Footer, PageNumber, UnderlineType
} = require('docx');
const fs = require('fs');

const BLUE      = "1F4E79";
const COLORS = { ai: "1565C0", social: "2E7D32", economy: "BF360C", politics: "4A148C", culture: "00695C" };

function sectionHeader(text, color) {
  return new Paragraph({
    spacing: { before: 400, after: 160 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color, space: 4 } },
    children: [new TextRun({ text, bold: true, size: 32, color, font: "Malgun Gothic" })]
  });
}

function newsItem(num, title, body, source, url) {
  return [
    new Paragraph({
      spacing: { before: 240, after: 60 },
      children: [new TextRun({ text: `${num}. ${title}`, bold: true, size: 22, font: "Malgun Gothic" })]
    }),
    new Paragraph({
      spacing: { before: 0, after: 60 },
      indent: { left: 360 },
      children: [new TextRun({ text: body, size: 20, font: "Malgun Gothic", color: "333333" })]
    }),
    new Paragraph({
      spacing: { before: 0, after: 160 },
      indent: { left: 360 },
      children: [
        new TextRun({ text: "출처: ", size: 18, color: "666666", font: "Malgun Gothic" }),
        new ExternalHyperlink({
          link: url,
          children: [new TextRun({ text: source, size: 18, color: "1565C0", underline: { type: UnderlineType.SINGLE }, font: "Malgun Gothic" })]
        })
      ]
    })
  ];
}

function divider() {
  return new Paragraph({
    spacing: { before: 80, after: 80 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: "DDDDDD", space: 1 } },
    children: []
  });
}

const children = [
  // ── 제목 ──
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 100 },
    children: [new TextRun({ text: "2026년 5월 11일", bold: true, size: 44, color: BLUE, font: "Malgun Gothic" })]
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 80 },
    children: [new TextRun({ text: "일일 뉴스 브리핑", bold: true, size: 52, color: BLUE, font: "Malgun Gothic" })]
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 80 },
    children: [new TextRun({ text: "AI  ·  사회  ·  경제  ·  정치  ·  문화", size: 24, color: "555555", font: "Malgun Gothic" })]
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 400 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: BLUE, space: 4 } },
    children: [new TextRun({ text: "작성일: 2026년 5월 11일 (월)  |  자동 생성 뉴스 다이제스트", size: 18, color: "777777", font: "Malgun Gothic" })]
  }),

  // ── AI 뉴스 ──
  sectionHeader("AI / 인공지능 뉴스", COLORS.ai),
  ...newsItem(1,
    "SI 업계, 1분기 AI·클라우드 바람 타고 호실적 — 수익성 대폭 개선",
    "국내 시스템통합(SI) 업계가 올해 1분기 인공지능(AI)과 클라우드 수요 급증에 힘입어 호실적을 기록했다. CJ올리브네트웍스는 AI 관련 사업 확대로 매출 2,118억원·영업이익 113억원을 달성해 전년 동기 대비 각각 12.1%, 31.4% 늘었다. AI 전환 수요가 SI 전반으로 확산되면서 수익성 대폭 개선이 이어질 것으로 전망된다.",
    "뉴스핌", "https://www.newspim.com/news/view/20260511001205"
  ),
  ...newsItem(2,
    "생성형 AI, 글로벌 일자리 3만 2천 개 감소·AI 인재 수요는 3.2배",
    "생성형 인공지능(AI)이 메타·마이크로소프트·아마존 등 빅테크에서 최근 수 주 사이 3만 2천여 개 일자리를 대체한 것으로 나타났다. 반면 AI 역량을 갖춘 인력에 대한 수요는 공급을 3.2대 1 비율로 앞서고 있으며, 엔트리 레벨 채용 공고의 AI 기술 요건은 지난해 가을 이후 세 배 가까이 늘었다. 일자리 감소와 신수요 창출이 동시에 나타나는 'AI 고용 패러독스'가 심화되고 있다.",
    "전국인력신문", "https://www.kjob.news/news/487771"
  ),
  ...newsItem(3,
    "「2026 AI 프라이버시 민·관 정책협의회」 공식 출범",
    "정부는 인공지능(AI) 서비스 확산에 따른 개인정보 보호 강화를 위해 「2026 AI 프라이버시 민·관 정책협의회」를 공식 출범시켰다. 협의회는 AI 개발·서비스 기업, 학계, 시민단체로 구성되며 AI 관련 개인정보 가이드라인 마련을 주요 과제로 삼는다. 개인정보보호위원회 주도로 분기별 정책 논의를 진행할 예정이다.",
    "대한민국 정책브리핑", "https://www.korea.kr/briefing/pressReleaseView.do?newsId=156742576&call_from=rsslink"
  ),
  ...newsItem(4,
    "EU, 교육 분야 AI 윤리 지침 2026년판 발표 — GDPR·AI법 결합 최초 문서",
    "EU가 교육 분야 AI 사용에 특화된 AI 윤리 지침 2026년판을 발표했다. GDPR(개인정보 보호 규정)과 AI법을 결합한 최초의 문서로, 학습 데이터 처리·학생 개인정보 보호·AI 활용 공정성을 핵심 원칙으로 담았다. 국내 교육부도 이를 참고한 AI 윤리 가이드라인 마련을 검토 중이다.",
    "한국IT산업뉴스", "https://www.koreaiin.com/news/900887"
  ),
  ...newsItem(5,
    "AI 포털 관련주 2026년 5월 주목 — 네오펙트·셀바스AI·라온피플 분석",
    "업스테이지의 '다음' 인수 확정으로 AI 포털 시대가 열리면서 관련 AI 중소형주에 투자자들의 관심이 집중되고 있다. 네오펙트, 셀바스AI, 라온피플 등이 주요 수혜 후보로 꼽히며, 5월 들어 거래량이 눈에 띄게 늘고 있다. 전문가들은 실적 기반이 뒤따라야 한다며 '묻지마 매수' 경계를 권고했다.",
    "21세기 이슈", "https://issue.cyberbabarian.com/ai-portal-stocks-2026-may/"
  ),

  divider(),

  // ── 사회 뉴스 ──
  sectionHeader("사회 주요 뉴스", COLORS.social),
  ...newsItem(1,
    "인천 서구, '서해구'로 명칭 변경 법안 국회 본회의 통과",
    "행정안전부는 인천광역시 '서구(西區)'를 '서해구(西海區)'로 변경하는 법률안이 국회 본회의에서 의결됐다고 밝혔다. 주민들이 오랫동안 요구해 온 사안으로, 인천 서부 지역의 지역 정체성을 강화하기 위한 조치다. 법 시행에 따라 행정 주소와 기관 명칭 등이 단계적으로 바뀐다.",
    "시선뉴스", "https://www.sisunnews.co.kr/news/articleView.html?idxno=238207"
  ),
  ...newsItem(2,
    "2차 고유가 피해지원금 18일부터 지급 — 3600만명 대상",
    "정부가 관계부처 합동 브리핑을 통해 오는 18일부터 국민 3,600만 명을 대상으로 제2차 고유가 피해지원금을 지급한다고 밝혔다. 1차와 지급 기준·신청 절차가 일부 달라졌으므로 해당자는 정부24 또는 주민센터를 통해 반드시 확인해야 한다.",
    "시선뉴스", "https://www.sisunnews.co.kr/news/articleView.html?idxno=238207"
  ),
  ...newsItem(3,
    "인천공항 통합 반대 — 인천시민 한목소리 궐기대회",
    "유정복 인천시장 후보를 비롯한 인천시민들이 인천공항 통합 정책에 반대하는 궐기대회를 개최했다. 시민들은 인천공항 관할권이 인천시에 귀속되어야 한다며 강도 높은 비판과 함께 정책 철회를 촉구했다. 지역 정치권도 이에 가세해 논란이 확산되는 양상이다.",
    "부평포스트", "http://www.bupyeongpost.com/post/47173"
  ),
  ...newsItem(4,
    "알코올 음주 경고 강화 — 국민건강증진법 개정, 11월 시행",
    "보건복지부와 한국건강증진개발원이 국민건강증진법 시행규칙·고시 개정을 완료해 오는 11월 9일부터 강화된 알코올 음주 경고 표시를 시행한다. 개정안은 음주로 인한 건강 위험과 음주운전 등 사회적 폐해 경감을 위해 경고 문구를 대폭 강화했다.",
    "경인방송", "https://news.ifm.kr/news/articleView.html?idxno=470738"
  ),
  ...newsItem(5,
    "5월 11일 입양의 날 — 보건복지부 입양 문화 캠페인",
    "매년 5월 11일은 보건복지부가 2005년 제정한 입양의 날이다. '한 가정(1)이 한 명(1)의 아이를 입양해 새 가정으로 거듭나자'는 의미를 담은 날로, 전국에서 입양 문화 캠페인과 기념 행사가 열렸다. 정부는 입양에 대한 긍정적 인식 확산을 위한 홍보 활동도 병행했다.",
    "이슈에디코", "https://www.issueedico.co.kr/news/article.html?no=37780"
  ),

  divider(),

  // ── 경제 뉴스 ──
  sectionHeader("경제 뉴스", COLORS.economy),
  ...newsItem(1,
    "코스피 7,498 돌파 — 연일 사상 최고치 경신, 7500 목전",
    "코스피 지수가 5월 11일 7,498.00선을 기록하며 사상 최고치를 재경신했다. 반도체 대형주에 대한 높은 수요와 외국인 투자자들의 꾸준한 자금 유입이 상승세를 견인하고 있다. 시장 전문가들은 7,500선 돌파를 앞두고 단기 차익 실현 매물이 출회될 수 있다며 변동성에 주의를 당부했다.",
    "오늘의 클릭", "https://www.cliktoday.com/2026/05/20260511-today-news.html"
  ),
  ...newsItem(2,
    "구윤철 부총리 '올해 성장률 2% 상회' — 중동 변수·반도체 강도 주시",
    "구윤철 경제부총리 겸 기획재정부 장관은 1분기 GDP가 전기 대비 1.7% 성장해 시장 예상을 상회했다며 올해 경제성장률이 2%를 넘을 것으로 전망했다. 다만 중동 전쟁 영향과 반도체 호황 강도에 따라 상회 폭이 달라질 수 있다며 하반기 전략 마련을 예고했다.",
    "뉴스핌", "https://www.newspim.com/news/view/20260511001083"
  ),
  ...newsItem(3,
    "다주택자 양도세 중과 재시행 — 최고 세율 82.5%, 막판 급매 몰려",
    "4년간의 유예가 끝나고 다주택자에 대한 양도소득세 중과 조치가 재시행됐다. 조정대상지역 내 2주택자는 기본세율에 20%p, 3주택 이상 보유자는 30%p가 가산돼 지방소득세 포함 최고 82.5%의 세율이 적용된다. 재시행 직전 막판 급매 물량과 토허구역 신청이 몰리며 부동산 시장에 혼란이 빚어졌다.",
    "이투데이", "https://www.etoday.co.kr/news/view/2583038"
  ),
  ...newsItem(4,
    "외국인 6조 매도 폭탄 속 개인·기관 철벽 방어 — 수급 주도권 변화 신호",
    "외국인 투자자가 최근 6조원 규모의 매도세를 쏟아냈지만, 개인과 기관이 이를 강력히 받아내며 코스피가 7,500선에 바짝 다가섰다. 증시 전문가들은 수급 주도권이 국내 투자자로 이동하는 신호일 수 있다고 분석하면서도 단기 변동성 확대에 주의를 당부했다.",
    "이투데이", "https://www.etoday.co.kr/news/view/2582850"
  ),
  ...newsItem(5,
    "2026년 5월 경제 전망 — 수출 호조 속 내수 회복이 과제",
    "반도체·AI 수출 호조에 힘입어 한국 경제의 외형 성장은 견조하지만, 내수 소비 회복이 더딘 점은 숙제로 꼽힌다. 전문가들은 고금리 장기화와 가계부채 부담이 소비 심리를 억누르고 있어 정책 당국의 균형 잡힌 대응이 필요하다고 지적했다.",
    "KTN 코리아타운뉴스", "https://koreatownnews.com/%EA%B2%BD%EC%A0%9C%EC%A0%84%EB%A7%9D-2026%EB%85%84-5%EC%9B%94-%EA%B2%BD%EC%A0%9C-%ED%98%84%ED%99%A9-%EB%B6%84%EC%84%9D%EA%B3%BC-%EC%A0%84%EB%A7%9D/"
  ),

  divider(),

  // ── 정치 뉴스 ──
  sectionHeader("정치 뉴스", COLORS.politics),
  ...newsItem(1,
    "부산 북갑 보궐선거 — 박민식·한동훈 개소식 600m 맞불 대전",
    "국민의힘 박민식 후보와 무소속 한동훈 후보가 부산 북구갑 보궐선거 개소식을 600m 거리를 두고 동시에 열며 보수 내전을 본격화했다. 박민식 측에는 당 지도부가 총출동했고, 한동훈 측은 시민 지지자를 전면에 내세우는 전략으로 맞서 대조적인 모습을 연출했다.",
    "서울신문", "https://www.seoul.co.kr/news/politics/local-election2026/2026/05/11/20260511003002"
  ),
  ...newsItem(2,
    "한동훈 '박민식 찍으면 장동혁 찍는 것' — 보수 내전 격화",
    "한동훈 후보는 '박민식에게 투표하는 것은 장동혁 대표에게 투표하는 것과 같다'며 국민의힘 지지자들의 전략 투표를 호소했다. 박민식 측은 '한동훈 후원회장을 퇴출 1순위로 검토하겠다'며 강경하게 맞받아쳤다.",
    "인사이트", "https://www.insight.co.kr/news/553514"
  ),
  ...newsItem(3,
    "한동훈, 단일화 질문에 '세상에 절대 안 되는 건 없어'",
    "한동훈 후보는 박민식 측과의 단일화 가능성을 묻는 질문에 '세상에 절대 안 되는 건 없다'며 여지를 남겼다. 그러나 박민식 측은 단일화 가능성이 '제로'라며 일축했고, 국민의힘 당 차원에서도 단일화 논의는 없다는 입장을 고수하고 있다.",
    "프레시안", "https://www.pressian.com/pages/articles/2026051115213172187"
  ),
  ...newsItem(4,
    "국민의힘 박민식에 전폭 지원 — 친한계 징계 가능성 거론",
    "국민의힘 지도부가 부산 북구갑 박민식 후보 지지를 당론으로 확정하고 총력 지원에 나섰다. 당내에서는 한동훈을 지지하는 친한계 의원들에 대한 징계 가능성도 거론되고 있어 당내 갈등이 심화되고 있는 양상이다.",
    "MBC 뉴스", "https://imnews.imbc.com/replay/2026/nwdesk/article/6820169_37004.html"
  ),
  ...newsItem(5,
    "정부, 5월 11일 주요 정책 브리핑 — 2차 지원금·서해구 명칭 변경 등",
    "이재명 정부는 11일 관계부처 합동 정책 브리핑을 통해 제2차 고유가 피해지원금 지급 기준, 인천 서구의 서해구 명칭 변경, 국민건강증진법 시행규칙 개정 등 주요 민생 정책 내용을 발표했다. 정부는 국민 생활 밀착형 정책으로 민심 잡기에 나서고 있다.",
    "시선뉴스", "https://www.sisunnews.co.kr/news/articleView.html?idxno=238207"
  ),

  divider(),

  // ── 문화 뉴스 ──
  sectionHeader("문화 / 엔터테인먼트 뉴스", COLORS.culture),
  ...newsItem(1,
    "BTS·코르티스, 5월 1주 한터차트 정상 — K팝 위력 과시",
    "방탄소년단(BTS)이 한터차트 5월 1주 주간 월드차트 1위에 오르며 건재함을 과시했다. 코르티스(CORTIS)는 두 번째 미니앨범 'GREENGREEN'으로 음반 지수 308만7829점(판매량 231만여 장)을 기록하며 음반 차트 정상을 차지했다. K팝이 글로벌 차트에서 독보적 존재감을 이어가고 있다.",
    "스포츠경향", "https://sports.khan.co.kr/article/202605112159003"
  ),
  ...newsItem(2,
    "BTS 정규 5집 발매·월드투어 한창 — 34개 도시 79회 공연 진행 중",
    "방탄소년단이 정규 5집 발매와 함께 2026~2027 월드투어를 진행 중이다. 고양 개막 공연을 시작으로 로스앤젤레스, 런던, 도쿄 등 전 세계 34개 도시 79회 공연이 예정되어 있다. 완전체 컴백에 전 세계 팬들의 열기가 고조되고 있다.",
    "한터 뉴스", "https://www.hanteonews.com/ko/article/91206"
  ),
  ...newsItem(3,
    "코르티스 'GREENGREEN', 음반 231만 장 돌풍 — 차세대 K팝 강자 부상",
    "신예 K팝 그룹 코르티스(CORTIS)의 두 번째 미니앨범 'GREENGREEN'이 5월 1주 음반 차트를 석권하며 231만 장이 넘는 초동 판매량을 기록했다. 데뷔 1주년을 앞두고 폭발적인 성장세를 보이며 K팝 차세대 강자로 급부상하고 있다.",
    "한터 뉴스", "https://www.hanteonews.com/ko/article/91206"
  ),
  ...newsItem(4,
    "윤가은 감독 '사랑의 세계', 프랑스 매체 호평 — K-독립영화 유럽 진출",
    "윤가은 감독의 코미디 드라마 영화 '사랑의 세계(The World of Love)'가 5월 6일 글로벌 개봉 후 프랑스 현지 매체의 호평을 받으며 주목받고 있다. K-영화의 유럽 독립영화 시장 진출이 확대되는 추세 속에서 고무적인 성과로 평가된다.",
    "Sortiraparis", "https://www.sortiraparis.com/ko/palieseo-hal-il/yeonghwa-silijeu-chugje/guides/325188-2026nyeon-5wol-11il-17il-tv-pyeonseongpyo-ibeon-ju-chucheonjag"
  ),
  ...newsItem(5,
    "이번 주(5월 11~17일) 주목할 TV·OTT 신작 — 드라마·예능 풍성",
    "이번 주 지상파와 OTT에서 다양한 드라마·예능 신작이 방영된다. 로맨틱 판타지 드라마를 비롯해 시청자들의 취향을 겨냥한 다양한 장르가 편성됐다. 연휴 이후 안방극장을 공략하는 콘텐츠가 쏟아지면서 시청자들의 선택 폭이 넓어졌다.",
    "Sortiraparis", "https://www.sortiraparis.com/ko/palieseo-hal-il/yeonghwa-silijeu-chugje/guides/325188-2026nyeon-5wol-11il-17il-tv-pyeonseongpyo-ibeon-ju-chucheonjag"
  ),

  // ── 푸터 ──
  new Paragraph({
    spacing: { before: 400, after: 120 },
    border: { top: { style: BorderStyle.SINGLE, size: 6, color: BLUE, space: 4 } },
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "※ 본 브리핑은 Claude Code 자동화 시스템이 2026년 5월 11일 공개된 정보를 기반으로 작성했습니다.", size: 16, color: "888888", font: "Malgun Gothic" })]
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "각 기사의 상세 내용은 출처 링크를 통해 확인하시기 바랍니다.", size: 16, color: "888888", font: "Malgun Gothic" })]
  })
];

const doc = new Document({
  styles: {
    default: { document: { run: { font: "Malgun Gothic", size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Malgun Gothic" },
        paragraph: { spacing: { before: 400, after: 160 }, outlineLevel: 0 } }
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
          children: [new TextRun({ text: "일일 뉴스 브리핑 | 2026.05.11", font: "Malgun Gothic", size: 18, color: "888888" })]
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
            new TextRun({ text: " -", font: "Malgun Gothic", size: 18, color: "888888" })
          ]
        })]
      })
    },
    children
  }]
});

const path = require('path');
const OUT = path.join(__dirname, '일간뉴스_2026-05-11.docx');

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(OUT, buf);
  console.log("생성 완료:", OUT);
}).catch(err => { console.error(err); process.exit(1); });
