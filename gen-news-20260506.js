const {
  Document, Packer, Paragraph, TextRun, Header, Footer,
  AlignmentType, HeadingLevel, BorderStyle, PageNumber
} = require('docx');
const fs = require('fs');
const path = require('path');

// Use __dirname to safely handle Korean path characters
const OUTPUT_DIR  = path.join(__dirname, '뉴스브리핑 모음');
const OUTPUT_FILE = path.join(OUTPUT_DIR, '일일뉴스브리핑_20260506.docx');

// Ensure output directory exists
if (!fs.existsSync(OUTPUT_DIR)) fs.mkdirSync(OUTPUT_DIR, { recursive: true });

/* ── Helpers ─────────────────────────────────────────── */
function sectionHeader(emoji, title) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 440, after: 180 },
    border: {
      bottom: { style: BorderStyle.SINGLE, size: 8, color: '2E4D7B', space: 4 }
    },
    children: [
      new TextRun({ text: emoji + '  ' + title, font: 'Malgun Gothic', size: 34, bold: true, color: '2E4D7B' })
    ]
  });
}

function newsItem(num, headline, body) {
  return [
    new Paragraph({
      spacing: { before: 220, after: 60 },
      children: [
        new TextRun({ text: num + '.  ', font: 'Malgun Gothic', size: 22, bold: true, color: '1F497D' }),
        new TextRun({ text: headline,    font: 'Malgun Gothic', size: 22, bold: true, color: '1F497D' })
      ]
    }),
    new Paragraph({
      spacing: { before: 0, after: 180 },
      indent:  { left: 380 },
      children: [new TextRun({ text: body, font: 'Malgun Gothic', size: 20, color: '3C3C3C' })]
    })
  ];
}

/* ── Content ─────────────────────────────────────────── */
const DATE_KO   = '2026년 5월 6일 (수요일)';
const DATE_YMDH = '2026.05.06';

const children = [

  /* Title block */
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 100 },
    children: [new TextRun({ text: '데일리 뉴스 브리핑', font: 'Malgun Gothic', size: 56, bold: true, color: '1A1A2E' })]
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 100 },
    children: [new TextRun({ text: DATE_KO, font: 'Malgun Gothic', size: 28, color: '555555' })]
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 480 },
    border: { bottom: { style: BorderStyle.DOUBLE, size: 6, color: '2E4D7B', space: 4 } },
    children: [new TextRun({ text: 'AI  ·  사회  ·  경제  ·  정치  ·  문화', font: 'Malgun Gothic', size: 22, color: '666666', italics: true })]
  }),

  /* ── Section 1: AI ── */
  sectionHeader('🤖', 'AI / 인공지능 뉴스'),

  ...newsItem('1',
    'AI EXPO KOREA 2026 개막 — 제9회 국제인공지능대전',
    '아시아 최대 규모의 AI 행사인 제9회 국제인공지능산업대전(AI EXPO KOREA 2026)이 5월 6일부터 8일까지 서울 코엑스 A홀에서 개막된다. 약 350개사 600부스 규모로 피지컈 AI와 AI 에이전트 등 최신 트렌드를 총망라하며, 참관객 5만 명을 목표로 한다.'),

  ...newsItem('2',
    '이재명 대통령, 구글 딥마인드 CEO 하사비스 접견 — AI 안전장치 MOU',
    '이재명 대통령이 구글 딥마인드의 데미스 하사비스 대표와 면담했다. 하사비스는 "엤내 5년 안에 범용 AI가 현실화될 것"이라면서 AI 통제를 위한 가드레일의 필요성을 강조했다. 양측은 AI 공동연구와 인재양성 MOU를 체결하고 서울 AI 캔퍼스 설립도 합의했다.'),

  ...newsItem('3',
    'AMD AI 칩 호실적 발표 — 코스피 7000선 돌파 기대감 고조',
    'AMD의 견실한 AI 칩 실적 발표에 가랍리의 글로벌 위험자산 선호 심리가 강화되면서 코스피가 7000선 돌파를 시도하고 있다. 니용스 증시 코스피는 AI 반도체 관련주를 중심으로 상승하며 외국인 매수세가 지속되고 있다.'),

  ...newsItem('4',
    '딕시크(DeepSeek) V4-Pro 가격 75% 인하 — AI 모델 가격 경쟁 재점화',
    '딕시크가 V4-Pro 모델을 5월 31일까지 75% 할인 공급한다고 발표했다. 입력 비용을 100만 토큰당 0.0036달러로 낙춰 GPT-5.5 대비 97% 저렴한 파격가를 제시하면서 글로벌 AI 모델 가격 경쟁이 다시 공겟화됐다.'),

  ...newsItem('5',
    'AI 에이전트 · 피지컈 AI 시대 본격화',
    '2026년이 AI 에이전트의 도약 원년으로 자리매김하는 가운데, 스스로 계획·실행·협업하는 에이전튱 AI가 업무 현장에 깊숭이 침투하고 있다. 디지털 공간을 넘어 기계를 매개로 인간과 직접 상호작용하는 피지컈 AI 시대도 본격 개막 중이다.'),

  /* ── Section 2: Society ── */
  sectionHeader('📰', '사회 뉴스'),

  ...newsItem('1',
    '서울고법 신종오 판사 변사, 경찰 수사 착수',
    '김건희 여사의 도이치모터스 주가조작 혜의 항소심에서 유죄를 선고한 신종오 서울고법 판사가 6일 서울고법에서 숫진 채 발견됐다. 경찰은 타살 정황은 없다고 밝히며 정확한 사인을 조사 중이다.'),

  ...newsItem('2',
    '삼성전자 노조, 영업이익 15% 성과급 요구하며 파업 예고',
    '삼성전자 노동조합이 영업이익의 15% 성과급 지급을 요구하며 파업을 예고했다. 1인당 약 6억 원에 달할 것으로 추산되는 요구안에 이사회 의장은 "기업 경쟁력 훼손, 국가 경제 악영향"라고 우려하며 노사 갈등 해법을 촉구하고 있다.'),

  ...newsItem('3',
    '트럼프, 호르무즈 해협 항행 자유 보장 해상작전 선언 — 동맹국 파병 요구',
    '도널드 트럼프 미국 대통령이 호르무즈 해협의 항행 자유를 보장하기 위한 대규모 해상작전을 선언하고 동맹국에 파병을 요구했다. 에너지 의존도가 높은 한국에 대한 에너지 공급망 위협이 근시화되며, 정부가 대응 방안을 발빠 마련하고 있다.'),

  ...newsItem('4',
    '청와대, 트럼프 \'프리덤 프로젝트\' 참여 제안 검토 중',
    '청와대가 트럼프 행정부가 제안한 \'프리덤 프로젝트\' 참여 여부에 대해 "한반도 태세와 국내법을 감안한 검토 중"이라고 밝혔다. 외교부관방 학계에서는 한미 동맹 관계에 미칠 엁향을 면밀히 분석해야 한다는 목소리도 나오고 있다.'),

  ...newsItem('5',
    '어린이날 계기 \'노키즈존 철폐\' 목소리 확산',
    '어린이날을 \'어린이 차별 철폐의 날\'로 지정하자는 제안이 제기됐다. 전국 노키즈존의 60% 이상이 서울·경기·제주에 집중되어 있으며, 쿨페·식당이 90% 이상을 차지하는 것으로 나타나 개선 요구가 가속화되고 있다.'),

  /* ── Section 3: Economy ── */
  sectionHeader('💰', '경제 뉴스'),

  ...newsItem('1',
    '코스피 7000선 돌파 시도 — AMD 호실적 · 외국인 매수세 강화',
    '바이론증시의 사상 최고치 경신과 AMD의 AI 칩 호실적 발표에 글로벌 위험자산 선호 심리가 강화되면서 코스피가 7000선 돌파를 시도하고 있다. 삼성증권 등은 코스피 상단을 8400까지 제시하는 낙관적 전망을 내놓고 있다.'),

  ...newsItem('2',
    '원·달러 환율 고환율 지속 — 에너지 의존·외국인 매도세 원인 분석',
    '중동 전쟁 이후 원·달러 환율이 다른 국가보다 큰 폭으로 오른 원인으로 한국의 높은 에너지 의존도와 외국인 투자자의 주식 매도세가 분석됐다. 전문가들은 호르무즈 문제가 장기화될 경우 환율 불안이 가중될 수 있다고 경고했다.'),

  ...newsItem('3',
    '4월 소비자물가 1년 9개월 만에 최대 상승 — 금리 인상론 부상',
    '4월 소비자물가지수가 119.37로 1년 9개월 만에 가장 큰 폭으로 올랐다. 유상대 한국은행 부옂재가 "금리 인상을 고민할 때가 됐다"고 공개 언급하면서 기준금리 인상 논의가 본격화될 전망이다.'),

  ...newsItem('4',
    '부동산 지역 양극화 심화 — 서울 둥세 둥화, 지방 미분양 적체',
    '고금리 장기화와 대출 규제 강화 영향으로 서울 아파트값 상승폭은 둔화된 반면 지방에서는 미분양 물량이 쌓이며 하락세가 지속되는 양극화 구도가 심화되고 있다. 전문가들은 지방 주택 시장 안정을 위한 선제적 지원설 마련을 촉구하고 있다.'),

  ...newsItem('5',
    '한국 1분기 GDP 1.7% 꺜짝 성장, JP모건 등 해외 IB 전망치 상향',
    '한국 2026년 1분기 GDP가 전분기 대비 1.7% 성장해 예상치(0.9%)를 크게 웃돈다. JP모건 등 해외 주요 투자은행들이 성장률 전망치를 상향 조정한 가운데, 중동 전쟁 영향이 2분기부터 본격화될 수 있어 선제적 재정 운용이 중요하다고 강조했다.'),

  /* ── Section 4: Politics ── */
  sectionHeader('🏛️', '정치 뉴스'),

  ...newsItem('1',
    '이재명 대통령, 국무회의 겸 비상경제점검회의 주재 — 중동 리스크 대응 논의',
    '이재명 대통령이 6일 오전 청와대에서 제20회 국무회의 겸 제7차 비상경제점검회의를 주재했다. 중동 리스크 대응과 에너지·물가 안정 방안이 주요 의제에 포함됐다.'),

  ...newsItem('2',
    '민주당 22대 제3기 원내대표 선출 — 한병도 단독 출마 사실상 확정',
    '민주당이 6일 오후 의원총회를 열어 22대 제3기 원내대표를 선출한다. 한병도 의원이 단독으로 출마해 사실상 확정됐다. 정청래 당대표 주재로 열린 최고위원회의도 오전에 진행된다.'),

  ...newsItem('3',
    '22대 후반기 국회의장 후보 출사표 — 조정식·박지원·김태년 경약',
    '조정식, 박지원, 김태년 의원이 22대 후반기 국회의장 선도 경쟁에 각각 출사표를 던졌다. 민주당 내부에서는 의장 선임을 두고 안배 경쟁이 본격화되는 양상이다.'),

  ...newsItem('4',
    '국민의힐, 6.3 지방선거 경기도 필승결의대회 — 장동혜 대표 총력전 선언',
    '장동혜 국민의힐 대표가 췄언석 원내대표와 함께 6.3 지방선거 경기도 필승결의대회에 참석해 총력전을 선언했다. 오후에는 의원총회를 열어 6.3 선거 대막희 논의를 진행할 예정이다.'),

  ...newsItem('5',
    '이준석·천하람 등 야권 원내대표 주요 방송 출연 — 지방선거 공세 예고',
    '개혁신당 이준석 당대표와 천하람 원내대표가 주요 방송 프로그램에 동시 출연해 6.3 지방선거 공세 전략을 내놀았다. 조국혁신당 서왕진 원내대표도 KBS 인터븷 등에 출연해 입장을 명햙했다.'),

  /* ── Section 5: Culture ── */
  sectionHeader('🎭', '문화 뉴스'),

  ...newsItem('1',
    '2026 한국 영화 재도약 라인업 공개 — 정주리 감독 신작 등',
    '불황을 뚫고 재도약을 노리는 2026 한국 영화 라인업이 공개됐다. 정주리 감독의 신작 <도라>를 비롯해 다양한 장르의 화제작들이 도전장을 보이며, K영화의 부활을 알리고 있다.'),

  ...newsItem('2',
    '띌지컹 성수기 5월 — <6시 퇴근>·<오즈>·<브라더스 카라마조프> 동시 개막',
    '5월은 러븐스 띌지컹의 성수기로, <6시 퇴근>, <오즈>, <브라더스 카라마조프>, <극낙간에서> 등 익숫한 작품들이 관객을 만나 화제를 이뢨거나 기대를 모으고 있다.'),

  ...newsItem('3',
    '라이선스 신작 듀지컹 <V.L> 5월 개막 — 영국 고교 배경 청춥 성장기',
    '글림아티스트가 5월부터 링크더스페이스 2관에서 라이선스 신작 <V.L>을 선보인다. 영국 고등학교를 배경으로 10대 소년의 고민과 성장을 그린 작품으로, 영상 이미지와 라이브 퍼포먼스가 조화를 이론 주목받고 있다.'),

  ...newsItem('4',
    '예술의전당 5월 공연 — 오브르뉴론알프 국립 오케스트라 내한 연주',
    '토마스 첼헤트마이어가 이끄는 프랑스 오브르뉴론알프 국립 오케스트라가 첨리스트 양성원과 함께 예술의전당 스테이지에 오른다. 깊이 있는 프랑스 음악을 매평가들로부터 호평을 받으며 클래식 애호가들의 시선을 실천시계 수집하고 있다.'),

  ...newsItem('5',
    '아이유 1억 원·한지민 5청만 원 어린이날 기부 — 니년간 나눈 문화',
    '가수 겸 배우 아이유가 어린이날을 맞아 소외 어린이를 위한 1억 원을 기부했다. 배우 한지민도 19년 연속으로 어린이날 기부 활동을 이어가며 5청만 원을 나눐 선한 영향력을 포발했다.'),

  /* Footer */
  new Paragraph({
    spacing: { before: 480, after: 120 },
    border: { top: { style: BorderStyle.SINGLE, size: 4, color: '999999', space: 4 } },
    alignment: AlignmentType.CENTER,
    children: [new TextRun({
      text: '본 뉴스 브리핑은 2026년 5월 6일(수) 기준으로 자동 수집·정리된 자료입니다. 각 기사의 세부 내용은 원문 출처를 확인하시기 바랍니다.',
      font: 'Malgun Gothic', size: 18, color: '888888', italics: true
    })]
  }),
];

/* ── Document ─────────────────────────────────────────── */
const doc = new Document({
  styles: {
    default: { document: { run: { font: 'Malgun Gothic', size: 22 } } },
    paragraphStyles: [{
      id: 'Heading1', name: 'Heading 1',
      basedOn: 'Normal', next: 'Normal', quickFormat: true,
      run: { size: 34, bold: true, font: 'Malgun Gothic', color: '2E4D7B' },
      paragraph: { spacing: { before: 440, after: 180 }, outlineLevel: 0 }
    }]
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
          border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC', space: 2 } },
          children: [new TextRun({ text: '데일리 뉴스 브리핑  |  ' + DATE_YMDH, font: 'Malgun Gothic', size: 18, color: '888888' })]
        })]
      })
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          border: { top: { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC', space: 2 } },
          children: [
            new TextRun({ text: '—  ', font: 'Malgun Gothic', size: 18, color: '888888' }),
            new TextRun({ children: [PageNumber.CURRENT], font: 'Malgun Gothic', size: 18, color: '888888' }),
            new TextRun({ text: '  —', font: 'Malgun Gothic', size: 18, color: '888888' }),
          ]
        })]
      })
    },
    children
  }]
});

Packer.toBuffer(doc)
  .then(buffer => {
    fs.writeFileSync(OUTPUT_FILE, buffer);
    console.log('SUCCESS: ' + OUTPUT_FILE);
  })
  .catch(err => {
    console.error('ERROR:', err.message);
    process.exit(1);
  });
