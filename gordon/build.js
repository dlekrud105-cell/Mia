const pptxgen = require('pptxgenjs');
const React = require('react');
const RDS = require('react-dom/server');
const sharp = require('sharp');
const fa = require('react-icons/fa');
const { applyTheme } = require('/root/.claude/skills/synced/eb7b2f3f-445a-4846-acbc-22136f2f01b5_65af6b0e-d336-4b83-bddb-739ee7f577d1/pptx/scripts/apply_theme.js');

const THEME = { name: 'Gordon Eucalyptus', headFontFace: 'Malgun Gothic', bodyFontFace: 'Malgun Gothic',
  colors: { dk1: '1C2B24', lt1: 'FFFFFF', dk2: '1F4A3A', lt2: 'EEF3EF',
    accent1: '1F4A3A', accent2: 'C8964E', accent3: '6E9C84', accent4: '8A5A3C', accent5: '4F6B8A', accent6: 'B5C9BC',
    hlink: '1F4A3A', folHlink: '6E9C84' } };
const HEX = THEME.colors;

async function icon(Comp, color, size = 256) {
  const svg = RDS.renderToStaticMarkup(React.createElement(Comp, { color: '#' + color, size }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return 'image/png;base64,' + buf.toString('base64');
}

(async () => {
  const pres = new pptxgen();
  pres.layout = 'LAYOUT_WIDE'; // 13.33 x 7.5
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  pres.title = 'Gordon NSW 부동산 분석';
  const C = pres.SchemeColor;
  const W = 13.333;

  pres.defineSlideMaster({ title: 'DARK', background: { color: C.text2 },
    objects: [{ placeholder: { options: { name: 'title', type: 'title', x: 0.8, y: 2.2, w: 11.7, h: 1.6, fontSize: 44, bold: true, color: C.background1, valign: 'bottom', align: 'left' }, text: '' } },
              { placeholder: { options: { name: 'body', type: 'body', x: 0.8, y: 4.0, w: 11.7, h: 1.2, fontSize: 20, color: C.accent6, valign: 'top', align: 'left' }, text: '' } }] });
  pres.defineSlideMaster({ title: 'CONTENT', background: { color: C.background1 },
    objects: [{ placeholder: { options: { name: 'title', type: 'title', x: 0.6, y: 0.4, w: 12.1, h: 0.9, fontSize: 32, bold: true, color: C.text2, valign: 'middle', align: 'left' }, text: '' } },
              { text: { text: 'Gordon NSW 2072 · 부동산 리서치', options: { x: 0.6, y: 7.0, w: 6, h: 0.3, fontSize: 10, color: C.accent3 } } }],
    slideNumber: { x: 12.2, y: 7.0, w: 0.6, h: 0.3, fontSize: 10, color: C.accent3, align: 'right' } });

  const I = {};
  for (const [k, comp, col] of [['train', fa.FaTrain, HEX.lt1], ['school', fa.FaGraduationCap, HEX.lt1], ['tree', fa.FaTree, HEX.lt1],
      ['home', fa.FaHome, HEX.lt1], ['landmark', fa.FaLandmark, HEX.lt1], ['city', fa.FaCity, HEX.lt1], ['shop', fa.FaShoppingBag, HEX.lt1],
      ['book', fa.FaBookOpen, HEX.lt1], ['gavel', fa.FaGavel, HEX.lt1], ['shield', fa.FaShieldAlt, HEX.lt1], ['up', fa.FaArrowUp, HEX.lt1], ['warn', fa.FaExclamationTriangle, HEX.lt1]]) {
    I[k] = await icon(comp, col);
  }
  const iconCircle = (s, key, x, y, d = 0.7, fill = C.accent1) => {
    s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { type: 'none' }, objectName: 'icon-bg-' + key });
    s.addImage({ data: I[key], x: x + d * 0.25, y: y + d * 0.25, w: d * 0.5, h: d * 0.5, objectName: 'icon-' + key, altText: key });
  };
  const T = (s, text, o) => s.addText(text, Object.assign({ isTextBox: true, margin: 0, color: C.text1, fontSize: 14, valign: 'top' }, o));
  const card = (s, x, y, w, h, fill = C.background2) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.12, fill: { color: fill }, line: { type: 'none' } });

  // 1. Title
  pres.addSection({ title: '개요' });
  let s = pres.addSlide({ masterName: 'DARK', sectionTitle: '개요' });
  s.addText('고든(Gordon), 왜 노스쇼어의 부촌인가', { placeholder: 'title' });
  s.addText('입지 · 역사 · Ku-ring-gai Council과의 관계 · 향후 개발 계획\n2026년 11월 기준 리서치', { placeholder: 'body' });
  iconCircle(s, 'tree', 0.8, 1.0, 0.9, C.accent2);
  s.addNotes('표지. 시드니 어퍼 노스쇼어의 고든 지역을 부동산 관점에서 정리한 자료입니다.');

  // 2. Summary
  s = pres.addSlide({ masterName: 'CONTENT', sectionTitle: '개요' });
  s.addText('핵심 요약: 희소성이 만든 프리미엄, 이제 고밀도 전환기', { placeholder: 'title' });
  const sum = [
    ['train', '입지', '1890년 노스쇼어 철도 개통 이후 형성된 기차역 중심 주거지. Pacific Hwy와 T1 노선이 교차하는 교통 거점'],
    ['home', '부촌 형성', '대형 필지 정원주택, 명문 학군, 숲. 100년 넘게 저밀도가 유지되며 희소가치 형성'],
    ['landmark', '카운슬 관계', 'Ku-ring-gai Council 청사가 고든에 위치. 문화유산 보존·수목 정책으로 동네 성격을 지켜옴'],
    ['city', '향후 계획', '2025.11 확정된 카운슬 대안안. 고든에 9,012세대 배분, 핵심부지 최고 28층'] ];
  sum.forEach((r, i) => { const x = 0.6 + i * 3.08; card(s, x, 1.7, 2.85, 4.3);
    iconCircle(s, r[0], x + 0.3, 2.0);
    T(s, r[1], { x: x + 0.3, y: 2.95, w: 2.3, h: 0.5, fontSize: 20, bold: true, color: C.text2 });
    T(s, r[2], { x: x + 0.3, y: 3.6, w: 2.3, h: 2.2, fontSize: 16 }); });

  // 3. At a glance - stats
  s = pres.addSlide({ masterName: 'CONTENT', sectionTitle: '개요' });
  s.addText('한눈에 보는 고든: 단독주택 중위 $3.4–3.8M', { placeholder: 'title' });
  const stats = [['$3.4–3.8M', '단독주택 중위가격\n(OnTheHouse · Microburbs · AusPropertyInsights, 2026)'],
                 ['~$0.95M', '유닛(아파트) 중위가격\n단독주택의 약 1/4 수준'],
                 ['$1,500/주', '단독주택 중위 임대료\n(OnTheHouse)'],
                 ['+0.1%', '최근 12개월 주택가격 변동\n(Microburbs, 70건 거래 기준)']];
  stats.forEach((r, i) => { const x = 0.6 + (i % 2) * 3.6, y = 1.6 + Math.floor(i / 2) * 2.55;
    card(s, x, y, 3.4, 2.35);
    T(s, r[0], { x: x + 0.25, y: y + 0.25, w: 3.0, h: 0.9, fontSize: 32, bold: true, color: i === 3 ? C.accent2 : C.text2 });
    T(s, r[1], { x: x + 0.25, y: y + 1.2, w: 3.0, h: 1.0, fontSize: 12, color: C.text1 }); });
  card(s, 8.0, 1.6, 4.7, 4.9, C.text2);
  T(s, '위치 개요', { x: 8.35, y: 1.9, w: 4.0, h: 0.5, fontSize: 20, bold: true, color: C.background1 });
  T(s, [
    { text: '시드니 CBD 북쪽 약 15km, Ku-ring-gai LGA', options: { bullet: true, breakLine: true } },
    { text: 'North Shore 철도(T1) 고든역 + Pacific Hwy 상업가', options: { bullet: true, breakLine: true } },
    { text: '동쪽 Garigal, 서쪽 Lane Cove 국립공원 사이의 숲 지대', options: { bullet: true, breakLine: true } },
    { text: '카운슬 청사·최대 도서관이 있는 행정 중심', options: { bullet: true } } ],
    { x: 8.35, y: 2.6, w: 4.0, h: 3.6, fontSize: 15, color: C.background1, paraSpaceAfter: 10 });
  T(s, '※ 데이터 출처마다 기준일·방법이 달라 범위로 표기', { x: 0.6, y: 6.6, w: 7, h: 0.3, fontSize: 10, color: C.accent3, italic: true });

  // 4. Why strong
  pres.addSection({ title: '왜 뛰어난가' });
  s = pres.addSlide({ masterName: 'CONTENT', sectionTitle: '왜 뛰어난가' });
  s.addText('부동산 가치를 받치는 다섯 가지 기둥', { placeholder: 'title' });
  const pillars = [
    ['train', '교통', '고든역은 T1 노선 정차역이자 버스 환승 거점. 노스시드니·채스우드·CBD로 직결'],
    ['school', '학군', 'Ravenswood School for Girls(K-12)·Gordon East PS 소재. 인근 Knox·PLC Pymble 등 명문 사립 접근성'],
    ['tree', '자연환경', '국립공원 사이의 수목 캐노피. 카운슬의 수목 보호 정책으로 녹지 유지'],
    ['home', '주택 품질', '페더레이션·전간기 양식의 대형 필지 단독주택. 상당수가 문화유산 보존구역(HCA)'],
    ['shop', '생활 인프라', 'Gordon Centre 쇼핑센터·식당가, 카운슬 도서관·행정시설이 역 주변에 집중'] ];
  pillars.forEach((r, i) => { const y = 1.55 + i * 1.05; iconCircle(s, r[0], 0.7, y, 0.75);
    T(s, r[1], { x: 1.75, y, w: 2.0, h: 0.75, fontSize: 20, bold: true, color: C.text2, valign: 'middle' });
    T(s, r[2], { x: 3.8, y, w: 8.9, h: 0.75, fontSize: 16, valign: 'middle' }); });

  // 5. History timeline
  pres.addSection({ title: '부촌의 형성' });
  s = pres.addSlide({ masterName: 'CONTENT', sectionTitle: '부촌의 형성' });
  s.addText('철도가 만든 부촌: 1879년부터 오늘까지', { placeholder: 'title' });
  const tl = [['1879', 'Gordondale 영지 이름을 따 우체국을 "Gordon"으로 명명'],
              ['1890', '노스쇼어 철도와 고든역 개통 → 주거지 개발 본격화'],
              ['1906', 'Ku-ring-gai Shire 설립, 첫 회의를 고든에서 개최'],
              ['1927–28', '노선 전철화, 시(Municipality) 승격과 고든 청사 개관'],
              ['20세기', '대형 필지 정원주택의 전문직 주거지로 정착'],
              ['2024–25', '역세권 고밀화(TOD) 논쟁 → 카운슬 대안안 확정']];
  s.addShape(pres.shapes.LINE, { x: 0.9, y: 3.65, w: 11.5, h: 0, line: { color: C.accent6, width: 2 } });
  tl.forEach((r, i) => { const x = 0.6 + i * 2.05;
    s.addShape(pres.shapes.OVAL, { x: x + 0.15, y: 3.45, w: 0.4, h: 0.4, fill: { color: i === 5 ? C.accent2 : C.accent1 }, line: { color: C.background1, width: 2 } });
    T(s, r[0], { x, y: 2.6, w: 1.9, h: 0.6, fontSize: 22, bold: true, color: i === 5 ? C.accent2 : C.text2 });
    T(s, r[1], { x, y: 4.15, w: 1.9, h: 2.4, fontSize: 15 }); });

  // 6. Why wealthy mechanism
  s = pres.addSlide({ masterName: 'CONTENT', sectionTitle: '부촌의 형성' });
  s.addText('부촌이 된 메커니즘: 접근성 × 공급 제약', { placeholder: 'title' });
  const flow = [['철도 개통', '도심 출퇴근이 가능한 언덕 위 숲속 주거지로 전문직·상류층 유입'],
                ['대형 필지 개발', '넓은 정원과 고급 단독주택 위주로 분양, 고급 주거지 이미지 정착'],
                ['보존 정책', '카운슬이 문화유산 보존구역·수목 보호·저밀도 조닝을 오래 유지'],
                ['희소성 프리미엄', '수요는 꾸준한데 신규 공급은 묶임 → 단독주택 가격 상승 누적']];
  flow.forEach((r, i) => { const x = 0.6 + i * 3.15;
    card(s, x, 1.7, 2.75, 3.4, i === 3 ? C.text2 : C.background2);
    T(s, String(i + 1).padStart(2, '0'), { x: x + 0.3, y: 1.95, w: 1, h: 0.6, fontSize: 28, bold: true, color: C.accent2 });
    T(s, r[0], { x: x + 0.3, y: 2.65, w: 2.2, h: 0.5, fontSize: 19, bold: true, color: i === 3 ? C.background1 : C.text2 });
    T(s, r[1], { x: x + 0.3, y: 3.25, w: 2.2, h: 1.7, fontSize: 15, color: i === 3 ? C.background1 : C.text1 });
    if (i < 3) s.addShape(pres.shapes.CHEVRON, { x: x + 2.84, y: 3.25, w: 0.25, h: 0.3, fill: { color: C.accent3 }, line: { type: 'none' } });
  });
  T(s, '결과: 고든은 2026년에도 단독주택 중위가격이 유닛의 약 4배인 "땅값 중심" 시장', { x: 0.6, y: 5.5, w: 12.1, h: 0.6, fontSize: 20, bold: true, color: C.text2 });

  // 7. Council relationship
  pres.addSection({ title: 'Ku-ring-gai Council' });
  s = pres.addSlide({ masterName: 'CONTENT', sectionTitle: 'Ku-ring-gai Council' });
  s.addText('고든은 Ku-ring-gai Council의 행정 심장부', { placeholder: 'title' });
  const rel = [['landmark', '행정 중심', '카운슬 청사(1928년 원형 건물 개관)가 Pacific Hwy 고든에 위치. 카운슬 탄생(1906)부터 고든과 함께함'],
               ['book', '공공시설', '고든 도서관은 4개 분관 중 최대(1992년 건립). 2024년 주정부 보조금 $312,315로 개선 사업'],
               ['city', '시빅 허브 부지', '828 Pacific Hwy 카운슬 소유지를 미래 "Gordon Civic Hub"용으로 지정(문화·도서관·행정 통합 구상)'],
               ['shield', '보존형 도시계획', '문화유산 보존구역·수목 캐노피 정책으로 고든의 녹지·주택 성격을 지켜 가격을 뒷받침']];
  rel.forEach((r, i) => { const x = 0.6 + (i % 2) * 6.15, y = 1.6 + Math.floor(i / 2) * 2.55;
    card(s, x, y, 5.95, 2.35); iconCircle(s, r[0], x + 0.3, y + 0.3);
    T(s, r[1], { x: x + 1.25, y: y + 0.4, w: 4.4, h: 0.5, fontSize: 19, bold: true, color: C.text2 });
    T(s, r[2], { x: x + 1.25, y: y + 0.95, w: 4.4, h: 1.3, fontSize: 15 }); });

  // 8. Policy conflict timeline
  s = pres.addSlide({ masterName: 'CONTENT', sectionTitle: 'Ku-ring-gai Council' });
  s.addText('주정부 TOD vs 카운슬 대안안: 갈등에서 합의로', { placeholder: 'title' });
  const steps = [['2024', 'NSW TOD 정책', '역 반경 400m 내 6층(어포더블 포함 시 7~8층) 일률 허용. 보존구역도 포함'],
                 ['2024', '카운슬 법적 대응', '토지환경법원(LEC)에 TOD SEPP 소송 제기, 동시에 대안 시나리오 마련'],
                 ['2024.12', '5개 시나리오 전시', '고든 최대 45층안까지 주민 의견 수렴'],
                 ['2025.06', '대안안 채택', '고든 28층·린드필드 18층·킬라라/로즈빌 8층. 6/13 TOD SSD 경로 종료'],
                 ['2025.11', '주정부 최종 확정', '11/14 시행. 고든·린드필드 핵심부지 높이·용적률 상향 반영']];
  steps.forEach((r, i) => { const y = 1.5 + i * 1.08;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y, w: 1.5, h: 0.85, rectRadius: 0.1, fill: { color: i === 4 ? C.accent2 : C.accent1 }, line: { type: 'none' } });
    T(s, r[0], { x: 0.6, y, w: 1.5, h: 0.85, fontSize: 16, bold: true, color: C.background1, align: 'center', valign: 'middle' });
    T(s, r[1], { x: 2.35, y: y + 0.05, w: 2.9, h: 0.75, fontSize: 17, bold: true, color: C.text2, valign: 'middle' });
    T(s, r[2], { x: 5.3, y: y + 0.05, w: 7.4, h: 0.75, fontSize: 15, valign: 'middle' }); });

  // 9. Future plan numbers + chart
  pres.addSection({ title: '향후 개발' });
  s = pres.addSlide({ masterName: 'CONTENT', sectionTitle: '향후 개발' });
  s.addText('고든 배분 9,012세대: 4개 역세권 중 2위 규모', { placeholder: 'title' });
  s.addChart(pres.charts.BAR, [{ name: '배분 세대 수', labels: ['린드필드', '고든', '로즈빌', '킬라라'], values: [9419, 9012, 3353, 2778] }], {
    x: 0.6, y: 1.5, w: 6.6, h: 4.9, barDir: 'bar', chartColors: [HEX.accent1], barGapWidthPct: 60,
    showTitle: true, title: '카운슬 대안안 역세권별 세대 수(20년, 총 24,562)', titleFontSize: 14, titleColor: HEX.dk2, titleFontFace: '+mn-lt',
    showValue: true, dataLabelPosition: 'outEnd', dataLabelFormatCode: '#,##0', dataLabelFontSize: 14, dataLabelColor: HEX.dk1, dataLabelFontFace: '+mn-lt',
    catAxisLabelColor: HEX.dk1, catAxisLabelFontSize: 14, catAxisLabelFontFace: '+mn-lt', catAxisOrientation: 'maxMin', valAxisHidden: true,
    valGridLine: { style: 'none' }, catGridLine: { style: 'none' }, showLegend: false, valAxisMinVal: 0, valAxisMaxVal: 11500 });
  T(s, '출처: Ku-ring-gai Council, TOD Alternative Preferred Scenario FAQ(2025) — SJB 컨설턴트 추정치, 주정부(DPHI) 검증 전 수치', { x: 0.6, y: 6.45, w: 6.6, h: 0.45, fontSize: 10, color: C.accent3, italic: true });
  const facts = [['28층', '고든 핵심부지(Gordon Centre) 최고 높이, 4개 역세권 3~28층'],
                 ['800m', '역 반경 800m까지 개발, 보존구역(HCA) 80% 보호'],
                 ['24,728', '2025.11 주정부 확정안 전체 수용력(기존 TOD 21,430 대비 +약 3,300)'],
                 ['7,600', '2029년까지 Ku-ring-gai 주택 공급 목표(국가 주택 협약)']];
  facts.forEach((r, i) => { const y = 1.5 + i * 1.25;
    T(s, r[0], { x: 7.6, y, w: 1.7, h: 1.0, fontSize: 26, bold: true, color: C.accent2, valign: 'middle' });
    T(s, r[1], { x: 9.35, y, w: 3.4, h: 1.0, fontSize: 14, valign: 'middle' }); });

  // 10. Projects
  s = pres.addSlide({ masterName: 'CONTENT', sectionTitle: '향후 개발' });
  s.addText('고든에서 진행 중인 주요 개발 프로젝트', { placeholder: 'title' });
  const rows = [['프로젝트', '내용', '진행 단계'],
    ['810 Pacific Highway', '복합개발 + 어포더블 주택 (SSD)', '심사 중'],
    ['Uniting Gordon', '시니어 주거 타워, 약 $4.45억 규모', 'EIS 준비'],
    ['25–27 Rosedale Rd · 1–1A Edward St', '공동주택 + 인필 어포더블 (SSD)', '의견 회신'],
    ['Burgoyne St · Pearson Ave', '공동주택 2개 동 (SSD-82395459)', '의견 회신'],
    ['3–9 Park Avenue', '공동주택 약 9층 + 어포더블 (SSD-78775458)', '의견 회신'],
    ['828 Pacific Hwy (카운슬 부지)', 'Gordon Civic Hub 예정지', '장기 구상']];
  s.addTable(rows.map((r, ri) => r.map(c => ({ text: c, options: ri === 0 ? { bold: true, color: C.background1, fill: { color: C.accent1 } } : { color: C.text1, fill: { color: ri % 2 ? C.background1 : C.background2 } } }))),
    { x: 0.6, y: 1.5, w: 12.1, colW: [3.9, 5.6, 2.6], fontSize: 14, rowH: 0.62, border: { type: 'solid', pt: 0.5, color: HEX.accent6 }, valign: 'middle', margin: [0, 0.12, 0, 0.12] });
  T(s, '단계는 NSW Planning Portal 기준이며 변동 가능 · 2025.6.13 이전 접수된 SSD 신청이 다수(TOD 경로 종료 전 몰림)', { x: 0.6, y: 6.1, w: 12.1, h: 0.5, fontSize: 11, color: C.accent3, italic: true });

  // 11. Implications
  pres.addSection({ title: '시사점' });
  s = pres.addSlide({ masterName: 'CONTENT', sectionTitle: '시사점' });
  s.addText('투자 관점 시사점: 기회와 리스크', { placeholder: 'title' });
  const cols = [['up', '기회', C.accent1, ['역세권 고밀도 조닝으로 부지 통합(amalgamation) 수요 → 대상 필지 토지가치 상승 여지', '신축 아파트 공급으로 다운사이저·젊은 가구 유입, 상권 활성화', '시빅 허브·도서관 등 공공시설 투자로 생활 편의 개선']],
                ['warn', '리스크', C.accent4, ['유닛 공급 급증 시 아파트 가격·임대료 상승 제한 가능', '최근 12개월 단독주택 가격 보합(+0.1%) — 단기 모멘텀 약함', '교통·학교·주차 등 인프라 부담과 동네 성격 변화 논쟁 지속']]];
  cols.forEach((c, i) => { const x = 0.6 + i * 6.15; card(s, x, 1.55, 5.95, 4.6);
    iconCircle(s, c[0], x + 0.3, 1.85, 0.7, c[2]);
    T(s, c[1], { x: x + 1.2, y: 1.95, w: 3, h: 0.5, fontSize: 22, bold: true, color: c[2] });
    T(s, c[3].map((t, j) => ({ text: t, options: { bullet: true, breakLine: j < 2 } })), { x: x + 0.35, y: 2.85, w: 5.3, h: 3.1, fontSize: 17, paraSpaceAfter: 14 }); });

  // 12. Closing
  s = pres.addSlide({ masterName: 'DARK', sectionTitle: '시사점' });
  s.addText('결론: 저밀도 부촌에서 역세권 중심지로', { placeholder: 'title' });
  s.addText('철도·학군·숲·보존 정책이 만든 희소가치는 여전하며, 2025년 확정된 카운슬 대안안으로 역 주변은 최고 28층의 고밀 중심지로 재편될 예정. 단독주택(특히 보존구역)은 희소성 유지, 역세권 필지는 개발 프리미엄, 아파트는 공급 증가 영향을 주시할 필요.', { placeholder: 'body', fontSize: 18 });

  // 13. Sources
  s = pres.addSlide({ masterName: 'CONTENT', sectionTitle: '시사점' });
  s.addText('출처', { placeholder: 'title' });
  const src = ['Microburbs, OnTheHouse, Aus Property Insights — Gordon NSW 2072 시세 (2026)',
    'Wikipedia — Gordon, NSW / Gordon railway station / Ku-ring-gai Council',
    'Ku-ring-gai Council — Council history; TOD Preferred Alternative Scenario FAQ; 도서관 보조금 뉴스(2024)',
    'Ku-ring-gai Council — Community Facilities Strategy (2018), Gordon Civic Hub 관련 보고서',
    'NSW Government 보도자료 — "3,300 extra homes unlocked in Ku-ring-gai"',
    'The Urban Developer, Inside Local Government, City Hub, STEP, FOKE — TOD 대안안 보도',
    'NSW Planning Portal — 고든 SSD 프로젝트(810 Pacific Hwy, Uniting Gordon, Rosedale Rd, Burgoyne St, Park Ave)',
    'Property Council of Australia — 카운슬 대안안 의견서 (2025.5)'];
  T(s, src.map((t, j) => ({ text: t, options: { bullet: true, breakLine: j < src.length - 1 } })), { x: 0.6, y: 1.5, w: 12.1, h: 4.8, fontSize: 14, paraSpaceAfter: 8 });
  T(s, '본 자료는 공개 정보 기반 리서치이며 투자 자문이 아닙니다. 가격·계획은 수시로 변동되니 최신 자료를 확인하세요.', { x: 0.6, y: 6.35, w: 12.1, h: 0.4, fontSize: 11, italic: true, color: C.accent3 });

  await pres.writeFile({ fileName: 'Gordon_부동산_분석.pptx' });
  await applyTheme('Gordon_부동산_분석.pptx', THEME);
  require('child_process').execFileSync('python3', [__dirname + '/fix_fonts.py', 'Gordon_부동산_분석.pptx']);
})();
