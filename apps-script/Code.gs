// ============================================
// 해커톤 작품 갤러리 - 데이터 API (Code.gs)
// GitHub Pages의 정적 갤러리가 이 웹앱을 호출해
// 시트 데이터를 JSON으로 받아갑니다.
// 시트 주소 가운데 긴 문자열이 시트 ID입니다.
// 예) docs.google.com/spreadsheets/d/이_부분이_ID/edit
// ============================================
const SHEET_ID = '여기에_시트_ID_붙여넣기';

// 회차 탭으로 취급하지 않는 특수 탭 이름들
const RESERVED_TABS = ['소개', '설명'];

// 정적 페이지(GitHub Pages)가 부르는 주소입니다.
// ?season=회차탭이름 (생략 시 최신 회차)
// &callback=함수이름  (JSONP - 도메인이 달라도 동작하게 해줍니다)
function doGet(e) {
  const p = (e && e.parameter) || {};
  const data = getGallery(p.season || '');
  const json = JSON.stringify(data);
  const cb = p.callback;
  if (cb && /^[A-Za-z_$][\w$]*$/.test(cb)) {
    return ContentService.createTextOutput(cb + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

// 회차 목록 = 시트 탭 이름들 ('소개', '설명' 탭 제외)
function getSeasons() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  return ss.getSheets()
    .map(s => s.getName())
    .filter(name => RESERVED_TABS.indexOf(name) < 0);
}

// 해커톤 설명 캐러셀 슬라이드 ('설명' 탭: A열=제목, B열=내용, 한 줄이 한 장)
function getSlides(ss) {
  const sh = ss.getSheetByName('설명');
  if (!sh) return [];
  const rows = sh.getDataRange().getValues();
  const slides = [];
  for (let i = 1; i < rows.length; i++) {   // 1행은 헤더
    const title = String(rows[i][0] || '').trim();
    const body  = String(rows[i][1] || '').trim();
    if (!title && !body) continue;
    slides.push({ title, body });
  }
  return slides;
}

// 갤러리 데이터: 소개 문구 + 설명 슬라이드 + 선택 회차의 작품 목록
// 작품 목록은 "줄 순서 = 모둠 자리"입니다. i번째 줄이 i번 자리 카드가 되고,
// 모둠 이름 칸을 바꾸면 갤러리에도 그대로 반영됩니다.
function getGallery(season) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const seasons = getSeasons();
  const current = (season && seasons.indexOf(season) >= 0)
    ? season
    : seasons[seasons.length - 1];  // 기본 = 마지막 탭(최신 회차)

  // 소개 탭: A1=소개 문구, A2=대회 이름(예: 1st SCH Hackathon), A3=날짜(예: 2026.10.24 토요일)
  let intro = '', eventTitle = '', eventDate = '';
  const introSheet = ss.getSheetByName('소개');
  if (introSheet) {
    intro = String(introSheet.getRange('A1').getValue() || '');
    eventTitle = String(introSheet.getRange('A2').getValue() || '');
    eventDate = String(introSheet.getRange('A3').getValue() || '');
  }

  const slides = getSlides(ss);

  const works = [];
  const sheet = ss.getSheetByName(current);
  if (sheet) {
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {   // 1행은 헤더
      const team  = String(rows[i][0] || '').trim();
      const title = String(rows[i][1] || '').trim();
      const desc  = String(rows[i][2] || '').trim();
      const link  = String(rows[i][3] || '').trim();
      const thumb = String(rows[i][4] || '').trim();
      if (!team && !title && !link) continue;
      works.push({ slot: works.length + 1, team, title, desc, link, thumb });
    }
  }
  return { intro, eventTitle, eventDate, slides, seasons, current, works };
}
