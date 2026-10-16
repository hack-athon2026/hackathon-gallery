// ============================================
// 해커톤 작품 갤러리 - 데이터 API (Code.gs)
// GitHub Pages의 정적 갤러리가 이 웹앱을 호출해
// 시트 데이터를 JSON으로 받아갑니다.
// 시트 주소 가운데 긴 문자열이 시트 ID입니다.
// 예) docs.google.com/spreadsheets/d/이_부분이_ID/edit
// ============================================
const SHEET_ID = '여기에_시트_ID_붙여넣기';

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

// 회차 목록 = 시트 탭 이름들 ('소개' 탭 제외)
function getSeasons() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  return ss.getSheets()
    .map(s => s.getName())
    .filter(name => name !== '소개');
}

// 갤러리 데이터: 소개 문구 + 선택 회차의 작품 목록
function getGallery(season) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const seasons = getSeasons();
  const current = (season && seasons.indexOf(season) >= 0)
    ? season
    : seasons[seasons.length - 1];  // 기본 = 마지막 탭(최신 회차)

  let intro = '';
  const introSheet = ss.getSheetByName('소개');
  if (introSheet) intro = String(introSheet.getRange('A1').getValue() || '');

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
      if (!team && !title) continue;
      works.push({ team, title, desc, link, thumb });
    }
  }
  return { intro, seasons, current, works };
}
