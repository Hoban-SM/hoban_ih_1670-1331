/*******************************************************
 * 인천검단 AB13 호반써밋Ⅲ 부부공동명의(분양권 전매) 예약 시스템
 * Google Apps Script 백엔드
 *
 * [배포 방법]
 * 1) 새 Google 스프레드시트를 만들고 아래 SPREADSHEET_ID 에 ID 입력
 *    (시트/탭은 자동 생성되므로 처음엔 빈 시트여도 됨)
 * 2) 확장프로그램 > Apps Script 에서 이 파일 내용을 Code.gs 에 붙여넣기
 * 3) 파일 상단 CONFIG 값들을 실제 값으로 수정
 * 4) 배포 > 웹 앱으로 배포
 *      - 실행 계정: 나(소유자)
 *      - 액세스 권한: 전체 공개(모든 사용자, Google 계정 없어도 됨)
 * 5) 배포 후 생성되는 웹앱 URL을 index.html / admin.html 의 API_URL 에 입력
 * 6) 코드 수정 시에는 "새 배포"가 아니라 기존 배포에서
 *    "배포 관리 > 편집(연필) > 버전:새 버전 > 배포" 로 갱신해야
 *    URL이 바뀌지 않음 (URL 바뀌면 GitHub Pages 쪽도 다시 수정해야 함)
 *******************************************************/

const CONFIG = {
  SPREADSHEET_ID: 'YOUR_SPREADSHEET_ID_HERE',   // 스프레드시트 ID
  DRIVE_FOLDER_ID: 'YOUR_DRIVE_FOLDER_ID_HERE', // 첨부파일(가족관계증명서) 저장 폴더 ID
  NOTIFY_EMAIL: '16701331@hobanenc.co.kr',      // 접수 시 메일 발송 대상
  ADMIN_PASSWORD: 'CHANGE_ME',                  // 관리자 페이지 비밀번호
  SLOT_SHEET: '슬롯',
  RESERVATION_SHEET: '예약',
  SITE_NAME: '인천검단 AB13블록 호반써밋Ⅲ',
  TIMEZONE: 'Asia/Seoul'
};

const SLOT_HEADERS = ['ID', '날짜', '요일', '시작시간', '종료시간', '정원', '예약자수', '상태'];
const RESV_HEADERS = [
  '예약ID', '접수일시', '예약일', '예약시간',
  '동', '호',
  '양도인_성명', '양도인_생년월일', '양도인_연락처', '양도인_주소',
  '양수인_성명', '양수인_생년월일', '양수인_연락처', '양수인_주소',
  '첨부파일명', '첨부파일링크',
  '개인정보동의_양도인', '개인정보동의_양수인', '상태'
];

/* ===================== 진입점 ===================== */

function doGet(e) {
  try {
    const action = e.parameter.action || 'getSlots';
    let result;
    if (action === 'getSlots') {
      result = { ok: true, data: getOpenSlots() };
    } else {
      result = { ok: false, error: 'Unknown GET action' };
    }
    return jsonOut_(result);
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    let result;

    switch (action) {
      case 'reserve':
        result = handleReserve_(body);
        break;
      case 'checkStatus':
        result = handleCheckStatus_(body);
        break;
      case 'adminLogin':
        result = { ok: checkPassword_(body.password) };
        break;
      case 'addSlots':
        checkAdmin_(body.password);
        result = { ok: true, data: handleAddSlots_(body) };
        break;
      case 'deleteSlot':
        checkAdmin_(body.password);
        result = { ok: true, data: handleDeleteSlot_(body.id) };
        break;
      case 'deleteSlotsByDate':
        checkAdmin_(body.password);
        result = { ok: true, data: handleDeleteSlotsByDate_(body.date) };
        break;
      case 'getAllSlots':
        checkAdmin_(body.password);
        result = { ok: true, data: getAllSlots_() };
        break;
      case 'getReservations':
        checkAdmin_(body.password);
        result = { ok: true, data: getAllReservations_() };
        break;
      case 'updateStatus':
        checkAdmin_(body.password);
        result = { ok: true, data: handleUpdateStatus_(body.reservationId, body.status) };
        break;
      default:
        result = { ok: false, error: 'Unknown POST action' };
    }
    return jsonOut_(result);
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ===================== 인증 ===================== */

function checkPassword_(pw) {
  return pw === CONFIG.ADMIN_PASSWORD;
}

function checkAdmin_(pw) {
  if (!checkPassword_(pw)) throw new Error('관리자 인증 실패');
}

/* ===================== 시트 유틸 ===================== */

function getSS_() {
  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
}

function getSheet_(name, headers) {
  const ss = getSS_();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.setFrozenRows(1);
  }
  return sh;
}

function sheetToObjects_(sh) {
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  const rows = values.slice(1);
  return rows
    .map((row, idx) => {
      const obj = {};
      headers.forEach((h, i) => (obj[h] = row[i]));
      obj._row = idx + 2; // 실제 시트 행 번호
      return obj;
    })
    .filter(o => o['ID'] !== '' && o['ID'] !== undefined || o['예약ID'] !== undefined);
}

/* ===================== 슬롯(예약 가능 일시) ===================== */

function getOpenSlots() {
  const sh = getSheet_(CONFIG.SLOT_SHEET, SLOT_HEADERS);
  const all = sheetToObjects_(sh);
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  return all
    .filter(s => {
      if (s['상태'] === '마감') return false;
      const d = new Date(s['날짜']);
      if (isNaN(d)) return false;
      d.setHours(0, 0, 0, 0);
      if (d < today) return false;
      const cap = Number(s['정원']) || 0;
      const used = Number(s['예약자수']) || 0;
      return used < cap;
    })
    .map(s => ({
      id: s['ID'],
      date: formatDate_(s['날짜']),
      weekday: s['요일'],
      start: formatTime_(s['시작시간']),
      end: formatTime_(s['종료시간']),
    }))
    .sort((a, b) => (a.date + a.start).localeCompare(b.date + b.start));
}

function getAllSlots_() {
  const sh = getSheet_(CONFIG.SLOT_SHEET, SLOT_HEADERS);
  return sheetToObjects_(sh).map(s => ({
    id: s['ID'],
    date: formatDate_(s['날짜']),
    weekday: s['요일'],
    start: formatTime_(s['시작시간']),
    end: formatTime_(s['종료시간']),
    row: s['_row']
  }));
}

const WEEKDAY_KR = ['일', '월', '화', '수', '목', '금', '토'];

function handleAddSlots_(body) {
  // body: { dateFrom, dateTo, startTime, endTime, lunchStart, lunchEnd,
  //         capacity, onlyWeekday(optional: '화' 등), excludeDates(optional: "2026-08-11,2026-09-08"), password }
  const sh = getSheet_(CONFIG.SLOT_SHEET, SLOT_HEADERS);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const startDate = new Date(body.dateFrom);
    const endDate = new Date(body.dateTo);
    const capacity = Number(body.capacity) || 1;
    const created = [];

    const excludeSet = new Set(
      String(body.excludeDates || '')
        .split(',')
        .map(s => s.trim())
        .filter(Boolean)
    );

    // 구글시트가 "10:00" 같은 문자열을 시간 형식으로 자동 인식해 저장하는 것을 막기 위해
    // 시작시간(D열)/종료시간(E열)을 미리 텍스트 서식으로 고정
    sh.getRange(2, 4, 5000, 2).setNumberFormat('@');

    for (let d = new Date(startDate); d <= endDate; d.setDate(d.getDate() + 1)) {
      const wd = WEEKDAY_KR[d.getDay()];
      if (body.onlyWeekday && body.onlyWeekday !== wd) continue;

      const dateStr = formatDate_(new Date(d));
      if (excludeSet.has(dateStr)) continue; // 공휴일/중도금 약정일 등 제외 날짜는 건너뜀

      const slots = buildTimeSlots_(body.startTime, body.endTime, body.lunchStart, body.lunchEnd);

      slots.forEach(t => {
        const id = Utilities.getUuid();
        sh.appendRow([id, dateStr, wd, t.start, t.end, capacity, 0, '사용']);
        created.push({ id, date: dateStr, start: t.start, end: t.end });
      });
    }
    return created;
  } finally {
    lock.releaseLock();
  }
}

// 30분 단위 시간대 생성, 점심시간(lunchStart~lunchEnd) 제외
function buildTimeSlots_(startTime, endTime, lunchStart, lunchEnd) {
  const toMin = t => {
    const [h, m] = t.split(':').map(Number);
    return h * 60 + m;
  };
  const toStr = m => {
    const h = Math.floor(m / 60);
    const mm = m % 60;
    return (h < 10 ? '0' : '') + h + ':' + (mm < 10 ? '0' : '') + mm;
  };

  const start = toMin(startTime);
  const end = toMin(endTime);
  const lStart = lunchStart ? toMin(lunchStart) : null;
  const lEnd = lunchEnd ? toMin(lunchEnd) : null;

  const slots = [];
  for (let t = start; t + 30 <= end; t += 30) {
    if (lStart !== null && lEnd !== null && t >= lStart && t < lEnd) continue;
    slots.push({ start: toStr(t), end: toStr(t + 30) });
  }
  return slots;
}

function handleDeleteSlot_(id) {
  const sh = getSheet_(CONFIG.SLOT_SHEET, SLOT_HEADERS);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sh.deleteRow(i + 1);
      return { deleted: id };
    }
  }
  return { deleted: null };
}

// 특정 날짜의 슬롯을 한 번에 삭제 (예: 공휴일로 뒤늦게 확인된 날짜, 중도금 약정일 포함 월 등)
// 이미 예약자가 있는 슬롯(예약자수 > 0)은 안전을 위해 건너뛰고 개수만 알려줌
function handleDeleteSlotsByDate_(dateStr) {
  const sh = getSheet_(CONFIG.SLOT_SHEET, SLOT_HEADERS);
  const data = sh.getDataRange().getValues();
  let deleted = 0;
  let skipped = 0;
  // 뒤에서부터 삭제해야 행 번호가 안 꼬임
  for (let i = data.length - 1; i >= 1; i--) {
    const rowDate = formatDate_(data[i][1]);
    if (rowDate !== dateStr) continue;
    const reserved = Number(data[i][6]) || 0;
    if (reserved > 0) {
      skipped++;
      continue;
    }
    sh.deleteRow(i + 1);
    deleted++;
  }
  return { date: dateStr, deleted, skipped };
}

function formatDate_(d) {
  if (typeof d === 'string') return d;
  const date = new Date(d);
  const y = date.getFullYear();
  const m = ('0' + (date.getMonth() + 1)).slice(-2);
  const day = ('0' + date.getDate()).slice(-2);
  return `${y}-${m}-${day}`;
}

// 시작시간/종료시간 셀 값을 항상 'HH:mm' 문자열로 반환.
// 구글시트가 "10:00" 같은 문자열을 시간 형식으로 자동 인식해 Date로 저장하는 경우가 있어
// (그 경우 getValues()가 1899-12-30 기준 Date 객체를 반환) 이를 안전하게 문자열로 변환해줌.
function formatTime_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, CONFIG.TIMEZONE, 'HH:mm');
  }
  return String(v);
}

/* ===================== 예약 처리 ===================== */

function handleReserve_(body) {
  // body: { slotId, dong, ho,
  //   transferor:{name,birth,phone,addr}, transferee:{name,birth,phone,addr},
  //   agreeTransferor, agreeTransferee,
  //   fileName, mimeType, fileData(base64) }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  let slotInfo;
  try {
    slotInfo = reserveSlotCapacity_(body.slotId);
  } finally {
    lock.releaseLock();
  }
  if (!slotInfo) {
    return { ok: false, error: '선택하신 시간대는 마감되었습니다. 다시 선택해주세요.' };
  }

  // 첨부파일 Drive 저장
  let fileUrl = '';
  let blob = null;
  if (body.fileData && body.fileName) {
    blob = Utilities.newBlob(
      Utilities.base64Decode(body.fileData),
      body.mimeType || 'application/octet-stream',
      body.fileName
    );
    const folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
    const file = folder.createFile(blob);
    fileUrl = file.getUrl();
  }

  const resvId = Utilities.getUuid();
  const now = new Date();
  const sh = getSheet_(CONFIG.RESERVATION_SHEET, RESV_HEADERS);

  const t = body.transferor || {};
  const e = body.transferee || {};

  sh.appendRow([
    resvId,
    now,
    slotInfo.date,
    slotInfo.start + '~' + slotInfo.end,
    body.dong || '',
    body.ho || '',
    t.name || '', t.birth || '', t.phone || '', t.addr || '',
    e.name || '', e.birth || '', e.phone || '', e.addr || '',
    body.fileName || '', fileUrl,
    body.agreeTransferor ? '동의' : '미동의',
    body.agreeTransferee ? '동의' : '미동의',
    '접수완료'
  ]);

  sendNotifyMail_(resvId, slotInfo, body, blob);

  return { ok: true, data: { reservationId: resvId, date: slotInfo.date, time: slotInfo.start + '~' + slotInfo.end } };
}

// 슬롯의 예약자수를 원자적으로 +1 (정원 초과시 null 반환)
function reserveSlotCapacity_(slotId) {
  const sh = getSheet_(CONFIG.SLOT_SHEET, SLOT_HEADERS);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === slotId) {
      const capacity = Number(data[i][5]) || 0;
      const reserved = Number(data[i][6]) || 0;
      if (reserved >= capacity || data[i][7] === '마감') return null;
      sh.getRange(i + 1, 7).setValue(reserved + 1); // '예약자수' 열
      return {
        date: formatDate_(data[i][1]),
        start: formatTime_(data[i][3]),
        end: formatTime_(data[i][4])
      };
    }
  }
  return null;
}

function sendNotifyMail_(resvId, slotInfo, body, blob) {
  const t = body.transferor || {};
  const e = body.transferee || {};
  const subject = `[분양권 전매신청 접수] ${CONFIG.SITE_NAME} ${body.dong || ''}동 ${body.ho || ''}호 - ${slotInfo.date} ${slotInfo.start}`;

  const htmlBody = `
    <h3>분양권 전매신청서 접수</h3>
    <p><b>현장:</b> ${CONFIG.SITE_NAME}</p>
    <p><b>예약ID:</b> ${resvId}</p>
    <p><b>희망 방문일시:</b> ${slotInfo.date} ${slotInfo.start} ~ ${slotInfo.end}</p>
    <p><b>동/호:</b> ${body.dong || ''}동 ${body.ho || ''}호</p>
    <hr>
    <p><b>[양도인(현 계약자)]</b><br>
    성명: ${t.name || ''}<br>
    생년월일: ${t.birth || ''}<br>
    연락처: ${t.phone || ''}<br>
    주소: ${t.addr || ''}</p>
    <p><b>[양수예정인(배우자)]</b><br>
    성명: ${e.name || ''}<br>
    생년월일: ${e.birth || ''}<br>
    연락처: ${e.phone || ''}<br>
    주소: ${e.addr || ''}</p>
    <hr>
    <p><b>개인정보 수집/이용 동의</b> - 양도인: ${body.agreeTransferor ? '동의' : '미동의'} / 양수인: ${body.agreeTransferee ? '동의' : '미동의'}</p>
    <p>첨부파일(가족관계증명서)이 함께 첨부되었습니다.</p>
  `;

  const options = { htmlBody };
  if (blob) options.attachments = [blob];

  MailApp.sendEmail({
    to: CONFIG.NOTIFY_EMAIL,
    subject,
    htmlBody,
    attachments: blob ? [blob] : []
  });
}

function getAllReservations_() {
  const sh = getSheet_(CONFIG.RESERVATION_SHEET, RESV_HEADERS);
  return sheetToObjects_(sh);
}

function handleUpdateStatus_(reservationId, status) {
  const sh = getSheet_(CONFIG.RESERVATION_SHEET, RESV_HEADERS);
  const data = sh.getDataRange().getValues();
  const statusCol = RESV_HEADERS.indexOf('상태') + 1; // 1-indexed
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === reservationId) {
      sh.getRange(i + 1, statusCol).setValue(status);
      return { reservationId, status };
    }
  }
  return { reservationId: null };
}

// 신청자 본인 접수확인: 동/호/계약자(양도인) 성명/휴대폰 뒷 4자리로 조회
// - 개인정보 최소화를 위해 4가지 값이 모두 정확히 일치하는 건만 반환
// - 응답에는 상세 주소 등 민감정보는 포함하지 않음
function handleCheckStatus_(body) {
  const dong = String(body.dong || '').trim();
  const ho = String(body.ho || '').trim();
  const name = String(body.name || '').trim();
  const phoneLast4 = String(body.phoneLast4 || '').trim();

  if (!dong || !ho || !name || phoneLast4.length !== 4) {
    return { ok: false, error: '동/호, 계약자 성명, 휴대폰 뒷 4자리를 모두 정확히 입력해주세요.' };
  }

  const sh = getSheet_(CONFIG.RESERVATION_SHEET, RESV_HEADERS);
  const all = sheetToObjects_(sh);

  const matches = all.filter(r => {
    const rDong = String(r['동'] || '').trim();
    const rHo = String(r['호'] || '').trim();
    const rName = String(r['양도인_성명'] || '').trim();
    const rPhone = String(r['양도인_연락처'] || '').replace(/[^0-9]/g, '');
    return rDong === dong && rHo === ho && rName === name && rPhone.slice(-4) === phoneLast4;
  });

  if (matches.length === 0) {
    return { ok: true, data: [] };
  }

  const data = matches
    .sort((a, b) => new Date(b['접수일시']) - new Date(a['접수일시']))
    .map(r => ({
      reservationId: r['예약ID'],
      submittedAt: r['접수일시'] instanceof Date ? r['접수일시'].toISOString() : r['접수일시'],
      visitDate: r['예약일'],
      visitTime: r['예약시간'],
      dong: r['동'],
      ho: r['호'],
      status: r['상태']
    }));

  return { ok: true, data };
}
