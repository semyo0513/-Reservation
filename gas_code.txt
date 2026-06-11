/**
 * 체험학습 프로그램 예약 시스템 - Google Apps Script 백엔드
 * 
 * [사용 방법]
 * 1. Google Apps Script (script.google.com)에서 새 프로젝트 생성
 * 2. 이 코드 전체를 붙여넣기
 * 3. SPREADSHEET_ID를 실제 Google Sheets ID로 변경
 * 4. 최초 1회 initializeAll() 함수 실행 (시트 구조 + 초기 데이터 생성)
 * 5. 배포 > 웹 앱으로 배포 (액세스: 모든 사용자)
 * 6. 배포된 URL을 index.html의 GAS_URL 상수에 입력
 */

// =====================================================================
// 설정값
// =====================================================================
const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE'; // Google Sheets ID로 교체
const ADMIN_DEFAULT_PASSWORD = 'admin1234';         // 초기 관리자 비밀번호
const TOKEN_EXPIRE_HOURS = 4;                       // 세션 토큰 만료 시간

// =====================================================================
// 진입점
// =====================================================================

function doGet(e) {
  const params = e.parameter;
  const action = params.action || '';
  try {
    let result;
    switch (action) {
      case 'getPrograms':    result = getPrograms(params); break;
      case 'checkOpenTime':  result = checkOpenTime(params); break;
      case 'getBookings':    result = getBookings(params); break;
      case 'getSettings':    result = getSettings(params); break;
      case 'checkBooking':   result = checkBooking(params); break;
      default: result = { success: false, message: '알 수 없는 액션입니다.' };
    }
    return jsonResponse(result);
  } catch (err) {
    return jsonResponse({ success: false, message: err.message });
  }
}

function doPost(e) {
  let params;
  try {
    params = JSON.parse(e.postData.contents);
  } catch (_) {
    params = e.parameter;
  }
  const action = params.action || '';
  try {
    let result;
    switch (action) {
      case 'submitBooking':   result = submitBooking(params); break;
      case 'adminLogin':      result = adminLogin(params); break;
      case 'saveProgram':     result = saveProgram(params); break;
      case 'deleteProgram':   result = deleteProgram(params); break;
      case 'saveSettings':    result = saveSettings(params); break;
      case 'cancelBooking':   result = cancelBooking(params); break;
      case 'manualBooking':   result = manualBooking(params); break;
      case 'changePassword':  result = changePassword(params); break;
      default: result = { success: false, message: '알 수 없는 액션입니다.' };
    }
    return jsonResponse(result);
  } catch (err) {
    return jsonResponse({ success: false, message: err.message });
  }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// =====================================================================
// 스프레드시트 헬퍼
// =====================================================================

function getSheet(name) {
  if (!SPREADSHEET_ID || SPREADSHEET_ID === 'YOUR_SPREADSHEET_ID_HERE') {
    throw new Error('SPREADSHEET_ID가 설정되지 않았습니다. gas_code.gs 상단의 SPREADSHEET_ID를 실제 Google Sheets ID로 교체하고, initializeAll()을 실행해 주세요.');
  }
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(name);
  if (!sheet) {
    throw new Error('"' + name + '" 시트를 찾을 수 없습니다. GAS 편집기에서 initializeAll() 함수를 한 번 실행해 시트를 생성해 주세요.');
  }
  return sheet;
}

function getSheetSafe(name) {
  // 시트가 없어도 null을 반환 (크래시 없음)
  try {
    if (!SPREADSHEET_ID || SPREADSHEET_ID === 'YOUR_SPREADSHEET_ID_HERE') return null;
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    return ss.getSheetByName(name);
  } catch(e) {
    return null;
  }
}

function getAllData(sheetName) {
  const sheet = getSheet(sheetName);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return [];
  const headers = rows[0];
  return rows.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}

function appendRow(sheetName, rowObj) {
  const sheet = getSheet(sheetName);
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) throw new Error('"' + sheetName + '" 시트에 헤더가 없습니다. initializeAll()을 다시 실행해 주세요.');
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const row = headers.map(h => rowObj[h] !== undefined ? rowObj[h] : '');
  sheet.appendRow(row);
}

function getSetting(key) {
  const data = getAllData('settings');
  const item = data.find(r => r.key === key);
  return item ? item.value : null;
}

function setSetting(key, value) {
  const sheet = getSheet('settings');
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow >= 2 && lastCol >= 1) {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === key) {
        sheet.getRange(i + 1, 2).setValue(value);
        return;
      }
    }
  }
  sheet.appendRow([key, value]);
}

// =====================================================================
// SHA-256 해시 (GAS 내장)
// =====================================================================

function sha256(text) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    text,
    Utilities.Charset.UTF_8
  );
  return bytes.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

// =====================================================================
// UUID 생성
// =====================================================================

function generateUUID() {
  return Utilities.getUuid();
}

// =====================================================================
// 관리자 인증
// =====================================================================

function adminLogin(params) {
  const { password } = params;
  if (!password) return { success: false, message: '비밀번호를 입력하세요.' };

  const storedHash = getSetting('adminPasswordHash');
  const inputHash = sha256(password);

  if (storedHash !== inputHash) {
    return { success: false, message: '비밀번호가 올바르지 않습니다.' };
  }

  const token = generateUUID();
  const expire = new Date(Date.now() + TOKEN_EXPIRE_HOURS * 60 * 60 * 1000).toISOString();
  setSetting('sessionToken', token);
  setSetting('sessionExpire', expire);

  return { success: true, token };
}

function validateToken(token) {
  if (!token) return false;
  const storedToken = getSetting('sessionToken');
  const expireStr = getSetting('sessionExpire');
  if (!storedToken || !expireStr) return false;
  if (storedToken !== token) return false;
  if (new Date() > new Date(expireStr)) return false;
  return true;
}

function changePassword(params) {
  const { token, newPassword } = params;
  if (!validateToken(token)) return { success: false, message: '인증이 필요합니다.' };
  if (!newPassword || newPassword.length < 4) return { success: false, message: '비밀번호는 4자 이상이어야 합니다.' };
  setSetting('adminPasswordHash', sha256(newPassword));
  return { success: true, message: '비밀번호가 변경되었습니다.' };
}

// =====================================================================
// 프로그램 관련
// =====================================================================

function getPrograms(params) {
  const { classNum } = params;
  const programs = getAllData('programs');
  const bookings = getAllData('bookings');

  // 신청 수 집계
  const countMap = {};
  bookings.forEach(b => {
    if (b.programId) {
      countMap[b.programId] = (countMap[b.programId] || 0) + 1;
    }
  });

  let filtered = programs;
  if (classNum) {
    const cn = String(classNum);
    filtered = programs.filter(p => {
      const targets = String(p.targetClasses).split(',').map(s => s.trim());
      return targets.includes(cn);
    });
  }

  const result = filtered.map(p => ({
    id: p.id,
    name: p.name,
    date: p.date,
    targetClasses: p.targetClasses,
    capacity: Number(p.capacity),
    bookedCount: countMap[p.id] || 0,
    isFull: (countMap[p.id] || 0) >= Number(p.capacity)
  }));

  return { success: true, programs: result };
}

function saveProgram(params) {
  const { token, program } = params;
  if (!validateToken(token)) return { success: false, message: '인증이 필요합니다.' };

  const sheet = getSheet('programs');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  if (program.id) {
    // 수정
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === program.id) {
        const row = headers.map(h => program[h] !== undefined ? program[h] : data[i][headers.indexOf(h)]);
        sheet.getRange(i + 1, 1, 1, headers.length).setValues([row]);
        return { success: true, message: '프로그램이 수정되었습니다.' };
      }
    }
  }

  // 추가
  const newId = 'p' + String(Date.now());
  const newProgram = {
    id: newId,
    name: program.name,
    date: program.date,
    targetClasses: program.targetClasses,
    capacity: program.capacity,
    createdAt: new Date().toISOString()
  };
  appendRow('programs', newProgram);
  return { success: true, message: '프로그램이 추가되었습니다.', id: newId };
}

function deleteProgram(params) {
  const { token, programId } = params;
  if (!validateToken(token)) return { success: false, message: '인증이 필요합니다.' };

  const bookings = getAllData('bookings');
  const hasBookings = bookings.some(b => b.programId === programId);
  if (hasBookings) {
    // 경고만 하고 force 플래그 없으면 차단
    if (!params.force) {
      return { success: false, message: '이 프로그램에 신청자가 있습니다. 강제 삭제하시겠습니까?', needConfirm: true };
    }
  }

  const sheet = getSheet('programs');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === programId) {
      sheet.deleteRow(i + 1);
      return { success: true, message: '프로그램이 삭제되었습니다.' };
    }
  }
  return { success: false, message: '프로그램을 찾을 수 없습니다.' };
}

// =====================================================================
// 오픈 시간 관련
// =====================================================================

function checkOpenTime(params) {
  const { date } = params;
  if (!date) return { success: false, message: 'date 파라미터가 필요합니다.' };

  const openTimeStr = getSetting('openTime_' + date);
  const forceOpen = getSetting('forceOpen_' + date);
  const forceClose = getSetting('forceClose_' + date);

  if (forceClose === 'true') {
    return { success: true, isOpen: false, isClosed: true, message: '예매가 마감되었습니다.' };
  }
  if (forceOpen === 'true') {
    return { success: true, isOpen: true };
  }
  if (!openTimeStr) {
    return { success: true, isOpen: false, openTime: null, message: '오픈 시간이 설정되지 않았습니다.' };
  }

  const openTime = new Date(openTimeStr);
  const now = new Date();
  const isOpen = now >= openTime;

  return {
    success: true,
    isOpen,
    openTime: openTimeStr,
    serverTime: now.toISOString()
  };
}

// =====================================================================
// 신청 관련
// =====================================================================

function submitBooking(params) {
  const { grade, classNum, studentNum, studentName, programId } = params;

  // 입력값 검증
  if (!grade || !classNum || !studentNum || !studentName || !programId) {
    return { success: false, message: '모든 항목을 입력해 주세요.' };
  }
  if (!/^[가-힣]{2,5}$/.test(studentName)) {
    return { success: false, message: '이름은 한글 2~5자로 입력해 주세요.' };
  }
  if (Number(studentNum) < 1 || Number(studentNum) > 40) {
    return { success: false, message: '번호는 1~40 사이로 입력해 주세요.' };
  }

  // 오픈 시간 확인
  const programs = getAllData('programs');
  const targetProgram = programs.find(p => p.id === programId);
  if (!targetProgram) return { success: false, message: '존재하지 않는 프로그램입니다.' };

  const openCheck = checkOpenTime({ date: targetProgram.date });
  if (!openCheck.isOpen) {
    return { success: false, message: '아직 예매 오픈 전입니다.' };
  }

  // LockService로 동시성 제어
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (e) {
    return { success: false, message: '서버가 혼잡합니다. 잠시 후 다시 시도해 주세요.' };
  }

  try {
    const bookings = getAllData('bookings');

    // 중복 신청 확인
    const alreadyBooked = bookings.find(b =>
      String(b.grade) === String(grade) &&
      String(b.classNum) === String(classNum) &&
      String(b.studentNum) === String(studentNum)
    );
    if (alreadyBooked) {
      const prog = programs.find(p => p.id === alreadyBooked.programId);
      return {
        success: false,
        message: `이미 "${prog ? prog.name : alreadyBooked.programId}" 프로그램에 신청하셨습니다.`,
        alreadyBooked: true
      };
    }

    // 정원 확인
    const bookedCount = bookings.filter(b => b.programId === programId).length;
    if (bookedCount >= Number(targetProgram.capacity)) {
      return { success: false, message: '정원이 마감되었습니다.' };
    }

    // 신청 등록
    const newBooking = {
      id: 'b' + String(Date.now()),
      grade: String(grade),
      classNum: String(classNum),
      studentNum: String(studentNum),
      studentName: String(studentName),
      programId: String(programId),
      bookedAt: new Date().toISOString()
    };
    appendRow('bookings', newBooking);

    return {
      success: true,
      message: '신청이 완료되었습니다.',
      booking: {
        ...newBooking,
        programName: targetProgram.name,
        programDate: targetProgram.date
      }
    };
  } finally {
    lock.releaseLock();
  }
}

function checkBooking(params) {
  const { grade, classNum, studentNum } = params;
  if (!grade || !classNum || !studentNum) {
    return { success: false, message: '학년, 반, 번호를 입력해 주세요.' };
  }

  const bookings = getAllData('bookings');
  const programs = getAllData('programs');

  const booking = bookings.find(b =>
    String(b.grade) === String(grade) &&
    String(b.classNum) === String(classNum) &&
    String(b.studentNum) === String(studentNum)
  );

  if (!booking) {
    return { success: true, found: false, message: '신청 내역이 없습니다.' };
  }

  const program = programs.find(p => p.id === booking.programId);
  return {
    success: true,
    found: true,
    booking: {
      ...booking,
      programName: program ? program.name : booking.programId,
      programDate: program ? program.date : ''
    }
  };
}

function getBookings(params) {
  const { token } = params;
  if (!validateToken(token)) return { success: false, message: '인증이 필요합니다.' };

  const bookings = getAllData('bookings');
  const programs = getAllData('programs');

  const result = bookings.map(b => {
    const prog = programs.find(p => p.id === b.programId);
    return {
      ...b,
      programName: prog ? prog.name : b.programId,
      programDate: prog ? prog.date : ''
    };
  });

  return { success: true, bookings: result };
}

function cancelBooking(params) {
  const { token, bookingId } = params;
  if (!validateToken(token)) return { success: false, message: '인증이 필요합니다.' };

  const sheet = getSheet('bookings');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: false, message: '신청 내역이 없습니다.' };
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === bookingId) {
      sheet.deleteRow(i + 1);
      return { success: true, message: '신청이 취소되었습니다.' };
    }
  }
  return { success: false, message: '신청 내역을 찾을 수 없습니다.' };
}

function manualBooking(params) {
  const { token, grade, classNum, studentNum, studentName, programId } = params;
  if (!validateToken(token)) return { success: false, message: '인증이 필요합니다.' };

  // 중복 확인
  const bookings = getAllData('bookings');
  const alreadyBooked = bookings.find(b =>
    String(b.grade) === String(grade) &&
    String(b.classNum) === String(classNum) &&
    String(b.studentNum) === String(studentNum)
  );
  if (alreadyBooked && !params.force) {
    return { success: false, message: '이미 신청 내역이 있습니다.', needConfirm: true };
  }

  const newBooking = {
    id: 'b' + String(Date.now()),
    grade: String(grade),
    classNum: String(classNum),
    studentNum: String(studentNum),
    studentName: String(studentName),
    programId: String(programId),
    bookedAt: new Date().toISOString()
  };
  appendRow('bookings', newBooking);
  return { success: true, message: '수동 신청이 완료되었습니다.' };
}

// =====================================================================
// 설정 관련
// =====================================================================

function getSettings(params) {
  const { token } = params;
  if (!validateToken(token)) return { success: false, message: '인증이 필요합니다.' };

  const data = getAllData('settings');
  const settings = {};
  data.forEach(r => {
    if (!r.key.startsWith('adminPassword') && !r.key.startsWith('session')) {
      settings[r.key] = r.value;
    }
  });
  return { success: true, settings };
}

function saveSettings(params) {
  const { token, settings } = params;
  if (!validateToken(token)) return { success: false, message: '인증이 필요합니다.' };

  Object.entries(settings).forEach(([key, value]) => {
    setSetting(key, value);
  });
  return { success: true, message: '설정이 저장되었습니다.' };
}

// =====================================================================
// 초기화 함수 (최초 1회 실행)
// =====================================================================

function initializeAll() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // 시트 생성 헬퍼
  function ensureSheet(name, headers) {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
    } else {
      sheet.clear();
    }
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#4a90d9')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    return sheet;
  }

  // programs 시트
  ensureSheet('programs', ['id', 'name', 'date', 'targetClasses', 'capacity', 'createdAt']);

  // bookings 시트
  ensureSheet('bookings', ['id', 'grade', 'classNum', 'studentNum', 'studentName', 'programId', 'bookedAt']);

  // settings 시트
  ensureSheet('settings', ['key', 'value']);

  // 초기 설정값 입력
  const settingsSheet = ss.getSheetByName('settings');
  const initialSettings = [
    ['adminPasswordHash', sha256(ADMIN_DEFAULT_PASSWORD)],
    ['openTime_2025-06-30', ''],
    ['openTime_2025-07-01', ''],
    ['openTime_2025-07-02', ''],
    ['forceOpen_2025-06-30', 'false'],
    ['forceOpen_2025-07-01', 'false'],
    ['forceOpen_2025-07-02', 'false'],
    ['forceClose_2025-06-30', 'false'],
    ['forceClose_2025-07-01', 'false'],
    ['forceClose_2025-07-02', 'false'],
    ['sessionToken', ''],
    ['sessionExpire', '']
  ];
  initialSettings.forEach(row => settingsSheet.appendRow(row));

  // 초기 프로그램 데이터
  const programsSheet = ss.getSheetByName('programs');
  const now = new Date().toISOString();
  const initialPrograms = [
    // 6월 30일 - 1, 4, 7반
    ['p001', '전동 선풍기 조립', '2025-06-30', '1,4,7', 30, now],
    ['p002', '조립 드론 만들기', '2025-06-30', '1,4,7', 30, now],
    ['p003', '미래 자동차와 코딩', '2025-06-30', '1,4,7', 30, now],
    ['p004', '디지털 굿즈 디자인', '2025-06-30', '1,4,7', 30, now],
    // 7월 1일 - 2, 5, 8반
    ['p005', 'AI캐릭터 포토카드 디자인 체험', '2025-07-01', '2,5,8', 30, now],
    ['p006', '버튼 제어형 Arduino 개발 실습', '2025-07-01', '2,5,8', 30, now],
    ['p007', '휴먼로봇 제작을 위한 코딩', '2025-07-01', '2,5,8', 30, now],
    ['p008', '3D프린팅 메이커 제작 체험', '2025-07-01', '2,5,8', 30, now],
    ['p009', '숏폼 제작을 위한 드론 활용', '2025-07-01', '2,5,8', 30, now],
    // 7월 2일 - 3, 6, 9반
    ['p010', '공유압제어 체험', '2025-07-02', '3,6,9', 30, now],
    ['p011', '미래 자동차를 만드는 3차원 측정', '2025-07-02', '3,6,9', 30, now],
    ['p012', 'MBTI 함수만들기', '2025-07-02', '3,6,9', 30, now],
    ['p013', '드론 비행 기초 입문', '2025-07-02', '3,6,9', 30, now]
  ];
  initialPrograms.forEach(row => programsSheet.appendRow(row));

  Logger.log('✅ 초기화 완료! programs: ' + initialPrograms.length + '개, settings: ' + initialSettings.length + '개');
  return '초기화가 완료되었습니다.';
}
