/**
 * 서울제일교회 주차관리앱 — 구글 시트 연동 스크립트 (Google Apps Script)
 *
 * 두 가지 사용자
 *  - 교인    : MEMBER_KEY 로 접속. 자기 차량 등록, 주차 신청, 내 신청 보기·취소
 *  - 관리자  : ADMIN_KEY  로 접속. 전체 명부, 차주 찾기·연락, 신청 승인/반려, 연락 기록
 *
 * 설치
 * 1) 구글 시트 > 확장 프로그램 > Apps Script 에 이 파일 내용을 붙여넣고 저장
 * 2) 함수 목록에서 setup 선택 → 실행 (권한 허용) — 시트 3개와 머리글이 만들어집니다
 * 3) 프로젝트 설정(톱니바퀴) > 스크립트 속성 에 세 가지 추가
 *      MEMBER_KEY = 교인용 암호 (주보·단톡방에 알릴 것)
 *      ADMIN_KEY  = 관리자용 암호 (사무실·주차봉사팀만)
 *      CAPACITY   = 하루 주차 가능 대수 (예: 40)
 * 4) 배포 > 새 배포 > 웹 앱 / 실행: 나 / 액세스 권한: 모든 사용자 → 웹 앱 URL 복사
 *    (코드를 고친 뒤에는 배포 > 배포 관리 > 수정 > 새 버전 으로 다시 배포)
 */

const SHEET_CARS = '차량';
const SHEET_APPS = '주차신청';
const SHEET_LOGS = '연락기록';
const CAR_HEADERS = ['ID', '차량번호', '이름', '연락처', '소속', '차종', '메모', '수정일'];
const APP_HEADERS = ['ID', '신청일시', '주차날짜', '차량번호', '이름', '연락처', '사유', '상태', '배정구역', '관리메모'];
const LOG_HEADERS = ['일시', '차량번호', '이름', '방법', '담당자'];
const STATUS = { WAIT: '대기', OK: '승인', NO: '반려', CANCEL: '취소' };

function setup() {
  const ss = SpreadsheetApp.getActive();
  [[SHEET_CARS, CAR_HEADERS], [SHEET_APPS, APP_HEADERS], [SHEET_LOGS, LOG_HEADERS]].forEach(([name, headers]) => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  });
  // 번호·날짜를 글자로 저장 (앞자리 0, 날짜 형식이 바뀌지 않도록)
  ss.getSheetByName(SHEET_CARS).getRange('B:D').setNumberFormat('@');
  ss.getSheetByName(SHEET_APPS).getRange('C:F').setNumberFormat('@');
}

function doGet() {
  return json({ ok: true, message: '서울제일교회 주차관리 API가 작동 중입니다.' });
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); }
  catch (err) { return json({ ok: false, error: '잘못된 요청입니다.' }); }

  const props = PropertiesService.getScriptProperties();
  const role = req.key && req.key === props.getProperty('ADMIN_KEY') ? 'admin'
             : req.key && req.key === props.getProperty('MEMBER_KEY') ? 'member' : null;
  if (!role) return json({ ok: false, error: '접속 암호가 맞지 않습니다.' });

  const admin = role === 'admin';
  const need = () => { if (!admin) throw new Error('관리자만 할 수 있습니다.'); };

  try {
    let data;
    switch (req.action) {
      case 'whoami':    data = { role, capacity: capacity() }; break;

      // 교인·관리자 공통
      case 'myCars':    data = listCars().filter(c => samePhone(c.phone, req.phone)); break;
      case 'saveMyCar': data = withLock(() => saveMyCar(req.car, req.phone)); break;
      case 'removeMyCar': data = withLock(() => removeMyCar(req.id, req.phone)); break;
      case 'apply':     data = withLock(() => applyParking(req.app)); break;
      case 'myApps':    data = listApps().filter(a => samePhone(a.phone, req.phone)); break;
      case 'cancelApp': data = withLock(() => cancelApp(req.id, req.phone)); break;
      case 'dayCount':  data = dayCount(req.date); break;

      // 관리자 전용
      case 'list':      need(); data = listCars(); break;
      case 'save':      need(); data = withLock(() => saveCar(req.car)); break;
      case 'remove':    need(); data = withLock(() => deleteById(SHEET_CARS, req.id)); break;
      case 'apps':      need(); data = listApps().filter(a => !req.date || a.date === req.date); break;
      case 'setApp':    need(); data = withLock(() => setApp(req.id, req.status, req.spot, req.note)); break;
      case 'log':       need(); data = withLock(() => addLog(req.log)); break;
      case 'logs':      need(); data = listLogs(req.limit || 100); break;
      default: throw new Error('알 수 없는 요청: ' + req.action);
    }
    return json({ ok: true, data });
  } catch (err) {
    return json({ ok: false, error: err.message || String(err) });
  }
}

/* ---------- 차량 ---------- */

function listCars() {
  return rows(SHEET_CARS, CAR_HEADERS.length).map(r => ({
    id: r[0], plate: r[1], name: r[2], phone: r[3], group: r[4], model: r[5], memo: r[6], updated: r[7]
  }));
}

function saveCar(car) {
  if (!car || !car.plate || !car.name) throw new Error('차량번호와 이름은 꼭 입력해 주세요.');
  const sh = sheet(SHEET_CARS);
  const id = car.id || newId();
  const row = [id, clean(car.plate), clean(car.name), clean(car.phone), clean(car.group),
               clean(car.model), clean(car.memo), now()];
  const r = findRow(sh, id);
  if (r) sh.getRange(r, 1, 1, row.length).setValues([row]);
  else {
    const dup = listCars().find(c => normPlate(c.plate) === normPlate(car.plate));
    if (dup) throw new Error(`${car.plate}는 이미 ${dup.name} 님 차량으로 등록되어 있습니다.`);
    sh.appendRow(row);
  }
  return { id };
}

// 교인은 자기 연락처로 등록된 차량만 만들고 고칠 수 있음
function saveMyCar(car, phone) {
  if (!digits(phone)) throw new Error('설정에서 내 연락처를 먼저 입력해 주세요.');
  car.phone = phone;
  if (car.id) {
    const old = listCars().find(c => c.id === car.id);
    if (!old || !samePhone(old.phone, phone)) throw new Error('내 차량만 고칠 수 있습니다.');
  }
  return saveCar(car);
}

function removeMyCar(id, phone) {
  const old = listCars().find(c => c.id === id);
  if (!old || !samePhone(old.phone, phone)) throw new Error('내 차량만 삭제할 수 있습니다.');
  return deleteById(SHEET_CARS, id);
}

/* ---------- 주차 신청 ---------- */

function listApps() {
  return rows(SHEET_APPS, APP_HEADERS.length).map(r => ({
    id: r[0], at: r[1], date: r[2], plate: r[3], name: r[4], phone: r[5],
    reason: r[6], status: r[7], spot: r[8], note: r[9]
  })).sort((a, b) => (a.date < b.date ? 1 : a.date > b.date ? -1 : (a.at < b.at ? 1 : -1)));
}

function applyParking(app) {
  if (!app || !app.date || !app.plate || !app.name || !digits(app.phone))
    throw new Error('날짜, 차량, 이름, 연락처를 모두 넣어 주세요.');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(app.date)) throw new Error('날짜 형식이 맞지 않습니다.');
  const today = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
  if (app.date < today) throw new Error('지난 날짜는 신청할 수 없습니다.');

  const live = listApps().filter(a => a.date === app.date && a.status !== STATUS.NO && a.status !== STATUS.CANCEL);
  if (live.some(a => normPlate(a.plate) === normPlate(app.plate)))
    throw new Error('이 날짜에 이미 신청한 차량입니다.');

  const id = newId();
  sheet(SHEET_APPS).appendRow([id, now(), app.date, clean(app.plate), clean(app.name),
    clean(app.phone), clean(app.reason), STATUS.WAIT, '', '']);
  const cap = capacity();
  return { id, waiting: live.length + 1, capacity: cap, full: cap > 0 && live.length >= cap };
}

function cancelApp(id, phone) {
  const a = listApps().find(x => x.id === id);
  if (!a || !samePhone(a.phone, phone)) throw new Error('내 신청만 취소할 수 있습니다.');
  return setApp(id, STATUS.CANCEL, a.spot, a.note);
}

function setApp(id, status, spot, note) {
  if (!Object.values(STATUS).includes(status)) throw new Error('상태 값이 맞지 않습니다.');
  const sh = sheet(SHEET_APPS);
  const r = findRow(sh, id);
  if (!r) throw new Error('신청을 찾을 수 없습니다.');
  sh.getRange(r, 8, 1, 3).setValues([[status, clean(spot), clean(note)]]);
  return { id, status };
}

function dayCount(date) {
  const day = listApps().filter(a => a.date === date);
  return {
    date, capacity: capacity(),
    approved: day.filter(a => a.status === STATUS.OK).length,
    waiting: day.filter(a => a.status === STATUS.WAIT).length
  };
}

/* ---------- 연락 기록 ---------- */

function addLog(log) {
  sheet(SHEET_LOGS).appendRow([now(), clean(log.plate), clean(log.name), clean(log.method), clean(log.by)]);
  return {};
}

function listLogs(limit) {
  return rows(SHEET_LOGS, LOG_HEADERS.length).reverse().slice(0, limit)
    .map(r => ({ at: r[0], plate: r[1], name: r[2], method: r[3], by: r[4] }));
}

/* ---------- 도우미 ---------- */

function sheet(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error(`'${name}' 시트가 없습니다. setup 함수를 먼저 실행해 주세요.`);
  return sh;
}

function rows(name, width) {
  const sh = sheet(name);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, width).getDisplayValues().filter(r => r[0] || r[1]);
}

function findRow(sh, id) {
  if (!id) return 0;
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const ids = sh.getRange(2, 1, last - 1, 1).getDisplayValues();
  for (let i = 0; i < ids.length; i++) if (ids[i][0] === String(id)) return i + 2;
  return 0;
}

function deleteById(name, id) {
  const sh = sheet(name);
  const r = findRow(sh, id);
  if (r) sh.deleteRow(r);
  return { id };
}

function capacity() {
  return Number(PropertiesService.getScriptProperties().getProperty('CAPACITY') || 0);
}

function withLock(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try { return fn(); } finally { lock.releaseLock(); }
}

const digits = s => String(s || '').replace(/\D/g, '');
const normPlate = s => String(s || '').replace(/\s|-/g, '');
const samePhone = (a, b) => digits(a).length >= 8 && digits(a) === digits(b);
const newId = () => Utilities.getUuid().slice(0, 8);
const now = () => Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm');

// 수식 주입 방지: = + - @ 로 시작하는 값 앞에 ' 를 붙임
function clean(v) {
  const s = String(v == null ? '' : v).trim();
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
