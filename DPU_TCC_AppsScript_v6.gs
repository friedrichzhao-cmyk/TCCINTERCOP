// ════════════════════════════════════════════════════════════════
//  DPU–TCC Apps Script v6
//  修复: 已填写课程无法被重新编辑/覆盖的问题
//  根本原因: v5的touched判断会忽略"清空"或"改回pending"的操作
// ════════════════════════════════════════════════════════════════

const SHEET_ID   = '1mkWCr08ctBFO-E7wZrZlbrjWOBUbTp19UuXSUoEHFi8';
const MAIN_SHEET = 'CourseConfirmation';
const LOG_SHEET  = 'Log';

// ── doGet ────────────────────────────────────────────────────────
function doGet(e) {
  if (e.parameter.action === 'get') return getCurrentState();
  return jsonResponse({error: 'Unknown action'});
}

// ── doPost ───────────────────────────────────────────────────────
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.action === 'set') return saveChanges(body.data);
    return jsonResponse({error: 'Unknown action'});
  } catch(err) {
    return jsonResponse({error: err.toString()});
  }
}

// ── getCurrentState ───────────────────────────────────────────────
function getCurrentState() {
  const sheet  = getSheet(MAIN_SHEET);
  const rows   = sheet.getDataRange().getValues();
  const latest = {};

  for (let i = 1; i < rows.length; i++) {
    const [ts, courseId, teacher, status, submitter, email] = rows[i];
    if (courseId) {
      latest[courseId] = {
        teacher: teacher || '',
        status:  status  || 'pending',
        email:   email   || '',
      };
    }
  }
  return jsonResponse(latest);
}

// ── saveChanges ────────────────────────────────────────────────────
// ✅ 修复: 不再用 touched 判断是否记录。
//    只要这门课在本次提交里出现过（前端只发送有变化的课程），
//    就一律写入新行 —— 包括清空姓名、改回 pending 的情况。
//    "覆盖"的本质 = 新的一行记录，getCurrentState 永远取最后一行。
function saveChanges(payload) {
  if (!payload || !payload.courses) {
    return jsonResponse({error: 'Missing courses data'});
  }

  const sheet     = getSheet(MAIN_SHEET);
  const submitter = payload.submitter || '未署名';
  const timestamp = new Date();
  const courses   = payload.courses;
  const changedIds = [];

  // 读取当前最新状态，用于对比是否真的有变化
  const current = JSON.parse(getCurrentState().getContent());

  Object.entries(courses).forEach(([id, val]) => {
    const teacher = (val.teacher || '').trim();
    const status  = val.status  || 'pending';
    const email   = (val.email  || '').trim();

    const prev = current[id] || {teacher:'', status:'pending', email:''};

    // ✅ 新判断: 只要任何字段与"当前最新状态"不同，就记录
    //    这样允许: 改名字、清空、改状态(含改回pending)、改email
    const changed = (teacher !== prev.teacher) ||
                    (status  !== prev.status)  ||
                    (email   !== prev.email);

    if (changed) {
      sheet.appendRow([timestamp, id, teacher, status, submitter, email]);
      changedIds.push(id);
    }
  });

  if (changedIds.length > 0) {
    writeLog(submitter, timestamp, changedIds);
  }

  return jsonResponse({ok: true, updated: changedIds.length, changedIds, by: submitter});
}

// ── getSheet ────────────────────────────────────────────────────────
function getSheet(name) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  let   sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (name === MAIN_SHEET) {
      sheet.appendRow([
        '时间戳 Timestamp', '课程ID CourseID',
        '外派教师 Teacher (Passport Name)', '状态 Status',
        '填写人工号 Submitter ID', 'DPU E-Mail'
      ]);
      sheet.getRange('A1:F1').setFontWeight('bold')
           .setBackground('#1A1A2E').setFontColor('#FFFFFF');
      [180,160,180,120,130,200].forEach((w,i)=>sheet.setColumnWidth(i+1,w));
    } else if (name === LOG_SHEET) {
      sheet.appendRow([
        '时间戳 Timestamp', '填写人工号 Submitter ID',
        '更新课程数 Count', '改动课程ID Changed Course IDs'
      ]);
      sheet.getRange('A1:D1').setFontWeight('bold')
           .setBackground('#2D2875').setFontColor('#FFFFFF');
      [180,130,100,320].forEach((w,i)=>sheet.setColumnWidth(i+1,w));
    }
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ── writeLog ──────────────────────────────────────────────────────
function writeLog(submitter, timestamp, changedIds) {
  try {
    const log = getSheet(LOG_SHEET);
    log.appendRow([timestamp, submitter, changedIds.length, changedIds.join(', ')]);
  } catch(e) {}
}

// ── jsonResponse ──────────────────────────────────────────────────
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── 测试函数 ──────────────────────────────────────────────────────
function testOverwrite() {
  // 第一次提交
  saveChanges({
    submitter: '111111',
    courses: { 'EN_25B1_AD': {teacher: '张三', status: 'confirmed', email: 'a@dpu.ac.th'} }
  });
  // 第二次提交 — 修改同一门课
  const result = saveChanges({
    submitter: '222222',
    courses: { 'EN_25B1_AD': {teacher: '李四', status: 'pending', email: 'b@dpu.ac.th'} }
  });
  Logger.log(result.getContent());
  // 应该记录2次（2行），getCurrentState应返回李四的最新数据
  Logger.log(getCurrentState().getContent());
}

function testRead() {
  Logger.log(getCurrentState().getContent());
}

/*
════════════════════════════════════════════════════════
  DEPLOY GUIDE v6 — 修复"无法覆盖已填写内容"问题

  Step 1 — Apps Script 编辑器
  全选旧代码 → 删除 → 粘贴本文件全部内容 → Ctrl+S

  Step 2 — 测试
  函数选 testOverwrite → ▶ 运行
  查看日志 (View → Logs):
  应看到第二次返回 updated:1，changedIds含EN_25B1_AD
  去Sheet检查: 应有2行EN_25B1_AD记录(张三+李四)

  Step 3 — 部署新版本
  部署 → 管理部署 → ✏️ → 版本选「新版本」→ 部署

  ──────────────────────────────────────────────────
  v5 → v6 核心修复说明:
  v5用 "teacher!=='' || status!=='pending' || email!==''"
  判断是否记录 → 如果老师想清空姓名或改回pending，
  这个判断为false，导致修改被忽略，看起来像"无法覆盖"。

  v6改为: 对比"新值"与"当前Sheet最新值"是否不同，
  只要不同就记录 → 任何修改(包括清空、改回pending)
  都能被正确保存。
════════════════════════════════════════════════════════
*/
