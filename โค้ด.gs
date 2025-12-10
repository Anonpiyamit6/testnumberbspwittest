const SHEET_ID = '1WNcL2AWLNvXH8ojJwWFDHpaBVXunxw4KycndV2RAEhY'; 
const SHEET_NAME = 'Students';

function doGet(e) { return ContentService.createTextOutput(JSON.stringify({ status: 'active' })).setMimeType(ContentService.MimeType.JSON); }

function doPost(e) {
  const lock = LockService.getScriptLock(); lock.tryLock(10000);
  try {
    const request = JSON.parse(e.postData.contents);
    const action = request.action;
    let result = {};

    if (action === 'getAllStudents') result = getAllStudents();
    else if (action === 'createStudent') result = createStudent(request.data);
    else if (action === 'createStudentsBulk') result = createStudentsBulk(request.data);
    else if (action === 'updateStudent') result = updateStudent(request.data);
    else if (action === 'updateScoresBulk') result = updateScoresBulk(request.data);
    else if (action === 'deleteStudent') result = deleteStudent(request.id);
    else if (action === 'deleteStudentsBulk') result = deleteStudentsBulk(request.ids);
    // --- เพิ่ม Action สำหรับประมวลผลจัดห้อง ---
    else if (action === 'runAdmission') result = runAdmission();

    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch (err) { return ContentService.createTextOutput(JSON.stringify({ success: false, message: err.toString() })).setMimeType(ContentService.MimeType.JSON); } finally { lock.releaseLock(); }
}

function getSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['ID','Exam ID','Full Name','Previous School','Grade Level','Thai','Math','Science','English','Aptitude','Total','Rank','National ID', 'Choice 1', 'Choice 2', 'Choice 3', 'Admission Result', 'Practical Score']);
  }
  return sheet;
}

function getAllStudents() {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: true, students: [] };
  const data = sheet.getRange(2, 1, lastRow - 1, 18).getDisplayValues();
  
  const students = data.map(row => ({
    id: row[0], 
    exam_id: String(row[1]).replace(/^'/, ''), 
    full_name: row[2], 
    previous_school: row[3], 
    grade_level: row[4],
    thai_score: parseFloat(row[5])||0, 
    math_score: parseFloat(row[6])||0, 
    science_score: parseFloat(row[7])||0, 
    english_score: parseFloat(row[8])||0, 
    aptitude_score: parseFloat(row[9])||0, 
    total_score: parseFloat(row[10])||0, 
    rank: parseInt(row[11])||0, 
    national_id: String(row[12]||'').replace(/^'/, ''),
    choice_1: String(row[13]||'').trim(), 
    choice_2: String(row[14]||'').trim(), 
    choice_3: String(row[15]||'').trim(), 
    admission_result: row[16]||'',
    practical_score: parseFloat(row[17])||0 
  }));
  return { success: true, students: students };
}

/**
 * ฟังก์ชันประมวลผลการคัดเลือกเข้าห้องเรียนโครงการ
 */
function runAdmission() {
  const sheet = getSheet();
  const res = getAllStudents();
  if (!res.success || res.students.length === 0) return { success: false, message: 'ไม่มีข้อมูลนักเรียน' };

  let students = res.students;
  
  // 1. แบ่งกลุ่ม ม.1 และ ม.4
  const m1Students = students.filter(s => ['p4','p5','p6','m1'].includes(s.grade_level.toLowerCase()));
  const m4Students = students.filter(s => ['m3','m4'].includes(s.grade_level.toLowerCase()));

  // 2. กำหนดโควตาที่นั่ง
  const quotaM1 = { 'S&M': 80, 'GATE': 60, 'EIS': 60 };
  const quotaM4 = { 'SME': 80, 'GATE': 60 };

  // ล้างค่าผลการคัดเลือกเดิม
  students.forEach(s => s.admission_result = 'ไม่ผ่านการคัดเลือก');

  // ฟังก์ชันบรรจุคนเข้าโครงการ
  function processLogic(group, quota) {
    let filled = {};
    Object.keys(quota).forEach(key => filled[key] = 0);

    // วน 3 รอบ ตามลำดับการเลือก (Round 1 = อันดับ 1)
    for (let round = 1; round <= 3; round++) {
      // เฉพาะคนที่ยังไม่มีที่เรียน
      let available = group.filter(s => s.admission_result === 'ไม่ผ่านการคัดเลือก');
      // เรียงคะแนนสูงไปต่ำ
      available.sort((a, b) => b.total_score - a.total_score);

      available.forEach(s => {
        let choice = s['choice_' + round];
        if (choice && quota[choice] !== undefined) {
          if (filled[choice] < quota[choice]) {
            s.admission_result = choice;
            filled[choice]++;
          }
        }
      });
    }
  }

  processLogic(m1Students, quotaM1);
  processLogic(m4Students, quotaM4);

  // 3. เขียนข้อมูลกลับลง Sheet (คอลัมน์ Q คือ index 16)
  const lastRow = sheet.getLastRow();
  const range = sheet.getRange(2, 1, lastRow - 1, 18);
  const rawValues = range.getValues();

  const studentMap = {};
  students.forEach(s => studentMap[s.id] = s.admission_result);

  const updatedValues = rawValues.map(row => {
    const id = row[0];
    if (studentMap[id]) row[16] = studentMap[id];
    return row;
  });

  range.setValues(updatedValues);
  return { success: true, message: 'ประมวลผลเสร็จสิ้น (ม.1: S&M 80, GATE 60, EIS 60 | ม.4: SME 80, GATE 60)' };
}

// ฟังก์ชัน CRUD อื่นๆ คงเดิม
function createStudent(data) {
  const sheet = getSheet(); const id = Utilities.getUuid();
  const newRow = [id, "'" + data.exam_id, data.full_name, data.previous_school, data.grade_level, data.thai_score||0, data.math_score||0, data.science_score||0, data.english_score||0, data.aptitude_score||0, data.total_score||0, data.rank||0, "'" + (data.national_id||''), data.choice_1||'', data.choice_2||'', data.choice_3||'', '', data.practical_score||0];
  sheet.appendRow(newRow); return { success: true, message: 'บันทึกสำเร็จ', id: id };
}
function createStudentsBulk(students) {
  const sheet = getSheet(); if (students.length === 0) return { success: false, message: 'ไม่พบข้อมูล' };
  const newRows = students.map(data => [Utilities.getUuid(), "'" + data.exam_id, data.full_name, data.previous_school, data.grade_level, data.thai_score||0, data.math_score||0, data.science_score||0, data.english_score||0, data.aptitude_score||0, data.total_score||0, data.rank||0, "'" + (data.national_id||''), data.choice_1||'', data.choice_2||'', data.choice_3||'', '', data.practical_score||0]);
  sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, 18).setValues(newRows); return { success: true, message: `นำเข้าสำเร็จ ${newRows.length} รายการ` };
}
function updateStudent(data) {
  const sheet = getSheet(); const allData = sheet.getDataRange().getValues();
  for (let i = 1; i < allData.length; i++) {
    if (allData[i][0] === data.id) {
      sheet.getRange(i + 1, 1, 1, 18).setValues([[data.id, "'" + data.exam_id, data.full_name, data.previous_school, data.grade_level, data.thai_score||0, data.math_score||0, data.science_score||0, data.english_score||0, data.aptitude_score||0, data.total_score||0, data.rank||0, "'" + (data.national_id||''), data.choice_1||'', data.choice_2||'', data.choice_3||'', data.admission_result||'', data.practical_score||0]]);
      return { success: true, message: 'อัปเดตสำเร็จ' };
    }
  } return { success: false, message: 'ไม่พบข้อมูล' };
}
function deleteStudent(id) { const sheet = getSheet(); const allData = sheet.getDataRange().getValues(); for (let i = 1; i < allData.length; i++) { if (allData[i][0] === id) { sheet.deleteRow(i + 1); return { success: true, message: 'ลบสำเร็จ' }; } } return { success: false, message: 'ไม่พบข้อมูล' }; }
function deleteStudentsBulk(ids) { const sheet = getSheet(); const data = sheet.getDataRange().getValues(); let count=0; for(let i=data.length-1; i>=1; i--){ if(ids.includes(data[i][0])){ sheet.deleteRow(i+1); count++; } } return count>0 ? {success:true, message:`ลบ ${count} รายการ`} : {success:false, message:'ไม่พบข้อมูล'}; }
function updateScoresBulk(d){ const s=getSheet(); const r=s.getDataRange(); const v=r.getValues(); let m={}; for(let i=1;i<v.length;i++) m[String(v[i][1]).replace(/^'/,'').trim()]=i; let c=0; d.forEach(x=>{ let k=String(x.exam_id).trim(); if(m[k]){ let i=m[k]; v[i][6]=x.math_score; v[i][7]=x.science_score; v[i][8]=x.english_score; v[i][17]=x.practical_score||0; v[i][10]=parseFloat(x.math_score)+parseFloat(x.science_score)+parseFloat(x.english_score)+(parseFloat(x.practical_score)||0); c++; } }); if(c>0){r.setValues(v); return{success:true, message:`อัปเดต ${c} รายการ`};} return{success:false, message:'ไม่พบข้อมูล'}; }
