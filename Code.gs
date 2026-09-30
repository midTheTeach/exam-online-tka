const SPREADSHEET_ID = 'GANTI_DENGAN_ID_SPREADSHEET_ANDA';

function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Ujian Online TKA')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function checkStudentExists(studentName) {
  const name = String(studentName || '').trim().toUpperCase();
  if (!name) return { exists: false };
  const sheet = getSpreadsheet().getSheetByName('Peserta');
  if (!sheet || sheet.getLastRow() < 2) return { exists: false };
  const names = sheet.getRange(2, 3, sheet.getLastRow() - 1, 1).getDisplayValues();
  const exists = names.some(row => String(row[0]).trim().toUpperCase() === name);
  return {
    exists: exists,
    message: exists ? 'Nama tersebut sudah menyelesaikan ujian dan tidak dapat mengerjakan lagi.' : ''
  };
}

function getQuestions() {
  try {
    const sheet = getSpreadsheet().getSheetByName('Soal');
    if (!sheet) return { error: "Sheet 'Soal' tidak ditemukan." };
    if (sheet.getLastRow() < 2) return { error: 'Sheet Soal masih kosong.' };
    return sheet.getRange(2, 1, sheet.getLastRow() - 1, 17).getValues()
      .filter(row => row[0])
      .map(parseQuestion)
      .filter(Boolean);
  } catch (e) {
    return { error: String(e) };
  }
}

function parseQuestion(row) {
  const [id,type,question,a,b,c,d,e,correct,detail,min,max,tableRows,tableAnswers,left,right,matchAnswers] = row;
  if (!id || !type || !question) return null;
  const q = { id: Number(id), type: String(type).trim().toLowerCase(), question: String(question) };
  if (q.type === 'multiple_choice' || q.type === 'multiple_choice_complex') {
    q.options = [a,b,c,d,e].map((text,i) => ({key:String.fromCharCode(65+i),text:String(text || '')})).filter(x => x.text.trim());
    q.correctAnswer = String(correct || '').split(',').map(x => x.trim().toUpperCase()).filter(Boolean);
    if (q.type === 'multiple_choice') q.correctAnswer = q.correctAnswer[0] || '';
    q.minCorrect = Number(min) || 2; q.maxCorrect = Number(max) || 3;
  } else if (q.type === 'true_false') {
    q.correctAnswer = String(correct || 'B').trim().toUpperCase();
  } else if (q.type === 'true_false_table') {
    q.rows = String(tableRows || '').split('|').filter(Boolean).map((text,i) => ({label:i+1,text:text.trim()}));
    q.correctAnswer = String(tableAnswers || '').split(',').map(x => x.trim().toUpperCase());
  } else if (q.type === 'matching') {
    q.leftItems = String(left || '').split('|').filter(Boolean).map((text,i) => ({id:i+1,text:text.trim()}));
    q.rightItems = String(right || '').split('|').filter(Boolean).map((text,i) => ({id:String.fromCharCode(65+i),text:text.trim()}));
    q.correctAnswer = String(matchAnswers || '').split(',').map(x => x.trim());
  } else if (q.type === 'essay') {
    q.placeholder = String(detail || 'Tulis jawaban Anda di sini...');
  } else return null;
  return q;
}

function submitAnswers(studentClass, studentName, answers) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const name = String(studentName || '').trim().toUpperCase();
    const kelas = String(studentClass || '').trim().toUpperCase();
    if (!name || !kelas) return {success:false,error:'Nama dan kelas wajib diisi.'};
    const duplicate = checkStudentExists(name);
    if (duplicate.exists) return {success:false,error:duplicate.message};
    const questions = getQuestions();
    if (questions.error) return {success:false,error:questions.error};
    let score = 0, details = [];
    questions.forEach((q,i) => {
      const correct = checkAnswer(q, answers[i]);
      if (correct) score++;
      details.push({questionId:q.id,answer:answers[i],isCorrect:correct});
    });
    const sheet = getSpreadsheet().getSheetByName('Peserta');
    if (!sheet) return {success:false,error:"Sheet 'Peserta' tidak ditemukan."};
    sheet.appendRow([new Date(), kelas, name, score, questions.length,
      (score / questions.length * 100).toFixed(2), JSON.stringify(details)]);
    return {success:true,score:score,total:questions.length,
      percentage:(score / questions.length * 100).toFixed(2),details:details};
  } finally { lock.releaseLock(); }
}

function checkAnswer(q, answer) {
  if (answer === undefined || answer === null || answer === '') return false;
  if (q.type === 'essay') return false;
  if (q.type === 'multiple_choice' || q.type === 'true_false') return String(answer).toUpperCase() === String(q.correctAnswer).toUpperCase();
  if (q.type === 'multiple_choice_complex') return sameSet(answer, q.correctAnswer);
  if (q.type === 'true_false_table') return Array.isArray(answer) && answer.length === q.correctAnswer.length && answer.every((x,i) => String(x).toUpperCase() === q.correctAnswer[i]);
  if (q.type === 'matching') return Array.isArray(answer) && answer.length === q.correctAnswer.length && answer.every((x,i) => String(x).trim() === q.correctAnswer[i]);
  return false;
}

function sameSet(a,b) {
  const x = (Array.isArray(a) ? a : [a]).map(String).map(s => s.toUpperCase()).sort();
  const y = (Array.isArray(b) ? b : [b]).map(String).map(s => s.toUpperCase()).sort();
  return x.length === y.length && x.every((v,i) => v === y[i]);
}