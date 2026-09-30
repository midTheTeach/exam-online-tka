const SPREADSHEET_ID = 'GANTI_ID_SPREADSHEET_ANDA_DI_SINI';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Ujian Online TKA')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getSpreadsheet() {
  try {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch (e) {
    Logger.log('Error opening spreadsheet: ' + e);
    throw new Error('Spreadsheet tidak ditemukan. Periksa SPREADSHEET_ID di Code.gs');
  }
}

function checkStudentExists(studentName) {
  try {
    const name = String(studentName || '').trim().toUpperCase();
    if (!name) return { exists: false, message: '' };
    
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('Peserta');
    
    if (!sheet) return { exists: false, message: '' };
    if (sheet.getLastRow() < 2) return { exists: false, message: '' };
    
    const data = sheet.getRange(2, 3, sheet.getLastRow() - 1, 1).getDisplayValues();
    const found = data.some(row => String(row[0] || '').trim().toUpperCase() === name);
    
    return {
      exists: found,
      message: found ? 'Nama ini sudah pernah ujian. Setiap peserta hanya boleh ujian 1 kali.' : ''
    };
  } catch (e) {
    Logger.log('checkStudentExists error: ' + e);
    return { exists: false, message: '' };
  }
}

function getQuestions() {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('Soal');
    
    if (!sheet) {
      return { error: 'Sheet "Soal" tidak ditemukan' };
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return { error: 'Sheet Soal kosong' };
    }
    
    const data = sheet.getRange(2, 1, lastRow - 1, 17).getValues();
    const questions = [];
    
    data.forEach(row => {
      if (row[0]) {
        const q = parseQuestion(row);
        if (q) questions.push(q);
      }
    });
    
    return questions.length > 0 ? questions : { error: 'Tidak ada soal di sheet Soal' };
  } catch (e) {
    Logger.log('getQuestions error: ' + e);
    return { error: 'Error: ' + String(e) };
  }
}

function parseQuestion(row) {
  try {
    const [id, type, question, optA, optB, optC, optD, optE, correct, detail, min, max, tableRows, tableAnswers, left, right, matchAnswers] = row;
    
    if (!id || !type || !question) return null;
    
    const typeStr = String(type).toLowerCase().trim();
    const q = {
      id: Number(id),
      type: typeStr,
      question: String(question)
    };
    
    if (typeStr === 'multiple_choice') {
      q.options = [optA, optB, optC, optD]
        .map((text, i) => ({ key: String.fromCharCode(65 + i), text: String(text || '') }))
        .filter(x => x.text.trim());
      q.correctAnswer = String(correct || '').trim().toUpperCase();
    }
    else if (typeStr === 'multiple_choice_complex') {
      q.options = [optA, optB, optC, optD, optE]
        .map((text, i) => ({ key: String.fromCharCode(65 + i), text: String(text || '') }))
        .filter(x => x.text.trim());
      q.correctAnswer = String(correct || '').split(',').map(x => x.trim().toUpperCase()).filter(x => x);
      q.minCorrect = Number(min) || 2;
      q.maxCorrect = Number(max) || 3;
    }
    else if (typeStr === 'true_false') {
      q.correctAnswer = String(correct || 'B').trim().toUpperCase();
    }
    else if (typeStr === 'true_false_table') {
      q.rows = String(tableRows || '').split('|').filter(x => x).map((text, i) => ({ label: i + 1, text: text.trim() }));
      q.correctAnswer = String(tableAnswers || '').split(',').map(x => x.trim().toUpperCase());
    }
    else if (typeStr === 'matching') {
      q.leftItems = String(left || '').split('|').filter(x => x).map((text, i) => ({ id: i + 1, text: text.trim() }));
      q.rightItems = String(right || '').split('|').filter(x => x).map((text, i) => ({ id: String.fromCharCode(65 + i), text: text.trim() }));
      q.correctAnswer = String(matchAnswers || '').split(',').map(x => x.trim());
    }
    else if (typeStr === 'essay') {
      q.placeholder = String(detail || 'Tulis jawaban Anda...');
    }
    else {
      return null;
    }
    
    return q;
  } catch (e) {
    Logger.log('parseQuestion error: ' + e);
    return null;
  }
}

function submitAnswers(studentClass, studentName, answers) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  
  try {
    const name = String(studentName || '').trim().toUpperCase();
    const kelas = String(studentClass || '').trim().toUpperCase();
    
    if (!name || !kelas) {
      return { success: false, error: 'Nama dan kelas harus diisi' };
    }
    
    const dup = checkStudentExists(name);
    if (dup.exists) {
      return { success: false, error: dup.message };
    }
    
    const ss = getSpreadsheet();
    const pesertaSheet = ss.getSheetByName('Peserta');
    
    if (!pesertaSheet) {
      return { success: false, error: 'Sheet Peserta tidak ditemukan' };
    }
    
    const questionsData = getQuestions();
    if (questionsData.error) {
      return { success: false, error: questionsData.error };
    }
    
    let score = 0;
    const details = [];
    
    questionsData.forEach((q, i) => {
      const isCorrect = checkAnswer(q, answers[i]);
      if (isCorrect) score++;
      details.push({
        questionId: q.id,
        answer: answers[i],
        isCorrect: isCorrect
      });
    });
    
    const total = questionsData.length;
    const percentage = (score / total * 100).toFixed(2);
    
    pesertaSheet.appendRow([
      new Date(),
      kelas,
      name,
      score,
      total,
      percentage,
      JSON.stringify(details)
    ]);
    
    return {
      success: true,
      score: score,
      total: total,
      percentage: percentage,
      details: details
    };
  } catch (e) {
    Logger.log('submitAnswers error: ' + e);
    return { success: false, error: 'Error: ' + String(e) };
  } finally {
    lock.releaseLock();
  }
}

function checkAnswer(question, answer) {
  if (!answer) return false;
  if (question.type === 'essay') return false;
  
  if (question.type === 'multiple_choice' || question.type === 'true_false') {
    return String(answer).toUpperCase() === String(question.correctAnswer).toUpperCase();
  }
  
  if (question.type === 'multiple_choice_complex') {
    const ans = Array.isArray(answer) ? answer : [answer];
    const correct = question.correctAnswer || [];
    if (ans.length !== correct.length) return false;
    const aSet = new Set(ans.map(x => String(x).toUpperCase()));
    const cSet = new Set(correct.map(x => String(x).toUpperCase()));
    return aSet.size === cSet.size && [...aSet].every(x => cSet.has(x));
  }
  
  if (question.type === 'true_false_table') {
    const ans = Array.isArray(answer) ? answer : [];
    const correct = question.correctAnswer || [];
    if (ans.length !== correct.length) return false;
    return ans.every((x, i) => String(x).toUpperCase() === String(correct[i]).toUpperCase());
  }
  
  if (question.type === 'matching') {
    const ans = Array.isArray(answer) ? answer : [];
    const correct = question.correctAnswer || [];
    if (ans.length !== correct.length) return false;
    return ans.every((x, i) => String(x).trim() === String(correct[i]).trim());
  }
  
  return false;
}
