// ===== KONFIGURASI SPREADSHEET =====
// GANTI NILAI INI DENGAN SPREADSHEET ID ANDA
const SPREADSHEET_ID = '1abc2def3ghi4jkl5mno6pqr7stu8vwx'; // GANTI INI!

// Fungsi untuk mendapatkan spreadsheet
function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

// ===== MAIN FUNCTION =====
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ===== FUNGSI UNTUK MENGAMBIL SOAL =====
function getQuestions() {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('Soal');
    
    if (!sheet) {
      return { error: "Sheet 'Soal' tidak ditemukan. Pastikan sheet bernama 'Soal' sudah dibuat!" };
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return { error: "Sheet 'Soal' kosong. Silakan isi soal terlebih dahulu!" };
    }
    
    const data = sheet.getRange(2, 1, lastRow - 1, 17).getValues();
    const questions = [];
    
    data.forEach((row, index) => {
      if (row[0]) { // Jika ID soal ada
        const questionObj = parseQuestion(row);
        if (questionObj) {
          questions.push(questionObj);
        }
      }
    });
    
    return questions;
  } catch (e) {
    Logger.log("Error getQuestions: " + e);
    return { error: "Error: " + e.toString() };
  }
}

// ===== FUNGSI PARSE SOAL SESUAI TIPE =====
function parseQuestion(row) {
  const [id, type, question, optA, optB, optC, optD, optE, correctAnswer, 
          details, minCorrect, maxCorrect, tableRows, tableAnswers, 
          matchLeft, matchRight, matchAnswers] = row;
  
  if (!id || !type || !question) return null;
  
  const baseObj = {
    id: parseInt(id),
    type: type.toLowerCase().trim(),
    question: question.toString()
  };
  
  switch(type.toLowerCase().trim()) {
    case 'multiple_choice':
      return {
        ...baseObj,
        options: [
          { key: 'A', text: optA ? optA.toString() : '' },
          { key: 'B', text: optB ? optB.toString() : '' },
          { key: 'C', text: optC ? optC.toString() : '' },
          { key: 'D', text: optD ? optD.toString() : '' }
        ].filter(o => o.text && o.text.trim()),
        correctAnswer: correctAnswer ? correctAnswer.toString().toUpperCase().trim() : ''
      };
      
    case 'multiple_choice_complex':
      return {
        ...baseObj,
        options: [
          { key: 'A', text: optA ? optA.toString() : '' },
          { key: 'B', text: optB ? optB.toString() : '' },
          { key: 'C', text: optC ? optC.toString() : '' },
          { key: 'D', text: optD ? optD.toString() : '' },
          { key: 'E', text: optE ? optE.toString() : '' }
        ].filter(o => o.text && o.text.trim()),
        correctAnswer: correctAnswer ? correctAnswer.toString().split(',').map(x => x.trim().toUpperCase()) : [],
        minCorrect: parseInt(minCorrect) || 2,
        maxCorrect: parseInt(maxCorrect) || 3,
        instruction: 'Pilih ' + (parseInt(minCorrect) || 2) + '-' + (parseInt(maxCorrect) || 3) + ' jawaban yang benar!'
      };
      
    case 'true_false':
      return {
        ...baseObj,
        correctAnswer: correctAnswer ? correctAnswer.toString().toUpperCase().trim() : 'B'
      };
      
    case 'true_false_table':
      const rows = tableRows ? parseTableRows(tableRows.toString()) : [];
      const answers = tableAnswers ? tableAnswers.toString().split(',').map(x => x.trim().toUpperCase()) : [];
      return {
        ...baseObj,
        rows: rows,
        correctAnswer: answers
      };
      
    case 'matching':
      const leftItems = matchLeft ? matchLeft.toString().split('|').map((x, idx) => ({
        id: idx + 1,
        text: x.trim()
      })) : [];
      const rightItems = matchRight ? matchRight.toString().split('|').map((x, idx) => ({
        id: String.fromCharCode(65 + idx),
        text: x.trim()
      })) : [];
      const matchingAnswers = matchAnswers ? matchAnswers.toString().split(',').map(x => x.trim()) : [];
      return {
        ...baseObj,
        leftItems: leftItems,
        rightItems: rightItems,
        correctAnswer: matchingAnswers
      };
      
    case 'essay':
      return {
        ...baseObj,
        placeholder: details ? details.toString() : 'Tulis jawaban Anda di sini...'
      };
      
    default:
      return null;
  }
}

function parseTableRows(rowsString) {
  if (!rowsString) return [];
  return rowsString.split('|').map((item, index) => ({
    label: (index + 1).toString(),
    text: item.trim()
  }));
}

// ===== FUNGSI SUBMIT JAWABAN DAN HITUNG SKOR =====
function submitAnswers(studentEmail, studentName, answers) {
  try {
    const ss = getSpreadsheet();
    const soalSheet = ss.getSheetByName('Soal');
    const pesertaSheet = ss.getSheetByName('Peserta');
    
    if (!soalSheet) {
      return { success: false, error: "Sheet 'Soal' tidak ditemukan!" };
    }
    
    if (!pesertaSheet) {
      return { success: false, error: "Sheet 'Peserta' tidak ditemukan!" };
    }
    
    const questions = getQuestions();
    
    if (questions.error) {
      return { success: false, error: questions.error };
    }
    
    let score = 0;
    let details = [];
    
    questions.forEach((q, index) => {
      const isCorrect = checkAnswer(q, answers[index]);
      if (isCorrect) {
        score++;
      }
      details.push({
        questionId: q.id,
        answer: answers[index],
        isCorrect: isCorrect
      });
    });
    
    // Simpan hasil ke sheet Peserta
    const timestamp = new Date();
    pesertaSheet.appendRow([
      timestamp,
      studentEmail,
      studentName,
      score,
      questions.length,
      (score / questions.length * 100).toFixed(2),
      JSON.stringify(details)
    ]);
    
    return {
      success: true,
      score: score,
      total: questions.length,
      percentage: (score / questions.length * 100).toFixed(2),
      details: details
    };
    
  } catch (e) {
    Logger.log("Error submitAnswers: " + e);
    return { success: false, error: "Error: " + e.toString() };
  }
}

// ===== FUNGSI CEK JAWABAN BERDASARKAN TIPE SOAL =====
function checkAnswer(question, userAnswer) {
  if (userAnswer === null || userAnswer === undefined || userAnswer === '') return false;
  
  switch(question.type) {
    case 'multiple_choice':
      return userAnswer.toString().toUpperCase() === question.correctAnswer.toString().toUpperCase();
      
    case 'true_false':
      return userAnswer.toString().toUpperCase() === question.correctAnswer.toString().toUpperCase();
      
    case 'multiple_choice_complex':
      const userAnswers = Array.isArray(userAnswer) ? userAnswer : [userAnswer];
      const correctAnswers = question.correctAnswer || [];
      
      if (userAnswers.length !== correctAnswers.length) return false;
      if (userAnswers.length === 0) return false;
      
      const userSet = new Set(userAnswers.map(x => x.toString().toUpperCase()).sort());
      const correctSet = new Set(correctAnswers.map(x => x.toString().toUpperCase()).sort());
      
      return userSet.size === correctSet.size && 
             [...userSet].every(item => correctSet.has(item));
      
    case 'true_false_table':
      const tableAnswers = Array.isArray(userAnswer) ? userAnswer : [];
      const correctTableAnswers = question.correctAnswer || [];
      
      if (tableAnswers.length !== correctTableAnswers.length) return false;
      if (tableAnswers.length === 0) return false;
      
      return tableAnswers.every((ans, idx) => 
        ans.toString().toUpperCase() === correctTableAnswers[idx].toString().toUpperCase()
      );
      
    case 'matching':
      const matchingAns = Array.isArray(userAnswer) ? userAnswer : [];
      const correctMatching = question.correctAnswer || [];
      
      if (matchingAns.length !== correctMatching.length) return false;
      if (matchingAns.length === 0) return false;
      
      return matchingAns.every((ans, idx) => 
        ans.toString().trim() === correctMatching[idx].toString().trim()
      );
      
    case 'essay':
      return false; // Essay dikoreksi manual
      
    default:
      return false;
  }
}