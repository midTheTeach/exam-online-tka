// Main Apps Script File - Backend untuk Ujian Online TKA
// Deploy sebagai Web App

function doGet(e) {
  return HtmlService.createHtmlOutput(getHtmlTemplate())
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getHtmlTemplate() {
  return HtmlService.createTemplateFromFile('Index').evaluate();
}

// ===== FUNGSI UNTUK MENGAMBIL SOAL =====
function getQuestions() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Soal');
    
    if (!sheet) {
      return { error: "Sheet 'Soal' tidak ditemukan" };
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return [];
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
    return [];
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
    question: question
  };
  
  switch(type.toLowerCase().trim()) {
    case 'multiple_choice':
      return {
        ...baseObj,
        options: [
          { key: 'A', text: optA },
          { key: 'B', text: optB },
          { key: 'C', text: optC },
          { key: 'D', text: optD }
        ].filter(o => o.text && o.text.toString().trim()),
        correctAnswer: correctAnswer ? correctAnswer.toString().toUpperCase().trim() : ''
      };
      
    case 'multiple_choice_complex':
      return {
        ...baseObj,
        options: [
          { key: 'A', text: optA },
          { key: 'B', text: optB },
          { key: 'C', text: optC },
          { key: 'D', text: optD },
          { key: 'E', text: optE }
        ].filter(o => o.text && o.text.toString().trim()),
        correctAnswer: correctAnswer ? correctAnswer.toString().split(',').map(x => x.trim().toUpperCase()) : [],
        minCorrect: parseInt(minCorrect) || 2,
        maxCorrect: parseInt(maxCorrect) || 3,
        instruction: 'Pilih ' + (minCorrect || 2) + '-' + (maxCorrect || 3) + ' jawaban yang benar!'
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
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const soalSheet = ss.getSheetByName('Soal');
    const pesertaSheet = ss.getSheetByName('Peserta');
    
    if (!soalSheet || !pesertaSheet) {
      return { success: false, error: "Sheet tidak ditemukan" };
    }
    
    const questions = getQuestions();
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
    return { success: false, error: e.toString() };
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
      // Essay biasanya dikoreksi manual, return false untuk auto-scoring
      return false;
      
    default:
      return false;
  }
}

// ===== FUNGSI SETUP SHEETS (OPTIONAL) =====
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Buat sheet Soal jika belum ada
  if (!ss.getSheetByName('Soal')) {
    const soalSheet = ss.insertSheet('Soal', 0);
    const headers = [
      'ID Soal', 'Tipe', 'Pertanyaan', 'Opsi A', 'Opsi B', 'Opsi C', 'Opsi D', 'Opsi E',
      'Jawaban Benar', 'Detail Soal', 'Min Jawaban Benar', 'Max Jawaban Benar',
      'Baris Tabel (|)', 'Jawaban Tabel', 'Item Kiri Pasangan', 'Item Kanan Pasangan', 'Jawaban Pasangan'
    ];
    soalSheet.appendRow(headers);
  }
  
  // Buat sheet Peserta jika belum ada
  if (!ss.getSheetByName('Peserta')) {
    const pesertaSheet = ss.insertSheet('Peserta', 1);
    pesertaSheet.appendRow([
      'Waktu', 'Email', 'Nama Peserta', 'Skor', 'Total Soal', 'Persentase', 'Detail Jawaban'
    ]);
  }
}