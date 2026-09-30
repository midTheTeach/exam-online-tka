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

// Ambil semua soal dari Google Sheets
function getQuestions() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Soal');
    
    if (!sheet) {
      return { error: "Sheet 'Soal' tidak ditemukan" };
    }
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 20).getValues();
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
    Logger.log(e);
    return { error: e.toString() };
  }
}

// Parse data soal sesuai dengan tipe
function parseQuestion(row) {
  const [id, type, question, optA, optB, optC, optD, optE, correctAnswer, 
          details, minCorrect, maxCorrect, tableRows, tableAnswers, 
          matchLeft, matchRight, matchAnswers] = row;
  
  if (!id || !type || !question) return null;
  
  const baseObj = {
    id: id,
    type: type,
    question: question
  };
  
  switch(type) {
    case 'multiple_choice':
      return {
        ...baseObj,
        options: [
          { key: 'A', text: optA },
          { key: 'B', text: optB },
          { key: 'C', text: optC },
          { key: 'D', text: optD }
        ].filter(o => o.text),
        correctAnswer: correctAnswer ? correctAnswer.split(',').map(x => x.trim()) : []
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
        ].filter(o => o.text),
        correctAnswer: correctAnswer ? correctAnswer.split(',').map(x => x.trim()) : [],
        minCorrect: minCorrect || 2,
        maxCorrect: maxCorrect || 3
      };
      
    case 'true_false':
      return {
        ...baseObj,
        correctAnswer: correctAnswer ? correctAnswer.toUpperCase() : 'B' // B untuk Benar, S untuk Salah
      };
      
    case 'true_false_table':
      const rows = tableRows ? parseTableRows(tableRows) : [];
      const answers = tableAnswers ? tableAnswers.split(',').map(x => x.trim().toUpperCase()) : [];
      return {
        ...baseObj,
        rows: rows,
        correctAnswer: answers
      };
      
    case 'matching':
      const leftItems = matchLeft ? matchLeft.split('|').map(x => x.trim()) : [];
      const rightItems = matchRight ? matchRight.split('|').map(x => x.trim()) : [];
      const matchingAnswers = matchAnswers ? matchAnswers.split(',').map(x => x.trim()) : [];
      return {
        ...baseObj,
        leftItems: leftItems,
        rightItems: rightItems,
        correctAnswer: matchingAnswers
      };
      
    case 'essay':
      return {
        ...baseObj,
        placeholder: details || 'Tulis jawaban Anda di sini...'
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

// Submit jawaban dan hitung skor
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
    Logger.log(e);
    return { success: false, error: e.toString() };
  }
}

// Cek jawaban berdasarkan tipe soal
function checkAnswer(question, userAnswer) {
  if (!userAnswer) return false;
  
  switch(question.type) {
    case 'multiple_choice':
    case 'true_false':
      return userAnswer === question.correctAnswer[0];
      
    case 'multiple_choice_complex':
      const userAnswers = Array.isArray(userAnswer) ? userAnswer : [userAnswer];
      const correctAnswers = question.correctAnswer || [];
      
      if (userAnswers.length !== correctAnswers.length) return false;
      
      const userSet = new Set(userAnswers.sort());
      const correctSet = new Set(correctAnswers.sort());
      
      return userSet.size === correctSet.size && 
             [...userSet].every(item => correctSet.has(item));
      
    case 'true_false_table':
      const tableAnswers = Array.isArray(userAnswer) ? userAnswer : [];
      const correctTableAnswers = question.correctAnswer || [];
      
      if (tableAnswers.length !== correctTableAnswers.length) return false;
      
      return tableAnswers.every((ans, idx) => 
        ans.toUpperCase() === correctTableAnswers[idx].toUpperCase()
      );
      
    case 'matching':
      const matchingAns = Array.isArray(userAnswer) ? userAnswer : [];
      const correctMatching = question.correctAnswer || [];
      
      if (matchingAns.length !== correctMatching.length) return false;
      
      return matchingAns.every((ans, idx) => 
        ans === correctMatching[idx]
      );
      
    case 'essay':
      // Essay biasanya dikoreksi manual
      return false;
      
    default:
      return false;
  }
}

// Function untuk setup Google Sheets template
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Buat sheet Soal jika belum ada
  if (!ss.getSheetByName('Soal')) {
    const soalSheet = ss.insertSheet('Soal');
    const headers = [
      'ID Soal', 'Tipe', 'Pertanyaan', 'Opsi A', 'Opsi B', 'Opsi C', 'Opsi D', 'Opsi E',
      'Jawaban Benar', 'Detail Soal', 'Min Jawaban Benar', 'Max Jawaban Benar',
      'Baris Tabel (|)', 'Jawaban Tabel', 'Item Kiri Pasangan', 'Item Kanan Pasangan', 'Jawaban Pasangan'
    ];
    soalSheet.appendRow(headers);
    
    // Contoh soal
    soalSheet.appendRow([
      1, 'multiple_choice', 'Ibu kota Indonesia adalah?', 'Jakarta', 'Surabaya', 'Bandung', 'Medan', '',
      'A', '', '', '', '', '', '', '', ''
    ]);
  }
  
  // Buat sheet Peserta jika belum ada
  if (!ss.getSheetByName('Peserta')) {
    const pesertaSheet = ss.insertSheet('Peserta');
    pesertaSheet.appendRow([
      'Waktu', 'Email', 'Nama Peserta', 'Skor', 'Total Soal', 'Persentase', 'Detail Jawaban'
    ]);
  }
}

// Jalankan setup saat pertama kali
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Ujian Online')
    .addItem('Setup Sheet', 'setupSheets')
    .addToUi();
}