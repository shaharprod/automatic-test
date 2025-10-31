/**
 * Apps Script - יצירת Google Form עם חישוב ציונים אוטומטי
 * עם ממשק גרפי פשוט למורים
 */

/**
 * פונקציה לפתיחת ממשק גרפי - הרץ את זה במקום createFormFromText
 */
function showQuestionnaireForm() {
  const html = HtmlService.createHtmlOutputFromFile('QuestionnaireGUI')
    .setWidth(600)
    .setHeight(800)
    .setTitle('יצירת שאלון עם ציונים אוטומטיים');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * יצירת פורם מהנתונים שהמורה הזין בממשק
 */
function createFormFromGUI(questionnaireText, correctAnswersJson) {
  try {
    // פרסור השאלון
    const questions = parseQuestionnaire(questionnaireText);
    if (questions.length === 0) {
      return { success: false, message: 'לא נמצאו שאלות בטקסט' };
    }

    // יצירת Google Sheet
    const sheet = createSpreadsheet(questions);
    
    // יצירת Google Form
    const form = createForm(questions);
    
    // חיבור הפורם ל-Sheet
    const ss = SpreadsheetApp.openByUrl(sheet.url);
    const formObj = FormApp.openById(form.id);
    
    try {
      formObj.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
    } catch (error) {
      // התעלם משגיאות
    }

    // המתן קצת
    Utilities.sleep(3000);

    // שמירת קישורים ב-Sheet
    const sheetTab = ss.getSheets()[0];
    sheetTab.getRange(1, 9).setValue('תשובה נכונה');
    sheetTab.getRange(1, 10).setValue('Form URL');
    sheetTab.getRange(1, 11).setValue('Form ID');
    sheetTab.getRange(2, 10).setValue(form.url);
    sheetTab.getRange(2, 11).setValue(form.id);

    // הוספת תשובות נכונות
    const correctAnswers = JSON.parse(correctAnswersJson || '{}');
    for (let i = 0; i < questions.length; i++) {
      const qNum = String(i + 1).padStart(2, '0');
      const correctAns = correctAnswers[qNum] || '';
      sheetTab.getRange(i + 2, 9).setValue(correctAns);
    }

    // נסה ליצור טריגר
    try {
      const triggers = ScriptApp.getProjectTriggers();
      triggers.forEach(function(trigger) {
        if (trigger.getHandlerFunction() === 'onFormSubmit') {
          try {
            ScriptApp.deleteTrigger(trigger);
          } catch (e) {}
        }
      });

      try {
        ScriptApp.newTrigger('onFormSubmit')
          .forForm(formObj)
          .onFormSubmit()
          .create();
      } catch (e1) {
        try {
          ScriptApp.newTrigger('onFormSubmit')
            .onFormSubmit()
            .create();
        } catch (e2) {
          // לא הצליח - צריך להוסיף ידנית
        }
      }
    } catch (error) {
      // התעלם
    }

    return {
      success: true,
      sheetUrl: sheet.url,
      formUrl: form.url,
      message: '✅ הפורם נוצר בהצלחה!'
    };

  } catch (error) {
    return { success: false, message: 'שגיאה: ' + error.toString() };
  }
}

/**
 * פרסור השאלון - תומך בפורמטים: "שאלה 1." או "Q1." או "ש1."
 */
function parseQuestionnaire(text) {
  const questions = [];
  const lines = text.split('\n').map(l => l.trim()).filter(l => l);
  
  let currentQuestion = null;
  
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    
    // זיהוי שאלה: "שאלה 1.", "Q1.", "ש1."
    const questionMatch = line.match(/^(?:שאלה|Q|ש)\s*(\d+)\.\s*(.+)/i);
    if (questionMatch) {
      if (currentQuestion) questions.push(currentQuestion);
      currentQuestion = {
        number: questionMatch[1],
        text: questionMatch[2],
        answers: []
      };
      continue;
    }
    
    // זיהוי תשובה: "א.", "ב.", "A1.", "א1."
    const answerMatch = line.match(/^([אבגדה]|A\d+|א\d+)\.\s*(.+)/);
    if (answerMatch && currentQuestion) {
      currentQuestion.answers.push(answerMatch[2]);
    }
  }
  
  if (currentQuestion) questions.push(currentQuestion);
  return questions;
}

/**
 * יצירת Google Sheet עם השאלות
 */
function createSpreadsheet(questions) {
  const ss = SpreadsheetApp.create('שאלון לטפסים - ' + new Date().toLocaleDateString('he-IL'));
  const sheet = ss.getSheets()[0];
  sheet.setName('שאלון');
  
  // כותרות
  sheet.getRange(1, 1, 1, 8).setValues([['מספר', 'שאלה', 'סוג', 'תשובה א', 'תשובה ב', 'תשובה ג', 'תשובה ד', 'תשובה ה']]);
  sheet.getRange(1, 1, 1, 8).setFontWeight('bold');
  
  // שאלות
  questions.forEach(function(q, index) {
    const row = index + 2;
    const answers = q.answers.concat(['', '', '', '']).slice(0, 5);
    sheet.getRange(row, 1, 1, 8).setValues([[
      q.number,
      q.text,
      'multiple_choice',
      answers[0] || '',
      answers[1] || '',
      answers[2] || '',
      answers[3] || '',
      answers[4] || ''
    ]]);
  });
  
  return { url: ss.getUrl(), id: ss.getId() };
}

/**
 * יצירת Google Form עם השאלות
 */
function createForm(questions) {
  const form = FormApp.create('שאלון - ' + new Date().toLocaleDateString('he-IL'));
  form.setDescription('שאלון שנוצר אוטומטית');
  
  form.setAcceptingResponses(true);
  form.setAllowResponseEdits(true);
  
  // הוסף שדות חובה: שם מלא ואימייל
  const nameItem = form.addTextItem();
  nameItem.setTitle('שם מלא');
  nameItem.setRequired(true);
  nameItem.setHelpText('הזן את שמך המלא');
  
  const emailItem = form.addTextItem();
  emailItem.setTitle('כתובת אימייל');
  emailItem.setRequired(true);
  emailItem.setHelpText('הזן את כתובת האימייל שלך');
  emailItem.setValidation(FormApp.createTextValidation()
    .setHelpText('הזן כתובת אימייל תקינה (למשל: name@example.com)')
    .requireTextMatchesPattern('[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\\.[a-zA-Z]{2,}')
    .build());
  
  const separatorItem = form.addSectionHeaderItem();
  separatorItem.setTitle('שאלות השאלון');
  
  const hebrewLetters = ['א', 'ב', 'ג', 'ד', 'ה'];
  
  questions.forEach(function(q) {
    const item = form.addMultipleChoiceItem();
    item.setTitle('שאלה ' + q.number + ': ' + q.text);
    item.setRequired(false);
    
    const choices = [];
    q.answers.forEach(function(answer, index) {
      const letter = hebrewLetters[index] || String(index + 1);
      choices.push(item.createChoice(letter + '. ' + answer));
    });
    item.setChoices(choices);
  });
  
  try {
    const formFile = DriveApp.getFileById(form.getId());
    formFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (error) {
    // התעלם
  }
  
  return { url: form.getPublishedUrl(), id: form.getId() };
}

/**
 * חישוב ציונים אוטומטי - מופעל בעת הגשת פורם
 */
function onFormSubmit(e) {
  try {
    if (!e || !e.response || !e.source) {
      Logger.log('⚠️ הפרמטר לא תקין');
      return;
    }

    const formResponse = e.response;
    const form = e.source;

    const destinationId = form.getDestinationId();
    if (!destinationId) {
      Logger.log('⚠️ הפורם לא מחובר ל-Sheet');
      return;
    }

    const ss = SpreadsheetApp.openById(destinationId);
    
    const sheets = ss.getSheets();
    let responsesSheet = null;
    let settingsSheet = null;
    
    for (let i = 0; i < sheets.length; i++) {
      const name = sheets[i].getName();
      if (name === 'טופס 1' || name.includes('תשובות') || name === 'Form Responses 1') {
        responsesSheet = sheets[i];
      } else if (name === 'שאלון') {
        settingsSheet = sheets[i];
      }
    }

    if (!responsesSheet || !settingsSheet) {
      Logger.log('⚠️ לא נמצאו הטאבים הנדרשים');
      return;
    }

    const lastRow = settingsSheet.getLastRow();
    const questionData = settingsSheet.getRange(2, 1, lastRow - 1, 9).getValues();
    const itemResponses = formResponse.getItemResponses();
    const lastResponseRow = responsesSheet.getLastRow();

    const headers = responsesSheet.getRange(1, 1, 1, responsesSheet.getLastColumn()).getValues()[0];
    let scoreCol = headers.indexOf('ציון') + 1;
    let percentCol = headers.indexOf('אחוז') + 1;

    if (scoreCol === 0) {
      scoreCol = responsesSheet.getLastColumn() + 1;
      responsesSheet.getRange(1, scoreCol).setValue('ציון').setFontWeight('bold');
    }
    if (percentCol === 0) {
      percentCol = responsesSheet.getLastColumn() + 1;
      if (percentCol === scoreCol) percentCol++;
      responsesSheet.getRange(1, percentCol).setValue('אחוז').setFontWeight('bold');
    }

    let totalQuestions = 0;
    let correctCount = 0;
    const hebrewLetters = ['א', 'ב', 'ג', 'ד', 'ה'];

    for (let i = 0; i < questionData.length; i++) {
      const questionNum = questionData[i][0];
      const correctAnswer = questionData[i][8];
      
      if (!correctAnswer) continue;
      
      totalQuestions++;
      const questionHeader = 'שאלה ' + questionNum + ':';
      
      let userAnswer = null;
      for (let j = 0; j < itemResponses.length; j++) {
        const title = itemResponses[j].getItem().getTitle();
        if (title && title.includes(questionHeader)) {
          userAnswer = itemResponses[j].getResponse();
          break;
        }
      }

      if (!userAnswer) continue;

      const answerOptions = [
        questionData[i][3], questionData[i][4], 
        questionData[i][5], questionData[i][6], questionData[i][7]
      ].filter(opt => opt && opt.trim());

      let correctIndex = -1;
      if (correctAnswer === 'א' || correctAnswer === 'א1' || correctAnswer === 'A1') correctIndex = 0;
      else if (correctAnswer === 'ב' || correctAnswer === 'א2' || correctAnswer === 'A2') correctIndex = 1;
      else if (correctAnswer === 'ג' || correctAnswer === 'א3' || correctAnswer === 'A3') correctIndex = 2;
      else if (correctAnswer === 'ד' || correctAnswer === 'א4' || correctAnswer === 'A4') correctIndex = 3;
      else {
        let num = correctAnswer.replace(/[אבגדAא]/, '');
        correctIndex = parseInt(num) - 1;
      }

      if (correctIndex >= 0 && correctIndex < answerOptions.length) {
        const correctText = answerOptions[correctIndex];
        const correctWithLetter = hebrewLetters[correctIndex] + '. ' + correctText;
        
        const userClean = userAnswer.toString().trim();
        if (userClean === correctWithLetter.trim() || 
            userClean === correctText.trim() ||
            userClean.startsWith(hebrewLetters[correctIndex] + '.')) {
          correctCount++;
        }
      }
    }

    const score = correctCount + '/' + totalQuestions;
    const percentage = totalQuestions > 0 ? Math.round((correctCount / totalQuestions) * 100) : 0;
    
    responsesSheet.getRange(lastResponseRow, scoreCol).setValue(score);
    responsesSheet.getRange(lastResponseRow, percentCol).setValue(percentage + '%');

    if (percentage >= 80) {
      responsesSheet.getRange(lastResponseRow, scoreCol, 1, 2).setBackground('#c6efce');
    } else if (percentage >= 60) {
      responsesSheet.getRange(lastResponseRow, scoreCol, 1, 2).setBackground('#ffeb9c');
    } else {
      responsesSheet.getRange(lastResponseRow, scoreCol, 1, 2).setBackground('#ffc7ce');
    }

    Logger.log('✅ ציון מחושב: ' + score + ' (' + percentage + '%)');

  } catch (error) {
    Logger.log('❌ שגיאה: ' + error.toString());
  }
}

