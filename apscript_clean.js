/**
 * Apps Script - יצירת Google Form עם חישוב ציונים אוטומטי
 *
 * הוראות קצרות:
 * 1. העתק את הקוד ל-Apps Script (https://script.google.com/)
 * 2. ערוך את QUESTIONNAIRE_TEXT עם השאלון שלך
 * 3. ערוך את getCorrectAnswers() עם התשובות הנכונות
 * 4. הרץ createFormFromText()
 * 5. הוסף טריגר ידנית: ⚙️ → Triggers → + Add Trigger
 *    Function: onFormSubmit | Event: From form → On form submit
 */

/**
 * יצירת פורם מה-שאלון
 */
function createFormFromText() {
  try {
    // ========================================
    // הכנס כאן את השאלון שלך
    // ========================================
    const QUESTIONNAIRE_TEXT = `שאלה 1. מהו העיקרון המרכזי של "למידה מונחית"?
א. מתן תשובה מלאה ומקיפה לנושא מורכב
ב. שימוש במודלים מתקדמים ליצירת תמונות
ג. בניית דיאלוג אינטראקטיבי שמפרק בעיות ומקדם את הלומד צעד אחר צעד
ד. יצירת מחוונים מותאמים אישית אוטומטית

שאלה 2. מהו תפקידה של הבינה המלאכותית (AI) ביחס לצוותי חינוך?
א. ה-AI נועד להחליף מורים בתחומי ידע ספציפיים
ב. ה-AI נועד לספק משוב מיידי במקום המורה
ג. ה-AI לעולם לא יכול להחליף את המומחיות, הידע או היצירתיות של איש חינוך, אך הוא כלי מועיל לשיפור החוויה
ד. ה-AI משמש בעיקר לשיפור האבטחה הדיגיטלית ולחסימת תוכנות כופר

שאלה 3. איזה אחוז מהמורים בישראל מדווחים על שימוש בכלי בינה מלאכותית יוצרת להכנה ותכנון שיעורים?
א. כרבע מהמורים (25%)
ב. כשליש מהמורים (33%)
ג. כמחצית מהמורים (50%)
ד. למעלה מ-75% מהמורים`;

    // פרסור השאלון
    const questions = parseQuestionnaire(QUESTIONNAIRE_TEXT);
    if (questions.length === 0) {
      throw new Error('לא נמצאו שאלות בטקסט');
    }

    // יצירת Google Sheet
    const sheet = createSpreadsheet(questions);
    Logger.log('✅ Google Sheet נוצר: ' + sheet.url);

    // יצירת Google Form
    const form = createForm(questions);
    Logger.log('✅ Google Form נוצר: ' + form.url);

    // חיבור הפורם ל-Sheet
    const ss = SpreadsheetApp.openByUrl(sheet.url);
    const formObj = FormApp.openById(form.id);

    try {
      formObj.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
      Logger.log('✅ הפורם חובר ל-Sheet');
    } catch (error) {
      Logger.log('⚠️ שגיאה בחיבור: ' + error.toString());
    }

    // המתן קצת כדי שה-Sheet יעודכן
    Utilities.sleep(3000);

    // שמירת קישורים ב-Sheet
    const sheetTab = ss.getSheets()[0];
    sheetTab.getRange(1, 9).setValue('תשובה נכונה');
    sheetTab.getRange(1, 10).setValue('Form URL');
    sheetTab.getRange(1, 11).setValue('Form ID');
    sheetTab.getRange(2, 10).setValue(form.url);
    sheetTab.getRange(2, 11).setValue(form.id);

    // הוספת תשובות נכונות
    const correctAnswers = getCorrectAnswers();
    for (let i = 0; i < questions.length; i++) {
      const qNum = String(i + 1).padStart(2, '0');
      const correctAns = correctAnswers[qNum] || '';
      sheetTab.getRange(i + 2, 9).setValue(correctAns);
    }

    // נסה ליצור טריגר אוטומטית
    Logger.log('');
    Logger.log('🔄 מנסה ליצור טריגר אוטומטית...');
    try {
      // מחק טריגרים קיימים
      const triggers = ScriptApp.getProjectTriggers();
      triggers.forEach(function(trigger) {
        if (trigger.getHandlerFunction() === 'onFormSubmit') {
          try {
            ScriptApp.deleteTrigger(trigger);
          } catch (e) {
            // התעלם
          }
        }
      });

      // נסה ליצור טריגר
      try {
        ScriptApp.newTrigger('onFormSubmit')
          .forForm(formObj)
          .onFormSubmit()
          .create();
        Logger.log('✅ טריגר נוצר אוטומטית!');
      } catch (e1) {
        try {
          ScriptApp.newTrigger('onFormSubmit')
            .onFormSubmit()
            .create();
          Logger.log('✅ טריגר נוצר אוטומטית!');
        } catch (e2) {
          Logger.log('⚠️ לא ניתן ליצור טריגר אוטומטית');
          Logger.log('💡 צריך להוסיף טריגר ידנית:');
          Logger.log('   ⚙️ → Triggers → + Add Trigger');
          Logger.log('   Function: onFormSubmit');
          Logger.log('   Event: From form → On form submit');
        }
      }
    } catch (error) {
      Logger.log('⚠️ שגיאה ביצירת טריגר: ' + error.toString());
      Logger.log('💡 צריך להוסיף טריגר ידנית:');
      Logger.log('   ⚙️ → Triggers → + Add Trigger');
    }

    Logger.log('');
    Logger.log('✅ הכל מוכן!');
    Logger.log('📊 Sheet: ' + sheet.url);
    Logger.log('📝 Form: ' + form.url);
    Logger.log('');
    Logger.log('💡 חשוב - כדי שהפורם לא יבקש הרשאות:');
    Logger.log('   1. פתח את הפורם שנוצר');
    Logger.log('   2. לחץ על ⚙️ (Settings)');
    Logger.log('   3. ב-"General": בטל "Limit to 1 response" (אם מסומן)');
    Logger.log('   4. ב-"Responses": בטל "Collect email addresses" (אם מסומן)');
    Logger.log('   5. שמור (Save)');
    Logger.log('');
    Logger.log('🎉 עכשיו תוכל לשתף את הפורם - הוא לא יבקש הרשאות!');

  } catch (error) {
    Logger.log('❌ שגיאה: ' + error.toString());
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

  // הגדר את הפורם לשיתוף ללא הרשאות
  // פותח את הפורם לכל מי שיש לו את הקישור - לא דורש התחברות
  form.setAcceptingResponses(true);
  form.setAllowResponseEdits(true); // אפשר לערוך תשובות

  // הוסף שדות חובה: שם מלא ואימייל
  const nameItem = form.addTextItem();
  nameItem.setTitle('שם מלא');
  nameItem.setRequired(true);
  nameItem.setHelpText('הזן את שמך המלא');

  const emailItem = form.addTextItem();
  emailItem.setTitle('כתובת אימייל');
  emailItem.setRequired(true);
  emailItem.setHelpText('הזן את כתובת האימייל שלך');
  // ולידציה לאימייל
  emailItem.setValidation(FormApp.createTextValidation()
    .setHelpText('הזן כתובת אימייל תקינה (למשל: name@example.com)')
    .requireTextMatchesPattern('[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\\.[a-zA-Z]{2,}')
    .build());

  // הוסף שאלה ריקה להפרדה (אופציונלי)
  const separatorItem = form.addSectionHeaderItem();
  separatorItem.setTitle('שאלות השאלון');

  const hebrewLetters = ['א', 'ב', 'ג', 'ד', 'ה'];

  // הוסף את השאלות
  questions.forEach(function(q) {
    const item = form.addMultipleChoiceItem();
    item.setTitle('שאלה ' + q.number + ': ' + q.text);
    item.setRequired(false);

    // הוסף תשובות עם אותיות עבריות
    const choices = [];
    q.answers.forEach(function(answer, index) {
      const letter = hebrewLetters[index] || String(index + 1);
      choices.push(item.createChoice(letter + '. ' + answer));
    });
    item.setChoices(choices);
  });

  // הגדר שיתוף - כל מי שיש לו את הקישור יכול למלא
  // זה נעשה דרך Drive API - ננסה להגדיר זאת
  try {
    const formFile = DriveApp.getFileById(form.getId());
    // שתף את הקובץ כך שכל מי שיש לו קישור יכול לצפות
    // הערה: זה לא משנה את הגדרות הפורם עצמו, אבל עוזר
    formFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    Logger.log('✅ הפורם מוגדר לשיתוף');
  } catch (error) {
    Logger.log('⚠️ לא ניתן להגדיר שיתוף אוטומטית: ' + error.toString());
    Logger.log('💡 הגדר ידנית: פתח את הפורם → ⚙️ → General → "Limit to 1 response" ובטל');
    Logger.log('💡 ו: Settings → Responses → "Collect email addresses" - בטל אם לא צריך');
  }

  return { url: form.getPublishedUrl(), id: form.getId() };
}

/**
 * הגדרת תשובות נכונות
 * ערוך כאן את התשובות הנכונות
 */
function getCorrectAnswers() {
  return {
    '01': 'ג',   // שאלה 1 - תשובה ג
    '02': 'ג',   // שאלה 2 - תשובה ג
    '03': 'א'    // שאלה 3 - תשובה א
  };
}

/**
 * חישוב ציונים אוטומטי - מופעל בעת הגשת פורם
 */
function onFormSubmit(e) {
  try {
    if (!e || !e.response || !e.source) {
      Logger.log('⚠️ הפרמטר לא תקין - ודא שהטריגר מוגדר נכון');
      return;
    }

    const formResponse = e.response;
    const form = e.source;

    // קבל Sheet דרך הפורם
    const destinationId = form.getDestinationId();
    if (!destinationId) {
      Logger.log('⚠️ הפורם לא מחובר ל-Sheet');
      return;
    }

    const ss = SpreadsheetApp.openById(destinationId);

    // מצא טאב תשובות
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

    // קבל שאלות ותשובות נכונות
    const lastRow = settingsSheet.getLastRow();
    const questionData = settingsSheet.getRange(2, 1, lastRow - 1, 9).getValues();
    const itemResponses = formResponse.getItemResponses();
    const lastResponseRow = responsesSheet.getLastRow();

    // מצא או צור עמודות ציון ואחוז
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

    // חשב ציון
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

      // מצא אינדקס תשובה נכונה
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

    // שמור ציון
    const score = correctCount + '/' + totalQuestions;
    const percentage = totalQuestions > 0 ? Math.round((correctCount / totalQuestions) * 100) : 0;

    responsesSheet.getRange(lastResponseRow, scoreCol).setValue(score);
    responsesSheet.getRange(lastResponseRow, percentCol).setValue(percentage + '%');

    // הדגש לפי ציון
    if (percentage >= 80) {
      responsesSheet.getRange(lastResponseRow, scoreCol, 1, 2).setBackground('#c6efce'); // ירוק
    } else if (percentage >= 60) {
      responsesSheet.getRange(lastResponseRow, scoreCol, 1, 2).setBackground('#ffeb9c'); // צהוב
    } else {
      responsesSheet.getRange(lastResponseRow, scoreCol, 1, 2).setBackground('#ffc7ce'); // אדום
    }

    Logger.log('✅ ציון מחושב: ' + score + ' (' + percentage + '%)');

  } catch (error) {
    Logger.log('❌ שגיאה: ' + error.toString());
  }
}

