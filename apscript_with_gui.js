/**
 * Apps Script - יצירת Google Form עם חישוב ציונים אוטומטי
 * עם ממשק גרפי פשוט למורים
 */

/**
 * פונקציה ליצירת תפריט מותאם אישית - תתוסף אוטומטית בעת פתיחת Sheet
 */
function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('📝 יצירת שאלון')
      .addItem('פתח ממשק גרפי', 'showQuestionnaireForm')
      .addToUi();
  } catch (error) {
    Logger.log('⚠️ לא ניתן ליצור תפריט: ' + error.toString());
  }
}

/**
 * פונקציה לפתיחת ממשק גרפי - הרץ את זה במקום createFormFromText
 * אם לא עובד, הרץ: createFormFromGUI(questionnaire, correctAnswersJson)
 */
function showQuestionnaireForm() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('QuestionnaireGUI')
      .setWidth(600)
      .setHeight(800)
      .setTitle('יצירת שאלון עם ציונים אוטומטיים');

    // נסה לפתוח ב-sidebar
    try {
      SpreadsheetApp.getUi().showSidebar(html);
    } catch (e) {
      // אם sidebar לא עובד, נסה modal dialog
      try {
        SpreadsheetApp.getUi().showModalDialog(html, 'יצירת שאלון עם ציונים אוטומטיים');
      } catch (e2) {
        // אם גם זה לא עובד, נסה Web App
        Logger.log('⚠️ לא ניתן לפתוח ממשק גרפי מההקשר הזה');
        Logger.log('💡 פתרון: פתח Sheet חדש והפעל את showQuestionnaireForm() מתוכו');
        throw new Error('לא ניתן לפתוח ממשק גרפי מההקשר הזה. פתח Sheet חדש והריץ את הפונקציה מתוכו.');
      }
    }
  } catch (error) {
    Logger.log('❌ שגיאה בפתיחת ממשק גרפי: ' + error.toString());
    throw error;
  }
}

/**
 * פונקציה להרצה ישירה (ללא ממשק גרפי) - למורים שמעדיפים טקסט
 * דוגמה:
 * createFormDirectly(
 *   "שאלה 1. שאלה?\nא. תשובה א\nב. תשובה ב",
 *   "01:א, 02:ב"
 * );
 */
function createFormDirectly(questionnaireText, correctAnswersText) {
  let correctAnswers = {};

  if (correctAnswersText) {
    const pairs = correctAnswersText.split(',').map(s => s.trim());
    pairs.forEach(function(pair) {
      const match = pair.match(/(\d+)[:=]\s*([אבגד])/);
      if (match) {
        const qNum = match[1].padStart(2, '0');
        const answer = match[2];
        correctAnswers[qNum] = answer;
      }
    });
  }

  // אם לא ניתן שם, השתמש בברירת מחדל
  const defaultName = 'שאלון - ' + new Date().toLocaleDateString('he-IL');
  return createFormFromGUI(questionnaireText, JSON.stringify(correctAnswers), defaultName);
}

/**
 * יצירת פורם מהנתונים שהמורה הזין בממשק
 */
function createFormFromGUI(questionnaireText, correctAnswersJson, questionnaireName) {
  try {
    // פרסור השאלון
    const questions = parseQuestionnaire(questionnaireText);
    if (questions.length === 0) {
      return { success: false, message: 'לא נמצאו שאלות בטקסט' };
    }

    // אם לא ניתן שם, השתמש בברירת מחדל
    if (!questionnaireName || questionnaireName.trim() === '') {
      questionnaireName = 'שאלון - ' + new Date().toLocaleDateString('he-IL');
    }

    // לוג לבדיקה
    Logger.log('✅ נמצאו ' + questions.length + ' שאלות');
    Logger.log('📝 שם השאלון: ' + questionnaireName);
    questions.forEach(function(q, i) {
      Logger.log('שאלה ' + (i+1) + ': ' + q.number + ' - ' + q.answers.length + ' תשובות');
    });

    // יצירת Google Sheet
    const sheet = createSpreadsheet(questions);

    // יצירת Google Form
    const form = createForm(questions, questionnaireName);

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
 * פרסור השאלון - תומך בפורמטים:
 * - "1. שאלה\nתשובות: א. תשובה ב. תשובה"
 * - "שאלה 1. שאלה\nא. תשובה ב. תשובה"
 * - "1. שאלה? א. תשובה ב. תשובה" (בשורה אחת)
 */
function parseQuestionnaire(text) {
  const questions = [];
  const lines = text.split('\n').map(l => l.trim()).filter(l => l);

  let currentQuestion = null;

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];

    // זיהוי שאלה: "1.", "שאלה 1.", "שאלה 1:", "Q1.", "ש1." (מספר עם נקודה או נקודתיים)
    const questionMatch = line.match(/^(?:שאלה\s*)?(\d+)[\.:]\s*(.+)/i);
    if (questionMatch) {
      // אם יש שאלה קודמת, נשמור אותה
      if (currentQuestion) {
        questions.push(currentQuestion);
      }

      currentQuestion = {
        number: questionMatch[1],
        text: questionMatch[2],
        answers: []
      };

      // נסה לזהות תשובות באותה שורה כמו השאלה (אם יש)
      extractAnswersFromLine(currentQuestion, line);

      Logger.log('✅ זוהתה שאלה ' + currentQuestion.number + ' עם טקסט: ' + currentQuestion.text.substring(0, 50));

      continue;
    }

    // זיהוי שורה של תשובות: "תשובות: א. תשובה ב. תשובה ג. תשובה ד. תשובה"
    const answersLineMatch = line.match(/^תשובות:\s*(.+)/i);
    if (answersLineMatch && currentQuestion) {
      const answersText = answersLineMatch[1];
      Logger.log('✅ נמצאה שורת תשובות: ' + answersText);

      // נסה לפרסר את התשובות ישירות
      // פורמט: "א. תשובה ב. תשובה ג. תשובה ד. תשובה"
      // גישה פשוטה: חיפוש כל "א.", "ב.", "ג.", "ד." ואז הטקסט עד האות הבאה

      const directAnswers = [];
      const hebrewLetters = ['א', 'ב', 'ג', 'ד', 'ה'];

      // עבור כל אות עברית, מצא את התשובה שלה
      // גישה חדשה: פירוק לפי "א. ", "ב. ", "ג. ", "ד. " - הטקסט ביניהם הוא התשובה

      // גישה פשוטה: מצא כל "א. ", "ב. ", "ג. ", "ד. " וקח את הטקסט ביניהם
      const answerPositions = [];

      // חיפוש כל האותיות עם נקודה ורווח
      for (let i = 0; i < hebrewLetters.length; i++) {
        const letter = hebrewLetters[i];
        const searchStr = letter + '. ';
        let index = answersText.indexOf(searchStr);

        while (index !== -1) {
          answerPositions.push({
            letter: letter,
            index: index,
            start: index + searchStr.length
          });
          // חפש את המופע הבא מאותה אות
          index = answersText.indexOf(searchStr, index + 1);
        }
      }

      // מיין לפי מיקום (משמאל לימין)
      answerPositions.sort(function(a, b) { return a.index - b.index; });

      Logger.log('📌 נמצאו ' + answerPositions.length + ' מיקומי תשובות');

      // עבור כל תשובה, קח את הטקסט עד התשובה הבאה
      for (let i = 0; i < answerPositions.length; i++) {
        const current = answerPositions[i];
        const next = answerPositions[i + 1];

        let answerText;
        if (next) {
          // יש תשובה הבאה - קח את הטקסט ביניהן
          answerText = answersText.substring(current.start, next.index).trim();
        } else {
          // זו התשובה האחרונה - קח את כל מה שנשאר
          answerText = answersText.substring(current.start).trim();
        }

        if (answerText && answerText.length > 0) {
          const cleanAnswer = answerText.replace(/^\s+|\s+$/g, '').replace(/\s+/g, ' ');
          if (cleanAnswer && cleanAnswer.length > 0) {
            // ודא שהתשובה לא מתחילה במספר (שאלה)
            if (!/^\d+[\.:]/.test(cleanAnswer)) {
              directAnswers.push(cleanAnswer);
              Logger.log('  ✅ מצא תשובה ' + current.letter + ': ' + cleanAnswer.substring(0, 60));
            }
          }
        }
      }

      // אם לא מצאנו תשובות, ננסה גם עם רווח בלי נקודה "א "
      if (directAnswers.length === 0) {
        Logger.log('⚠️ לא נמצאו תשובות עם נקודה, מנסה בלי נקודה...');
        for (let i = 0; i < hebrewLetters.length; i++) {
          const letter = hebrewLetters[i];
          const searchStr = letter + ' ';
          let index = answersText.indexOf(searchStr);

          if (index !== -1) {
            const nextLetter = hebrewLetters[i + 1];
            const nextSearchStr = nextLetter ? nextLetter + ' ' : null;
            let nextIndex = nextSearchStr ? answersText.indexOf(nextSearchStr, index + 1) : -1;

            let answerText;
            if (nextIndex !== -1) {
              answerText = answersText.substring(index + searchStr.length, nextIndex).trim();
            } else {
              answerText = answersText.substring(index + searchStr.length).trim();
            }

            if (answerText && answerText.length > 0 && !/^\d+/.test(answerText)) {
              directAnswers.push(answerText);
              Logger.log('  ✅ מצא תשובה (בלי נקודה) ' + letter + ': ' + answerText.substring(0, 50));
            }
          }
        }
      }

      Logger.log('📊 סה"כ ' + directAnswers.length + ' תשובות בשורת "תשובות:"');

      // אם מצאנו תשובות, נוסיף אותן
      if (directAnswers.length > 0) {
        directAnswers.forEach(function(ans) {
          if (currentQuestion.answers.indexOf(ans) === -1) {
            currentQuestion.answers.push(ans);
          }
        });
        Logger.log('✅ הוספתי ' + directAnswers.length + ' תשובות לשאלה ' + currentQuestion.number);
      } else {
        // אם לא מצאנו, ננסה עם extractAnswersFromLine
        extractAnswersFromLine(currentQuestion, answersText);
      }
      continue;
    }

    // זיהוי תשובה בשורה נפרדת: "א.", "ב.", "A1.", "א1." - תומך גם בלי נקודה
    const answerMatch = line.match(/^([אבגדה]|A\d+|א\d+)[\.:]?\s+(.+)/);
    if (answerMatch && currentQuestion) {
      const answerText = answerMatch[2].trim();
      // ודא שהתשובה לא ריקה ולא מתחילה במספר (שאלה)
      if (answerText && !/^\d+[\.:]/.test(answerText)) {
        currentQuestion.answers.push(answerText);
        Logger.log('  הוספת תשובה: ' + answerText.substring(0, 30));
      }
    } else if (currentQuestion && /^[אבגדה]\s/.test(line)) {
      // תשובה שמתחילה באות עברית ואז רווח (בלי נקודה או נקודתיים)
      const answerMatch2 = line.match(/^([אבגדה])\s+(.+)/);
      if (answerMatch2) {
        const answerText = answerMatch2[2].trim();
        if (answerText && !/^\d+[\.:]/.test(answerText)) {
          currentQuestion.answers.push(answerText);
          Logger.log('  הוספת תשובה (ללא נקודה): ' + answerText.substring(0, 30));
        }
      }
    }
  }

  if (currentQuestion) {
    // נסה לזהות תשובות בשורה האחרונה אם יש
    if (currentQuestion.answers.length === 0 && lines.length > 0) {
      extractAnswersFromLine(currentQuestion, lines[lines.length - 1]);
    }
    questions.push(currentQuestion);
  }

  // בדיקה נוספת - אם שאלות ללא תשובות
  for (let i = 0; i < questions.length; i++) {
    if (questions[i].answers.length === 0) {
      Logger.log('⚠️ אזהרה: שאלה ' + questions[i].number + ' ללא תשובות - טקסט: ' + questions[i].text.substring(0, 50));
    } else {
      Logger.log('✅ שאלה ' + questions[i].number + ': ' + questions[i].answers.length + ' תשובות');
    }
  }

  Logger.log('📊 סה"כ שאלות שנמצאו: ' + questions.length);

  return questions;
}

/**
 * פונקציה עזר לזיהוי תשובות מתוך טקסט (אותו שורה)
 */
function extractAnswersFromLine(question, lineText) {
  if (!lineText || !lineText.trim()) return;

  Logger.log('🔍 מחפש תשובות בטקסט: ' + lineText.substring(0, 100));

  // מחפש תשובות בפורמט: "א. תשובה ב. תשובה ג. תשובה ד. תשובה"
  // או: "א תשובה ב תשובה" (בלי נקודה)
  // מחפש גם תשובות שבאות אחרי מילה אחרת, כמו "א. היא מאביקה פרחי בר..."

  // פורמט 1: "א. טקסט ב. טקסט ג. טקסט ד. טקסט"
  // משפר: מחפש גם תשובות שבאות אחרי מילים/סימנים, כמו "? א. אפריקה ב. אוסטרליה"
  // שינוי: מחפש `([אבגדה])\.\s+` ללא דרישת רווח לפני - יכול לבוא אחרי כל תו
  // משפר: מחפש גם "תשובות: א. תשובה ב. תשובה"
  // שינוי חשוב: [^\d] במקום [^אבגד] כדי לא לכלול מספרים (שאלות) אבל לאפשר כל תו אחר
  // שינוי נוסף: lookahead מדויק יותר - `(?=\s+[אבגדה]\.\s|$)`
  const answerPattern1 = /([אבגדה])\.\s+([^\d]+?)(?=\s+[אבגדה]\.\s|$)/g;
  let match;
  const foundAnswers = [];

  // איפוס regex state
  answerPattern1.lastIndex = 0;

  while ((match = answerPattern1.exec(lineText)) !== null) {
    const answerText = match[2].trim();
    // ודא שהתשובה לא ריקה ולא מתחילה במספר (שאלה)
    // הסר את הבדיקה של length > 2 כדי לתמוך בתשובות קצרות
    if (answerText && answerText.length > 0 && !/^\d+[\.:]/.test(answerText) && !/^תשובות:/i.test(answerText)) {
      // הסר רווחים מיותרים בהתחלה ובסוף
      const cleanAnswer = answerText.replace(/^\s+|\s+$/g, '');
      if (cleanAnswer && cleanAnswer.length > 0 && foundAnswers.indexOf(cleanAnswer) === -1) {
        foundAnswers.push(cleanAnswer);
        Logger.log('  ✅ מצא תשובה בפורמט 1: ' + cleanAnswer.substring(0, 40));
      }
    }
  }

  Logger.log('📊 פורמט 1 מצא: ' + foundAnswers.length + ' תשובות');

  // אם עדיין לא מצאנו מספיק תשובות, ננסה פורמט אחר
  if (foundAnswers.length < 2) {
    Logger.log('⚠️ לא מספיק תשובות - מנסה פורמטים נוספים...');
  }

  // פורמט 2: "א טקסט ב טקסט" (בלי נקודה אבל עם רווח)
  if (foundAnswers.length === 0 || foundAnswers.length < 2) {
    const answerPattern2 = /\s([אבגדה])\s+([^אבגד]+?)(?=\s+[אבגדה]\s|$)/g;
    while ((match = answerPattern2.exec(lineText)) !== null) {
      const answerText = match[2].trim();
      // הסר את הבדיקה של length > 2
      if (answerText && answerText.length > 0 && !/^\d+\./.test(answerText)) {
        const cleanAnswer = answerText.replace(/^\s+|\s+$/g, '');
        if (cleanAnswer && foundAnswers.indexOf(cleanAnswer) === -1) {
          foundAnswers.push(cleanAnswer);
        }
      }
    }
    Logger.log('פורמט 2 מצא: ' + foundAnswers.length + ' תשובות');
  }

  // פורמט 3: "א.תשובה ב.תשובה" (בלי רווח אחרי הנקודה)
  if (foundAnswers.length === 0 || foundAnswers.length < 2) {
    const answerPattern3 = /([אבגדה])\.([^אבגד]+?)(?=[אבגדה]\.|$)/g;
    while ((match = answerPattern3.exec(lineText)) !== null) {
      const answerText = match[2].trim();
      // הסר את הבדיקה של length > 2
      if (answerText && answerText.length > 0 && !/^\d+\./.test(answerText)) {
        const cleanAnswer = answerText.replace(/^\s+|\s+$/g, '');
        if (cleanAnswer && foundAnswers.indexOf(cleanAnswer) === -1) {
          foundAnswers.push(cleanAnswer);
        }
      }
    }
    Logger.log('פורמט 3 מצא: ' + foundAnswers.length + ' תשובות');
  }

  // אם מצאנו תשובות, נוסיף אותן
  if (foundAnswers.length > 0) {
    Logger.log('✅ סה"כ ' + foundAnswers.length + ' תשובות נמצאו לשורה: ' + lineText.substring(0, 50));
    foundAnswers.forEach(function(ans) {
      if (question.answers.indexOf(ans) === -1) {
        question.answers.push(ans);
      }
    });
  } else {
    Logger.log('⚠️ לא נמצאו תשובות בשורה: ' + lineText.substring(0, 80));
  }
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
function createForm(questions, questionnaireName) {
  // אם לא ניתן שם, השתמש בברירת מחדל
  if (!questionnaireName || questionnaireName.trim() === '') {
    questionnaireName = 'שאלון - ' + new Date().toLocaleDateString('he-IL');
  }

  const form = FormApp.create(questionnaireName);
  form.setDescription('שאלון שנוצר אוטומטית');

  form.setAcceptingResponses(true);
  form.setAllowResponseEdits(true);

  // הגדר את הפורם לאסוף כתובות אימייל אוטומטית
  // זה יוצר עמודת "Email address" ב-Sheet ושם את האימייל שם אוטומטית
  form.setCollectEmail(true);

  // הוסף שדה חובה: שם מלא
  const nameItem = form.addTextItem();
  nameItem.setTitle('שם מלא');
  nameItem.setRequired(true);
  nameItem.setHelpText('הזן את שמך המלא');

  const separatorItem = form.addSectionHeaderItem();
  separatorItem.setTitle('שאלות השאלון');

  const hebrewLetters = ['א', 'ב', 'ג', 'ד', 'ה'];

  questions.forEach(function(q) {
    // לוג לבדיקה
    Logger.log('מעבד שאלה ' + q.number + ' עם ' + (q.answers ? q.answers.length : 0) + ' תשובות');

    // ודא שיש תשובות לשאלה
    if (!q.answers || q.answers.length === 0) {
      Logger.log('⚠️ שאלה ' + q.number + ' ללא תשובות - מדלגת');
      // במקום לדלג, נוסיף שאלה ללא תשובות (רק טקסט)
      const item = form.addTextItem();
      item.setTitle('שאלה ' + q.number + ': ' + q.text);
      item.setRequired(false);
      item.setHelpText('⚠️ שאלה זו לא כוללת תשובות - נא למלא תשובה חופשית');
      return;
    }

    const item = form.addMultipleChoiceItem();
    item.setTitle('שאלה ' + q.number + ': ' + q.text);
    item.setRequired(false);

    const choices = [];
    q.answers.forEach(function(answer, index) {
      if (answer && answer.trim()) {
        const letter = hebrewLetters[index] || String(index + 1);
        const choiceText = letter + '. ' + answer.trim();
        Logger.log('  הוספת תשובה: ' + choiceText);
        choices.push(item.createChoice(choiceText));
      }
    });

    // אם אין תשובות תקפות, הוסף תשובה ברירת מחדל
    if (choices.length === 0) {
      Logger.log('⚠️ שאלה ' + q.number + ' - הוספת תשובה ברירת מחדל');
      choices.push(item.createChoice('א. אין תשובות'));
    }

    Logger.log('  הוספתי ' + choices.length + ' תשובות לשאלה ' + q.number);
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

    // קבל את האימייל אוטומטית מה-Sheet (האימייל האוטומטי שהמשתמש רואה בפורם)
    // Google Forms שומר את האימייל האוטומטי בעמודה "Email address" ב-Sheet
    const headers = responsesSheet.getRange(1, 1, 1, responsesSheet.getLastColumn()).getValues()[0];

    // חפש עמודת אימייל קיימת (Google Forms יוצר "Email address" אוטומטית)
    let emailCol = headers.indexOf('Email address') + 1;
    if (emailCol === 0) {
      emailCol = headers.indexOf('כתובת אימייל') + 1;
    }
    if (emailCol === 0) {
      emailCol = headers.indexOf('Email') + 1;
    }

    // קרא את האימייל אוטומטית מהעמודה "Email address" ב-Sheet
    // האימייל תמיד נמצא שם (למשל: "1004389819@manhischools.org.il")
    let respondentEmail = null;
    if (emailCol > 0) {
      const emailFromSheet = responsesSheet.getRange(lastResponseRow, emailCol).getValue();
      if (emailFromSheet && emailFromSheet.toString().trim() !== '') {
        respondentEmail = emailFromSheet.toString().trim();
        Logger.log('✅ אימייל אוטומטי נמצא ב-Sheet (Email address): ' + respondentEmail);
      } else {
        Logger.log('⚠️ לא נמצא אימייל בשורה ' + lastResponseRow + ' בעמודה ' + emailCol);
      }
    } else {
      Logger.log('⚠️ לא נמצאה עמודת Email address ב-Sheet');
    }

    // אם לא מצאנו ב-Sheet, נסה לקבל מהתגובה
    if (!respondentEmail || respondentEmail.trim() === '') {
      const emailFromResponse = formResponse.getRespondentEmail();
      if (emailFromResponse && emailFromResponse.trim() !== '') {
        respondentEmail = emailFromResponse.trim();
        Logger.log('✅ אימייל נמצא מהתגובה: ' + respondentEmail);
      } else {
        Logger.log('⚠️ לא נמצא אימייל מהתגובה');
      }
    }

    // שמור את האימייל האוטומטי ב-Sheet אוטומטית - יחד עם כל תוצאות השאלון
    if (respondentEmail && respondentEmail.trim() !== '') {
      if (emailCol > 0) {
        // ודא שהאימייל נשמר בעמודה בשורה האחרונה (יחד עם תוצאות השאלון)
        const existingEmail = responsesSheet.getRange(lastResponseRow, emailCol).getValue();
        if (!existingEmail || existingEmail === '' || existingEmail === null) {
          // אם האימייל לא שם, שמור אותו ידנית בשורה האחרונה
          responsesSheet.getRange(lastResponseRow, emailCol).setValue(respondentEmail);
          Logger.log('✅ אימייל האוטומטי נשמר ב-Sheet בשורה ' + lastResponseRow + ': ' + respondentEmail);
        } else {
          // האימייל כבר שם - ודא שאנחנו משתמשים באימייל הנכון
          respondentEmail = existingEmail.toString().trim();
          Logger.log('✅ אימייל כבר נשמר ב-Sheet ונשתמש בו: ' + respondentEmail);
        }
      } else {
        // אם אין עמודה, הוסף אחת חדשה ושמור את האימייל בשורה האחרונה
        emailCol = responsesSheet.getLastColumn() + 1;
        responsesSheet.getRange(1, emailCol).setValue('כתובת אימייל').setFontWeight('bold');
        responsesSheet.getRange(lastResponseRow, emailCol).setValue(respondentEmail);
        Logger.log('✅ אימייל האוטומטי נשמר בעמודה חדשה בשורה ' + lastResponseRow + ': ' + respondentEmail);
      }
    } else {
      Logger.log('⚠️ לא נמצא אימייל אוטומטי של הנבחן - לא נשמר ב-Sheet');
    }

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

    // שליחת אימייל לנבחן עם הציון (שימוש באימייל האוטומטי מה-Sheet)
    // האימייל תמיד נמצא בעמודת "Email address" ב-Sheet - אחרת הנבחן לא היה רואה את השאלון
    // ודא שיש לנו את האימייל מהעמודה "Email address" ב-Sheet
    if (!respondentEmail || respondentEmail.trim() === '' || emailCol === 0) {
      // קרא את האימייל מהעמודה "Email address" ב-Sheet
      if (emailCol > 0) {
        const emailFromSheet = responsesSheet.getRange(lastResponseRow, emailCol).getValue();
        if (emailFromSheet && emailFromSheet.toString().trim() !== '') {
          respondentEmail = emailFromSheet.toString().trim();
          Logger.log('✅ אימייל נקרא מהעמודה Email address: ' + respondentEmail);
        }
      }
    }

    // שלח את האימייל עם הציון - האימייל תמיד זמין
    try {
      // ודא שיש לנו את האימייל לפני שליחה
      if (!respondentEmail || respondentEmail.trim() === '') {
        Logger.log('⚠️ לא נמצא אימייל - מנסה לקרוא מהעמודה Email address');
        if (emailCol > 0) {
          const emailFromSheet = responsesSheet.getRange(lastResponseRow, emailCol).getValue();
          if (emailFromSheet && emailFromSheet.toString().trim() !== '') {
            respondentEmail = emailFromSheet.toString().trim();
            Logger.log('✅ אימייל נקרא מהעמודה Email address: ' + respondentEmail);
          }
        }
      }

      // אם עדיין אין אימייל, נסה מהתגובה
      if (!respondentEmail || respondentEmail.trim() === '') {
        const emailFromResponse = formResponse.getRespondentEmail();
        if (emailFromResponse && emailFromResponse.trim() !== '') {
          respondentEmail = emailFromResponse.trim();
          Logger.log('✅ אימייל נמצא מהתגובה: ' + respondentEmail);
        }
      }

      // שלח את האימייל רק אם יש אימייל
      if (respondentEmail && respondentEmail.trim() !== '' && respondentEmail !== 'null' && respondentEmail !== null) {
        // קבל את שם הנבחן מהתגובה
        let respondentName = '';
        for (let j = 0; j < itemResponses.length; j++) {
          const item = itemResponses[j].getItem();
          if (item.getTitle() === 'שם מלא') {
            respondentName = itemResponses[j].getResponse();
            break;
          }
        }

        const subject = 'הטופס הוגש - הציון שלך';
        const greeting = respondentName ? 'שלום ' + respondentName + ',\n\n' : 'שלום,\n\n';
        const message = greeting +
          'הטופס הוגש בהצלחה!\n\n' +
          'הציון שלך: ' + score + ' (' + percentage + '%)\n\n' +
          'תודה על ההשתתפות!';

        MailApp.sendEmail({
          to: respondentEmail,
          subject: subject,
          body: message
        });

        Logger.log('✅ אימייל נשלח בהצלחה לנבחן: ' + respondentEmail + ' עם ציון ' + score + ' (' + percentage + '%)');
      } else {
        Logger.log('⚠️ לא ניתן לשלוח אימייל - לא נמצא אימייל');
        Logger.log('   respondentEmail: ' + (respondentEmail ? respondentEmail : 'null'));
        Logger.log('   emailCol: ' + emailCol);
      }
    } catch (sendError) {
      Logger.log('❌ שגיאה בשליחת האימייל: ' + sendError.toString());
      Logger.log('   האימייל: ' + respondentEmail);
      Logger.log('   הציון: ' + score);
      Logger.log('   שגיאה מלאה: ' + sendError.stack);
    }

  } catch (error) {
    Logger.log('❌ שגיאה: ' + error.toString());
  }
}

