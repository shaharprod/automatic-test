/**
 * Apps Script מלא - יצירת Google Sheet ו-Google Form מטקסט שאלון
 *
 * שימוש:
 * 1. פתח Google Drive
 * 2. לחץ על New > More > Google Apps Script
 * 3. העתק את הקוד הזה
 * 4. הרץ את הפונקציה createFormFromText()
 */

/**
 * פונקציה ראשית - יוצרת Sheet ו-Form מטקסט
 * ערוך את המשתנה QUESTIONNAIRE_TEXT למטה עם השאלון שלך
 */
function createFormFromText() {
  try {
    // ========================================
    // הכנס כאן את השאלון שלך
    // ========================================
    const QUESTIONNAIRE_TEXT = `Q1. מהו העיקרון המרכזי של "למידה מונחית" (Guided Learning) כפי שהיא מיושמת בג'מיני?
A1. מתן תשובה מלאה ומקיפה לנושא מורכב באמצעות סיכום טקסטואלי.
A2. שימוש במודלים מתקדמים (כמו Gemini Pro) כדי ליצור תמונות וסרטוני וידאו.
A3. בניית דיאלוג אינטראקטיבי שמפרק בעיות ומקדם את הלומד צעד אחר צעד באמצעות שאלות קטנות ומדויקות.
A4. יצירת מחוונים מותאמים אישית אוטומטית, ללא צורך בהתערבות המורה.

Q2. על פי עקרונות גוגל, מהו תפקידה של הבינה המלאכותית (AI) ביחס לצוותי חינוך?
A1. ה-AI נועד להחליף מורים בתחומי ידע ספציפיים כדי לחסוך זמן.
A2. ה-AI נועד לספק משוב מיידי במקום המורה ולהוות את מרכז ההערכה.
A3. ה-AI לעולם לא יכול להחליף את המומחיות, הידע או היצירתיות של איש חינוך, אך הוא כלי מועיל לשיפור החוויה.
A4. ה-AI משמש בעיקר לשיפור האבטחה הדיגיטלית ולחסימת תוכנות כופר ב-Chromebooks.

Q3. איזה אחוז מהמורים בישראל מדווחים, על פי מחקר לאומי, כי הם משתמשים בפועל בכלי בינה מלאכותית יוצרת להכנה ותכנון שיעורים?
A1. כרבע מהמורים (25%).
A2. כשליש מהמורים (33%).
A3. כמחצית מהמורים (50%).
A4. למעלה מ-75% מהמורים.

Q4. איזה מהבאים אינו דוגמה לשימוש ש-Gemini יכול לבצע עבור מורה ב-Google Classroom?
A1. יצירת ראשי פרקים למערך שיעור.
A2. שינוי הרמה של טקסט כדי שיתאים לרמת קריאה מסוימת.
A3. המרת קריטריונים להערכה מקבצי Drive לפורמט של Google Classroom.
A4. ניתוח ציוני תלמידים והוצאת גרפים אוטומטיים מתוך Google Sheets.

Q5. על פי המחקר הישראלי, איזה חשש הוא אחד המרכזיים של המורים בישראל בנוגע ל-AI, שאינו קשור ישירות לאיום על מקום עבודתם האישי?
A1. התלות הגבוהה ב-API חיצוניים.
A2. המשך קיום המין האנושי, אפליה ופגיעה בפרטיות.
A3. הדרישה לעבור משיטות הערכה מסורתיות לשיטות חלופיות.
A4. המחסור ביכולת לשלב מדיה ויזואלית בתגובות של המודל.

Q6. מהו המונח המתאר את משפחת המודלים שעברה כוונון עדין (fine-tuned) עבור למידה, מודל המוטמע ישירות בג'מיני ומונחה על ידי עקרונות מדעי הלמידה?
A1. Gemini Pro.
A2. Gemini Flash.
A3. LearnLM.
A4. Copilot.

Q7. מדוע קיימת דאגה מפני "פצצת זמן חברתית" בהקשר של אימוץ AI במערכת החינוך בישראל, על פי המחקר?
A1. מפני שהמורים "הפתוחים, אך חרדים" הם הקבוצה הגדולה ביותר.
A2. מפני שבבתי ספר ממעמדות חברתיים-כלכליים "חלשים", מורים נוטים פחות לאמץ בינה מלאכותית בהוראה, מה שמעמיק פערים.
A3. מפני שרק רבע מהמורים מדווחים על שינוי משמעותי בשיטות ההוראה.
A4. מפני שהמורים משתמשים בטכנולוגיה חדשנית כדי לייעל פדגוגיה ישנה.

Q8. מהי ההגדרה של בינה מלאכותית יוצרת (Gen AI)?
A1. מודל למידת מכונה (ML) שיכול רק להבין ולנתח שפה טבעית.
A2. טכניקה המאפשרת למכונות ללמוד באופן אוטונומי מנתונים קיימים.
A3. סוג של AI המתמקד ביצירת תוכן חדש (כגון טקסט, תמונות, קוד או וידאו) באמצעות הנחיה פשוטה.
A4. מערכת מחשב המלמדת לחקות אינטליגנציה טבעית.

Q9. כיצד מוגנים נתוני צ'אט (שיחה) של משתמשים המחוברים ל-Gemini באמצעות חשבון Google Workspace for Education?
A1. הנתונים נשמרים אך משמשים לשיפור כללי של מודלי ה-AI בהמשך.
A2. הנתונים עוברים בדיקה על ידי בודקים אנושיים כדי למנוע יצירת תוכן מזיק.
A3. הנתונים נמחקים אוטומטית מיד לאחר סיום סשן הצ'אט.
A4. השיחות אינן נבדקות על ידי בני אדם ואינן משמשות לאימון או שיפור מודלי ה-AI.

Q10. איזה פרופיל של מורים זיהה המחקר הישראלי כקבוצה הקטנה ביותר, המשתמשת ברמה גבוהה ב-GenAI?
A1. "סגורים ואדישים" (בעלי תפיסת תועלת נמוכה).
A2. "פתוחים, אך חרדים" (מכירים בתועלת אך מביעים חששות רבות).
A3. "פתוחים ובטוחים" (עושים שימוש גבוה ב-GenAI).
A4. "הלומדים המונחים".`;

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

    // חיבור הפורם ל-Sheet (נדרש לקבלת טריגרים)
    const ss = SpreadsheetApp.openByUrl(sheet.url);
    const formObj = FormApp.openById(form.id);

    // חיבור הפורם ל-Sheet - זה יוצר טאב תשובות אוטומטית
    try {
      formObj.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
      Logger.log('✅ הפורם חובר ל-Sheet בהצלחה');
    } catch (error) {
      Logger.log('⚠️ שגיאה בחיבור הפורם ל-Sheet: ' + error.toString());
    }

    // המתן קצת כדי שה-Sheet יעודכן
    Utilities.sleep(2000);

    // הוספת עמודה לתשובה הנכונה ב-Sheet
    const sheetTab = ss.getSheets()[0];
    sheetTab.getRange(1, 9).setValue('תשובה נכונה');

    // הוספת פונקציה לחישוב ציונים - רק אחרי שהפורם מחובר ל-Sheet
    // חשוב: הטריגר צריך להיות מוגדר אחרי שהפורם מחובר
    try {
      setupGrading(ss.getId(), formObj);
    } catch (error) {
      Logger.log('⚠️ לא ניתן ליצור טריגר אוטומטית: ' + error.toString());
      Logger.log('💡 פתרון: הוסף טריגר ידנית - ראה הוראות ב-פתרון_בעיות.md');
    }

    // שמירת קישורים ב-Sheet
    sheetTab.getRange(1, 10).setValue('Form URL');
    sheetTab.getRange(1, 11).setValue('Form ID');
    sheetTab.getRange(2, 10).setValue(form.url);
    sheetTab.getRange(2, 11).setValue(form.id);

    // הצגת הודעה ב-Logger (תמיד עובד)
    const successMessage =
      '✅ הושלם בהצלחה!\n\n' +
      '📊 Google Sheet:\n' + sheet.url + '\n\n' +
      '📝 Google Form:\n' + form.url + '\n\n' +
      'הקישורים נשמרו גם בעמודות J ו-K ב-Sheet\n\n' +
      '📝 הוראות להגדרת תשובות נכונות:\n' +
      'ערוך את הפונקציה getCorrectAnswers() והכנס את התשובות הנכונות בפורמט:\n' +
      '"מספר_שאלה: תשובה" - למשל: "06: א" (שאלה 6, תשובה א)\n' +
      'שימוש: א, ב, ג, ד (או א1, ב1 - שתי הצורות נתמכות)\n\n' +
      'הציונים יחושבו אוטומטית בעת הגשת כל תשובה!';

    Logger.log(successMessage);

    // נסה להציג חלון הודעה (אם אפשר)
    try {
      Browser.msgBox('✅ הושלם בהצלחה!',
        '📊 Google Sheet: ' + sheet.url + '\n\n' +
        '📝 Google Form: ' + form.url + '\n\n' +
        'הקישורים נשמרו גם בעמודות J ו-K ב-Sheet',
        Browser.Buttons.OK);
    } catch (e) {
      // אם Browser.msgBox לא עובד, רק נשתמש ב-Logger
      Logger.log('ℹ️ בדוק את היומן (Logs) לראות את הקישורים');
    }

    return {
      success: true,
      sheet: sheet,
      form: form
    };

  } catch (error) {
    const errorMessage = '❌ שגיאה: ' + error.toString();
    Logger.log(errorMessage);

    // נסה להציג הודעת שגיאה (אם אפשר)
    try {
      Browser.msgBox('❌ שגיאה', errorMessage, Browser.Buttons.OK);
    } catch (e) {
      // אם Browser.msgBox לא עובד, רק נשתמש ב-Logger
      Logger.log('ℹ️ בדוק את היומן (Logs) לראות את פרטי השגיאה');
    }

    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * מפרסר טקסט שאלון לפורמט מובנה
 */
function parseQuestionnaire(text) {
  const questions = [];
  const lines = text.split('\n');

  let currentQuestion = null;
  let currentAnswers = [];

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();

    if (!line) {
      continue; // דלג על שורות ריקות
    }

    // בדיקה אם זו שאלה (שאלה 1., שאלה 2., וכו' או ש1, Q1 למידה לאחור)
    if (line.match(/^שאלה\s+\d+\.\s*(.+)/) || line.match(/^ש\d+\.\s*(.+)/) || line.match(/^Q\d+\.\s*(.+)/)) {
      // שמור שאלה קודמת אם יש
      if (currentQuestion) {
        questions.push({
          number: questions.length + 1,
          question: currentQuestion,
          answers: currentAnswers,
          type: currentAnswers.length > 0 ? 'multiple_choice' : 'text'
        });
      }

      // התחל שאלה חדשה - תמיכה בכמה פורמטים
      let match = line.match(/^שאלה\s+(\d+)\.\s*(.+)/);
      if (match) {
        currentQuestion = match[2].trim();
        currentAnswers = [];
      } else {
        match = line.match(/^ש\d+\.\s*(.+)/) || line.match(/^Q\d+\.\s*(.+)/);
        if (match) {
          currentQuestion = match[1].trim();
          currentAnswers = [];
        }
      }
    }
    // בדיקה אם זו תשובה (א., ב., ג., ד. או א1, A1 למידה לאחור)
    else if (line.match(/^[אבגד]\.\s*(.+)/) || line.match(/^א\d+\.\s*(.+)/) || line.match(/^A\d+\.\s*(.+)/)) {
      let match = line.match(/^([אבגד])\.\s*(.+)/);
      if (match) {
        currentAnswers.push(match[2].trim());
      } else {
        match = line.match(/^א\d+\.\s*(.+)/) || line.match(/^A\d+\.\s*(.+)/);
        if (match) {
          currentAnswers.push(match[1].trim());
        }
      }
    }
  }

  // שמור שאלה אחרונה
  if (currentQuestion) {
    questions.push({
      number: questions.length + 1,
      question: currentQuestion,
      answers: currentAnswers,
      type: currentAnswers.length > 0 ? 'multiple_choice' : 'text'
    });
  }

  return questions;
}

/**
 * יוצר Google Sheet עם הנתונים
 */
function createSpreadsheet(questions) {
  const ss = SpreadsheetApp.create('שאלון לטפסים - ' + new Date().toLocaleDateString('he-IL'));
  const sheet = ss.getSheets()[0];
  sheet.setName('שאלון');

  // כותרות
  const headers = [['מספר', 'שאלה', 'סוג', 'תשובה 1', 'תשובה 2', 'תשובה 3', 'תשובה 4', 'תשובה 5']];
  sheet.getRange(1, 1, 1, 8).setValues(headers);
  sheet.getRange(1, 1, 1, 8).setFontWeight('bold');

  // הוספת כותרות לתשובות נכונות ולציון בסוף Sheet
  sheet.getRange(1, 9).setValue('תשובה נכונה');
  sheet.getRange(1, 9).setFontWeight('bold');

  // קבלת תשובות נכונות
  const correctAnswers = getCorrectAnswers();

  // נתונים
  const data = [];
  questions.forEach(function(q) {
    // קבלת תשובה נכונה - פורמט מספר שאלה עם אפס מוביל (01, 02, וכו')
    const questionNum = String(q.number).padStart(2, '0');
    const correctAnswer = correctAnswers[questionNum] || '';

    const row = [
      q.number,
      q.question,
      q.type,
      q.answers[0] || '',
      q.answers[1] || '',
      q.answers[2] || '',
      q.answers[3] || '',
      q.answers[4] || '',
      correctAnswer
    ];
    data.push(row);
  });

  if (data.length > 0) {
    sheet.getRange(2, 1, data.length, 9).setValues(data);
  }

  // עיצוב
  sheet.setColumnWidth(1, 60);  // מספר
  sheet.setColumnWidth(2, 400); // שאלה
  sheet.setColumnWidth(3, 120); // סוג
  sheet.setColumnWidths(4, 5, 200); // תשובות

  return {
    id: ss.getId(),
    url: ss.getUrl()
  };
}

/**
 * הגדרת תשובות נכונות - עדכן כאן את התשובות הנכונות
 * פורמט: "מספר_שאלה: תשובה" - למשל: "06: א" (שאלה 6, תשובה א)
 *
 * הערה: התשובות הנכונות נשמרות גם בעמודה 9 ב-Sheet (טאב "שאלון")
 * ניתן לערוך שם ישירות במקום בקוד - הפונקציה onFormSubmit תקרא משם
 *
 * שימוש: א, ב, ג, ד (או א1, ב1, וכו' - שתי הצורות נתמכות)
 */
function getCorrectAnswers() {
  return {
    '01': 'א',   // שאלה 1 - תשובה א
    '02': 'ג',   // שאלה 2 - תשובה ג
    '03': 'א',   // שאלה 3 - תשובה א
    '04': 'ד',   // שאלה 4 - תשובה ד
    '05': 'ב',   // שאלה 5 - תשובה ב
    '06': 'א',   // שאלה 6 - תשובה א (מונח LearnLM)
    '07': 'ב',   // שאלה 7 - תשובה ב
    '08': 'ג',   // שאלה 8 - תשובה ג
    '09': 'ד',   // שאלה 9 - תשובה ד
    '10': 'א'    // שאלה 10 - תשובה א
  };
}

/**
 * הגדרת חישוב ציונים - נקרא אוטומטית בעת הגשת פורם
 */
function setupGrading(spreadsheetId, form) {
  try {
    // פתח את ה-Sheet
    const ss = SpreadsheetApp.openById(spreadsheetId);

    // מחק טריגרים קיימים (אם יש) עבור אותו פורם
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(function(trigger) {
      if (trigger.getHandlerFunction() === 'onFormSubmit') {
        try {
          ScriptApp.deleteTrigger(trigger);
          Logger.log('🗑️ טריגר קיים נמחק');
        } catch (e) {
          // התעלם משגיאות מחיקה
        }
      }
    });

    // הוסף טריגר חדש - מופעל בעת הגשת פורם
    // חשוב: צריך שהפורם יהיה מחובר ל-Sheet לפני יצירת הטריגר
    // ננסה כמה דרכים ליצירת הטריגר

    try {
      // דרך 1: טריגר על הפורם ישירות
      ScriptApp.newTrigger('onFormSubmit')
        .forForm(form)
        .onFormSubmit()
        .create();
      Logger.log('✅ טריגר לחישוב ציונים הוגדר בהצלחה (על הפורם)!');
    } catch (error1) {
      try {
        // דרך 2: טריגר על ה-Sheet (כשהפורם מחובר)
        ScriptApp.newTrigger('onFormSubmit')
          .onFormSubmit()
          .create();
        Logger.log('✅ טריגר לחישוב ציונים הוגדר בהצלחה (על ה-Sheet)!');
      } catch (error2) {
        // אם גם זה לא עובד, נתן הוראות ידניות
        throw new Error('לא ניתן ליצור טריגר אוטומטית: ' + error1.toString() + ' | ' + error2.toString());
      }
    }

    Logger.log('ℹ️ הטריגר יפעיל את הפונקציה onFormSubmit() בכל פעם שמישהו מגיש את הפורם');

  } catch (error) {
    Logger.log('❌ שגיאה בהגדרת טריגר: ' + error.toString());
    Logger.log('💡 פתרון: הוסף טריגר ידנית:');
    Logger.log('   1. ב-Apps Script: ⚙️ → Triggers → + Add Trigger');
    Logger.log('   2. Choose function: onFormSubmit');
    Logger.log('   3. Select event source: From form');
    Logger.log('   4. Select event type: On form submit');
    Logger.log('   5. Save');
    Logger.log('');
    Logger.log('💡 ודא שהפורם מחובר ל-Sheet (בתפריט הפורם: Responses → Link to Sheets)');
  }
}

/**
 * פונקציה המופעלת אוטומטית בעת הגשת פורם
 * חישוב ציונים על בסיס תשובות נכונות
 */
function onFormSubmit(e) {
  try {
    Logger.log('🔄 פונקציה onFormSubmit הופעלה - מתחיל חישוב ציונים...');

    // בדיקה שהפרמטר תקין
    if (!e || !e.response || !e.source) {
      Logger.log('❌ שגיאה: הפרמטר של onFormSubmit לא תקין');
      Logger.log('📋 זה יכול לקרות אם:');
      Logger.log('   1. הטריגר לא מוגדר נכון');
      Logger.log('   2. הפורם לא מחובר ל-Sheet');
      Logger.log('   3. הפונקציה נקראת ידנית (לא דרך טריגר)');
      Logger.log('');
      Logger.log('💡 פתרון 1: ודא שהטריגר מוגדר:');
      Logger.log('   ב-Apps Script → ⚙️ → Triggers');
      Logger.log('   ודא שיש: onFormSubmit → From form → On form submit');
      Logger.log('');
      Logger.log('💡 פתרון 2: ודא שהפורם מחובר ל-Sheet:');
      Logger.log('   פתח את הפורם → Responses → לחץ "Link to Sheets"');
      Logger.log('');
      Logger.log('💡 פתרון 3: חשב ציונים ידנית:');
      Logger.log('   הרץ את calculateScoreManually() (ראה הוראות בקוד)');
      Logger.log('');
      Logger.log('📝 פרטים טכניים:');
      Logger.log('   e = ' + (e ? 'קיים אבל לא תקין' : 'undefined'));
      Logger.log('   e.response = ' + (e && e.response ? 'קיים' : 'לא קיים'));
      Logger.log('   e.source = ' + (e && e.source ? 'קיים' : 'לא קיים'));
      return;
    }

    const formResponse = e.response;
    const form = e.source;

    // קבלת ה-Sheet המקושר לפורם
    let ss = null;
    try {
      const destinationId = form.getDestinationId();
      if (destinationId) {
        ss = SpreadsheetApp.openById(destinationId);
        Logger.log('✅ Sheet נמצא: ' + destinationId);
      } else {
        Logger.log('⚠️ הפורם לא מקושר ל-Sheet');
        Logger.log('💡 פתרון: פתח את הפורם → Responses → לחץ "Link to Sheets"');
        return;
      }
    } catch (error) {
      Logger.log('❌ שגיאה בקבלת Sheet: ' + error.toString());
      return;
    }

    if (!ss) {
      Logger.log('❌ לא נמצא Sheet מקושר לפורם');
      return;
    }

    // מצא את הטאב של התשובות (נוצר אוטומטית ע"י Google Forms)
    const sheets = ss.getSheets();
    let responsesSheet = null;
    for (let i = 0; i < sheets.length; i++) {
      const sheetName = sheets[i].getName();
      if (sheetName === 'טופס 1' || sheetName.includes('תשובות') || sheetName === 'Form Responses 1') {
        responsesSheet = sheets[i];
        break;
      }
    }

    if (!responsesSheet) {
      Logger.log('⚠️ לא נמצא טאב תשובות');
      return;
    }

    // מצא את הטאב עם הגדרות התשובות הנכונות
    let settingsSheet = null;
    for (let i = 0; i < sheets.length; i++) {
      const sheetName = sheets[i].getName();
      if (sheetName === 'שאלון') {
        settingsSheet = sheets[i];
        break;
      }
    }

    if (!settingsSheet) {
      Logger.log('⚠️ לא נמצא טאב "שאלון" עם הגדרות');
      return;
    }

    // קבלת כל השאלות והתשובות הנכונות
    const lastRow = settingsSheet.getLastRow();
    const questionData = settingsSheet.getRange(2, 1, lastRow - 1, 9).getValues();

    let totalQuestions = 0;
    let correctCount = 0;

    // קבלת התשובות מהפורם
    const itemResponses = formResponse.getItemResponses();
    const lastResponseRow = responsesSheet.getLastRow();

    // קבלת כותרות
    const headers = responsesSheet.getRange(1, 1, 1, responsesSheet.getLastColumn()).getValues()[0];

    // מצא עמודות ציון ואחוז (אם עוד לא קיימות)
    let scoreCol = headers.indexOf('ציון') + 1;
    let percentCol = headers.indexOf('אחוז') + 1;

    if (scoreCol === 0 || headers.indexOf('ציון') === -1) {
      scoreCol = responsesSheet.getLastColumn() + 1;
      responsesSheet.getRange(1, scoreCol).setValue('ציון');
      responsesSheet.getRange(1, scoreCol).setFontWeight('bold');
    }

    if (percentCol === 0 || headers.indexOf('אחוז') === -1) {
      percentCol = responsesSheet.getLastColumn() + 1;
      if (percentCol === scoreCol) percentCol++;
      responsesSheet.getRange(1, percentCol).setValue('אחוז');
      responsesSheet.getRange(1, percentCol).setFontWeight('bold');
    }

    Logger.log('📊 נמצאו ' + questionData.length + ' שאלות בטאב "שאלון"');
    Logger.log('📝 נמצאו ' + itemResponses.length + ' תשובות מהפורם');

    // בדוק כל שאלה
    for (let i = 0; i < questionData.length; i++) {
      const questionNum = questionData[i][0]; // מספר שאלה
      const correctAnswer = questionData[i][8]; // תשובה נכונה (עמודה 9)

      if (!correctAnswer) {
        Logger.log('⚠️ שאלה ' + questionNum + ' - אין תשובה נכונה מוגדרת');
        continue;
      }

      totalQuestions++;

      // מצא את התשובה של המשתמש לפי מספר השאלה
      // תמיכה בשני פורמטים: "ש1:" או "שאלה 1:"
      const questionHeader1 = 'ש' + questionNum + ':';
      const questionHeader2 = 'שאלה ' + questionNum + ':';
      let userAnswer = null;

      // חיפוש בתגובות הפורם
      for (let j = 0; j < itemResponses.length; j++) {
        const itemTitle = itemResponses[j].getItem().getTitle();
        if (itemTitle && (itemTitle.includes(questionHeader1) || itemTitle.includes(questionHeader2))) {
          userAnswer = itemResponses[j].getResponse();
          Logger.log('✅ שאלה ' + questionNum + ' - תשובה נמצאת: ' + userAnswer);
          break;
        }
      }

      if (!userAnswer) {
        Logger.log('⚠️ שאלה ' + questionNum + ' - לא נמצאה תשובה מהמשתמש');
        continue;
      }

      // השווה תשובות - חיפוש במערך התשובות האפשריות
      const answerOptions = [
        questionData[i][3], // א1 / A1
        questionData[i][4], // א2 / A2
        questionData[i][5], // א3 / A3
        questionData[i][6], // א4 / A4
        questionData[i][7]  // א5 / A5
      ].filter(function(opt) { return opt && opt.trim() !== ''; });

      // מצא את אינדקס התשובה הנכונה - תמיכה בא, ב, ג, ד וגם א1, A1, וכו'
      let correctAnswerIndex = -1;
      const hebrewLetters = ['א', 'ב', 'ג', 'ד', 'ה'];

      // פורמט חדש: א, ב, ג, ד
      if (correctAnswer === 'א' || correctAnswer === 'א1' || correctAnswer === 'A1') {
        correctAnswerIndex = 0; // תשובה ראשונה
      } else if (correctAnswer === 'ב' || correctAnswer === 'א2' || correctAnswer === 'A2') {
        correctAnswerIndex = 1; // תשובה שנייה
      } else if (correctAnswer === 'ג' || correctAnswer === 'א3' || correctAnswer === 'A3') {
        correctAnswerIndex = 2; // תשובה שלישית
      } else if (correctAnswer === 'ד' || correctAnswer === 'א4' || correctAnswer === 'A4') {
        correctAnswerIndex = 3; // תשובה רביעית
      } else {
        // פורמט ישן: א1, א2, וכו'
        let answerNum = correctAnswer;
        if (answerNum.includes('א')) {
          answerNum = answerNum.replace('א', '');
        } else if (answerNum.includes('A')) {
          answerNum = answerNum.replace('A', '');
        }
        correctAnswerIndex = parseInt(answerNum) - 1;
      }

      if (correctAnswerIndex >= 0 && correctAnswerIndex < answerOptions.length) {
        // בניה של התשובה הנכונה עם האות העברית (כפי שמופיעה בפורם)
        const correctAnswerText = answerOptions[correctAnswerIndex];
        const correctAnswerWithLetter = hebrewLetters[correctAnswerIndex] + '. ' + correctAnswerText;

        Logger.log('🔍 שאלה ' + questionNum + ':');
        Logger.log('   תשובה נכונה: ' + correctAnswerWithLetter);
        Logger.log('   תשובת משתמש: ' + userAnswer);

        // השווה - התשובה מהפורם יכולה להיות עם או בלי האות
        // בדוק אם התשובה מתחילה באות עברית (כפי שמופיעה בפורם)
        const userAnswerClean = userAnswer.toString().trim();
        const correctAnswerClean = correctAnswerWithLetter.trim();
        const correctAnswerCleanWithoutLetter = correctAnswerText.trim();

        // בדוק התאמה מלאה (עם אות) או התאמה של הטקסט בלבד (בלי אות)
        if (userAnswerClean === correctAnswerClean ||
            userAnswerClean === correctAnswerCleanWithoutLetter ||
            userAnswerClean.startsWith(hebrewLetters[correctAnswerIndex] + '.') ||
            userAnswerClean.endsWith(correctAnswerText.trim())) {
          correctCount++;
          Logger.log('   ✅ תשובה נכונה!');
        } else {
          Logger.log('   ❌ תשובה שגויה');
          Logger.log('   השוואה: [' + userAnswerClean + '] vs [' + correctAnswerClean + ']');
        }
      } else {
        Logger.log('⚠️ שאלה ' + questionNum + ' - אינדקס תשובה נכונה לא תקין: ' + correctAnswerIndex);
      }
    }

    // חשב ציון ואחוז
    const score = correctCount + '/' + totalQuestions;
    const percentage = totalQuestions > 0 ? Math.round((correctCount / totalQuestions) * 100) : 0;

    // עדכן את הציון בשורה האחרונה
    responsesSheet.getRange(lastResponseRow, scoreCol).setValue(score);
    responsesSheet.getRange(lastResponseRow, percentCol).setValue(percentage + '%');

    // הדגש שורה לפי ציון
    if (percentage >= 80) {
      responsesSheet.getRange(lastResponseRow, scoreCol, 1, 2).setBackground('#c6efce'); // ירוק
    } else if (percentage >= 60) {
      responsesSheet.getRange(lastResponseRow, scoreCol, 1, 2).setBackground('#ffeb9c'); // צהוב
    } else {
      responsesSheet.getRange(lastResponseRow, scoreCol, 1, 2).setBackground('#ffc7ce'); // אדום
    }

    Logger.log('✅ ציון מחושב: ' + score + ' (' + percentage + '%)');
    Logger.log('🎉 סיום חישוב ציונים בהצלחה!');

  } catch (error) {
    Logger.log('❌ שגיאה בחישוב ציונים: ' + error.toString());
    Logger.log('📋 Stack trace: ' + error.stack);
  }
}

/**
 * פונקציה לחישוב ציונים ידנית - עם Sheet ID ידוע
 * שימוש: עדכן את SHEET_ID למטה והרץ את הפונקציה
 */
function calculateScoreWithSheetId() {
  try {
    Logger.log('🔍 מחשב ציונים ידנית...');

    // מצא את ה-Sheet אוטומטית דרך הטריגרים או דרך פורמים מחוברים
    let SHEET_ID = null;
    let FORM_ID = null;

    // נסה למצוא מהטריגרים
    try {
      const triggers = ScriptApp.getProjectTriggers();
      for (let i = 0; i < triggers.length; i++) {
        if (triggers[i].getHandlerFunction() === 'onFormSubmit') {
          const sourceId = triggers[i].getSourceId();
          if (sourceId) {
            try {
              // נסה לפתוח כ-Sheet
              const ss = SpreadsheetApp.openById(sourceId);
              SHEET_ID = sourceId;
              Logger.log('✅ נמצא Sheet ID מהטריגר: ' + sourceId);

              // נסה למצוא Form ID מה-Sheet
              const sheets = ss.getSheets();
              for (let j = 0; j < sheets.length; j++) {
                if (sheets[j].getName() === 'שאלון') {
                  const formIdCell = sheets[j].getRange(2, 11); // עמודה K
                  if (formIdCell.getValue()) {
                    FORM_ID = formIdCell.getValue().toString();
                    Logger.log('✅ נמצא Form ID מה-Sheet: ' + FORM_ID);
                    break;
                  }
                }
              }
              break;
            } catch (e) {
              // זה לא Sheet, נסה Form
              try {
                const form = FormApp.openById(sourceId);
                FORM_ID = sourceId;
                Logger.log('✅ נמצא Form ID מהטריגר: ' + sourceId);

                const destinationId = form.getDestinationId();
                if (destinationId) {
                  SHEET_ID = destinationId;
                  Logger.log('✅ נמצא Sheet ID מהפורם: ' + destinationId);
                }
                break;
              } catch (e2) {
                // לא זה ולא זה
              }
            }
          }
        }
      }
    } catch (error) {
      Logger.log('⚠️ לא ניתן למצוא ID מהטריגרים: ' + error.toString());
    }

    // אם לא מצאנו דרך טריגר, נסה למצוא דרך פורמים פתוחים
    if (!SHEET_ID || !FORM_ID) {
      try {
        // נסה למצוא פורמים שפתוחים ובדוק אם יש להם Sheet מחובר
        const forms = FormApp.getForms();
        for (let i = 0; i < forms.length; i++) {
          try {
            const destinationId = forms[i].getDestinationId();
            if (destinationId) {
              SHEET_ID = destinationId;
              FORM_ID = forms[i].getId();
              Logger.log('✅ נמצא Sheet ID מפורם: ' + SHEET_ID);
              Logger.log('✅ נמצא Form ID: ' + FORM_ID);
              break;
            }
          } catch (e) {
            // דלג על פורמים בלי Sheet
          }
        }
      } catch (error) {
        Logger.log('⚠️ לא ניתן למצוא ID מפורמים: ' + error.toString());
      }
    }

    if (!SHEET_ID || !FORM_ID) {
      Logger.log('❌ לא ניתן למצוא Sheet ID או Form ID אוטומטית');
      Logger.log('');
      Logger.log('💡 הסיבות האפשריות:');
      Logger.log('   1. אין טריגר מוגדר - הוסף טריגר: ⚙️ → Triggers → + Add Trigger');
      Logger.log('   2. הפורם לא מחובר ל-Sheet - פתח הפורם → Responses → Link to Sheets');
      Logger.log('   3. לא ניתן לגשת לפרויקט - ודא שיש הרשאות');
      Logger.log('');
      Logger.log('💡 פתרון: ודא שהפורם מחובר ל-Sheet ויש טריגר מוגדר');
      return;
    }

    Logger.log('📋 משתמש ב-Sheet ID: ' + SHEET_ID);

    const ss = SpreadsheetApp.openById(SHEET_ID);
    Logger.log('✅ Sheet נפתח בהצלחה: ' + ss.getName());

    // מצא את הטאב של התשובות
    const sheets = ss.getSheets();
    let responsesSheet = null;
    for (let i = 0; i < sheets.length; i++) {
      const sheetName = sheets[i].getName();
      if (sheetName === 'טופס 1' || sheetName.includes('תשובות') || sheetName === 'Form Responses 1') {
        responsesSheet = sheets[i];
        break;
      }
    }

    if (!responsesSheet) {
      Logger.log('❌ לא נמצא טאב תשובות');
      Logger.log('💡 ודא שהפורם מחובר ל-Sheet (פתח את הפורם → Responses → Link to Sheets)');
      return;
    }

    // מצא את הטאב עם השאלות
    let settingsSheet = null;
    for (let i = 0; i < sheets.length; i++) {
      if (sheets[i].getName() === 'שאלון') {
        settingsSheet = sheets[i];
        break;
      }
    }

    if (!settingsSheet) {
      Logger.log('❌ לא נמצא טאב "שאלון"');
      return;
    }

    // אם עדיין אין Form ID, נסה למצוא אותו מה-Sheet
    if (!FORM_ID) {
      const formIdCell = settingsSheet.getRange(2, 11); // עמודה K
      const foundFormId = formIdCell.getValue();
      if (foundFormId) {
        FORM_ID = foundFormId.toString();
        Logger.log('✅ Form ID נמצא מה-Sheet: ' + FORM_ID);
      }
    }

    if (!FORM_ID) {
      Logger.log('❌ לא נמצא Form ID');
      Logger.log('💡 ודא שרצת את createFormFromText ושנוצר Form ID');
      return;
    }

    Logger.log('✅ Form ID נמצא: ' + FORM_ID);

    const form = FormApp.openById(FORM_ID);

    // חשב ציונים לכל השורות
    const lastRow = responsesSheet.getLastRow();
    if (lastRow <= 1) {
      Logger.log('⚠️ אין תשובות לחשב');
      return;
    }

    Logger.log('📊 מחשב ציונים...');
    Logger.log('   מספר שורות ב-Sheet: ' + (lastRow - 1));

    // קבל את כל התשובות מהפורם
    const formResponses = form.getResponses();
    Logger.log('   מספר תשובות בפורם: ' + formResponses.length);

    if (formResponses.length === 0) {
      Logger.log('⚠️ אין תשובות בפורם');
      Logger.log('💡 הפרס את הפורם קודם (כמו תלמיד)');
      return;
    }

    // חשב ציון לכל תשובה
    // נחשב על בסיס מספר התשובות בפורם, לא מספר השורות ב-Sheet
    for (let i = 0; i < formResponses.length; i++) {
      const row = i + 2; // שורה 2 = תשובה ראשונה, שורה 3 = תשובה שנייה, וכו'

      // ודא שיש שורה ב-Sheet
      if (row <= lastRow || row === lastRow + 1) {
        Logger.log('📝 מחשב ציון לשורה ' + row + ' (תשובה ' + (i + 1) + ')...');
        try {
          calculateScoreForRow(responsesSheet, settingsSheet, row, formResponses[i]);
        } catch (error) {
          Logger.log('⚠️ שגיאה בחישוב ציון לשורה ' + row + ': ' + error.toString());
        }
      } else {
        Logger.log('⚠️ שורה ' + row + ' לא קיימת ב-Sheet (יש רק ' + lastRow + ' שורות)');
      }
    }

    Logger.log('✅ חישוב ציונים הושלם!');
    Logger.log('💡 פתח את ה-Sheet → טאב "טופס 1" → בדוק עמודות "ציון" ו"אחוז"');

  } catch (error) {
    Logger.log('❌ שגיאה בחישוב ידני: ' + error.toString());
    Logger.log('📋 Stack trace: ' + error.stack);
  }
}

/**
 * פונקציה לחישוב ציונים ידנית - אם הטריגר לא עובד
 * שימוש: הרץ את calculateScoreManually() - היא תמצא את ה-ID בעצמה!
 */
function calculateScoreManually() {
  try {
    Logger.log('🔍 מחשב ציונים ידנית...');

    // נסה למצוא את ה-Sheet וה-Form אוטומטית
    let FORM_ID = null;
    let SPREADSHEET_ID = null;

    // נסה למצוא מהטריגרים או מה-Sheet
    try {
      const triggers = ScriptApp.getProjectTriggers();
      for (let i = 0; i < triggers.length; i++) {
        if (triggers[i].getHandlerFunction() === 'onFormSubmit') {
          const sourceId = triggers[i].getSourceId();
          if (sourceId) {
            // זה יכול להיות Sheet ID או Form ID - נבדוק
            try {
              const ss = SpreadsheetApp.openById(sourceId);
              SPREADSHEET_ID = sourceId;
              Logger.log('✅ נמצא Sheet ID מהטריגר: ' + sourceId);

              // נסה למצוא Form ID מה-Sheet
              const sheets = ss.getSheets();
              for (let j = 0; j < sheets.length; j++) {
                if (sheets[j].getName() === 'שאלון') {
                  const formIdCell = sheets[j].getRange(2, 11); // עמודה K
                  if (formIdCell.getValue()) {
                    FORM_ID = formIdCell.getValue().toString();
                    Logger.log('✅ נמצא Form ID מה-Sheet: ' + FORM_ID);
                    break;
                  }
                }
              }
              break;
            } catch (e) {
              // זה לא Sheet, נסה Form
              try {
                const form = FormApp.openById(sourceId);
                FORM_ID = sourceId;
                Logger.log('✅ נמצא Form ID מהטריגר: ' + sourceId);

                // נסה למצוא Sheet ID מהפורם
                const destinationId = form.getDestinationId();
                if (destinationId) {
                  SPREADSHEET_ID = destinationId;
                  Logger.log('✅ נמצא Sheet ID מהפורם: ' + destinationId);
                }
                break;
              } catch (e2) {
                // לא זה ולא זה
              }
            }
          }
        }
      }
    } catch (error) {
      Logger.log('⚠️ לא ניתן למצוא ID אוטומטית: ' + error.toString());
    }

    // אם לא מצאנו, נבקש מהמשתמש
    if (!FORM_ID || !SPREADSHEET_ID) {
      Logger.log('⚠️ לא ניתן למצוא Form ID או Sheet ID אוטומטית');
      Logger.log('');
      Logger.log('💡 פתרון: עדכן את הפונקציה calculateScoreManually() עם ה-ID:');
      Logger.log('');
      Logger.log('📋 איך למצוא את ה-ID:');
      Logger.log('   1. Form ID: פתח את ה-Sheet → טאב "שאלון" → עמודה K (11)');
      Logger.log('   2. Sheet ID: מהכתובת של ה-Sheet: docs.google.com/spreadsheets/d/XXXXXXXXXX/edit');
      Logger.log('      העתק את XXXXXXX (זה ה-Sheet ID)');
      Logger.log('');
      Logger.log('📝 דוגמה לעדכון:');
      Logger.log('   const FORM_ID = \'1abc...xyz\';');
      Logger.log('   const SPREADSHEET_ID = \'1xyz...abc\';');
      Logger.log('');
      Logger.log('   או עדכן את השורות 629-630 בקוד');
      return;
    }

    // עדכן כאן את הערכים אם צריך (אופציונלי):
    // אם לא נמצא אוטומטית, עדכן כאן את ה-ID:
    // const FORM_ID = 'העתק_כאן_את_Form_ID';
    // const SPREADSHEET_ID = 'העתק_כאן_את_Sheet_ID';

    // נסה לזהות אם המשתמש עדכן ID ידנית
    // אם לא נמצא אוטומטית, ננסה לבדוק אם זה Sheet ID או Form ID
    if (!FORM_ID || !SPREADSHEET_ID) {
      // נסה למצוא מה-Sheet ישירות (אם יש Sheet ID)
      try {
        // ננסה לבדוק אם זה Sheet ID או Form ID
        // Sheet ID בדרך כלל קצר יותר, Form ID ארוך יותר
        const testId = '10_CEpT7FriwghO3IcLhgDZKuioZG2xpygr_AYR_2R90'; // ID שסופק
        try {
          const testSheet = SpreadsheetApp.openById(testId);
          SPREADSHEET_ID = testId;
          Logger.log('✅ נמצא Sheet ID: ' + testId);

          // נסה למצוא Form ID מה-Sheet
          const sheets = testSheet.getSheets();
          for (let j = 0; j < sheets.length; j++) {
            if (sheets[j].getName() === 'שאלון') {
              const formIdCell = sheets[j].getRange(2, 11); // עמודה K
              if (formIdCell.getValue()) {
                FORM_ID = formIdCell.getValue().toString();
                Logger.log('✅ נמצא Form ID מה-Sheet: ' + FORM_ID);
                break;
              }
            }
          }
        } catch (e) {
          // זה לא Sheet, נסה Form
          try {
            const testForm = FormApp.openById(testId);
            FORM_ID = testId;
            Logger.log('✅ נמצא Form ID: ' + testId);

            const destinationId = testForm.getDestinationId();
            if (destinationId) {
              SPREADSHEET_ID = destinationId;
              Logger.log('✅ נמצא Sheet ID מהפורם: ' + destinationId);
            }
          } catch (e2) {
            Logger.log('⚠️ ה-ID שסופק לא מזוהה כ-Sheet או Form');
          }
        }
      } catch (error) {
        // התעלם - נמשיך עם מה שיש
      }
    }

    if (!FORM_ID || !SPREADSHEET_ID) {
      Logger.log('❌ לא ניתן למצוא Form ID או Sheet ID');
      Logger.log('💡 עדכן את השורות 719-720 עם ה-ID שלך:');
      Logger.log('   const FORM_ID = \'העתק_כאן_את_Form_ID\';');
      Logger.log('   const SPREADSHEET_ID = \'10_CEpT7FriwghO3IcLhgDZKuioZG2xpygr_AYR_2R90\';');
      return;
    }

    const form = FormApp.openById(FORM_ID);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // מצא את הטאב של התשובות
    const sheets = ss.getSheets();
    let responsesSheet = null;
    for (let i = 0; i < sheets.length; i++) {
      const sheetName = sheets[i].getName();
      if (sheetName === 'טופס 1' || sheetName.includes('תשובות') || sheetName === 'Form Responses 1') {
        responsesSheet = sheets[i];
        break;
      }
    }

    if (!responsesSheet) {
      Logger.log('❌ לא נמצא טאב תשובות');
      return;
    }

    // מצא את הטאב עם השאלות
    let settingsSheet = null;
    for (let i = 0; i < sheets.length; i++) {
      if (sheets[i].getName() === 'שאלון') {
        settingsSheet = sheets[i];
        break;
      }
    }

    if (!settingsSheet) {
      Logger.log('❌ לא נמצא טאב "שאלון"');
      return;
    }

    // חשב ציונים לכל השורות
    const lastRow = responsesSheet.getLastRow();
    if (lastRow <= 1) {
      Logger.log('⚠️ אין תשובות לחשב');
      return;
    }

    Logger.log('📊 מחשב ציונים ל-' + (lastRow - 1) + ' תשובות...');

    // קבל את כל התשובות מהפורם
    const formResponses = form.getResponses();

    // חשב ציון לכל תשובה
    for (let row = 2; row <= lastRow; row++) {
      calculateScoreForRow(responsesSheet, settingsSheet, row, formResponses[row - 2]);
    }

    Logger.log('✅ חישוב ציונים הושלם!');

  } catch (error) {
    Logger.log('❌ שגיאה בחישוב ידני: ' + error.toString());
    Logger.log('📋 Stack trace: ' + error.stack);
  }
}

/**
 * מחשב ציון לשורה ספציפית
 */
function calculateScoreForRow(responsesSheet, settingsSheet, row, formResponse) {
  try {
    if (!formResponse) {
      Logger.log('⚠️ שורה ' + row + ': אין formResponse');
      return;
    }

    // קבל את כל השאלות והתשובות הנכונות
    const lastRow = settingsSheet.getLastRow();
    if (lastRow < 2) {
      Logger.log('⚠️ שורה ' + row + ': אין שאלות ב-Sheet');
      return;
    }

    const questionData = settingsSheet.getRange(2, 1, lastRow - 1, 9).getValues();

    let totalQuestions = 0;
    let correctCount = 0;

    const itemResponses = formResponse.getItemResponses();
    const headers = responsesSheet.getRange(1, 1, 1, responsesSheet.getLastColumn()).getValues()[0];

    // מצא עמודות ציון ואחוז
    let scoreCol = headers.indexOf('ציון') + 1;
    let percentCol = headers.indexOf('אחוז') + 1;

    if (scoreCol === 0 || headers.indexOf('ציון') === -1) {
      scoreCol = responsesSheet.getLastColumn() + 1;
      responsesSheet.getRange(1, scoreCol).setValue('ציון');
      responsesSheet.getRange(1, scoreCol).setFontWeight('bold');
    }

    if (percentCol === 0 || headers.indexOf('אחוז') === -1) {
      percentCol = responsesSheet.getLastColumn() + 1;
      if (percentCol === scoreCol) percentCol++;
      responsesSheet.getRange(1, percentCol).setValue('אחוז');
      responsesSheet.getRange(1, percentCol).setFontWeight('bold');
    }

    const hebrewLetters = ['א', 'ב', 'ג', 'ד', 'ה'];

    // בדוק כל שאלה
    for (let i = 0; i < questionData.length; i++) {
      const questionNum = questionData[i][0];
      const correctAnswer = questionData[i][8];

      if (!correctAnswer) continue;

      totalQuestions++;

      const questionHeader1 = 'שאלה ' + questionNum + ':';
      const questionHeader2 = 'ש' + questionNum + ':';
      let userAnswer = null;

      for (let j = 0; j < itemResponses.length; j++) {
        const itemTitle = itemResponses[j].getItem().getTitle();
        if (itemTitle && (itemTitle.includes(questionHeader1) || itemTitle.includes(questionHeader2))) {
          userAnswer = itemResponses[j].getResponse();
          break;
        }
      }

      if (!userAnswer) continue;

      const answerOptions = [
        questionData[i][3],
        questionData[i][4],
        questionData[i][5],
        questionData[i][6],
        questionData[i][7]
      ].filter(function(opt) { return opt && opt.trim() !== ''; });

      let correctAnswerIndex = -1;

      if (correctAnswer === 'א' || correctAnswer === 'א1' || correctAnswer === 'A1') {
        correctAnswerIndex = 0;
      } else if (correctAnswer === 'ב' || correctAnswer === 'א2' || correctAnswer === 'A2') {
        correctAnswerIndex = 1;
      } else if (correctAnswer === 'ג' || correctAnswer === 'א3' || correctAnswer === 'A3') {
        correctAnswerIndex = 2;
      } else if (correctAnswer === 'ד' || correctAnswer === 'א4' || correctAnswer === 'A4') {
        correctAnswerIndex = 3;
      } else {
        let answerNum = correctAnswer;
        if (answerNum.includes('א')) {
          answerNum = answerNum.replace('א', '');
        } else if (answerNum.includes('A')) {
          answerNum = answerNum.replace('A', '');
        }
        correctAnswerIndex = parseInt(answerNum) - 1;
      }

      if (correctAnswerIndex >= 0 && correctAnswerIndex < answerOptions.length) {
        const correctAnswerText = answerOptions[correctAnswerIndex];
        const correctAnswerWithLetter = hebrewLetters[correctAnswerIndex] + '. ' + correctAnswerText;

        const userAnswerClean = userAnswer.toString().trim();
        const correctAnswerClean = correctAnswerWithLetter.trim();
        const correctAnswerCleanWithoutLetter = correctAnswerText.trim();

        if (userAnswerClean === correctAnswerClean ||
            userAnswerClean === correctAnswerCleanWithoutLetter ||
            userAnswerClean.startsWith(hebrewLetters[correctAnswerIndex] + '.') ||
            userAnswerClean.endsWith(correctAnswerText.trim())) {
          correctCount++;
        }
      }
    }

    const score = correctCount + '/' + totalQuestions;
    const percentage = totalQuestions > 0 ? Math.round((correctCount / totalQuestions) * 100) : 0;

    responsesSheet.getRange(row, scoreCol).setValue(score);
    responsesSheet.getRange(row, percentCol).setValue(percentage + '%');

    if (percentage >= 80) {
      responsesSheet.getRange(row, scoreCol, 1, 2).setBackground('#c6efce');
    } else if (percentage >= 60) {
      responsesSheet.getRange(row, scoreCol, 1, 2).setBackground('#ffeb9c');
    } else {
      responsesSheet.getRange(row, scoreCol, 1, 2).setBackground('#ffc7ce');
    }

    Logger.log('✅ שורה ' + row + ': ' + score + ' (' + percentage + '%)');

  } catch (error) {
    Logger.log('❌ שגיאה בחישוב ציון לשורה ' + row + ': ' + error.toString());
  }
}

/**
 * פונקציה לבדיקת הטריגר - הרץ ידנית כדי לבדוק שהכל מוגדר נכון
 */
function checkTrigger() {
  try {
    Logger.log('🔍 בודק טריגרים...');

    const triggers = ScriptApp.getProjectTriggers();
    Logger.log('📊 נמצאו ' + triggers.length + ' טריגרים בסך הכל');

    let foundFormTrigger = false;
    triggers.forEach(function(trigger) {
      if (trigger.getHandlerFunction() === 'onFormSubmit') {
        foundFormTrigger = true;
        Logger.log('✅ נמצא טריגר:');
        Logger.log('   פונקציה: ' + trigger.getHandlerFunction());
        Logger.log('   מקור אירוע: ' + trigger.getEventType());
        Logger.log('   מקור: ' + trigger.getSource());
      }
    });

    if (!foundFormTrigger) {
      Logger.log('⚠️ לא נמצא טריגר עבור onFormSubmit!');
      Logger.log('💡 פתרון: הרץ את createFormFromText() שוב, או הוסף טריגר ידנית:');
      Logger.log('   1. ⚙️ → Triggers → + Add Trigger');
      Logger.log('   2. Choose function: onFormSubmit');
      Logger.log('   3. Select event source: From form');
      Logger.log('   4. Select event type: On form submit');
    } else {
      Logger.log('✅ הכל תקין! הטריגר מוגדר ונכון');
    }

  } catch (error) {
    Logger.log('❌ שגיאה בבדיקת טריגר: ' + error.toString());
  }
}

/**
 * יוצר Google Form מהשאלות
 */
function createForm(questions) {
  const form = FormApp.create('שאלון בינה מלאכותית בחינוך');
  form.setDescription('שאלון שנוצר אוטומטית - ' + new Date().toLocaleDateString('he-IL'));
  form.setProgressBar(true);

  questions.forEach(function(q) {
    let item;

    if (q.type === 'multiple_choice' && q.answers.length > 0) {
      // שאלה רב-ברירה
      item = form.addMultipleChoiceItem();
      item.setTitle('שאלה ' + q.number + ': ' + q.question);

      // הוסף את האותיות העבריות (א, ב, ג, ד) לפני כל תשובה
      const hebrewLetters = ['א', 'ב', 'ג', 'ד', 'ה'];
      const choices = [];
      q.answers.forEach(function(answer, index) {
        // הוסף אות עברית לפני התשובה: "א. תשובה..."
        const letter = hebrewLetters[index] || String(index + 1);
        const answerWithLetter = letter + '. ' + answer;
        choices.push(item.createChoice(answerWithLetter));
      });
      item.setChoices(choices);
    } else {
      // שאלת טקסט
      item = form.addTextItem();
      item.setTitle('שאלה ' + q.number + ': ' + q.question);
    }

    item.setRequired(false);
  });

  return {
    id: form.getId(),
    url: form.getPublishedUrl()
  };
}

/**
 * פונקציה לקריאת שאלון מקובץ טקסט ב-Google Drive
 * (אופציונלי - אם רוצים לקרוא מקובץ Drive במקום מהקוד)
 */
function createFormFromDriveFile(fileName) {
  try {
    // חפש קובץ טקסט ב-Drive
    const files = DriveApp.getFilesByName(fileName);

    if (!files.hasNext()) {
      throw new Error('קובץ לא נמצא: ' + fileName);
    }

    const file = files.next();
    const text = file.getBlob().getDataAsString();

    // פרסור ויצירה
    const questions = parseQuestionnaire(text);
    const sheet = createSpreadsheet(questions);
    const form = createForm(questions);

    Logger.log('✅ הושלם! Sheet: ' + sheet.url + ', Form: ' + form.url);

    return {
      success: true,
      sheet: sheet,
      form: form
    };

  } catch (error) {
    Logger.log('❌ שגיאה: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

