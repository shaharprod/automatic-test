/**
 * Apps Script ליצירת Google Form מ-Google Sheets
 *
 * הוראות שימוש:
 * 1. פתח Google Sheets עם השאלון
 * 2. לחץ על Extensions > Apps Script
 * 3. העתק את הקוד הזה
 * 4. לחץ על Run > createFormFromSheet
 */

/**
 * פונקציה ראשית ליצירת Google Form מהשיטס הנוכחי
 */
function createFormFromSheet() {
  try {
    // קבל את השיטס הנוכחי
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();

    // קרא את הנתונים
    const data = sheet.getDataRange().getValues();

    // דלג על שורת הכותרות
    const rows = data.slice(1);

    // צור Google Form חדש
    const form = FormApp.create('שאלון מ-Google Sheets');
    form.setDescription('שאלון שנוצר אוטומטית מ-Google Sheets');

    // עבר על כל שאלה
    rows.forEach(function(row) {
      if (row.length < 3 || !row[1]) return; // דלג על שורות ריקות

      const questionNumber = row[0] || '';
      const questionText = row[1] || '';
      const questionType = row[2] || 'text';

      // הוסף שאלה לפי סוג
      let item;

      if (questionType === 'multiple_choice' || hasAnswers(row)) {
        // שאלה רב-ברירה
        item = form.addMultipleChoiceItem();
        item.setTitle(questionText);

        // הוסף תשובות
        const choices = [];
        for (let i = 3; i < row.length; i++) {
          if (row[i] && row[i].trim()) {
            choices.push(item.createChoice(row[i].trim()));
          }
        }

        if (choices.length > 0) {
          item.setChoices(choices);
        } else {
          // אם אין תשובות, שנה לשאלת טקסט
          form.removeItem(item);
          item = form.addTextItem();
          item.setTitle(questionText);
        }
      } else {
        // שאלת טקסט
        item = form.addTextItem();
        item.setTitle(questionText);
      }

      item.setRequired(false); // ניתן לשנות ל-true אם רוצים חובה
    });

    // קבל את קישור הטופס
    const formUrl = form.getPublishedUrl();
    const formId = form.getId();

    // הדפס קישור בקונסולה
    Logger.log('✅ Google Form נוצר בהצלחה!');
    Logger.log('🔗 קישור הטופס: ' + formUrl);
    Logger.log('📝 Form ID: ' + formId);

    // שמור את הקישור והמזהה בשיטס (עמודה חדשה או גיליון חדש)
    sheet.getRange(1, 9).setValue('Form URL');
    sheet.getRange(1, 10).setValue('Form ID');
    sheet.getRange(2, 9).setValue(formUrl);
    sheet.getRange(2, 10).setValue(formId);

    // הצג הודעה למשתמש
    SpreadsheetApp.getUi().alert(
      '✅ Google Form נוצר בהצלחה!\n\n' +
      '🔗 קישור הטופס נשמר בתא J2\n' +
      '📝 Form ID נשמר בתא K2\n\n' +
      'ניתן גם לפתוח את הטופס ישירות:\n' + formUrl
    );

    return {
      success: true,
      formUrl: formUrl,
      formId: formId
    };

  } catch (error) {
    Logger.log('❌ שגיאה: ' + error.toString());
    SpreadsheetApp.getUi().alert('❌ שגיאה ביצירת הטופס: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * בדוק אם בשורה יש תשובות (עמודות 3 ואילך)
 */
function hasAnswers(row) {
  for (let i = 3; i < row.length; i++) {
    if (row[i] && row[i].trim()) {
      return true;
    }
  }
  return false;
}

/**
 * פונקציה אוטומטית שמופעלת כאשר קובץ חדש נוצר או נתונים משתנים
 * (אופציונלי - ניתן להפעיל ידנית)
 */
function onEdit(e) {
  // ניתן להוסיף כאן לוגיקה לניטור שינויים
  // כרגע לא מופעל אוטומטית
}

/**
 * פונקציה עזר ליצירת טופס עם הגדרות מותאמות אישית
 */
function createFormWithSettings() {
  const settings = {
    title: 'שאלון מותאם אישית',
    description: 'תיאור מותאם',
    collectEmail: false,
    showProgressBar: true,
    shuffleQuestions: false
  };

  const result = createFormFromSheet();

  if (result.success) {
    const form = FormApp.openByUrl(result.formUrl);
    form.setTitle(settings.title);
    form.setDescription(settings.description);
    form.setCollectEmail(settings.collectEmail);
    form.setProgressBar(settings.showProgressBar);
    form.setShuffleQuestions(settings.shuffleQuestions);

    Logger.log('✅ הגדרות מותאמות אישית הוחלו');
  }

  return result;
}

