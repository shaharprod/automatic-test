# 🔧 תיקון שגיאת UI - "Cannot call SpreadsheetApp.getUi()"

## ⚠️ הבעיה:
```
Exception: Cannot call SpreadsheetApp.getUi() from this context.
```

## ✅ הפתרון:

### שיטה 1: הרצה מתוך Google Sheet (מומלץ)

1. **פתח Google Sheet חדש** (לא Forms, לא Script Editor)
   - לך ל: https://sheets.google.com
   - לחץ על **"Blank"** ליצירת Sheet חדש

2. **פתח Apps Script**
   - בתוך ה-Sheet: **Extensions → Apps Script**
   - או: https://script.google.com

3. **העתק את הקוד**
   - העתק את `apscript_with_gui.js` ל-Apps Script
   - העתק את `QuestionnaireGUI.html` ל-Apps Script (File → New → HTML file)

4. **שמור את הפרויקט**
   - לחץ Ctrl+S או File → Save
   - תן שם לפרויקט (למשל: "יצירת שאלון")

5. **הרץ את הפונקציה**
   - בתפריט: **Run → showQuestionnaireForm**
   - או: לחץ על תפריט **"📝 יצירת שאלון"** שיופיע בתוך ה-Sheet

---

### שיטה 2: שימוש ישיר בלי ממשק גרפי

אם הממשק הגרפי עדיין לא עובד, השתמש בפונקציה הישירה:

```javascript
createFormDirectly(
  `שאלה 1. מהו העיקרון המרכזי?
א. תשובה א
ב. תשובה ב
ג. תשובה ג
ד. תשובה ד

שאלה 2. שאלה נוספת?
א. תשובה א
ב. תשובה ב
ג. תשובה ג
ד. תשובה ד`,
  "01:ג, 02:א"
);
```

**איך להשתמש:**
1. פתח Apps Script
2. העתק את `apscript_with_gui.js`
3. ערוך את `createFormDirectly()` עם השאלון שלך
4. לחץ Run

---

### שיטה 3: תפריט אוטומטי

אם אתה פותח Google Sheet חדש, התפריט **"📝 יצירת שאלון"** יתווסף אוטומטית!

**למצוא את התפריט:**
- בשורת התפריטים העליונה
- לחץ על **"📝 יצירת שאלון"**
- בחר **"פתח ממשק גרפי"**

---

## ✅ סיכום - מה לעשות עכשיו:

1. ✅ פתח **Google Sheet חדש** (לא Script Editor!)
2. ✅ **Extensions → Apps Script**
3. ✅ העתק את `apscript_with_gui.js` + `QuestionnaireGUI.html`
4. ✅ **Run → showQuestionnaireForm**
5. ✅ או: השתמש ב-**"📝 יצירת שאלון"** מהתפריט

---

## 💡 למה זה קורה?

`SpreadsheetApp.getUi()` עובד רק כאשר:
- ✅ הקוד רץ מתוך Google Sheet (לא Forms)
- ✅ הקוד רץ מ-extensions או מתפריט
- ❌ לא עובד כאשר רץ ישירות מ-Script Editor או מ-trigger

---

**🎉 אחרי התיקון - הכל יעבוד!**

