# 📋 איך למצוא Sheet ID - מדריך מהיר

## 🎯 מה זה Sheet ID?
Sheet ID הוא מזהה ייחודי של Google Sheet - זה המספר שיש בכתובת של ה-Sheet.

---

## ✅ דרך 1: מהכתובת של ה-Sheet (הכי פשוט)

### שלב 1: פתח את ה-Sheet
1. פתח את ה-**Google Sheet** ב-Google Sheets
2. מסתכל בכתובת הדפדפן (address bar)

### שלב 2: מצא את ה-ID בכתובת
הכתובת תיראה כך:
```
https://docs.google.com/spreadsheets/d/10_CEpT7FriwghO3IcLhgDZKuioZG2xpygr_AYR_2R90/edit
```

**העבר:** `10_CEpT7FriwghO3IcLhgDZKuioZG2xpygr_AYR_2R90`

זה ה-**Sheet ID**!

### שלב 3: העתק את ה-ID
1. בחר את כל הטקסט בין `/d/` ו-`/edit`
2. העתק (Ctrl+C)

---

## ✅ דרך 2: מה-Execution log

### שלב 1: הרץ את `createFormFromText`
1. ב-Apps Script, בחר `createFormFromText`
2. לחץ **Run** (▶)
3. פתח את **Execution log** (יומן ביצוע)

### שלב 2: מצא את הקישור
ב-Execution log, תראה:
```
✅ Google Sheet נוצר: https://docs.google.com/spreadsheets/d/10_CEpT7FriwghO3IcLhgDZKuioZG2xpygr_AYR_2R90/edit
```

**העבר:** `10_CEpT7FriwghO3IcLhgDZKuioZG2xpygr_AYR_2R90`

---

## ✅ דרך 3: מה-Sheet עצמו

### שלב 1: פתח את ה-Sheet
1. פתח את ה-Sheet שנוצר
2. עקור לטאב **"שאלון"**

### שלב 2: מצא את הקישור
1. עמודה **J (10)**: Form URL - זה הקישור ל-Sheet
2. **לחץ על הקישור** - זה ייקח אותך ל-Sheet
3. מסתכל בכתובת הדפדפן - שם תראה את ה-ID

---

## ✅ דרך 4: מ-Google Drive

### שלב 1: פתח Google Drive
1. לך ל-**https://drive.google.com/**
2. חפש את ה-Sheet (השם: "שאלון לטפסים - [תאריך]")

### שלב 2: לחץ ימין על הקובץ
1. **לחץ ימין** על הקובץ
2. בחר **"Get link"** (קבל קישור) או **"Open"** (פתח)
3. אם בחרת "Get link" - העתק את הקישור וחפש את ה-ID
4. אם בחרת "Open" - מסתכל בכתובת הדפדפן וחפש את ה-ID

---

## 🔍 איך לזהות Sheet ID?

### Sheet ID בדרך כלל:
- מתחיל עם מספר או אות
- באורך של כ-44 תווים
- לא כולל רווחים או תווים מיוחדים (חוץ מ-`_`)

### דוגמה:
```
✅ נכון: 10_CEpT7FriwghO3IcLhgDZKuioZG2xpygr_AYR_2R90
❌ שגוי: 10 CEpT7FriwghO3IcLhgDZKuioZG2xpygr AYR 2R90 (עם רווחים)
```

---

## ⚠️ בעיות נפוצות:

### "Document is missing"
**זה אומר שה-ID שגוי או שהקובץ נמחק!**

**פתרון:**
1. ודא שהעתקת נכון את כל ה-ID (בלי רווחים)
2. ודא שהקובץ קיים ב-Google Drive
3. ודא שיש לך הרשאות לקובץ

### "You don't have read access"
**זה אומר שאין לך הרשאות לקובץ!**

**פתרון:**
1. ודא שהקובץ שייך לך או שמישהו שיתף אותו איתך
2. ודא שיש לך הרשאות "Viewer" או יותר
3. נסה לפתוח את הקובץ ידנית ב-Google Sheets

---

## 💡 טיפים:

### 1. שמור את ה-ID
לאחר שמצאת את ה-ID, שמור אותו במקום בטוח (קובץ טקסט או הערה)

### 2. בדוק שהעתקת נכון
העתק את כל ה-ID (בלי רווחים, בלי סימנים נוספים)

### 3. ודא שהקובץ קיים
לפני השימוש, ודא שהקובץ קיים ב-Google Drive

---

## ✅ סיכום מהיר:

1. ✅ פתח את ה-Sheet ב-Google Sheets
2. ✅ מסתכל בכתובת הדפדפן
3. ✅ העתק את הטקסט בין `/d/` ו-`/edit`
4. ✅ זה ה-Sheet ID!

---

**🎉 בהצלחה!**

