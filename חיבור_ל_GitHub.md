# 🔗 חיבור ל-GitHub - הוראות קצרות

## 🎯 המטרה:
לחבר את הפרויקט ל-repository ב-GitHub: https://github.com/shaharprod/automatic-test

---

## ✅ שלבים מהירים:

### שלב 1: ודא ש-Git מותקן
1. פתח **PowerShell** או **Command Prompt**
2. הקלד: `git --version`
3. אם רואה מספר גרסה - Git מותקן ✅
4. אם לא - הורד מ: https://git-scm.com/download/win

### שלב 2: פתח את התיקייה בפקודה
1. פתח **PowerShell** או **Command Prompt**
2. הקלד:
   ```bash
   cd "C:\Users\User\Downloads\automate american test"
   ```
3. Enter

### שלב 3: אתחל repository מקומי
```bash
git init
```

### שלב 4: הוסף קבצים
```bash
git add .
```
זה מוסיף את כל הקבצים בתיקייה.

### שלב 5: בצע commit ראשון
```bash
git commit -m "Initial commit - שאלון עם ציונים אוטומטיים"
```

### שלב 6: חבר ל-GitHub
```bash
git remote add origin https://github.com/shaharprod/automatic-test.git
```

### שלב 7: שנה את שם ה-branch הראשי (אם צריך)
```bash
git branch -M main
```

### שלב 8: שלח ל-GitHub
```bash
git push -u origin main
```

**הערה:** ייתכן שתתבקש להתחבר ל-GitHub:
- אם יש לך GitHub CLI - זה יעבוד אוטומטית
- אם לא - תצטרך להשתמש ב-personal access token

---

## 🔐 אם צריך אישור (Authentication):

### אופציה 1: Personal Access Token (מומלץ)
1. לך ל: https://github.com/settings/tokens
2. לחץ **"Generate new token (classic)"**
3. תן שם: "automatic-test"
4. בחר הרשאות: **`repo`** (כל הסימונים)
5. לחץ **"Generate token"**
6. **העתק את ה-token** (תראה רק פעם אחת!)
7. כשאתה מריץ `git push`, השתמש ב-token כסיסמה

### אופציה 2: GitHub CLI
```bash
gh auth login
```
עקוב אחר ההוראות על המסך.

---

## ✅ סיכום - פקודות מהירות:

```bash
cd "C:\Users\User\Downloads\automate american test"
git init
git add .
git commit -m "Initial commit"
git remote add origin https://github.com/shaharprod/automatic-test.git
git branch -M main
git push -u origin main
```

---

## 📋 אם הקבצים כבר קיימים ב-GitHub:

אם ה-repository כבר לא ריק, תצטרך לעשות pull קודם:

```bash
git pull origin main --allow-unrelated-histories
git push -u origin main
```

---

## 💡 עדכון קבצים בעתיד:

לאחר שינויים בקבצים:
```bash
git add .
git commit -m "תיאור השינויים"
git push
```

---

**🎉 בהצלחה!**

