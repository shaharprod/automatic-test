# 📤 הוראות - Commit ו-Push ל-GitHub

## ✅ דרך 1: קובץ אוטומטי (הכי פשוט)

**לחץ כפול על:** `git_commands.bat`

זה יעשה הכל אוטומטית!

---

## ✅ דרך 2: ידנית (אם צריך)

### פתח PowerShell או CMD:
```powershell
cd "C:\Users\User\Downloads\automate american test"
```

### הרץ את הפקודות:

```bash
# 1. אתחל repository (אם עדיין לא)
git init

# 2. הוסף קבצים
git add .

# 3. צור commit
git commit -m "Initial commit - שאלון עם ציונים אוטומטיים"

# 4. חבר ל-GitHub
git remote add origin https://github.com/shaharprod/automatic-test.git

# 5. שנה branch ל-main
git branch -M main

# 6. שלח ל-GitHub
git push -u origin main
```

---

## 🔐 אם צריך אישור:

כשאתה מריץ `git push`, תתבקש:

### Username:
```
shaharprod
```

### Password:
**לא סיסמה רגילה!** צריך Personal Access Token:

1. לך ל: https://github.com/settings/tokens
2. לחץ **"Generate new token (classic)"**
3. תן שם: "automatic-test"
4. בחר: **`repo`** (כל הסימונים)
5. לחץ **"Generate token"**
6. **העתק את ה-token**
7. השתמש בו כ-password (לא סיסמה רגילה!)

---

## ✅ אחרי שהכל עלה:

**פתח:** https://github.com/shaharprod/automatic-test

**תראה:**
- ✅ כל הקבצים
- ✅ README.md
- ✅ הקוד המלא

---

## 💡 עדכון קבצים בעתיד:

לאחר שינויים:
```bash
git add .
git commit -m "תיאור השינויים"
git push
```

---

**🎉 בהצלחה!**

