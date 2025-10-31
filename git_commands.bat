@echo off
chcp 65001 >nul
echo 🔗 חיבור ל-GitHub...
echo.

cd /d "%~dp0"

echo ✅ 1. בודק אם repository קיים...
if exist ".git" (
    echo ✅ Repository כבר קיים
) else (
    echo ✅ מאתחל repository...
    git init
)

echo ✅ 2. מוסיף קבצים...
git add .

echo ✅ 3. יוצר commit...
git commit -m "Initial commit - שאלון עם ציונים אוטומטיים"

echo ✅ 4. מחבר ל-GitHub...
git remote remove origin 2>nul
git remote add origin https://github.com/shaharprod/automatic-test.git

echo ✅ 5. משנה branch ל-main...
git branch -M main

echo.
echo 📤 6. שולח ל-GitHub...
echo ⚠️ אם תתבקש - הזן username: shaharprod
echo ⚠️ Password: השתמש ב-personal access token (לא סיסמה רגילה!)
echo.
git push -u origin main

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ✅ הצלחה! הקבצים עלו ל-GitHub
    echo 🔗 https://github.com/shaharprod/automatic-test
) else (
    echo.
    echo ⚠️ יש בעיה - אולי צריך אישור
    echo 💡 ראה הוראות ב-חיבור_ל_GitHub.md
)

echo.
echo ✅ סיום!
pause

