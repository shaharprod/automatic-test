"""
סקריפט להעלאת שאלון מטקסט ל-Google Sheets
"""
import json
import os
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import pickle

# הגדרות
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 
          'https://www.googleapis.com/auth/forms.body']
SAMPLE_SPREADSHEET_ID = None  # יוגדר לאחר יצירת השיטס
CREDENTIALS_FILE = 'credentials.json'
TOKEN_FILE = 'token.pickle'

def authenticate():
    """אימות עם Google APIs"""
    creds = None
    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE, 'rb') as token:
            creds = pickle.load(token)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(CREDENTIALS_FILE):
                print(f"❌ לא נמצא קובץ {CREDENTIALS_FILE}")
                print("📝 נא להוריד credentials מ-Google Cloud Console:")
                print("   1. לך ל-https://console.cloud.google.com/")
                print("   2. צור פרויקט חדש או בחר קיים")
                print("   3. הפעל Google Sheets API ו-Google Forms API")
                print("   4. צור OAuth 2.0 credentials והורד כ-JSON")
                print(f"   5. שמור את הקובץ כ-{CREDENTIALS_FILE}")
                return None
            flow = InstalledAppFlow.from_client_secrets_file(
                CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        
        with open(TOKEN_FILE, 'wb') as token:
            pickle.dump(creds, token)
    
    return creds

def parse_questionnaire(text_file):
    """
    קורא שאלון מקובץ טקסט
    פורמט צפוי:
    Q1. שאלה ראשונה?
    A1. תשובה 1
    A2. תשובה 2
    
    Q2. שאלה שנייה?
    ...
    """
    questions = []
    
    if not os.path.exists(text_file):
        print(f"❌ קובץ {text_file} לא נמצא")
        return questions
    
    with open(text_file, 'r', encoding='utf-8') as f:
        content = f.read()
    
    lines = content.split('\n')
    current_question = None
    current_answers = []
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        # בדיקה אם זו שאלה (Q1, Q2, וכו')
        if line.startswith('Q') and '.' in line:
            # שמור שאלה קודמת אם יש
            if current_question:
                questions.append({
                    'question': current_question,
                    'answers': current_answers,
                    'type': 'multiple_choice' if current_answers else 'text'
                })
            
            # התחל שאלה חדשה
            parts = line.split('.', 1)
            if len(parts) > 1:
                current_question = parts[1].strip()
                current_answers = []
        
        # בדיקה אם זו תשובה (A1, A2, וכו')
        elif line.startswith('A') and '.' in line:
            parts = line.split('.', 1)
            if len(parts) > 1:
                current_answers.append(parts[1].strip())
    
    # שמור שאלה אחרונה
    if current_question:
        questions.append({
            'question': current_question,
            'answers': current_answers,
            'type': 'multiple_choice' if current_answers else 'text'
        })
    
    return questions

def create_spreadsheet(service, title="שאלון"):
    """יוצר Google Sheet חדש"""
    spreadsheet = {
        'properties': {
            'title': title
        },
        'sheets': [{
            'properties': {
                'title': 'שאלון'
            }
        }]
    }
    
    spreadsheet = service.spreadsheets().create(body=spreadsheet).execute()
    spreadsheet_id = spreadsheet.get('spreadsheetId')
    print(f"✅ נוצר Google Sheet: {spreadsheet_id}")
    print(f"🔗 קישור: https://docs.google.com/spreadsheets/d/{spreadsheet_id}")
    return spreadsheet_id

def upload_to_sheets(service, spreadsheet_id, questions):
    """מעלה שאלות ל-Google Sheets"""
    # הכנת הנתונים
    values = [['מספר', 'שאלה', 'סוג', 'תשובה 1', 'תשובה 2', 'תשובה 3', 'תשובה 4', 'תשובה 5']]
    
    for i, q in enumerate(questions, 1):
        row = [str(i), q['question'], q['type']]
        # הוסף תשובות
        for ans in q['answers']:
            row.append(ans)
        # מלא ריקים עד 5 תשובות
        while len(row) < 8:
            row.append('')
        values.append(row)
    
    body = {
        'values': values
    }
    
    range_name = 'שאלון!A1'
    result = service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=range_name,
        valueInputOption='RAW',
        body=body
    ).execute()
    
    print(f"✅ עודכנו {result.get('updatedCells')} תאים")
    return spreadsheet_id

def main():
    print("🚀 מתחיל תהליך העלאת שאלון ל-Google Sheets...")
    
    # אימות
    creds = authenticate()
    if not creds:
        return
    
    service = build('sheets', 'v4', credentials=creds)
    
    # קרא שאלון מטקסט - בדוק קבצים אפשריים
    text_file = None
    for possible_file in ['apscript.txt', 'example_questionnaire.txt']:
        if os.path.exists(possible_file):
            text_file = possible_file
            break
    
    if not text_file:
        print("❌ לא נמצא קובץ שאלון")
        print("📝 נא ליצור קובץ טקסט בשם apscript.txt או example_questionnaire.txt")
        return
    
    print(f"📄 קורא שאלון מקובץ: {text_file}")
    questions = parse_questionnaire(text_file)
    
    if not questions:
        print("❌ לא נמצאו שאלות בקובץ")
        return
    
    print(f"📋 נמצאו {len(questions)} שאלות")
    
    # צור Google Sheet
    spreadsheet_id = create_spreadsheet(service, "שאלון לטפסים")
    
    # העלה לשרת
    upload_to_sheets(service, spreadsheet_id, questions)
    
    # שמור את ה-ID לקובץ
    with open('spreadsheet_id.txt', 'w') as f:
        f.write(spreadsheet_id)
    
    print(f"\n✅ הושלם! Spreadsheet ID נשמר ב-spreadsheet_id.txt")
    print("📝 כעת תוכל להריץ את ה-Apps Script ליצירת Google Form")

if __name__ == '__main__':
    main()

