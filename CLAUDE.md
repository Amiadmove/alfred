# CLAUDE.md — Alfred Onboarding Project

## מה הפרויקט

מערכת onboarding לשותפים של Move (חברת Alfred Travel).
שותפים ממלאים טופס מקוון, הנתונים נשמרים ונשלחים במייל כ-Excel.
יש גם Admin Dashboard ו-CRM לניהול הלקוחות.

---

## מבנה הפרויקט

```
alfred/
├── index.html           # דף בחירת פלטפורמה (HolidayHeroes / Ratecore)
├── holidayheroes.html   # טופס onboarding לשותפי HolidayHeroes (4 שלבים, ~15 דק')
├── ratecore.html        # טופס onboarding לשותפי Ratecore API (4 שלבים, ~20 דק')
├── admin.html           # Submissions Dashboard (עם OTP login)
├── crm.html             # Partner CRM - ניהול לקוחות
├── server.js            # Express backend
├── package.json
├── vercel.json          # הגדרות deployment ל-Vercel
├── submissions/         # JSON + Excel files של טפסים שנשלחו
└── clients/             # JSON files של לקוחות מנוהלים
    ├── alfred.json
    ├── ruefa.json
    └── setur.json
```

---

## Stack טכנולוגי

**Frontend:**
- Vanilla HTML + TailwindCSS (CDN)
- AlpineJS v3.14.1 (CDN) — לוגיקה ו-reactivity
- xlsx.js (CDN) — ייצוא Excel בצד הלקוח
- hCaptcha — הגנה על טפסים

**Backend (server.js):**
- Node.js + Express
- nodemailer — שליחת מייל עם Excel מצורף
- xlsx — בניית קבצי Excel בצד השרת
- @vercel/blob — (מותקן אך לא בשימוש פעיל)

**Deployment:** Vercel

---

## Backend API (server.js)

| Route | Method | תיאור |
|-------|--------|--------|
| `/submit/ratecore` | POST | קבלת טופס Ratecore → שמירה + מייל |
| `/submit/holidayheroes` | POST | קבלת טופס HolidayHeroes → שמירה + מייל |
| `/submit` | POST | backward compat → Ratecore |
| `/admin/submissions` | GET | רשימת כל הטפסים שנשלחו |
| `/admin/clients` | GET | רשימת לקוחות (דורש auth) |
| `/api/admin-clients` | GET | זהה ל-admin/clients |
| `/api/save-client` | POST | עדכון לקוח קיים (דורש auth) |

**אימות Admin:** `x-admin-password` header, לפי env var `ADMIN_PASSWORD`.

---

## משתני סביבה (.env.local)

```
SMTP_HOST=smtp.gmail.com
SMTP_PORT=587
SMTP_USER=...
SMTP_PASS=...
ADMIN_PASSWORD=...
PORT=3000
```

מייל התראות נשלח אל: `amiad@alfredtravel.io`

---

## OTP Login ב-admin.html

**מצב נוכחי:** ה-UI של OTP בנוי מלא (2 שלבים: email → קוד 6 ספרות).
רק כתובות `@wearemove.io` מורשות להיכנס.

**הבעיה:** Resend (ספק מייל שנבחר) חוסם שליחת מיילים לדומיינים שאינם מאומתים בתוכנית החינמית.
- Backend routes ל-OTP (`/api/send-otp`, `/api/verify-otp`) עדיין לא מומשו בserver.js.

**אפשרויות להמשך:**
1. **Nodemailer + Gmail** — כבר מותקן בפרויקט, פשוט להוסיף routes לOTP
2. **Resend בתוכנית בתשלום** — מאפשר שליחה לכל דומיין
3. **דילוג על OTP** — auth פשוטה עם password בלבד (כבר קיים)

---

## לקוחות קיימים

- alfred
- ruefa
- setur

---

## איפה עצרנו

עצרנו בזמן מימוש OTP לדשבורד admin.html.
ה-UI מוכן, ה-backend routes חסרים.
הכיוון המועדף: Nodemailer (כבר קיים) במקום Resend.

---

## הנחיות עבודה

- אל תשנה את ה-design system (צבעים: `#152656`, `#3689FB`, `#F7F9FC`)
- שמור על Vanilla HTML/Alpine — אל תכניס frameworks כמו React/Vue
- הכל צריך לעבוד ב-Vercel (serverless או Node server)
- בדוק תמיד שה-Excel export עובד גם בצד הלקוח וגם בצד השרת
