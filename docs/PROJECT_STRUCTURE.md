# מבנה הפרויקט

```text
sales-statistics-analyzer-github/
├─ index.html
├─ README.md
├─ LICENSE
├─ .gitignore
├─ .nojekyll
├─ start_windows.bat
├─ start_mac_linux.sh
├─ assets/
│  ├─ app.js
│  └─ styles.css
├─ data/
│  └─ sample-leads.csv
└─ docs/
   ├─ GITHUB_UPLOAD.md
   └─ PROJECT_STRUCTURE.md
```

## תפקיד כל חלק

- `index.html` - דף המערכת הראשי.
- `assets/app.js` - כל חישובי הניתוח, הסינונים, הדוחות והייצוא.
- `assets/styles.css` - עיצוב הדשבורד, טבלאות, גרפים וכרטיסים.
- `data/sample-leads.csv` - נתוני דוגמה אנונימיים להפעלת ניתוח מהיר.
- `docs/` - מסמכי עזר לגיטאהב ולתחזוקה.

## עצמאות מהמערכת המקורית

התיקייה הזו לא משתמשת בקבצים מתוך מערכת ה-CRM ולא תלויה ב-`app.py`, במסדי נתונים, בתיקיות `templates`, `static` או בסביבה וירטואלית.

אפשר להעתיק אותה כפי שהיא למחשב אחר או לגיטאהב.

