# העלאה לגיטאהב

הפרויקט מוכן כתיקייה עצמאית בשם:

```text
sales-statistics-analyzer-github
```

## דרך 1: העלאה דרך האתר של GitHub

1. צור Repository חדש בגיטאהב.
2. בחר שם, למשל:

```text
sales-statistics-analyzer
```

3. אל תוסיף README דרך GitHub, כי כבר יש כאן README.
4. לחץ `Create repository`.
5. לחץ `uploading an existing file`.
6. גרור את כל הקבצים מתוך התיקייה `sales-statistics-analyzer-github`.
7. לחץ `Commit changes`.

## דרך 2: העלאה דרך Git במחשב

מתוך התיקייה `sales-statistics-analyzer-github`:

```sh
git init
git add .
git commit -m "Initial sales statistics analyzer"
git branch -M main
git remote add origin https://github.com/YOUR_USER/YOUR_REPO.git
git push -u origin main
```

החלף את:

```text
YOUR_USER/YOUR_REPO
```

בשם המשתמש והריפוזיטורי שלך.

## חשוב לפני העלאה

- לא להעלות CSV אמיתי של לידים.
- לא להעלות טלפונים אמיתיים.
- לא להעלות מיילים אמיתיים.
- לא להעלות הערות לקוח אמיתיות.
- קובץ `data/sample-leads.csv` הוא אנונימי ומותר להעלאה.

