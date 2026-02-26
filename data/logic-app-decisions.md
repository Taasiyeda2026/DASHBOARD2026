# מסמך החלטות מחייב – מערכת דיווח נוכחות ואישורים חודשיים

גרסה: 1.0  
סטטוס: מחייב (Source of Truth)

---

## 1. מטרת המערכת
המערכת מאפשרת דיווח נוכחות, אישורים חודשיים ונעילה לשכר באמצעות Azure Logic App אחד עם HTTP Trigger ו-Switch לפי action.

---

## 2. ישויות וטבלאות

### 2.1 Attendance
טבלת דיווחי נוכחות יומיים.

שדות עיקריים:
- employeeId
- employeeName
- attendanceDate
- startTime / endTime
- workHours
- activityType
- schoolName
- municipality
- programName
- status
- team
- employmentType
- attachmentsNames
- approvedBy
- approvedDate

---

### 2.2 MonthlyApprovals
טבלת אישורים חודשיים בלבד.
כל רשומה מייצגת: עובד אחד × חודש אחד.

שדות:
- employeeId
- employeeName
- monthKey
- team
- instructorApprovalDate
- managerApprovalDate
- released (boolean)

---

### 2.3 Employees
טבלת עובדים (לצורך אימות בלבד).

---

## 3. תפקידי משתמשים

### עובד
- יוצר / מעדכן / מוחק דיווחים בחודש פתוח בלבד
- שולח אישור חודשי (setinstructorapproval)
- אינו יכול לערוך לאחר נעילה

### מנהל
- מאשר חודש לצוות (setmanagerapproval)
- אינו עורך דיווחים לאחר נעילה

### אדמין
- עורך ומוחק גם לאחר נעילה
- מאפס חודש (resetteammonth)
- מייצא נתונים

---

## 4. חוקים רוחביים

### 4.1 אישור חודשי
- רשומת MonthlyApproval אחת בלבד לעובד לכל monthKey
- אישור מדריך יוצר/מעדכן רשומה
- אישור מנהל קובע released=true ונועל חודש

### 4.2 נעילה חודשית
חודש נעול אם קיימת רשומת MonthlyApproval עם released=true.
לאחר נעילה:
- עובד ומנהל אינם יכולים לעדכן/למחוק
- אדמין יכול

### 4.3 הפרדה מוחלטת
- update/delete אינם מאשרים
- approval אינו משנה Attendance

### 4.4 אחריות UI מול API
- UI: הצגת כפתורים וחסימות
- API: אכיפת נעילה ושלמות נתונים

---

## 5. קייסים

### auth
אימות עובד מול Employees. ללא שינוי נתונים.

### submit
יצירת דיווח Attendance בחודש פתוח בלבד.

### getemployeerecords
שליפת דיווחים לעובד.

### getsummary
סיכום כללי – מנהל/אדמין.

### updaterecord
עדכון דיווח – עובד רק בחודש פתוח, אדמין תמיד.

### deleterecord
מחיקת דיווח – אותם חוקים כמו עדכון.

### uploadfile
העלאת קובץ ושיוכו לרשומה.

### exportdata
ייצוא נתונים – אדמין בלבד.

### getmonthlyapproval
שליפת אישור חודשי לעובד.

### setinstructorapproval
שליחת אישור חודשי ע"י עובד.

### setmanagerapproval
אישור מנהל – נועל חודש.

### getteamapprovalstatus
סטטוס אישורים לצוות.

### resetteammonth
איפוס חודש – אדמין בלבד.

---

## 6. קודי שגיאה
- 200 הצלחה
- 400 קלט שגוי
- 401 אימות נכשל
- 403 פעולה אסורה
- 409 קונפליקט
- 500 שגיאת מערכת

---

## 7. כלל ברזל
אין לקבל החלטות בקוד. כל החלטה חייבת להופיע במסמך זה.
