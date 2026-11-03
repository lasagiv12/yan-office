---
name: meeting-summary
description: Takes what the consultant wrote in a meeting - in the meeting's own document (including the marks in the "בוצע" column), in the first-meeting document of a new client, pasted, attached, or in the Outlook event - and files it straight into the client file (tasks, decisions, principles, open issues, metrics, and for a first meeting also the profile and goals) without asking first, then reports what was updated so he can correct it. The other skills do the same by themselves before using the client file, so asking is optional. Does not write summary documents (that is client-summary). Use for "תעדכן בתיק את הפגישה עם...", "תעביר לתיק את מה שכתבתי בפגישה", "הנה הנקודות מהפגישה", a pasted list of meeting notes, an attached notes/transcript file, or a correction like "T-008 לא נסגרה".
---

# עדכון התיק מפגישה

קרא קודם: `references/base-rules.md`, `references/data-access.md`, `references/client-file-schema.md`, `references/meetings-doc.md`, `references/meeting-filing.md`, `references/approval-flow.md`, `references/hebrew-documents.md`.

יאן כותב נקודות קצרות וחופשיות, בדרך שלו. לא מבקשים ממנו לכתוב אחרת, והוא לא צריך לבקש כלום אחרי הפגישה: ההכנה הבאה, או כל שאלה על הלקוח, מעדכנות את התיק לבד. הסקיל הזה הוא למי שרוצה לעדכן עכשיו, להדביק נקודות או לתקן.

## 1. מאיפה הנקודות
לפי הסדר הזה:
1. **מסמכי הפגישות ב-`02 פגישות/`:**
   - מסמך הפגישה `<YYYY-MM-DD> פגישה – <לקוח>.docx`: עמודת "בוצע" בטבלאות של ①, והמשבצות הצהובות של ②.
   - מסמך הפגישה הראשונה `פגישה ראשונה – <לקוח>.docx`: המשבצת הצהובה של כל כותרת.
   - מסמך ישן `פגישות – <לקוח>.docx`: הפרקים שבו, לפי `meetings-doc.md`.
   - בלי תאריך בבקשה: כל מה שמחכה לתיק, מהישן לחדש.
2. **מודבקות בשיחה או קובץ מצורף.**
3. **קובץ** ב-`07 לתיוק` או ב-`02 פגישות/<תאריך>/הערות ותמלול`.
4. **גוף האירוע ביומן,** אם יש חיבור (ראו `data-access.md`).

אם הנקודות לא הגיעו ממסמך (2 עד 4), ואין מסמך לאותה פגישה, יוצרים אותו לפי `meetings-doc.md` ("פגישה בלי מסמך"), עם הנקודות כמו שהן. כך לכל פגישה יש מסמך.

אם לא ברור לאיזה לקוח הנקודות שייכות, שואלים. אם לא ברור מאיזו פגישה, שואלים רק על התאריך.

## 2. מעדכנים
לפי `meeting-filing.md`, בלי לשאול לפני: סיווג, גיבוי, כתיבה לתיק, סטטוס במסמך, שורה ביומן הפעולות.

## 3. סיכום מסודר: רק כשמבקשים
העדכון לתיק לא כותב מסמך סיכום. כשמבקשים סיכום (של הפגישה, של תקופה או כללי), עוברים לסקיל `client-summary`.

## בסוף
לפי `meeting-filing.md` ("מה אומרים בסוף"): מה עודכן, מה לא ועל מה צריך תשובה, ו"משהו לא נכון? כתוב לי מה לתקן". אחריהם ההתקדמות: "הושלמו N מתוך M (P%), נשארו K, מהן L באיחור" (לפי "התקדמות במשימות" ב-`client-file-schema.md`). אחר כך מציעים, בהתאם למה שעלה:
- **פגישה הבאה:** אם נאמר מועד, לקבוע אותה ביומן (לפי `mail-and-calendar`, אחרי אישור).
- **מייל סיכום ללקוח:** טיוטה ב-Outlook עם ההחלטות והמשימות של הלקוח, בלי פריטים חסויים ובלי משימות של יאן. רק טיוטה. השליחה לפי ההגדרות ובאישור.
- **סיכום מסודר לשליחה:** לפי הסקיל `client-summary`.
