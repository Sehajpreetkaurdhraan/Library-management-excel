# Library Management System in Excel

An Excel workbook for running a small library. It records books, members, and each time a book is issued or returned. It works out due dates, late fines and overdue books on its own, and sums everything up on a dashboard with charts.

Everything is done with Excel formulas, data validation and conditional formatting. There are **no macros**, so the file opens safely in any version of Excel.

![Dashboard](docs/dashboard.png)

## Features

- **Books** – each book's title, author, genre, price, rating and number of copies. The workbook keeps count of how many copies are issued and available.
- **Members** – each member's details, the books they hold now, and their total fine
- **Issue / Return log** – pick a member and a book from drop-down lists and type the issue date. The due date, status, days late and fine are filled in for you.
- **Overdue tracking** – overdue rows turn red. A book with no copies left is marked *All Issued*.
- **Search** – find a book by its ID or by part of its title (capital letters don't matter), or look up a member
- **Dashboard** – 8 summary tiles, summaries by genre and by month, the top 5 most borrowed books, and 3 charts
- **Settings** – change the loan period, the fine per day or the report date in one place
- **Input checks** – no duplicate IDs, ratings must be 0–5, prices can't be negative, and a return date can't be before the issue date

## How to use

1. Download `Library_Management_System.xlsx` and open it in Microsoft Excel. Excel 2010 or later works, and so does Excel Online.
2. **Yellow cells** are where you type. **Grey cells** are calculated by formulas.
3. **To add a book or member**, type in the next empty yellow row of the **Books** or **Members** sheet.
4. **To issue a book**, go to the next empty row of **Issue_Log** and enter an Issue ID, choose the Member ID and Book ID, and type the Issue Date.
5. **To return a book**, type the Return Date in that book's row.
6. The **Dashboard** and **Book_Search** sheets update automatically.

The report date is `=TODAY()`, so overdue days and fines update every day. To freeze the numbers, type a fixed date in **Settings → Report date**.

## Sheets

| Sheet | What it does |
|---|---|
| Dashboard | Totals, overdue count, fines, genre and monthly summaries, top 5, charts |
| Books | List of books, with copies issued and available counted automatically |
| Members | List of members, with books held and total fine |
| Issue_Log | Every issue and return, with due date, status, days late, fine and overdue flag |
| Book_Search | Search by Book ID, by part of a title, or by Member ID |
| Settings | Loan period (14 days), fine per day (₹5), report date |
| Guide | Instructions inside the workbook |

## How it works

All the data flows through the **Issue_Log**. You type the facts (who, which book, when), and every other sheet counts and adds up rows from the log:

```
Books + Members + Settings  →  Issue_Log  →  Books (issued / available)
                                             Members (books held / fine)
                                             Dashboard and Book_Search
```

Some of the main formulas, as they appear in row 5:

| Purpose | Formula |
|---|---|
| Member name | `=IFERROR(VLOOKUP(B5,Members!$A$5:$B$60,2,FALSE),"Invalid Member ID")` |
| Due date | `=F5+Settings!$B$5` |
| Status | `=IF(H5="","Issued","Returned")` |
| Days late | `=MAX(0,IF(H5="",Settings!$B$7,H5)-G5)` |
| Fine | `=J5*Settings!$B$6` |
| Copies issued | `=COUNTIFS(Issue_Log!$D$5:$D$300,A5,Issue_Log!$I$5:$I$300,"Issued")` |
| Member's total fine | `=SUMIF(Issue_Log!$B$5:$B$300,A5,Issue_Log!$K$5:$K$300)` |
| Search by part of a title | `=MATCH("*"&C19&"*",Books!$B$5:$B$100,0)` |
| Top 5 most borrowed | `=INDEX(Books!$B$5:$B$100,MATCH(LARGE(Books!$L$5:$L$100,1),Books!$L$5:$L$100,0))` |

**Excel features used:** VLOOKUP, INDEX, MATCH, IF, IFERROR, COUNTIF, COUNTIFS, COUNTA, SUM, SUMIF, MAX, LARGE, TODAY, EDATE and TEXT, plus data validation, conditional formatting, charts, freeze panes and filters.

For a full explanation of every sheet and formula, worked examples, practice exercises and viva questions, see the **[Learning Guide (PDF)](docs/Library_Management_Learning_Guide.pdf)**.

## Project structure

```
library-management-excel/
├── Library_Management_System.xlsx      # the workbook
├── README.md                           # this file
└── docs/
    ├── dashboard.png                   # dashboard screenshot
    └── Library_Management_Learning_Guide.pdf   # detailed guide
```

## Sample data

The member names and transactions are made-up sample data for demonstration. The book titles and authors are real.

## Related project

This is the Excel version of my **Book Management System in C**, which has the same features: add, search, issue and return books, plus file storage.

## Author

Sehajpreet Kaur
