\# 📄 Lazy Report Generator



Automatically converts an Excel findings table into a formatted Word report (.docx) — no manual formatting needed.



\---



\## 📁 Files in This Folder



| File | Description |

|---|---|

| `lazy.py` | Original script — passes data as command argument (use for small datasets) |

| `lazy2.0.py` | Updated script — passes data via temp JSON file (use for large datasets) |

| `generate\_report.js` | The document builder — \*\*do not delete or move this file\*\* |

| `results.xlsx` | Your input Excel file (edit the data inside will do) |



\---



\## ✅ Requirements



Before running, make sure you have the following installed:



\### 1. Python

Download from https://www.python.org/downloads/



Install required Python library:

```

pip install pandas openpyxl

```



\### 2. Node.js

Download from https://nodejs.org/



\### 3. docx (Node package)

Run this \*\*once\*\* inside your `lazy` folder:

```

npm init -y

npm install docx

```

This creates a `node\_modules` folder — keep it in the same directory.



\---



\## 📊 Excel Format



Your Excel file \*\*must\*\* have exactly these two column headers:



| Misconfiguration | CSTP Justification |

|---|---|



\- Column names are \*\*case-sensitive\*\* — must match exactly

\- You can have as many rows as needed

\- Leave the justification blank if not applicable (it will just be empty in the report)



\---



\## ▶️ How to Run



Open \*\*Command Prompt\*\*, navigate to your `lazy` folder:



```

cd "C:\\Users\\YourName\\Desktop\\lazy"

```



\### Using `lazy.py` (small datasets)

```

py lazy.py results.xlsx report.docx        <----you can name any docx even it is not in ur folder

```



\### Using `lazy2.0.py` (large datasets, more than 200 row - recommended)

```

py lazy2.0.py results.xlsx report.docx

```



Replace `results.xlsx` with your input file name and `report.docx` with your desired output file name.



\---



\## 📝 Output



The generated `report.docx` will contain one block per row in your Excel, formatted as:



\*\*1.1.1 (L1) Ensure 'Enforce password history' is set to '24 or more password(s)'\*\*



| Justification | |

|---|---|

| Explanation | Password is ISO controlled document. |



Each block includes:

\- Bold underlined misconfiguration title

\- Dark navy (`#002060`) Justification header

\- Two-column table with Explanation label and content



\---



\## ❗ Troubleshooting



\*\*`Error: Cannot find module 'docx'`\*\*

→ Run `npm install docx` inside the `lazy` folder



\*\*`Error: Missing columns in Excel`\*\*

→ Check that your Excel headers are exactly `Misconfiguration` and `CSTP Justification`



\*\*`FileNotFoundError: The filename or extension is too long`\*\*

→ Switch to `lazy2.0.py` — this is a Windows limit on command length, the 2.0 version fixes it



\*\*`'node' is not recognized as an internal or external command`\*\*

→ Node.js is not installed or not added to PATH — reinstall from https://nodejs.org/ and restart Command Prompt



\*\*`'py' is not recognized`\*\*

→ Python is not installed — download from https://www.python.org/downloads/



\---



\## 📂 Recommended Folder Structure



```

lazy/

├── lazy.py

├── lazy2.0.py

├── generate\_report.js	    ← v1

├──  format.js	 	    ← v2

├── node\_modules/          ← created by npm install

├── package.json           ← created by npm init

├── results.xlsx           ← your input

├── report.docx            ← your output

└── README.md

```

