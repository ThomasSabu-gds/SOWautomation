# SOW Automation – Setup & Usage Guide

This project automates the generation of  SOW (Statement of Work) documents using:
- an Excel input workbook
- a Word template
- user selections from the UI (Yes / No / N/A / Text)

The tool removes drafting notes, processes highlighted placeholders, conditionally deletes
entire Schedules, and outputs a clean, client‑ready SOW.

##  What You Get

- ✔ Removes all **[NOTE TO DRAFT]** instructions  
- ✔ Replaces or deletes content based on **Yes/No/N/A/Textbox** inputs  
- ✔ Deletes full **Schedule A/B/C/…** sections when required  
- ✔ Preserves document formatting (headings, numbering, tables, TOC)  
- ✔ Generates a clean final Word document (no highlights, no draft notes)

---

## 🚀 Quick Start (Clone & Run)

### 1) Clone the repository
```sh
# HTTPS
git clone https://github.com/<your-org>/<your-repo>.git
    
# or SSH (if configured)
# git clone git@github.com:<your-org>/<your-repo>.git

Prerequisites
✔ Windows 10/11
✔ Visual Studio 2022 workloads
✔ .NET 8 and above
✔ OpenXML/Interop conditions

🛠️ Install Using Package Manager Console
----Install-Package DocumentFormat.OpenXml
----Install-Package ClosedXML


SOW_Automation/
│
├── Controllers/          → C# controllers & orchestration
├── Models/               → Data models (UI input, mappings)
├── Services/             → Document automation (OpenXML/Interop), Excel parsing (ClosedXML)
├── Views/                → Razor UI
├── wwwroot/              → CSS/JS (Bootstrap, jQuery)
│
├── SOW_Input/            → 📁 INPUTS used by automation (see below)
│   ├── Templates/        → Word base templates (e.g., FAAS122.docx)
│   ├── Mappings/         → Excel workbooks (clauses, flags, schedules)
│   └── Samples/          → Example files for testing
│
├── .gitignore
└── README.md


    Functionality

        --Upload Excel mapping & Word template (or load defaults from SOW_Input)
        --Map highlighted placeholders in Word to Excel rows/keys
        --Collect user inputs in the UI (Yes/No/N/A/Text/Date/Dropdown)
        --Apply rules
        --Replace placeholders ({{KEY}})
        --Include/remove clauses based on Yes/No/N/A
        --Delete entire Schedule sections (A/B/C/…) as per toggles
        --Strip all [NOTE TO DRAFT] content and highlights
        --Output a final, clean SOW .docx
