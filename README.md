# AutoInterview Flow

> **Automated Interview Invitation Engine, Master Document Generator & WhatsApp Dispatcher**

AutoInterview Flow is a Python and Windows automation utility designed to streamline the mass generation and distribution of formal institution/corporate interview invitation letters. 

It solves notorious Word Mail Merge layout issues—such as awkward address block vertical gaps, unaligned header dates, and line wrapping—while automatically generating both an archived Master Word document and individual candidate PDFs ready for WhatsApp dispatch.

---

## Key Features

- **Combined Master Document**: Compiles every candidate's invitation letter into a single, cohesive Microsoft Word document (`Output/All_Interview_Invitations.docx`) with headers, crests, and continuous layout intact for record-keeping and mass printing.
- **Individual Candidate PDFs**: Automatically converts and saves each letter as a high-fidelity PDF (`Output/PDFs/[Candidate Name] - [Phone Number].pdf`) for digital dispatch.
- **Smart Address Block Formatting**: Uses Word XML soft line breaks (`<w:br/>`) and 0pt spacing to ensure clean, professional multi-line address blocks without excessive paragraph gaps.
- **Right-Aligned Header Alignment**: Dynamically computes right-aligned tab stops (6.5") for letter reference numbers and dates, preventing awkward wrapping for long dates (e.g., *"Monday, 14th September, 2026"*).
- **Interactive WhatsApp Dispatcher**: Generates a sleek, responsive HTML dashboard (`WhatsApp_Dispatcher.html`) with pre-composed, one-click `wa.me` WhatsApp chat links and direct local links to each candidate's PDF.
- **Zero-Setup Standalone Executable**: Packaged with PyInstaller so non-technical staff can run the automation with a single click—no Python, pip, or dependency installation required.
- **Flexible Batch Processing**: Supports interactive CLI prompts, drag-and-drop CSV execution, or a native Windows file explorer dialog.

---

## Directory Structure

```text
AutoInterview/
│
├── AutoInterview.exe            # Standalone executable (Zero Python setup)
├── Run_AutoInterview.bat        # One-click Windows launcher & menu
├── auto_interview.py            # Core automation source code
├── sample_template.docx         # Generic sample template (Public safe)
├── applicants_sample.csv        # Sample candidate dataset
├── WhatsApp_Dispatcher.html     # Interactive dispatch dashboard
│
├── Output/                      # Generated deliverables (Ignored by Git)
│   ├── All_Interview_Invitations.docx  # Master combined letters
│   ├── All_Interview_Invitations.pdf   # Master PDF
│   └── PDFs/                           # Individual candidate PDFs
│       ├── Jane Doe - 08031234567.pdf
│       └── John Doe - 08029876543.pdf
│
└── .gitignore                   # Protects institutional templates & candidate PII
```

---

## Quick Start

### 1. Using the Standalone Launcher (Recommended)
1. Double-click `Run_AutoInterview.bat`.
2. Select or enter your candidate CSV file (you can press `B` to browse using the Windows file picker, or drag and drop your file into the window).
3. Confirm or customize the Interview Date, Time, Venue, and Reference Number.
4. AutoInterview will generate the master document, individual candidate PDFs, and open the post-run menu.

### 2. Running from Python Source
If you are running from source:
```bash
# Clone the repository
git clone https://github.com/lurdslurv/AutoInterview_Flow.git
cd AutoInterview_Flow

# Install dependencies
pip install python-docx docx2pdf

# Run the automation
python auto_interview.py applicants_sample.csv
```

---

## CSV Data Format

Prepare your applicant list as a standard UTF-8 CSV file with the following columns:

| Column Name | Description | Example |
| :--- | :--- | :--- |
| `Name` | Candidate full name (will be rendered in bold) | `Edowhoghon Irekpono` |
| `Phone` | Contact phone number for WhatsApp | `08033467912` |
| `Address` | Multi-line candidate address (comma or pipe separated) | `20 Eghosa Street, GRA, Benin City` |
| `Interview_Date` | (Optional) Specific interview date for this candidate | `Tuesday, 4th August, 2026` |
| `Interview_Time` | (Optional) Specific interview time slot | `10:00 AM` |

*Note: If `Interview_Date` or `Interview_Time` are left blank in the CSV, the program will prompt you for batch default values during startup.*

---

## Customizing the Word Template

AutoInterview searches for templates in the following order:
1. `template_with_placeholders.docx` *(Highest priority - for your organization's official letterhead)*
2. `invite_template.docx`
3. `sample_template.docx` *(Included default generic template)*

You can design your own `.docx` template using standard curly-bracket tags:
- `{REF_NO}` - Document reference number (e.g., `IUO/REG/PERS/IM/26`)
- `{TODAYS_DATE}` - Date of issuance (e.g., `14th September, 2026`)
- `{CANDIDATE_NAME}` - Applicant name
- `{ADDRESS}` - Formatted multi-line address block
- `{INTERVIEW_DATE}` - Date of the interview
- `{INTERVIEW_TIME}` - Time slot
- `{VENUE}` - Location / hall / room

---

## WhatsApp Dispatcher

After generating the letters, open `WhatsApp_Dispatcher.html` in any web browser. 

The dashboard provides:
- Live search and filter by candidate name or phone number.
- Instant click-to-chat links formatted with international dial codes (`wa.me/234...`).
- Direct file links to preview and attach the candidate's personalized PDF letter.

---

## Privacy & Security

- **No Proprietary Letterheads in Public Git**: Official institutional templates (`template_with_placeholders.docx`, `invite_template.docx`) are excluded via `.gitignore`.
- **Candidate PII Protection**: Real candidate data (`*.csv` files, except the mock `applicants_sample.csv`) and generated PDFs in `Output/` are strictly ignored and never committed to version control.

---

## License

MIT License. Developed for automated institutional correspondence and HR interview workflows.
