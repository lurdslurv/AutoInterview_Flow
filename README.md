# AutoInterview Flow

> **Automated Interview Invitation Engine, Master Document Generator & WhatsApp Dispatcher**

AutoInterview Flow is a Python and Windows automation utility designed to streamline the mass generation and distribution of formal institution and corporate interview invitation letters. 

It solves notorious Word Mail Merge layout issues—such as awkward address block vertical gaps, unaligned header dates, and line wrapping—while automatically generating both an archived Master Word document and individual candidate PDFs ready for WhatsApp dispatch.

---

## Download & Setup (For Non-Technical Users)

**You do NOT need to install Python, Git, or write any code to use this tool.**

### Step 1: Download the Ready-to-Run App
1. Go to the **[Releases](https://github.com/lurdslurv/AutoInterview_Flow/releases)** section on the right-hand side of this GitHub page (or click [Latest Release](https://github.com/lurdslurv/AutoInterview_Flow/releases/latest)).
2. Under **Assets**, click to download **`AutoInterview-v1.0-Windows.zip`**.

### Step 2: Extract the Folder
1. Locate the downloaded `.zip` file on your computer (usually in your `Downloads` folder).
2. Right-click the `.zip` file and select **"Extract All..."** (or unzip it to your Desktop or Documents).
3. Open the newly extracted folder.

### Step 3: Run the Program
1. Inside the folder, simply double-click **`Run_AutoInterview.bat`**.
2. A friendly window will appear:
   - Type `B` to browse for your candidate CSV file using the familiar Windows file picker, or drag and drop your CSV file right into the window.
   - You can test it immediately using the included `applicants_sample.csv`.
3. Confirm or type your Interview Date, Time, Venue, and Reference Number.
4. The system will automatically build your documents and display a simple menu to view your output!

### Step 4: Access Your Deliverables
Everything generated will be neatly saved inside the **`Output`** folder:
- **`Output/All_Interview_Invitations.docx`**: A single master Word document containing all applicant letters sequentially for printing or signing.
- **`Output/All_Interview_Invitations.pdf`**: The combined master PDF.
- **`Output/PDFs/`**: Individual candidate PDFs named `[Candidate Name] - [Phone Number].pdf`.
- **`WhatsApp_Dispatcher.html`**: Double-click this in your browser to search applicants and click one-touch WhatsApp links to dispatch their invitations!

---

## How to Use Your Own Organization's Letterhead

By default, the application uses the included `sample_template.docx`. To use your institution's official letterhead:
1. Open your official letterhead template in Microsoft Word.
2. Insert these tags wherever you want candidate details to appear:
   - `{REF_NO}` – Reference number (e.g. `IUO/REG/PERS/IM/26`)
   - `{TODAYS_DATE}` – Date of the letter
   - `{CANDIDATE_NAME}` – Candidate's full name
   - `{ADDRESS}` – Formatted multi-line address block
   - `{INTERVIEW_DATE}` – Date of the interview
   - `{INTERVIEW_TIME}` – Time of the interview
   - `{VENUE}` – Hall, room, or location
3. Save the document in the AutoInterview folder as **`template_with_placeholders.docx`**.
4. The next time you run `Run_AutoInterview.bat`, it will automatically use your official template!

---

## CSV Data Format

Prepare your applicant list as a standard Excel CSV file with these column headers:

| Column Name | Description | Example |
| :--- | :--- | :--- |
| `Name` | Candidate full name (will be rendered in bold) | `Edowhoghon Irekpono` |
| `Phone` | WhatsApp contact phone number | `08033467912` |
| `Address` | Multi-line candidate address (comma or pipe separated) | `20 Eghosa Street, GRA, Benin City` |
| `Interview_Date` | *(Optional)* Custom date for this applicant | `Tuesday, 4th August, 2026` |
| `Interview_Time` | *(Optional)* Custom time slot | `10:00 AM` |

*Tip: If `Interview_Date` or `Interview_Time` are left blank in the spreadsheet, the program will simply ask you for the default date and time when you run it.*

---

## For Developers (Running from Python Source)

If you are a developer and prefer to run or modify the Python source code directly:

```bash
# 1. Clone the repository
git clone https://github.com/lurdslurv/AutoInterview_Flow.git
cd AutoInterview_Flow

# 2. Install required Python packages
pip install python-docx docx2pdf

# 3. Run the automation
python auto_interview.py applicants_sample.csv
```

---

## Key Features Under the Hood

- **Combined Master Document**: Compiles every candidate's invitation letter into a single Word document (`Output/All_Interview_Invitations.docx`) with headers, crests, and continuous layout intact for record-keeping and mass printing.
- **Individual Candidate PDFs**: Automatically converts and saves each letter as a high-fidelity PDF (`Output/PDFs/[Candidate Name] - [Phone Number].pdf`) for digital dispatch.
- **Smart Address Block Formatting**: Uses Word XML soft line breaks (`<w:br/>`) and 0pt spacing to ensure clean, professional multi-line address blocks without excessive paragraph gaps.
- **Right-Aligned Header Alignment**: Dynamically computes right-aligned tab stops (6.5") for letter reference numbers and dates, preventing awkward wrapping for long dates (e.g., *"Monday, 14th September, 2026"*).
- **Interactive WhatsApp Dispatcher**: Generates a sleek, responsive HTML dashboard (`WhatsApp_Dispatcher.html`) with pre-composed, one-click `wa.me` WhatsApp chat links and direct local links to each candidate's PDF.

---

## Privacy & Security

- **No Proprietary Letterheads in Public Git**: Official institutional templates (`template_with_placeholders.docx`, `invite_template.docx`) are excluded via `.gitignore`.
- **Candidate PII Protection**: Real candidate data (`*.csv` files, except the mock `applicants_sample.csv`) and generated PDFs in `Output/` are strictly ignored and never committed to version control.

---

## License

MIT License. Developed for automated institutional correspondence and HR interview workflows.
