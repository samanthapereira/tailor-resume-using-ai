# Resume Tailor AI — User Guide

Resume Tailor AI is a Streamlit app that rewrites your resume to better match a job posting. It uses the OpenAI API to improve ATS (Applicant Tracking System) keyword alignment while keeping your experience factual.

---

## What you need

- **Python 3.11+** (3.12 recommended; avoid very new versions if packages fail to install)
- An **OpenAI API key** with billing enabled
- Your **base resume** (file or pasted text)
- The **job description** you are applying for

---

## First-time setup

### 1. Open a terminal in the project folder

```powershell
cd "c:\Users\saman\OneDrive\Desktop\Repo\Resume Customizing\resume-tailor-ai"
```

### 2. Create and activate a virtual environment (recommended)

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

### 3. Install dependencies

```powershell
pip install -r requirements.txt
```

### 4. Configure your API key

Copy the example file and add your real key:

```powershell
copy .env.example .env
```

Edit `.env`:

```env
OPENAI_API_KEY=sk-your-actual-key-here
OPENAI_MODEL=gpt-4o-mini
```

| Variable | Required | Description |
|----------|----------|-------------|
| `OPENAI_API_KEY` | Yes | Your OpenAI API key |
| `OPENAI_MODEL` | No | Model name (default: `gpt-4o-mini`) |

**Never commit `.env` to Git.** It is listed in `.gitignore`.

---

## Start the app

With the virtual environment activated:

```powershell
streamlit run app.py
```

Your browser should open automatically (usually `http://localhost:8501`). If it does not, open that URL manually.

To stop the app, press `Ctrl+C` in the terminal.

---

## How to use the app (step by step)

### Step 1 — Sidebar settings

On the left sidebar:

| Field | Purpose |
|-------|---------|
| **OpenAI model** | Optional override (default from `.env`, e.g. `gpt-4o-mini`) |
| **Company name** | Used in the output filename (e.g. `google`) |
| **Your first name** | Used in the output filename (e.g. `Samantha`) |
| **Boost Skills for ATS** | Adds JD keywords to Skills when your base resume supports them |
| **Second review pass** | Hiring-manager-style review for human tone, brevity, metrics, readability |

Example output filename: `google_resume_Samantha_0516.docx`  
(company + `_resume_` + first name + month/day)

### Step 2 — Add your base resume (left column)

Choose **one or both** of these options:

| Method | How |
|--------|-----|
| **Upload** | Click **Upload resume** and choose `.txt`, `.pdf`, or `.docx` |
| **Paste** | Type or paste your full resume into **Or paste resume** |

**Tips:**

- Plain `.txt` files give the most reliable text extraction.
- PDFs from scanned images may extract poorly; paste the text manually if needed.
- If you both upload and paste, both sources are combined.

Uploaded files are also saved locally in the `resumes/` folder (see [Where files are saved](#where-files-are-saved)).

### Step 3 — Paste the job description (right column)

Copy the full job posting from the company site or job board and paste it into **Paste job posting**.

Include as much detail as possible:

- Role title and summary
- Required and preferred skills
- Responsibilities and qualifications

### Step 4 — Click **Tailor resume**

The app runs up to **three passes** (see `rules.md` for full tailoring rules):

1. **Tailor** — minimal edits only; preserves structure, metrics, leadership tone, and bullet rhythm (does not rewrite the whole resume).
2. **Boost Skills** (optional) — adds missing JD keywords to Skills when your base resume supports them.
3. **Second review** (optional) — checks human tone, verbosity, buzzwords, and quantified achievements; shortens or restores bullets as needed.

Processing may take 30–90 seconds depending on model and which passes are enabled.

When finished, you will see:

- The **tailored resume** in a large text area (preview)
- A path showing where the `.docx` was saved on disk
- A **Download as .docx** button (named like `google_resume_Samantha_0516.docx`)

### Step 5 — Review and download

1. **Read the output carefully.** The app is instructed not to invent experience, but you should verify every fact, date, title, and metric.
2. Open the **Optimization report** expander to see skills added and second-review adjustments.
3. Click **Download as .docx**.
   - **If you uploaded a .docx:** your file is cloned as a template; only changed paragraph text is updated (styles, bullets, fonts, margins, spacing preserved).
   - **If you pasted or used PDF/txt:** layout is auto-generated; **always upload `.docx`** for formatting that matches your original.
4. Fine-tune in Word if needed (margins, one-page fit).

### Run again for another company

You do not need to restart Streamlit. For each new application:

1. Change **Company name** in the sidebar  
2. Paste the **new job description**  
3. Update the base resume if needed (or keep the same one)  
4. Click **Tailor resume** again  
5. Download the new `.docx` (new filename for that company)

---

## What the app changes (and what it does not)

| The app will | The app will not |
|--------------|------------------|
| Make targeted bullet edits for ATS alignment | Rewrite the entire resume from scratch |
| Preserve metrics, structure, and senior tone | Invent jobs, employers, or dates |
| Add Skills keywords when your base resume supports them | Add skills you do not have |
| Export a formatted `.docx` (headers, bullets, spacing) | Sound generic or keyword-stuffed (by design) |
| Run a second review for readability and human tone | Guarantee a specific ATS score |

Always treat the output as a **draft** you approve before submitting.

---

## Where files are saved

| Folder | Contents |
|--------|----------|
| `resumes/` | Copies of uploaded resume files (`YYYYMMDD_HHMMSS_filename`) |
| `outputs/` | Tailored Word files (e.g. `google_resume_Samantha_0516.docx`) |

These folders are gitignored so your personal files are not committed. Only empty `.gitkeep` files are tracked in Git.

---

## Troubleshooting

### `Set OPENAI_API_KEY in .env`

- Ensure `.env` exists in `resume-tailor-ai/` (same folder as `app.py`).
- Ensure the line is `OPENAI_API_KEY=sk-...` with no quotes unless your key contains special characters.
- Restart Streamlit after editing `.env`.

### `OpenAI error: ...` (authentication, quota, rate limit)

- Check your key at [platform.openai.com](https://platform.openai.com).
- Confirm your account has credits or a valid payment method.
- **`insufficient_quota` / 429:** Add billing or credits on your OpenAI account — the app cannot call the API without quota.
- Wait a moment and try again if you hit rate limits.

### `No module named 'pypdf'`

- Run `pip install -r requirements.txt` from the `resume-tailor-ai` folder (needed for PDF uploads).
- Restart Streamlit after installing.

### Company name or first name warnings

- Fill in **Company name** and **Your first name** in the sidebar before clicking **Tailor resume**.

### `Unsupported file type`

- Use `.txt`, `.pdf`, or `.docx` only.

### PDF text is garbled or empty

- The PDF may be image-based. Paste the resume text manually instead of uploading.

### `Provide a resume` or `Paste a job description`

- You must supply at least one resume source and a non-empty job description before clicking **Tailor resume**.

### Pip install fails with `No space left on device`

- Free disk space on your system drive, then retry `pip install -r requirements.txt`.

### App does not update after code changes

- Stop Streamlit (`Ctrl+C`) and run `streamlit run app.py` again.

---

## Privacy and security

- Your resume and job description are sent to **OpenAI’s API** for processing.
- Keep your API key in `.env` only; do not share it or commit it to Git.
- Uploaded and generated files stay on your machine in `resumes/` and `outputs/` unless you share them elsewhere.

---

## Quick reference

```powershell
cd resume-tailor-ai
.\.venv\Scripts\Activate.ps1
streamlit run app.py
```

1. Sidebar: company name + your first name  
2. Add resume (upload or paste)  
3. Paste job description  
4. Click **Tailor resume**  
5. Review → **Download as .docx**
