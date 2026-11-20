# ✉️ Cover Letter Generator

A sleek, AI-powered Reflex web app that generates tailored cover letters from your resume and a job description.

![Python](https://img.shields.io/badge/Python-3.9+-blue?logo=python&logoColor=white)
![Reflex](https://img.shields.io/badge/Reflex-0.11.0-6E56CF?logo=reflex&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-green)

## ✨ Features

- **AI-Powered Generation** — Runs `Qwen/Qwen2.5-1.5B-Instruct` locally on GPU with CUDA acceleration
- **Resume-Grounded** — Only uses real details from your resume; no hallucinated experience
- **~250-Word Letters** — Professional, concise, and tailored to the specific job description
- **Download as .docx** — One-click download with clean formatting
- **Dark Aesthetic UI** — Custom Reflex interface with glassmorphism cards, gradient accents, and responsive layout

## 🚀 Quick Start

### 1. Clone the repo
```bash
git clone https://github.com/<your-username>/Cover-Letter-Generator-.git
cd Cover-Letter-Generator-
```

### 2. Install dependencies
```bash
pip install -r requirements.txt
```

### 3. Run the app
```bash
reflex run
```

The app will open at `http://localhost:3000`.

The first generation downloads the Qwen2.5-1.5B-Instruct model (~3 GB) from Hugging Face for fast GPU-accelerated local inference.

## 📖 How to Use

1. **Enter the Company Name** — e.g. *Google*, *Microsoft*
2. **Paste your Resume** — full text of your resume
3. **Paste the Job Description** — the role you're applying for
4. **Click Generate** — the AI crafts a tailored cover letter
5. **Download** — save as a `.docx` file, ready to send

## 🛠️ Tech Stack

| Component        | Technology                              |
| ---------------- | --------------------------------------- |
| Frontend & Backend | Reflex (React + Granian/FastAPI)      |
| AI Model         | `Qwen/Qwen2.5-1.5B-Instruct`           |
| Model execution  | Local, GPU-accelerated (CUDA only)       |
| Document Export  | `python-docx`                           |

## 📂 Project Structure

```
Cover-Letter-Generator-/
├── assets/              # Static styling and background images
├── cover_letter/        # Reflex application code
│   └── cover_letter.py  # Main app, state, and UI components
├── rxconfig.py          # Reflex configuration
├── requirements.txt     # Python dependencies
└── README.md            # Project documentation
```

## 📝 License

MIT License — feel free to use and modify.
