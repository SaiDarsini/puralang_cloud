# 🚀 PuraLang Cloud

> ℹ️ **Note:** **PuraLang Cloud** serves as a cloud-hosted web extension and API wrapper for the core **PuraLang Engine**, extending its domain-specific language (DSL) execution capabilities into web applications, authentication frameworks, and multi-format document transformation pipelines.

[![Live Demo](https://img.shields.io/badge/Live%20Demo-Render-00e676?style=for-the-badge&logo=render)](https://puralang-cloud.onrender.com)
[![Python](https://img.shields.io/badge/Python-3.11-3776AB?style=for-the-badge&logo=python&logoColor=white)](https://www.python.org/)
[![FastAPI](https://img.shields.io/badge/FastAPI-0.100+-009688?style=for-the-badge&logo=fastapi&logoColor=white)](https://fastapi.tiangolo.com/)
[![Google Gemini](https://img.shields.io/badge/AI-Google%20Gemini-4285F4?style=for-the-badge&logo=googlegemini&logoColor=white)](https://ai.google.dev/)

> An end-to-end full-stack SaaS platform for multi-format data manipulation, document transformation, and OCR analysis powered by FastAPI, natural language DSL compilation via Google Gemini AI, and dynamic document processing engines.

---

## 🔗 Live Application
* **Production Deployment:** [https://puralang-cloud.onrender.com](https://puralang-cloud.onrender.com)
* **API Documentation:** [https://puralang-cloud.onrender.com/docs](https://puralang-cloud.onrender.com/docs)

---

## ✨ Key Features

* 📊 **CSV & Excel Analytics:** Instant row filtering, column remapping, custom sorting, summary metrics calculation, and automatic Matplotlib visualization generation.
* 📝 **In-Place Word (.docx) & PPT (.pptx) Editing:** Targeted text replacements, status updates, and automated summary slide creation.
* 📕 **PDF Processing & Watermarking:** Text/tabular extraction to structured CSV and dynamic PDF watermarking (`CONFIDENTIAL`, `APPROVED`).
* 🖼️ **Image Vision & OCR Scanning:** Vision-based table parsing and printed document conversion into structured JSON outputs using Google Gemini Flash.
* 🔐 **Secure Authentication & Logging:** JWT token-based session management, password hashing (`pbkdf2_sha256` / `bcrypt`), and automated execution tracking saved to SQLite (`puralang.db`).

---

## 🛠️ Architecture & Tech Stack
Frontend UI (HTML5 / CSS3 / JS)            │
│       Landing Page  •  Auth Portal  •  Workspace       │
└───────────────────────────┬─────────────────────────────┘
│
▼
┌─────────────────────────────────────────────────────────┐
│                    FastAPI Backend                      │
│      JWT Auth  •  Multi-Format Engines  •  REST APIs    │
└─────────────┬─────────────────────────────┬─────────────┘
│                             │
▼                             ▼
┌───────────────────────────┐  ┌───────────────────────────┐
│    Google Gemini Flash    │  │     SQLite Database       │
│  NL Prompt ➔ DSL Compiler │  │   Users & Activity Logs   │

* **Backend:** FastAPI, Uvicorn, Python 3.11
* **AI & OCR Pipelines:** Google Generative AI (`gemini-1.5-flash`), Pillow (PIL)
* **Data & Document Engines:** Pandas, OpenPyXL, Python-Docx, Python-PPTX, PyPDF, PDFPlumber, ReportLab, Matplotlib
* **Database & Security:** SQLite3, Passlib, PyJWT, Python-Multipart

---

## 🚀 Local Installation & Setup

### Prerequisites
* Python 3.11+
* Git
* Google Gemini API Key


## 📡 API Endpoints Summary

| Method | Endpoint | Description |
| :--- | :--- | :--- |
| `POST` | `/api/v1/register` | User account creation & database insertion |
| `POST` | `/api/v1/login` | Form authentication & JWT access token issue |
| `POST` | `/api/v1/execute` | Data processing for CSV/Excel datasets |
| `POST` | `/api/v1/process-word` | In-place editing and text replacement for `.docx` |
| `POST` | `/api/v1/process-ppt` | Slide text updates and summary creation for `.pptx` |
| `POST` | `/api/v1/process-pdf` | Watermarking and table extraction for `.pdf` |
| `POST` | `/api/v1/process-image` | Vision OCR table parsing and text extraction |
| `GET` | `/api/v1/admin/logs` | SQLite activity log retrieval |

---

## 🎓 Author

**Sai Darsini Sathuluru**  
*Computer Science and Engineering (Data Science)*
