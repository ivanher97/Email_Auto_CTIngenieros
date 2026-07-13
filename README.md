# Automated Email Campaign & Template Manager (Email Auto)

![Python](https://img.shields.io/badge/Python-3.12-blue?style=flat-square&logo=python)
![PyQt6](https://img.shields.io/badge/UI-PyQt6-green?style=flat-square&logo=qt)
![Outlook](https://img.shields.io/badge/Mail-Outlook_COM-0078D4?style=flat-square&logo=microsoftoutlook)
![Office](https://img.shields.io/badge/Templates-DocxTemplate-orange?style=flat-square&logo=microsoftoffice)

## Project Overview
This project implements a desktop **Automated Email Campaign Manager** designed to streamline and automate bulk personalized email deliveries via Outlook. By pairing a robust **PyQt6 GUI Wizard** with **openpyxl** for Excel database validation and **docxtpl/mammoth** for rendering Word template documents into beautiful, context-populated HTML emails, the application ensures high-throughput email workflows with maximum customization and reliability.

---

## Architecture & Workflow

### Phase 1: Database Parsing & Validation (Excel)
* **Data Load:** Secure ingestion of Excel records containing contact details (e.g., `email ct`, `email personal`, `nombre completo`).
* **Header Normalization:** Automated normalization of column headers to prevent casing and special character mismatches.
* **Integrity Validation:** Pre-flight validation logic to verify that mandatory fields (destinations and contact names) are fully present before starting campaigns.

### Phase 2: Template Rendering (Word & HTML)
* **Placeholders Parsing:** Scans `.docx` templates using **docxtpl** (Jinja2 syntax) to extract variables and prompt for data mapping.
* **High-Fidelity Conversion:** Uses **mammoth** to transform rendered Word documents into responsive HTML bodies, ensuring alignment and format integrity.
* **Automatic Formatting:** Localized pre-processing formatting rules (e.g. converting `YYYY-MM-DD` date strings to user-friendly `DD/MM/YYYY` layouts).

### Phase 3: COM Mail Delivery (Outlook Engine)
* **Thread Isolation:** Runs mail generation on a dedicated background thread (`MailSendingWorker` subclassing `QThread`) to keep the desktop interface responsive.
* **Outlook COM Integration:** Dispatches via `win32com` API to create Outlook items, set recipient properties (`To`, `CC`), attach files, and automatically route the communication through the correct SMTP account.
* **Diagnostics & Logs:** Active status tracking, including real-time progress updates, operation logging, and crash-safe global exception handlers.

---

## Technical Stack
* **GUI Framework:** PyQt6 (using the Fusion flat style).
* **Database Handler:** openpyxl.
* **Rendering Engine:** docxtpl (Jinja2) & mammoth.
* **Mail Integration:** win32com (Outlook COM API).
* **Language:** Python 3.12.
