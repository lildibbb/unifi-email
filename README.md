# NextGen VIP Email Blast

A premium Customer Experience onboarding tool developed for **Telekom Malaysia (TM)** to generate and automate VIP onboarding emails. It features a responsive Flask web application for real-time template editing/previewing and a custom Google Chrome Extension that integrates directly with **Microsoft Outlook Web** to automate the dispatch flow.

---

## 🚀 Key Features

*   **Dual Language Support**: Supports both **English (ENG)** and **Bahasa Melayu (BM)** onboarding templates.
*   **Live Preview**: Real-time rendering of email templates inside the browser with toggles for **Desktop** and **Mobile** viewports.
*   **Modern Shadcn-style UI**: Sleek Telekom Malaysia branded theme with native **Light** and **Dark** modes.
*   **Outlook Web Automation**: A lightweight helper Chrome Extension bridges the portal with Outlook Web (`outlook.office.com` or `outlook.cloud.microsoft`), allowing direct composing, HTML injection, and auto-sending.
*   **Matched-Height Table Layouts**: Custom HTML tables designed for full responsiveness and layout height consistency (no mismatched heights in Outlook Desktop or legacy clients).

---

## 📁 Repository Structure

```text
├── server.py              # Flask server running the web dashboard
├── renderer.py            # Backend engine performing HTML template parsing and replacement
├── static/                # Static assets (stylesheets and screenshots used in user guide)
│   ├── style.css          # Common styling sheet
│   └── preview_screenshot.png
├── templates/             # HTML dashboard files and template codebases
│   ├── index.html         # Main dashboard interface page
│   └── NextGen VIP Onboarding_HTML/   # Source onboarding templates (ENG/BM)
└── extension/             # Google Chrome extension source files
    ├── manifest.json      # Extension metadata configuration
    ├── background.js      # Background script running chrome API event handlers
    ├── bridge.js          # DOM messaging bridge content script
    ├── compose.js         # Script injected into Outlook Web Compose tab to inject HTML
    └── icons/             # Chrome Extension logo icons
```

---

## 💻 Web App Setup & Installation

### 1. Prerequisites
Ensure you have **Python 3.8+** installed on your system.

### 2. Install Dependencies
Install Flask using pip:
```bash
pip install flask
```

### 3. Run the Flask Server
Navigate to the root directory and run the application server:
```bash
python server.py
```
Open **[http://localhost:5000](http://localhost:5000)** in your web browser.

---

## 🔌 Chrome Extension Setup Guide

To automate sending emails directly through Outlook, you need to load the companion extension in Google Chrome.

### How to Load Unpacked (Developer Mode)
1.  Open Google Chrome and navigate to:
    ```text
    chrome://extensions/
    ```
2.  Enable the **Developer mode** toggle switch in the top-right corner.
3.  Click the **Load unpacked** button in the top-left corner.
4.  Navigate to the repository folder on your machine, select the `extension` folder, and click **Select Folder**.
5.  The extension is now active and will communicate with your local dashboard.

---

## 📬 Usage Workflow

1.  **Start the Flask server** and open the web dashboard:
    ![User Guide Modal Screenshot](static/user_guide_modal.png)
2.  **Ensure your Chrome Extension is active**. Click **Connect Outlook** in the header. If prompted, authenticate your Outlook Web account.
3.  **Fill in the dynamic client form fields** (Name, Router, Package, Address, etc.) and preview responsively in real time:
    ![Desktop Preview Screenshot](static/desktop_preview.png)
4.  **Switch views between ENG and BM or Desktop and Mobile** to verify alignment:
    ![Mobile Preview Screenshot](static/mobile_preview.png)
5.  **Click Send Email via Outlook**. The extension will automatically open the Outlook compose window, insert the recipient, subject line, inject the custom HTML template, and prepare the email for dispatch.
