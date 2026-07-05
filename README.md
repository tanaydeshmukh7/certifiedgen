# 🎓 CertifiedGen — Certificate Generator

**CertifiedGen** is a modern, high-performance web application designed to generate, download, and deliver PDF certificates in bulk from Excel data. It features a premium dark glassmorphic design and operates seamlessly in two modes:
1. **Static Browser Mode** (Fully client-side, zero backend required) — perfect for hosting on **GitHub Pages**.
2. **Local Server Mode** (Node.js/Express backend) — provides additional capability to automate emailing certificates directly to recipients.

---

## ✨ Features

- **📂 Excel File Upload**: Drag and drop any `.xlsx` spreadsheet. It parses columns for `Name`, `Email`, `Course`, and `Date` instantly.
- **🖼️ Custom Template Support**: Upload your own certificate template image (PNG/JPG/JPEG). If none is uploaded, a default design is loaded.
- **⚡ Client-Side PDF Generation**: Generates standard A4 landscape PDF certificates on-the-fly inside your browser. No server overhead, data stays private.
- **📦 Bulk ZIP Archive**: Compile and package all generated certificates into a single ZIP file for immediate download with one click.
- **📧 Email Automation**: Connect to a local Node.js Express server to email certificates automatically to their respective recipients (requires setup of environment variables).
- **📱 Responsive Glassmorphism UI**: Beautifully optimized for both desktop and mobile screens, styled with custom gradients and smooth micro-animations.

---

## 🚀 Live Demo & Deployment

This application is deployed and hosted on **GitHub Pages**:
🔗 **[CertifiedGen Live Website](https://tanaydeshmukh7.github.io/certifiedgen/)** *(Update with your repository link)*

---

## 📂 Project Structure

```text
certifiedgen/
├── templates/
│   └── certificate.png      # Default certificate template image
├── .gitignore               # Excludes node_modules, temp uploads, and credentials
├── index.html               # Main frontend entry point
├── script.js                # Core frontend logic (parsing, generation, ZIP)
├── style.css                # Custom premium CSS styling
├── server.js                # Node.js Express server script (for Email Automation)
├── package.json             # Backend dependencies and scripts
└── README.md                # Project documentation
```

---

## 🛠️ Excel File Format

The uploaded Excel (`.xlsx`) sheet **must** contain a sheet where the first row acts as headers. The sheet must include the following columns (case-insensitive):

| Column | Description | Example |
| :--- | :--- | :--- |
| **Name** | Recipient's full name | John Doe |
| **Email** | Recipient's email address | john.doe@example.com |
| **Course** | Title of the course or achievement | Certified Generative AI Specialist |
| **Date** | Date of generation/completion | October 24, 2026 |

*Note: You can download the example templates from this repository to test.*

---

## 💻 Local Setup & Running Server Mode

To run the application locally and enable the **Email Automation** feature, follow these steps:

### Prerequisites
- Install [Node.js](https://nodejs.org/) (v16 or higher recommended).

### Installation
1. Clone the repository:
   ```bash
   git clone https://github.com/tanaydeshmukh7/certifiedgen.git
   cd certifiedgen
   ```
2. Install dependencies:
   ```bash
   npm install
   ```
3. Set up environment variables. Create a `.env` file in the root directory:
   ```env
   EMAIL_USER=your-email@gmail.com
   EMAIL_PASS=your-gmail-app-password
   ```
   *Note: If using Gmail, you must generate and use an [App Password](https://support.google.com/accounts/answer/185833).*

4. Start the server:
   ```bash
   npm run start
   ```
5. Open your browser and navigate to `http://localhost:3000`.

---

## 🎨 Technologies Used

- **Frontend**: HTML5, Vanilla CSS3 (Custom Variables, Flexbox, CSS Grid), Vanilla JavaScript.
- **Client-side Libraries**:
  - [SheetJS (XLSX)](https://sheetjs.com/) for spreadsheet parsing.
  - [pdf-lib](https://pdf-lib.js.org/) for programmatic PDF rendering.
  - [JSZip](https://stuk.github.io/jszip/) for local ZIP compilation.
- **Backend (Optional)**: Node.js, Express.js, Multer, Nodemailer, Archiver.

---

## 📄 License

This project is open-source and available under the [MIT License](LICENSE).
