/**
 * Certificate Generator - Express Server
 * 
 * Handles Excel upload, PDF certificate generation, downloading, and emailing.
 */

const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const { PDFDocument, rgb, StandardFonts } = require('pdf-lib');
const nodemailer = require('nodemailer');
const archiver = require('archiver');
const cors = require('cors');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = 3000;

// ── Middleware ────────────────────────────────────────────────────────
app.use(cors());
app.use(express.json());
app.use(express.static(__dirname)); // serve index.html, style.css, script.js

// ── Ensure required directories exist ────────────────────────────────
const dirs = ['uploads', 'certificates', 'templates'];
dirs.forEach(dir => {
    const dirPath = path.join(__dirname, dir);
    if (!fs.existsSync(dirPath)) fs.mkdirSync(dirPath, { recursive: true });
});

// ── Multer config (accept only .xlsx) ────────────────────────────────
const storage = multer.diskStorage({
    destination: (req, file, cb) => cb(null, path.join(__dirname, 'uploads')),
    filename: (req, file, cb) => cb(null, `upload_${Date.now()}.xlsx`)
});

const upload = multer({
    storage,
    fileFilter: (req, file, cb) => {
        const ext = path.extname(file.originalname).toLowerCase();
        if (ext !== '.xlsx') {
            return cb(new Error('Only .xlsx files are allowed.'));
        }
        cb(null, true);
    },
    limits: { fileSize: 10 * 1024 * 1024 } // 10 MB max
});

// ══════════════════════════════════════════════════════════════════════
// ROUTES
// ══════════════════════════════════════════════════════════════════════

/**
 * POST /upload
 * Accepts an Excel file, parses it, and returns JSON rows.
 */
app.post('/upload', (req, res) => {
    upload.single('file')(req, res, (err) => {
        if (err) {
            return res.status(400).json({ error: err.message });
        }
        if (!req.file) {
            return res.status(400).json({ error: 'No file uploaded.' });
        }

        try {
            const workbook = XLSX.readFile(req.file.path);
            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];
            const data = XLSX.utils.sheet_to_json(sheet);

            // Validate required columns
            if (data.length === 0) {
                return res.status(400).json({ error: 'The Excel file is empty.' });
            }

            const requiredCols = ['Name', 'Email', 'Course', 'Date'];
            const headers = Object.keys(data[0]);
            const missing = requiredCols.filter(col =>
                !headers.some(h => h.toLowerCase() === col.toLowerCase())
            );

            if (missing.length > 0) {
                return res.status(400).json({
                    error: `Missing required columns: ${missing.join(', ')}`
                });
            }

            // Normalize column names to title case
            const normalized = data.map(row => {
                const obj = {};
                for (const key of Object.keys(row)) {
                    const lower = key.toLowerCase();
                    if (lower === 'name') obj.Name = row[key];
                    else if (lower === 'email') obj.Email = row[key];
                    else if (lower === 'course') obj.Course = row[key];
                    else if (lower === 'date') obj.Date = row[key];
                    else obj[key] = row[key];
                }
                return obj;
            });

            res.json({ success: true, data: normalized, count: normalized.length });
        } catch (parseErr) {
            res.status(500).json({ error: 'Failed to parse Excel file.' });
        }
    });
});

/**
 * POST /generate
 * Accepts an array of user objects and generates a PDF certificate for each.
 */
app.post('/generate', async (req, res) => {
    try {
        const { users } = req.body;
        if (!users || !Array.isArray(users) || users.length === 0) {
            return res.status(400).json({ error: 'No user data provided.' });
        }

        // Clear previous certificates
        const certDir = path.join(__dirname, 'certificates');
        const existing = fs.readdirSync(certDir);
        existing.forEach(f => fs.unlinkSync(path.join(certDir, f)));

        // Load the template image
        const templatePath = path.join(__dirname, 'templates', 'certificate.png');
        if (!fs.existsSync(templatePath)) {
            return res.status(500).json({ error: 'Certificate template not found.' });
        }
        const templateBytes = fs.readFileSync(templatePath);

        const generated = [];

        for (let i = 0; i < users.length; i++) {
            const user = users[i];
            const { Name, Course, Date: dateStr } = user;

            // Create a new PDF (landscape A4-ish)
            const pdfDoc = await PDFDocument.create();
            const page = pdfDoc.addPage([842, 595]); // A4 landscape

            // Embed the template image as background
            const templateImage = await pdfDoc.embedPng(templateBytes);
            page.drawImage(templateImage, {
                x: 0,
                y: 0,
                width: 842,
                height: 595
            });

            // Load fonts
            const fontBold = await pdfDoc.embedFont(StandardFonts.HelveticaBold);
            const fontRegular = await pdfDoc.embedFont(StandardFonts.Helvetica);
            const fontItalic = await pdfDoc.embedFont(StandardFonts.HelveticaOblique);

            // ── Draw Name (centered, large, dark navy) ──
            const nameSize = 32;
            const nameWidth = fontBold.widthOfTextAtSize(Name || '', nameSize);
            const nameX = (842 - nameWidth) / 2;
            page.drawText(Name || '', {
                x: nameX,
                y: 290,
                size: nameSize,
                font: fontBold,
                color: rgb(0.1, 0.1, 0.35)
            });

            // ── Draw a decorative line under the name ──
            page.drawRectangle({
                x: 271,
                y: 280,
                width: 300,
                height: 1.5,
                color: rgb(0.72, 0.57, 0.2)
            });

            // ── Draw Course (centered) ──
            const courseText = Course || '';
            const courseSize = 20;
            const courseWidth = fontRegular.widthOfTextAtSize(courseText, courseSize);
            const courseX = (842 - courseWidth) / 2;
            page.drawText(courseText, {
                x: courseX,
                y: 240,
                size: courseSize,
                font: fontRegular,
                color: rgb(0.2, 0.2, 0.2)
            });

            // ── Draw Date (centered, smaller) ──
            const dateText = dateStr ? String(dateStr) : '';
            const dateSize = 14;
            const dateWidth = fontItalic.widthOfTextAtSize(dateText, dateSize);
            const dateX = (842 - dateWidth) / 2;
            page.drawText(dateText, {
                x: dateX,
                y: 190,
                size: dateSize,
                font: fontItalic,
                color: rgb(0.35, 0.35, 0.35)
            });

            // Save PDF
            const pdfBytes = await pdfDoc.save();
            const safeName = (Name || 'unknown').replace(/[^a-zA-Z0-9]/g, '_');
            const filename = `certificate_${safeName}.pdf`;
            const filePath = path.join(certDir, filename);
            fs.writeFileSync(filePath, pdfBytes);

            generated.push({ name: Name, filename });
        }

        res.json({ success: true, certificates: generated, count: generated.length });
    } catch (err) {
        console.error('Generate error:', err);
        res.status(500).json({ error: 'Failed to generate certificates: ' + err.message });
    }
});

/**
 * GET /download/:filename
 * Download a single certificate PDF.
 */
app.get('/download/:filename', (req, res) => {
    const filePath = path.join(__dirname, 'certificates', req.params.filename);
    if (!fs.existsSync(filePath)) {
        return res.status(404).json({ error: 'Certificate not found.' });
    }
    res.download(filePath);
});

/**
 * GET /download-all
 * Zip all certificates and stream the archive.
 */
app.get('/download-all', (req, res) => {
    const certDir = path.join(__dirname, 'certificates');
    const files = fs.readdirSync(certDir).filter(f => f.endsWith('.pdf'));

    if (files.length === 0) {
        return res.status(404).json({ error: 'No certificates found.' });
    }

    res.setHeader('Content-Type', 'application/zip');
    res.setHeader('Content-Disposition', 'attachment; filename=certificates.zip');

    const archive = archiver('zip', { zlib: { level: 9 } });
    archive.on('error', err => res.status(500).send({ error: err.message }));
    archive.pipe(res);

    files.forEach(file => {
        archive.file(path.join(certDir, file), { name: file });
    });

    archive.finalize();
});

/**
 * POST /send-emails
 * Send each user their certificate as an email attachment.
 * Requires EMAIL_USER and EMAIL_PASS environment variables.
 */
app.post('/send-emails', async (req, res) => {
    try {
        const { users } = req.body;
        if (!users || !Array.isArray(users) || users.length === 0) {
            return res.status(400).json({ error: 'No user data provided.' });
        }

        const emailUser = process.env.EMAIL_USER;
        const emailPass = process.env.EMAIL_PASS;

        if (!emailUser || !emailPass) {
            return res.status(400).json({
                error: 'Email credentials not configured. Set EMAIL_USER and EMAIL_PASS environment variables.'
            });
        }

        // Create SMTP transporter (Gmail by default)
        const transporter = nodemailer.createTransport({
            service: 'gmail',
            auth: {
                user: emailUser,
                pass: emailPass
            }
        });

        const results = [];
        for (const user of users) {
            const safeName = (user.Name || 'unknown').replace(/[^a-zA-Z0-9]/g, '_');
            const filename = `certificate_${safeName}.pdf`;
            const filePath = path.join(__dirname, 'certificates', filename);

            if (!fs.existsSync(filePath)) {
                results.push({ name: user.Name, email: user.Email, status: 'error', message: 'Certificate not found' });
                continue;
            }

            try {
                await transporter.sendMail({
                    from: `"Certificate Generator" <${emailUser}>`,
                    to: user.Email,
                    subject: 'Your Certificate',
                    text: 'Congratulations! Please find your certificate attached.',
                    html: `
                        <div style="font-family: Arial, sans-serif; padding: 20px;">
                            <h2 style="color: #1a1a5e;">Congratulations, ${user.Name}!</h2>
                            <p>Please find your certificate for <strong>${user.Course}</strong> attached to this email.</p>
                            <p style="color: #888; font-size: 12px;">This email was sent by Certificate Generator.</p>
                        </div>
                    `,
                    attachments: [{
                        filename: filename,
                        path: filePath
                    }]
                });
                results.push({ name: user.Name, email: user.Email, status: 'sent' });
            } catch (mailErr) {
                results.push({ name: user.Name, email: user.Email, status: 'error', message: mailErr.message });
            }
        }

        res.json({ success: true, results });
    } catch (err) {
        console.error('Email error:', err);
        res.status(500).json({ error: 'Failed to send emails: ' + err.message });
    }
});

// ── Start server ─────────────────────────────────────────────────────
app.listen(PORT, () => {
    console.log(`\n  🎓 Certificate Generator is running!`);
    console.log(`  ➜ Open http://localhost:${PORT} in your browser\n`);
});
