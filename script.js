/**
 * Certificate Generator – Frontend Logic
 * Handles file upload, template selection, client-side PDF generation, local ZIP download, and local backend email support.
 */

// ── State ────────────────────────────────────────────────────
let parsedUsers = [];
let templateFile = null;
let templateType = 'png'; // 'png' or 'jpeg'
let generatedCerts = [];

// ── DOM Elements ─────────────────────────────────────────────
// Excel File Upload elements
const dropZone = document.getElementById('drop-zone');
const fileInput = document.getElementById('file-input');
const fileInfo = document.getElementById('file-info');
const fileName = document.getElementById('file-name');
const fileSize = document.getElementById('file-size');
const clearFileBtn = document.getElementById('clear-file-btn');

// Template Image Upload elements
const templateDropZone = document.getElementById('template-drop-zone');
const templateInput = document.getElementById('template-input');
const templateInfo = document.getElementById('template-info');
const templateFileName = document.getElementById('template-file-name');
const templateFileSize = document.getElementById('template-file-size');
const clearTemplateBtn = document.getElementById('clear-template-btn');

const previewSection = document.getElementById('preview-section');
const previewTbody = document.getElementById('preview-tbody');
const rowCountBadge = document.getElementById('row-count-badge');

const actionsSection = document.getElementById('actions-section');
const generateBtn = document.getElementById('generate-btn');
const downloadAllBtn = document.getElementById('download-all-btn');
const sendEmailsBtn = document.getElementById('send-emails-btn');

const certsSection = document.getElementById('certs-section');
const certsGrid = document.getElementById('certs-grid');
const certCountBadge = document.getElementById('cert-count-badge');

const toastContainer = document.getElementById('toast-container');

// ════════════════════════════════════════════════════════════════
// TOAST NOTIFICATIONS
// ════════════════════════════════════════════════════════════════

/**
 * Show a toast notification.
 * @param {string} message - The message to display
 * @param {'success'|'error'|'info'|'warning'} type - Toast type
 * @param {number} duration - Duration in ms (default 4000)
 */
function showToast(message, type = 'info', duration = 4000) {
    const icons = { success: '✅', error: '❌', info: 'ℹ️', warning: '⚠️' };
    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;
    toast.innerHTML = `<span>${icons[type] || ''}</span><span>${message}</span>`;
    toastContainer.appendChild(toast);

    setTimeout(() => {
        toast.classList.add('toast-exit');
        setTimeout(() => toast.remove(), 300);
    }, duration);
}

// ════════════════════════════════════════════════════════════════
// EXCEL FILE UPLOAD (DRAG & DROP + BROWSE)
// ════════════════════════════════════════════════════════════════

// Drag & Drop events for Excel
dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('drag-over');
});

dropZone.addEventListener('dragleave', () => {
    dropZone.classList.remove('drag-over');
});

dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('drag-over');
    const files = e.dataTransfer.files;
    if (files.length > 0) handleFile(files[0]);
});

// Click to browse Excel
dropZone.addEventListener('click', (e) => {
    if (e.target.closest('#browse-btn')) return;
    fileInput.click();
});

fileInput.addEventListener('change', () => {
    if (fileInput.files.length > 0) handleFile(fileInput.files[0]);
});

// Clear Excel file
clearFileBtn.addEventListener('click', () => {
    resetExcelState();
});

// ════════════════════════════════════════════════════════════════
// TEMPLATE IMAGE UPLOAD (DRAG & DROP + BROWSE)
// ════════════════════════════════════════════════════════════════

// Drag & Drop events for Template
templateDropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    templateDropZone.classList.add('drag-over');
});

templateDropZone.addEventListener('dragleave', () => {
    templateDropZone.classList.remove('drag-over');
});

templateDropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    templateDropZone.classList.remove('drag-over');
    const files = e.dataTransfer.files;
    if (files.length > 0) handleTemplateFile(files[0]);
});

// Click to browse Template
templateDropZone.addEventListener('click', (e) => {
    if (e.target.closest('#template-browse-btn')) return;
    templateInput.click();
});

templateInput.addEventListener('change', () => {
    if (templateInput.files.length > 0) handleTemplateFile(templateInput.files[0]);
});

// Clear Template file
clearTemplateBtn.addEventListener('click', () => {
    resetTemplateState();
});

/**
 * Handle template image file upload.
 */
function handleTemplateFile(file) {
    const ext = file.name.split('.').pop().toLowerCase();
    if (ext === 'png') {
        templateType = 'png';
    } else if (ext === 'jpg' || ext === 'jpeg') {
        templateType = 'jpeg';
    } else {
        showToast('Only PNG and JPG/JPEG images are supported as templates.', 'error');
        return;
    }

    if (file.size > 10 * 1024 * 1024) {
        showToast('Template is too large. Maximum size is 10 MB.', 'error');
        return;
    }

    templateFile = file;
    templateFileName.textContent = file.name;
    templateFileSize.textContent = formatBytes(file.size);
    templateInfo.style.display = 'flex';
    templateDropZone.style.display = 'none';
    showToast('Template selected successfully!', 'success');
}

function resetTemplateState() {
    templateFile = null;
    templateType = 'png';
    templateInput.value = '';
    templateInfo.style.display = 'none';
    templateDropZone.style.display = '';
}

/**
 * Fetch and return the template bytes.
 */
async function getTemplateBytes() {
    if (templateFile) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => resolve(new Uint8Array(e.target.result));
            reader.onerror = (e) => reject(new Error('Failed to read template file'));
            reader.readAsArrayBuffer(templateFile);
        });
    } else {
        // Fetch default template relative to the site root
        try {
            const response = await fetch('templates/certificate.png');
            if (!response.ok) {
                throw new Error('Default certificate template could not be loaded.');
            }
            const buffer = await response.arrayBuffer();
            return new Uint8Array(buffer);
        } catch (fetchErr) {
            throw new Error('Default template not found. Please upload a custom template.');
        }
    }
}

// ════════════════════════════════════════════════════════════════
// FILE HANDLING & EXCEL PARSING (CLIENT-SIDE)
// ════════════════════════════════════════════════════════════════

/**
 * Handle selected Excel file: validate and parse.
 */
function handleFile(file) {
    if (!file.name.toLowerCase().endsWith('.xlsx')) {
        showToast('Only .xlsx files are supported.', 'error');
        return;
    }

    if (file.size > 10 * 1024 * 1024) {
        showToast('File is too large. Maximum size is 10 MB.', 'error');
        return;
    }

    fileName.textContent = file.name;
    fileSize.textContent = formatBytes(file.size);
    fileInfo.style.display = 'flex';
    dropZone.style.display = 'none';

    showToast('Parsing Excel file...', 'info', 1500);

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];
            const rawData = XLSX.utils.sheet_to_json(sheet);

            if (rawData.length === 0) {
                showToast('The Excel file is empty.', 'error');
                resetExcelState();
                return;
            }

            const requiredCols = ['Name', 'Email', 'Course', 'Date'];
            const headers = Object.keys(rawData[0]);
            const missing = requiredCols.filter(col =>
                !headers.some(h => h.toLowerCase() === col.toLowerCase())
            );

            if (missing.length > 0) {
                showToast(`Missing required columns: ${missing.join(', ')}`, 'error');
                resetExcelState();
                return;
            }

            // Normalize column names to title case
            parsedUsers = rawData.map(row => {
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

            showToast(`Successfully parsed ${parsedUsers.length} records!`, 'success');
            renderPreview();
            showActions();
        } catch (err) {
            console.error(err);
            showToast('Failed to parse Excel file. Is it corrupted?', 'error');
            resetExcelState();
        }
    };
    reader.onerror = function() {
        showToast('Failed to read Excel file.', 'error');
        resetExcelState();
    };
    reader.readAsArrayBuffer(file);
}

function resetExcelState() {
    parsedUsers = [];
    fileInput.value = '';
    fileInfo.style.display = 'none';
    dropZone.style.display = '';
    previewSection.style.display = 'none';
    actionsSection.style.display = 'none';
    certsSection.style.display = 'none';
    downloadAllBtn.disabled = true;
    sendEmailsBtn.disabled = true;
}

function resetAll() {
    resetExcelState();
    resetTemplateState();
}

// ════════════════════════════════════════════════════════════════
// DATA PREVIEW TABLE
// ════════════════════════════════════════════════════════════════

/**
 * Render the parsed Excel data into the preview table.
 */
function renderPreview() {
    previewTbody.innerHTML = '';
    parsedUsers.forEach((user, idx) => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${idx + 1}</td>
            <td>${escapeHtml(user.Name || '')}</td>
            <td>${escapeHtml(user.Email || '')}</td>
            <td>${escapeHtml(user.Course || '')}</td>
            <td>${escapeHtml(user.Date ? String(user.Date) : '')}</td>
        `;
        previewTbody.appendChild(tr);
    });
    rowCountBadge.textContent = `${parsedUsers.length} records`;
    previewSection.style.display = '';
}

function showActions() {
    actionsSection.style.display = '';
}

// ════════════════════════════════════════════════════════════════
// CERTIFICATE GENERATION (CLIENT-SIDE)
// ════════════════════════════════════════════════════════════════

generateBtn.addEventListener('click', async () => {
    if (parsedUsers.length === 0) {
        showToast('No data to generate certificates.', 'warning');
        return;
    }

    generateBtn.classList.add('loading');
    generateBtn.disabled = true;

    try {
        showToast('Loading template image...', 'info', 1500);
        const templateBytes = await getTemplateBytes();
        
        showToast('Generating PDF certificates...', 'info', 1500);
        
        // Revoke old object URLs to release memory
        generatedCerts.forEach(cert => {
            if (cert.url) URL.revokeObjectURL(cert.url);
        });
        generatedCerts = [];

        // Grab variables from PDFLib
        const { PDFDocument, rgb, StandardFonts } = PDFLib;

        for (let i = 0; i < parsedUsers.length; i++) {
            const user = parsedUsers[i];
            const { Name, Course, Date: dateVal } = user;

            // Format Date values cleanly (handles Excel Serial Dates)
            let dateStr = '';
            if (dateVal) {
                if (typeof dateVal === 'number' && dateVal > 20000) {
                    const dateObj = new Date((dateVal - 25569) * 86400 * 1000);
                    dateStr = dateObj.toLocaleDateString('en-US', {
                        year: 'numeric', month: 'long', day: 'numeric'
                    });
                } else {
                    dateStr = String(dateVal);
                }
            }

            // Create a new PDF (A4 landscape)
            const pdfDoc = await PDFDocument.create();
            const page = pdfDoc.addPage([842, 595]);

            // Embed the template image as background
            let templateImage;
            if (templateType === 'png') {
                templateImage = await pdfDoc.embedPng(templateBytes);
            } else {
                templateImage = await pdfDoc.embedJpg(templateBytes);
            }

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
            const nameText = Name || '';
            const nameSize = 32;
            const nameWidth = fontBold.widthOfTextAtSize(nameText, nameSize);
            const nameX = (842 - nameWidth) / 2;
            page.drawText(nameText, {
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
            const dateSize = 14;
            const dateWidth = fontItalic.widthOfTextAtSize(dateStr, dateSize);
            const dateX = (842 - dateWidth) / 2;
            page.drawText(dateStr, {
                x: dateX,
                y: 190,
                size: dateSize,
                font: fontItalic,
                color: rgb(0.35, 0.35, 0.35)
            });

            // Save PDF
            const pdfBytes = await pdfDoc.save();
            const blob = new Blob([pdfBytes], { type: 'application/pdf' });
            const url = URL.createObjectURL(blob);
            const safeName = (Name || 'unknown').replace(/[^a-zA-Z0-9]/g, '_');
            const filename = `certificate_${safeName}.pdf`;

            generatedCerts.push({
                name: Name || 'Unknown',
                filename,
                blob,
                url
            });
        }

        showToast(`Successfully generated ${generatedCerts.length} certificates!`, 'success');
        renderCertificates();

        // Enable action buttons
        downloadAllBtn.disabled = false;
        sendEmailsBtn.disabled = false;
    } catch (err) {
        console.error(err);
        showToast('Failed to generate certificates: ' + err.message, 'error');
    }

    generateBtn.classList.remove('loading');
    generateBtn.disabled = false;
});

// ════════════════════════════════════════════════════════════════
// CERTIFICATES LIST
// ════════════════════════════════════════════════════════════════

/**
 * Render the list of generated certificates with individual download links.
 */
function renderCertificates() {
    certsGrid.innerHTML = '';
    generatedCerts.forEach(cert => {
        const div = document.createElement('div');
        div.className = 'cert-item';
        div.innerHTML = `
            <div class="cert-icon">📜</div>
            <div class="cert-details">
                <div class="cert-name">${escapeHtml(cert.name)}</div>
                <div class="cert-filename">${escapeHtml(cert.filename)}</div>
            </div>
            <a class="cert-download-btn" href="${cert.url}" download="${cert.filename}">
                ⬇ Download
            </a>
        `;
        certsGrid.appendChild(div);
    });
    certCountBadge.textContent = `${generatedCerts.length} certificates`;
    certsSection.style.display = '';
}

// ════════════════════════════════════════════════════════════════
// DOWNLOAD ALL (ZIP) - CLIENT-SIDE
// ════════════════════════════════════════════════════════════════

downloadAllBtn.addEventListener('click', async () => {
    if (generatedCerts.length === 0) {
        showToast('No certificates to download.', 'warning');
        return;
    }

    downloadAllBtn.classList.add('loading');
    downloadAllBtn.disabled = true;
    showToast('Creating ZIP archive...', 'info', 1500);

    try {
        const zip = new JSZip();
        generatedCerts.forEach(cert => {
            zip.file(cert.filename, cert.blob);
        });

        const zipBlob = await zip.generateAsync({ type: 'blob' });
        const url = URL.createObjectURL(zipBlob);
        
        const a = document.createElement('a');
        a.href = url;
        a.download = 'certificates.zip';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        
        showToast('ZIP archive downloaded successfully!', 'success');
    } catch (err) {
        console.error(err);
        showToast('Failed to create ZIP archive: ' + err.message, 'error');
    }

    downloadAllBtn.classList.remove('loading');
    downloadAllBtn.disabled = false;
});

// ════════════════════════════════════════════════════════════════
// SEND EMAILS (VIA LOCAL SERVER FALLBACK)
// ════════════════════════════════════════════════════════════════

sendEmailsBtn.addEventListener('click', async () => {
    if (parsedUsers.length === 0 || generatedCerts.length === 0) {
        showToast('Generate certificates first.', 'warning');
        return;
    }

    if (!confirm(`Send certificates to ${parsedUsers.length} recipients via email?\n\nNote: This feature is only functional when the local Node.js server is running and configured with SMTP credentials.`)) return;

    sendEmailsBtn.classList.add('loading');
    sendEmailsBtn.disabled = true;
    showToast('Attempting to connect to local email server...', 'info', 3000);

    try {
        const response = await fetch('/send-emails', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ users: parsedUsers })
        });

        const result = response.ok ? await response.json() : null;

        if (response.ok && result && result.success) {
            const sent = result.results.filter(r => r.status === 'sent').length;
            const failed = result.results.filter(r => r.status === 'error').length;
            if (failed > 0) {
                showToast(`Sent ${sent} emails. ${failed} failed. Check server console.`, 'warning', 6000);
            } else {
                showToast(`All ${sent} emails sent successfully!`, 'success');
            }
        } else {
            const errorMsg = result ? result.error : 'Server error';
            showToast(`Email sending failed: ${errorMsg}. Make sure your local server is running and configured with EMAIL_USER/EMAIL_PASS.`, 'error', 6000);
        }
    } catch (err) {
        console.error(err);
        showToast('Email feature is only available when running the local Node.js backend server. See the project README for setup instructions.', 'warning', 6000);
    }

    sendEmailsBtn.classList.remove('loading');
    sendEmailsBtn.disabled = false;
});

// ════════════════════════════════════════════════════════════════
// UTILITIES
// ════════════════════════════════════════════════════════════════

/** Format bytes into a human-readable string */
function formatBytes(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
}

/** Escape HTML entities to prevent XSS */
function escapeHtml(str) {
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
}
