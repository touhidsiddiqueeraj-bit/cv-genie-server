const express = require('express');
const cors = require('cors');
const { Document, Packer, Paragraph, TextRun, AlignmentType, BorderStyle } = require('docx');

const app = express();
app.use(cors());
app.use(express.json({ limit: '5mb' }));

// Puppeteer setup
let puppeteer;
let chromium;

async function initPuppeteer() {
    if (puppeteer) return { puppeteer, chromium };
    try {
        chromium = require('@sparticuz/chromium');
        puppeteer = require('puppeteer-core');
        console.log('✅ Puppeteer + Chromium loaded');
        return { puppeteer, chromium };
    } catch (e) {
        console.error('❌ Failed to load Puppeteer:', e.message);
        return { puppeteer: null, chromium: null };
    }
}

app.get('/', async (req, res) => {
    const { puppeteer: pptr } = await initPuppeteer();
    res.json({ status: 'running', version: '2.2', puppeteer: !!pptr });
});

app.get('/test-pdf', async (req, res) => {
    const { puppeteer: pptr, chromium: chrom } = await initPuppeteer();
    if (!pptr) return res.status(500).json({ error: 'Puppeteer not loaded' });
    
    let browser;
    try {
        browser = await pptr.launch({
            args: [...chrom.args, '--no-sandbox', '--disable-setuid-sandbox'],
            defaultViewport: chrom.defaultViewport,
            executablePath: await chrom.executablePath(),
            headless: true,
        });
        const page = await browser.newPage();
        await page.setContent('<h1 style="color:blue;font-family:Georgia">CV Genie Server ✅</h1><p style="font-size:16px">PDF is working with formatting!</p>');
        const pdf = await page.pdf({ format: 'A4', printBackground: true });
        res.writeHead(200, { 'Content-Type': 'application/pdf', 'Content-Length': pdf.length });
        res.end(pdf);
    } catch (e) {
        res.status(500).json({ error: e.message });
    } finally {
        if (browser) await browser.close();
    }
});

// PDF export
app.post('/export-pdf', async (req, res) => {
    const { html, color = '#2563eb', fontHeading = 'Georgia, serif', fontBody = 'Georgia, serif' } = req.body;
    if (!html) return res.status(400).json({ error: 'No HTML provided' });

    const { puppeteer: pptr, chromium: chrom } = await initPuppeteer();
    if (!pptr) return res.status(500).json({ error: 'Puppeteer not loaded' });

    let browser;
    try {
        browser = await pptr.launch({
            args: [...chrom.args, '--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage'],
            defaultViewport: chrom.defaultViewport,
            executablePath: await chrom.executablePath(),
            headless: true,
        });

        const page = await browser.newPage();
        
        // Build a complete styled HTML document with all CSS resolved
        const fullHTML = buildPDFHTML(html, color, fontHeading, fontBody);
        
        await page.setContent(fullHTML, { waitUntil: 'networkidle0', timeout: 30000 });
        
        // Wait a bit for any fonts/images to settle
        await page.evaluate(() => document.fonts.ready);
        
        const pdf = await page.pdf({
            format: 'A4',
            margin: { top: '0mm', bottom: '0mm', left: '0mm', right: '0mm' },
            printBackground: true,
            preferCSSPageSize: true,
        });

        console.log('PDF size:', pdf.length);
        
        res.writeHead(200, {
            'Content-Type': 'application/pdf',
            'Content-Length': pdf.length,
            'Content-Disposition': 'attachment; filename="cv-genie.pdf"',
            'Cache-Control': 'no-cache'
        });
        res.end(pdf);
    } catch (error) {
        console.error('PDF error:', error.message);
        res.status(500).json({ error: error.message });
    } finally {
        if (browser) await browser.close();
    }
});

// Build a complete HTML document with all styling inlined
function buildPDFHTML(previewHTML, cvColor, fontHeading, fontBody) {
    return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<style>
  /* ============ RESET & BASE ============ */
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { 
    font-family: ${fontBody}; 
    font-size: 9pt; 
    line-height: 1.5; 
    color: #1e293b; 
    background: white; 
    margin: 0; 
    padding: 0; 
    display: flex;
    justify-content: center;
  }
  
  /* ============ CV PREVIEW CONTAINER ============ */
  .cv-preview {
    width: 210mm;
    min-height: 297mm;
    background: white;
    padding: 14mm 12mm;
    font-family: ${fontBody};
    position: relative;
  }
  
  /* ============ COMMON ELEMENTS ============ */
  .cv-preview .cv-photo {
    width: 22mm; height: 22mm; 
    border-radius: 50%; object-fit: cover;
  }
  .cv-preview .cv-section {
    margin-bottom: 3mm;
  }
  .cv-preview .cv-section h2 {
    font-family: ${fontHeading};
    font-size: 9.5pt;
    border-bottom: 1.5px solid #333333;
    padding-bottom: 0.8mm;
    margin-bottom: 1.8mm;
    text-transform: uppercase;
    letter-spacing: 0.06em;
    color: #1a1a1a;
  }
  .cv-preview .cv-item {
    margin-bottom: 2mm;
  }
  .cv-preview .cv-item-header {
    display: flex;
    justify-content: space-between;
    font-weight: bold;
    flex-wrap: wrap;
  }
  .cv-preview .cv-item-sub {
    color: #555555;
    font-size: 8.5pt;
  }
  .cv-preview p {
    margin: 0.5mm 0;
  }
  
  /* ============ CLASSIC TEMPLATE ============ */
  .cv-preview.cv-classic .cv-name {
    font-family: ${fontHeading};
    font-size: 19pt;
    text-align: center;
    margin-bottom: 2mm;
    font-weight: bold;
  }
  .cv-preview.cv-classic .cv-title {
    text-align: center;
    color: #555555;
    font-size: 9.5pt;
    margin-bottom: 2.5mm;
  }
  .cv-preview.cv-classic .cv-contact {
    text-align: center;
    font-size: 8pt;
    color: #666666;
    margin-bottom: 3.5mm;
    border-top: 1px solid #dddddd;
    padding-top: 1.5mm;
  }
  
  /* ============ MODERN TEMPLATE ============ */
  .cv-preview.cv-modern .cv-header-modern {
    background: ${cvColor};
    color: white;
    padding: 8mm 12mm;
    margin: -14mm -12mm 4mm -12mm;
  }
  .cv-preview.cv-modern .cv-name {
    font-family: ${fontHeading};
    font-size: 20pt;
    font-weight: 300;
    color: white;
  }
  .cv-preview.cv-modern .cv-title {
    font-size: 9pt;
    opacity: 0.9;
    color: white;
  }
  .cv-preview.cv-modern .cv-body {
    display: flex;
    gap: 4mm;
  }
  .cv-preview.cv-modern .cv-sidebar {
    width: 35%;
  }
  .cv-preview.cv-modern .cv-main {
    width: 65%;
  }
  .cv-preview.cv-modern .cv-section h2 {
    color: ${cvColor};
    font-size: 8.5pt;
    border-bottom-color: ${cvColor};
  }
  
  /* ============ TABULAR TEMPLATE ============ */
  .cv-preview.cv-tabular .cv-name {
    font-family: ${fontHeading};
    font-size: 17pt;
    text-align: center;
  }
  .cv-preview.cv-tabular table {
    width: 100%;
    border-collapse: collapse;
    margin-bottom: 2.5mm;
  }
  .cv-preview.cv-tabular th {
    background: #2c3e50;
    color: white;
    padding: 1.2mm 1.5mm;
    font-size: 7.5pt;
    text-transform: uppercase;
  }
  .cv-preview.cv-tabular td {
    padding: 1.2mm 1.5mm;
    border-bottom: 1px solid #dddddd;
    font-size: 8pt;
  }
  
  /* ============ MINIMAL TEMPLATE ============ */
  .cv-preview.cv-minimal .cv-name {
    font-family: ${fontHeading};
    font-size: 15pt;
    font-weight: 300;
    text-transform: uppercase;
    letter-spacing: 0.05em;
  }
  .cv-preview.cv-minimal .cv-section h2 {
    font-size: 8pt;
    font-weight: 700;
    letter-spacing: 0.1em;
    color: #666666;
  }
  
  /* ============ COLORFUL TEMPLATE ============ */
  .cv-preview.cv-colorful .cv-header-colorful {
    background: linear-gradient(135deg, ${cvColor}, #a855f7);
    color: white;
    padding: 8mm 12mm;
    margin: -14mm -12mm 4mm -12mm;
    border-radius: 2.5mm;
  }
  .cv-preview.cv-colorful .cv-name {
    font-family: ${fontHeading};
    font-size: 20pt;
    font-weight: 700;
    color: white;
  }
  .cv-preview.cv-colorful .cv-body {
    display: flex;
    gap: 4mm;
  }
  .cv-preview.cv-colorful .cv-sidebar-c { width: 35%; }
  .cv-preview.cv-colorful .cv-main-c { width: 65%; }
  .cv-preview.cv-colorful .cv-section h2 { color: ${cvColor}; }
  .cv-preview.cv-colorful .skill-bar {
    height: 4px;
    background: #e5e7eb;
    border-radius: 2px;
    margin: 1mm 0;
    overflow: hidden;
  }
  .cv-preview.cv-colorful .skill-bar-fill {
    height: 100%;
    background: ${cvColor};
    border-radius: 2px;
  }
  
  /* ============ EXECUTIVE TEMPLATE ============ */
  .cv-preview.cv-executive .cv-name {
    font-family: ${fontHeading};
    font-size: 18pt;
    font-variant: small-caps;
    letter-spacing: 0.07em;
  }
  
  /* ============ TIMELINE TEMPLATE ============ */
  .cv-preview.cv-timeline .cv-name {
    font-family: ${fontHeading};
    font-size: 19pt;
    text-align: center;
    font-weight: 700;
  }
  .cv-preview.cv-timeline .timeline-line {
    position: relative;
    padding-left: 11mm;
  }
  .cv-preview.cv-timeline .timeline-line::before {
    content: '';
    position: absolute;
    left: 4.5mm;
    top: 0;
    bottom: 0;
    width: 1.5px;
    background: ${cvColor};
  }
  .cv-preview.cv-timeline .timeline-item {
    position: relative;
    margin-bottom: 2.5mm;
  }
  .cv-preview.cv-timeline .timeline-item::before {
    content: '';
    position: absolute;
    left: -7mm;
    top: 1.5mm;
    width: 5px;
    height: 5px;
    border-radius: 50%;
    background: ${cvColor};
    border: 2px solid white;
  }
  
  /* ============ NEWSPAPER TEMPLATE ============ */
  .cv-preview.cv-newspaper {
    column-count: 2;
    column-gap: 4mm;
  }
  .cv-preview.cv-newspaper .cv-name {
    font-family: ${fontHeading};
    font-size: 17pt;
    text-align: center;
    column-span: all;
  }
  
  /* ============ INFOGRAPHIC TEMPLATE ============ */
  .cv-preview.cv-infographic .cv-name {
    font-family: ${fontHeading};
    font-size: 20pt;
    font-weight: 800;
  }
  .cv-preview.cv-infographic .info-card {
    background: #f8fafc;
    border-left: 3px solid ${cvColor};
    padding: 1.8mm;
    margin-bottom: 1.8mm;
  }
  .cv-preview.cv-infographic .stat-tag {
    display: inline-block;
    background: ${cvColor};
    color: white;
    padding: 0.8mm 1.5mm;
    border-radius: 7px;
    font-size: 7pt;
    margin: 0.8mm;
  }
  
  /* ============ SIDEBAR BOLD TEMPLATE ============ */
  .cv-preview.cv-sidebar-bold {
    display: flex;
    padding: 0;
  }
  .cv-preview.cv-sidebar-bold .cv-sidebar-left {
    width: 35%;
    background: ${cvColor};
    color: white;
    padding: 14mm 4mm 14mm 12mm;
  }
  .cv-preview.cv-sidebar-bold .cv-sidebar-left h2 {
    border-bottom: 1px solid rgba(255,255,255,0.3);
    padding-bottom: 1mm;
    margin-bottom: 2mm;
    font-size: 8.5pt;
    color: white;
  }
  .cv-preview.cv-sidebar-bold .cv-main-right {
    width: 65%;
    padding: 14mm 12mm 14mm 4mm;
  }
  .cv-preview.cv-sidebar-bold .cv-main-right h2 {
    color: ${cvColor};
    font-size: 9pt;
  }
  
  /* ============ CORPORATE TEMPLATE ============ */
  .cv-preview.cv-corporate {
    border: 2px solid #1a3a5c;
    padding: 8mm;
  }
  .cv-preview.cv-corporate .cv-section h2 {
    background: #1a3a5c;
    color: white;
    padding: 0.8mm 1.5mm;
    font-size: 8.5pt;
    text-transform: uppercase;
  }
  
  /* ============ CREATIVE TEMPLATE ============ */
  .cv-preview.cv-creative .cv-name {
    font-family: ${fontHeading};
    font-size: 22pt;
    font-weight: 800;
    text-align: center;
  }
  .cv-preview.cv-creative .tag {
    display: inline-block;
    background: ${cvColor};
    color: white;
    padding: 0.8mm 1.5mm;
    border-radius: 7px;
    font-size: 6.5pt;
    margin: 0.8mm;
  }
  
  /* ============ MAGAZINE TEMPLATE ============ */
  .cv-preview.cv-magazine .cv-name {
    font-family: ${fontHeading};
    font-size: 22pt;
    text-align: center;
    font-weight: 900;
  }

  /* ============ PRINT OPTIMIZATION ============ */
  @media print {
    body { background: white; margin: 0; padding: 0; }
    .cv-preview { box-shadow: none; margin: 0; }
  }
  
  @page {
    size: A4;
    margin: 0;
  }
</style>
</head>
<body>
  ${previewHTML}
</body>
</html>`;
}

// DOCX export (unchanged — already working)
app.post('/export-docx', async (req, res) => {
    const data = req.body;
    if (!data) return res.status(400).json({ error: 'No data provided' });

    try {
        const children = [];

        children.push(new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 120 },
            children: [new TextRun({ text: data.name || 'Your Name', bold: true, size: 36, font: 'Georgia' })]
        }));

        if (data.title) {
            children.push(new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 40 },
                children: [new TextRun({ text: data.title, size: 24, color: '555555', font: 'Georgia' })]
            }));
        }

        if (data.contact) {
            children.push(new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 200 },
                border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC', space: 8 } },
                children: [new TextRun({ text: data.contact, size: 18, color: '666666', font: 'Georgia' })]
            }));
        }

        if (data.summary && data.summary.trim()) {
            children.push(new Paragraph({
                spacing: { before: 100, after: 60 },
                border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '333333', space: 4 } },
                children: [new TextRun({ text: 'PROFESSIONAL SUMMARY', bold: true, size: 24, font: 'Georgia' })]
            }));
            children.push(new Paragraph({
                spacing: { after: 160 },
                children: [new TextRun({ text: data.summary, size: 22, font: 'Georgia' })]
            }));
        }

        if (data.experiences && data.experiences.length) {
            children.push(new Paragraph({
                spacing: { before: 200, after: 60 },
                border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '333333', space: 4 } },
                children: [new TextRun({ text: 'EXPERIENCE', bold: true, size: 24, font: 'Georgia' })]
            }));
            data.experiences.forEach(e => {
                if (e.title) {
                    children.push(new Paragraph({
                        spacing: { before: 100, after: 0 },
                        children: [
                            new TextRun({ text: e.title, bold: true, size: 24, font: 'Georgia' }),
                            ...(e.company || e.dates ? [new TextRun({ text: ' | ' + (e.company || '') + (e.dates ? ' (' + e.dates + ')' : ''), size: 20, color: '555555', font: 'Georgia' })] : [])
                        ]
                    }));
                }
                if (e.description) {
                    children.push(new Paragraph({
                        spacing: { after: 100 },
                        indent: { left: 360 },
                        children: [new TextRun({ text: e.description, size: 22, font: 'Georgia' })]
                    }));
                }
            });
        }

        if (data.education && data.education.length) {
            children.push(new Paragraph({
                spacing: { before: 200, after: 60 },
                border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '333333', space: 4 } },
                children: [new TextRun({ text: 'EDUCATION', bold: true, size: 24, font: 'Georgia' })]
            }));
            data.education.forEach(e => {
                children.push(new Paragraph({
                    spacing: { after: 60 },
                    children: [new TextRun({ text: (e.degree || '') + ' — ' + (e.institution || '') + (e.year ? ' (' + e.year + ')' : ''), size: 22, font: 'Georgia' })]
                }));
            });
        }

        if (data.skills && data.skills.trim()) {
            children.push(new Paragraph({
                spacing: { before: 200, after: 60 },
                border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '333333', space: 4 } },
                children: [new TextRun({ text: 'SKILLS', bold: true, size: 24, font: 'Georgia' })]
            }));
            children.push(new Paragraph({
                spacing: { after: 120 },
                children: [new TextRun({ text: data.skills, size: 22, font: 'Georgia' })]
            }));
        }

        const doc = new Document({
            sections: [{
                properties: {
                    page: {
                        size: { width: 11906, height: 16838 },
                        margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 }
                    }
                },
                children: children
            }]
        });

        const buffer = await Packer.toBuffer(doc);
        console.log('DOCX size:', buffer.length);
        
        res.writeHead(200, {
            'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            'Content-Length': buffer.length,
            'Content-Disposition': 'attachment; filename="cv-genie.docx"'
        });
        res.end(buffer);
    } catch (error) {
        console.error('DOCX error:', error.message);
        res.status(500).json({ error: error.message });
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
