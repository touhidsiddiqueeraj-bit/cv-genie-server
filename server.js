const express = require('express');
const cors = require('cors');
const { Document, Packer, Paragraph, TextRun, AlignmentType, BorderStyle } = require('docx');

const app = express();
app.use(cors());
app.use(express.json({ limit: '5mb' }));

// ============ PUPPETEER SETUP ============
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

// ============ ROUTES ============
app.get('/', async (req, res) => {
    const { puppeteer: pptr } = await initPuppeteer();
    res.json({
        status: 'running',
        version: '2.0',
        puppeteer: !!pptr,
        endpoints: ['/export-pdf', '/export-docx']
    });
});

// PDF export
app.post('/export-pdf', async (req, res) => {
    const { html } = req.body;
    if (!html) return res.status(400).json({ error: 'No HTML provided' });

    const { puppeteer: pptr, chromium: chrom } = await initPuppeteer();
    
    if (!pptr || !chrom) {
        return res.status(500).json({
            error: 'PDF service unavailable — Puppeteer failed to load. Check server logs.'
        });
    }

    let browser;
    try {
        console.log('Launching browser...');
        console.log('Chromium path:', await chrom.executablePath());
        
        browser = await pptr.launch({
            args: [...chrom.args, '--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage'],
            defaultViewport: chrom.defaultViewport,
            executablePath: await chrom.executablePath(),
            headless: true,
            ignoreHTTPSErrors: true,
        });

        console.log('Browser launched, creating page...');
        const page = await browser.newPage();
        await page.setContent(html, { waitUntil: 'networkidle0', timeout: 30000 });
        console.log('Page loaded, generating PDF...');
        
        const pdf = await page.pdf({
            format: 'A4',
            margin: { top: '8mm', bottom: '8mm', left: '8mm', right: '8mm' },
            printBackground: true,
        });

        console.log('PDF generated, size:', pdf.length, 'bytes');
        res.set({
            'Content-Type': 'application/pdf',
            'Content-Disposition': 'attachment; filename="cv-genie.pdf"',
            'Content-Length': pdf.length
        });
        res.send(pdf);
    } catch (error) {
        console.error('PDF error:', error);
        res.status(500).json({ error: 'PDF generation failed: ' + error.message });
    } finally {
        if (browser) {
            await browser.close();
            console.log('Browser closed');
        }
    }
});

// DOCX export
app.post('/export-docx', async (req, res) => {
    const data = req.body;
    if (!data) return res.status(400).json({ error: 'No data provided' });

    try {
        const children = [];

        // Name
        children.push(new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 120 },
            children: [new TextRun({ text: data.name || 'Your Name', bold: true, size: 36, font: 'Georgia' })]
        }));

        // Title
        if (data.title) {
            children.push(new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 40 },
                children: [new TextRun({ text: data.title, size: 24, color: '555555', font: 'Georgia' })]
            }));
        }

        // Contact
        if (data.contact) {
            children.push(new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 200 },
                border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC', space: 8 } },
                children: [new TextRun({ text: data.contact, size: 18, color: '666666', font: 'Georgia' })]
            }));
        }

        // Summary
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

        // Experience
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

        // Education
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

        // Skills
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
        console.log('DOCX generated, size:', buffer.length, 'bytes');
        res.set({
            'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            'Content-Disposition': 'attachment; filename="cv-genie.docx"',
            'Content-Length': buffer.length
        });
        res.send(buffer);
    } catch (error) {
        console.error('DOCX error:', error);
        res.status(500).json({ error: 'DOCX failed: ' + error.message });
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
