const express = require('express');
const puppeteer = require('puppeteer-core');
const cors = require('cors');
const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, BorderStyle, TabStopPosition, TabStopType, Indent } = require('docx');

const app = express();
app.use(cors());
app.use(express.json({ limit: '5mb' }));

// Lazy-load chromium to avoid errors on import
let chromium;
async function getChromium() {
    if (!chromium) {
        chromium = require('@sparticuz/chromium');
    }
    return chromium;
}

app.get('/', (req, res) => {
    res.send('✅ CV Genie Server v2.0 — PDF + DOCX ready');
});

// PDF export
app.post('/export-pdf', async (req, res) => {
    const { html } = req.body;
    if (!html) return res.status(400).json({ error: 'No HTML provided' });

    let browser;
    try {
        const chrom = await getChromium();
        browser = await puppeteer.launch({
            args: [...chrom.args, '--no-sandbox', '--disable-setuid-sandbox'],
            defaultViewport: chrom.defaultViewport,
            executablePath: await chrom.executablePath(),
            headless: chrom.headless,
        });

        const page = await browser.newPage();
        await page.setContent(html, { waitUntil: 'networkidle0', timeout: 20000 });
        
        const pdf = await page.pdf({
            format: 'A4',
            margin: { top: '8mm', bottom: '8mm', left: '8mm', right: '8mm' },
            printBackground: true,
            preferCSSPageSize: true
        });

        res.set({
            'Content-Type': 'application/pdf',
            'Content-Disposition': 'attachment; filename="cv-genie.pdf"',
            'Content-Length': pdf.length
        });
        res.send(pdf);
    } catch (error) {
        console.error('PDF error:', error.message);
        res.status(500).json({ error: 'PDF generation failed: ' + error.message });
    } finally {
        if (browser) await browser.close();
    }
});

// DOCX export
app.post('/export-docx', async (req, res) => {
    const data = req.body;
    if (!data || !data.name) return res.status(400).json({ error: 'No data provided' });

    try {
        const children = [];

        // Name — large, bold, centered
        children.push(new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 120 },
            children: [new TextRun({ text: data.name || 'Your Name', bold: true, size: 36, font: 'Georgia' })]
        }));

        // Title
        if (data.title) {
            children.push(new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 80 },
                children: [new TextRun({ text: data.title, size: 24, color: '555555', font: 'Georgia' })]
            }));
        }

        // Contact line with separator
        if (data.contact) {
            children.push(new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 60 },
                border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC', space: 8 } },
                children: [new TextRun({ text: data.contact, size: 18, color: '666666', font: 'Georgia' })]
            }));
        }

        // Summary section
        if (data.summary && data.summary.trim()) {
            children.push(new Paragraph({
                spacing: { before: 240, after: 60 },
                border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '333333', space: 4 } },
                children: [new TextRun({ text: 'PROFESSIONAL SUMMARY', bold: true, size: 24, font: 'Georgia' })]
            }));
            children.push(new Paragraph({
                spacing: { after: 160 },
                children: [new TextRun({ text: data.summary, size: 22, font: 'Georgia' })]
            }));
        }

        // Experience section
        if (data.experiences && data.experiences.length) {
            children.push(new Paragraph({
                spacing: { before: 300, after: 60 },
                border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '333333', space: 4 } },
                children: [new TextRun({ text: 'EXPERIENCE', bold: true, size: 24, font: 'Georgia' })]
            }));
            data.experiences.forEach(e => {
                if (e.title || e.company) {
                    children.push(new Paragraph({
                        spacing: { before: 120, after: 0 },
                        children: [
                            new TextRun({ text: e.title || '', bold: true, size: 24, font: 'Georgia' }),
                            ...(e.company || e.dates ? [new TextRun({ text: ' | ' + (e.company || '') + (e.dates ? ' (' + e.dates + ')' : ''), size: 20, color: '555555', font: 'Georgia' })] : [])
                        ]
                    }));
                }
                if (e.description) {
                    children.push(new Paragraph({
                        spacing: { after: 120 },
                        indent: { left: 360 },
                        children: [new TextRun({ text: e.description, size: 22, font: 'Georgia', italics: false })]
                    }));
                }
            });
        }

        // Education section
        if (data.education && data.education.length) {
            children.push(new Paragraph({
                spacing: { before: 300, after: 60 },
                border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '333333', space: 4 } },
                children: [new TextRun({ text: 'EDUCATION', bold: true, size: 24, font: 'Georgia' })]
            }));
            data.education.forEach(e => {
                children.push(new Paragraph({
                    spacing: { after: 60 },
                    children: [
                        new TextRun({ text: (e.degree || '') + ' — ' + (e.institution || '') + (e.year ? ' (' + e.year + ')' : ''), size: 22, font: 'Georgia' })
                    ]
                }));
            });
        }

        // Skills section
        if (data.skills && data.skills.trim()) {
            children.push(new Paragraph({
                spacing: { before: 300, after: 60 },
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
        res.set({
            'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            'Content-Disposition': 'attachment; filename="cv-genie.docx"',
            'Content-Length': buffer.length
        });
        res.send(buffer);
    } catch (error) {
        console.error('DOCX error:', error.message);
        res.status(500).json({ error: 'DOCX generation failed: ' + error.message });
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`CV Genie Server running on port ${PORT}`));
