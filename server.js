const express = require('express');
const puppeteer = require('puppeteer-core');
const chromium = require('@sparticuz/chromium');
const cors = require('cors');
const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow, TableCell, WidthType, BorderStyle } = require('docx');

const app = express();
app.use(cors());
app.use(express.json({ limit: '5mb' }));

// Health check
app.get('/', (req, res) => {
  res.send('✅ CV Genie Server is running — PDF + DOCX ready');
});

// PDF export endpoint
app.post('/export-pdf', async (req, res) => {
  const { html, format = 'A4' } = req.body;
  if (!html) return res.status(400).json({ error: 'No HTML provided' });

  let browser;
  try {
    browser = await puppeteer.launch({
      args: [...chromium.args, '--no-sandbox', '--disable-setuid-sandbox'],
      defaultViewport: chromium.defaultViewport,
      executablePath: await chromium.executablePath(),
      headless: chromium.headless,
    });

    const page = await browser.newPage();
    await page.setContent(html, { waitUntil: 'networkidle0', timeout: 20000 });
    
    const pdf = await page.pdf({ format: 'A4', margin: { top: '8mm', bottom: '8mm', left: '8mm', right: '8mm' }, printBackground: true });
    
    res.set({ 'Content-Type': 'application/pdf', 'Content-Disposition': 'attachment; filename="cv-genie.pdf"', 'Content-Length': pdf.length });
    res.send(pdf);
  } catch (error) {
    console.error('PDF error:', error);
    res.status(500).json({ error: 'PDF generation failed: ' + error.message });
  } finally {
    if (browser) await browser.close();
  }
});

// DOCX export endpoint — generates proper editable Word document
app.post('/export-docx', async (req, res) => {
  const data = req.body;
  if (!data) return res.status(400).json({ error: 'No data provided' });

  try {
    const doc = new Document({
      styles: {
        default: { document: { run: { font: 'Georgia', size: 22 } } },
      },
      sections: [{
        properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } },
        children: [
          // Name
          new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: data.name || 'Your Name', bold: true, size: 36, font: 'Georgia' })] }),
          // Title
          ...(data.title ? [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [new TextRun({ text: data.title, size: 24, color: '555555', font: 'Georgia' })] })] : []),
          // Contact
          ...(data.contact ? [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 120 }, children: [new TextRun({ text: data.contact, size: 20, color: '666666' })] })] : []),
          // Summary
          ...(data.summary ? [
            new Paragraph({ spacing: { before: 200, after: 80 }, border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '333333' } }, children: [new TextRun({ text: 'PROFESSIONAL SUMMARY', bold: true, size: 24, font: 'Georgia', allCaps: true })] }),
            new Paragraph({ spacing: { after: 160 }, children: [new TextRun({ text: data.summary, size: 22, font: 'Georgia' })] }),
          ] : []),
          // Experience
          ...(data.experiences && data.experiences.length ? [
            new Paragraph({ spacing: { before: 300, after: 80 }, border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '333333' } }, children: [new TextRun({ text: 'EXPERIENCE', bold: true, size: 24, font: 'Georgia', allCaps: true })] }),
            ...data.experiences.flatMap(e => [
              new Paragraph({ spacing: { before: 120, after: 0 }, children: [new TextRun({ text: e.title || '', bold: true, size: 24, font: 'Georgia' }), new TextRun({ text: '  |  ' + (e.company || '') + '  (' + (e.dates || '') + ')', size: 20, color: '555555' })] }),
              ...(e.description ? [new Paragraph({ spacing: { after: 120 }, indent: { left: 360 }, children: [new TextRun({ text: e.description, size: 22, font: 'Georgia' })] })] : []),
            ]),
          ] : []),
          // Education
          ...(data.education && data.education.length ? [
            new Paragraph({ spacing: { before: 300, after: 80 }, border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '333333' } }, children: [new TextRun({ text: 'EDUCATION', bold: true, size: 24, font: 'Georgia', allCaps: true })] }),
            ...data.education.map(e => new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: (e.degree || '') + ' — ' + (e.institution || '') + ' (' + (e.year || '') + ')', size: 22, font: 'Georgia' })] })),
          ] : []),
          // Skills
          ...(data.skills ? [
            new Paragraph({ spacing: { before: 300, after: 80 }, border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '333333' } }, children: [new TextRun({ text: 'SKILLS', bold: true, size: 24, font: 'Georgia', allCaps: true })] }),
            new Paragraph({ children: [new TextRun({ text: data.skills, size: 22, font: 'Georgia' })] }),
          ] : []),
        ],
      }],
    });

    const buffer = await Packer.toBuffer(doc);
    res.set({ 'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', 'Content-Disposition': 'attachment; filename="cv-genie.docx"' });
    res.send(buffer);
  } catch (error) {
    console.error('DOCX error:', error);
    res.status(500).json({ error: 'DOCX generation failed: ' + error.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`CV Genie Server running on port ${PORT}`));
