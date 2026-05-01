const express = require('express');
const puppeteer = require('puppeteer');
const cors = require('cors');
const app = express();

// Allow CV Genie to call this server from any domain
app.use(cors());
app.use(express.json({ limit: '5mb' }));

// Health check
app.get('/', (req, res) => {
  res.send('✅ CV Genie PDF Server is running');
});

// PDF export endpoint
app.post('/export-pdf', async (req, res) => {
  const { html, format = 'A4' } = req.body;
  
  if (!html) {
    return res.status(400).json({ error: 'No HTML provided' });
  }

  let browser;
  try {
    // Launch headless Chrome
    browser = await puppeteer.launch({
      headless: 'new',
      args: [
        '--no-sandbox',
        '--disable-setuid-sandbox',
        '--disable-dev-shm-usage',
        '--disable-gpu'
      ]
    });

    const page = await browser.newPage();
    
    // Set viewport to A4 width at print resolution
    await page.setViewport({
      width: 794,  // A4 width in pixels at 96dpi
      height: 1123, // A4 height
      deviceScaleFactor: 2
    });

    // Load the HTML content
    await page.setContent(html, {
      waitUntil: 'networkidle0',
      timeout: 15000
    });

    // Generate PDF with exact A4 dimensions
    const pdf = await page.pdf({
      format: 'A4',
      margin: {
        top: '8mm',
        bottom: '8mm',
        left: '8mm',
        right: '8mm'
      },
      printBackground: true,
      preferCSSPageSize: true
    });

    // Send the PDF back
    res.set({
      'Content-Type': 'application/pdf',
      'Content-Disposition': 'attachment; filename="cv-genie.pdf"',
      'Content-Length': pdf.length
    });
    res.send(pdf);

  } catch (error) {
    console.error('PDF generation failed:', error);
    res.status(500).json({ error: 'PDF generation failed: ' + error.message });
  } finally {
    if (browser) await browser.close();
  }
});

// Start server
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`CV Genie PDF Server running on port ${PORT}`);
});
