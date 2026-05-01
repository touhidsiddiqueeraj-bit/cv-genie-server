// This runs after npm install and ensures Chromium is ready
const { execSync } = require('child_process');
try {
  console.log('Setting up Chromium...');
  // This finds the chromium binary from @sparticuz/chromium
  const chromium = require('@sparticuz/chromium');
  console.log('Chromium path:', chromium.path);
  console.log('Setup complete!');
} catch (e) {
  console.log('Postinstall note:', e.message);
}
