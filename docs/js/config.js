/**
 * docs/js/config.js
 *
 * ⚠️  EDIT THIS FILE after deploying the Apps Script web app.
 *
 * 1. Open your Google Spreadsheet → Extensions → Apps Script
 * 2. Paste the contents of appsscript/Code.gs
 * 3. Click Deploy → New deployment → Web App
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 4. Copy the URL that looks like:
 *    https://script.google.com/macros/s/AKfycb.../exec
 * 5. Replace the placeholder below with that URL.
 */

const CONFIG = {
  // Replace with your deployed Apps Script URL
  APPS_SCRIPT_URL: 'https://script.google.com/macros/s/AKfycbwzNr39q-mLc-FPuJRSVkEAjeQdsBbMUHp9laMXFBVthNJk6iZxZwo50gLlYLBCrI9pgg/exec',

  // Pharmacy name shown in the PDF header and WhatsApp message
  FARMACIA_NOME: 'Farmácia',

  // Batch size for CSV upload (rows per request – keeps requests under 6 MB)
  BATCH_SIZE: 200,
};
