/**
 * Kings Equestrian - Complete Email System with Consent & Payment Tracking
 * Features:
 * - HTML consent page (replaces Google Form)
 * - Registration number generation on consent acceptance
 * - PDF attachment of Google Docs
 * - Payment details tracking in separate sheet
 * - Queue management with error logging
 */

// ============== CONFIGURATION ==============
const CONFIG = {
  // Email settings
  emailSubject: 'Welcome {name} to Kings Equestrian 🏇',
  
  // Google Docs ID for email content
  emailTemplateDocId: '1qALQ8RlVlMrpEhjLYjaf3txDUo5BZLoHjOoRTrS191g',
  
  // Web App URL (will be set after deployment)
  webAppUrl: '', // Leave empty, will be auto-filled
  
  // Links
  websiteLink: 'https://kingsfarmequestrian.com',
  instagramLink: '@kingsequestrianfoundation',
  reviewsLink: 'https://maps.app.goo.gl/EVyzEfhh3tdJ2BTX7?g_st=iwb',
  
  // Contact
  whatsappNumbers: '9980771166 | 9980895533',
  email: 'info@kingsequestrian.com',
  
  // Location codes for registration numbers
  locationCodes: {
    'bangalore': 'BLR',
    'hyderabad': 'HYD',
    'pune': 'PNE'
  },
  
  // Default payment amount (can be customized per registration)
  defaultPaymentAmount: 2500,
  
  // Email quota management
  DAILY_EMAIL_LIMIT: 95,
  
  colors: {
    primary: '#1a472a',
    secondary: '#d4af37',
    accent: '#2d5a3d',
    background: '#f8f9fa',
    text: '#333333'
  }
};

// ============== WEB APP HANDLERS ==============

/**
 * Handles GET requests - serves the consent page
 */
function doGet(e) {
  const email = e.parameter.email || '';
  const name = e.parameter.name || '';
  const location = e.parameter.location || '';
  
  if (!email) {
    return HtmlService.createHtmlOutput('<h1>Invalid Access</h1><p>Missing email parameter.</p>');
  }
  
  // Serve the consent form
  const template = HtmlService.createTemplateFromFile('ConsentPage');
  template.email = email;
  template.name = name;
  template.location = location;
  template.colors = CONFIG.colors;
  
  return template.evaluate()
    .setTitle('Terms & Conditions - Kings Equestrian')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Handles POST requests - processes consent acceptance
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const email = data.email;
    const name = data.name;
    const location = data.location;
    const consentAccepted = data.consentAccepted;
    
    if (!consentAccepted) {
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        message: 'Consent must be accepted'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // Generate registration number
    const registrationNumber = generateRegistrationNumber(location);
    
    // Save to Payment Details sheet
    const paymentAmount = CONFIG.defaultPaymentAmount;
    saveToPaymentDetails(registrationNumber, paymentAmount, email, name, location);
    
    // Update main sheet with registration number and consent status
    updateMainSheetWithConsent(email, registrationNumber, consentAccepted);
    
    // Generate payment link
    const paymentLink = createPaymentLink(registrationNumber, paymentAmount);
    
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      registrationNumber: registrationNumber,
      paymentAmount: paymentAmount,
      paymentLink: paymentLink,
      message: 'Consent accepted successfully!'
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    Logger.log('Error in doPost: ' + error.toString());
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: 'Error processing consent: ' + error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ============== REGISTRATION NUMBER GENERATION ==============

/**
 * Generates unique registration number based on location
 * Format: [LOCATION_CODE][4-digit serial number]
 */
function generateRegistrationNumber(location) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const paymentSheet = getPaymentDetailsSheet();
  
  // Get location code
  const locationKey = Object.keys(CONFIG.locationCodes).find(key => 
    location.toLowerCase().includes(key) || key.includes(location.toLowerCase())
  );
  const locationCode = locationKey ? CONFIG.locationCodes[locationKey] : 'GEN';
  
  // Find highest existing number for this location
  const lastRow = paymentSheet.getLastRow();
  let maxNumber = 0;
  
  if (lastRow > 1) {
    const existingNumbers = paymentSheet.getRange(2, 1, lastRow - 1, 1).getValues();
    const pattern = new RegExp(`^${locationCode}(\\d+)$`);
    
    existingNumbers.forEach(row => {
      const regNumber = row[0].toString();
      const match = regNumber.match(pattern);
      if (match) {
        const num = parseInt(match[1]);
        if (num > maxNumber) {
          maxNumber = num;
        }
      }
    });
  }
  
  const newNumber = maxNumber + 1;
  const paddedNumber = newNumber.toString().padStart(4, '0');
  
  return `${locationCode}${paddedNumber}`;
}

// ============== PAYMENT DETAILS SHEET ==============

/**
 * Get or create Payment Details sheet
 */
function getPaymentDetailsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let paymentSheet = ss.getSheetByName('Payment Details');
  
  if (!paymentSheet) {
    paymentSheet = ss.insertSheet('Payment Details');
    
    const headers = [
      'Registration Number',
      'Amount to be Paid',
      'Email ID',
      'Student Name',
      'Location',
      'Consent Accepted',
      'Consent Timestamp',
      'Payment Status',
      'Payment Link'
    ];
    
    paymentSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    paymentSheet.getRange(1, 1, 1, headers.length)
      .setBackground(CONFIG.colors.primary)
      .setFontColor('white')
      .setFontWeight('bold');
    
    paymentSheet.setFrozenRows(1);
    paymentSheet.autoResizeColumns(1, headers.length);
  }
  
  return paymentSheet;
}

/**
 * Save consent and payment details
 */
function saveToPaymentDetails(registrationNumber, amount, email, name, location) {
  const paymentSheet = getPaymentDetailsSheet();
  
  const paymentLink = createPaymentLink(registrationNumber, amount);
  
  const rowData = [
    registrationNumber,
    amount,
    email,
    name,
    location || '',
    'Yes',
    new Date(),
    'Pending',
    paymentLink
  ];
  
  paymentSheet.appendRow(rowData);
  
  // Format the new row
  const lastRow = paymentSheet.getLastRow();
  paymentSheet.getRange(lastRow, 7).setNumberFormat("dd-MMM-yyyy HH:mm:ss"); // Timestamp
  paymentSheet.getRange(lastRow, 8).setBackground('#fff3cd').setFontColor('#856404'); // Payment Status
  
  Logger.log(`Payment details saved for ${registrationNumber}`);
}

/**
 * Create payment link
 */
function createPaymentLink(registrationNumber, amount) {
  // Replace with your actual payment page URL
  return `https://kings-equestrian.web.app/pay?ref=${registrationNumber}&am=${amount}`;
}

// ============== MAIN SHEET UPDATES ==============

/**
 * Update main sheet with registration number and consent status
 */
function updateMainSheetWithConsent(email, registrationNumber, consentAccepted) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheets()[0]; // Main registration sheet
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Add columns if they don't exist
    let regCol = headers.indexOf('Registration Number') + 1;
    if (regCol === 0) {
      regCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, regCol)
        .setValue('Registration Number')
        .setBackground(CONFIG.colors.primary)
        .setFontColor('white')
        .setFontWeight('bold');
    }
    
    let consentCol = headers.indexOf('Consent Accepted') + 1;
    if (consentCol === 0) {
      consentCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, consentCol)
        .setValue('Consent Accepted')
        .setBackground(CONFIG.colors.primary)
        .setFontColor('white')
        .setFontWeight('bold');
    }
    
    let consentTimeCol = headers.indexOf('Consent Timestamp') + 1;
    if (consentTimeCol === 0) {
      consentTimeCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, consentTimeCol)
        .setValue('Consent Timestamp')
        .setBackground(CONFIG.colors.primary)
        .setFontColor('white')
        .setFontWeight('bold');
    }
    
    // Find the row with this email
    const emailCol = headers.indexOf('Email ID') + 1;
    const dataRange = sheet.getRange(2, emailCol, sheet.getLastRow() - 1, 1);
    const emails = dataRange.getValues();
    
    for (let i = 0; i < emails.length; i++) {
      if (emails[i][0] === email) {
        const row = i + 2;
        
        sheet.getRange(row, regCol).setValue(registrationNumber);
        sheet.getRange(row, consentCol).setValue(consentAccepted ? 'Yes' : 'No');
        sheet.getRange(row, consentTimeCol).setValue(new Date()).setNumberFormat("dd-MMM-yyyy HH:mm:ss");
        
        break;
      }
    }
    
  } catch (error) {
    Logger.log('Error updating main sheet: ' + error.toString());
  }
}

// ============== PDF GENERATION ==============

/**
 * Convert Google Docs to PDF and return as blob
 */
function getDocAsPDF() {
  try {
    const docId = CONFIG.emailTemplateDocId;
    const doc = DriveApp.getFileById(docId);
    
    // Export as PDF
    const pdfBlob = doc.getAs('application/pdf');
    pdfBlob.setName('Kings_Equestrian_Information.pdf');
    
    return pdfBlob;
  } catch (error) {
    Logger.log('Error creating PDF: ' + error.toString());
    return null;
  }
}

// ============== CUSTOM MENU ==============

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📧 Kings Equestrian')
    .addItem('🔄 Resend Selected Emails', 'resendSelectedEmails')
    .addItem('📊 View Email Usage', 'showEmailUsage')
    .addItem('⚙️ Process Email Queue', 'processEmailQueue')
    .addItem('🧪 Test Email Template', 'testEmailTemplate')
    .addItem('🔗 Get Consent Page URL', 'showConsentPageURL')
    .addToUi();
}

/**
 * Show consent page URL for testing
 */
function showConsentPageURL() {
  const ui = SpreadsheetApp.getUi();
  const webAppUrl = getWebAppUrl();
  
  const testUrl = `${webAppUrl}?email=test@example.com&name=Test%20User&location=bangalore`;
  
  ui.alert(
    'Consent Page URL',
    `Base URL: ${webAppUrl}\n\nTest URL:\n${testUrl}\n\nUse this URL in your emails.`,
    ui.ButtonSet.OK
  );
}

/**
 * Get deployed web app URL
 */
function getWebAppUrl() {
  const url = ScriptApp.getService().getUrl();
  return url;
}

// ============== EMAIL QUOTA TRACKING ==============

function getEmailUsageToday() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const today = new Date().toDateString();
  const key = `emailCount_${today}`;
  const count = scriptProperties.getProperty(key);
  return count ? parseInt(count) : 0;
}

function incrementEmailUsage() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const today = new Date().toDateString();
  const key = `emailCount_${today}`;
  const currentCount = getEmailUsageToday();
  scriptProperties.setProperty(key, (currentCount + 1).toString());
}

function canSendEmailToday() {
  return getEmailUsageToday() < CONFIG.DAILY_EMAIL_LIMIT;
}

function resetDailyEmailCounter() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const yesterday = new Date(Date.now() - 24 * 60 * 60 * 1000).toDateString();
  const oldKey = `emailCount_${yesterday}`;
  scriptProperties.deleteProperty(oldKey);
  Logger.log('Daily email counter reset');
}

// ============== GOOGLE DOCS CONTENT FETCHER ==============

function getEmailContentFromDocs() {
  try {
    const doc = DocumentApp.openById(CONFIG.emailTemplateDocId);
    const body = doc.getBody();
    const text = body.getText();
    
    return {
      success: true,
      content: text,
      html: formatDocsContentToHTML(text)
    };
  } catch (error) {
    Logger.log('Error fetching Google Docs content: ' + error.toString());
    return {
      success: false,
      error: error.toString(),
      content: null
    };
  }
}

function formatDocsContentToHTML(text) {
  const sections = text.split(/(?=🤍|✨|🌿|📸|🌄|🎉)/);
  
  let html = '';
  
  sections.forEach(section => {
    if (section.trim()) {
      if (section.includes('Horse Riding') || 
          section.includes('Horse Safari') || 
          section.includes('Photography') ||
          section.includes('Training') ||
          section.includes('Leadership') ||
          section.includes('Events')) {
        
        const lines = section.split('\n');
        const title = lines[0];
        const content = lines.slice(1).join('<br>');
        
        html += `
          <div style="margin: 25px 0; padding: 20px; background: ${CONFIG.colors.background}; border-radius: 8px; border-left: 4px solid ${CONFIG.colors.secondary};">
            <h3 style="color: ${CONFIG.colors.primary}; margin-top: 0;">${title}</h3>
            <p style="margin: 0; color: ${CONFIG.colors.text};">${content}</p>
          </div>
        `;
      } else {
        html += `<p style="margin: 15px 0; line-height: 1.6;">${section.trim().replace(/\n/g, '<br>')}</p>`;
      }
    }
  });
  
  return html;
}

// ============== EMAIL QUEUE MANAGEMENT ==============

function getEmailQueueSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let queueSheet = ss.getSheetByName('Email Queue');
  
  if (!queueSheet) {
    queueSheet = ss.insertSheet('Email Queue');
    
    const headers = [
      'Timestamp',
      'Email ID',
      'Student Name',
      'Location',
      'Status',
      'Attempts',
      'Last Attempt',
      'Error Message'
    ];
    
    queueSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    queueSheet.getRange(1, 1, 1, headers.length)
      .setBackground(CONFIG.colors.primary)
      .setFontColor('white')
      .setFontWeight('bold');
    
    queueSheet.setFrozenRows(1);
    queueSheet.autoResizeColumns(1, headers.length);
  }
  
  return queueSheet;
}

function addToEmailQueue(email, name, location, errorMessage = '') {
  const queueSheet = getEmailQueueSheet();
  
  const rowData = [
    new Date(),
    email,
    name,
    location,
    errorMessage ? 'Failed' : 'Pending',
    0,
    '',
    errorMessage
  ];
  
  queueSheet.appendRow(rowData);
  Logger.log(`Email queued for ${email}`);
  
  updateEmailStatusInMainSheet(email, errorMessage ? 'Failed - Queued' : 'Queued for Tomorrow', errorMessage);
}

function updateEmailStatusInMainSheet(email, status, errorMessage = '') {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheets()[0];
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    let statusCol = headers.indexOf('Email Status') + 1;
    if (statusCol === 0) {
      statusCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, statusCol)
        .setValue('Email Status')
        .setBackground(CONFIG.colors.primary)
        .setFontColor('white')
        .setFontWeight('bold');
    }

    let timestampCol = headers.indexOf('Email Sent Timestamp') + 1;
    if (timestampCol === 0) {
      timestampCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, timestampCol)
        .setValue('Email Sent Timestamp')
        .setBackground(CONFIG.colors.primary)
        .setFontColor('white')
        .setFontWeight('bold');
    }
    
    let errorCol = headers.indexOf('Error Message') + 1;
    if (errorCol === 0) {
      errorCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, errorCol)
        .setValue('Error Message')
        .setBackground(CONFIG.colors.primary)
        .setFontColor('white')
        .setFontWeight('bold');
    }

    const emailCol = headers.indexOf('Email ID') + 1;
    const dataRange = sheet.getRange(2, emailCol, sheet.getLastRow() - 1, 1);
    const emails = dataRange.getValues();

    for (let i = 0; i < emails.length; i++) {
      if (emails[i][0] === email) {
        const row = i + 2;

        const statusCell = sheet.getRange(row, statusCol);
        statusCell.setValue(status);

        const timestampCell = sheet.getRange(row, timestampCol);
        timestampCell.setValue(new Date()).setNumberFormat("dd-MMM-yyyy HH:mm:ss");
        
        if (errorMessage) {
          sheet.getRange(row, errorCol).setValue(errorMessage);
        }

        if (status === 'Sent') {
          statusCell.setBackground('#d4edda').setFontColor('#155724');
        } else if (status.includes('Queued')) {
          statusCell.setBackground('#fff3cd').setFontColor('#856404');
        } else if (status.includes('Failed')) {
          statusCell.setBackground('#f8d7da').setFontColor('#721c24');
        }
        break;
      }
    }
  } catch (error) {
    Logger.log('Error updating email status: ' + error.toString());
  }
}

function processEmailQueue() {
  const queueSheet = getEmailQueueSheet();
  const lastRow = queueSheet.getLastRow();
  
  if (lastRow <= 1) {
    SpreadsheetApp.getUi().alert('Email queue is empty');
    return;
  }
  
  const data = queueSheet.getRange(2, 1, lastRow - 1, 8).getValues();
  const rowsToDelete = [];
  let emailsSentCount = 0;
  
  for (let i = 0; i < data.length; i++) {
    if (!canSendEmailToday()) {
      Logger.log(`Daily limit reached. ${data.length - i} emails remain`);
      break;
    }
    
    const [timestamp, email, name, location, status, attempts] = data[i];
    
    if (status === 'Sent') {
      rowsToDelete.push(i + 2);
      continue;
    }
    
    try {
      const result = sendWelcomeEmail(email, name, location);
      
      if (result.success) {
        incrementEmailUsage();
        emailsSentCount++;
        
        queueSheet.getRange(i + 2, 5).setValue('Sent');
        queueSheet.getRange(i + 2, 6).setValue(attempts + 1);
        queueSheet.getRange(i + 2, 7).setValue(new Date());
        queueSheet.getRange(i + 2, 8).setValue('');
        queueSheet.getRange(i + 2, 5).setBackground('#d4edda').setFontColor('#155724');
        
        updateEmailStatusInMainSheet(email, 'Sent');
        rowsToDelete.push(i + 2);
        
        Logger.log(`✅ Email sent to ${email}`);
      } else {
        throw new Error(result.error);
      }
      
      Utilities.sleep(1000);
      
    } catch (error) {
      const errorMsg = error.toString();
      Logger.log(`❌ Failed: ${email}: ${errorMsg}`);
      
      queueSheet.getRange(i + 2, 5).setValue('Failed');
      queueSheet.getRange(i + 2, 6).setValue(attempts + 1);
      queueSheet.getRange(i + 2, 7).setValue(new Date());
      queueSheet.getRange(i + 2, 8).setValue(errorMsg.substring(0, 200));
      queueSheet.getRange(i + 2, 5).setBackground('#f8d7da').setFontColor('#721c24');
      
      updateEmailStatusInMainSheet(email, 'Failed', errorMsg);
    }
  }
  
  rowsToDelete.reverse().forEach(row => {
    queueSheet.deleteRow(row);
  });
  
  Logger.log(`Queue complete. Sent: ${emailsSentCount}, Remaining: ${queueSheet.getLastRow() - 1}`);
  SpreadsheetApp.getUi().alert(`Sent: ${emailsSentCount}\nRemaining: ${queueSheet.getLastRow() - 1}`);
}

// ============== RESEND FUNCTIONALITY ==============

function resendSelectedEmails() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const selection = sheet.getActiveRange();
  
  if (!selection) {
    ui.alert('Please select rows to resend');
    return;
  }
  
  const response = ui.alert(
    'Resend Emails',
    `Resend emails for ${selection.getNumRows()} row(s)?`,
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  const startRow = selection.getRow();
  const numRows = selection.getNumRows();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  const emailCol = headers.indexOf('Email ID') + 1;
  const nameCol = headers.indexOf('Student Name') + 1;
  const locationCol = headers.indexOf('Location') + 1;
  
  if (emailCol === 0 || nameCol === 0) {
    ui.alert('Required columns not found');
    return;
  }
  
  let successCount = 0;
  let failCount = 0;
  
  for (let i = 0; i < numRows; i++) {
    const row = startRow + i;
    if (row === 1) continue;
    
    const email = sheet.getRange(row, emailCol).getValue();
    const name = sheet.getRange(row, nameCol).getValue();
    const location = locationCol > 0 ? sheet.getRange(row, locationCol).getValue() : '';
    
    if (!email || !isValidEmail(email)) {
      failCount++;
      continue;
    }
    
    if (!canSendEmailToday()) {
      ui.alert(`Limit reached! Sent: ${successCount}, Failed: ${failCount}`);
      
      for (let j = i; j < numRows; j++) {
        const qRow = startRow + j;
        if (qRow === 1) continue;
        
        const qEmail = sheet.getRange(qRow, emailCol).getValue();
        const qName = sheet.getRange(qRow, nameCol).getValue();
        const qLocation = locationCol > 0 ? sheet.getRange(qRow, locationCol).getValue() : '';
        
        if (qEmail && isValidEmail(qEmail)) {
          addToEmailQueue(qEmail, qName, qLocation);
        }
      }
      break;
    }
    
    try {
      const result = sendWelcomeEmail(email, name, location);
      
      if (result.success) {
        incrementEmailUsage();
        updateEmailStatusInMainSheet(email, 'Sent');
        successCount++;
      } else {
        throw new Error(result.error);
      }
      
      Utilities.sleep(1000);
      
    } catch (error) {
      updateEmailStatusInMainSheet(email, 'Failed', error.toString());
      failCount++;
    }
  }
  
  ui.alert(`Complete!\n✅ Sent: ${successCount}\n❌ Failed: ${failCount}`);
}

function isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

// ============== EMAIL SENDING ==============

function onFormSubmit(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const row = e.range.getRow();
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const emailCol = headers.indexOf('Email ID') + 1;
    const nameCol = headers.indexOf('Student Name') + 1;
    const locationCol = headers.indexOf('Location') + 1;
    
    if (emailCol === 0 || nameCol === 0) {
      throw new Error('Required columns not found');
    }
    
    const email = sheet.getRange(row, emailCol).getValue();
    const name = sheet.getRange(row, nameCol).getValue();
    const location = locationCol > 0 ? sheet.getRange(row, locationCol).getValue() : '';
    
    if (!email || !isValidEmail(email)) {
      const error = 'Invalid email: ' + email;
      updateEmailStatusInMainSheet(email, 'Failed', error);
      addToEmailQueue(email, name, location, error);
      return;
    }
    
    if (canSendEmailToday()) {
      const result = sendWelcomeEmail(email, name, location);
      
      if (result.success) {
        incrementEmailUsage();
        updateEmailStatusInMainSheet(email, 'Sent');
        Logger.log(`✅ Email sent to ${email}`);
      } else {
        throw new Error(result.error);
      }
    } else {
      addToEmailQueue(email, name, location);
      Logger.log(`Limit reached. Queued: ${email}`);
    }
    
  } catch (error) {
    Logger.log('❌ Error: ' + error.toString());
    
    try {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
      const row = e.range.getRow();
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const emailCol = headers.indexOf('Email ID') + 1;
      const email = sheet.getRange(row, emailCol).getValue();
      
      updateEmailStatusInMainSheet(email, 'Failed', error.toString());
      addToEmailQueue(email, name || 'Unknown', location || '', error.toString());
    } catch (e) {
      Logger.log('Could not update error: ' + e.toString());
    }
  }
}

function sendWelcomeEmail(email, name, location) {
  try {
    if (!email || !isValidEmail(email)) {
      return { success: false, error: 'Invalid email: ' + email };
    }
    
    const docsContent = getEmailContentFromDocs();
    if (!docsContent.success) {
      return { success: false, error: 'Failed to fetch content: ' + docsContent.error };
    }
    
    const webAppUrl = getWebAppUrl();
    const consentUrl = `${webAppUrl}?email=${encodeURIComponent(email)}&name=${encodeURIComponent(name)}&location=${encodeURIComponent(location)}`;
    
    const htmlBody = createEmailTemplate(name, location, docsContent.html, consentUrl);
    const subject = CONFIG.emailSubject.replace('{name}', name);
    
    // Attach PDF
    const pdfBlob = getDocAsPDF();
    
    const options = {
      htmlBody: htmlBody,
      name: 'Kings Equestrian'
    };
    
    if (pdfBlob) {
      options.attachments = [pdfBlob];
    }
    
    MailApp.sendEmail(email, subject, 'Please view in HTML client.', options);
    
    return { success: true, message: 'Email sent to ' + email };
    
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

function createEmailTemplate(name, location, docsHtmlContent, consentUrl) {
  const locationName = location ? (location.charAt(0).toUpperCase() + location.slice(1)) : 'Our Center';
  
  return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <style>
    body {
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
      line-height: 1.6;
      color: ${CONFIG.colors.text};
      margin: 0;
      padding: 0;
      background-color: #f4f4f4;
    }
    .container {
      max-width: 650px;
      margin: 20px auto;
      background-color: white;
      border-radius: 12px;
      overflow: hidden;
      box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
    }
    .header {
      background: linear-gradient(135deg, ${CONFIG.colors.primary} 0%, ${CONFIG.colors.accent} 100%);
      padding: 40px 30px;
      text-align: center;
      color: white;
    }
    .header h1 {
      margin: 10px 0 0 0;
      font-size: 28px;
      font-weight: 600;
    }
    .content {
      padding: 40px 30px;
    }
    .greeting {
      font-size: 20px;
      color: ${CONFIG.colors.primary};
      margin-bottom: 20px;
      font-weight: 600;
    }
    .cta-button {
      display: inline-block;
      background: linear-gradient(135deg, ${CONFIG.colors.primary} 0%, ${CONFIG.colors.accent} 100%);
      color: white;
      padding: 15px 35px;
      text-decoration: none;
      border-radius: 8px;
      font-weight: 600;
      margin: 20px 0;
      box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
    }
    .section-box {
      background: ${CONFIG.colors.background};
      padding: 25px;
      border-radius: 8px;
      margin: 25px 0;
      border-left: 4px solid ${CONFIG.colors.secondary};
    }
    .footer {
      background-color: ${CONFIG.colors.primary};
      color: white;
      padding: 30px;
      text-align: center;
      font-size: 14px;
    }
    @media only screen and (max-width: 600px) {
      .container { margin: 10px; }
      .content { padding: 25px 20px; }
    }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <div>
        <img src="https://kingsfarmequestrian.com/wp-content/uploads/2023/08/Logo2.jpg" width="70" height="70" style="border-radius: 50%;" alt="Logo"/>
      </div>
      <h1>KINGS EQUESTRIAN</h1>
      <p style="margin: 10px 0 0 0; font-size: 16px;">
        Where horses don't just carry you — they change you. 🤍🐎
      </p>
    </div>

    <div class="content">
      <div class="greeting">Dear ${name},</div>
      
      <p>Welcome to <strong>Kings Equestrian</strong>! We're thrilled to have you join our equestrian family.</p>
      
      <p>📎 <strong>Please find attached</strong> our detailed information document for your reference.</p>
      
      <div style="margin: 30px 0;">
        ${docsHtmlContent}
      </div>

      <div class="section-box">
        <h3 style="color: ${CONFIG.colors.primary}; margin-top: 0;">
          ✅ Next Step: Accept Terms & Conditions
        </h3>
        <p style="margin: 15px 0;">
          To proceed with registration and receive your payment details, please review and accept our Terms & Conditions.
        </p>
        <div style="text-align: center;">
          <a href="${consentUrl}" class="cta-button" style="color: white;">
            Review & Accept Terms
          </a>
        </div>
        <p style="margin-top: 15px; font-size: 12px; color: #777; text-align: center;">
          You'll receive your registration number and payment link after acceptance.
        </p>
      </div>

      <div class="section-box">
        <h3 style="color: ${CONFIG.colors.primary}; margin-top: 0;">📞 Contact Us</h3>
        <p style="margin: 10px 0;">
          <strong>📱 WhatsApp:</strong> ${CONFIG.whatsappNumbers}<br>
          <strong>📩 Instagram:</strong> ${CONFIG.instagramLink}<br>
          <strong>🌐 Website:</strong> <a href="${CONFIG.websiteLink}">${CONFIG.websiteLink}</a><br>
          <strong>⭐ Reviews:</strong> <a href="${CONFIG.reviewsLink}">Read experiences</a>
        </p>
      </div>

      <p style="margin-top: 30px; font-style: italic; color: #666;">
        Come for the ride. Leave with a feeling that stays for life. 🤍🐎✨
      </p>
    </div>

    <div class="footer">
      <p style="font-size: 16px; font-weight: 600;">KINGS EQUESTRIAN</p>
      <p>📍 ${locationName}</p>
      <p>📞 ${CONFIG.whatsappNumbers} | ✉️ ${CONFIG.email}</p>
      <p style="margin-top: 15px; font-size: 12px;">
        © ${new Date().getFullYear()} Kings Equestrian. All rights reserved.
      </p>
    </div>
  </div>
</body>
</html>
  `;
}

// ============== UTILITY FUNCTIONS ==============

function showEmailUsage() {
  const usage = getEmailUsageToday();
  SpreadsheetApp.getUi().alert(
    'Email Usage',
    `Sent: ${usage}/${CONFIG.DAILY_EMAIL_LIMIT}\n\n${canSendEmailToday() ? '✅ Can send more' : '❌ Limit reached'}`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function testEmailTemplate() {
  const result = sendWelcomeEmail('jyothikondupally@gmail.com', 'Test User', 'bangalore');
  
  if (result.success) {
    SpreadsheetApp.getUi().alert('✅ Test email sent!');
  } else {
    SpreadsheetApp.getUi().alert('❌ Failed:\n\n' + result.error);
  }
}