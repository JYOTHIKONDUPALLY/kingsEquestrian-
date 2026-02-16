/**
 * Kings Equestrian - Complete Email System with Consent & Payment Tracking
 * FIXES:
 * - Added missing doPost() function
 * - T&C content fetched from Google Docs
 * - Proper error handling for consent submission
 */

// ============== CONFIGURATION ==============
const CONFIG = {
  // Email settings
  emailSubject: 'Welcome {name} to Kings Equestrian 🏇',
  
  // Google Docs ID for email content
  emailTemplateDocId: '1qALQ8RlVlMrpEhjLYjaf3txDUo5BZLoHjOoRTrS191g',
  
  // Google Docs ID for Terms & Conditions (ADD THIS!)
  termsConditionsDocId: '1ePc_Rb62vcRRN5Z8SybaJ-MEZlBGQnwDiJ9qNstzs4c', // ← UPDATE THIS
  
  // Web App URL (auto-filled after deployment)
  // webAppUrl: 'https://script.google.com/macros/s/AKfycbwI0YjkNLMpOEF9Tww_4i7-A4Gx1Bgd5TihrfC9Y-chW1ORb3pnw-xDKFJ1pBrwjVi6/exec',
  webAppUrl:'https://script.google.com/macros/s/AKfycbwv2BAYZI9qblkaWc8kT-3yzeAw17t26kAN_ogxILqrZUq0TBw40ITsysXYIBdY3KQG/exec',
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
  
  // Default payment amount
  defaultPaymentAmount: 950000,
  
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
  
  // Fetch T&C content from Google Docs
  const termsContent = getTermsFromDocs();
  
  // Serve the consent form
  const template = HtmlService.createTemplateFromFile('ConsentPage');
  template.email = email;
  template.name = name;
  template.location = location;
  template.colors = CONFIG.colors;
  template.termsHtml = termsContent.html;
  
  return template.evaluate()
    .setTitle('Terms & Conditions - Kings Equestrian')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * ✅ FIXED: Added missing doPost() function
 * Handles POST requests - processes consent acceptance
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const result = submitConsentToServer(data);
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}



/**
 * ✅ NEW: Fetch Terms & Conditions from Google Docs
 */
function getTermsFromDocs() {
  try {
    const doc = DocumentApp.openById(CONFIG.termsConditionsDocId);
    const body = doc.getBody();
    const text = body.getText();
    
    // Convert to HTML with proper formatting
    const html = formatTermsToHTML(text);
    
    return {
      success: true,
      html: html,
      text: text
    };
  } catch (error) {
    Logger.log('Error fetching T&C from Docs: ' + error.toString());
    
    // Fallback to default terms
    return {
      success: false,
      html: getDefaultTermsHTML(),
      text: 'Error loading terms'
    };
  }
}

/**
 * ✅ NEW: Format Terms & Conditions text to HTML
 */
function formatTermsToHTML(text) {
  // Split by numbered sections (1. 2. 3. etc.)
  const sections = text.split(/(?=\d+\.\s)/);
  
  let html = '';
  
  sections.forEach(section => {
    if (section.trim()) {
      // Check if it's a numbered section
      const match = section.match(/^(\d+)\.\s+(.+?)$/m);
      
      if (match) {
        const sectionNumber = match[1];
        const sectionTitle = match[2];
        const content = section.replace(match[0], '').trim();
        
        html += `
          <h3>${sectionNumber}. ${sectionTitle}</h3>
          <p>${content.replace(/\n/g, '<br>')}</p>
        `;
      } else {
        // Regular paragraph
        html += `<p>${section.trim().replace(/\n/g, '<br>')}</p>`;
      }
    }
  });
  
  return html;
}

/**
 * ✅ NEW: Default terms if Google Docs fails
 */
function getDefaultTermsHTML() {
  return `
    <h3>1. Registration & Membership</h3>
    <ul>
      <li>Registration is valid for one year from payment date</li>
      <li>Membership is non-transferable</li>
      <li>All information must be accurate</li>
      <li>Registration fees are non-refundable</li>
    </ul>
    
    <h3>2. Safety & Liability</h3>
    <ul>
      <li>Follow all safety instructions</li>
      <li>Proper riding gear is mandatory</li>
      <li>Kings Equestrian is not liable for injuries</li>
      <li>Participants engage at their own risk</li>
    </ul>
    
    <h3>3. Payment Terms</h3>
    <ul>
      <li>Payment must be completed within 7 days</li>
      <li>Session packages paid in full before commencement</li>
      <li>No refunds for unused sessions</li>
    </ul>
    
    <p style="margin-top: 20px; font-weight: bold;">
      By accepting, you acknowledge that you have read and agree to these terms.
    </p>
  `;
}

// ============== EXISTING FUNCTIONS (Keep all your existing code) ==============

function getExistingRegistrationByEmail(email) {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('Mainsheet');

  if (!sheet) {
    return { found: false };
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const emailColIndex = headers.indexOf('Email ID');

  if (emailColIndex === -1) {
    return { found: false };
  }

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowEmail = row[emailColIndex];

    if (rowEmail && rowEmail.toString().trim().toLowerCase() === email.trim().toLowerCase()) {
      return {
        found: true,
        registrationNumber: row[headers.indexOf('Registration Number')],
        name: row[headers.indexOf('Student Name')],
        amount: row[headers.indexOf('Amount to be Paid')],
      };
    }
  }

  return { found: false };
}

function submitConsentToServer(data) {
  try {
    if (!data.consentAccepted) {
      return { success: false, message: 'Consent must be accepted' };
    }

    // 🔍 Check if already registered
    const existing = getExistingRegistrationByEmail(data.email);

    if (existing.found) {
      return {
        success: false,
        alreadyRegistered: true,
        message: 'You are already registered with this email.',
        details: {
          registrationNumber: existing.registrationNumber,
          name: existing.name,
          amount: existing.amount,
        }
      };
    }

    // 🆕 New registration
    const registrationNumber = generateRegistrationNumber(data.location);
    const paymentAmount = CONFIG.defaultPaymentAmount;
    const parentName ='';
    const contact='';

    saveToMainSheet(
      registrationNumber,
      paymentAmount,
      data.email,
      data.name,
      data.location
    );
     updateMainSheetWithConsent(
      data.email,
      registrationNumber,
      data.consentAccepted
    );
     saveToPaymentDetails(
      registrationNumber,
      paymentAmount,
      data.email,
      data.name,
      data.location
    );
      try {
    const enquireSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Enquire Response');
    if (enquireSheet) {
      const enquireData = enquireSheet.getDataRange().getValues();
      const headers = enquireData[0];
      
      // Find column indices
      const regNumCol = headers.indexOf('Registration Number') || headers.indexOf('Reg#') || 0;
      const parentCol = headers.indexOf('Parent Name');
      const phoneCol = headers.indexOf('Phone Number');
      
      // Find matching row
      for (let i = 1; i < enquireData.length; i++) {
        if (enquireData[i][regNumCol]?.toString().trim() === registrationNumber.toString().trim()) {
          parentName = enquireData[i][parentCol] || parentName;
          contact = enquireData[i][phoneCol] || contact;
          break;
        }
      }
    }
  } catch (e) {
    Logger.log('Enquire lookup failed: ' + e);
  }
  Logger.log('=== DEBUG START ===');
Logger.log('Full data object: ' + JSON.stringify(data));
Logger.log('data.email type: ' + typeof data.email);
Logger.log('data.email value: "' + data.email + '"');
Logger.log('data.email is null: ' + (data.email === null));
Logger.log('data.email is undefined: ' + (data.email === undefined));
Logger.log('registrationNumber: ' + registrationNumber);
Logger.log('parentName: ' + parentName);
Logger.log('contact: ' + contact);
Logger.log('=== DEBUG END ===');

sendPaymentRequestEmail({
  registrationNumber: registrationNumber,
  email: data.email,
  name: data.name,
  parentName: parentName,
  contact: contact,
  location: data.location,
  amount: paymentAmount,
  consentDate: data.consentDate
});

 

    return {
      success: true,
      registrationNumber,
      paymentAmount
    };

  } catch (err) {
    return { success: false, message: err.toString() };
  }
}


function saveToPaymentDetails(registrationNumber, amount, email, name, location) {
    const paymentSheet = getPaymentDetailsSheet();
    const rowData = [
    registrationNumber,
    amount,
    email,
    name,
    location || '',
    '',
    '',
    'Pending'
  ];
  paymentSheet.appendRow(rowData);
    
  Logger.log(`Payment details saved for ${registrationNumber}`);
}

function generateRegistrationNumber(location) {
  const EnquireSheet = getEnquireDetailsSheet();
  
  const locationKey = Object.keys(CONFIG.locationCodes).find(key => 
    location.toLowerCase().includes(key) || key.includes(location.toLowerCase())
  );
  const locationCode = locationKey ? CONFIG.locationCodes[locationKey] : 'GEN';
  
  const lastRow = EnquireSheet.getLastRow();
  let maxNumber = 0;
  
  if (lastRow > 1) {
    const existingNumbers = EnquireSheet.getRange(2, 1, lastRow - 1, 1).getValues();
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
      'Payment Status',
      'Send Payment Link',
      'Sent Timestamps',
      'Payment Status'
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

function getEnquireDetailsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let EnquireSheet = ss.getSheetByName('Enquire Response');
  
  if (!EnquireSheet) {
    EnquireSheet = ss.insertSheet('Enquire Response');
    
    const headers = [
      'Timestamp',
      'Email address',
      'Student Name',
      'Grade & Section',
      'Parent Name',
      'Email ID',
      'Phone Number',
      'Location',
      'Email Sent',
      'Email Sent Timestamps',
      'Error Message'
    ];
    
   EnquireSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    EnquireSheet.getRange(1, 1, 1, headers.length)
      .setBackground(CONFIG.colors.primary)
      .setFontColor('white')
      .setFontWeight('bold');
    
    EnquireSheet.setFrozenRows(1);
    EnquireSheet.autoResizeColumns(1, headers.length);
  }
  
  return EnquireSheet;
}

function saveToMainSheet(registrationNumber, amount, email, name, location) {
   const ss = SpreadsheetApp.getActiveSpreadsheet();
  const enquireSheet = ss.getSheetByName('Enquire Response');
  const mainSheet = ss.getSheetByName('Mainsheet');
    if (!enquireSheet || !mainSheet) {
    throw new Error('Required sheets not found');
  }
  const enquireHeaders = getHeaderIndexMap(enquireSheet);
  const mainHeaders = getHeaderIndexMap(mainSheet);

  const enquireValues = enquireSheet.getDataRange().getValues();

  // 🔍 Find matching enquiry row
  let enquiryRow = null;

  for (let i = 1; i < enquireValues.length; i++) {
    const row = enquireValues[i];

    if (
      row[enquireHeaders['Email ID']]?.toString().toLowerCase() === email.toLowerCase() &&
      row[enquireHeaders['Location']] === location &&
      row[enquireHeaders['Student Name']] === name
    ) {
      enquiryRow = row;
      break;
    }
  }

  if (!enquiryRow) {
    throw new Error(`No enquiry found for ${email} / ${location} / ${name}`);
  }

  // 🧾 Build Mainsheet row (based on column names)
  const newRow = Array(mainSheet.getLastColumn()).fill('');

  newRow[mainHeaders['Registration Number']] = registrationNumber;
  newRow[mainHeaders['Student Name']] = enquiryRow[enquireHeaders['Student Name']];
  newRow[mainHeaders['Grade & Section']] = enquiryRow[enquireHeaders['Grade & Section']];
  newRow[mainHeaders['Parent Name']] = enquiryRow[enquireHeaders['Parent Name']];
  newRow[mainHeaders['Email ID']] = enquiryRow[enquireHeaders['Email ID']];
  newRow[mainHeaders['Phone Number']] = enquiryRow[enquireHeaders['Phone Number']];
  newRow[mainHeaders['Location']] = enquiryRow[enquireHeaders['Location']];

  newRow[mainHeaders['Consent Accepted']] = '';
  newRow[mainHeaders['Consent Timestamp']] = '';
  newRow[mainHeaders['Program Selected']] = '';
  newRow[mainHeaders['Amount To be Paid']] = amount;
  newRow[mainHeaders['Total Amount Paid']] = '';
  newRow[mainHeaders['PAN/AAdhar']] = '';

  // ➕ Append to Mainsheet
  mainSheet.appendRow(newRow);
  
}


function updateMainSheetWithConsent(email, registrationNumber, consentAccepted) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Mainsheet');
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
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
    
    const emailCol = headers.indexOf('Email ID') + 1;
    const dataRange = sheet.getRange(2, emailCol, sheet.getLastRow() - 1, 1);
    const emails = dataRange.getValues();
    const registerationNumbers = sheet.getRange(2, regCol, sheet.getLastRow() - 1, 1).getValues();
    
    for (let i = 0; i < emails.length; i++) {
      if (emails[i][0] === email && registerationNumbers[i][0] === registrationNumber) {
        const row = i + 2;
        sheet.getRange(row, consentCol).setValue(consentAccepted ? 'Yes' : 'No');
        sheet.getRange(row, consentTimeCol).setValue(new Date()).setNumberFormat("dd-MMM-yyyy HH:mm:ss");
        
        break;
      }
    }
    
  } catch (error) {
    Logger.log('Error updating main sheet: ' + error.toString());
  }
}

function getDocAsPDF() {
  try {
    const docId = CONFIG.emailTemplateDocId;
    const doc = DriveApp.getFileById(docId);
    const pdfBlob = doc.getAs('application/pdf');
    pdfBlob.setName('Kings_Equestrian_Information.pdf');
    return pdfBlob;
  } catch (error) {
    Logger.log('Error creating PDF: ' + error.toString());
    return null;
  }
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📧 Kings Equestrian')
    .addItem('🔄 Resend Selected Emails', 'resendSelectedEmails')
    .addItem('📊 View Email Usage', 'showEmailUsage')
    .addItem('⚙️ Process Email Queue', 'processEmailQueue')
    .addItem('🧪 Test Email Template', 'testEmailTemplate')
    .addItem('🔗 Get Consent Page URL', 'showConsentPageURL')
    .addItem('🔄 RESEND PAYMENT EMAIL', 'resendPaymentEmail')
    .addItem('📧SEND RECEIPT', 'SendPaymentReceipt')
    .addToUi();
}
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

function showConsentPageURL() {
  const ui = SpreadsheetApp.getUi();
  const webAppUrl = getWebAppUrl();
  const testUrl = `${getWebAppUrl()}?email=test@example.com&name=Test%20User&location=bangalore`;
  
  ui.alert(
    'Consent Page URL',
    `Base URL: ${webAppUrl}\n\nTest URL:\n${testUrl}\n\nUse this URL in your emails.`,
    ui.ButtonSet.OK
  );
}

function getWebAppUrl() {
  // return ScriptApp.getService().getUrl();
  return CONFIG.webAppUrl;
}

// [Keep all other existing functions: email quota, docs fetcher, queue, etc.]
// I'm not repeating them here to save space - they remain the same

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
function createEmptyBlock() {
  return { title: '', subtitle: '', content: [], list: [] };
}

function flushBlock(block) {
  if (!block) return '';

  return `
    <div style="
      margin: 25px 0;
      padding: 20px;
      background: ${CONFIG.colors.background};
      border-radius: 8px;
      border-left: 4px solid ${CONFIG.colors.secondary};
    ">
      ${block.title ? `<h3 style="color:${CONFIG.colors.primary};margin-top:0;">${block.title}</h3>` : ''}
      ${block.subtitle ? `<h4 style="margin:5px 0;color:${CONFIG.colors.secondary};">${block.subtitle}</h4>` : ''}
      ${block.content.length ? `<p style="line-height:1.6;color:${CONFIG.colors.text};">${block.content.join('<br>')}</p>` : ''}
      ${block.list.length ? `<ul>${block.list.map(li => `<li>${li}</li>`).join('')}</ul>` : ''}
    </div>
  `;
}



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
  }
}

function getHeaderIndexMap(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};

  headers.forEach((header, index) => {
    if (header) {
      map[header.toString().trim()] = index;
    }
  });

  return map; // 0-based indexes
}

function sendWelcomeEmail(email, name, location) {
  let status = 'Sent';
  let errorMsg = '';
  const timestamp = new Date();

  try {
    if (!email || !isValidEmail(email)) {
      throw new Error('Invalid email: ' + email);
    }

    const docsContent = getEmailContentFromDocs();
    if (!docsContent.success) {
      throw new Error('Failed to fetch content: ' + docsContent.error);
    }

    const consentUrl =
      `${getWebAppUrl()}?email=${encodeURIComponent(email)}&name=${encodeURIComponent(name)}&location=${encodeURIComponent(location)}`;

    const htmlBody = createEmailTemplate(name, location, docsContent.html, consentUrl);
    const subject = CONFIG.emailSubject.replace('{name}', name);

    const pdfBlob = getDocAsPDF();

    const options = {
      htmlBody,
      name: 'Kings Equestrian',
      ...(pdfBlob && { attachments: [pdfBlob] })
    };

    MailApp.sendEmail(email, subject, 'Please view in HTML client.', options);

  } catch (err) {
    status = 'Failed';
    errorMsg = err.toString();
  }

  // 🔁 Update sheet using header names
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Enquire Response');
    if (!sheet) throw new Error('Sheet "Enquire Response" not found');

    const headerMap = getHeaderIndexMap(sheet);
    const values = sheet.getDataRange().getValues();

    for (let i = 1; i < values.length; i++) {
      const rowEmail = values[i][headerMap['Email']]?.toString().toLowerCase();

      if (rowEmail === email.toLowerCase()) {

        if (headerMap['Email Status'] !== undefined && headerMap['Email Status'] !== "") {
          sheet.getRange(i + 1, headerMap['Email Status'] + 1).setValue(status);
        }

        if (headerMap['Email Sent Timestamp'] !== undefined && headerMap['Email Sent Timestamp'] !== "") {
          sheet.getRange(i + 1, headerMap['Email Sent Timestamp'] + 1)
            .setValue(timestamp)
            .setNumberFormat("dd-MMM-yyyy HH:mm:ss");
        }

        if (status === 'Failed' && headerMap['Error Message'] !== undefined) {
          sheet.getRange(i + 1, headerMap['Error Message'] + 1)
            .setValue(errorMsg.substring(0, 200));
        }

        Logger.log(`✅ Sheet updated - ${email}: ${status}`);
        break;
      }
    }

  } catch (sheetErr) {
    Logger.log('❌ Sheet update failed: ' + sheetErr.toString());
  }

  return { success: status === 'Sent', status, error: errorMsg };
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
    
    .footer {
      background-color: ${CONFIG.colors.primary};
      color: white;
      padding: 30px;
      text-align: center;
      font-size: 14px;
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

      <div class=""style="border: 4px solid ${CONFIG.colors.secondary}; padding:30px;" >
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

      <p style="margin-top: 30px; font-style: italic; color: #666;">
        Come for the ride. Leave with a feeling that stays for life. 🤍🐎✨
      </p>
    </div>

    <div class="footer">
      <p style="font-size: 16px; font-weight: 600;">KINGS EQUESTRIAN</p>
      <p>📞 ${CONFIG.whatsappNumbers} | ✉️ ${CONFIG.email}</p>
      <p style="margin-top: 15px; font-size: 12px;">
        © ${new Date().getFullYear()} Kings Equestrian. All rights reserved.
      </p>
      <p style="margin: 10px 0;">
          <strong>📱 WhatsApp:</strong> ${CONFIG.whatsappNumbers}<br>
          <strong>📩 Instagram:</strong> ${CONFIG.instagramLink}<br>
        </p>
        <p style="margin: 10px 0;">
          <strong>🌐 Website:</strong> <a href="${CONFIG.websiteLink}">${CONFIG.websiteLink}</a><br>
          <strong>⭐ Reviews:</strong> <a href="${CONFIG.reviewsLink}">Read experiences</a>
        </p>
    </div>
  </div>
</body>
</html>
  `;
}

function isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function updateEmailStatusInMainSheet(email, status, errorMessage = '') {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Enquire Response');
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
function testEmailTemplate() {
  const result = sendWelcomeEmail('jyothikondupally@gmail.com', 'Test User', 'bangalore');
  
  if (result.success) {
    SpreadsheetApp.getUi().alert('✅ Test email sent!');
  } else {
    SpreadsheetApp.getUi().alert('❌ Failed:\n\n' + result.error);
  }
}

function testConsentPage() {
  const testUrl = `${getWebAppUrl()}?email=jyothikondupally@gmail.com&name=Jyothi&location=hyderabad`;
  console.log('Test this URL: ' + testUrl);
}