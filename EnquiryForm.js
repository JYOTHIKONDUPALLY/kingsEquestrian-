/**
 * Kings Equestrian - Advanced Email Queue System
 * Features:
 * - Fetches email content from Google Docs
 * - No registration number generation
 * - Detailed error logging
 * - UI button to resend failed emails
 * - Automatic queue processing
 */

// ============== CONFIGURATION ==============
const CONFIG = {
  // Email settings
  emailSubject: 'Welcome {name} to Kings Equestrian 🏇',
  
  // Google Docs ID for email content (extract from URL)
  emailTemplateDocId: '1qALQ8RlVlMrpEhjLYjaf3txDUo5BZLoHjOoRTrS191g',
  
  // Links
  consentFormLink: 'https://forms.gle/SRfZmVsc3qHNJf3i7',
  websiteLink: 'https://kingsfarmequestrian.com',
  instagramLink: '@kingsequestrianfoundation',
  reviewsLink: 'https://maps.app.goo.gl/EVyzEfhh3tdJ2BTX7?g_st=iwb',
  
  // Contact
  whatsappNumbers: '9980771166 | 9980895533',
  email: 'info@kingsequestrian.com',
  
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

// ============== ADD CUSTOM MENU ==============

/**
 * Creates custom menu when spreadsheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📧 Kings Equestrian')
    .addItem('🔄 Resend Selected Emails', 'resendSelectedEmails')
    .addItem('📊 View Email Usage', 'showEmailUsage')
    .addItem('⚙️ Process Email Queue', 'processEmailQueue')
    .addItem('🧪 Test Email Template', 'testEmailTemplate')
    .addToUi();
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

/**
 * Fetches and formats content from Google Docs
 */
function getEmailContentFromDocs() {
  try {
    const doc = DocumentApp.openById(CONFIG.emailTemplateDocId);
    const body = doc.getBody();
    const text = body.getText();
    
    // Return the raw text from the document
    // You can format it as needed
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

/**
 * Converts plain text from Docs to formatted HTML
 */
function formatDocsContentToHTML(text) {
  // Split by sections using emojis as markers
  const sections = text.split(/(?=🤍|✨|🌿|📸|🌄|🎉)/);
  
  let html = '';
  
  sections.forEach(section => {
    if (section.trim()) {
      // Check for section titles (with emojis)
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
        // Regular paragraph
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
  
  Logger.log(`Email queued for ${email} - ${errorMessage || 'Will be sent tomorrow'}`);
  
  // Update main sheet status
  updateEmailStatusInMainSheet(email, errorMessage ? 'Failed - Queued' : 'Queued for Tomorrow', errorMessage);
}

/**
 * Update email status in main registration sheet
 */
function updateEmailStatusInMainSheet(email, status, errorMessage = '') {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheets()[0];
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Email Status column
    let statusCol = headers.indexOf('Email Status') + 1;
    if (statusCol === 0) {
      statusCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, statusCol)
        .setValue('Email Status')
        .setBackground(CONFIG.colors.primary)
        .setFontColor('white')
        .setFontWeight('bold');
    }

    // Email Timestamp column
    let timestampCol = headers.indexOf('Email Sent Timestamp') + 1;
    if (timestampCol === 0) {
      timestampCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, timestampCol)
        .setValue('Email Sent Timestamp')
        .setBackground(CONFIG.colors.primary)
        .setFontColor('white')
        .setFontWeight('bold');
    }
    
    // Error Message column
    let errorCol = headers.indexOf('Error Message') + 1;
    if (errorCol === 0) {
      errorCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, errorCol)
        .setValue('Error Message')
        .setBackground(CONFIG.colors.primary)
        .setFontColor('white')
        .setFontWeight('bold');
    }

    // Find Email column
    const emailCol = headers.indexOf('Email ID') + 1;
    const dataRange = sheet.getRange(2, emailCol, sheet.getLastRow() - 1, 1);
    const emails = dataRange.getValues();

    for (let i = 0; i < emails.length; i++) {
      if (emails[i][0] === email) {
        const row = i + 2;

        // Set status
        const statusCell = sheet.getRange(row, statusCol);
        statusCell.setValue(status);

        // Set timestamp
        const timestampCell = sheet.getRange(row, timestampCol);
        timestampCell.setValue(new Date())
          .setNumberFormat("dd-MMM-yyyy HH:mm:ss");
        
        // Set error message if exists
        if (errorMessage) {
          sheet.getRange(row, errorCol).setValue(errorMessage);
        }

        // Color code status
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

/**
 * Process pending emails from queue
 */
function processEmailQueue() {
  const queueSheet = getEmailQueueSheet();
  const lastRow = queueSheet.getLastRow();
  
  if (lastRow <= 1) {
    SpreadsheetApp.getUi().alert('Email queue is empty');
    Logger.log('Email queue is empty');
    return;
  }
  
  const data = queueSheet.getRange(2, 1, lastRow - 1, 8).getValues();
  const rowsToDelete = [];
  let emailsSentCount = 0;
  
  Logger.log(`Processing ${data.length} queued emails...`);
  
  for (let i = 0; i < data.length; i++) {
    if (!canSendEmailToday()) {
      Logger.log(`Daily email limit reached. ${data.length - i} emails remain in queue.`);
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
      Logger.log(`❌ Failed to send email to ${email}: ${errorMsg}`);
      
      queueSheet.getRange(i + 2, 5).setValue('Failed');
      queueSheet.getRange(i + 2, 6).setValue(attempts + 1);
      queueSheet.getRange(i + 2, 7).setValue(new Date());
      queueSheet.getRange(i + 2, 8).setValue(errorMsg.substring(0, 200));
      queueSheet.getRange(i + 2, 5).setBackground('#f8d7da').setFontColor('#721c24');
      
      updateEmailStatusInMainSheet(email, 'Failed - Check Queue', errorMsg);
    }
  }
  
  // Delete sent emails from queue
  rowsToDelete.reverse().forEach(row => {
    queueSheet.deleteRow(row);
  });
  
  Logger.log(`Queue processing complete. Sent: ${emailsSentCount}, Remaining: ${queueSheet.getLastRow() - 1}`);
  
  SpreadsheetApp.getUi().alert(`Queue processed!\n\nEmails sent: ${emailsSentCount}\nRemaining in queue: ${queueSheet.getLastRow() - 1}`);
}

// ============== RESEND SELECTED EMAILS ==============

/**
 * Resend emails for selected rows in the sheet
 */
function resendSelectedEmails() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const selection = sheet.getActiveRange();
  
  if (!selection) {
    ui.alert('Please select rows to resend emails');
    return;
  }
  
  const response = ui.alert(
    'Resend Emails',
    `Are you sure you want to resend emails for ${selection.getNumRows()} selected row(s)?`,
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  const startRow = selection.getRow();
  const numRows = selection.getNumRows();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  const emailCol = headers.indexOf('Email ID') + 1;
  const nameCol = headers.indexOf('Student Name') + 1;
  const locationCol = headers.indexOf('Location') + 1;
  
  if (emailCol === 0 || nameCol === 0) {
    ui.alert('Error: Required columns (Email ID, Student Name) not found');
    return;
  }
  
  let successCount = 0;
  let failCount = 0;
  const errors = [];
  
  for (let i = 0; i < numRows; i++) {
    const row = startRow + i;
    
    // Skip header row
    if (row === 1) continue;
    
    const email = sheet.getRange(row, emailCol).getValue();
    const name = sheet.getRange(row, nameCol).getValue();
    const location = locationCol > 0 ? sheet.getRange(row, locationCol).getValue() : '';
    
    if (!email || !isValidEmail(email)) {
      errors.push(`Row ${row}: Invalid or missing email`);
      failCount++;
      continue;
    }
    
    if (!canSendEmailToday()) {
      ui.alert(`Daily email limit reached!\n\nSent: ${successCount}\nFailed: ${failCount}\n\nRemaining emails added to queue.`);
      
      // Add remaining to queue
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
        Logger.log(`✅ Resent email to ${email}`);
      } else {
        throw new Error(result.error);
      }
      
      Utilities.sleep(1000);
      
    } catch (error) {
      const errorMsg = error.toString();
      errors.push(`Row ${row} (${email}): ${errorMsg}`);
      updateEmailStatusInMainSheet(email, 'Failed', errorMsg);
      failCount++;
      Logger.log(`❌ Failed to resend to ${email}: ${errorMsg}`);
    }
  }
  
  // Show summary
  let message = `Resend Complete!\n\n✅ Sent: ${successCount}\n❌ Failed: ${failCount}`;
  
  if (errors.length > 0 && errors.length <= 5) {
    message += '\n\nErrors:\n' + errors.join('\n');
  } else if (errors.length > 5) {
    message += '\n\nCheck Error Message column for details.';
  }
  
  ui.alert(message);
}

/**
 * Validate email format
 */
function isValidEmail(email) {
  const emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailPattern.test(email);
}

// ============== MAIN TRIGGER FUNCTION ==============

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
    
    // Validate email
    if (!email || !isValidEmail(email)) {
      const error = 'Invalid email format: ' + email;
      updateEmailStatusInMainSheet(email, 'Failed', error);
      addToEmailQueue(email, name, location, error);
      Logger.log('❌ ' + error);
      return;
    }
    
    if (canSendEmailToday()) {
      const result = sendWelcomeEmail(email, name, location);
      
      if (result.success) {
        incrementEmailUsage();
        updateEmailStatusInMainSheet(email, 'Sent');
        
        const currentUsage = getEmailUsageToday();
        Logger.log(`✅ Email sent to ${email}. Daily count: ${currentUsage}/${CONFIG.DAILY_EMAIL_LIMIT}`);
      } else {
        throw new Error(result.error);
      }
    } else {
      addToEmailQueue(email, name, location);
      Logger.log(`Daily limit reached. Email queued for ${email}`);
    }
    
  } catch (error) {
    const errorMsg = error.toString();
    Logger.log('❌ Error in onFormSubmit: ' + errorMsg);
    
    try {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
      const row = e.range.getRow();
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const emailCol = headers.indexOf('Email ID') + 1;
      const email = sheet.getRange(row, emailCol).getValue();
      
      updateEmailStatusInMainSheet(email, 'Failed', errorMsg);
      addToEmailQueue(email, name || 'Unknown', location || '', errorMsg);
    } catch (e) {
      Logger.log('Could not update error status: ' + e.toString());
    }
    
    sendAdminNotification(errorMsg);
  }
}

// ============== EMAIL TEMPLATE ==============

function sendWelcomeEmail(email, name, location) {
  try {
    // Validate email first
    if (!email || !isValidEmail(email)) {
      return {
        success: false,
        error: 'Invalid email address: ' + email
      };
    }
    
    // Get content from Google Docs
    const docsContent = getEmailContentFromDocs();
    
    if (!docsContent.success) {
      return {
        success: false,
        error: 'Failed to fetch content from Google Docs: ' + docsContent.error
      };
    }
    
    const htmlBody = createEmailTemplate(name, location, docsContent.html);
    
    const subject = CONFIG.emailSubject.replace('{name}', name);
    
    const options = {
      htmlBody: htmlBody,
      name: 'Kings Equestrian',
    };
    
    MailApp.sendEmail(
      email,
      subject,
      'Please view this email in an HTML-enabled email client.',
      options
    );
    
    return {
      success: true,
      message: 'Email sent successfully to ' + email
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

function createEmailTemplate(name, location, docsHtmlContent) {
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
      letter-spacing: 0.5px;
    }
    .crown-icon img {
      width: 70px;
      height: 70px;
      border-radius: 50%;
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
    .docs-content {
      margin: 30px 0;
      line-height: 1.8;
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
    .footer p { margin: 5px 0; }
    @media only screen and (max-width: 600px) {
      .container { margin: 10px; }
      .content { padding: 25px 20px; }
      .header { padding: 30px 20px; }
    }
  </style>
</head>
<body>
  <div class="container">
    <!-- Header -->
    <div class="header">
      <div class="crown-icon">
        <img src="https://kingsfarmequestrian.com/wp-content/uploads/2023/08/Logo2.jpg" alt="Kings Equestrian Logo"/>
      </div>
      <h1>KINGS EQUESTRIAN</h1>
      <p style="margin: 10px 0 0 0; font-size: 16px; opacity: 0.95;">
        Where horses don't just carry you — they change you. 🤍🐎
      </p>
    </div>

    <!-- Main Content -->
    <div class="content">
      <div class="greeting">Dear ${name},</div>
      
      <p>Welcome to <strong>Kings Equestrian</strong>! We're thrilled to have you join our equestrian family.</p>
      
      <!-- Dynamic content from Google Docs -->
      <div class="docs-content">
        ${docsHtmlContent}
      </div>

      <!-- Terms & Conditions -->
      <div class="section-box">
        <h3 style="color: ${CONFIG.colors.primary}; margin-top: 0;">
          ✅ Terms & Conditions Acceptance
        </h3>
        <p style="margin: 15px 0;">
          To proceed further and receive the payment request, please review and accept our Terms & Conditions.
        </p>
        <div style="text-align: center;">
          <a href="${CONFIG.consentFormLink}" class="cta-button" style="color: white;" target="_blank">
            Review & Accept Terms & Conditions
          </a>
        </div>
        <p style="margin-top: 15px; font-size: 12px; color: #777; text-align: center;">
          Payment request will be shared after terms acceptance.
        </p>
      </div>

      <!-- Contact & Social Links -->
      <div class="section-box">
        <h3 style="color: ${CONFIG.colors.primary}; margin-top: 0;">📞 Get In Touch</h3>
        <p style="margin: 10px 0;">
          <strong>📱 WhatsApp:</strong> ${CONFIG.whatsappNumbers}<br>
          <strong>📩 Instagram:</strong> ${CONFIG.instagramLink}<br>
          <strong>🌐 Website:</strong> <a href="${CONFIG.websiteLink}" style="color: ${CONFIG.colors.primary};">${CONFIG.websiteLink}</a><br>
          <strong>⭐ Reviews:</strong> <a href="${CONFIG.reviewsLink}" style="color: ${CONFIG.colors.primary};">Read customer experiences</a>
        </p>
      </div>

      <p style="margin-top: 30px; font-style: italic; color: #666;">
        Come for the ride. Leave with a feeling that stays for life. 🤍🐎✨
      </p>
      
      <p style="margin-top: 20px;">
        <strong>Ride with Pride!</strong><br>
        <span style="color: ${CONFIG.colors.secondary};">The Kings Equestrian Team</span>
      </p>
    </div>

    <!-- Footer -->
    <div class="footer">
      <p style="font-size: 16px; font-weight: 600; margin-bottom: 10px;">KINGS EQUESTRIAN</p>
      <p>📍 ${locationName}</p>
      <p>📞 ${CONFIG.whatsappNumbers} | ✉️ ${CONFIG.email}</p>
      <p style="margin-top: 15px; font-size: 12px; opacity: 0.9;">
        © ${new Date().getFullYear()} Kings Equestrian. All rights reserved.
      </p>
      <div style="margin-top: 15px;">
        <p style="font-size: 12px;">Follow us: ${CONFIG.instagramLink}</p>
      </div>
    </div>
  </div>
</body>
</html>
  `;
}

// ============== ADMIN & UTILITY FUNCTIONS ==============

function sendAdminNotification(error) {
  const adminEmail = 'jyothikondupally@gmail.com';
  
  try {
    MailApp.sendEmail(
      adminEmail,
      'Kings Equestrian - Email System Error',
      `An error occurred:\n\n${error}\n\nTime: ${new Date()}`
    );
  } catch (e) {
    Logger.log('Could not send admin notification: ' + e.toString());
  }
}

function showEmailUsage() {
  const usage = getEmailUsageToday();
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Email Usage Today',
    `Emails sent: ${usage} / ${CONFIG.DAILY_EMAIL_LIMIT}\n\n${canSendEmailToday() ? '✅ Can send more emails' : '❌ Daily limit reached'}`,
    ui.ButtonSet.OK
  );
}

// ============== TEST FUNCTIONS ==============

function testEmailTemplate() {
  const testEmail = 'jyothikondupally@gmail.com';
  const testName = 'John Doe';
  const testLocation = 'bangalore';
  
  const result = sendWelcomeEmail(testEmail, testName, testLocation);
  
  if (result.success) {
    Logger.log('✅ Test email sent successfully to: ' + testEmail);
    SpreadsheetApp.getUi().alert('✅ Test email sent successfully to: ' + testEmail);
  } else {
    Logger.log('❌ Test email failed: ' + result.error);
    SpreadsheetApp.getUi().alert('❌ Test email failed:\n\n' + result.error);
  }
}

function testDocsContent() {
  const content = getEmailContentFromDocs();
  
  if (content.success) {
    Logger.log('✅ Successfully fetched content from Google Docs');
    Logger.log('Content length: ' + content.content.length + ' characters');
    SpreadsheetApp.getUi().alert('✅ Successfully fetched content from Google Docs\n\nContent length: ' + content.content.length + ' characters');
  } else {
    Logger.log('❌ Failed to fetch content: ' + content.error);
    SpreadsheetApp.getUi().alert('❌ Failed to fetch content:\n\n' + content.error);
  }
}