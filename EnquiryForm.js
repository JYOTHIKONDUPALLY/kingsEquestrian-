/**
 * Kings Equestrian - Advanced Email Queue System
 * Handles email limits with automatic retry the next day
 * 
 * FEATURES:
 * - Sends first 100 emails immediately
 * - Queues remaining emails for next day
 * - Automatically processes queue daily
 * - Tracks email status in sheet
 * - Personalized subject with name and registration number
 */

// ============== CONFIGURATION ==============
const CONFIG = {
  locationCodes: {
    'bangalore': 'BLR',
    'hyderabad': 'HYD',
    'pune': 'PNE'
  },
  
  // Email settings - {name} and {regNumber} will be replaced with actual values
  emailSubject: 'Welcome {name} - Registration {regNumber} Confirmed | Kings Equestrian 🏇',
  documentLink: 'https://docs.google.com/document/d/1qALQ8RlVlMrpEhjLYjaf3txDUo5BZLoHjOoRTrS191g/edit?tab=t.0',
  consentFormLink: 'https://forms.gle/SRfZmVsc3qHNJf3i7',

  
  // Email quota management
  DAILY_EMAIL_LIMIT: 95, // Set slightly below 100 to be safe
  
  colors: {
    primary: '#1a472a',
    secondary: '#d4af37',
    accent: '#2d5a3d',
    background: '#f8f9fa',
    text: '#333333'
  }
};

// ============== EMAIL QUOTA TRACKING ==============

/**
 * Get current email usage for today
 */
function getEmailUsageToday() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const today = new Date().toDateString();
  const key = `emailCount_${today}`;
  const count = scriptProperties.getProperty(key);
  return count ? parseInt(count) : 0;
}

/**
 * Increment email counter
 */
function incrementEmailUsage() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const today = new Date().toDateString();
  const key = `emailCount_${today}`;
  const currentCount = getEmailUsageToday();
  scriptProperties.setProperty(key, (currentCount + 1).toString());
}

/**
 * Check if we can send more emails today
 */
function canSendEmailToday() {
  return getEmailUsageToday() < CONFIG.DAILY_EMAIL_LIMIT;
}

/**
 * Reset daily counter (runs automatically at midnight)
 */
function resetDailyEmailCounter() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const yesterday = new Date(Date.now() - 24 * 60 * 60 * 1000).toDateString();
  const oldKey = `emailCount_${yesterday}`;
  scriptProperties.deleteProperty(oldKey);
  Logger.log('Daily email counter reset');
}

// ============== EMAIL QUEUE MANAGEMENT ==============

/**
 * Get or create the Email Queue sheet
 */
function getEmailQueueSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let queueSheet = ss.getSheetByName('Email Queue');
  
  if (!queueSheet) {
    queueSheet = ss.insertSheet('Email Queue');
    
    // Set up headers
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
    
    // Format header
    queueSheet.getRange(1, 1, 1, headers.length)
      .setBackground(CONFIG.colors.primary)
      .setFontColor('white')
      .setFontWeight('bold');
    
    queueSheet.setFrozenRows(1);
    queueSheet.autoResizeColumns(1, headers.length);
  }
  
  return queueSheet;
}

/**
 * Add email to queue when limit is reached
 */
function addToEmailQueue(email, name,  location) {
  const queueSheet = getEmailQueueSheet();
  
  const rowData = [
    new Date(),
    email,
    name,
    location,
    'Pending',
    0,
    '',
    ''
  ];
  
  queueSheet.appendRow(rowData);
  
  Logger.log(`Email queued for ${email} - Will be sent tomorrow`);
  
  // Update main sheet status
  updateEmailStatusInMainSheet(registerNumber, 'Queued for Tomorrow');
}

/**
 * Update email status in main registration sheet
 */
function updateEmailStatusInMainSheet(registerNumber, status) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheets()[0]; // Main sheet (first sheet)
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // --- Email Status column ---
    let statusCol = headers.indexOf('Email Status') + 1;
    if (statusCol === 0) {
      statusCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, statusCol)
        .setValue('Email Status')
        .setBackground(CONFIG.colors.primary)
        .setFontColor('white')
        .setFontWeight('bold');
    }

    // --- Email Timestamp column ---
    let timestampCol = headers.indexOf('Email Sent Timestamp') + 1;
    if (timestampCol === 0) {
      timestampCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, timestampCol)
        .setValue('Email Sent Timestamp')
        .setBackground(CONFIG.colors.primary)
        .setFontColor('white')
        .setFontWeight('bold');
    }

    // Find Register Number column
    const registerCol = headers.indexOf('Register Number') + 1;
    const dataRange = sheet.getRange(2, registerCol, sheet.getLastRow() - 1, 1);
    const registerNumbers = dataRange.getValues();

    for (let i = 0; i < registerNumbers.length; i++) {
      if (registerNumbers[i][0] === registerNumber) {
        const row = i + 2;

        // Set status
        const statusCell = sheet.getRange(row, statusCol);
        statusCell.setValue(status);

        // Set timestamp (date + time)
        const timestampCell = sheet.getRange(row, timestampCol);
        timestampCell.setValue(new Date())
          .setNumberFormat("dd-MMM-yyyy HH:mm:ss");

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
 * Set this to run daily via trigger (recommended: 1 AM)
 */
function processEmailQueue() {
  const queueSheet = getEmailQueueSheet();
  const lastRow = queueSheet.getLastRow();
  
  if (lastRow <= 1) {
    Logger.log('Email queue is empty');
    return;
  }
  
  const data = queueSheet.getRange(2, 1, lastRow - 1, 9).getValues();
  const rowsToDelete = [];
  let emailsSentCount = 0;
  
  Logger.log(`Processing ${data.length} queued emails...`);
  
  for (let i = 0; i < data.length; i++) {
    // Check if we've hit today's limit
    if (!canSendEmailToday()) {
      Logger.log(`Daily email limit reached. ${data.length - i} emails remain in queue.`);
      break;
    }
    
    const [timestamp, email, name, registerNumber, location, status, attempts] = data[i];
    
    // Skip if already sent
    if (status === 'Sent') {
      rowsToDelete.push(i + 2);
      continue;
    }
    
    // Try to send the email
    try {
      sendWelcomeEmail(email, name, registerNumber, location);
      incrementEmailUsage();
      emailsSentCount++;
      
      // Update queue sheet
      queueSheet.getRange(i + 2, 6).setValue('Sent'); // Status
      queueSheet.getRange(i + 2, 7).setValue(attempts + 1); // Attempts
      queueSheet.getRange(i + 2, 8).setValue(new Date()); // Last Attempt
      queueSheet.getRange(i + 2, 6).setBackground('#d4edda').setFontColor('#155724');
      
      // Update main sheet
      updateEmailStatusInMainSheet(registerNumber, 'Sent');
      
      rowsToDelete.push(i + 2);
      
      Logger.log(`Email sent to ${email} (${registerNumber})`);
      
      // Small delay to avoid triggering spam filters
      Utilities.sleep(1000);
      
    } catch (error) {
      Logger.log(`Failed to send email to ${email}: ` + error.toString());
      
      // Update queue with error
      queueSheet.getRange(i + 2, 6).setValue('Failed');
      queueSheet.getRange(i + 2, 7).setValue(attempts + 1);
      queueSheet.getRange(i + 2, 8).setValue(new Date());
      queueSheet.getRange(i + 2, 9).setValue(error.toString().substring(0, 100));
      queueSheet.getRange(i + 2, 6).setBackground('#f8d7da').setFontColor('#721c24');
      
      updateEmailStatusInMainSheet(registerNumber, 'Failed - Check Queue');
    }
  }
  
  // Delete sent emails from queue (in reverse to maintain indices)
  rowsToDelete.reverse().forEach(row => {
    queueSheet.deleteRow(row);
  });
  
  Logger.log(`Queue processing complete. Sent: ${emailsSentCount}, Remaining: ${queueSheet.getLastRow() - 1}`);
  
  // Send summary to admin if needed
  if (emailsSentCount > 0 || (queueSheet.getLastRow() - 1) > 0) {
    sendQueueSummaryToAdmin(emailsSentCount, queueSheet.getLastRow() - 1);
  }
}

/**
 * Send queue processing summary to admin
 */
function sendQueueSummaryToAdmin(sentCount, remainingCount) {
  const adminEmail = 'jyothikondupally@gmail.com'; // Update with actual admin email
  
  const currentUsage = getEmailUsageToday();
  const subject = `Kings Equestrian - Daily Email Queue Summary`;
  
  const body = `
Daily Email Queue Processing Report
Date: ${new Date().toDateString()}

📊 Summary:
- Emails sent today: ${currentUsage} / ${CONFIG.DAILY_EMAIL_LIMIT}
- Queued emails processed: ${sentCount}
- Emails still pending: ${remainingCount}

${remainingCount > 0 ? '⚠️ Note: Some emails are still in queue and will be processed tomorrow.' : '✅ All queued emails have been sent.'}

View the Email Queue sheet for details.
  `;
  
  try {
    MailApp.sendEmail(adminEmail, subject, body);
  } catch (error) {
    Logger.log('Could not send admin summary: ' + error.toString());
  }
}

// ============== MAIN TRIGGER FUNCTION ==============

/**
 * This function runs when a form is submitted
 * Set as onFormSubmit trigger
 */
function onFormSubmit(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const row = e.range.getRow();
    
    // Get column indices
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const locationCol = headers.indexOf('Location') + 1;
    const registerCol = headers.indexOf('Register Number') + 1;
    const emailCol = headers.indexOf('Email ID') + 1;
    const nameCol = headers.indexOf('Student Name') + 1;
    
    if (locationCol === 0 || registerCol === 0 || emailCol === 0) {
      throw new Error('Required columns not found. Please check column names.');
    }
    
    // Get submitted data
    const location = sheet.getRange(row, locationCol).getValue().toString().toLowerCase().trim();
    const email = sheet.getRange(row, emailCol).getValue();
    const name = nameCol > 0 ? sheet.getRange(row, nameCol).getValue() : 'Valued Customer';
    
    // Generate and set registration number
    const registerNumber = generateRegistrationNumber(sheet, location, registerCol);
    sheet.getRange(row, registerCol).setValue(registerNumber);
    
    // Check if we can send email today
    if (canSendEmailToday()) {
      // Send email immediately
      sendWelcomeEmail(email, name, registerNumber, location);
      incrementEmailUsage();
      updateEmailStatusInMainSheet(registerNumber, 'Sent' );
      
      const currentUsage = getEmailUsageToday();
      Logger.log(`Email sent immediately to ${email}. Daily count: ${currentUsage}/${CONFIG.DAILY_EMAIL_LIMIT}`);
    } else {
      // Add to queue for tomorrow
      addToEmailQueue(email, name, registerNumber, location);
      Logger.log(`Daily limit reached (${CONFIG.DAILY_EMAIL_LIMIT}). Email queued for tomorrow.`);
    }
    
  } catch (error) {
    Logger.log('Error in onFormSubmit: ' + error.toString());
    sendAdminNotification(error);
  }
}

// ============== REGISTRATION NUMBER GENERATOR ==============

/**
 * Generates a unique registration number based on location
 * Format: [LOCATION_CODE][4-digit serial number]
 */
function generateRegistrationNumber(sheet, location, registerCol) {
  const locationKey = Object.keys(CONFIG.locationCodes).find(key => 
    location.includes(key) || key.includes(location)
  );
  
  const locationCode = locationKey ? CONFIG.locationCodes[locationKey] : 'GEN';
  
  const lastRow = sheet.getLastRow();
  const existingNumbers = sheet.getRange(2, registerCol, lastRow - 1, 1).getValues();
  
  let maxNumber = 0;
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
  
  const newNumber = maxNumber + 1;
  const paddedNumber = newNumber.toString().padStart(4, '0');
  
  return `${locationCode}${paddedNumber}`;
}

// ============== EMAIL TEMPLATE ==============

/**
 * Sends a beautifully formatted welcome email to the new registrant
 */
function sendWelcomeEmail(email, name, registerNumber, location) {
  const htmlBody = createEmailTemplate(name, registerNumber, location);
  
  // Replace placeholders in subject
  const subject = CONFIG.emailSubject
    .replace('{name}', name)
    .replace('{regNumber}', registerNumber);
  
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
}

/**
 * Creates the HTML email template
 */
function createEmailTemplate(name, registerNumber, location) {
  const locationName = location.charAt(0).toUpperCase() + location.slice(1);
  
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
    .logo-section {
      margin-bottom: 20px;
    }
    .header h1 {
      margin: 0;
      font-size: 28px;
      font-weight: 600;
      letter-spacing: 0.5px;
    }
    .crown-icon {
      font-size: 48px;
      margin-bottom: 10px;
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
    .registration-box {
      background: linear-gradient(135deg, ${CONFIG.colors.background} 0%, #fff 100%);
      border-left: 4px solid ${CONFIG.colors.secondary};
      padding: 20px;
      margin: 25px 0;
      border-radius: 8px;
    }
    .registration-box p {
      margin: 5px 0;
    }
    .reg-number {
      font-size: 24px;
      font-weight: 700;
      color: ${CONFIG.colors.primary};
      letter-spacing: 1px;
    }
    .section-title {
      color: ${CONFIG.colors.primary};
      font-size: 20px;
      font-weight: 600;
      margin-top: 30px;
      margin-bottom: 15px;
      border-bottom: 2px solid ${CONFIG.colors.secondary};
      padding-bottom: 8px;
    }
    .benefits-grid {
      display: grid;
      gap: 15px;
      margin: 20px 0;
    }
    .benefit-item {
      background-color: ${CONFIG.colors.background};
      padding: 18px;
      border-radius: 8px;
      border-left: 3px solid ${CONFIG.colors.secondary};
    }
    .benefit-item h3 {
      margin: 0 0 8px 0;
      color: ${CONFIG.colors.primary};
      font-size: 16px;
    }
    .benefit-item p {
      margin: 0;
      color: ${CONFIG.colors.text};
      font-size: 14px;
    }
    .price {
      color: ${CONFIG.colors.secondary};
      font-weight: 700;
      font-size: 16px;
    }
    .info-section {
      background-color: #f9f9f9;
      padding: 20px;
      border-radius: 8px;
      margin: 20px 0;
    }
    .info-section ul {
      margin: 10px 0;
      padding-left: 20px;
    }
    .info-section li {
      margin: 8px 0;
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
      transition: transform 0.2s;
    }
    .cta-button:hover {
      transform: translateY(-2px);
    }
    .document-link {
      text-align: center;
      margin: 30px 0;
      padding: 25px;
      background: linear-gradient(135deg, #fff 0%, ${CONFIG.colors.background} 100%);
      border-radius: 8px;
      border: 2px dashed ${CONFIG.colors.secondary};
    }
    .footer {
      background-color: ${CONFIG.colors.primary};
      color: white;
      padding: 30px;
      text-align: center;
      font-size: 14px;
    }
    .footer p {
      margin: 5px 0;
    }
    .social-links {
      margin-top: 15px;
    }
    .highlight {
      color: ${CONFIG.colors.secondary};
      font-weight: 600;
    }
    @media only screen and (max-width: 600px) {
      .container {
        margin: 10px;
      }
      .content {
        padding: 25px 20px;
      }
      .header {
        padding: 30px 20px;
      }
    }
  </style>
</head>
<body>
  <div class="container">
    <!-- Header -->
    <div class="header">
      <div class="logo-section">
        <div class="crown-icon"><img src="https://kingsfarmequestrian.com/wp-content/uploads/2023/08/Logo2.jpg" width="70px" height="70px"/></div>
        <h1>KINGS EQUESTRIAN</h1>
      </div>
      <p style="margin: 10px 0 0 0; font-size: 16px; opacity: 0.95;">Welcome to the Premier Equestrian Experience</p>
    </div>

    <!-- Main Content -->
    <div class="content">
      <div class="greeting">Dear ${name},</div>
      
      <p>Welcome to <strong>Kings Equestrian</strong>! We're thrilled to have you join our equestrian family. Your journey towards mastering the art of horse riding begins here.</p>
      
      <!-- Registration Details -->
      <div class="registration-box">
        <p style="margin-bottom: 10px; font-size: 14px; color: #666;">Your Registration Number</p>
        <div class="reg-number">${registerNumber}</div>
        <p style="margin-top: 10px; font-size: 14px; color: #666;">Location: <strong>${locationName}</strong></p>
      </div>

      <!-- Benefits & Classes -->
      <div class="section-title">🏇 Our Classes & Benefits</div>
      
      <div class="benefits-grid">
        <div class="benefit-item">
          <h3>🌟 Beginner Classes</h3>
          <p>Perfect for first-time riders. Learn the fundamentals of horse riding with our expert instructors.</p>
          <p class="price">₹2,500/session | ₹20,000/month (8 sessions)</p>
        </div>
        
        <div class="benefit-item">
          <h3>🏆 Intermediate Training</h3>
          <p>Advance your skills with jumping, dressage, and advanced riding techniques.</p>
          <p class="price">₹3,500/session | ₹28,000/month (8 sessions)</p>
        </div>
        
        <div class="benefit-item">
          <h3>💎 Advanced/Competition Prep</h3>
          <p>Elite training for competitive riders with personalized coaching and competition guidance.</p>
          <p class="price">₹5,000/session | ₹40,000/month (8 sessions)</p>
        </div>
        
        <div class="benefit-item">
          <h3>👨‍👩‍👧‍👦 Kids Special Programs</h3>
          <p>Age-appropriate lessons for children (6+ years) focusing on safety and fun.</p>
          <p class="price">₹2,000/session | ₹16,000/month (8 sessions)</p>
        </div>
      </div>

      <!-- What You Get -->
      <div class="section-title">✨ What Makes Us Special</div>
      
      <div class="info-section">
        <ul>
          <li><strong>World-Class Facilities:</strong> State-of-the-art stables, training arenas, and equipment</li>
          <li><strong>Expert Instructors:</strong> Certified trainers with international experience</li>
          <li><strong>Well-Trained Horses:</strong> Gentle, well-maintained horses suitable for all skill levels</li>
          <li><strong>Safety First:</strong> Complete safety gear provided and strict safety protocols</li>
          <li><strong>Flexible Scheduling:</strong> Morning, evening, and weekend batches available</li>
          <li><strong>Community Events:</strong> Regular riding events, competitions, and social gatherings</li>
          <li><strong>Complementary Benefits:</strong> Access to grooming sessions, stable tours, and horse care workshops</li>
        </ul>
      </div>

      <!-- Additional Information -->
      <div class="section-title">📋 Important Information</div>
      
      <div class="info-section">
        <p><strong>What to Bring:</strong></p>
        <ul>
          <li>Comfortable athletic wear (long pants recommended)</li>
          <li>Closed-toe shoes with a small heel (riding boots preferred)</li>
          <li>Water bottle and towel</li>
          <li>Safety gear (available at facility if needed)</li>
        </ul>
        
        <p style="margin-top: 15px;"><strong>Booking & Schedule:</strong></p>
        <ul>
          <li>Book your sessions in advance through our portal or contact us directly</li>
          <li>Cancellations must be made 24 hours in advance</li>
          <li>First session includes orientation and facility tour</li>
        </ul>
      </div>

      <!-- Document Link -->
      <div class="document-link">
        <p style="margin: 0 0 15px 0; font-size: 16px; color: ${CONFIG.colors.primary}; font-weight: 600;">
          📄 Complete Details & Guidelines
        </p>
        <a href="https://kings-equestrian.web.app/pay?ref=KE-1640991234567&am=500" class="cta-button" style="color: white;">
          View Detailed Information
        </a>
        <p style="margin: 15px 0 0 0; font-size: 13px; color: #666;">
          Access our comprehensive guide with facility rules, safety guidelines, and more
        </p>
      </div>
      <!-- Terms & Conditions Button -->
<div class="document-link" style="margin-top: 35px;">
  <p style="margin: 0 0 15px 0; font-size: 16px; color: ${CONFIG.colors.primary}; font-weight: 600;">
    ✅ Terms & Conditions Acceptance
  </p>

  <p style="font-size: 14px; color: #555; margin-bottom: 20px;">
    To proceed further and receive the payment request, please review and accept our
    Terms & Conditions.
  </p>

  <a 
    href="${CONFIG.consentFormLink}"
    class="cta-button"
    style="color: white;"
    target="_blank"
  >
    Review & Accept Terms & Conditions
  </a>

  <p style="margin-top: 15px; font-size: 12px; color: #777;">
    Payment request will be shared after terms acceptance.
  </p>
</div>


      <!-- Next Steps -->
      <div style="background-color: #fffbf0; padding: 20px; border-radius: 8px; border-left: 4px solid ${CONFIG.colors.secondary}; margin-top: 30px;">
        <h3 style="margin-top: 0; color: ${CONFIG.colors.primary};">🎯 Next Steps</h3>
        <ol style="margin: 10px 0; padding-left: 20px;">
          <li>Review the detailed document for facility guidelines</li>
          <li>Review the Terms & Conditions and Accept it To get the Payment Link</li>
          <li>Complete the Payment and share the Details</li>
          <li>Meet your instructor and begin your equestrian journey!</li>
        </ol>
      </div>

      <p style="margin-top: 30px;">If you have any questions or need assistance, please don't hesitate to reach out. We're here to ensure you have an exceptional experience!</p>
      
      <p style="margin-top: 20px;">
        <strong>Ride with Pride!</strong><br>
        <span class="highlight">The Kings Equestrian Team</span>
      </p>
    </div>

    <!-- Footer -->
    <div class="footer">
      <p style="font-size: 16px; font-weight: 600; margin-bottom: 10px;">KINGS EQUESTRIAN</p>
      <p>📍 ${locationName} Branch</p>
      <p>📞 Contact: +91-XXXXXXXXXX | ✉️ info@kingsequestrian.com</p>
      <p style="margin-top: 15px; font-size: 12px; opacity: 0.9;">
        © ${new Date().getFullYear()} Kings Equestrian. All rights reserved.
      </p>
      <div class="social-links">
        <p style="font-size: 12px;">Follow us: Instagram | Facebook | YouTube</p>
      </div>
    </div>
  </div>
</body>
</html>
  `;
}

// ============== ADMIN NOTIFICATION ==============

/**
 * Sends an error notification to admin
 */
function sendAdminNotification(error) {
  const adminEmail = 'jyothikondupally@gmail.com'; // Update with actual admin email
  
  MailApp.sendEmail(
    adminEmail,
    'Kings Equestrian - Form Submission Error',
    `An error occurred during form submission processing:\n\n${error.toString()}\n\nTime: ${new Date()}`
  );
}

// ============== MANUAL TEST FUNCTIONS ==============

/**
 * Test email template
 */
function testEmailTemplate() {
  const testEmail = 'jyothikondupally@gmail.com'; // Change to your email
  const testName = 'John Doe';
  const testRegNumber = 'BLR0001';
  const testLocation = 'bangalore';
  
  sendWelcomeEmail(testEmail, testName, testRegNumber, testLocation);
  Logger.log('Test email sent to: ' + testEmail);
}

/**
 * Test registration number generation
 */
function testRegistrationNumber() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const registerCol = 2; // Adjust based on your sheet
  
  const regNumber = generateRegistrationNumber(sheet, 'bangalore', registerCol);
  Logger.log('Generated registration number: ' + regNumber);
}

/**
 * Test queue processing manually
 */
function testQueueProcessing() {
  processEmailQueue();
}

/**
 * View current email usage
 */
function viewEmailUsage() {
  const usage = getEmailUsageToday();
  Logger.log(`Emails sent today: ${usage} / ${CONFIG.DAILY_EMAIL_LIMIT}`);
  Logger.log(`Can send more: ${canSendEmailToday()}`);
}