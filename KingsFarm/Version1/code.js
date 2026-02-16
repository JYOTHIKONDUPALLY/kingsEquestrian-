// ============================================
// KINGS EQUESTRIAN - ENHANCED BOOKING SYSTEM 
// ============================================

// --------------- CONFIG ---------------

const CONFIG = {
  UPI_ID: "vyapar.176548151976@hdfcbank",
  BUSINESS_NAME: "KingsEquestrian",
  PAYMENT_FORM_LINK: "https://forms.gle/XXhQ5oaLt65VBfwK6", // Update this
  EMAIL_TEMPLATE_DOC_ID: "1bUTpk9QCR4n1uUmMuoSRRcflTShG3jawuhemE28aTio", // Doc with email body template
  webAppUrl:"https://script.google.com/macros/s/AKfycbxGNi137N_vvd6kFWe0CL2clALwKLp7QKsLgiWUd9fGcvYhTlaeQIy15n2vai_1g-PIig/exec",
  // Sheet names
  SHEETS: {
    BOOKING_FORM: "Booking Form Response",
    PAYMENT_FORM: "Payment Form Response",
    PRICING: "Pricing"
  },
  
  // Column indices for Booking Form
  BOOKING_COLS: {
    TIMESTAMP: 0,
    EMAIL_ADDRESS: 1,
    RIDER_NAME: 2,
    EMAIL_ID: 3,
    PHONE: 4,
    SERVICE: 5,
    DATE: 6,
    TIME_SLOTS: 7,
    CONSENT: 8,
    AMOUNT: 9,
    REFERENCE: 10,
    UPI_LINK: 11,
    QR_CODE: 12,
    PAYMENT_STATUS: 13,
    RECEIPT_NO: 14,
    WELCOME_EMAIL_SENT: 15,
    WELCOME_EMAIL_TIMESTAMP: 16,
     CALENDAR_EVENT_ID: 17,
    CHANGE_REQUESTS: 18,
    LAST_CHANGED_DATE: 19,
    CHANGE_HISTORY: 20
  },
  
  // Column indices for Payment Form
  PAYMENT_COLS: {
    TIMESTAMP: 0,
    EMAIL_ADDRESS: 1,
    REFERENCE_NO: 2,
    EMAIL_ID: 3,
    AMOUNT_PAID: 4,
    TRANSACTION_ID: 5,
    SCREENSHOT: 6,
    PAYMENT_DATE: 7,
    PAN_AADHAAR: 8,
    TRANSACTION_VERIFIED: 9,
    RECEIPT_SENT: 10,
    RECEIPT_SENT_TIMESTAMP: 11,
    PAYMENT_RECEIPT_NO: 12
  }
};

// --------------- UTILITY FUNCTIONS ---------------

/**
 * Generate simple, user-friendly reference number
 */
function generateReference() {
  const date = new Date();
  const year = date.getFullYear().toString().substr(-2);
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  const random = Math.floor(Math.random() * 9000) + 1000; // 4-digit random
  
  return `KE${year}${month}${day}${random}`; // Example: KE26021012345
}

/**
 * Create UPI payment link
 */
function createUPILink(amount, reference) {
  return `upi://pay?pa=${CONFIG.UPI_ID}&pn=${encodeURIComponent(CONFIG.BUSINESS_NAME)}&am=${amount}&cu=INR&tn=${encodeURIComponent(reference)}`;
}

/**
 * Create QR code URL
 */
function createQRCode(link) {
  return `https://api.qrserver.com/v1/create-qr-code/?size=400x400&data=${encodeURIComponent(link)}`;
}

/**
 * Get pricing data from Pricing sheet
 */
function getPricingData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pricingSheet = ss.getSheetByName(CONFIG.SHEETS.PRICING);
  
  if (!pricingSheet) {
    throw new Error('Pricing sheet not found');
  }
  
  const data = pricingSheet.getDataRange().getValues();
  const pricingMap = {};
  
  // Skip header row
  for (let i = 1; i < data.length; i++) {
    const serviceId = data[i][0]; // Service name
    const service = data[i][1];  // Service Name
    const pricePerHalfHour = data[i][2]; // Price per 30 min
    const docId = data[i][3]; // Google Doc ID
    
    if (service) {
      pricingMap[service] = {
        price: pricePerHalfHour,
        docId: docId
      };
    }
  }
  
  return pricingMap;
}

/**
 * Get service details from Google Doc
 */
function getServiceDetailsFromDoc(docId) {
  try {
    const doc = DocumentApp.openById(docId);
    const body = doc.getBody();
    const text = body.getText();
    
    // Extract summary section
    const summaryMatch = text.match(/summary[:\s]*([\s\S]*?)(?=\n\n|$)/i);
    const summary = summaryMatch ? summaryMatch[1].trim() : '';
    
    return {
      summary: summary,
      fullText: text
    };
  } catch (error) {
    Logger.log('Error fetching service details: ' + error);
    return {
      summary: 'Professional equestrian service at Kings Equestrian.',
      fullText: ''
    };
  }
}

/**
 * Get service PDF from Google Doc
 */
function getServicePDF(docId, serviceName) {
  try {
    const doc = DocumentApp.openById(docId);
    const blob = doc.getAs('application/pdf');
    blob.setName(`${serviceName.replace(/\s+/g, '_')}_Details.pdf`);
    return blob;
  } catch (error) {
    Logger.log('Error creating PDF: ' + error);
    return null;
  }
}

/**
 * Calculate amount and slot distribution
 */
function calculateBookingAmount(services, totalSlots) {
  const pricingData = getPricingData();
  
  // Parse services (can be single or multiple)
  const serviceList = typeof services === 'string' 
    ? services.split(',').map(s => s.trim()) 
    : [services];
  
  const numServices = serviceList.length;
  const slotsPerService = Math.floor(totalSlots / numServices);
  const remainingSlots = totalSlots % numServices;
  
  let totalAmount = 0;
  const breakdown = [];
  
  serviceList.forEach((service, index) => {
    const slots = slotsPerService + (index < remainingSlots ? 1 : 0);
    const pricing = pricingData[service];
    
    if (!pricing) {
      throw new Error(`Pricing not found for service: ${service}`);
    }
    
    const amount = pricing.price * slots;
    totalAmount += amount;
    
    breakdown.push({
      service: service,
      slots: slots,
      pricePerSlot: pricing.price,
      amount: amount,
      docId: pricing.docId
    });
  });
  
  return {
    totalAmount: totalAmount,
    breakdown: breakdown
  };
}

/**
 * Get email template from Google Doc
 */
function getEmailTemplate() {
  try {
    const doc = DocumentApp.openById(CONFIG.EMAIL_TEMPLATE_DOC_ID);
    const body = doc.getBody();
    return body.getText();
  } catch (error) {
    Logger.log('Error fetching email template: ' + error);
    return '';
  }
}

/**
 * Create header mapping from sheet
 */
function getHeaderIndexMap(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((header, index) => {
    if (header) {
      map[header.toString().trim()] = index;
    }
  });
  return map;
}

// --------------- MAIN FORM SUBMIT HANDLER ---------------

/**
 * Trigger: On booking form submit
 */
function onBookingFormSubmit(e) {
  try {
    const sheet = e.range.getSheet();
    const row = e.range.getRow();
    
    // Get form data
    const riderName = sheet.getRange(row, CONFIG.BOOKING_COLS.RIDER_NAME + 1).getValue();
    const email = sheet.getRange(row, CONFIG.BOOKING_COLS.EMAIL_ID + 1).getValue();
    const phone = sheet.getRange(row, CONFIG.BOOKING_COLS.PHONE + 1).getValue();
    const services = sheet.getRange(row, CONFIG.BOOKING_COLS.SERVICE + 1).getValue();
    const selectedDate = sheet.getRange(row, CONFIG.BOOKING_COLS.DATE + 1).getValue();
    const timeSlots = sheet.getRange(row, CONFIG.BOOKING_COLS.TIME_SLOTS + 1).getValue();
    
    // Parse time slots count
    let slotCount = 1;
    if (typeof timeSlots === 'string') {
      slotCount = timeSlots.split(',').length;
    }
    
    // Calculate amount with breakdown
    const calculation = calculateBookingAmount(services, slotCount);
    const amount = calculation.totalAmount;
    
    // Generate payment details
    const reference = generateReference();
    const upiLink = createUPILink(amount, reference);
    const qrCode = createQRCode(upiLink);
    
    // Update sheet
    sheet.getRange(row, CONFIG.BOOKING_COLS.AMOUNT + 1).setValue(amount);
    sheet.getRange(row, CONFIG.BOOKING_COLS.REFERENCE + 1).setValue(reference);
    sheet.getRange(row, CONFIG.BOOKING_COLS.UPI_LINK + 1).setValue(upiLink);
    sheet.getRange(row, CONFIG.BOOKING_COLS.QR_CODE + 1).setValue(qrCode);
    sheet.getRange(row, CONFIG.BOOKING_COLS.PAYMENT_STATUS + 1).setValue("Pending");
    
     const calendarEventId = createCalendarEvent({
      name: riderName,
      email: email,
      phone: phone,
      services: services,
      date: selectedDate,
      timeSlots: timeSlots,
      reference: reference
    });

    if (calendarEventId) {
      const headerMap = getHeaderIndexMap(sheet);
      if (headerMap['Calendar Event ID'] === undefined) {
        const lastCol = sheet.getLastColumn();
        sheet.getRange(1, lastCol + 1).setValue('Calendar Event ID');
      }
      sheet.getRange(row, sheet.getLastColumn()).setValue(calendarEventId);
    }
    // Send welcome email with booking confirmation
    sendWelcomeEmail({
      name: riderName,
      email: email,
      phone: phone,
      services: services,
      date: selectedDate,
      timeSlots: timeSlots,
      amount: amount,
      reference: reference,
      upiLink: upiLink,
      qrCode: qrCode,
      breakdown: calculation.breakdown,
      row: row,
      sheet: sheet
    });
    
    Logger.log(`Booking processed for ${riderName} - Reference: ${reference}`);
    
  } catch (error) {
    Logger.log('Error in onBookingFormSubmit: ' + error);
    SpreadsheetApp.getUi().alert('Error processing booking: ' + error.message);
  }
}

// --------------- EMAIL FUNCTIONS ---------------

/**
 * Send welcome email with booking details
 */
function sendWelcomeEmail(data) {
  const subject = `🐎 Welcome to Kings Equestrian - Booking ${data.reference}`;
  
  // Get service details and PDFs
  let serviceDetailsHTML = '';
  const attachments = [];
  
  data.breakdown.forEach(item => {
    const details = getServiceDetailsFromDoc(item.docId);
    serviceDetailsHTML += `
      <div style="margin: 15px 0; padding: 15px; background: #f9f9f9; border-radius: 8px;">
        <h3 style="color: #2c5f2d; margin: 0 0 10px 0;">${item.service}</h3>
        <p style="margin: 5px 0; font-size: 14px;">${details.summary}</p>
        <p style="margin: 5px 0; color: #666; font-size: 13px;">
          <strong>Slots allocated:</strong> ${item.slots} × 30 minutes<br>
          <strong>Cost:</strong> ₹${item.pricePerSlot} per slot = ₹${item.amount}
        </p>
      </div>
    `;
    
    // Add PDF attachment
    const pdf = getServicePDF(item.docId, item.service);
    if (pdf) {
      attachments.push(pdf);
    }
  });

  const changeRequestSection = `
          <div style="background: #fff3cd; border: 2px solid #ffc107; padding: 20px; border-radius: 10px; margin: 25px 0;">
            <h3 style="margin-top: 0; color: #856404;">📅 Need to Change Your Date or Time?</h3>
            <p style="margin: 10px 0;">
              If the selected date or time doesn't work for you, you can easily request a change:
            </p>
            <div style="text-align: center; margin: 20px 0;">
              <a href="${CONFIG.webAppUrl}?ref=${data.reference}" 
                 class="button" 
                 style="background-color: #ffc107; color: #333;">
                🔄 Request Date/Time Change
              </a>
            </div>
            <p style="font-size: 13px; color: #666; margin: 10px 0;">
              <strong>Note:</strong> You can also call us at <strong>+91-9980895533</strong> to discuss available slots.
            </p>
          </div>
`;
  
  // Create breakdown table
  let breakdownTable = `
    <table style="width: 100%; border-collapse: collapse; margin: 20px 0; font-size: 14px;">
      <thead>
        <tr style="background: #2c5f2d; color: white;">
          <th style="padding: 12px; text-align: left; border: 1px solid #ddd;">Service</th>
          <th style="padding: 12px; text-align: center; border: 1px solid #ddd;">Slots</th>
          <th style="padding: 12px; text-align: right; border: 1px solid #ddd;">Rate</th>
          <th style="padding: 12px; text-align: right; border: 1px solid #ddd;">Amount</th>
        </tr>
      </thead>
      <tbody>
  `;
  
  data.breakdown.forEach(item => {
    breakdownTable += `
      <tr>
        <td style="padding: 10px; border: 1px solid #ddd;">${item.service}</td>
        <td style="padding: 10px; text-align: center; border: 1px solid #ddd;">${item.slots}</td>
        <td style="padding: 10px; text-align: right; border: 1px solid #ddd;">₹${item.pricePerSlot}</td>
        <td style="padding: 10px; text-align: right; border: 1px solid #ddd;">₹${item.amount}</td>
      </tr>
    `;
  });
  
  breakdownTable += `
      <tr style="background: #f0f0f0; font-weight: bold;">
        <td colspan="3" style="padding: 12px; text-align: right; border: 1px solid #ddd;">Total Amount:</td>
        <td style="padding: 12px; text-align: right; border: 1px solid #ddd; color: #2c5f2d; font-size: 16px;">₹${data.amount}</td>
      </tr>
      </tbody>
    </table>
  `;
  
  const htmlBody = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <style>
        body {
          font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
          background-color: #f4f4f4;
          margin: 0;
          padding: 0;
          color: #333;
        }
        .container {
          max-width: 650px;
          margin: 20px auto;
          background-color: #ffffff;
          border-radius: 12px;
          overflow: hidden;
          box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }
        .header {
          background: linear-gradient(135deg, #1f4e3d 0%, #4f9c7a 100%);
          padding: 30px;
          text-align: center;
          color: #fff;
        }
        .header img {
          width: 70px;
          height: 70px;
          border-radius: 10%;
          margin-bottom: 10px;
        }
        .header h1 {
          margin: 0;
          font-size: 26px;
        }
        .banner {
          background: #fff3cd;
          padding: 20px;
          text-align: center;
          border-bottom: 3px solid #ffc107;
        }
        .banner h2 {
          margin: 0;
          color: #856404;
          font-size: 24px;
        }
        .content {
          padding: 35px 30px;
        }
        .welcome-box {
          background: linear-gradient(135deg, #2c5f2d 0%, #4f9c7a 100%);
          color: white;
          padding: 25px;
          border-radius: 10px;
          text-align: center;
          margin: 20px 0;
        }
        .info-box {
          background: #f8f8f8;
          border-left: 4px solid #2c5f2d;
          padding: 18px;
          margin: 20px 0;
          border-radius: 6px;
        }
        .payment-section {
          background: #fff9e6;
          border: 2px solid #ffc107;
          padding: 25px;
          border-radius: 10px;
          margin: 25px 0;
        }
        .qr-container {
          text-align: center;
          margin: 20px 0;
        }
        .qr-container img {
          max-width: 300px;
          border: 3px solid #2c5f2d;
          border-radius: 10px;
          padding: 10px;
          background: white;
        }
        .reference-box {
          background: #2c5f2d;
          color: white;
          padding: 20px;
          border-radius: 8px;
          text-align: center;
          margin: 20px 0;
        }
        .reference-number {
          font-size: 32px;
          font-weight: bold;
          letter-spacing: 3px;
          margin: 10px 0;
          font-family: 'Courier New', monospace;
        }
        .button {
          background-color: #2c5f2d;
          color: white;
          padding: 15px 35px;
          text-decoration: none;
          display: inline-block;
          margin: 10px 5px;
          border-radius: 8px;
          font-weight: bold;
          font-size: 16px;
        }
        .button:hover {
          background-color: #1f4e3d;
        }
        .footer {
          background-color: #1f4e3d;
          color: #fff;
          padding: 25px;
          text-align: center;
          font-size: 13px;
        }
        .note-box {
          background: #e7f3ff;
          border-left: 4px solid #0066cc;
          padding: 15px;
          margin: 15px 0;
          border-radius: 5px;
          font-size: 14px;
        }
      </style>
    </head>
    <body>
      <div class="container">
        
        <!-- Header -->
        <div class="header">
          <img src="https://kingsfarmequestrian.com/wp-content/uploads/2023/08/Logo2.jpg" alt="Kings Equestrian Logo">
          <h1>Kings Equestrian Foundation</h1>
          <p style="margin:8px 0 0;">Where horses don't just carry you — they change you</p>
        </div>

        <!-- Welcome Banner -->
        <div class="banner">
          <h2>🎉 Welcome to the Kings Equestrian Family! 🐎</h2>
        </div>
        
        <!-- Content -->
        <div class="content">
          
          <div class="welcome-box">
            <h2 style="margin: 0 0 10px 0;">🌟 Thank You, ${data.name}! 🌟</h2>
            <p style="margin: 5px 0; font-size: 16px;">
              We're thrilled to have you join our equestrian community!
            </p>
          </div>

          <p style="font-size: 16px; line-height: 1.6;">
            Dear ${data.name},
          </p>

          <p>
            Thank you for choosing Kings Equestrian Foundation! Your booking request has been received successfully. 
            We're excited to welcome you to an unforgettable equestrian experience.
          </p>
          
          <!-- Reference Number -->
          <div class="reference-box">
            <div style="font-size: 14px; margin-bottom: 5px;">Your Booking Reference</div>
            <div class="reference-number">${data.reference}</div>
            <div style="font-size: 12px; margin-top: 5px; opacity: 0.9;">
              Please save this number for payment submission and future reference
            </div>
          </div>

          <h3 style="color: #2c5f2d;">📋 Your Booking Details:</h3>
          <div class="info-box">
            <p style="margin: 8px 0;"><strong>Name:</strong> ${data.name}</p>
            <p style="margin: 8px 0;"><strong>Contact:</strong> ${data.phone}</p>
            <p style="margin: 8px 0;"><strong>Preferred Date:</strong> ${formatDate(data.date)}</p>
            <p style="margin: 8px 0;"><strong>Time Slots:</strong> ${data.timeSlots}</p>
          </div>

          <h3 style="color: #2c5f2d;">🏇 Service Details:</h3>
          ${serviceDetailsHTML}
           <h3 style="color: #2c5f2d;">📅 Need to Change Your Booking?</h3>
  ${changeRequestSection}

          <h3 style="color: #2c5f2d;">💰 Payment Breakdown:</h3>
          ${breakdownTable}
          
          <div class="note-box">
            <strong>📌 Note:</strong> The time slots have been automatically distributed among your selected services. 
            If you'd like to adjust the time allocation for specific services, please contact us at 
            <strong>+91-9980895533</strong> or <a href="mailto:info@kingsequestrian.com">info@kingsequestrian.com</a>
          </div>

          <!-- Payment Section -->
          <div class="payment-section">
            <h3 style="margin-top: 0; color: #856404;">💳 Complete Your Payment</h3>
            <p>To confirm your booking, please pay <strong style="font-size: 24px; color: #2c5f2d;">₹${data.amount}</strong></p>
            
            <div class="qr-container">
              <p style="margin: 10px 0; font-weight: bold;">Scan QR Code to Pay via UPI:</p>
              <img src="${data.qrCode}" alt="Payment QR Code" />
            </div>
            
            <div style="text-align: center; margin: 20px 0;">
              <p style="margin: 10px 0;"><strong>OR</strong></p>
              <a href="${data.upiLink}" class="button">📱 Pay via UPI App</a>
            </div>
            
            <div style="text-align: center; margin: 25px 0; padding-top: 20px; border-top: 2px dashed #ffc107;">
              <p style="font-weight: bold; margin-bottom: 15px;">After payment, submit your confirmation:</p>
              <a href="${CONFIG.PAYMENT_FORM_LINK}" class="button" style="background-color: #0066cc;">
                ✅ Submit Payment Details
              </a>
            </div>
            
            <div style="background: white; padding: 15px; border-radius: 5px; margin-top: 20px;">
              <p style="margin: 5px 0; font-size: 13px; color: #666;">
                <strong>Payment Instructions:</strong>
              </p>
              <ol style="margin: 10px 0; padding-left: 20px; font-size: 13px; color: #666;">
                <li>Use the QR code or UPI link above to make payment</li>
                <li>Use your reference number: <strong>${data.reference}</strong></li>
                <li>After successful payment, click "Submit Payment Details"</li>
                <li>Fill in the transaction details and upload screenshot</li>
              </ol>
            </div>
          </div>

          <h3 style="color: #2c5f2d;">📚 Attached Documents:</h3>
          <div class="info-box">
            <p>We've attached detailed information about your selected services. 
            Please review them to prepare for your visit.</p>
          </div>

          <h3 style="color: #2c5f2d;">📌 Important Information:</h3>
          <ul style="line-height: 1.8;">
            <li>Each time slot is <strong>30 minutes</strong></li>
            <li>Please arrive <strong>15 minutes before</strong> your scheduled time</li>
            <li>Wear <strong>comfortable clothing</strong> and <strong>closed-toe shoes</strong></li>
            <li>Your booking will be confirmed once payment is verified</li>
            <li>For any changes or queries, contact us with your reference number</li>
          </ul>
          

          <div class="welcome-box" style="margin-top: 30px;">
            <p style="margin: 0; font-size: 16px;">
              🐴 We can't wait to see you at Kings Equestrian! 🐴
            </p>
          </div>

        </div>

        <!-- Footer -->
        <div class="footer">
          <p><strong>Kings Equestrian Foundation</strong></p>
          <p>📍 Karnataka, India</p>
          <p>📞 +91-9980895533 | ✉️ info@kingsequestrian.com</p>
          <p style="margin-top: 10px; font-size: 11px;">
            © ${new Date().getFullYear()} Kings Equestrian Foundation. All rights reserved.
          </p>
        </div>

      </div>
    </body>
    </html>
  `;
  
  // Plain text version
  const plainBody = `
Welcome to Kings Equestrian Foundation! 🐎

Dear ${data.name},

Your booking reference: ${data.reference}

BOOKING DETAILS:
Name: ${data.name}
Contact: ${data.phone}
Date: ${formatDate(data.date)}
Time Slots: ${data.timeSlots}

TOTAL AMOUNT: ₹${data.amount}

PAYMENT BREAKDOWN:
${data.breakdown.map(item => 
  `${item.service}: ${item.slots} slots × ₹${item.pricePerSlot} = ₹${item.amount}`
).join('\n')}

Please pay using:
- UPI Link: ${data.upiLink}
- QR Code: ${data.qrCode}

After payment, submit details at: ${CONFIG.PAYMENT_FORM_LINK}

Kings Equestrian Foundation
Where horses don't just carry you — they change you
  `;
  
  // Send email
  MailApp.sendEmail({
    to: data.email,
    subject: subject,
    body: plainBody,
    htmlBody: htmlBody,
    attachments: attachments,
    name: "Kings Equestrian Foundation"
  });
  
  // Update welcome email sent status
  if (data.sheet && data.row) {
    data.sheet.getRange(data.row, CONFIG.BOOKING_COLS.WELCOME_EMAIL_SENT + 1)
      .setValue('Yes')
      .setBackground('#d4edda')
      .setFontColor('#155724')
      .setFontWeight('bold');
    
    data.sheet.getRange(data.row, CONFIG.BOOKING_COLS.WELCOME_EMAIL_TIMESTAMP + 1)
      .setValue(new Date())
      .setNumberFormat("dd-MMM-yyyy HH:mm:ss");
  }
  
  Logger.log(`Welcome email sent to: ${data.email}`);
}

// --------------- RECEIPT GENERATION (Using your existing code) ---------------

function generateReceiptNumber() {
  const year = new Date().getFullYear();
  const props = PropertiesService.getScriptProperties();
  const key = `LAST_RECEIPT_${year}`;
  const lastNumber = Number(props.getProperty(key)) || 0;
  const newNumber = lastNumber + 1;
  props.setProperty(key, newNumber);
  return `${year}/R/${String(newNumber).padStart(4, '0')}`;
}

function getImageAsBase64(fileId) {
  try {
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    const base64 = Utilities.base64Encode(blob.getBytes());
    const mimeType = blob.getContentType();
    return `data:${mimeType};base64,${base64}`;
  } catch (error) {
    Logger.log('Error getting image: ' + error);
    return '';
  }
}

function getImageFromUrlAsBase64(url) {
  try {
    const response = UrlFetchApp.fetch(url);
    const blob = response.getBlob();
    const base64 = Utilities.base64Encode(blob.getBytes());
    const mimeType = blob.getContentType();
    return `data:${mimeType};base64,${base64}`;
  } catch (error) {
    Logger.log('Error fetching image from URL: ' + error);
    return '';
  }
}

function numberToWords(num) {
  const ones = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine'];
  const teens = ['Ten', 'Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen'];
  const tens = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety'];
  
  function convert(num) {
    if (num === 0) return 'Zero';
    if (num < 10) return ones[num];
    if (num < 20) return teens[num - 10];
    if (num < 100) return tens[Math.floor(num / 10)] + (num % 10 ? ' ' + ones[num % 10] : '');
    if (num < 1000) return ones[Math.floor(num / 100)] + ' Hundred' + (num % 100 ? ' ' + convert(num % 100) : '');
    
    if (num < 100000) {
      return convert(Math.floor(num / 1000)) + ' Thousand' + (num % 1000 ? ' ' + convert(num % 1000) : '');
    }
    if (num < 10000000) {
      return convert(Math.floor(num / 100000)) + ' Lakh' + (num % 100000 ? ' ' + convert(num % 100000) : '');
    }
    return convert(Math.floor(num / 10000000)) + ' Crore' + (num % 10000000 ? ' ' + convert(num % 10000000) : '');
  }
  
  return convert(num).trim() + ' Rupees';
}

function generate80GReceipt(riderName, pan, amount, transactionRef, receiptNumber) {
  Logger.log('Converting logo to base64...');
  const logoBase64 = getImageFromUrlAsBase64('https://kingsfarmequestrian.com/wp-content/uploads/2023/08/Logo2.jpg');
  
  Logger.log('Converting stamp to base64...');
  const stampBase64 = getImageAsBase64('1kkDoebRZYYJDW76jYNZT1rWX65_DjSMs');
  
  Logger.log('Converting signature to base64...');
  const signBase64 = getImageAsBase64('1z8rGx3HkgyBb-nqIXIT-_BgY8cqiQDRR');
  
  const htmlContent = createReceiptHTML(riderName, pan, amount, transactionRef, receiptNumber, logoBase64, stampBase64, signBase64);
  
  const htmlFile = DriveApp.createFile(`receipt_temp_${new Date().getTime()}.html`, htmlContent, MimeType.HTML);
  const blob = htmlFile.getAs('application/pdf');
  blob.setName(`80G_Receipt_${riderName.replace(/\s+/g, '_')}_${receiptNumber}.pdf`);
  htmlFile.setTrashed(true);
  
  return blob;
}

function createReceiptHTML(donorName, pan, amount, transactionRef, receiptNumber, logoBase64, stampBase64, signBase64) {
  const currentDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yy');
  const amountInWords = numberToWords(amount);
  
  const html = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <style>
    @page {
      size: A4;
      margin: 0;
    }
    
    body {
      font-family: Arial, sans-serif;
      margin: 0;
      padding: 30px;
      background: #fff;
    }
    
    .receipt-container {
      border: 2px solid #000;
      border-radius: 25px;
      padding: 25px;
      max-width: 750px;
      margin: 0 auto;
    }
    
    .header {
      display: flex;
      align-items: flex-start;
      margin-bottom: 15px;
      position: relative;
    }
    
    .logo-section {
      flex: 0 0 130px;
      text-align: center;
    }
    
    .logo-img {
      width: 100px;
      height: 100px;
      margin-bottom: 5px;
    }
    
    .logo-text {
      font-size: 11px;
      font-weight: bold;
      line-height: 1.2;
    }
    
    .header-center {
      flex: 1;
      text-align: center;
      padding: 0 15px;
    }
    
    .org-name {
      font-size: 22px;
      font-weight: bold;
      margin-bottom: 3px;
    }
    
    .registration-info {
      font-size: 10px;
      margin-bottom: 2px;
    }
    
    .location {
      font-size: 9px;
    }
    
    .receipt-number {
      position: absolute;
      right: 0;
      top: 0;
      background: #ff4444;
      color: white;
      font-size: 22px;
      font-weight: bold;
      padding: 8px 18px;
      border-radius: 5px;
    }
    
    .receipt-box {
      border: 1px solid #000;
      border-radius: 10px;
      padding: 12px;
      text-align: center;
      margin: 15px 0;
    }
    
    .receipt-title {
      font-size: 18px;
      font-weight: bold;
      margin-bottom: 3px;
    }
    
    .receipt-subtitle {
      font-size: 9px;
      font-style: italic;
    }
    
    .main-content {
      display: flex;
      gap: 20px;
    }
    
    .left-column {
      flex: 1;
    }
    
    .right-column {
      flex: 1;
    }
    
    .section-title {
      font-size: 11px;
      font-weight: bold;
      margin-bottom: 8px;
    }
    
    .checkbox-list {
      font-size: 10px;
      margin-bottom: 12px;
    }
    
    .checkbox-item {
      margin: 4px 0;
      display: flex;
      align-items: center;
    }
    
    .checkbox {
      width: 12px;
      height: 12px;
      border: 1px solid #000;
      display: inline-block;
      margin-right: 6px;
    }
    
    .checkbox.checked {
      background: #000;
      position: relative;
    }
    
    .checkbox.checked::after {
      content: '✓';
      color: white;
      font-size: 10px;
      position: absolute;
      top: -2px;
      left: 1px;
    }
    
    .amount-section {
      border: 2px solid #000;
      padding: 15px;
      margin: 15px 0;
      position: relative;
      min-height: 60px;
    }
    
    .rupee-symbol {
      position: absolute;
      left: 15px;
      top: 50%;
      transform: translateY(-50%);
      font-size: 36px;
      font-weight: bold;
      color: #ffa500;
    }
    
    .amount-value {
      text-align: center;
      font-size: 32px;
      font-weight: bold;
      padding-top: 5px;
    }
    
    .payment-mode {
      font-size: 9px;
      margin: 10px 0 5px 0;
      font-weight: bold;
    }
    
    .declaration-section {
      font-size: 8.5px;
      line-height: 1.4;
      text-align: justify;
      margin-top: 8px;
    }
    
    .donor-details {
      font-size: 10px;
      line-height: 1.6;
    }
    
    .detail-row {
      margin: 5px 0;
    }
    
    .detail-label {
      font-weight: bold;
    }
    
    .signature-section {
      margin-top: 30px;
      display: flex;
      justify-content: space-between;
      align-items: flex-end;
    }
    
    .left-signature {
      flex: 1;
    }
    
    .right-signature {
      flex: 1;
      text-align: center;
    }
    
    .org-label {
      font-size: 11px;
      font-weight: bold;
      margin-bottom: 5px;
    }
    
    .stamp-and-sign {
      position: relative;
      width: 150px;
      height: 150px;
      margin:auto 0;
    }
    
    .stamp-img {
      position: absolute;
      width: 120px;
      height: 120px;
      left: 15px;
      top: 50px;
    }
    
    .sign-img {
      position: absolute;
      width: 100px;
      height: 40px;
      left:25px;
      top: -45px;
    }
    
    .authorized-text {
      font-size: 10px;
      margin-top: 70px;
      text-decoration: underline;
    }
  </style>
</head>
<body>
  <div class="receipt-container">
    <div class="header">
      <div class="logo-section">
        <img src="${logoBase64}" class="logo-img" alt="Logo">
        <div class="logo-text">KINGS EQUESTRIAN<br>SADOLI, UP.</div>
      </div>
      
      <div class="header-center">
        <div class="org-name">Kings Equestrian Foundation</div>
        <div class="registration-info">Registered u/s 80G of Income-tax Act, 1961, PAN: AAJCK7191E</div>
        <div class="location">Karnataka, India</div>
        <div class="receipt-box">
          <div class="receipt-title">Receipt</div>
          <div class="receipt-subtitle">This receipt is issued in compliance with Rule 18AB and Form 10BD requirements.</div>
        </div>
      </div>
      
      <div class="receipt-number">${receiptNumber}</div>
    </div>
    
    <div>
      <div class="main-content">
        <div class="left-column">
          <div class="section-title">Donation Purpose ( ✓ Tick Applicable)</div>
          <div class="checkbox-list" style="display:flex;flex-wrap:wrap; gap:10px;">
            <div class="checkbox-item">
              <span class="checkbox checked"></span>
              <span>All</span>
            </div>
            <div class="checkbox-item">
              <span class="checkbox"></span>
              <span>Horse welfare</span>
            </div>
            <div class="checkbox-item">
              <span class="checkbox"></span>
              <span>Recovery training</span>
            </div>
            <div class="checkbox-item">
              <span class="checkbox"></span>
              <span>Nutrition & feed</span>
            </div>
            <div class="checkbox-item">
              <span class="checkbox"></span>
              <span>Rehabilitation including timely veterinary care</span>
            </div>
            <div class="checkbox-item">
              <span class="checkbox"></span>
              <span>Non-commercial equestrian skill and sports development</span>
            </div>
          </div>
          
          <div class="section-title">Donor Category ( ✓ Tick Applicable)</div>
          <div class="checkbox-list">
            <div class="checkbox-item">
              <span class="checkbox checked"></span>
              <span>Resident Indian Donor</span>
            </div>
            <div class="checkbox-item">
              <span class="checkbox"></span>
              <span>Non Resident Indian (NRI)</span>
            </div>
          </div>
        </div>
        
        <div class="right-column">
          <div class="section-title">Donor Details</div>
          <div class="donor-details">
            <div class="detail-row">
              <span class="detail-label">Date:</span> ${currentDate}
            </div>
            <div class="detail-row">
              <span class="detail-label">Name of Donor:</span> ${donorName}
            </div>
            <div class="detail-row">
              <span class="detail-label">PAN / Aadhaar:</span> ${pan}
            </div>
            <div class="detail-row">
              <span class="detail-label">Amount in Words:</span> ${amountInWords}
            </div>
          </div>
        </div>
      </div>
      
      <div class="amount-section">
        <span class="rupee-symbol">₹</span>
        <div class="amount-value">${amount.toLocaleString('en-IN')}</div>
      </div>
      
      <div class="payment-mode">Mode of Payment (✓ Tick above): <span>${transactionRef}</span></div>
      <div class="checkbox-list" style="display:flex;flex-wrap:wrap; gap:10px;">
        <div class="checkbox-item">
          <span class="checkbox checked"></span>
          <span>UPI</span>
        </div>
        <div class="checkbox-item">
          <span class="checkbox"></span>
          <span>RTGS</span>
        </div>
        <div class="checkbox-item">
          <span class="checkbox"></span>
          <span>NEFT</span>
        </div>
        <div class="checkbox-item">
          <span class="checkbox"></span>
          <span>DD</span>
        </div>
        <div class="checkbox-item">
          <span class="checkbox"></span>
          <span>Cheque</span>
        </div>
      </div>
      
      <div class="declaration-section">
        Certified that the above donation is received by trust for charitable purposes only. This donation is eligible for deduction under Section 80G of the Income Tax Act, 1961, subject to applicable limits. This receipt will be reported in Form 10BD and a certificate in Form 10BE will be issued to the donor.
      </div>
      
      <div style="margin-top: 40px;">
        <div class="org-label">For Kings Equestrian Foundation</div>
        <div class="stamp-and-sign">
          <div class="authorized-text">Authorized Signatory</div>
          <img src="${signBase64}" class="sign-img" alt="Signature">
          <img src="${stampBase64}" class="stamp-img" alt="Stamp">
        </div>
      </div>
    </div>
  </div>
</body>
</html>
  `;
  
  return html;
}

// --------------- PAYMENT RECEIPT FUNCTION ---------------

function SendPaymentReceipt() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const paymentSheet = ss.getSheetByName(CONFIG.SHEETS.PAYMENT_FORM);
  const bookingSheet = ss.getSheetByName(CONFIG.SHEETS.BOOKING_FORM);
  
  if (!paymentSheet) {
    ui.alert('❌ Payment Form Response sheet not found');
    return;
  }
  
  const selection = paymentSheet.getActiveRange();
  
  if (!selection) {
    ui.alert('Please select rows to send receipts');
    return;
  }
  
  const startRow = selection.getRow();
  const numRows = selection.getNumRows();
  
  if (startRow === 1) {
    ui.alert('Cannot send receipts for header row');
    return;
  }
  
  const response = ui.alert(
    'Send Payment Receipts',
    `Send receipts for ${numRows} row(s)?`,
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  const paymentHeaderMap = getHeaderIndexMap(paymentSheet);
  const bookingHeaderMap = getHeaderIndexMap(bookingSheet);
  const bookingValues = bookingSheet.getDataRange().getValues();
  
  let successCount = 0;
  let failCount = 0;
  const errors = [];
  
  for (let i = 0; i < numRows; i++) {
    const rowIndex = startRow + i;
    let email = '';
    let referenceNumber = '';
    
    try {
      const row = paymentSheet.getRange(rowIndex, 1, 1, paymentSheet.getLastColumn()).getValues()[0];
      
      referenceNumber = row[paymentHeaderMap['Reference No']];
      const amount = row[paymentHeaderMap['Amount Paid (₹)']];
      const transactionId = row[paymentHeaderMap['Transaction Reference ID']];
      const pan = row[paymentHeaderMap['Pan / AAdhar Number']];
      email = row[paymentHeaderMap['Email Id']];
      const transactionVerified = row[paymentHeaderMap['TRANSACTION_VERIFIED']];
      
      if (!referenceNumber || !email || !amount || !transactionId) {
        throw new Error('Missing required fields');
      }
      
      if (transactionVerified !== 'Yes') {
        throw new Error('Transaction Not verified');
      }
      
      // Find rider name from Booking sheet
      let riderName = '';
      
      for (let j = 1; j < bookingValues.length; j++) {
        const bookingRow = bookingValues[j];
        
        if (bookingRow[bookingHeaderMap['Reference']] === referenceNumber) {
          riderName = bookingRow[bookingHeaderMap['Rider Name']];
          
          // Update payment status in booking sheet
          bookingSheet.getRange(j + 1, bookingHeaderMap['Payment Status'] + 1)
            .setValue('Paid')
            .setBackground('#d4edda')
            .setFontColor('#155724')
            .setFontWeight('bold');
          
          break;
        }
      }
      
      if (!riderName) {
        throw new Error(`Rider name not found for ${referenceNumber}`);
      }
      
      // Generate receipt
      const receiptNumber = generateReceiptNumber();
      const receiptPDF = generate80GReceipt(
        riderName,
        pan,
        amount,
        transactionId,
        receiptNumber
      );
      
      const subject = `Payment Receipt - ${riderName} - Ref: ${referenceNumber}`;
      
      const htmlBody = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <style>
    body {
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
      background-color: #f4f4f4;
      margin: 0;
      padding: 0;
      color: #333;
    }
    .container {
      max-width: 650px;
      margin: 20px auto;
      background-color: #ffffff;
      border-radius: 12px;
      overflow: hidden;
      box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .header {
      background: linear-gradient(135deg, #1f4e3d 0%, #4f9c7a 100%);
      padding: 30px;
      text-align: center;
      color: #fff;
    }
    .header img {
      width: 70px;
      height: 70px;
      border-radius: 10%;
      margin-bottom: 10px;
    }
    .header h1 {
      margin: 0;
      font-size: 26px;
    }
    .content {
      padding: 35px 30px;
    }
    .greeting {
      font-size: 18px;
      font-weight: 600;
      margin-bottom: 15px;
      color: #1f4e3d;
    }
    .success-banner {
      text-align: center;
      margin: 25px 0;
    }
    .info-box {
      background: #f8f8f8;
      border-left: 4px solid #1f4e3d;
      padding: 18px;
      margin: 20px 0;
      border-radius: 6px;
      font-size: 14px;
    }
    .footer {
      background-color: #1f4e3d;
      color: #fff;
      padding: 25px;
      text-align: center;
      font-size: 13px;
    }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <img src="https://kingsfarmequestrian.com/wp-content/uploads/2023/08/Logo2.jpg" alt="Kings Equestrian Logo">
      <h1>Kings Equestrian Foundation</h1>
      <p style="margin:8px 0 0;">Where horses don't just carry you — they change you</p>
    </div>

    <div class="content">
      <div class="greeting">Dear ${riderName},</div>

      <div class="success-banner">
        <img 
          src="https://i.pinimg.com/736x/69/3c/20/693c200ad675967032f941cf76953b3e.jpg"
          alt="Payment Successful"
          width="200"
          height="150"
        />
        <div style="font-size:18px; font-weight:600; color:#1f7a3f; margin-top:10px;">
          ✅ Payment Confirmed - Booking Complete!
        </div>
      </div>

      <p>
        Thank you for your payment! Your booking with <strong>Kings Equestrian Foundation</strong> 
        is now confirmed. We're excited to welcome you!
      </p>

      <div class="info-box">
        📎 <strong>Your Payment Receipt</strong> is attached to this email for your records.
      </div>

      <p>
        <strong>Payment Details:</strong><br>
        Booking Reference: ${referenceNumber}<br>
        Receipt No: ${receiptNumber}<br>
        Amount Paid: ₹${amount.toLocaleString('en-IN')}<br>
        Transaction ID: ${transactionId}
      </p>

      <p style="margin-top: 25px;">
        We look forward to seeing you soon! If you have any questions, feel free to reach out.
      </p>
    </div>

    <div class="footer">
      <p><strong>Kings Equestrian Foundation</strong></p>
      <p>📍 Karnataka, India</p>
      <p>📞 +91-9980895533 | ✉️ info@kingsequestrian.com</p>
      <p style="margin-top: 10px; font-size: 11px;">
        © ${new Date().getFullYear()} Kings Equestrian Foundation. All rights reserved.
      </p>
    </div>
  </div>
</body>
</html>
      `;
      
      MailApp.sendEmail({
        to: email,
        subject: subject,
        htmlBody: htmlBody,
        attachments: [receiptPDF]
      });
      
      // Update payment sheet
      if (paymentHeaderMap['RECEIPT_SENT'] !== undefined) {
        paymentSheet.getRange(rowIndex, paymentHeaderMap['RECEIPT_SENT'] + 1)
          .setValue('Yes')
          .setBackground('#d4edda')
          .setFontColor('#155724')
          .setFontWeight('bold');
      }
      
      if (paymentHeaderMap['RECEIPT_SENT_TIMESTAMP'] !== undefined) {
        paymentSheet.getRange(rowIndex, paymentHeaderMap['RECEIPT_SENT_TIMESTAMP'] + 1)
          .setValue(new Date())
          .setNumberFormat("dd-MMM-yyyy HH:mm:ss");
      }
      
      if (paymentHeaderMap['PAYMENT_RECEIPT_NO'] !== undefined) {
        paymentSheet.getRange(rowIndex, paymentHeaderMap['PAYMENT_RECEIPT_NO'] + 1)
          .setValue(receiptNumber);
      }
      
      successCount++;
      Logger.log(`✅ Receipt sent to: ${email} for ${referenceNumber}`);
      
      Utilities.sleep(1000);
      
    } catch (error) {
      failCount++;
      const errorMsg = `Row ${rowIndex} (${email || 'no email'}): ${error.message}`;
      errors.push(errorMsg);
      Logger.log(`❌ ${errorMsg}`);
    }
  }
  
  let message = `Complete!\n✅ Sent: ${successCount}\n❌ Failed: ${failCount}`;
  
  if (errors.length > 0) {
    message += '\n\nErrors:\n' + errors.slice(0, 5).join('\n');
    if (errors.length > 5) {
      message += `\n... and ${errors.length - 5} more`;
    }
  }
  
  ui.alert(message);
}

// --------------- RESEND WELCOME EMAIL FUNCTION ---------------

function ResendWelcomeEmail() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bookingSheet = ss.getSheetByName(CONFIG.SHEETS.BOOKING_FORM);
  
  if (!bookingSheet) {
    ui.alert('❌ Booking Form Response sheet not found');
    return;
  }
  
  const selection = bookingSheet.getActiveRange();
  
  if (!selection) {
    ui.alert('Please select rows to resend welcome emails');
    return;
  }
  
  const startRow = selection.getRow();
  const numRows = selection.getNumRows();
  
  if (startRow === 1) {
    ui.alert('Cannot send emails for header row');
    return;
  }
  
  const response = ui.alert(
    'Resend Welcome Emails',
    `Resend welcome emails for ${numRows} row(s)?`,
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  let successCount = 0;
  let failCount = 0;
  
  for (let i = 0; i < numRows; i++) {
    const rowIndex = startRow + i;
    
    try {
      const riderName = bookingSheet.getRange(rowIndex, CONFIG.BOOKING_COLS.RIDER_NAME + 1).getValue();
      const email = bookingSheet.getRange(rowIndex, CONFIG.BOOKING_COLS.EMAIL_ID + 1).getValue();
      const phone = bookingSheet.getRange(rowIndex, CONFIG.BOOKING_COLS.PHONE + 1).getValue();
      const services = bookingSheet.getRange(rowIndex, CONFIG.BOOKING_COLS.SERVICE + 1).getValue();
      const selectedDate = bookingSheet.getRange(rowIndex, CONFIG.BOOKING_COLS.DATE + 1).getValue();
      const timeSlots = bookingSheet.getRange(rowIndex, CONFIG.BOOKING_COLS.TIME_SLOTS + 1).getValue();
      const amount = bookingSheet.getRange(rowIndex, CONFIG.BOOKING_COLS.AMOUNT + 1).getValue();
      const reference = bookingSheet.getRange(rowIndex, CONFIG.BOOKING_COLS.REFERENCE + 1).getValue();
      const upiLink = bookingSheet.getRange(rowIndex, CONFIG.BOOKING_COLS.UPI_LINK + 1).getValue();
      const qrCode = bookingSheet.getRange(rowIndex, CONFIG.BOOKING_COLS.QR_CODE + 1).getValue();
      
      if (!email || !reference) {
        throw new Error('Missing email or reference');
      }
      
      // Parse time slots
      let slotCount = 1;
      if (typeof timeSlots === 'string') {
        slotCount = timeSlots.split(',').length;
      }
      
      // Recalculate breakdown
      const calculation = calculateBookingAmount(services, slotCount);
      
      sendWelcomeEmail({
        name: riderName,
        email: email,
        phone: phone,
        services: services,
        date: selectedDate,
        timeSlots: timeSlots,
        amount: amount,
        reference: reference,
        upiLink: upiLink,
        qrCode: qrCode,
        breakdown: calculation.breakdown,
        row: rowIndex,
        sheet: bookingSheet
      });
      
      successCount++;
      Utilities.sleep(1000);
      
    } catch (error) {
      failCount++;
      Logger.log(`❌ Error at row ${rowIndex}: ${error.message}`);
    }
  }
  
  ui.alert(`Complete!\n✅ Sent: ${successCount}\n❌ Failed: ${failCount}`);
}

// --------------- HELPER FUNCTIONS ---------------

function formatDate(date) {
  if (!date) return 'N/A';
  if (typeof date === 'string') return date;
  return Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), 'dd MMM yyyy');
}

// --------------- MENU SETUP ---------------

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🐎 Kings Equestrian')
    .addItem('📧 Resend Welcome Email', 'ResendWelcomeEmail')
    .addItem('🧾 Send Payment Receipt', 'SendPaymentReceipt')
    .addSeparator()
    .addItem('⚙️ Setup Triggers', 'setupTriggers')
    .addToUi();
}

function setupTriggers() {
  // Delete existing triggers
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
  
  // Create new form submit trigger for booking form
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger('onBookingFormSubmit')
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();
  
  SpreadsheetApp.getUi().alert('✅ Triggers set up successfully!');
}


// --------------- DATE/TIME CHANGE REQUEST HANDLER ---------------

/**
 * Web App Handler - Deploy this as a web app
 * Deployment: Deploy → New deployment → Web app
 * Execute as: Me
 * Who has access: Anyone
 */
function doGet(e) {
  const referenceNo = e.parameter.ref;
  
  if (!referenceNo) {
    return HtmlService.createHtmlOutput('<h3>Invalid request - Missing reference number</h3>');
  }
  
  // Get booking details
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bookingSheet = ss.getSheetByName(CONFIG.SHEETS.BOOKING_FORM);
  const headerMap = getHeaderIndexMap(bookingSheet);
  const data = bookingSheet.getDataRange().getValues();
  
  let bookingData = null;
  let rowIndex = -1;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][headerMap['Reference']] === referenceNo) {
      bookingData = {
        name: data[i][headerMap['Rider Name']],
        email: data[i][headerMap['Email ID']],
        service: data[i][headerMap['Service']],
        currentDate: data[i][headerMap['Date']],
        currentTimeSlots: data[i][headerMap['Time Slots']],
        reference: referenceNo
      };
      rowIndex = i + 1;
      break;
    }
  }
  
  if (!bookingData) {
    return HtmlService.createHtmlOutput('<h3>Booking not found</h3>');
  }
  
  // Create HTML form
  const template = HtmlService.createTemplateFromFile('ChangeRequestForm');
  template.bookingData = bookingData;
  
  return template.evaluate()
    .setTitle('Change Date/Time - Kings Equestrian')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Process change request from web form
 */
function processChangeRequest(formData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const bookingSheet = ss.getSheetByName(CONFIG.SHEETS.BOOKING_FORM);
    const headerMap = getHeaderIndexMap(bookingSheet);
    const data = bookingSheet.getDataRange().getValues();
    
    // Find booking row
    let rowIndex = -1;
    let oldDate = '';
    let oldTimeSlots = '';
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][headerMap['Reference']] === formData.reference) {
        rowIndex = i + 1;
        oldDate = data[i][headerMap['Date']];
        oldTimeSlots = data[i][headerMap['Time Slots']];
        break;
      }
    }
    
    if (rowIndex === -1) {
      return { success: false, message: 'Booking not found' };
    }
    
    // Add change tracking columns if they don't exist
    if (headerMap['Change Requests'] === undefined) {
      const lastCol = bookingSheet.getLastColumn();
      bookingSheet.getRange(1, lastCol + 1).setValue('Change Requests');
      bookingSheet.getRange(1, lastCol + 2).setValue('Last Changed Date');
      bookingSheet.getRange(1, lastCol + 3).setValue('Change History');
      headerMap['Change Requests'] = lastCol;
      headerMap['Last Changed Date'] = lastCol + 1;
      headerMap['Change History'] = lastCol + 2;
    }
    
    // Create change log entry
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yyyy HH:mm:ss');
    const changeLog = `[${timestamp}] Date: ${formatDate(oldDate)} → ${formData.newDate} | Time: ${oldTimeSlots} → ${formData.newTimeSlots} | Reason: ${formData.reason}`;
    
    // Get existing history
    const existingHistory = bookingSheet.getRange(rowIndex, headerMap['Change History'] + 1).getValue();
    const newHistory = existingHistory ? existingHistory + '\n' + changeLog : changeLog;
    
    // Update booking sheet
    bookingSheet.getRange(rowIndex, headerMap['Date'] + 1).setValue(formData.newDate);
    bookingSheet.getRange(rowIndex, headerMap['Time Slots'] + 1).setValue(formData.newTimeSlots);
    bookingSheet.getRange(rowIndex, headerMap['Change Requests'] + 1).setValue('Yes').setBackground('#fff3cd');
    bookingSheet.getRange(rowIndex, headerMap['Last Changed Date'] + 1).setValue(new Date()).setNumberFormat('dd-MMM-yyyy HH:mm:ss');
    bookingSheet.getRange(rowIndex, headerMap['Change History'] + 1).setValue(newHistory);
    
    // Recalculate amount if time slots changed
    const services = data[rowIndex - 1][headerMap['Service']];
    const newSlotCount = formData.newTimeSlots.split(',').length;
    const calculation = calculateBookingAmount(services, newSlotCount);
    const newAmount = calculation.totalAmount;
    
    bookingSheet.getRange(rowIndex, headerMap['Amount'] + 1).setValue(newAmount);
    
    // Update Google Calendar event if exists
    const calendarEventId = bookingSheet.getRange(rowIndex, headerMap['Calendar Event ID'] + 1).getValue();
    if (calendarEventId) {
      updateCalendarEvent(calendarEventId, formData.newDate, formData.newTimeSlots, formData.reference);
    }
    
    // Send confirmation email
    sendChangeConfirmationEmail({
      email: formData.email,
      name: formData.name,
      reference: formData.reference,
      oldDate: formatDate(oldDate),
      newDate: formData.newDate,
      oldTimeSlots: oldTimeSlots,
      newTimeSlots: formData.newTimeSlots,
      newAmount: newAmount,
      reason: formData.reason
    });
    
    return { 
      success: true, 
      message: 'Your date/time change request has been confirmed! A confirmation email has been sent.',
      newAmount: newAmount
    };
    
  } catch (error) {
    Logger.log('Error in processChangeRequest: ' + error);
    return { success: false, message: 'Error processing request: ' + error.message };
  }
}

/**
 * Send change confirmation email
 */
function sendChangeConfirmationEmail(data) {
  const subject = `Booking Updated - ${data.reference} - Kings Equestrian`;
  
  const htmlBody = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <style>
    body {
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
      background-color: #f4f4f4;
      margin: 0;
      padding: 0;
      color: #333;
    }
    .container {
      max-width: 650px;
      margin: 20px auto;
      background-color: #ffffff;
      border-radius: 12px;
      overflow: hidden;
      box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .header {
      background: linear-gradient(135deg, #1f4e3d 0%, #4f9c7a 100%);
      padding: 30px;
      text-align: center;
      color: #fff;
    }
    .header img {
      width: 70px;
      height: 70px;
      border-radius: 10%;
      margin-bottom: 10px;
    }
    .header h1 {
      margin: 0;
      font-size: 26px;
    }
    .content {
      padding: 35px 30px;
    }
    .change-box {
      background: #e7f3ff;
      border: 2px solid #0066cc;
      padding: 20px;
      border-radius: 10px;
      margin: 20px 0;
    }
    .change-item {
      display: flex;
      align-items: center;
      margin: 15px 0;
      font-size: 15px;
    }
    .old-value {
      text-decoration: line-through;
      color: #999;
      margin-right: 15px;
    }
    .new-value {
      color: #2c5f2d;
      font-weight: bold;
    }
    .arrow {
      margin: 0 10px;
      color: #0066cc;
      font-size: 18px;
    }
    .footer {
      background-color: #1f4e3d;
      color: #fff;
      padding: 25px;
      text-align: center;
      font-size: 13px;
    }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <img src="https://kingsfarmequestrian.com/wp-content/uploads/2023/08/Logo2.jpg" alt="Kings Equestrian Logo">
      <h1>Kings Equestrian Foundation</h1>
      <p style="margin:8px 0 0;">Where horses don't just carry you — they change you</p>
    </div>

    <div class="content">
      <h2 style="color: #2c5f2d;">✅ Booking Updated Successfully!</h2>
      
      <p>Dear ${data.name},</p>
      
      <p>Your booking date/time change request has been confirmed. Here are the updated details:</p>
      
      <div class="change-box">
        <h3 style="margin-top: 0; color: #0066cc;">📅 Changes Made</h3>
        
        <div class="change-item">
          <strong>Date:</strong>
          <span class="arrow">→</span>
          <span class="old-value">${data.oldDate}</span>
          <span class="arrow">➜</span>
          <span class="new-value">${data.newDate}</span>
        </div>
        
        <div class="change-item">
          <strong>Time Slots:</strong>
          <span class="arrow">→</span>
          <span class="old-value">${data.oldTimeSlots}</span>
          <span class="arrow">➜</span>
          <span class="new-value">${data.newTimeSlots}</span>
        </div>
        
        <div style="margin-top: 20px; padding-top: 15px; border-top: 1px solid #ddd;">
          <p style="margin: 5px 0;"><strong>Booking Reference:</strong> ${data.reference}</p>
          ${data.newAmount ? `<p style="margin: 5px 0;"><strong>Updated Amount:</strong> ₹${data.newAmount}</p>` : ''}
        </div>
        
        ${data.reason ? `
        <div style="margin-top: 15px; padding: 15px; background: #f9f9f9; border-radius: 5px;">
          <strong>Your Reason:</strong><br>
          <em>${data.reason}</em>
        </div>
        ` : ''}
      </div>
      
      <p><strong>Important:</strong></p>
      <ul>
        <li>Your calendar event has been updated</li>
        <li>No additional payment required if amount hasn't changed</li>
        <li>If you've already paid, your payment is still valid</li>
        <li>Please arrive 15 minutes before your new scheduled time</li>
      </ul>
      
      <p>If you have any questions or need further assistance, please don't hesitate to contact us.</p>
    </div>

    <div class="footer">
      <p><strong>Kings Equestrian Foundation</strong></p>
      <p>📍 Karnataka, India</p>
      <p>📞 +91-9980895533 | ✉️ info@kingsequestrian.com</p>
      <p style="margin-top: 10px; font-size: 11px;">
        © ${new Date().getFullYear()} Kings Equestrian Foundation. All rights reserved.
      </p>
    </div>
  </div>
</body>
</html>
  `;
  
  MailApp.sendEmail({
    to: data.email,
    subject: subject,
    htmlBody: htmlBody,
    name: "Kings Equestrian Foundation"
  });
}

// --------------- GOOGLE CALENDAR INTEGRATION ---------------

/**
 * Create calendar event when booking is confirmed
 */
function createCalendarEvent(bookingData) {
  try {
    const calendar = CalendarApp.getDefaultCalendar();
    
    // Parse date and first time slot
    const date = new Date(bookingData.date);
    const timeSlots = bookingData.timeSlots.split(',');
    const firstSlot = timeSlots[0].trim();
    
    // Parse time (assuming format like "10:00 AM")
    const timeParts = firstSlot.match(/(\d+):(\d+)\s*(AM|PM)/i);
    if (!timeParts) {
      Logger.log('Invalid time format');
      return null;
    }
    
    let hours = parseInt(timeParts[1]);
    const minutes = parseInt(timeParts[2]);
    const period = timeParts[3].toUpperCase();
    
    if (period === 'PM' && hours !== 12) hours += 12;
    if (period === 'AM' && hours === 12) hours = 0;
    
    const startTime = new Date(date);
    startTime.setHours(hours, minutes, 0);
    
    // Calculate end time (30 minutes per slot)
    const endTime = new Date(startTime);
    endTime.setMinutes(endTime.getMinutes() + (timeSlots.length * 30));
    
    // Create event
    const event = calendar.createEvent(
      `Kings Equestrian - ${bookingData.name} (${bookingData.reference})`,
      startTime,
      endTime,
      {
        description: `Service: ${bookingData.services}\nReference: ${bookingData.reference}\nPhone: ${bookingData.phone}\nEmail: ${bookingData.email}`,
        location: 'Kings Equestrian Foundation, Karnataka',
        guests: bookingData.email,
        sendInvites: true
      }
    );
    
    Logger.log('Calendar event created: ' + event.getId());
    return event.getId();
    
  } catch (error) {
    Logger.log('Error creating calendar event: ' + error);
    return null;
  }
}

/**
 * Update existing calendar event
 */
function updateCalendarEvent(eventId, newDate, newTimeSlots, reference) {
  try {
    const calendar = CalendarApp.getDefaultCalendar();
    const event = calendar.getEventById(eventId);
    
    if (!event) {
      Logger.log('Event not found: ' + eventId);
      return false;
    }
    
    // Parse new date and time
    const date = new Date(newDate);
    const timeSlots = newTimeSlots.split(',');
    const firstSlot = timeSlots[0].trim();
    
    const timeParts = firstSlot.match(/(\d+):(\d+)\s*(AM|PM)/i);
    if (!timeParts) {
      Logger.log('Invalid time format');
      return false;
    }
    
    let hours = parseInt(timeParts[1]);
    const minutes = parseInt(timeParts[2]);
    const period = timeParts[3].toUpperCase();
    
    if (period === 'PM' && hours !== 12) hours += 12;
    if (period === 'AM' && hours === 12) hours = 0;
    
    const startTime = new Date(date);
    startTime.setHours(hours, minutes, 0);
    
    const endTime = new Date(startTime);
    endTime.setMinutes(endTime.getMinutes() + (timeSlots.length * 30));
    
    // Update event
    event.setTime(startTime, endTime);
    event.setDescription(event.getDescription() + `\n\nUpdated: ${new Date().toLocaleString()}`);
    
    Logger.log('Calendar event updated: ' + eventId);
    return true;
    
  } catch (error) {
    Logger.log('Error updating calendar event: ' + error);
    return false;
  }
}