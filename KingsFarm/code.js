// ============================================
// KINGS EQUESTRIAN - ENHANCED BOOKING SYSTEM 
// ============================================

// --------------- CONFIG ---------------

const CONFIG = {
  UPI_ID: "vyapar.176548151976@hdfcbank",
  BUSINESS_NAME: "KingsEquestrian",
  PAYMENT_FORM_LINK: "https://forms.gle/XXhQ5oaLt65VBfwK6",
  EMAIL_TEMPLATE_DOC_ID: "1bUTpk9QCR4n1uUmMuoSRRcflTShG3jawuhemE28aTio",
  webAppUrl:"https://script.google.com/macros/s/AKfycbxGNi137N_vvd6kFWe0CL2clALwKLp7QKsLgiWUd9fGcvYhTlaeQIy15n2vai_1g-PIig/exec",
  
  SHEETS: {
    BOOKING_FORM: "Booking Form Response",
    PAYMENT_FORM: "Payment Form Response",
    PRICING: "Pricing"
  },
  
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
    NUMBER_OF_PARTICIPANTS: 9,
    AMOUNT: 10,
    REFERENCE: 11,
    UPI_LINK: 12,
    QR_CODE: 13,
    PAYMENT_STATUS: 14,
    RECEIPT_NO: 15,
    WELCOME_EMAIL_SENT: 16,
    WELCOME_EMAIL_TIMESTAMP: 17,
    CALENDAR_EVENT_ID: 18,
    CHANGE_REQUESTS: 19,
    LAST_CHANGED_DATE: 20,
    CHANGE_HISTORY: 21
  },
  
  PAYMENT_COLS: {
    TIMESTAMP: 0,
    EMAIL_ADDRESS: 1,
    REGISTRATION_NO: 2,
    AMOUNT_PAID: 3,
    SCREENSHOT: 4,
    PAYMENT_DATE: 5,
    PAN_AADHAAR: 6,
    TRANSACTION_VERIFIED: 7,
    RECEIPT_SENT: 8,
    RECEIPT_SENT_TIMESTAMP: 9,
    PAYMENT_RECEIPT_NO: 10
  }
};

// --------------- UTILITY FUNCTIONS ---------------

function generateReference() {
  const date = new Date();
  const year = date.getFullYear().toString().substr(-2);
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  const random = Math.floor(Math.random() * 9000) + 1000;
  
  return `KE${year}${month}${day}${random}`;
}

function createUPILink(amount, reference) {
  return `upi://pay?pa=${CONFIG.UPI_ID}&pn=${encodeURIComponent(CONFIG.BUSINESS_NAME)}&am=${amount}&cu=INR&tn=${encodeURIComponent(reference)}`;
}

function createQRCode(link) {
  return `https://api.qrserver.com/v1/create-qr-code/?size=400x400&data=${encodeURIComponent(link)}`;
}

function getPricingData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pricingSheet = ss.getSheetByName(CONFIG.SHEETS.PRICING);
  
  if (!pricingSheet) {
    throw new Error('Pricing sheet not found');
  }
  
  const data = pricingSheet.getDataRange().getValues();
  const pricingMap = {};
  
  for (let i = 1; i < data.length; i++) {
    const serviceId = data[i][0];
    const service = data[i][1];
    const pricePerHalfHour = data[i][2];
    const docId = data[i][3];
    
    if (service) {
      pricingMap[service] = {
        price: pricePerHalfHour,
        docId: docId
      };
    }
  }
  
  return pricingMap;
}

function getServiceDetailsFromDoc(docId) {
  try {
    const doc = DocumentApp.openById(docId);
    const body = doc.getBody();
    const text = body.getText();
    
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
 * ENHANCED: Calculate booking amount with proper participants handling
 */
function calculateBookingAmount(services, totalSlots, participants) {
  const pricingData = getPricingData();
  
  // Ensure safe values
  const safeParticipants = Math.max(1, Number(participants) || 1);
  const safeTotalSlots = Math.max(1, Number(totalSlots) || 1);
  
  // Parse services - handle both string and array input
  const serviceList = Array.isArray(services)
    ? services.map(s => String(s).trim()).filter(Boolean)
    : String(services || '').split(',').map(s => s.trim()).filter(Boolean);
  
  if (!serviceList.length) {
    throw new Error('No service selected');
  }
  
  const numServices = serviceList.length;
  const slotsPerService = Math.floor(safeTotalSlots / numServices);
  const remainingSlots = safeTotalSlots % numServices;
  
  let totalAmount = 0;
  let needsContact = false;
  const breakdown = [];
  
  serviceList.forEach((service, index) => {
    const slots = slotsPerService + (index < remainingSlots ? 1 : 0);
    const pricing = pricingData[service];
    
    if (!pricing) {
      throw new Error(`Pricing not found for service: ${service}`);
    }
    
    const price = Number(pricing.price);
    if (Number.isNaN(price)) {
      throw new Error(`Invalid pricing configured for service: ${service}`);
    }
    
    // Check if pricing requires contact (negative price)
    if (price < 0) {
      needsContact = true;
      breakdown.push({
        service: service,
        slots: slots,
        participants: safeParticipants,
        pricePerSlot: 'Contact us',
        pricePerSlotPerParticipant: 'Contact us',
        totalPricePerSlot: 0,
        amount: 0,
        docId: pricing.docId,
        needsContact: true
      });
      return;
    }
    
    // Calculate: price per slot per participant * participants * slots
    const amountPerSlot = price * safeParticipants;
    const amount = amountPerSlot * slots;
    totalAmount += amount;
    
    breakdown.push({
      service: service,
      slots: slots,
      participants: safeParticipants,
      pricePerSlot: price,
      pricePerSlotPerParticipant: price,
      totalPricePerSlot: amountPerSlot,
      amount: amount,
      docId: pricing.docId,
      needsContact: false
    });
  });
  
  return {
    needsContact: needsContact,
    totalAmount: totalAmount,
    breakdown: breakdown,
    message: needsContact ? 'One or more services require direct pricing confirmation' : ''
  };
}

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
    
    // FIXED: Properly get participants value
    const participantsValue = sheet.getRange(row, CONFIG.BOOKING_COLS.NUMBER_OF_PARTICIPANTS + 1).getValue();
    const participants = participantsValue ? Number(participantsValue) : 1;
    
    // Parse time slots count
    let slotCount = 1;
    if (typeof timeSlots === 'string') {
      slotCount = timeSlots.split(',').length;
    }
    
    // Calculate amount with breakdown
    const calculation = calculateBookingAmount(services, slotCount, participants);
    const amount = calculation.totalAmount;
    const needsContact = calculation.needsContact;
    
    // Generate payment details
    const reference = generateReference();
    const upiLink = needsContact ? '' : createUPILink(amount, reference);
    const qrCode = needsContact ? '' : createQRCode(upiLink);
    const paymentStatus = needsContact ? 'Contact for Pricing' : 'Pending';
    
    // Update sheet
    sheet.getRange(row, CONFIG.BOOKING_COLS.AMOUNT + 1).setValue(needsContact ? '' : amount);
    sheet.getRange(row, CONFIG.BOOKING_COLS.REFERENCE + 1).setValue(reference);
    sheet.getRange(row, CONFIG.BOOKING_COLS.UPI_LINK + 1).setValue(upiLink);
    sheet.getRange(row, CONFIG.BOOKING_COLS.QR_CODE + 1).setValue(qrCode);
    sheet.getRange(row, CONFIG.BOOKING_COLS.PAYMENT_STATUS + 1).setValue(paymentStatus);
    
    // Create calendar event
    const calendarEventId = createCalendarEvent({
      name: riderName,
      email: email,
      phone: phone,
      services: services,
      date: selectedDate,
      timeSlots: timeSlots,
      reference: reference,
      participants: participants
    });

    if (calendarEventId) {
      const headerMap = getHeaderIndexMap(sheet);
      if (headerMap['Calendar Event ID'] === undefined) {
        const lastCol = sheet.getLastColumn();
        sheet.getRange(1, lastCol + 1).setValue('Calendar Event ID');
      }
      sheet.getRange(row, CONFIG.BOOKING_COLS.CALENDAR_EVENT_ID + 1).setValue(calendarEventId);
    }
    
    // Send welcome email
    sendWelcomeEmail({
      name: riderName,
      email: email,
      phone: phone,
      services: services,
      date: selectedDate,
      timeSlots: timeSlots,
      participants: participants,
      amount: amount,
      reference: reference,
      upiLink: upiLink,
      qrCode: qrCode,
      needsContact: needsContact,
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

// --------------- ENHANCED EMAIL FUNCTIONS ---------------

/**
 * ENHANCED: Send welcome email with proper participants display
 */
function sendWelcomeEmail(data) {
  const subject = `Welcome to Kings Equestrian - Booking ${data.reference}`;
  const breakdownItems = Array.isArray(data.breakdown) ? data.breakdown : [];
  const needsContact = Boolean(data.needsContact) || breakdownItems.some(item => item && item.needsContact);
  const participants = data.participants || 1;

  let serviceDetailsHTML = '';
  const attachments = [];

  breakdownItems.forEach(item => {
    const details = getServiceDetailsFromDoc(item.docId);
    
    // ENHANCED: Show participants in service details
    const participantText = item.participants > 1 ? ` (${item.participants} participants)` : '';
    
    const costText = item.needsContact
      ? '<span style="color: #ff6b00; font-weight: bold;">Contact us for pricing</span>'
      : `₹${item.pricePerSlot} per slot per participant × ${item.participants} participant${item.participants > 1 ? 's' : ''} × ${item.slots} slot${item.slots > 1 ? 's' : ''} = <strong>₹${item.amount.toLocaleString('en-IN')}</strong>`;

    serviceDetailsHTML += `
      <div style="margin: 15px 0; padding: 15px; background: #f9f9f9; border-radius: 8px;">
        <h3 style="color: #2c5f2d; margin: 0 0 10px 0;">${item.service}${participantText}</h3>
        <p style="margin: 5px 0; font-size: 14px;">${details.summary}</p>
        <p style="margin: 5px 0; color: #666; font-size: 13px;">
          <strong>Slots allocated:</strong> ${item.slots} × 30 minutes<br>
          <strong>Cost breakdown:</strong> ${costText}
        </p>
      </div>
    `;

    const pdf = getServicePDF(item.docId, item.service);
    if (pdf) attachments.push(pdf);
  });

  // ENHANCED: Breakdown table with participants column
  let breakdownTable = `
    <table style="width: 100%; border-collapse: collapse; margin: 20px 0; font-size: 14px;">
      <thead>
        <tr style="background: #2c5f2d; color: white;">
          <th style="padding: 12px; text-align: left; border: 1px solid #ddd;">Service</th>
          <th style="padding: 12px; text-align: center; border: 1px solid #ddd;">Participants</th>
          <th style="padding: 12px; text-align: center; border: 1px solid #ddd;">Slots</th>
          <th style="padding: 12px; text-align: right; border: 1px solid #ddd;">Rate/Slot</th>
          <th style="padding: 12px; text-align: right; border: 1px solid #ddd;">Amount</th>
        </tr>
      </thead>
      <tbody>
  `;

  breakdownItems.forEach(item => {
    const rateDisplay = item.needsContact 
      ? '<span style="color: #ff6b00;">Contact us</span>' 
      : `₹${item.pricePerSlot}`;
    const amountDisplay = item.needsContact 
      ? '<span style="color: #ff6b00;">Contact us</span>' 
      : `₹${item.amount.toLocaleString('en-IN')}`;

    breakdownTable += `
      <tr>
        <td style="padding: 10px; border: 1px solid #ddd;">${item.service}</td>
        <td style="padding: 10px; text-align: center; border: 1px solid #ddd;">${item.participants}</td>
        <td style="padding: 10px; text-align: center; border: 1px solid #ddd;">${item.slots}</td>
        <td style="padding: 10px; text-align: right; border: 1px solid #ddd;">${rateDisplay}</td>
        <td style="padding: 10px; text-align: right; border: 1px solid #ddd;">${amountDisplay}</td>
      </tr>
    `;
  });

  breakdownTable += `
      <tr style="background: #f0f0f0; font-weight: bold;">
        <td colspan="4" style="padding: 12px; text-align: right; border: 1px solid #ddd;">Total Amount:</td>
        <td style="padding: 12px; text-align: right; border: 1px solid #ddd; color: #2c5f2d; font-size: 16px;">
          ${needsContact ? '<span style="color: #ff6b00;">Contact us</span>' : `₹${data.amount.toLocaleString('en-IN')}`}
        </td>
      </tr>
      </tbody>
    </table>
  `;

  const paymentSection = needsContact
    ? `
      <div style="background: #fff9e6; border: 2px solid #ffc107; padding: 20px; border-radius: 8px; margin: 20px 0;">
        <h3 style="margin-top: 0; color: #856404;">📞 Pricing Confirmation Required</h3>
        <p>One or more selected services require custom pricing based on your requirements.</p>
        <p><strong>Next Steps:</strong></p>
        <ol style="margin: 10px 0; padding-left: 20px;">
          <li>Our team will contact you within 24 hours</li>
          <li>We'll discuss your specific needs and provide exact pricing</li>
          <li>Once confirmed, you'll receive payment details</li>
        </ol>
        <p style="margin-top: 15px;">
          For immediate assistance, contact us at:<br>
          📞 <strong>+91-9980895533</strong><br>
          ✉️ <a href="mailto:info@kingsequestrian.com">info@kingsequestrian.com</a><br>
          <strong>Reference:</strong> ${data.reference}
        </p>
      </div>
    `
    : `
      <div style="background: #e8f5e9; border: 2px solid #4caf50; padding: 20px; border-radius: 8px; margin: 20px 0;">
        <h3 style="margin-top: 0; color: #2e7d32;">💳 Complete Your Payment</h3>
        <p style="font-size: 16px;">To confirm your booking, please pay:</p>
        <p style="text-align: center; font-size: 32px; font-weight: bold; color: #2c5f2d; margin: 15px 0;">
          ₹${data.amount.toLocaleString('en-IN')}
        </p>
        
        <div style="display: flex; gap: 15px; margin: 20px 0; flex-wrap: wrap;">
          <div style="flex: 1; min-width: 200px; background: white; padding: 15px; border-radius: 8px; text-align: center;">
            // <p style="margin: 0 0 10px 0; font-weight: bold;">Quick Pay via UPI</p>
            // <a href="${data.upiLink}" style="display: inline-block; background: #2c5f2d; color: white; padding: 12px 24px; text-decoration: none; border-radius: 5px; font-weight: bold;">
            //   Pay Now
            // </a>
          </div>
          
          <div style="flex: 1; min-width: 200px; background: white; padding: 15px; border-radius: 8px; text-align: center;">
            <p style="margin: 0 0 10px 0; font-weight: bold;">Scan QR Code</p>
            <img src="${data.qrCode}" alt="QR Code" style="width: 150px; height: 150px;">
          </div>
        </div>
        
        <div style="background: #fff3cd; padding: 15px; border-radius: 5px; margin-top: 15px;">
          <p style="margin: 0; font-size: 14px;"><strong>⚠️ Important:</strong></p>
          <p style="margin: 5px 0 0 0; font-size: 13px;">
            After making payment, please submit your payment details at:<br>
            <a href="${CONFIG.PAYMENT_FORM_LINK}" style="color: #856404; font-weight: bold;">Submit Payment details</a>
          </p>
        </div>
      </div>
    `;

  const htmlBody = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
    </head>
    <body style="font-family: Arial, sans-serif; color: #333; line-height: 1.6; margin: 0; padding: 0; background: #f5f5f5;">
      <div style="max-width: 650px; margin: 20px auto; background: white; border-radius: 10px; overflow: hidden; box-shadow: 0 2px 10px rgba(0,0,0,0.1);">
        
        <!-- Header -->
        <div style="background: linear-gradient(135deg, #1f4e3d 0%, #4f9c7a 100%); padding: 30px; text-align: center; color: white;">
          <img src="https://kingsfarmequestrian.com/wp-content/uploads/2023/08/Logo2.jpg" alt="Kings Equestrian" style="width: 80px; height: 80px; border-radius: 50%; margin-bottom: 15px;">
          <h1 style="margin: 0; font-size: 28px;">Welcome to Kings Equestrian!</h1>
          <p style="margin: 10px 0 0 0; font-size: 14px; opacity: 0.9;">Where horses don't just carry you — they change you</p>
        </div>
        
        <!-- Content -->
        <div style="padding: 30px;">
          <h2 style="color: #2c5f2d; margin-top: 0;">Hello ${data.name}! 👋</h2>
          
          <p>Thank you for choosing Kings Equestrian Foundation. Your booking has been received and is being processed.</p>
          
          <div style="background: #f0f8ff; border-left: 4px solid #2c5f2d; padding: 15px; margin: 20px 0;">
            <p style="margin: 0; font-size: 14px;">
              <strong>Booking Reference:</strong> <span style="font-size: 18px; color: #2c5f2d; font-weight: bold;">${data.reference}</span><br>
              <strong>Date:</strong> ${formatDate(data.date)}<br>
              <strong>Time Slots:</strong> ${data.timeSlots}<br>
              <strong>Participants:</strong> ${participants}
            </p>
          </div>
          
          <h3 style="color: #2c5f2d; border-bottom: 2px solid #2c5f2d; padding-bottom: 10px;">📋 Service Details</h3>
          ${serviceDetailsHTML}
          
          <h3 style="color: #2c5f2d; border-bottom: 2px solid #2c5f2d; padding-bottom: 10px;">💰 Payment Breakdown</h3>
          ${breakdownTable}
          
          ${paymentSection}
          
          <div style="background: #f9f9f9; padding: 20px; border-radius: 8px; margin-top: 20px;">
            <h4 style="margin: 0 0 10px 0; color: #2c5f2d;">📌 What's Next?</h4>
            <ul style="margin: 0; padding-left: 20px;">
              ${needsContact 
                ? '<li>Our team will contact you to confirm pricing</li><li>Once pricing is confirmed, you\'ll receive payment details</li>'
                : '<li>Complete your payment using the options above</li><li>Submit payment confirmation through the form</li>'
              }
              <li>Arrive 15 minutes before your scheduled time</li>
              <li>Bring comfortable clothing and closed-toe shoes</li>
            </ul>
          </div>
          
          <p style="margin-top: 20px; font-size: 14px; color: #666;">
            If you have any questions, feel free to reach out to us anytime.
          </p>
        </div>
        
        <!-- Footer -->
        <div style="background: #1f4e3d; color: white; padding: 20px; text-align: center; font-size: 13px;">
          <p style="margin: 0 0 10px 0;"><strong>Kings Equestrian Foundation</strong></p>
          <p style="margin: 0;">📍 Karnataka, India</p>
          <p style="margin: 5px 0;">📞 +91-9980895533 | ✉️ info@kingsequestrian.com</p>
          <p style="margin: 10px 0 0 0; opacity: 0.8; font-size: 11px;">
            © ${new Date().getFullYear()} Kings Equestrian Foundation. All rights reserved.
          </p>
        </div>
      </div>
    </body>
    </html>
  `;

  const plainBody = `
Welcome to Kings Equestrian Foundation!

Dear ${data.name},

Your booking reference: ${data.reference}

BOOKING DETAILS:
Name: ${data.name}
Contact: ${data.phone}
Date: ${formatDate(data.date)}
Time Slots: ${data.timeSlots}
Participants: ${participants}

TOTAL AMOUNT: ${needsContact ? 'Contact us for pricing' : `₹${data.amount.toLocaleString('en-IN')}`}

PAYMENT BREAKDOWN:
${breakdownItems.map(item => {
  if (item.needsContact) {
    return `${item.service}: ${item.slots} slots × ${item.participants} participants - Contact us for pricing`;
  }
  return `${item.service}: ${item.slots} slots × ${item.participants} participant(s) × ₹${item.pricePerSlot} = ₹${item.amount.toLocaleString('en-IN')}`;
}).join('\n')}

${needsContact
  ? `PRICING CONFIRMATION REQUIRED:\nOur team will contact you within 24 hours to discuss pricing.\nFor immediate assistance: +91-9980895533 or info@kingsequestrian.com`
  : `PAYMENT INSTRUCTIONS:\n1. Pay ₹${data.amount.toLocaleString('en-IN')} using UPI: ${data.upiLink}\n2. Or scan QR code: ${data.qrCode}\n3. Submit payment details: ${CONFIG.PAYMENT_FORM_LINK}`
}

WHAT'S NEXT:
${needsContact ? '- Our team will contact you to confirm pricing' : '- Complete your payment'}
${needsContact ? '- You\'ll receive payment details once confirmed' : '- Submit payment confirmation through the form'}
- Arrive 15 minutes before your scheduled time
- Bring comfortable clothing and closed-toe shoes

Kings Equestrian Foundation
Karnataka, India
+91-9980895533 | info@kingsequestrian.com
  `;

  MailApp.sendEmail({
    to: data.email,
    subject: subject,
    body: plainBody,
    htmlBody: htmlBody,
    attachments: attachments,
    name: 'Kings Equestrian Foundation'
  });

  if (data.sheet && data.row) {
    data.sheet.getRange(data.row, CONFIG.BOOKING_COLS.WELCOME_EMAIL_SENT + 1)
      .setValue('Yes')
      .setBackground('#d4edda')
      .setFontColor('#155724')
      .setFontWeight('bold');

    data.sheet.getRange(data.row, CONFIG.BOOKING_COLS.WELCOME_EMAIL_TIMESTAMP + 1)
      .setValue(new Date())
      .setNumberFormat('dd-MMM-yyyy HH:mm:ss');
  }

  Logger.log(`Welcome email sent to: ${data.email}`);
}

// --------------- RECEIPT GENERATION ---------------

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
    @page { size: A4; margin: 0; }
    body { font-family: Arial, sans-serif; margin: 0; padding: 30px; background: #fff; }
    .receipt-container { border: 2px solid #000; border-radius: 25px; padding: 25px; max-width: 750px; margin: 0 auto; }
    .header { display: flex; align-items: flex-start; margin-bottom: 15px; position: relative; }
    .logo-section { flex: 0 0 130px; text-align: center; }
    .logo-img { width: 100px; height: 100px; margin-bottom: 5px; }
    .logo-text { font-size: 11px; font-weight: bold; line-height: 1.2; }
    .header-center { flex: 1; text-align: center; padding: 0 15px; }
    .org-name { font-size: 22px; font-weight: bold; margin-bottom: 3px; }
    .registration-info { font-size: 10px; margin-bottom: 2px; }
    .location { font-size: 9px; }
    .receipt-number { position: absolute; right: 0; top: 0; background: #ff4444; color: white; font-size: 22px; font-weight: bold; padding: 8px 18px; border-radius: 5px; }
    .receipt-box { border: 1px solid #000; border-radius: 10px; padding: 12px; text-align: center; margin: 15px 0; }
    .receipt-title { font-size: 18px; font-weight: bold; margin-bottom: 3px; }
    .receipt-subtitle { font-size: 9px; font-style: italic; }
    .main-content { display: flex; gap: 20px; }
    .left-column { flex: 1; }
    .right-column { flex: 1; }
    .section-title { font-size: 11px; font-weight: bold; margin-bottom: 8px; }
    .checkbox-list { font-size: 10px; margin-bottom: 12px; }
    .checkbox-item { margin: 4px 0; display: flex; align-items: center; }
    .checkbox { width: 12px; height: 12px; border: 1px solid #000; display: inline-block; margin-right: 6px; }
    .checkbox.checked { background: #000; position: relative; }
    .checkbox.checked::after { content: '✓'; color: white; font-size: 10px; position: absolute; top: -2px; left: 1px; }
    .amount-section { border: 2px solid #000; padding: 15px; margin: 15px 0; position: relative; min-height: 60px; }
    .rupee-symbol { position: absolute; left: 15px; top: 50%; transform: translateY(-50%); font-size: 36px; font-weight: bold; color: #ffa500; }
    .amount-value { text-align: center; font-size: 32px; font-weight: bold; padding-top: 5px; }
    .payment-mode { font-size: 9px; margin: 10px 0 5px 0; font-weight: bold; }
    .declaration-section { font-size: 8.5px; line-height: 1.4; text-align: justify; margin-top: 8px; }
    .donor-details { font-size: 10px; line-height: 1.6; }
    .detail-row { margin: 5px 0; }
    .detail-label { font-weight: bold; }
    .signature-section { margin-top: 30px; display: flex; justify-content: space-between; align-items: flex-end; }
    .left-signature { flex: 1; }
    .right-signature { flex: 1; text-align: center; }
    .org-label { font-size: 11px; font-weight: bold; margin-bottom: 5px; }
    .stamp-and-sign { position: relative; width: 150px; height: 150px; margin:auto 0; }
    .stamp-img { position: absolute; width: 120px; height: 120px; left: 15px; top: 50px; }
    .sign-img { position: absolute; width: 100px; height: 40px; left:25px; top: -45px; }
    .authorized-text { font-size: 10px; margin-top: 70px; text-decoration: underline; }
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
          <div class="section-title">Donation Purpose (✓ Tick Applicable)</div>
          <div class="checkbox-list" style="display:flex;flex-wrap:wrap; gap:10px;">
            <div class="checkbox-item"><span class="checkbox checked"></span><span>All</span></div>
            <div class="checkbox-item"><span class="checkbox"></span><span>Horse welfare</span></div>
            <div class="checkbox-item"><span class="checkbox"></span><span>Recovery training</span></div>
            <div class="checkbox-item"><span class="checkbox"></span><span>Nutrition & feed</span></div>
            <div class="checkbox-item"><span class="checkbox"></span><span>Rehabilitation including timely veterinary care</span></div>
            <div class="checkbox-item"><span class="checkbox"></span><span>Non-commercial equestrian skill and sports development</span></div>
          </div>
          
          <div class="section-title">Donor Category (✓ Tick Applicable)</div>
          <div class="checkbox-list">
            <div class="checkbox-item"><span class="checkbox checked"></span><span>Resident Indian Donor</span></div>
            <div class="checkbox-item"><span class="checkbox"></span><span>Non Resident Indian (NRI)</span></div>
          </div>
        </div>
        
        <div class="right-column">
          <div class="section-title">Donor Details</div>
          <div class="donor-details">
            <div class="detail-row"><span class="detail-label">Date:</span> ${currentDate}</div>
            <div class="detail-row"><span class="detail-label">Name of Donor:</span> ${donorName}</div>
            <div class="detail-row"><span class="detail-label">PAN / Aadhaar:</span> ${pan}</div>
            <div class="detail-row"><span class="detail-label">Amount in Words:</span> ${amountInWords}</div>
          </div>
        </div>
      </div>
      
      <div class="amount-section">
        <span class="rupee-symbol">₹</span>
        <div class="amount-value">${amount.toLocaleString('en-IN')}</div>
      </div>
      
      <div class="payment-mode">Mode of Payment (✓ Tick above): <span>${transactionRef}</span></div>
      <div class="checkbox-list" style="display:flex;flex-wrap:wrap; gap:10px;">
        <div class="checkbox-item"><span class="checkbox checked"></span><span>UPI</span></div>
        <div class="checkbox-item"><span class="checkbox"></span><span>RTGS</span></div>
        <div class="checkbox-item"><span class="checkbox"></span><span>NEFT</span></div>
        <div class="checkbox-item"><span class="checkbox"></span><span>DD</span></div>
        <div class="checkbox-item"><span class="checkbox"></span><span>Cheque</span></div>
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

  if (!paymentSheet || !bookingSheet) {
    ui.alert('Required sheets not found');
    return;
  }

  const selection = paymentSheet.getActiveRange();
  if (!selection || selection.getRow() === 1) {
    ui.alert('Please select valid rows to send receipts');
    return;
  }

  const startRow = selection.getRow();
  const numRows = selection.getNumRows();

  const response = ui.alert(
    'Send Payment Receipts',
    `Send receipts for ${numRows} row(s)?`,
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  const paymentHeaderMap = getHeaderIndexMap(paymentSheet);
  const bookingHeaderMap = getHeaderIndexMap(bookingSheet);
  const bookingValues = bookingSheet.getDataRange().getValues();
  const bookingReferenceCol = bookingHeaderMap['Reference'] !== undefined ? bookingHeaderMap['Reference'] : bookingHeaderMap['Reference No'];
  const bookingNameCol = bookingHeaderMap['Rider Name'] !== undefined ? bookingHeaderMap['Rider Name'] : bookingHeaderMap['Name'];
  const bookingEmailCol = bookingHeaderMap['Email ID'] !== undefined
    ? bookingHeaderMap['Email ID']
    : (bookingHeaderMap['Email Id'] !== undefined ? bookingHeaderMap['Email Id'] : bookingHeaderMap['Email']);
  const bookingPaymentStatusCol = bookingHeaderMap['Payment Status'] !== undefined
    ? bookingHeaderMap['Payment Status']
    : bookingHeaderMap['PAYMENT_STATUS'];

  if (bookingReferenceCol === undefined || bookingNameCol === undefined) {
    ui.alert('Booking sheet is missing required columns');
    return;
  }

  const pickCell = (row, map, headers) => {
    for (let i = 0; i < headers.length; i++) {
      const idx = map[headers[i]];
      if (idx !== undefined) return row[idx];
    }
    return '';
  };

  const findBookingByReference = (referenceNumber) => {
    for (let j = 1; j < bookingValues.length; j++) {
      const bookingRow = bookingValues[j];
      if (String(bookingRow[bookingReferenceCol] || '').trim() === String(referenceNumber || '').trim()) {
        return { rowIndex: j + 1, row: bookingRow };
      }
    }
    return null;
  };

  let successCount = 0;
  let failCount = 0;
  const errors = [];

  for (let i = 0; i < numRows; i++) {
    const rowIndex = startRow + i;
    let email = '';
    let referenceNumber = '';

    try {
      const row = paymentSheet.getRange(rowIndex, 1, 1, paymentSheet.getLastColumn()).getValues()[0];

      referenceNumber = pickCell(row, paymentHeaderMap, [
        'Reference No',
        'Reference Number',
        'Reference',
        'Booking Reference',
        'Registration No',
        'REGISTRATION_NO'
      ]);

      if (!referenceNumber) {
        throw new Error('Reference number missing');
      }

      const bookingMatch = findBookingByReference(referenceNumber);
      if (!bookingMatch) {
        throw new Error(`Booking not found for reference ${referenceNumber}`);
      }

      const riderName = bookingMatch.row[bookingNameCol];
      const bookingEmail = bookingEmailCol !== undefined ? bookingMatch.row[bookingEmailCol] : '';

      email = pickCell(row, paymentHeaderMap, ['Email Id', 'Email ID', 'Email']) || bookingEmail;
      if (!email) {
        throw new Error('Email not found');
      }

      const amountValue = pickCell(row, paymentHeaderMap, ['Amount Paid (₹)', 'Amount Paid (₹)', 'Amount Paid', 'Amount']);
      const amount = Number(amountValue);
      if (!amount || Number.isNaN(amount)) {
        throw new Error('Valid amount is required');
      }

      const transactionId = pickCell(row, paymentHeaderMap, [
        'Transaction Reference ID',
        'Transaction ID',
        'UTR',
        'UTR Number'
      ]) || 'N/A';

      const pan = pickCell(row, paymentHeaderMap, ['Pan / AAdhar Number', 'PAN / Aadhar', 'PAN', 'PAN Number']) || '';

      const verificationHeaders = ['TRANSACTION_VERIFIED', 'Transaction Verified', 'Verified'];
      let transactionVerified = 'Yes';
      for (let v = 0; v < verificationHeaders.length; v++) {
        const idx = paymentHeaderMap[verificationHeaders[v]];
        if (idx !== undefined) {
          transactionVerified = row[idx];
          break;
        }
      }

      if (String(transactionVerified || '').toLowerCase() !== 'yes') {
        throw new Error('Transaction not verified');
      }

      if (bookingPaymentStatusCol !== undefined) {
        bookingSheet.getRange(bookingMatch.rowIndex, bookingPaymentStatusCol + 1)
          .setValue('Paid')
          .setBackground('#d4edda')
          .setFontColor('#155724')
          .setFontWeight('bold');
      }

      const receiptNumber = generateReceiptNumber();
      bookingSheet.getRange(bookingMatch.rowIndex, bookingHeaderMap['Receipt No'] + 1).setValue(receiptNumber);
      
      const receiptPDF = generate80GReceipt(riderName, pan, amount, transactionId, receiptNumber);

      const subject = `Payment Receipt - ${riderName} - Ref: ${referenceNumber}`;
      const htmlBody = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <style>
    body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f4f4f4; margin: 0; padding: 0; color: #333; }
    .container { max-width: 650px; margin: 20px auto; background-color: #ffffff; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
    .header { background: linear-gradient(135deg, #1f4e3d 0%, #4f9c7a 100%); padding: 30px; text-align: center; color: #fff; }
    .header img { width: 70px; height: 70px; border-radius: 10%; margin-bottom: 10px; }
    .header h1 { margin: 0; font-size: 26px; }
    .content { padding: 35px 30px; }
    .greeting { font-size: 18px; font-weight: 600; margin-bottom: 15px; color: #1f4e3d; }
    .info-box { background: #f8f8f8; border-left: 4px solid #1f4e3d; padding: 18px; margin: 20px 0; border-radius: 6px; font-size: 14px; }
    .footer { background-color: #1f4e3d; color: #fff; padding: 25px; text-align: center; font-size: 13px; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <img src="https://kingsfarmequestrian.com/wp-content/uploads/2023/08/Logo2.jpg" alt="Kings Equestrian Logo">
      <h1>Kings Equestrian Foundation</h1>
      <p style="margin:8px 0 0;">Where horses don't just carry you - they change you</p>
    </div>
    <div class="content">
      <div class="greeting">Dear ${riderName},</div>
      <p>Thank you for your payment. Your booking is confirmed.</p>
      <div class="info-box">
        <strong>Your Payment Receipt</strong> is attached to this email.
      </div>
      <p>
        <strong>Payment Details:</strong><br>
        Booking Reference: ${referenceNumber}<br>
        Receipt No: ${receiptNumber}<br>
        Amount Paid: ₹${amount.toLocaleString('en-IN')}<br>
        Transaction ID: ${transactionId}
      </p>
    </div>
    <div class="footer">
      <p><strong>Kings Equestrian Foundation</strong></p>
      <p>Karnataka, India</p>
      <p>+91-9980895533 | info@kingsequestrian.com</p>
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

      const sentCol = paymentHeaderMap['RECEIPT_SENT'] !== undefined ? paymentHeaderMap['RECEIPT_SENT'] : paymentHeaderMap['Receipt Sent'];
      const sentTsCol = paymentHeaderMap['RECEIPT_SENT_TIMESTAMP'] !== undefined ? paymentHeaderMap['RECEIPT_SENT_TIMESTAMP'] : paymentHeaderMap['Receipt Sent Timestamp'];
      const receiptNoCol = paymentHeaderMap['PAYMENT_RECEIPT_NO'] !== undefined ? paymentHeaderMap['PAYMENT_RECEIPT_NO'] : paymentHeaderMap['Payment Receipt No'];

      if (sentCol !== undefined) {
        paymentSheet.getRange(rowIndex, sentCol + 1)
          .setValue('Yes')
          .setBackground('#d4edda')
          .setFontColor('#155724')
          .setFontWeight('bold');
      }

      if (sentTsCol !== undefined) {
        paymentSheet.getRange(rowIndex, sentTsCol + 1)
          .setValue(new Date())
          .setNumberFormat('dd-MMM-yyyy HH:mm:ss');
      }

      if (receiptNoCol !== undefined) {
        paymentSheet.getRange(rowIndex, receiptNoCol + 1).setValue(receiptNumber);
      }

      successCount++;
      Logger.log(`Receipt sent to: ${email} for ${referenceNumber}`);
      Utilities.sleep(1000);
    } catch (error) {
      failCount++;
      const errorMsg = `Row ${rowIndex} (${email || 'no email'}): ${error.message}`;
      errors.push(errorMsg);
      Logger.log(`Receipt failed: ${errorMsg}`);
    }
  }

  let message = `Complete!\n✅ Sent: ${successCount}\n❌ Failed: ${failCount}`;
  if (errors.length > 0) {
    message += '\n\nErrors:\n' + errors.slice(0, 5).join('\n');
    if (errors.length > 5) message += `\n... and ${errors.length - 5} more`;
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
      const reference = bookingSheet.getRange(rowIndex, CONFIG.BOOKING_COLS.REFERENCE + 1).getValue();
      
      // FIXED: Get participants
      const participantsValue = bookingSheet.getRange(rowIndex, CONFIG.BOOKING_COLS.NUMBER_OF_PARTICIPANTS + 1).getValue();
      const participants = participantsValue ? Number(participantsValue) : 1;
      
      if (!email || !reference) {
        throw new Error('Missing email or reference');
      }
      
      let slotCount = 1;
      if (typeof timeSlots === 'string') {
        slotCount = timeSlots.split(',').length;
      }
      
      const calculation = calculateBookingAmount(services, slotCount, participants);
      const amount = calculation.totalAmount;
      const needsContact = calculation.needsContact;
      const upiLink = needsContact ? '' : createUPILink(amount, reference);
      const qrCode = needsContact ? '' : createQRCode(upiLink);
      
      sendWelcomeEmail({
        name: riderName,
        email: email,
        phone: phone,
        services: services,
        date: selectedDate,
        timeSlots: timeSlots,
        participants: participants,
        amount: amount,
        reference: reference,
        upiLink: upiLink,
        qrCode: qrCode,
        needsContact: needsContact,
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
  ui.createMenu('🎠 Kings Equestrian')
    .addItem('📧 Resend Welcome Email', 'ResendWelcomeEmail')
    .addItem('🧾 Send Payment Receipt', 'SendPaymentReceipt')
    .addSeparator()
    .addItem('⚙️ Setup Triggers', 'setupTriggers')
    .addToUi();
}

function setupTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger('onBookingFormSubmit')
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();
  
  SpreadsheetApp.getUi().alert('✅ Triggers set up successfully!');
}

// --------------- DATE/TIME CHANGE REQUEST HANDLER ---------------

function doGet(e) {
  const referenceNo = e.parameter.ref;
  
  if (!referenceNo) {
    return HtmlService.createHtmlOutput('<h3>Invalid request - Missing reference number</h3>');
  }
  
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
        participants: data[i][headerMap['Number of Participants']] || 1,
        reference: referenceNo
      };
      rowIndex = i + 1;
      break;
    }
  }
  
  if (!bookingData) {
    return HtmlService.createHtmlOutput('<h3>Booking not found</h3>');
  }
  
  const template = HtmlService.createTemplateFromFile('ChangeRequestForm');
  template.bookingData = bookingData;
  
  return template.evaluate()
    .setTitle('Change Date/Time - Kings Equestrian')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * ENHANCED: Process change request with participants recalculation
 */
function processChangeRequest(formData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const bookingSheet = ss.getSheetByName(CONFIG.SHEETS.BOOKING_FORM);
    const headerMap = getHeaderIndexMap(bookingSheet);
    const data = bookingSheet.getDataRange().getValues();
    
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
    
    if (headerMap['Change Requests'] === undefined) {
      const lastCol = bookingSheet.getLastColumn();
      bookingSheet.getRange(1, lastCol + 1).setValue('Change Requests');
      bookingSheet.getRange(1, lastCol + 2).setValue('Last Changed Date');
      bookingSheet.getRange(1, lastCol + 3).setValue('Change History');
      headerMap['Change Requests'] = lastCol;
      headerMap['Last Changed Date'] = lastCol + 1;
      headerMap['Change History'] = lastCol + 2;
    }
    
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yyyy HH:mm:ss');
    const changeLog = `[${timestamp}] Date: ${formatDate(oldDate)} → ${formData.newDate} | Time: ${oldTimeSlots} → ${formData.newTimeSlots} | Reason: ${formData.reason}`;
    
    const existingHistory = bookingSheet.getRange(rowIndex, headerMap['Change History'] + 1).getValue();
    const newHistory = existingHistory ? existingHistory + '\n' + changeLog : changeLog;
    
    bookingSheet.getRange(rowIndex, headerMap['Date'] + 1).setValue(formData.newDate);
    bookingSheet.getRange(rowIndex, headerMap['Time Slots'] + 1).setValue(formData.newTimeSlots);
    bookingSheet.getRange(rowIndex, headerMap['Change Requests'] + 1).setValue('Yes').setBackground('#fff3cd');
    bookingSheet.getRange(rowIndex, headerMap['Last Changed Date'] + 1).setValue(new Date()).setNumberFormat('dd-MMM-yyyy HH:mm:ss');
    bookingSheet.getRange(rowIndex, headerMap['Change History'] + 1).setValue(newHistory);
    
    // FIXED: Recalculate with participants
    const services = data[rowIndex - 1][headerMap['Service']];
    const participants = data[rowIndex - 1][headerMap['Number of Participants']] || 1;
    const newSlotCount = formData.newTimeSlots.split(',').length;
    const calculation = calculateBookingAmount(services, newSlotCount, participants);
    const newAmount = calculation.totalAmount;
    
    const reference = data[rowIndex - 1][headerMap['Reference']];
    const currentStatus = bookingSheet.getRange(rowIndex, CONFIG.BOOKING_COLS.PAYMENT_STATUS + 1).getValue();
    const updatedUPILink = calculation.needsContact ? '' : createUPILink(newAmount, reference);
    const updatedQRCode = calculation.needsContact ? '' : createQRCode(updatedUPILink);
    
    bookingSheet.getRange(rowIndex, CONFIG.BOOKING_COLS.AMOUNT + 1).setValue(calculation.needsContact ? '' : newAmount);
    bookingSheet.getRange(rowIndex, CONFIG.BOOKING_COLS.UPI_LINK + 1).setValue(updatedUPILink);
    bookingSheet.getRange(rowIndex, CONFIG.BOOKING_COLS.QR_CODE + 1).setValue(updatedQRCode);
    
    if (currentStatus !== 'Paid') {
      bookingSheet.getRange(rowIndex, CONFIG.BOOKING_COLS.PAYMENT_STATUS + 1)
        .setValue(calculation.needsContact ? 'Contact for Pricing' : 'Pending');
    }
    
    const calendarEventId = bookingSheet.getRange(rowIndex, headerMap['Calendar Event ID'] + 1).getValue();
    if (calendarEventId) {
      updateCalendarEvent(calendarEventId, formData.newDate, formData.newTimeSlots, formData.reference);
    }
    
    sendChangeConfirmationEmail({
      email: formData.email,
      name: formData.name,
      reference: formData.reference,
      oldDate: formatDate(oldDate),
      newDate: formData.newDate,
      oldTimeSlots: oldTimeSlots,
      newTimeSlots: formData.newTimeSlots,
      newAmount: calculation.needsContact ? 'Contact us' : newAmount,
      reason: formData.reason
    });
    
    return { 
      success: true, 
      message: 'Your date/time change request has been confirmed! A confirmation email has been sent.',
      newAmount: calculation.needsContact ? 'Contact us' : newAmount
    };
    
  } catch (error) {
    Logger.log('Error in processChangeRequest: ' + error);
    return { success: false, message: 'Error processing request: ' + error.message };
  }
}

function sendChangeConfirmationEmail(data) {
  const subject = `Booking Updated - ${data.reference} - Kings Equestrian`;
  
  const htmlBody = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <style>
    body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f4f4f4; margin: 0; padding: 0; color: #333; }
    .container { max-width: 650px; margin: 20px auto; background-color: #ffffff; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
    .header { background: linear-gradient(135deg, #1f4e3d 0%, #4f9c7a 100%); padding: 30px; text-align: center; color: #fff; }
    .header img { width: 70px; height: 70px; border-radius: 10%; margin-bottom: 10px; }
    .header h1 { margin: 0; font-size: 26px; }
    .content { padding: 35px 30px; }
    .change-box { background: #e7f3ff; border: 2px solid #0066cc; padding: 20px; border-radius: 10px; margin: 20px 0; }
    .change-item { display: flex; align-items: center; margin: 15px 0; font-size: 15px; }
    .old-value { text-decoration: line-through; color: #999; margin-right: 15px; }
    .new-value { color: #2c5f2d; font-weight: bold; }
    .arrow { margin: 0 10px; color: #0066cc; font-size: 18px; }
    .footer { background-color: #1f4e3d; color: #fff; padding: 25px; text-align: center; font-size: 13px; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <img src="https://kingsfarmequestrian.com/wp-content/uploads/2023/08/Logo2.jpg" alt="Kings Equestrian Logo">
      <h1>Kings Equestrian Foundation</h1>
      <p style="margin:8px 0 0;">Where horses don't just carry you – they change you</p>
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
          ${(data.newAmount !== undefined && data.newAmount !== null && data.newAmount !== '')
            ? `<p style="margin: 5px 0;"><strong>Updated Amount:</strong> ${typeof data.newAmount === 'number' ? `₹${data.newAmount.toLocaleString('en-IN')}` : data.newAmount}</p>`
            : ''}
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
      <p style="margin-top: 10px; font-size: 11px;">© ${new Date().getFullYear()} Kings Equestrian Foundation. All rights reserved.</p>
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

function createCalendarEvent(bookingData) {
  try {
    const calendar = CalendarApp.getDefaultCalendar();
    
    const date = new Date(bookingData.date);
    const timeSlots = bookingData.timeSlots.split(',');
    const firstSlot = timeSlots[0].trim();
    
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
    
    const endTime = new Date(startTime);
    endTime.setMinutes(endTime.getMinutes() + (timeSlots.length * 30));
    
    const participants = bookingData.participants || 1;
    const participantText = participants > 1 ? ` (${participants} participants)` : '';
    
    const event = calendar.createEvent(
      `Kings Equestrian - ${bookingData.name}${participantText} (${bookingData.reference})`,
      startTime,
      endTime,
      {
        description: `Service: ${bookingData.services}\nParticipants: ${participants}\nReference: ${bookingData.reference}\nPhone: ${bookingData.phone}\nEmail: ${bookingData.email}`,
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

function updateCalendarEvent(eventId, newDate, newTimeSlots, reference) {
  try {
    const calendar = CalendarApp.getDefaultCalendar();
    const event = calendar.getEventById(eventId);
    
    if (!event) {
      Logger.log('Event not found: ' + eventId);
      return false;
    }
    
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
    
    event.setTime(startTime, endTime);
    event.setDescription(event.getDescription() + `\n\nUpdated: ${new Date().toLocaleString()}`);
    
    Logger.log('Calendar event updated: ' + eventId);
    return true;
    
  } catch (error) {
    Logger.log('Error updating calendar event: ' + error);
    return false;
  }
}