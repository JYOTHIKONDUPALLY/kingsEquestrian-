// ============================================
// KINGS EQUESTRIAN - ENHANCED BOOKING SYSTEM
// ============================================

// --------------- CONFIG ---------------

const CONFIG = {
    UPI_ID: "vyapar.176548151976@hdfcbank",
    BUSINESS_NAME: "KingsEquestrian",
    PAYMENT_FORM_LINK: "https://forms.gle/WxskpjCcDQWkA7L57",
    EMAIL_TEMPLATE_DOC_ID: "1bUTpk9QCR4n1uUmMuoSRRcflTShG3jawuhemE28aTio",
    TERMS_CONDITIONS_DOC_ID: "1QbJHA5keyTLvgw-5stTY74i92BQ89TYya-NvtJ4YGx4",
    ADVANCE_BOOKING_AMOUNT: 1000, // Fixed advance booking fee
    webAppUrl: "https://script.google.com/macros/s/AKfycbxGNi137N_vvd6kFWe0CL2clALwKLp7QKsLgiWUd9fGcvYhTlaeQIy15n2vai_1g-PIig/exec",
    
    SHEETS: {
        BOOKING_FORM: "Booking Form Response",
        PAYMENT_FORM: "Payment Form Response",
        PRICING: "Pricing",
        MAIL_INFO: "Mail Info",
        PAYMENT_LOG: "Payment Logs"
    },
    
    BOOKING_COLS: {
        TIMESTAMP: 0,
        NAME: 1,
        EMAIL_ID: 2,
        PHONE_NUMBER: 3,
        OUR_SERVICES: 4,
        NUMBER_OF_PARTICIPANTS: 5,
        CONSENT: 6,
        REFERENCE: 8,
        WELCOME_EMAIL_SENT: 9,
        WELCOME_EMAIL_TIMESTAMP: 10,
    },
    
    PAYMENT_COLS: {
        TIMESTAMP: 0,
        REGISTRATION_NO: 1,
        AMOUNT_PAID: 2,
        SCREENSHOT: 3,
        PAYMENT_DATE: 4,
        TRANSACTION_REFERENCE_NUMBER: 5,
        PAN_AADHAAR: 6,
        PREFERRED_SERVICE_DATE: 7,
        PREFERRED_TIME_SLOT: 8,
        TRANSACTION_VERIFIED: 10,
        RECEIPT_SENT: 11,
        RECEIPT_SENT_TIMESTAMP: 12,
        PAYMENT_RECEIPT_NO: 13,
        PAYMENT_RECEIPT_DRIVER_LINK: 14
    },
    
    PAYMENT_LOG_COLS: {
        TIMESTAMP: 0,
        REFERENCE_NO: 1,
        AMOUNT: 2,
        QR_CODE: 3,
        UPI_LINK: 4,
        PAYMENT_STATUS: 5,
        MAIL_SENT: 6,
        MAIL_TIMESTAMPS: 7,
        RECEIPT_NO: 8
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

function getTermsAndConditionsPDF() {
    try {
        const doc = DocumentApp.openById(CONFIG.TERMS_CONDITIONS_DOC_ID);
        const blob = doc.getAs('application/pdf');
        blob.setName('Terms_and_Conditions.pdf');
        return blob;
    } catch (error) {
        Logger.log('Error creating T&C PDF: ' + error);
        return null;
    }
}

function getCCRecipients(mailType) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const mailInfoSheet = ss.getSheetByName(CONFIG.SHEETS.MAIL_INFO);
        if (!mailInfoSheet) {
            Logger.log('Mail Info sheet not found');
            return [];
        }
        const data = mailInfoSheet.getDataRange().getValues();
        const ccEmails = [];
        // Assuming columns: Email ID | Mail Type
        for (let i = 1; i < data.length; i++) {
            const email = data[i][0];
            const type = data[i][1];
            if (email && type && type.toLowerCase().includes(mailType.toLowerCase())) {
                ccEmails.push(email);
            }
        }
        return ccEmails;
    } catch (error) {
        Logger.log('Error getting CC recipients: ' + error);
        return [];
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
        const name = sheet.getRange(row, CONFIG.BOOKING_COLS.NAME + 1).getValue();
        const email = sheet.getRange(row, CONFIG.BOOKING_COLS.EMAIL_ID + 1).getValue();
        const phone = sheet.getRange(row, CONFIG.BOOKING_COLS.PHONE_NUMBER + 1).getValue();
        const services = sheet.getRange(row, CONFIG.BOOKING_COLS.OUR_SERVICES + 1).getValue();
        const participants = Number(sheet.getRange(row, CONFIG.BOOKING_COLS.NUMBER_OF_PARTICIPANTS + 1).getValue()) || 1;

        // Fixed advance booking amount
        const amount = CONFIG.ADVANCE_BOOKING_AMOUNT;

        // Generate payment details
        const reference = generateReference();
        const upiLink = createUPILink(amount, reference);
        const qrCode = createQRCode(upiLink);
        const paymentStatus = 'Pending';

        // Update BOOKING sheet with reference number only
        sheet.getRange(row, CONFIG.BOOKING_COLS.REFERENCE + 1).setValue(reference);

        // Get Payment Log sheet and create entry
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const paymentLogSheet = ss.getSheetByName(CONFIG.SHEETS.PAYMENT_LOG);
        if (!paymentLogSheet) {
            throw new Error('Payment Log sheet not found. Please check sheet name: "' + CONFIG.SHEETS.PAYMENT_LOG + '"');
        }

        // Append new row to Payment Log
        const paymentLogData = [
            new Date(), // TIMESTAMP
            reference, // REFERENCE_NO
            amount, // AMOUNT
            qrCode, // QR_CODE
            upiLink, // UPI_LINK
            paymentStatus, // PAYMENT_STATUS
            '', // MAIL_SENT (will update after email)
            '', // MAIL_TIMESTAMPS (will update after email)
            '' // RECEIPT_NO (will update after receipt sent)
        ];
        paymentLogSheet.appendRow(paymentLogData);
        const logRowIndex = paymentLogSheet.getLastRow();
        Logger.log(`Payment Log entry created at row ${logRowIndex} for reference: ${reference}`);

        // Get booking timestamp for consent PDF
        const bookingDate = sheet.getRange(row, CONFIG.BOOKING_COLS.TIMESTAMP + 1).getValue();

        // Send welcome email
        sendWelcomeEmail({
            name: name,
            email: email,
            phone: phone,
            services: services,
            participants: participants,
            amount: amount,
            reference: reference,
            upiLink: upiLink,
            qrCode: qrCode,
            row: row,
            sheet: sheet,
            logRow: logRowIndex,
            logSheet: paymentLogSheet,
            bookingDate: bookingDate
        });

        Logger.log(`Booking processed successfully for ${name} - Reference: ${reference}`);
    } catch (error) {
        Logger.log('Error in onBookingFormSubmit: ' + error);
        Logger.log('Stack trace: ' + error.stack);
        SpreadsheetApp.getUi().alert('Error processing booking: ' + error.message);
    }
}

// --------------- PAYMENT FORM SUBMIT HANDLER ---------------

function onPaymentFormSubmit(e) {
    try {
        const sheet = e.range.getSheet();
        const row = e.range.getRow();

        Logger.log(`Payment form submitted at row ${row}`);

        // Get registration number from payment form
        const referenceNumber = sheet.getRange(row, CONFIG.PAYMENT_COLS.REGISTRATION_NO + 1).getValue();
        const amount = Number(sheet.getRange(row, CONFIG.PAYMENT_COLS.AMOUNT_PAID + 1).getValue());
        const paymentDate = sheet.getRange(row, CONFIG.PAYMENT_COLS.PAYMENT_DATE + 1).getValue();
        const Timestamp = sheet.getRange(row, CONFIG.PAYMENT_COLS.TIMESTAMP + 1).getValue();
        if (!referenceNumber) {
            Logger.log('No registration number found in payment form submission');
            return;
        }

        Logger.log(`Processing payment for reference: ${referenceNumber}, amount: ${amount}`);

        // Check if this is a duplicate submission
        if (isDuplicateReceipt(referenceNumber, amount, paymentDate, Timestamp)) {
            Logger.log(`Duplicate receipt detected for ${referenceNumber} with amount ${amount} on ${paymentDate}. Skipping.`);
            
            // Mark the row to indicate duplicate
            sheet.getRange(row, CONFIG.PAYMENT_COLS.RECEIPT_SENT + 1)
                .setValue('Duplicate')
                .setBackground('#fff3cd')
                .setFontColor('#856404');
            
            return;
        }

        // Auto-verify transaction and send receipt
        // Mark as verified
        sheet.getRange(row, CONFIG.PAYMENT_COLS.TRANSACTION_VERIFIED + 1)
            .setValue('Yes')
            .setBackground('#d4edda')
            .setFontColor('#155724')
            .setFontWeight('bold');

        Logger.log('Transaction auto-verified, proceeding to send receipt');

        // Wait a moment for the verification to save
        Utilities.sleep(500);

        // Send receipt automatically
        sendReceiptForRow(row);

    } catch (error) {
        Logger.log('Error in onPaymentFormSubmit: ' + error);
        Logger.log('Stack trace: ' + error.stack);
    }
}

// --------------- DUPLICATE DETECTION ---------------

function isDuplicateReceipt(referenceNumber, amount, paymentDate, timestamps) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const paymentSheet = ss.getSheetByName(CONFIG.SHEETS.PAYMENT_FORM);
        
        if (!paymentSheet) {
            return false;
        }

        const data = paymentSheet.getDataRange().getValues();
        
        // Normalize the payment date for comparison
        const normalizedDate = normalizeDate(paymentDate);
        
        let matchCount = 0;
        
        // Start from row 2 (skip header)
        for (let i = 1; i < data.length; i++) {
            const rowRef = String(data[i][CONFIG.PAYMENT_COLS.REGISTRATION_NO] || '').trim();
            const rowAmount = Number(data[i][CONFIG.PAYMENT_COLS.AMOUNT_PAID]);
            const rowDate = data[i][CONFIG.PAYMENT_COLS.PAYMENT_DATE];
            const rowReceiptSent = String(data[i][CONFIG.PAYMENT_COLS.RECEIPT_SENT] || '').trim();
            const rowTimestamp = data[i][CONFIG.PAYMENT_COLS.TIMESTAMP];
            // Only count if registration number, amount, and date match
            // and receipt was already sent (not 'Duplicate' marker)
            if (rowRef === String(referenceNumber).trim() && 
                rowAmount === amount &&
                normalizeDate(rowDate) === normalizedDate &&
                rowReceiptSent.toLowerCase() === 'yes' && 
                rowTimestamp === timestamps) {
                matchCount++;
            }
        }
        
        // If we found at least one match with receipt already sent, this is a duplicate
        return matchCount > 0;
        
    } catch (error) {
        Logger.log('Error checking for duplicate: ' + error);
        return false;
    }
}

function normalizeDate(dateValue) {
    if (!dateValue) return '';
    
    try {
        const date = new Date(dateValue);
        if (isNaN(date.getTime())) return String(dateValue);
        
        // Return as YYYY-MM-DD for comparison
        return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    } catch (e) {
        return String(dateValue);
    }
}

// --------------- ENHANCED EMAIL FUNCTIONS ---------------

function sendWelcomeEmail(data) {
    const subject = `Welcome to Kings Equestrian - Booking ${data.reference}`;
    const participants = data.participants || 1;

    // Prepare attachments
    const attachments = [];

    // Add Terms & Conditions PDF
    const termsPDF = getTermsAndConditionsPDF();
    if (termsPDF) attachments.push(termsPDF);

    // Add Consent Form PDF
    try {
        const consentPDF = generateConsentPDF(data.name, data.email, data.phone, data.bookingDate);
        if (consentPDF) {
            attachments.push(consentPDF);
            Logger.log('Consent form PDF generated and added to attachments');
        }
    } catch (error) {
        Logger.log('Error generating consent PDF: ' + error);
    }

    // Add service-related PDFs
    const pricingData = getPricingData();
    const serviceList = Array.isArray(data.services) ? data.services.map(s => String(s).trim()).filter(Boolean) : String(data.services || '').split(',').map(s => s.trim()).filter(Boolean);

    serviceList.forEach(service => {
        const pricing = pricingData[service];
        if (pricing && pricing.docId) {
            const pdf = getServicePDF(pricing.docId, service);
            if (pdf) attachments.push(pdf);
        }
    });

    // Simple service list HTML (no pricing breakdown)
    const servicesHTML = serviceList.map(s => `<li>${s}</li>`).join('');

    const serviceDetailsHTML = `
        <div style="margin: 15px 0; padding: 15px; background: #f9f9f9; border-radius: 8px;">
      <h3 style="color: #2c5f2d; margin: 0 0 10px 0;">Selected Services</h3>
      <ul style="margin: 0; padding-left: 20px; font-size: 14px;">
                ${servicesHTML}
            </ul>
            <p style="color: #666; font-size: 14px; margin-bottom: 0;">See attached PDFs for detailed service information</p>
        </div>
    `;

    const paymentSection = `
         <div style="background: #e8f5e9; border: 2px solid #4caf50; padding: 20px; border-radius: 8px; margin: 20px 0;">
      <h3 style="margin-top: 0; color: #2e7d32;">💳 Reserve Your Slot</h3>
      <p style="font-size: 16px;">To confirm your booking, please pay the advance amount:</p>
      <p style="text-align: center; font-size: 32px; font-weight: bold; color: #2c5f2d; margin: 15px 0;">
        ₹${data.amount.toLocaleString('en-IN')}
      </p>
      <p style="text-align: center; font-size: 11px; color: #666; margin: 10px 0; font-style: italic;">
        This advance amount is non-refundable and can be used towards any Kings Equestrian service.
      </p>
      
      <div style=" margin: 25px 0;">
        <!-- QR Code -->
        <div style="flex: 1; min-width: 180px; background: white; padding: 20px; border-radius: 8px; text-align: center; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
          <p style="margin: 0 0 12px 0; font-weight: bold; font-size: 14px; color: #2c5f2d;">Scan to Pay</p>
          <img src="${data.qrCode}" alt="QR Code" style="width: 150px; height: 150px; border: 2px solid #e0e0e0; border-radius: 4px;">
        </div>
        
        <!-- Submit Payment Button -->
        <div style="flex: 1; min-width: 180px; text-align: center;">
          <p style="margin: 0 0 15px 0; font-size: 14px; color: #333;">After making payment:</p>
          <a href="${CONFIG.PAYMENT_FORM_LINK}" 
             style="display: inline-block; background: #2c5f2d; color: white; padding: 14px 28px; text-decoration: none; border-radius: 6px; font-weight: bold; font-size: 15px; box-shadow: 0 3px 6px rgba(44,95,45,0.3); transition: all 0.3s;">
            📝 Submit Payment & Select Slot
          </a>
          <p style="margin: 12px 0 0 0; font-size: 11px; color: #666; font-style: italic;">
            Don't forget to select your preferred date & time!
          </p>
        </div>
      </div>
      
      <div style="background: #fff3cd; padding: 15px; border-radius: 5px; margin-top: 15px; border-left: 4px solid #ffc107;">
        <p style="margin: 0; font-size: 13px; line-height: 1.6;">
          <strong>⚠️ Important:</strong> After scanning the QR code and making payment, click the button above to submit your payment screenshot, transaction details, and select your preferred date & time slot.
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
          
          <p>Thank you for choosing Kings Equestrian Foundation. Your booking request has been received.</p>
          
          <div style="background: #f0f8ff; border-left: 4px solid #2c5f2d; padding: 15px; margin: 20px 0;">
            <p style="margin: 0; font-size: 14px;">
              <strong>Booking Reference:</strong> <span style="font-size: 18px; color: #2c5f2d; font-weight: bold;">${data.reference}</span><br>
              <strong>Participants:</strong> ${participants}
            </p>
          </div>
          
          <h3 style="color: #2c5f2d; border-bottom: 2px solid #2c5f2d; padding-bottom: 10px;">📋 Service Details</h3>
          ${serviceDetailsHTML}
          
          ${paymentSection}
          
          <div style="background: #f9f9f9; padding: 20px; border-radius: 8px; margin-top: 20px;">
            <h4 style="margin: 0 0 10px 0; color: #2c5f2d;">📌 What's Next?</h4>
            <ul style="margin: 0; padding-left: 20px;">
              <li>Pay the advance booking fee of ₹${data.amount.toLocaleString('en-IN')}</li>
              <li>Submit payment confirmation and select your preferred date & time through the form</li>
              <li>Review the Terms & Conditions (attached)</li>
              <li>Wait for our confirmation email with your receipt</li>
              <li>Arrive 15 minutes before your scheduled time</li>
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
Services: ${data.services}
Participants: ${participants}

ADVANCE BOOKING AMOUNT: ₹${data.amount.toLocaleString('en-IN')}
(Non-refundable - Can be used towards any Kings Equestrian service)

PAYMENT INSTRUCTIONS:
1. Pay ₹${data.amount.toLocaleString('en-IN')} using UPI
2. Scan QR code or use UPI link
3. Submit payment details and select your preferred date & time: ${CONFIG.PAYMENT_FORM_LINK}

WHAT'S NEXT:
- Pay the advance booking fee
- Submit payment confirmation through the form
- Select your preferred date and time slot
- Review the Terms & Conditions (attached)
- Wait for our confirmation email with receipt
- Arrive 15 minutes before your scheduled time

Kings Equestrian Foundation
Karnataka, India
+91-9980895533 | info@kingsequestrian.com
  `;

    // Get CC recipients for welcome emails
    const ccEmails = getCCRecipients('Welcome Mail');

  MailApp.sendEmail({
    to: data.email,
    cc: ccEmails.join(','),
    subject: subject,
    body: plainBody,
    htmlBody: htmlBody,
    attachments: attachments,
    name: 'Kings Equestrian Foundation'
  });

    // Update BOOKING sheet email status
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

    // Update PAYMENT LOG sheet email status
    if (data.logSheet && data.logRow) {
        data.logSheet.getRange(data.logRow, CONFIG.PAYMENT_LOG_COLS.MAIL_SENT + 1)
            .setValue('Yes')
            .setBackground('#d4edda')
            .setFontColor('#155724')
            .setFontWeight('bold');

        data.logSheet.getRange(data.logRow, CONFIG.PAYMENT_LOG_COLS.MAIL_TIMESTAMPS + 1)
            .setValue(new Date())
            .setNumberFormat('dd-MMM-yyyy HH:mm:ss');
    }

    Logger.log(`Welcome email sent to: ${data.email} with CC to: ${ccEmails.join(', ')}`);
}

// --------------- RECEIPT GENERATION ---------------

function generateReceiptNumber(referenceNumber) {
    // Extract the serial number from reference (last 4 digits)
    const serialMatch = referenceNumber.match(/\d{4}$/);
    const serial = serialMatch ? serialMatch[0] : '0000';
    
    // Format: ReferenceNumber/SerialNumber
    return `${referenceNumber}/${serial}`;
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
    blob.setName(`80G_Receipt_${riderName.replace(/\s+/g, '_')}_${receiptNumber.replace(/\//g, '_')}.pdf`);

    htmlFile.setTrashed(true);

    return blob;
}

function createReceiptHTML(donorName, pan, amount, transactionRef, receiptNumber, logoBase64, stampBase64, signBase64) {
  const currentDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yy');
  const amountInWords = numberToWords(amount);
  
  // Determine payment mode based on transaction reference
  const isUPI = transactionRef && transactionRef !== 'N/A';
  
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
    .payment-mode { font-size: 10px; margin: 10px 0 5px 0; }
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
      
      <div class="payment-mode">
        <strong>Mode of Payment (✓ Tick above):</strong> Cheque / DD / NEFT / RTGS / UPI (Cash not eligible u/s 80G)<br><br>
        ${transactionRef && transactionRef !== 'N/A' ? `Transaction Reference No.: <strong>${transactionRef}</strong><br><br>` : ''}
        Amount in Words: <strong>${amountInWords}</strong>
      </div>
      <div class="checkbox-list" style="display:flex;flex-wrap:wrap; gap:10px;">
        <div class="checkbox-item"><span class="checkbox ${isUPI ? 'checked' : ''}"></span><span>UPI</span></div>
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

// --------------- SEND RECEIPT FOR SPECIFIC ROW ---------------

function sendReceiptForRow(rowIndex) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const paymentSheet = ss.getSheetByName(CONFIG.SHEETS.PAYMENT_FORM);
    const bookingSheet = ss.getSheetByName(CONFIG.SHEETS.BOOKING_FORM);
    const paymentLogSheet = ss.getSheetByName(CONFIG.SHEETS.PAYMENT_LOG);

    if (!paymentSheet || !bookingSheet || !paymentLogSheet) {
        Logger.log('Required sheets not found');
        return false;
    }

    const bookingValues = bookingSheet.getDataRange().getValues();
    const paymentLogValues = paymentLogSheet.getDataRange().getValues();

    let email = '';
    let referenceNumber = '';

    try {
        const row = paymentSheet.getRange(rowIndex, 1, 1, paymentSheet.getLastColumn()).getValues()[0];
        referenceNumber = row[CONFIG.PAYMENT_COLS.REGISTRATION_NO];

        if (!referenceNumber) {
            throw new Error('Reference number missing');
        }

        // Find booking match
        let bookingMatch = null;
        for (let j = 1; j < bookingValues.length; j++) {
            if (String(bookingValues[j][CONFIG.BOOKING_COLS.REFERENCE] || '').trim() === String(referenceNumber || '').trim()) {
                bookingMatch = {
                    rowIndex: j + 1,
                    row: bookingValues[j]
                };
                break;
            }
        }

        if (!bookingMatch) {
            throw new Error(`Booking not found for reference ${referenceNumber}`);
        }

        const riderName = bookingMatch.row[CONFIG.BOOKING_COLS.NAME];
        email = bookingMatch.row[CONFIG.BOOKING_COLS.EMAIL_ID];
        const phone = bookingMatch.row[CONFIG.BOOKING_COLS.PHONE_NUMBER];
        const services = bookingMatch.row[CONFIG.BOOKING_COLS.OUR_SERVICES];
        const participants = bookingMatch.row[CONFIG.BOOKING_COLS.NUMBER_OF_PARTICIPANTS] || 1;

        if (!email) {
            throw new Error('Email not found in booking');
        }

        const amount = Number(row[CONFIG.PAYMENT_COLS.AMOUNT_PAID]);
        if (!amount || Number.isNaN(amount)) {
            throw new Error('Valid amount is required');
        }

        const transactionId = row[CONFIG.PAYMENT_COLS.TRANSACTION_REFERENCE_NUMBER] || '';
        const pan = row[CONFIG.PAYMENT_COLS.PAN_AADHAAR] || '';
        const transactionVerified = row[CONFIG.PAYMENT_COLS.TRANSACTION_VERIFIED];
        const preferredDate = row[CONFIG.PAYMENT_COLS.PREFERRED_SERVICE_DATE];
        const preferredTimeSlots = row[CONFIG.PAYMENT_COLS.PREFERRED_TIME_SLOT];
        const paymentDate = row[CONFIG.PAYMENT_COLS.PAYMENT_DATE];
        const timeStamp = row[CONFIG.PAYMENT_COLS.TIMESTAMP];
        if (String(transactionVerified || '').toLowerCase() !== 'yes') {
            throw new Error('Transaction not verified. Please verify first.');
        }

        // Check for duplicate before proceeding
        if (isDuplicateReceipt(referenceNumber, amount, paymentDate, timeStamp)) {
            throw new Error('Duplicate receipt detected. Receipt already sent for this registration, amount, and date.');
        }

        // Find and update Payment Log entry
        let paymentLogRowIndex = -1;
        for (let k = 1; k < paymentLogValues.length; k++) {
            const logReference = String(paymentLogValues[k][CONFIG.PAYMENT_LOG_COLS.REFERENCE_NO] || '').trim();
            const logAmount = Number(paymentLogValues[k][CONFIG.PAYMENT_LOG_COLS.AMOUNT]);

            // Match by reference AND amount for accuracy
            if (logReference === String(referenceNumber).trim() && logAmount === amount) {
                paymentLogRowIndex = k + 1;
                break;
            }
        }

        if (paymentLogRowIndex === -1) {
            throw new Error(`Payment log entry not found for ${referenceNumber} with amount ${amount}`);
        }

        // Generate receipt with new format: RegistrationNumber/SerialNumber
        const receiptNumber = generateReceiptNumber(referenceNumber);
        const receiptPDF = generate80GReceipt(riderName, pan, amount, transactionId, receiptNumber);

        // Store receipt in Google Drive
        const driveInfo = storeReceiptInDrive(receiptPDF, riderName, receiptNumber, referenceNumber);
        if (driveInfo) {
            Logger.log(`Receipt stored in Drive: ${driveInfo.fileUrl}`);

            // Store Drive link in payment sheet
            paymentSheet.getRange(rowIndex, CONFIG.PAYMENT_COLS.PAYMENT_RECEIPT_DRIVER_LINK + 1).setValue(driveInfo.fileUrl);
        }

        // Update Payment Log sheet
        paymentLogSheet.getRange(paymentLogRowIndex, CONFIG.PAYMENT_LOG_COLS.PAYMENT_STATUS + 1)
            .setValue('Paid')
            .setBackground('#d4edda')
            .setFontColor('#155724')
            .setFontWeight('bold');

        paymentLogSheet.getRange(paymentLogRowIndex, CONFIG.PAYMENT_LOG_COLS.RECEIPT_NO + 1)
            .setValue(receiptNumber);

      // Send receipt email
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
    .success-banner { text-align: center; margin: 25px 0; }
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
        <p>Thank you for your payment. Your booking is confirmed.</p>
      </div>
      <div class="info-box">
        <strong>Your Payment Receipt (80G)</strong> is attached to this email for tax deduction purposes.
      </div>
      <p>
        <strong>Payment Details:</strong><br>
        Booking Reference: ${referenceNumber}<br>
        Receipt No: ${receiptNumber}<br>
        Amount Paid: ₹${amount.toLocaleString('en-IN')}<br>
        ${transactionId ? `Transaction ID: ${transactionId}<br>` : ''}
        ${preferredDate ? `Scheduled Date: ${formatDate(preferredDate)}<br>` : ''}
        ${preferredTimeSlots ? `Time Slot: ${preferredTimeSlots}<br>` : ''}
      </p>
      <p style="margin-top: 20px; font-size: 14px;">
        We look forward to welcoming you at Kings Equestrian. Please arrive 15 minutes before your scheduled time.
      </p>
      <p style="margin-top: 15px; font-size: 13px; color: #666;">
        <strong>What to bring:</strong><br>
        • Comfortable clothing<br>
        • Closed-toe shoes<br>
        • Your booking reference: ${referenceNumber}
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

        // Get CC recipients for receipt emails
        const ccEmails = getCCRecipients('Receipt Mail');

        MailApp.sendEmail({
            to: email,
            cc: ccEmails.join(','),
            subject: subject,
            htmlBody: htmlBody,
            attachments: [receiptPDF],
            name: 'Kings Equestrian Foundation'
        });

        // Update Payment Form Response sheet
        paymentSheet.getRange(rowIndex, CONFIG.PAYMENT_COLS.RECEIPT_SENT + 1)
            .setValue('Yes')
            .setBackground('#d4edda')
            .setFontColor('#155724')
            .setFontWeight('bold');

        paymentSheet.getRange(rowIndex, CONFIG.PAYMENT_COLS.RECEIPT_SENT_TIMESTAMP + 1)
            .setValue(new Date())
            .setNumberFormat('dd-MMM-yyyy HH:mm:ss');

        paymentSheet.getRange(rowIndex, CONFIG.PAYMENT_COLS.PAYMENT_RECEIPT_NO + 1)
            .setValue(receiptNumber);

        // Create calendar event with date/time from Payment Form
        if (preferredDate && preferredTimeSlots) {
            try {
                const calendarEventId = createCalendarEvent({
                    name: riderName,
                    email: email,
                    phone: phone,
                    services: services,
                    date: preferredDate,
                    timeSlots: preferredTimeSlots,
                    reference: referenceNumber,
                    participants: participants
                });

                if (calendarEventId) {
                    Logger.log(`Calendar event created: ${calendarEventId} for ${referenceNumber}`);
                }
            } catch (calError) {
                Logger.log(`Warning: Calendar event creation failed for ${referenceNumber}: ${calError.message}`);
                // Don't fail the entire receipt process if calendar fails
            }
        }

        Logger.log(`Receipt sent to: ${email} for ${referenceNumber} with CC to: ${ccEmails.join(', ')}`);
        return true;

    } catch (error) {
        Logger.log(`Receipt failed for row ${rowIndex} (Ref: ${referenceNumber || 'N/A'}, Email: ${email || 'N/A'}): ${error.message}`);
        throw error;
    }
}

// --------------- PAYMENT RECEIPT FUNCTION ---------------

function SendPaymentReceipt() {
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const paymentSheet = ss.getSheetByName(CONFIG.SHEETS.PAYMENT_FORM);

    if (!paymentSheet) {
        ui.alert('Payment Form Response sheet not found');
        return;
    }

    const selection = paymentSheet.getActiveRange();
    if (!selection || selection.getRow() === 1) {
        ui.alert('Please select valid rows to send receipts (not header row)');
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

    let successCount = 0;
    let failCount = 0;
    const errors = [];

    for (let i = 0; i < numRows; i++) {
        const rowIndex = startRow + i;

        try {
            sendReceiptForRow(rowIndex);
            successCount++;
            Utilities.sleep(1000); // Rate limiting
        } catch (error) {
            failCount++;
            errors.push(`Row ${rowIndex}: ${error.message}`);
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
    const paymentLogSheet = ss.getSheetByName(CONFIG.SHEETS.PAYMENT_LOG);

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
            const name = bookingSheet.getRange(rowIndex, CONFIG.BOOKING_COLS.NAME + 1).getValue();
            const email = bookingSheet.getRange(rowIndex, CONFIG.BOOKING_COLS.EMAIL_ID + 1).getValue();
            const phone = bookingSheet.getRange(rowIndex, CONFIG.BOOKING_COLS.PHONE_NUMBER + 1).getValue();
            const services = bookingSheet.getRange(rowIndex, CONFIG.BOOKING_COLS.OUR_SERVICES + 1).getValue();
            const reference = bookingSheet.getRange(rowIndex, CONFIG.BOOKING_COLS.REFERENCE + 1).getValue();
            const participants = Number(bookingSheet.getRange(rowIndex, CONFIG.BOOKING_COLS.NUMBER_OF_PARTICIPANTS + 1).getValue()) || 1;
            const bookingDate = bookingSheet.getRange(rowIndex, CONFIG.BOOKING_COLS.TIMESTAMP + 1).getValue();

            if (!email || !reference) {
                throw new Error('Missing email or reference');
            }

            const amount = CONFIG.ADVANCE_BOOKING_AMOUNT;
            const upiLink = createUPILink(amount, reference);
            const qrCode = createQRCode(upiLink);

            // Find payment log row for this reference
            let logRowIndex = null;
            if (paymentLogSheet) {
                const logValues = paymentLogSheet.getDataRange().getValues();
                for (let j = 1; j < logValues.length; j++) {
                    if (String(logValues[j][CONFIG.PAYMENT_LOG_COLS.REFERENCE_NO] || '').trim() === String(reference).trim()) {
                        logRowIndex = j + 1;
                        break;
                    }
                }
            }

            sendWelcomeEmail({
                name: name,
                email: email,
                phone: phone,
                services: services,
                participants: participants,
                amount: amount,
                reference: reference,
                upiLink: upiLink,
                qrCode: qrCode,
                row: rowIndex,
                sheet: bookingSheet,
                logRow: logRowIndex,
                logSheet: paymentLogSheet,
                bookingDate: bookingDate
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
    try {
        return Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), 'dd MMM yyyy');
    } catch (e) {
        return String(date);
    }
}

// --------------- GOOGLE CALENDAR INTEGRATION ---------------

function createCalendarEvent(bookingData) {
    try {
        const calendar = CalendarApp.getDefaultCalendar();

        const date = new Date(bookingData.date);
        const timeSlots = String(bookingData.timeSlots).split(',');
        const firstSlot = timeSlots[0].trim();

        const timeParts = firstSlot.match(/(\d+):(\d+)\s*(AM|PM)/i);
        if (!timeParts) {
            Logger.log('Invalid time format: ' + firstSlot);
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
            endTime, {
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

// --------------- CONSENT PDF GENERATION ---------------

function generateConsentPDF(name, email, phone, bookingDate) {

    const LABEL_FONT = 'Arial';
    const FONT_SIZE = 11;

    const doc = DocumentApp.create('Consent Form - ' + (name || 'Participant'));
    const body = doc.getBody();
    body.clear();

    body.setMarginTop(40);
    body.setMarginBottom(40);
    body.setMarginLeft(50);
    body.setMarginRight(50);

    function paragraph(textStr, size = FONT_SIZE, bold = false, spacing = 6, align = null) {
        const p = body.appendParagraph(textStr);
        const t = p.editAsText();
        t.setFontFamily(LABEL_FONT).setFontSize(size).setBold(bold);
        if (align) p.setAlignment(align);
        p.setSpacingAfter(spacing);
        return p;
    }

    function formatValue(textObj, fullText, value) {
        if (!value || value.toString().trim().length === 0) return;
        if (!fullText) return;

        const valStr = value.toString();
        const start = fullText.indexOf(valStr);

        if (start === -1) return;

        const end = start + valStr.length - 1;

        if (end >= start && start >= 0 && end < fullText.length) {
            textObj.setBold(start, end, true);
            textObj.setUnderline(start, end, true);
        }
    }

    function formatDateOnly(dateValue) {
        if (!dateValue) return null;

        if (typeof dateValue === 'string') {
            // Try to parse if it's a string
            try {
                const parsed = new Date(dateValue);
                if (!isNaN(parsed.getTime())) {
                    dateValue = parsed;
                } else {
                    return dateValue; // Return as-is if can't parse
                }
            } catch (e) {
                return dateValue;
            }
        }

        if (dateValue instanceof Date) {
            const day = String(dateValue.getDate()).padStart(2, '0');
            const month = String(dateValue.getMonth() + 1).padStart(2, '0');
            const year = dateValue.getFullYear();
            return `${day}/${month}/${year}`;
        }

        return dateValue.toString();
    }

    /* ========= HEADER WITH LOGO ========= */

    const logoUrl = 'https://kingsfarmequestrian.com/wp-content/uploads/2023/08/Logo2.jpg';
    let logoBlob;

    try {
        logoBlob = UrlFetchApp.fetch(logoUrl).getBlob();
    } catch (e) {
        Logger.log('Error fetching logo: ' + e);
    }

    if (logoBlob) {
        const logoPara = body.appendParagraph('');
        logoPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        const logoImg = logoPara.appendInlineImage(logoBlob);
        logoImg.setWidth(120);
        logoImg.setHeight(120);
        logoPara.setSpacingAfter(20);
    }

    /* ========= TITLE ========= */

    paragraph('KINGS EQUESTRIAN FOUNDATION', 16, true, 5, DocumentApp.HorizontalAlignment.CENTER);
    paragraph('Acknowledgement & Consent Form – Horse Riding Participants', 13, true, 3, DocumentApp.HorizontalAlignment.CENTER);
    paragraph('(Applicable for Individual / Group / Family Participants)', 10, false, 25, DocumentApp.HorizontalAlignment.CENTER);

    /* ========= CONTENT ========= */

    paragraph(
        'Kings Equestrian Foundation offers horse riding programs and related activities, which may include casual riding, dressage, jumping, workshops, clinics, and equine interaction.',
        11, false, 12
    );

    paragraph(
        'I/we understand and acknowledge that participation in equestrian activities involves inherent risks, including but not limited to falls, bruises, muscle strain, fractures, head injuries, or other serious injuries. I/we further acknowledge that horses are live animals and their behaviour can be unpredictable.',
        11, false, 12
    );

    paragraph(
        'I/we also acknowledge that Kings Equestrian Foundation follows reasonable safety precautions, provides trained supervision, and enforces established safety guidelines. However, despite all precautions, accidents may occasionally occur.',
        11, false, 20
    );

    /* ========= SEPARATOR ========= */

    let sepPara = body.appendParagraph('⸻');
    sepPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    sepPara.setSpacingAfter(20);

    /* ========= MEDICAL FITNESS & INSURANCE ========= */

    paragraph('Medical Fitness & Insurance Declaration', 12, true, 12);

    paragraph(
        'I/we hereby declare that I / my child / all participants covered under this consent are medically fit to participate in horse riding and equestrian-related activities. To the best of my/our knowledge, there are no undisclosed medical conditions, injuries, or health concerns that would prevent safe participation, except those disclosed in writing to Kings Equestrian Foundation prior to participation.',
        11, false, 12
    );

    paragraph(
        'I/we further confirm that I / my child / all participants are covered by valid medical and/or personal accident insurance, which will cover any injuries, medical treatment, or emergencies arising from participation.',
        11, false, 12
    );

    paragraph(
        'I/we understand and agree that Kings Equestrian Foundation is not responsible for medical expenses, and all such costs shall be borne by the participant(s) or covered under their insurance.',
        11, false, 20
    );

    /* ========= SEPARATOR ========= */

    sepPara = body.appendParagraph('⸻');
    sepPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    sepPara.setSpacingAfter(20);

    /* ========= ACKNOWLEDGEMENT & AGREEMENT ========= */

    paragraph('Acknowledgement & Agreement', 12, true, 12);

    paragraph('I/we confirm that:', 11, false, 8);

    const bulletPoints = [
        'I/we have carefully read and fully understood this consent form.',
        'I/we understand the nature of equestrian activities and the associated risks.',
        'I/we voluntarily consent to participation.',
        'For participants under 18 years of age, I/we am/are the parent(s) or legal guardian(s) and authorised to provide consent.',
        'All participants agree to follow safety instructions, rules, and guidelines issued by Kings Equestrian Foundation and its instructors at all times.'
    ];

    bulletPoints.forEach(point => {
        const p = body.appendParagraph('• ' + point);
        p.editAsText().setFontFamily(LABEL_FONT).setFontSize(11);
        p.setSpacingAfter(6);
        p.setIndentStart(20);
        p.setIndentFirstLine(0);
    });

    body.appendParagraph('').setSpacingAfter(8);

    paragraph(
        'I/we agree that Kings Equestrian Foundation, its trainers, staff, and associates shall not be held responsible for injuries arising from participation, except in cases of proven negligence.',
        11, false, 20
    );

    /* ========= SEPARATOR ========= */

    sepPara = body.appendParagraph('⸻');
    sepPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    sepPara.setSpacingAfter(20);

    /* ========= PRIMARY CONTACT DETAILS ========= */

    paragraph('Primary Contact / Parent / Guardian Details', 12, true, 12);

    let p = body.appendParagraph('');
    let t = p.editAsText();
    const nameSpaced = name ? `  ${name}  ` : '___________________________________';
    const nameLine = `Name: ${nameSpaced}`;
    t.setText(nameLine).setFontFamily(LABEL_FONT).setFontSize(FONT_SIZE);
    if (name) formatValue(t, nameLine, nameSpaced);
    p.setSpacingAfter(12);

    p = body.appendParagraph('');
    t = p.editAsText();
    const phoneSpaced = phone ? `  ${phone}  ` : '___________________________________';
    const phoneLine = `Contact Number: ${phoneSpaced}`;
    t.setText(phoneLine).setFontFamily(LABEL_FONT).setFontSize(FONT_SIZE);
    if (phone) formatValue(t, phoneLine, phoneSpaced);
    p.setSpacingAfter(12);

    p = body.appendParagraph('');
    t = p.editAsText();
    const emailSpaced = email ? `  ${email}  ` : '___________________________________';
    const emailLine = `Email ID: ${emailSpaced}`;
    t.setText(emailLine).setFontFamily(LABEL_FONT).setFontSize(FONT_SIZE);
    if (email) formatValue(t, emailLine, emailSpaced);
    p.setSpacingAfter(25);

    /* ========= SEPARATOR ========= */

    sepPara = body.appendParagraph('⸻');
    sepPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    sepPara.setSpacingAfter(25);

    /* ========= SIGNATURE SECTION ========= */

    p = body.appendParagraph('');
    t = p.editAsText();
    const signatureSpaced = name ? `  ${name}  ` : '___________________________________';
    const dateFormatted = formatDateOnly(bookingDate);
    const dateSpaced = dateFormatted ? `  ${dateFormatted}  ` : '_______________';
    const signatureLine = `Signature of Participant / Parent / Guardian: ${signatureSpaced}     Date: ${dateSpaced}`;
    t.setText(signatureLine).setFontFamily(LABEL_FONT).setFontSize(11);

    // Apply Dancing Script font to signature (name)
    if (name) {
        const sigStart = signatureLine.indexOf(signatureSpaced);
        if (sigStart !== -1) {
            const sigEnd = sigStart + signatureSpaced.length - 1;
            if (sigEnd >= sigStart && sigStart >= 0) {
                t.setFontFamily(sigStart, sigEnd, 'Dancing Script');
                t.setFontSize(sigStart, sigEnd, 16);
                t.setBold(sigStart, sigEnd, false);
                t.setUnderline(sigStart, sigEnd, false);
            }
        }
    }

    // Format the date
    if (dateFormatted) formatValue(t, signatureLine, dateSpaced);
    p.setSpacingAfter(30);

    /* ========= FOOTER ========= */

    const footerPara = paragraph(
        'Kings Equestrian Foundation | Karnataka, India | +91-9980895533 | info@kingsequestrian.com',
        9, false, 0, DocumentApp.HorizontalAlignment.CENTER
    );
    footerPara.editAsText().setForegroundColor('#666666');

    /* ========= SAVE ========= */

    doc.saveAndClose();

    const pdf = doc.getAs('application/pdf');
    pdf.setName(`Consent_Form_${(name || 'Participant').replace(/\s+/g, '_')}.pdf`);

    DriveApp.getFileById(doc.getId()).setTrashed(true);

    return pdf;
}

// --------------- DRIVE STORAGE ---------------

function getKingsFarmFolder() {
    const folderName = "Kings Farm Receipts";
    const year = new Date().getFullYear();

    let mainFolder = DriveApp.getFoldersByName(folderName);
    if (!mainFolder.hasNext()) {
        mainFolder = DriveApp.createFolder(folderName);
    } else {
        mainFolder = mainFolder.next();
    }

    // Create year subfolder
    const yearFolders = mainFolder.getFoldersByName(year.toString());
    if (yearFolders.hasNext()) {
        return yearFolders.next();
    } else {
        return mainFolder.createFolder(year.toString());
    }
}

function storeReceiptInDrive(receiptBlob, riderName, receiptNumber, referenceNumber) {
    try {
        const folder = getKingsFarmFolder();

        const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
        const fileName = `Receipt_${receiptNumber.replace(/\//g, '-')}_${riderName.replace(/\s+/g, '_')}_${timestamp}.pdf`;

        const file = folder.createFile(receiptBlob);
        file.setName(fileName);
        file.setDescription(`Receipt for ${riderName} | Reference: ${referenceNumber} | Receipt No: ${receiptNumber}`);

        Logger.log(`Receipt saved to Drive: ${fileName}`);

        return {
            fileId: file.getId(),
            fileUrl: file.getUrl(),
            fileName: fileName,
            folderId: folder.getId(),
            folderUrl: folder.getUrl()
        };
    } catch (error) {
        Logger.log('Error storing receipt in Drive: ' + error);
        return null;
    }
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
    // Delete existing triggers
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Create trigger for booking form submissions
    ScriptApp.newTrigger('onBookingFormSubmit')
        .forSpreadsheet(ss)
        .onFormSubmit()
        .create();

    // Create trigger for payment form submissions
    ScriptApp.newTrigger('onPaymentFormSubmit')
        .forSpreadsheet(ss)
        .onFormSubmit()
        .create();

    SpreadsheetApp.getUi().alert('✅ Triggers set up successfully!\n\n' +
        'The system will now automatically:\n' +
        '- Generate reference numbers and send welcome emails on booking\n' +
        '- Auto-verify and send receipts when payment form is submitted\n' +
        '- Prevent duplicate receipts\n' +
        '- Store receipts in Google Drive');
}