// ============================================
// KINGS EQUESTRIAN - BOOKING LOGIC
// File: 2_Booking.js
// Depends on: 1_Config.js
// ============================================

// --------------- BOOKING FORM SUBMIT HANDLER ---------------

function onBookingFormSubmit(e) {
    try {
        const sheet = e.range.getSheet();
        if (sheet.getName() !== CONFIG.SHEETS.BOOKING_FORM) {
            Logger.log('onBookingFormSubmit: Skipping — wrong sheet: ' + sheet.getName());
            return;
        }
        const row = e.range.getRow();

        const name        = sheet.getRange(row, CONFIG.BOOKING_COLS.NAME + 1).getValue();
        const email       = sheet.getRange(row, CONFIG.BOOKING_COLS.EMAIL_ID + 1).getValue();
        const phone       = sheet.getRange(row, CONFIG.BOOKING_COLS.PHONE_NUMBER + 1).getValue();
        const services    = sheet.getRange(row, CONFIG.BOOKING_COLS.OUR_SERVICES + 1).getValue();
        const participants = Number(sheet.getRange(row, CONFIG.BOOKING_COLS.NUMBER_OF_PARTICIPANTS + 1).getValue()) || 1;
        const bookingDate = sheet.getRange(row, CONFIG.BOOKING_COLS.TIMESTAMP + 1).getValue();

        const amount    = CONFIG.ADVANCE_BOOKING_AMOUNT;
        const reference = generateReference();
        const upiLink   = createUPILink(amount, reference);
        const qrCode    = createQRCode(upiLink);

        // Write reference back to the booking sheet
        sheet.getRange(row, CONFIG.BOOKING_COLS.REFERENCE + 1).setValue(reference);

        sendWelcomeEmail({
            name, email, phone, services, participants,
            amount, reference, upiLink, qrCode,
            row, sheet, bookingDate
        });

        Logger.log(`Booking processed for ${name} — Reference: ${reference}`);
    } catch (error) {
        Logger.log('Error in onBookingFormSubmit: ' + error);
        Logger.log('Stack: ' + error.stack);
        SpreadsheetApp.getUi().alert('Error processing booking: ' + error.message);
    }
}

// --------------- WELCOME EMAIL ---------------

function sendWelcomeEmail(data) {
    const subject      = `Welcome to Kings Equestrian - Booking ${data.reference}`;
    const participants = data.participants || 1;
    const attachments  = [];

    // Attach T&C
    const termsPDF = getTermsAndConditionsPDF();
    if (termsPDF) attachments.push(termsPDF);

    // Attach consent form
    try {
        const consentPDF = generateConsentPDF(data.name, data.email, data.phone, data.bookingDate);
        if (consentPDF) {
            attachments.push(consentPDF);
            Logger.log('Consent form PDF generated and added');
        }
    } catch (error) {
        Logger.log('Error generating consent PDF: ' + error);
    }

    // Build service list and attach service PDFs
    const pricingData  = getPricingData();
    const rawServices  = Array.isArray(data.services)
        ? data.services.join(', ')
        : String(data.services || '');

    const serviceList = Object.keys(pricingData).filter(key =>
        rawServices.toLowerCase().includes(key.toLowerCase())
    );

    serviceList.forEach(key => {
        const pricing = pricingData[key];
        if (pricing && pricing.docId) {
            const pdf = getServicePDF(pricing.docId, key);
            if (pdf) attachments.push(pdf);
        }
    });

    const servicesHTML      = serviceList.map(s => `<li>${s}</li>`).join('');
    const serviceDetailsHTML = `
        <div style="margin:15px 0;padding:15px;background:#f9f9f9;border-radius:8px;">
            <h3 style="color:#2c5f2d;margin:0 0 10px 0;">Selected Services</h3>
            <ul style="margin:0;padding-left:20px;font-size:14px;">${servicesHTML}</ul>
            <p style="color:#666;font-size:14px;margin-bottom:0;">See attached PDFs for detailed service information</p>
        </div>`;

    const paymentSection = `
        <div style="background:#e8f5e9;border:2px solid #4caf50;padding:20px;border-radius:8px;margin:20px 0;">
            <h3 style="margin-top:0;color:#2e7d32;">💳 Reserve Your Slot</h3>
            <p style="font-size:16px;">To confirm your booking, please pay the advance amount:</p>
            <p style="text-align:center;font-size:32px;font-weight:bold;color:#2c5f2d;margin:15px 0;">
                ₹${data.amount.toLocaleString('en-IN')}
            </p>
            <p style="text-align:center;font-size:11px;color:#666;margin:10px 0;font-style:italic;">
                This advance amount is non-refundable and can be used towards any Kings Equestrian service.
            </p>
            <div style="margin:25px 0;">
                <div style="flex:1;min-width:180px;background:white;padding:20px;border-radius:8px;text-align:center;box-shadow:0 2px 4px rgba(0,0,0,0.1);">
                    <p style="margin:0 0 12px 0;font-weight:bold;font-size:14px;color:#2c5f2d;">Scan to Pay</p>
                    <img src="${data.qrCode}" alt="QR Code" style="width:150px;height:150px;border:2px solid #e0e0e0;border-radius:4px;">
                </div>
                <div style="flex:1;min-width:180px;text-align:center;">
                    <p style="margin:0 0 15px 0;font-size:14px;color:#333;">After making payment:</p>
                    <a href="${CONFIG.PAYMENT_FORM_LINK}"
                       style="display:inline-block;background:#2c5f2d;color:white;padding:14px 28px;text-decoration:none;border-radius:6px;font-weight:bold;font-size:15px;">
                        📝 Submit Payment & Select Slot
                    </a>
                    <p style="margin:12px 0 0 0;font-size:11px;color:#666;font-style:italic;">
                        Don't forget to select your preferred date &amp; time!
                    </p>
                </div>
            </div>
            <div style="background:#fff3cd;padding:15px;border-radius:5px;margin-top:15px;border-left:4px solid #ffc107;">
                <p style="margin:0;font-size:13px;line-height:1.6;">
                    <strong>⚠️ Important:</strong> After scanning the QR code and making payment, click the button above
                    to submit your payment screenshot, transaction details, and select your preferred date &amp; time slot.
                </p>
            </div>
        </div>`;

    const htmlBody = `
<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0"></head>
<body style="font-family:Arial,sans-serif;color:#333;line-height:1.6;margin:0;padding:0;background:#f5f5f5;">
<div style="max-width:650px;margin:20px auto;background:white;border-radius:10px;overflow:hidden;box-shadow:0 2px 10px rgba(0,0,0,0.1);">
    <div style="background:linear-gradient(135deg,#1f4e3d 0%,#4f9c7a 100%);padding:30px;text-align:center;color:white;">
        <img src="https://kingsfarmequestrian.com/wp-content/uploads/2023/08/Logo2.jpg"
             alt="Kings Equestrian" style="width:80px;height:80px;border-radius:50%;margin-bottom:15px;">
        <h1 style="margin:0;font-size:28px;">Welcome to Kings Equestrian!</h1>
        <p style="margin:10px 0 0 0;font-size:14px;opacity:0.9;">Where horses don't just carry you — they change you</p>
    </div>
    <div style="padding:30px;">
        <h2 style="color:#2c5f2d;margin-top:0;">Hello ${data.name}! 👋</h2>
        <p>Thank you for choosing Kings Equestrian Foundation. Your booking request has been received.</p>
        <div style="background:#f0f8ff;border-left:4px solid #2c5f2d;padding:15px;margin:20px 0;">
            <p style="margin:0;font-size:14px;">
                <strong>Booking Reference:</strong>
                <span style="font-size:18px;color:#2c5f2d;font-weight:bold;">${data.reference}</span><br>
                <strong>Participants:</strong> ${participants}
            </p>
        </div>
        <h3 style="color:#2c5f2d;border-bottom:2px solid #2c5f2d;padding-bottom:10px;">📋 Service Details</h3>
        ${serviceDetailsHTML}
        ${paymentSection}
        <div style="background:#f9f9f9;padding:20px;border-radius:8px;margin-top:20px;">
            <h4 style="margin:0 0 10px 0;color:#2c5f2d;">📌 What's Next?</h4>
            <ul style="margin:0;padding-left:20px;">
                <li>Pay the advance booking fee of ₹${data.amount.toLocaleString('en-IN')}</li>
                <li>Submit payment confirmation and select your preferred date &amp; time through the form</li>
                <li>Review the Terms &amp; Conditions (attached)</li>
                <li>Wait for our confirmation email with your receipt</li>
                <li>Arrive 15 minutes before your scheduled time</li>
            </ul>
        </div>
        <p style="margin-top:20px;font-size:14px;color:#666;">
            If you have any questions, feel free to reach out to us anytime.
        </p>
    </div>
    <div style="background:#1f4e3d;color:white;padding:20px;text-align:center;font-size:13px;">
        <p style="margin:0 0 10px 0;"><strong>Kings Equestrian Foundation</strong></p>
        <p style="margin:0;">📍 Karnataka, India</p>
        <p style="margin:5px 0;">📞 +91-9980895533 | ✉️ info@kingsequestrian.com</p>
        <p style="margin:10px 0 0 0;opacity:0.8;font-size:11px;">
            © ${new Date().getFullYear()} Kings Equestrian Foundation. All rights reserved.
        </p>
    </div>
</div>
</body>
</html>`;

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
+91-9980895533 | info@kingsequestrian.com`;

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

    Logger.log(`Welcome email sent to: ${data.email} — CC: ${ccEmails.join(', ')}`);
}

// --------------- RESEND WELCOME EMAIL MENU FUNCTION ---------------

function ResendWelcomeEmail() {
    const ui           = SpreadsheetApp.getUi();
    const ss           = SpreadsheetApp.getActiveSpreadsheet();
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
    const numRows  = selection.getNumRows();

    if (startRow === 1) {
        ui.alert('Cannot send emails for header row');
        return;
    }

    const response = ui.alert('Resend Welcome Emails', `Resend welcome emails for ${numRows} row(s)?`, ui.ButtonSet.YES_NO);
    if (response !== ui.Button.YES) return;

    let successCount = 0;
    let failCount    = 0;

    for (let i = 0; i < numRows; i++) {
        const rowIndex = startRow + i;
        try {
            const name         = bookingSheet.getRange(rowIndex, CONFIG.BOOKING_COLS.NAME + 1).getValue();
            const email        = bookingSheet.getRange(rowIndex, CONFIG.BOOKING_COLS.EMAIL_ID + 1).getValue();
            const phone        = bookingSheet.getRange(rowIndex, CONFIG.BOOKING_COLS.PHONE_NUMBER + 1).getValue();
            const services     = bookingSheet.getRange(rowIndex, CONFIG.BOOKING_COLS.OUR_SERVICES + 1).getValue();
            const reference    = bookingSheet.getRange(rowIndex, CONFIG.BOOKING_COLS.REFERENCE + 1).getValue();
            const participants = Number(bookingSheet.getRange(rowIndex, CONFIG.BOOKING_COLS.NUMBER_OF_PARTICIPANTS + 1).getValue()) || 1;
            const bookingDate  = bookingSheet.getRange(rowIndex, CONFIG.BOOKING_COLS.TIMESTAMP + 1).getValue();

            if (!email || !reference) throw new Error('Missing email or reference');

            const amount  = CONFIG.ADVANCE_BOOKING_AMOUNT;
            const upiLink = createUPILink(amount, reference);
            const qrCode  = createQRCode(upiLink);

            sendWelcomeEmail({
                name, email, phone, services, participants,
                amount, reference, upiLink, qrCode,
                row: rowIndex, sheet: bookingSheet, bookingDate
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

// --------------- CONSENT PDF GENERATION ---------------

function generateConsentPDF(name, email, phone, bookingDate) {
    const LABEL_FONT = 'Arial';
    const FONT_SIZE  = 11;

    const doc  = DocumentApp.create('Consent Form - ' + (name || 'Participant'));
    const body = doc.getBody();
    body.clear();
    body.setMarginTop(40).setMarginBottom(40).setMarginLeft(50).setMarginRight(50);

    function paragraph(textStr, size = FONT_SIZE, bold = false, spacing = 6, align = null) {
        const p = body.appendParagraph(textStr);
        p.editAsText().setFontFamily(LABEL_FONT).setFontSize(size).setBold(bold);
        if (align) p.setAlignment(align);
        p.setSpacingAfter(spacing);
        return p;
    }

    function formatValue(textObj, fullText, value) {
        if (!value || !fullText) return;
        const valStr = value.toString();
        const start  = fullText.indexOf(valStr);
        if (start === -1) return;
        const end = start + valStr.length - 1;
        if (end >= start && start >= 0 && end < fullText.length) {
            textObj.setBold(start, end, true).setUnderline(start, end, true);
        }
    }

    function formatDateOnly(dateValue) {
        if (!dateValue) return null;
        if (typeof dateValue === 'string') {
            try {
                const parsed = new Date(dateValue);
                if (!isNaN(parsed.getTime())) dateValue = parsed;
                else return dateValue;
            } catch (e) { return dateValue; }
        }
        if (dateValue instanceof Date) {
            return `${String(dateValue.getDate()).padStart(2,'0')}/${String(dateValue.getMonth()+1).padStart(2,'0')}/${dateValue.getFullYear()}`;
        }
        return dateValue.toString();
    }

    // Logo
    try {
        const logoBlob = UrlFetchApp.fetch('https://kingsfarmequestrian.com/wp-content/uploads/2023/08/Logo2.jpg').getBlob();
        const logoPara = body.appendParagraph('');
        logoPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        const logoImg = logoPara.appendInlineImage(logoBlob);
        logoImg.setWidth(120).setHeight(120);
        logoPara.setSpacingAfter(20);
    } catch (e) {
        Logger.log('Error fetching logo: ' + e);
    }

    paragraph('KINGS EQUESTRIAN FOUNDATION', 16, true, 5, DocumentApp.HorizontalAlignment.CENTER);
    paragraph('Acknowledgement & Consent Form – Horse Riding Participants', 13, true, 3, DocumentApp.HorizontalAlignment.CENTER);
    paragraph('(Applicable for Individual / Group / Family Participants)', 10, false, 25, DocumentApp.HorizontalAlignment.CENTER);

    paragraph('Kings Equestrian Foundation offers horse riding programs and related activities, which may include casual riding, dressage, jumping, workshops, clinics, and equine interaction.', 11, false, 12);
    paragraph('I/we understand and acknowledge that participation in equestrian activities involves inherent risks, including but not limited to falls, bruises, muscle strain, fractures, head injuries, or other serious injuries. I/we further acknowledge that horses are live animals and their behaviour can be unpredictable.', 11, false, 12);
    paragraph('I/we also acknowledge that Kings Equestrian Foundation follows reasonable safety precautions, provides trained supervision, and enforces established safety guidelines. However, despite all precautions, accidents may occasionally occur.', 11, false, 20);

    let sep = body.appendParagraph('⸻');
    sep.setAlignment(DocumentApp.HorizontalAlignment.CENTER).setSpacingAfter(20);

    paragraph('Medical Fitness & Insurance Declaration', 12, true, 12);
    paragraph('I/we hereby declare that I / my child / all participants covered under this consent are medically fit to participate in horse riding and equestrian-related activities. To the best of my/our knowledge, there are no undisclosed medical conditions, injuries, or health concerns that would prevent safe participation, except those disclosed in writing to Kings Equestrian Foundation prior to participation.', 11, false, 12);
    paragraph('I/we further confirm that I / my child / all participants are covered by valid medical and/or personal accident insurance, which will cover any injuries, medical treatment, or emergencies arising from participation.', 11, false, 12);
    paragraph('I/we understand and agree that Kings Equestrian Foundation is not responsible for medical expenses, and all such costs shall be borne by the participant(s) or covered under their insurance.', 11, false, 20);

    sep = body.appendParagraph('⸻');
    sep.setAlignment(DocumentApp.HorizontalAlignment.CENTER).setSpacingAfter(20);

    paragraph('Acknowledgement & Agreement', 12, true, 12);
    paragraph('I/we confirm that:', 11, false, 8);

    [
        'I/we have carefully read and fully understood this consent form.',
        'I/we understand the nature of equestrian activities and the associated risks.',
        'I/we voluntarily consent to participation.',
        'For participants under 18 years of age, I/we am/are the parent(s) or legal guardian(s) and authorised to provide consent.',
        'All participants agree to follow safety instructions, rules, and guidelines issued by Kings Equestrian Foundation and its instructors at all times.'
    ].forEach(point => {
        const p = body.appendParagraph('• ' + point);
        p.editAsText().setFontFamily(LABEL_FONT).setFontSize(11);
        p.setSpacingAfter(6).setIndentStart(20).setIndentFirstLine(0);
    });

    body.appendParagraph('').setSpacingAfter(8);
    paragraph('I/we agree that Kings Equestrian Foundation, its trainers, staff, and associates shall not be held responsible for injuries arising from participation, except in cases of proven negligence.', 11, false, 20);

    sep = body.appendParagraph('⸻');
    sep.setAlignment(DocumentApp.HorizontalAlignment.CENTER).setSpacingAfter(20);

    paragraph('Primary Contact / Parent / Guardian Details', 12, true, 12);

    // Name
    let p = body.appendParagraph('');
    let t = p.editAsText();
    const nameSpaced = name ? `  ${name}  ` : '___________________________________';
    const nameLine   = `Name: ${nameSpaced}`;
    t.setText(nameLine).setFontFamily(LABEL_FONT).setFontSize(FONT_SIZE);
    if (name) formatValue(t, nameLine, nameSpaced);
    p.setSpacingAfter(12);

    // Phone
    p = body.appendParagraph('');
    t = p.editAsText();
    const phoneSpaced = phone ? `  ${phone}  ` : '___________________________________';
    const phoneLine   = `Contact Number: ${phoneSpaced}`;
    t.setText(phoneLine).setFontFamily(LABEL_FONT).setFontSize(FONT_SIZE);
    if (phone) formatValue(t, phoneLine, phoneSpaced);
    p.setSpacingAfter(12);

    // Email
    p = body.appendParagraph('');
    t = p.editAsText();
    const emailSpaced = email ? `  ${email}  ` : '___________________________________';
    const emailLine   = `Email ID: ${emailSpaced}`;
    t.setText(emailLine).setFontFamily(LABEL_FONT).setFontSize(FONT_SIZE);
    if (email) formatValue(t, emailLine, emailSpaced);
    p.setSpacingAfter(25);

    sep = body.appendParagraph('⸻');
    sep.setAlignment(DocumentApp.HorizontalAlignment.CENTER).setSpacingAfter(25);

    // Signature line
    p = body.appendParagraph('');
    t = p.editAsText();
    const signatureSpaced = name ? `  ${name}  ` : '___________________________________';
    const dateFormatted   = formatDateOnly(bookingDate);
    const dateSpaced      = dateFormatted ? `  ${dateFormatted}  ` : '_______________';
    const signatureLine   = `Signature of Participant / Parent / Guardian: ${signatureSpaced}     Date: ${dateSpaced}`;
    t.setText(signatureLine).setFontFamily(LABEL_FONT).setFontSize(11);

    if (name) {
        const sigStart = signatureLine.indexOf(signatureSpaced);
        if (sigStart !== -1) {
            const sigEnd = sigStart + signatureSpaced.length - 1;
            if (sigEnd >= sigStart && sigStart >= 0) {
                t.setFontFamily(sigStart, sigEnd, 'Dancing Script')
                    .setFontSize(sigStart, sigEnd, 16)
                    .setBold(sigStart, sigEnd, false)
                    .setUnderline(sigStart, sigEnd, false);
            }
        }
    }
    if (dateFormatted) formatValue(t, signatureLine, dateSpaced);
    p.setSpacingAfter(30);

    const footerPara = paragraph('Kings Equestrian Foundation | Karnataka, India | +91-9980895533 | info@kingsequestrian.com', 9, false, 0, DocumentApp.HorizontalAlignment.CENTER);
    footerPara.editAsText().setForegroundColor('#666666');

    doc.saveAndClose();

    const pdf = doc.getAs('application/pdf');
    pdf.setName(`Consent_Form_${(name || 'Participant').replace(/\s+/g, '_')}.pdf`);
    DriveApp.getFileById(doc.getId()).setTrashed(true);

    return pdf;
}

// --------------- GOOGLE CALENDAR INTEGRATION ---------------

function createBookingCalendarEvent(bookingData) {
    try {
        const calendar  = CalendarApp.getDefaultCalendar();
        const date      = new Date(bookingData.date);
        const timeSlots = String(bookingData.timeSlots).split(',');
        const firstSlot = timeSlots[0].trim();
        const timeParts = firstSlot.match(/(\d+):(\d+)\s*(AM|PM)/i);

        if (!timeParts) {
            Logger.log('Invalid time format: ' + firstSlot);
            return null;
        }

        let hours      = parseInt(timeParts[1]);
        const minutes  = parseInt(timeParts[2]);
        const period   = timeParts[3].toUpperCase();

        if (period === 'PM' && hours !== 12) hours += 12;
        if (period === 'AM' && hours === 12) hours = 0;

        const startTime = new Date(date);
        startTime.setHours(hours, minutes, 0);

        const endTime = new Date(startTime);
        endTime.setMinutes(endTime.getMinutes() + (timeSlots.length * 30));

        const participants    = bookingData.participants || 1;
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