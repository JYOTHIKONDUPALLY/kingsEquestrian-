// Google Apps Script for sending consent form PDFs via email
// Attach this to your Google Sheet (not the form)

function onFormSubmit(e) {
    try {
        // Get the active spreadsheet
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
        const lastRow = sheet.getLastRow();

        // Get headers from row 1
        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

        // Get the last submitted row data
        const rowData = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).getValues()[0];

        // Create a map of column names to values
        const data = {};
        headers.forEach((header, index) => {
            const cleanHeader = header
                .replace(/['']/g, "'")   // normalize smart quotes
                .trim();
            data[cleanHeader] = rowData[index];
        });


        // Extract data from the row - trying multiple possible column names
        const email = data['Email address'] || data['Email Address (for consent copy & communication)'] || data['Email Address'];
        const studentName = data['Student Name (Rider’s Name)'] || data['Student Name'];
        const program = data['Program Selection'];
        const horseLease = data['Horse Lease Option'];
        const age = data['Rider’s Age'];
        const dob = data['Date of Birth'];
        const grade = data['Grade & Section'];
        const location = data['Location'];
        const address = data['Residential Address'];
        const motherName = data['Mother’s Name'];
        const fatherName = data['Father’s Name'];
        const motherContact = data['Mother’s Contact Number'] || '';
        const motherWhatsApp = data['Mother’s WhatsApp Number'] || '';
        const fatherContact = data['Father’s Contact Number'] || '';
        const fatherWhatsApp = data['Father’s WhatsApp Number'] || '';
        const emergencyContact = data['Emergency Contact Name & Phone Number'];
        const parentName = data['Name of Parent/Guardian (Digital Signature)'];
        const relationship = data['Relationship to Rider'];
        const consentDate = data['Date of Consent Submission'];

        // Debug log to see what data we're getting
        Logger.log('Email found: ' + email);
        Logger.log('Student name found: ' + studentName);
        Logger.log('Program found: ' + program);
        Logger.log('Available headers: ' + headers.join(', '));

        // Validate required fields
        if (!email || !studentName) {
            Logger.log('Missing required fields: email or student name');
            return;
        }

        // Determine session dates based on location
        let sessionDates = '';
        if (location && location.toLowerCase().includes('hyderabad')) {
            sessionDates = '15th July 2025 – April 2026';
        } else {
            sessionDates = 'August 2025 – 15th May 2026';
        }

        // Generate PDF
        const pdf = generateConsentPDF(studentName, program, horseLease, age, dob, grade,
            location, address, motherName, fatherName,
            motherContact, motherWhatsApp, fatherContact, fatherWhatsApp,
            email, emergencyContact, parentName, relationship,
            consentDate, sessionDates);

        // Send email with PDF attachment
        sendConsentEmail(email, studentName, pdf, parentName);

        // Mark as sent in the sheet (optional - add a status column)
        // sheet.getRange(lastRow, sheet.getLastColumn() + 1).setValue('Sent');

        Logger.log('Consent form sent successfully to: ' + email);
    } catch (error) {
        Logger.log('Error: ' + error.toString());
        SpreadsheetApp.getUi().alert('Error sending email: ' + error.toString());
    }
}

function generateConsentPDF(
    studentName, program, horseLease, age, dob, grade, location,
    address, motherName, fatherName, motherContact, motherWhatsApp,
    fatherContact, fatherWhatsApp, email, emergencyContact,
    parentName, relationship, consentDate, sessionDates
) {

    const LABEL_FONT = 'Comic Sans MS';
    const VALUE_FONT = 'Arial';
    const FONT_SIZE = 11;

    const doc = DocumentApp.create('Consent Form - ' + (studentName || 'Student'));
    const body = doc.getBody();
    body.clear();

    body.setMarginTop(40);
    body.setMarginBottom(40);
    body.setMarginLeft(40);
    body.setMarginRight(40);

    /* ========= HELPERS ========= */

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
            textObj.setFontFamily(start, end, VALUE_FONT);
            textObj.setBold(start, end, false);
            textObj.setUnderline(start, end, true);
            
            // Add extra space before and after the value for better visibility
            // Add a thin space character before and after (optional visual enhancement)
        }
    }
    
    function formatDate(dateValue) {
        // Helper to format date objects to readable string (DD/MM/YYYY or DD-MMM-YYYY)
        if (!dateValue) return null;
        
        // If it's already a string, return as is
        if (typeof dateValue === 'string') return dateValue;
        
        // If it's a Date object, format it
        if (dateValue instanceof Date) {
            const day = String(dateValue.getDate()).padStart(2, '0');
            const month = String(dateValue.getMonth() + 1).padStart(2, '0');
            const year = dateValue.getFullYear();
            return `${day}/${month}/${year}`;
        }
        
        return dateValue.toString();
    }

    function addUnderlinedValue(text, value, withSpacing = true) {
        // Helper to add values with spacing
        if (withSpacing && value) {
            return `  ${value}  `;
        }
        return value || '';
    }

    /* ========= HEADER ========= */

   
    // Fetch images
    const leftImageUrl = 'https://iais.in/wp-content/uploads/2025/11/Indus-Altum-International-School-.png';
    const rightImageUrl = 'https://kingsfarmequestrian.com/wp-content/uploads/2023/08/Logo2.jpg';
    
    let leftImageBlob, rightImageBlob;
    
    try {
        leftImageBlob = UrlFetchApp.fetch(leftImageUrl).getBlob();
    } catch (e) {
        Logger.log('Error fetching left image: ' + e);
    }
    
    try {
        rightImageBlob = UrlFetchApp.fetch(rightImageUrl).getBlob();
    } catch (e) {
        Logger.log('Error fetching right image: ' + e);
    }

    // Create header table with images and text
    const headerTable = body.appendTable();
    const headerRow = headerTable.appendTableRow();
    
    // Left cell - Indus logo
    const leftCell = headerRow.appendTableCell();
    leftCell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
    leftCell.setPaddingRight(15);
    leftCell.setWidth(120);
    if (leftImageBlob) {
        const leftImg = leftCell.appendParagraph('').appendInlineImage(leftImageBlob);
        leftImg.setWidth(110);
        leftImg.setHeight(110);
    }
    
    // Center cell - Text in 3 lines
    const centerCell = headerRow.appendTableCell();
    centerCell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
    
    let centerPara = centerCell.appendParagraph('INDUS EQUESTRIAN CENTRE OF EXCELLENCE');
    centerPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    centerPara.editAsText().setFontFamily(LABEL_FONT).setFontSize(14).setBold(true);
    centerPara.setSpacingAfter(4);
    
    centerPara = centerCell.appendParagraph('KINGS EQUESTRIAN FOUNDATION HORSE RIDING');
    centerPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    centerPara.editAsText().setFontFamily(LABEL_FONT).setFontSize(12).setBold(true);
    centerPara.setSpacingAfter(4);
    
    centerPara = centerCell.appendParagraph('REGISTRATION (2025–26)');
    centerPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    centerPara.editAsText().setFontFamily(LABEL_FONT).setFontSize(11).setBold(true);
    centerPara.setSpacingAfter(0);
    
    // Right cell - Kings Equestrian logo
    const rightCell = headerRow.appendTableCell();
    rightCell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
    rightCell.setPaddingLeft(15);
    rightCell.setWidth(120);
    if (rightImageBlob) {
        const rightImg = rightCell.appendParagraph('').appendInlineImage(rightImageBlob);
        rightImg.setWidth(110);
        rightImg.setHeight(110);
    }
    
    // Remove table borders
    headerTable.setBorderWidth(0);
    
    // Add spacing after header
    body.appendParagraph('').setSpacingAfter(12);

    paragraph('Dear Equestrian Team,', 11, false, 8);

    /* ========= ENROLMENT ========= */

    let p = body.appendParagraph('');
    let t = p.editAsText();
    const nameVal = studentName ? `  ${studentName}  ` : '____________________________________________';
    const enrollText = `Please enrol ${nameVal} (Name) in the horse-riding program.`;
    t.setText(enrollText).setFontFamily(LABEL_FONT).setFontSize(FONT_SIZE);
    if (studentName) formatValue(t, enrollText, `  ${studentName}  `);
    p.setSpacingAfter(8);

    const sessionText = `Session: As per School Academic Year (${sessionDates || '15th July 2025 – April 2026'})`;
    paragraph(sessionText, 11, true, 10);

    /* ========= PROGRAM ========= */

    paragraph('Please register my ward for:', 11, false, 6);

    const tick = '☑';
    const empty = '☐';

    function selected(programValue, keyword) {
        return programValue && programValue.toLowerCase().includes(keyword);
    }

    p = body.appendParagraph('');
    t = p.editAsText();
    t.setText(`    ${horseLease ? tick : empty} Leasing the horse     Starting Month of Lease ________________`)
        .setFontFamily(LABEL_FONT).setFontSize(11);
    p.setSpacingAfter(4);

    p = body.appendParagraph('');
    t = p.editAsText();
    t.setText(`    ${selected(program, '3 classes') ? tick : empty} 3 classes per week program    Fee: 1,20,000/-`)
        .setFontFamily(LABEL_FONT).setFontSize(11);
    p.setSpacingAfter(4);

    p = body.appendParagraph('');
    t = p.editAsText();
    t.setText(`    ${selected(program, '2 classes') ? tick : empty} 2 classes per week program   Fee: 90,000/-`)
        .setFontFamily(LABEL_FONT).setFontSize(11);
    p.setSpacingAfter(8);

    paragraph(
        'Note: For PYP (Grade 1–5) 3 classes per week option is available. MYP (Grade 6 and above) can choose 2 or 3 classes per week. Lease option is open for all.',
        10.5,
        false,
        17
    );

    /* ========= RIDER INFO ========= */

    paragraph('RIDER’S INFORMATION', 11, false, 8);

    // Rider's Name, Age, DOB, Grade
    p = body.appendParagraph('');
    t = p.editAsText();
   const riderNameSpaced = studentName ? `  ${studentName}  ` : '____________________';
    const ageSpaced = age ? `  ${age}  ` : '__________';
    const dobFormatted = formatDate(dob);
    const dobSpaced = dobFormatted ? `  ${dobFormatted}  ` : '__________';
    const gradeSpaced = grade ? `  ${grade}  ` : '__________________';
    const riderLine = `Rider's Name: ${riderNameSpaced}     Age: ${ageSpaced}     \nDate of Birth: ${dobSpaced}     Grade & Section: ${gradeSpaced}`;
    t.setText(riderLine).setFontFamily(LABEL_FONT).setFontSize(FONT_SIZE);
    if (studentName) formatValue(t, riderLine, riderNameSpaced);
    if (age) formatValue(t, riderLine, ageSpaced);
    if (dobFormatted) formatValue(t, riderLine, dobSpaced);
    if (grade) formatValue(t, riderLine, gradeSpaced);
    p.setSpacingAfter(8);

    // Address
    p = body.appendParagraph('');
    t = p.editAsText();
    const addressSpaced = address ? `  ${address}  ` : '_______________________________________________________________________________';
    const addressLine = `Address: ${addressSpaced}`;
    t.setText(addressLine).setFontFamily(LABEL_FONT).setFontSize(FONT_SIZE);
    if (address) formatValue(t, addressLine, addressSpaced);
    p.setSpacingAfter(8);

    // Parent's Name
    p = body.appendParagraph('');
    t = p.editAsText();
    const motherNameSpaced = motherName ? `  ${motherName}  ` : '____________________';
    const fatherNameSpaced = fatherName ? `  ${fatherName}  ` : '____________________';
    const parentLine = `Parent's Name:   \nMother: ${motherNameSpaced}   \nFather: ${fatherNameSpaced}`;
    t.setText(parentLine).setFontFamily(LABEL_FONT).setFontSize(FONT_SIZE);
    if (motherName) formatValue(t, parentLine, motherNameSpaced);
    if (fatherName) formatValue(t, parentLine, fatherNameSpaced);
    p.setSpacingAfter(6);

    // Contact Number label
    paragraph('Parent’s Contact Number:', 11, false, 2);

    // Mother contact
    p = body.appendParagraph('');
    t = p.editAsText();
    const motherContactSpaced = motherContact ? `  ${motherContact}  ` : '____________________';
    const motherWhatsAppSpaced = motherWhatsApp ? `  ${motherWhatsApp}  ` : '____________________';
    const motherContactLine = `    Mother ${motherContactSpaced}     WhatsApp ${motherWhatsAppSpaced}`;
    t.setText(motherContactLine).setFontFamily(LABEL_FONT).setFontSize(FONT_SIZE);
    if (motherContact) formatValue(t, motherContactLine, motherContactSpaced);
    if (motherWhatsApp) formatValue(t, motherContactLine, motherWhatsAppSpaced);
    p.setSpacingAfter(4);

    // Father contact
    p = body.appendParagraph('');
    t = p.editAsText();
    const fatherContactSpaced = fatherContact ? `  ${fatherContact}  ` : '____________________';
    const fatherWhatsAppSpaced = fatherWhatsApp ? `  ${fatherWhatsApp}  ` : '____________________';
    const fatherContactLine = `    Father ${fatherContactSpaced}     WhatsApp ${fatherWhatsAppSpaced}`;
    t.setText(fatherContactLine).setFontFamily(LABEL_FONT).setFontSize(FONT_SIZE);
    if (fatherContact) formatValue(t, fatherContactLine, fatherContactSpaced);
    if (fatherWhatsApp) formatValue(t, fatherContactLine, fatherWhatsAppSpaced);
    p.setSpacingAfter(8);

    // Email
    p = body.appendParagraph('');
    t = p.editAsText();
    const emailSpaced = email ? `  ${email}  ` : '_______________________________________________________________________________';
    const emailLine = `Email: ${emailSpaced}`;
    t.setText(emailLine).setFontFamily(LABEL_FONT).setFontSize(FONT_SIZE);
    if (email) formatValue(t, emailLine, emailSpaced);
    p.setSpacingAfter(8);

    // Emergency Contact
    p = body.appendParagraph('');
    t = p.editAsText();
    const emergencySpaced = emergencyContact ? `  ${emergencyContact}  ` : '____________________';
    const emergencyLine = `In Case of Emergency Call: ${emergencySpaced} (Phone Number)`;
    t.setText(emergencyLine).setFontFamily(LABEL_FONT).setFontSize(FONT_SIZE);
    if (emergencyContact) formatValue(t, emergencyLine, emergencySpaced);
    p.setSpacingAfter(17);

    /* ========= CONSENT SECTION ========= */

    paragraph('ACKNOWLEDGEMENT / CONSENT FORM – HORSE RIDING PARTICIPANTS', 11, true, 8);

    paragraph(
        'Kings Equestrian at Indus International School offers horseback riding programs for those interested in Casual riding, Dressage, Jumping and related workshops and clinics. Programs of this sort involve risk of personal injury, including, but not limited to, bruises, broken bones, head injuries and death. All normal safety precautions are taken to protect our participants, but occasionally accidents do happen.',
        11, false, 8
    );

    paragraph(
        'Horses may, without warning or apparent cause, buck, rear, stumble, fall, spook or make unanticipated movements, jump obstacles in their path, bite, kick, step on a person’s foot, or push or shove a person, and saddles or bridles may loosen or break, all of which may result in injury.',
        11, false, 8
    );

    paragraph(
        'We further note that if you are pregnant or immunocompromised you may be at a greater risk of injury and/or of contracting possible zoonotic agents due to your close proximity to animals and should consult your healthcare provider before undertaking equestrian activities.',
        11, false, 8
    );

    paragraph(
        'The school does not provide insurance to program participants. We strongly suggest that you should be covered under your own private insurance plan.',
        11, false, 8
    );

    paragraph(
        'All fees/funds will be used for animal welfare and well-being including feeding the horse, maintenance of the stable, and horse upkeep. Fees will not be refunded after enrolment.',
        11, false, 8
    );

    paragraph(
        'Kings Equestrian will not be responsible for any cancellations due to unpredicted weather, school leaves, school camps, student health, etc.',
        11, false, 8
    );

    paragraph(
        'Missed classes can only be accommodated in that particular week/weekend and will automatically lapse.',
        11, false, 8
    );

    paragraph(
        'Helmet, body protector, and shoes are mandatory for horse riding. Students and parents are responsible for having the required riding gear.',
        11, false, 8
    );

    paragraph(
        'Competition-related travel and expenses are not included and will be paid separately.',
        11, false, 8
    );

    // Consent paragraph
    p = body.appendParagraph('');
    t = p.editAsText();
    const consentStudentNameSpaced = studentName ? `  ${studentName}  ` : '____________________';
    const consentRelationshipSpaced = relationship ? `  ${relationship}  ` : 'ward';
    const consentText = `I give my consent for ${consentStudentNameSpaced}, my ${consentRelationshipSpaced} ("RIDER"), to participate in the above-mentioned riding programs and/or workshop. I have read the information provided above and understand the inherent risks involved. I further attest that I am at least eighteen (18) years of age and fully authorized to sign this consent.`;
    t.setText(consentText).setFontFamily(LABEL_FONT).setFontSize(11);
    if (studentName) formatValue(t, consentText, consentStudentNameSpaced);
    if (relationship) formatValue(t, consentText, consentRelationshipSpaced);
    p.setSpacingAfter(10);

    /* ========= SIGNATURE SECTION ========= */

   
    // Parent Name
    p = body.appendParagraph('');
    t = p.editAsText();
    const parentNameSpaced = parentName ? `  ${parentName}  ` : '____________________';
    const parentNameLine = `Name (Parent): ${parentNameSpaced}`;
    t.setText(parentNameLine).setFontFamily(LABEL_FONT).setFontSize(11);
    if (parentName) formatValue(t, parentNameLine, parentNameSpaced);
    p.setSpacingAfter(10);

    // Signature and Date
    p = body.appendParagraph('');
    t = p.editAsText();
    const parentSignatureSpaced = parentName ? `  ${parentName}  ` : '____________________';
    const consentDateFormatted = formatDate(consentDate);
    const consentDateSpaced = consentDateFormatted ? `  ${consentDateFormatted}  ` : '____________________';
    const signatureLine = `Signature: ${parentSignatureSpaced}     Date: ${consentDateSpaced}`;
    t.setText(signatureLine).setFontFamily(LABEL_FONT).setFontSize(18);
    
    // Apply Dancing Script to the signature (parent name)
    if (parentName) {
        const sigStart = signatureLine.indexOf(parentSignatureSpaced);
        if (sigStart !== -1) {
            const sigEnd = sigStart + parentSignatureSpaced.length - 1;
            if (sigEnd >= sigStart && sigStart >= 0) {
                t.setFontFamily(sigStart, sigEnd, 'Dancing Script');
                t.setBold(sigStart, sigEnd, false);
                t.setUnderline(sigStart, sigEnd, false);
            }
        }
    }
    
    // Format the date
    if (consentDateFormatted) formatValue(t, signatureLine, consentDateSpaced);
    p.setSpacingAfter(20);


    /* ========= SAVE ========= */

    doc.saveAndClose();

    const pdf = doc.getAs('application/pdf');
    pdf.setName(`Consent_Form_${(studentName || 'Student').replace(/\s+/g, '_')}.pdf`);

    DriveApp.getFileById(doc.getId()).setTrashed(true);
    return pdf;
}

function sendConsentEmail(recipientEmail, studentName, pdfBlob, parentName) {
    const subject = 'Horse Riding Consent Form - ' + studentName;

    const body = `Dear ${parentName || 'Parent/Guardian'},

Thank you for enrolling ${studentName} in the Kings Equestrian horse riding program at Indus International School.

Please find attached the signed consent form for your records. This document confirms your registration and acknowledgment of the program terms and conditions.

Program Details:
- Session: 2025-26 Academic Year
- All fees contribute to animal welfare, stable maintenance, and horse upkeep
- Mandatory safety gear: Helmet, body protector, and appropriate shoes

If you have any questions or need further assistance, please don't hesitate to contact us.

Best regards,
Kings Equestrian Foundation
Indus Equestrian Centre of Excellence

---
This is an automated email. Please do not reply to this address.`;

    MailApp.sendEmail({
        to: recipientEmail,
        subject: subject,
        body: body,
        attachments: [pdfBlob]
    });
}

// Function to generate unique registration number
function generateRegistrationNumber(location, sheet) {
    // Determine location code
    let locationCode = '';
    const loc = (location || '').toString().toLowerCase();
    
    if (loc.includes('bangalore') || loc.includes('bengaluru')) {
        locationCode = 'BLR';
    } else if (loc.includes('hyderabad')) {
        locationCode = 'HYD';
    } else if (loc.includes('pune')) {
        locationCode = 'PUNE';
    } else if (loc.includes('farm')) {
        locationCode = 'FARM';
    } else {
        locationCode = 'OTH'; // Other locations
    }
    
    // Get all existing registration numbers for this location
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const regNumCol = headers.indexOf('Registration Number') + 1;
    
    if (regNumCol === 0) {
        throw new Error('Registration Number column not found');
    }
    
    // Get all registration numbers
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
        return locationCode + '1';
    }
    
    const allRegNums = sheet.getRange(2, regNumCol, lastRow - 1, 1).getValues();
    
    // Find highest number for this location
    let maxNum = 0;
    allRegNums.forEach(row => {
        const regNum = (row[0] || '').toString();
        if (regNum.startsWith(locationCode)) {
            const num = parseInt(regNum.replace(locationCode, ''));
            if (!isNaN(num) && num > maxNum) {
                maxNum = num;
            }
        }
    });
    
    return locationCode + (maxNum + 1);
}

// Function to process and send email for a specific row
function processRowEmail(sheet, rowNumber, isFormSubmit = false) {
    const ui = SpreadsheetApp.getUi();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const rowData = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Find column indices
    const emailSentCol = headers.indexOf('Email Sent') + 1;
    const regNumCol = headers.indexOf('Registration Number') + 1;
    
    if (emailSentCol === 0) {
        throw new Error('Email Sent column not found');
    }
    if (regNumCol === 0) {
        throw new Error('Registration Number column not found');
    }
    
    // Check if email already sent (skip if already sent and not manual trigger)
    const emailSentValue = rowData[emailSentCol - 1];
    if (!isFormSubmit && emailSentValue) {
        if (ui) {
            ui.alert('Email already sent for this row on: ' + emailSentValue);
        }
        return;
    }
    
    // Create data object
    const data = {};
    headers.forEach((header, index) => {
        data[header] = rowData[index];
    });
    
    const email = data['Email address'] || data['Email Address (for consent copy & communication)'] || data['Email Address'];
    const studentName = data['Student Name (Rider’s Name)'] || data['Student Name'];
    
    if (!email || !studentName) {
        throw new Error('Missing email or student name in this row');
    }
    
    // Get all required data
    const program = data['Program Selection'];
    const horseLease = data['Horse Lease Option'];
    const age = data['Rider’s Age'];
    const dob = data['Date of Birth'];
    const grade = data['Grade & Section'];
    const location = data['Location'];
    const address = data['Residential Address'];
    const motherName = data['Mother’s Name'];
    const fatherName = data['Father’s Name'];
    const motherContact = data['Mother’s Contact Number'] || '';
    const motherWhatsApp = data['Mother’s WhatsApp Number'] || '';
    const fatherContact = data['Father’s Contact Number'] || '';
    const fatherWhatsApp = data['Father’s WhatsApp Number'] || '';
    const emergencyContact = data['Emergency Contact Name & Phone Number'];
    const parentName = data['Name of Parent/Guardian (Digital Signature)'];
    const relationship = data['Relationship to Rider'];
    const consentDate = data['Date of Consent Submission'];
    
    // Determine session dates based on location
    let sessionDates = '';
    if (location && location.toLowerCase().includes('hyderabad')) {
        sessionDates = '15th July 2025 – April 2026';
    } else {
        sessionDates = 'August 2025 – 15th May 2026';
    }
    
    // Generate registration number if not exists
    let registrationNumber = data['Registration Number'];
    if (!registrationNumber) {
        registrationNumber = generateRegistrationNumber(location, sheet);
        sheet.getRange(rowNumber, regNumCol).setValue(registrationNumber);
    }
    
    // Generate PDF
    const pdf = generateConsentPDF(
        studentName, program, horseLease, age, dob, grade,
        location, address, motherName, fatherName,
        motherContact, motherWhatsApp, fatherContact, fatherWhatsApp,
        email, emergencyContact, parentName, relationship,
        consentDate, sessionDates
    );
    
    // Send email
    sendConsentEmail(email, studentName, pdf, parentName, registrationNumber);
    
    // Update Email Sent timestamp
    const timestamp = new Date();
    sheet.getRange(rowNumber, emailSentCol).setValue(timestamp);
    
    return {
        email: email,
        studentName: studentName,
        registrationNumber: registrationNumber,
        timestamp: timestamp
    };
}

// Manual function to send email for a specific row
function sendEmailForRow() {
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt('Enter Row Number', 'Enter the row number to send email for:', ui.ButtonSet.OK_CANCEL);
    
    if (response.getSelectedButton() == ui.Button.OK) {
        const rowNumber = parseInt(response.getResponseText());
        
        if (isNaN(rowNumber) || rowNumber < 2) {
            ui.alert('Please enter a valid row number (2 or greater)');
            return;
        }
        
        try {
            const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
            const result = processRowEmail(sheet, rowNumber, false);
            
            ui.alert(
                'Success!', 
                `Registration Complete!\n\n` +
                `Registration Number: ${result.registrationNumber}\n` +
                `Student: ${result.studentName}\n` +
                `Email sent to: ${result.email}\n` +
                `Time: ${result.timestamp}`,
                ui.ButtonSet.OK
            );
        } catch (error) {
            ui.alert('Error: ' + error.toString());
        }
    }
}

// Function to run on form submit
function onFormSubmit(e) {
    try {
        const sheet = e.range.getSheet();
        const rowNumber = e.range.getRow();
        
        // Process and send email
        const result = processRowEmail(sheet, rowNumber, true);
        
        Logger.log(`Form submitted - Registration: ${result.registrationNumber}, Email: ${result.email}`);
    } catch (error) {
        Logger.log('Error in onFormSubmit: ' + error.toString());
        
        // Optionally send error notification to admin
        // MailApp.sendEmail('admin@example.com', 'Form Submit Error', error.toString());
    }
}

// Updated email sending function with registration number
function sendConsentEmail(email, studentName, pdf, parentName, registrationNumber) {
    const subject = `Horse Riding Registration Confirmation - ${registrationNumber}`;
    
    const body = `Dear ${parentName || 'Parent/Guardian'},

Thank you for registering ${studentName} for our horse riding program!

Registration Number: ${registrationNumber}

Please find attached the completed consent form for your records. This form confirms ${studentName}'s enrollment in the Kings Equestrian Foundation Horse Riding program.

Important Information:
• Please keep this registration number for future reference
• The attached consent form should be kept for your records
• If you have any questions, please contact us with your registration number

We look forward to seeing ${studentName} at our equestrian center!

Best regards,
Kings Equestrian Foundation
Indus International School`;

    MailApp.sendEmail({
        to: email,
        subject: subject,
        body: body,
        attachments: [pdf]
    });
}

// Utility function to check and send emails for rows without Email Sent timestamp
function sendPendingEmails() {
    const ui = SpreadsheetApp.getUi();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const emailSentCol = headers.indexOf('Email Sent') + 1;
    
    if (emailSentCol === 0) {
        ui.alert('Email Sent column not found');
        return;
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
        ui.alert('No data rows found');
        return;
    }
    
    let sentCount = 0;
    let errorCount = 0;
    const errors = [];
    
    // Loop through all rows
    for (let i = 2; i <= lastRow; i++) {
        const emailSentValue = sheet.getRange(i, emailSentCol).getValue();
        
        // If Email Sent is empty, send email
        if (!emailSentValue) {
            try {
                processRowEmail(sheet, i, true);
                sentCount++;
                
                // Add a small delay to avoid hitting email quotas
                Utilities.sleep(1000);
            } catch (error) {
                errorCount++;
                errors.push(`Row ${i}: ${error.toString()}`);
                Logger.log(`Error processing row ${i}: ${error.toString()}`);
            }
        }
    }
    
    // Show summary
    let message = `Processing Complete!\n\n`;
    message += `Emails sent: ${sentCount}\n`;
    message += `Errors: ${errorCount}`;
    
    if (errors.length > 0) {
        message += `\n\nErrors:\n${errors.slice(0, 5).join('\n')}`;
        if (errors.length > 5) {
            message += `\n... and ${errors.length - 5} more`;
        }
    }
    
    ui.alert(message);
}

// Function to setup the form submit trigger (run this once)
function setupFormSubmitTrigger() {
    // Remove existing triggers for this function
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
        if (trigger.getHandlerFunction() === 'onFormSubmit') {
            ScriptApp.deleteTrigger(trigger);
        }
    });
    
    // Create new trigger
    const sheet = SpreadsheetApp.getActiveSpreadsheet();
    ScriptApp.newTrigger('onFormSubmit')
        .forSpreadsheet(sheet)
        .onFormSubmit()
        .create();
    
    SpreadsheetApp.getUi().alert('Form submit trigger has been set up successfully!');
}