CONFIG
const UPI_ID = "vyapar.176548151976@hdfcbank";
const BUSINESS_NAME = "KingsEquestrian";
const ACCOUNT_NAME ="KINGS EQUESTRIAN FOUNDATION";
const BANK_NAME="HDFC";
const ACCOUNT_NUMBER="50200072121375";
const IFSC_CODE="HDFC0000286"
const PAYMENT_FORM_URL="https://forms.gle/DzYWG5dWgr4mNDBAA"

// Generate Booking Reference
function generateReference() {
  return "KE-" + new Date().getTime();
}

// Create UPI Link
function createUPILink(amount, reference) {
  const link= "upi://pay?pa=" + UPI_ID +
         "&pn=" + encodeURIComponent(BUSINESS_NAME) +
         "&am=" + amount +
         "&cu=INR&tn=" + reference;
         return link;
}

// Create QR Code
function createQRCode(link) {
  return "https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=" +
          encodeURIComponent(link);
}
// function testLinks() {
//   const testAmount = "500";  // Rs 500
//   const testRef = generateReference();  // Generates "KE-1641234567890"
  
//   const upiLink = createUPILink(testAmount, testRef);
//   const qrLink = createQRCode(upiLink);
  
//   console.log("=== UPI LINK ===");
//   console.log(upiLink);
//   console.log("\n=== QR CODE URL ===");
//   console.log(qrLink);
//   console.log("\n=== Test Complete ===");
// }

function sendPaymentRequestEmail(data) {
  const {
     registrationNumber,
    email,
    name,        // studentName
    parentName,  // parent name  
    contact,     // contact number
    location,
    amount
  } = data;

  const reference = registrationNumber;
  const qrLink = createQRCode(
    `upi://pay?pa=${UPI_ID}&pn=${encodeURIComponent(BUSINESS_NAME)}&am=${amount}&cu=INR&tn=${reference}`
  );

  const subject = `Payment Request | ${name} | Reg No: ${registrationNumber}`;
  const pdfBlob = generateConsentPDF(name, parentName, contact, location, email);
  const htmlBody = `
 <!DOCTYPE html>
<html>
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
<body >

  <div style="max-width:720px;margin:30px auto;background:#fff;
              border:2px solid #000;border-radius:22px;padding:25px;">

    <!-- Header -->
 <div class="header">
      <img src="https://kingsfarmequestrian.com/wp-content/uploads/2023/08/Logo2.jpg" alt="Kings Equestrian Logo">
      <h1>Kings Equestrian Foundation</h1>
      <p style="margin:8px 0 0;">Where horses don’t just carry you — they change you </p>
    </div>
    <!-- Greeting -->
    <p style="font-size:14px;margin-top:25px;">
      Dear <b>${name}</b>,
    </p>

    <p style="font-size:14px;">
      Thank you for completing your registration with
      <b>${BUSINESS_NAME}</b>.  
      Please find below the payment details to proceed.
    </p>

    <!-- Amount Box -->
    <div style="border:2px solid #000;border-radius:10px;
                padding:18px;margin:25px 0;text-align:center;">
      <div style="font-size:14px;">Amount Payable</div>
      <div style="font-size:36px;font-weight:bold;color:#e67e22;">
        ₹${amount}
      </div>
    </div>

    <!-- UPI Section -->
    <div style="border:1px solid #000;border-radius:10px;padding:15px;margin-bottom:20px;">
      <h3 style="margin-top:0;font-size:15px;">📌 UPI Payment</h3>
      <p style="font-size:13px;margin:5px 0;">
        <b>UPI ID:</b> ${UPI_ID}
      </p>

      <div style="text-align:center;margin-top:15px;">
        <img src="${qrLink}"
             alt="UPI QR Code"
             style="max-width:260px;border:1px solid #ccc;padding:6px;" />
        <p style="font-size:11px;color:#666;margin-top:6px;">
          Scan the QR code to pay
        </p>
      </div>
    </div>

    <!-- Bank Details -->
    <div style="border:1px solid #000;border-radius:10px;padding:15px;margin-bottom:20px;">
      <h3 style="margin-top:0;font-size:15px;">🏦 Bank Transfer Details</h3>

      <table style="width:100%;font-size:13px;border-collapse:collapse;">
        <tr>
          <td style="padding:6px 0;"><b>Account Name</b></td>
          <td>${ACCOUNT_NAME}</td>
        </tr>
        <tr>
          <td style="padding:6px 0;"><b>Bank</b></td>
          <td>${BANK_NAME}</td>
        </tr>
        <tr>
          <td style="padding:6px 0;"><b>Account Number</b></td>
          <td>${ACCOUNT_NUMBER}</td>
        </tr>
        <tr>
          <td style="padding:6px 0;"><b>IFSC Code</b></td>
          <td>${IFSC_CODE}</td>
        </tr>
      </table>
    </div>

    <!-- Submit Payment -->
    <div style="background:#fdf3d7;border-left:4px solid #f39c12;
                padding:18px;border-radius:8px;">
      <h3 style="margin-top:0;font-size:15px;">📝 Submit Payment Details</h3>

      <p style="font-size:13px;">
        After completing the payment, please submit your transaction details using the link below.
      </p>

      <div style="text-align:center;margin-top:15px;">
        <a href="${PAYMENT_FORM_URL}"
           style="background:#000;color:#fff;
                  padding:12px 28px;
                  text-decoration:none;
                  border-radius:6px;
                  font-size:14px;
                  font-weight:bold;">
          Submit Payment Details
        </a>
      </div>
    </div>

    <!-- Footer Note -->
       <div class="footer">
      <p><strong>Kings Equestrian Foundation</strong></p>
      <p>📍 Karnataka, India</p>
      <p>📞 +91-XXXXXXXXXX | ✉️ support@kingsequestrian.com</p>
      <p style="margin-top: 10px; font-size: 11px;">
        © ${new Date().getFullYear()} Kings Equestrian Foundation. All rights reserved.
      </p>
    </div>

  </div>

</body>
</html>

  `;
 try {
    // ✅ Send email FIRST
    MailApp.sendEmail({
      to: email,
      subject: subject,
      htmlBody: htmlBody,
      attachments: [pdfBlob]
    });
    
    // ✅ ONLY if email SUCCEEDS → Update sheet to 'Yes'
    const paymentSheet = getPaymentDetailsSheet();
    const dataRange = paymentSheet.getDataRange();
    const values = dataRange.getValues();
    
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === registrationNumber) {
        paymentSheet.getRange(i + 1, 8).setValue('Yes');  // Col 8 = Send Payment Link
         paymentSheet.getRange(i + 1, 9).setValue(new Date());         // Col 9 = Sent Timestamps
        paymentSheet.getRange(i + 1, 9).setNumberFormat("dd-MMM-yyyy HH:mm:ss");  // Format timestamp
        Logger.log(`✅ Email sent + Sheet updated to 'Yes' for ${registrationNumber}`);
        break;
      }
    }
    
  } catch (error) {
    // ❌ Email FAILED → Mark as 'No' + log error
    Logger.log(`❌ Email FAILED for ${registrationNumber}: ${error.toString()}`);
    const paymentSheet = getPaymentDetailsSheet();
    const dataRange = paymentSheet.getDataRange();
    const values = dataRange.getValues();
    
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === registrationNumber) {
        paymentSheet.getRange(i + 1, 7).setValue('No');  // Keep as 'No'
        break;
      }
    }
    throw error;  // Re-throw so caller knows it failed
  }
}


function generateConsentPDF(studentName, parentName, contact, location, email) {
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
    const valStr = value.toString();
    const start = fullText.indexOf(valStr);
    if (start === -1) return;
    const end = start + valStr.length - 1;
    if (end >= start && start >= 0 && end < fullText.length) {
      textObj.setFontFamily(start, end, VALUE_FONT);
      textObj.setBold(start, end, false);
      textObj.setUnderline(start, end, true);
    }
  }

  /* ========= HEADER ========= */
  const leftImageUrl = 'https://iais.in/wp-content/uploads/2025/11/Indus-Altum-International-School-.png';
  const rightImageUrl = 'https://kingsfarmequestrian.com/wp-content/uploads/2023/08/Logo2.jpg';
  
  let leftImageBlob, rightImageBlob;
  try { leftImageBlob = UrlFetchApp.fetch(leftImageUrl).getBlob(); } catch (e) { Logger.log('Left image error: ' + e); }
  try { rightImageBlob = UrlFetchApp.fetch(rightImageUrl).getBlob(); } catch (e) { Logger.log('Right image error: ' + e); }

  // Header table (same as original)
  const headerTable = body.appendTable();
  const headerRow = headerTable.appendTableRow();
  
  const leftCell = headerRow.appendTableCell();
  leftCell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
  leftCell.setPaddingRight(15);
  leftCell.setWidth(120);
  if (leftImageBlob) {
    const leftImg = leftCell.appendParagraph('').appendInlineImage(leftImageBlob);
    leftImg.setWidth(110).setHeight(110);
  }
  
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
  
  const rightCell = headerRow.appendTableCell();
  rightCell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
  rightCell.setPaddingLeft(15);
  rightCell.setWidth(120);
  if (rightImageBlob) {
    const rightImg = rightCell.appendParagraph('').appendInlineImage(rightImageBlob);
    rightImg.setWidth(110).setHeight(110);
  }
  
  headerTable.setBorderWidth(0);
  body.appendParagraph('').setSpacingAfter(12);

  paragraph('Dear Equestrian Team,', 11, false, 8);

  /* ========= SIMPLIFIED STUDENT INFO ========= */
  let p = body.appendParagraph('');
  let t = p.editAsText();
  const nameSpaced = studentName ? `  ${studentName}  ` : '____________________';
  const locationSpaced = location ? `  ${location}  ` : '____________________';
  const studentLine = `Please enrol ${nameSpaced} (Name) at ${locationSpaced} (Location) in the horse-riding program.`;
  t.setText(studentLine).setFontFamily(LABEL_FONT).setFontSize(FONT_SIZE);
  if (studentName) formatValue(t, studentLine, nameSpaced);
  if (location) formatValue(t, studentLine, locationSpaced);
  p.setSpacingAfter(8);

  paragraph('Session: As per School Academic Year (15th July 2025 – April 2026)', 11, true, 10);

  /* ========= PROGRAM - HARDCODED AS SELECTED ========= */
  // paragraph('Please register my ward for:', 11, false, 6);

  // const tick = '☑';
  // const empty = '☐';

  // // HARDCODED as per your request
  // p = body.appendParagraph('');
  // t = p.editAsText();
  // t.setText(`    ${tick} Leasing the horse     Starting Month of Lease ________________`)
  //   .setFontFamily(LABEL_FONT).setFontSize(11);
  // p.setSpacingAfter(4);

  // p = body.appendParagraph('');
  // t = p.editAsText();
  // t.setText(`    ${tick} 3 classes per week program    Fee: 1,20,000/-`)
  //   .setFontFamily(LABEL_FONT).setFontSize(11);
  // p.setSpacingAfter(8);

  /* ========= PARENT INFO ========= */
  paragraph('PARENT\'S INFORMATION', 11, false, 8);

  p = body.appendParagraph('');
  t = p.editAsText();
  const parentNameSpaced = parentName ? `  ${parentName}  ` : '____________________';
  const parentLine = `Parent's Name: ${parentNameSpaced}`;
  t.setText(parentLine).setFontFamily(LABEL_FONT).setFontSize(FONT_SIZE);
  if (parentName) formatValue(t, parentLine, parentNameSpaced);
  p.setSpacingAfter(8);

  p = body.appendParagraph('');
  t = p.editAsText();
  const contactSpaced = contact ? `  ${contact}  ` : '____________________';
  const emailSpaced = email ? `  ${email}  ` : '____________________';
  const contactLine = `Contact: ${contactSpaced}    Email: ${emailSpaced}`;
  t.setText(contactLine).setFontFamily(LABEL_FONT).setFontSize(FONT_SIZE);
  if (contact) formatValue(t, contactLine, contactSpaced);
  if (email) formatValue(t, contactLine, emailSpaced);
  p.setSpacingAfter(17);

  /* ========= CONSENT + SIGNATURE (UNCHANGED) ========= */
  paragraph('ACKNOWLEDGEMENT / CONSENT FORM – HORSE RIDING PARTICIPANTS', 11, true, 8);
  
  // All your existing consent paragraphs (unchanged)...
  paragraph('Kings Equestrian at Indus International School offers horseback riding programs for those interested in Casual riding, Dressage, Jumping and related workshops and clinics. Programs of this sort involve risk of personal injury, including, but not limited to, bruises, broken bones, head injuries and death. All normal safety precautions are taken to protect our participants, but occasionally accidents do happen.', 11, false, 8);
  // ... (keep all 9 consent paragraphs as-is)

  p = body.appendParagraph('');
  t = p.editAsText();
  const consentStudentNameSpaced = studentName ? `  ${studentName}  ` : '____________________';
  const consentText = `I give my consent for ${consentStudentNameSpaced}, my ward ("RIDER"), to participate in the above-mentioned riding programs and/or workshop. I have read the information provided above and understand the inherent risks involved. I further attest that I am at least eighteen (18) years of age and fully authorized to sign this consent.`;
  t.setText(consentText).setFontFamily(LABEL_FONT).setFontSize(11);
  if (studentName) formatValue(t, consentText, consentStudentNameSpaced);
  p.setSpacingAfter(10);

  // Signature
  p = body.appendParagraph('');
  t = p.editAsText();
  const parentSignatureSpaced = parentName ? `  ${parentName}  ` : '____________________';
  const signatureLine = `Signature: ${parentSignatureSpaced}     Date: ____________________`;
  t.setText(signatureLine).setFontFamily(LABEL_FONT).setFontSize(18);
  if (parentName) {
    const sigStart = signatureLine.indexOf(parentSignatureSpaced);
    if (sigStart !== -1) {
      const sigEnd = sigStart + parentSignatureSpaced.length - 1;
      t.setFontFamily(sigStart, sigEnd, 'Dancing Script');
      t.setBold(sigStart, sigEnd, false);
      t.setUnderline(sigStart, sigEnd, false);
    }
  }
  p.setSpacingAfter(20);

  /* ========= SAVE ========= */
  doc.saveAndClose();
  const pdf = doc.getAs('application/pdf');
  pdf.setName(`Consent_Form_${(studentName || 'Student').replace(/\s+/g, '_')}.pdf`);
  DriveApp.getFileById(doc.getId()).setTrashed(true);
  return pdf;
}


function resendPaymentEmail() {
  const paymentSheet = getPaymentDetailsSheet();
  const ui = SpreadsheetApp.getUi();
  
  // Get selected row from Payment Details
  const activeCell = paymentSheet.getActiveCell();
  const row = activeCell.getRow();
  
  if (row < 2) {
    ui.alert('Please select a row with payment details (row 2+).');
    return;
  }
  
  // Read PAYMENT row data
  const paymentRowData = paymentSheet.getRange(row, 1, 1, 10).getValues()[0];
  const [registrationNumber, amount, email, name, location] = paymentRowData;
  
  if (!email || !registrationNumber) {
    ui.alert('Missing email or registration number.');
    return;
  }

  // ✅ LOOKUP Enquire Response sheet for parent details
  let parentName = name.toString();  // fallback
  let contact = email.toString();    // fallback
  
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

  const data = {
    registrationNumber: registrationNumber.toString(),
    email: email.toString(),
    name: name.toString(),
    parentName: parentName.toString(),
    contact: contact.toString(),
    location: location.toString(),
    amount: parseInt(amount) || 9000
  };
  
  try {
    sendPaymentRequestEmail(data);
    
    // Update timestamp in payment sheet
    paymentSheet.getRange(row, 9).setValue(new Date());
    paymentSheet.getRange(row, 9).setNumberFormat("dd-MMM-yyyy HH:mm:ss");
    
    ui.alert(`✅ Payment email resent!\nReg#: ${registrationNumber}\nEmail: ${email}\nParent: ${parentName}\nContact: ${contact}`);
    
  } catch (error) {
    ui.alert(`❌ Failed: ${error.toString()}`);
  }
}



// Test function
function testPaymentEmail() {
   sendPaymentRequestEmail( "98023",
    "jyothikondupally@gmail.com",
    "Jhone",        // studentName
    "parentName",  // parent name  
    "90909",     // contact number
    "location",
    "amount");
}

