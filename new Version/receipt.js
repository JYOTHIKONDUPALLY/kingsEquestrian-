function testGenerate80GReceipt() {
  const receiptNumber=generateReceiptNumber()
  const receiptPDF = generate80GReceipt('Mr. Anil Kumar', 'AABCD1234E', 20000, 'UPI-12345678', 'test@email.com', receiptNumber);
  const subject = "Hi Mr. Anil Kumar, Here is your payment receipt";
  const htmlBody = 'Find the below details of the payment receipt';
  Logger.log('✅ SUCCESS: ' + receiptPDF.getName());
  MailApp.sendEmail({
    to: 'jyothikondupally@gmail.com',
    subject: subject,
    htmlBody: htmlBody,
    attachments: [receiptPDF]
  });
}
function generateReceiptNumber() {
  const year = new Date().getFullYear();
  const props = PropertiesService.getScriptProperties();

  const key = `LAST_RECEIPT_${year}`;
  const lastNumber = Number(props.getProperty(key)) || 0;

  const newNumber = lastNumber + 1;
  props.setProperty(key, newNumber);

  // Format: 2026/REC/0001
  return `${year}/R/${String(newNumber).padStart(4, '0')}`;
}


/**
 * Convert Google Drive file to base64 string
 */
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

/**
 * Convert URL to base64 (for logo from website)
 */
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

function generate80GReceipt(donorName, pan, amount, transactionRef, email, receiptNumber) {
  // Convert images to base64
  Logger.log('Converting logo to base64...');
  const logoBase64 = getImageFromUrlAsBase64('https://kingsfarmequestrian.com/wp-content/uploads/2023/08/Logo2.jpg');
  
  Logger.log('Converting stamp to base64...');
  const stampBase64 = getImageAsBase64('1kkDoebRZYYJDW76jYNZT1rWX65_DjSMs');
  
  Logger.log('Converting signature to base64...');
  const signBase64 = getImageAsBase64('1z8rGx3HkgyBb-nqIXIT-_BgY8cqiQDRR');
  
  const htmlContent = createReceiptHTML(donorName, pan, amount, transactionRef, receiptNumber, logoBase64, stampBase64, signBase64);
  
  const htmlFile = DriveApp.createFile(`receipt_temp_${new Date().getTime()}.html`, htmlContent, MimeType.HTML);
  const blob = htmlFile.getAs('application/pdf');
  blob.setName(`80G_Receipt_${donorName.replace(/\s+/g, '_')}_${receiptNumber}.pdf`);
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
    <!-- Header -->
    <div class="header">
      <div class="logo-section">
        <img src="${logoBase64}" class="logo-img" alt="Logo">
        <div class="logo-text">KINGS EQUESTRIAN<br>SADOLI, UP.</div>
      </div>
      
      <div class="header-center">
        <div class="org-name">Kings Equestrian Foundation</div>
        <div class="registration-info">Registered u/s 80G of Income-tax Act, 1961, PAN: AA JCK7191E</div>
        <div class="location">Karnataka, India</div>
        <!-- Receipt Box -->
    <div class="receipt-box">
      <div class="receipt-title">Receipt</div>
      <div class="receipt-subtitle">This receipt is issued in compliance with Rule 18AB and Form 10BD requirements.</div>
    </div>
      </div>
      
      <div class="receipt-number">${receiptNumber}</div>
    </div>
    
    
    
    <!-- Main Content -->
     <div>
    <div class="main-content">
      <!-- Left Column -->
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
      
      <!-- Right Column -->
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
            <span class="detail-label">Address:</span> area, 1st cross, MG Road,
          </div>
          <div class="detail-row" style="margin-left: 55px;">
            Bangalore, India
          </div>
          <div class="detail-row">
            <span class="detail-label">PAN / Aadhaar (Min 80G): UPI:</span> ${pan}
          </div>
          <div class="detail-row">
            <span class="detail-label">Amount in Words:</span> ${amountInWords}
          </div>
        </div>
        </div>
        </div>
        <!-- Amount Box -->
        <div class="amount-section">
          <span class="rupee-symbol">₹</span>
          <div class="amount-value">${amount.toLocaleString('en-IN')}</div>
        </div>
        
        <div class="payment-mode">Mode of Payment (✓ Tick above):</div>
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
          Certified that the above donation is received by trust for charitable purposes only. This donation is eligible for deduction under Section 80G of the Income Tax Act, 1961, subject to applicable limits. This receipt will be reported in Form 10BD and a certificate in Form 10BE will be issued to the owner.
        </div>
        
        <!-- Signature Section -->
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


function onPaymentFormSubmit(e) {
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();

  // Run only for the required sheet
  if (sheetName !== 'payment Submission') return;

  const row = e.range.getRow();

  // 👉 Adjust column numbers as per your sheet
  // const name = sheet.getRange(row, 1).getValue();          // Name
  const pan = sheet.getRange(row, 6).getValue();           // PAN
  const amount = sheet.getRange(row, 2).getValue();        // Amount
  const transactionId = sheet.getRange(row, 3).getValue(); // Transaction ID
   const email = sheet.getRange(row, 1).getValue();         // Email
  const registerNumber = sheet.getRange(row, 11).getValue();// Register Number

  // Generate receipt PDF
  const receiptPDF = generate80GReceipt(
    name,
    pan,
    amount,
    transactionId,
    email,
    registerNumber
  );

  // Email subject
  const subject = `${name} - ${registerNumber} - Payment Receipt`;
receiptNumber
  // Dummy HTML body (replace later)
  const htmlBody = `
    <!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
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
</head>
<body>
  <div class="container">
    
    <!-- Header -->
    <div class="header">
      <img src="https://kingsfarmequestrian.com/wp-content/uploads/2023/08/Logo2.jpg" alt="Kings Equestrian Logo">
      <h1>Kings Equestrian Foundation</h1>
      <p style="margin:8px 0 0;">Where horses don’t just carry you — they change you </p>
    </div>

    <!-- Content -->
 <div class="content">
  <div class="greeting">Dear ${donorName},</div>

  <!-- Payment Success Indicator -->
  <div style="text-align:center; margin: 25px 0;">
    <img 
      src="https://i.pinimg.com/736x/69/3c/20/693c200ad675967032f941cf76953b3e.jpg"
      alt="Payment Successful"
      width="200"
      height="150"
      style="display:block; margin:0 auto;"
    />
    <div style="
      font-size:18px;
      font-weight:600;
      color:#1f7a3f;
      margin-top:10px;
    ">
      Transaction Successful
    </div>
  </div>

  <p>
    Thank you for your generous contribution to <strong>Kings Equestrian Foundation</strong>.
    Your support helps us continue our mission toward horse welfare and equestrian development.
  </p>

  <div class="info-box">
    📎 <strong>Your Donation Receipt</strong> is attached to this email for your records.
  </div>

  <p>
    <strong>Donation Details:</strong><br>
    Registration Number: ${registrationNumber}<br>
    Receipt No: ${receiptNumber}<br>
    Amount: ₹${amount.toLocaleString('en-IN')}<br>
    Transaction Reference: ${transactionRef}
  </p>

  <p style="margin-top: 25px;">
    If you have any questions or require assistance, feel free to reach out to us a


    <!-- Footer -->
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

  // Send email
  MailApp.sendEmail({
    to: email,
    subject: subject,
    htmlBody: htmlBody,
    attachments: [receiptPDF]
  });

  Logger.log('✅ Receipt sent to: ' + email);
}

