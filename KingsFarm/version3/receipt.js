// ============================================
// KINGS EQUESTRIAN - RECEIPT GENERATION & DRIVE STORAGE
// File: 4_Receipt.js
// Depends on: 1_Config.js
// ============================================

// --------------- IMAGE HELPERS ---------------

function getImageAsBase64(fileId) {
    try {
        const file     = DriveApp.getFileById(fileId);
        const blob     = file.getBlob();
        const base64   = Utilities.base64Encode(blob.getBytes());
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
        const blob     = response.getBlob();
        const base64   = Utilities.base64Encode(blob.getBytes());
        const mimeType = blob.getContentType();
        return `data:${mimeType};base64,${base64}`;
    } catch (error) {
        Logger.log('Error fetching image from URL: ' + error);
        return '';
    }
}

// --------------- NUMBER TO WORDS ---------------

function numberToWords(num) {
    const ones  = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine'];
    const teens = ['Ten', 'Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen'];
    const tens  = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety'];

    function convert(n) {
        if (n === 0)         return 'Zero';
        if (n < 10)          return ones[n];
        if (n < 20)          return teens[n - 10];
        if (n < 100)         return tens[Math.floor(n / 10)] + (n % 10 ? ' ' + ones[n % 10] : '');
        if (n < 1000)        return ones[Math.floor(n / 100)] + ' Hundred' + (n % 100 ? ' ' + convert(n % 100) : '');
        if (n < 100000)      return convert(Math.floor(n / 1000)) + ' Thousand' + (n % 1000 ? ' ' + convert(n % 1000) : '');
        if (n < 10000000)    return convert(Math.floor(n / 100000)) + ' Lakh' + (n % 100000 ? ' ' + convert(n % 100000) : '');
        return convert(Math.floor(n / 10000000)) + ' Crore' + (n % 10000000 ? ' ' + convert(n % 10000000) : '');
    }

    return convert(num).trim() + ' Rupees';
}

// --------------- 80G RECEIPT PDF GENERATION ---------------

function generate80GReceipt(riderName, pan, amount, transactionRef, receiptNumber) {
    Logger.log('Converting logo to base64...');
    const logoBase64  = getImageFromUrlAsBase64('https://kingsfarmequestrian.com/wp-content/uploads/2023/08/Logo2.jpg');
    Logger.log('Converting stamp to base64...');
    const stampBase64 = getImageAsBase64('1fQVqA1ABWCaTJs4uJVxiNqIGhl5iWugJ');
    Logger.log('Converting signature to base64...');
    const signBase64  = getImageAsBase64('1CI6H0JgysxanA0RimUwu7QwSSRospSwc');

    const htmlContent = createReceiptHTML(riderName, pan, amount, transactionRef, receiptNumber, logoBase64, stampBase64, signBase64);

    const htmlFile = DriveApp.createFile(`receipt_temp_${new Date().getTime()}.html`, htmlContent, MimeType.HTML);
    const blob     = htmlFile.getAs('application/pdf');
    blob.setName(`80G_Receipt_${riderName.replace(/\s+/g, '_')}_${receiptNumber.replace(/\//g, '_')}.pdf`);

    htmlFile.setTrashed(true);
    return blob;
}

// --------------- RECEIPT HTML TEMPLATE ---------------

function createReceiptHTML(donorName, pan, amount, transactionRef, receiptNumber, logoBase64, stampBase64, signBase64) {
    const currentDate   = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yy');
    const amountInWords = numberToWords(amount);

    return `
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
    @page { size: A4; margin: 0; }
    body { font-family: "Times New Roman", serif; margin: 0; padding: 25px; background: #fff; }
    .receipt-container { border: 2px solid #000; border-radius: 35px; padding: 25px 30px; max-width: 800px; margin: auto; position: relative; }
    .header { display: flex; align-items: flex-start; }
    .logo-section { width: 140px; text-align: center; }
    .logo-img { width: 110px; }
    .header-center { flex: 1; text-align: center; }
    .org-name { font-size: 28px; font-weight: bold; margin-bottom: 5px; }
    .registration-info { font-size: 13px; }
    .registration-subdetails { font-size: 13px; margin-top: 3px; }
    .subtext { margin-top: 10px; font-style: italic; font-weight: bold; text-decoration: underline; }
    .receipt-number { position: absolute; right: 30px; top: 15px; font-size: 16px; font-weight: bold; color: red; }
    .receipt-box { border: 2px solid #000; border-radius: 12px; text-align: center; padding: 10px; margin: 20px 0 10px; }
    .receipt-title { font-size: 20px; font-weight: bold; }
    .receipt-subtitle { font-size: 12px; }
    .date-row { text-align: right; font-size: 14px; margin-bottom: 10px; }
    .section-title { font-weight: bold; margin: 12px 0 6px; font-size: 15px; }
    .main-content { display: flex; gap: 30px; margin-top: 10px; }
    .left-column, .right-column { flex: 1; font-size: 14px; }
    .checkbox-item { margin: 5px 0; }
    .checkbox { display: inline-block; width: 13px; height: 13px; border: 1px solid #000; margin-right: 6px; vertical-align: middle; }
    .checkbox.checked { background: #000; position: relative; }
    .checkbox.checked::after { content: "✓"; color: #fff; font-size: 11px; position: absolute; left: 1px; top: -2px; }
    .detail-row { margin: 8px 0; }
    .detail-label { font-weight: bold; }
    .amount-section { border: 2px solid #000; margin: 20px 0; padding: 18px; position: relative; text-align: center; }
    .rupee-symbol { position: absolute; left: 20px; top: 50%; transform: translateY(-50%); font-size: 40px; color: goldenrod; font-weight: bold; }
    .amount-value { font-size: 34px; font-weight: bold; }
    .payment-mode { font-size: 14px; margin-top: 10px; }
    .declaration-section { margin-top: 15px; font-size: 13px; text-align: justify; }
    .signature-section { margin-top: 36px; text-align: right; }
    .org-label { font-weight: bold; margin-bottom: 5px; }
    .stamp-and-sign { position: relative; height: 120px; }
    .sign-img { width: 110px; }
    .stamp-img { width: 120px; }
</style>
</head>
<body>
<div class="receipt-container">
    <div class="receipt-number">${receiptNumber}</div>
    <div class="header">
        <div class="logo-section">
            <img src="${logoBase64}" class="logo-img" />
        </div>
        <div class="header-center">
            <div class="org-name">Kings Equestrian Foundation</div>
            <div class="registration-info">Registered u/s 80G of Income-tax Act Rg no:AAJCK7191GE20231, 1961, PAN: AAJCK7191G</div>
            <div class="registration-subdetails">K202, Tower-6, Jacaranda Block, Devarabisanahalli, Bellandur S.O, Bengaluru – 560103 Karnataka, India<br>kingsequestrianfoundation@gmail.com, kingsequestrianfoundation.com</div>
            <div class="subtext">We gratefully acknowledge your generous contribution in support of our programmes promoting education, well-being, and personal development through sport and experiential learning.</div>
        </div>
    </div>
    <div class="receipt-box">
        <div class="receipt-title">Receipt</div>
        <div class="receipt-subtitle">This receipt is issued in compliance with Rule 18AB and Form 10BD requirements</div>
    </div>
    <div class="date-row"><strong>Date:</strong> ${currentDate}</div>
    <div class="main-content">
        <div class="left-column">
            <div class="section-title">Donor Category (✓ Tick Applicable)</div>
            <div class="checkbox-item"><span class="checkbox checked"></span> Resident Indian Donor</div>
            <div class="checkbox-item"><span class="checkbox"></span> Non-Resident Indian (NRI)</div>
        </div>
        <div class="right-column">
            <div class="section-title">Donor Details</div>
            <div class="detail-row"><span class="detail-label">Name of Donor:</span> ${donorName}</div>
            <div class="detail-row"><span class="detail-label">PAN / Aadhaar:</span> ${pan}</div>
            <div class="detail-row"><span class="detail-label">Amount in Words:</span> ${amountInWords}</div>
        </div>
    </div>
    <div class="amount-section">
        <span class="rupee-symbol">₹</span>
        <div class="amount-value">${amount.toLocaleString('en-IN')}</div>
    </div>
    <div class="payment-mode">
        <strong>Mode of Payment:</strong> Cheque / DD / NEFT / RTGS / UPI (Cash not eligible u/s 80G)<br><br>
        ${transactionRef && transactionRef !== 'N/A' ? `Transaction Reference No.: <strong>${transactionRef}</strong><br><br>` : ''}
        <strong>Amount in Words:</strong> ${amountInWords}
    </div>
    <div class="declaration-section">
        Certified that the above donation is received by trust for charitable purposes only.
        This donation is eligible for deduction under Section 80G of the Income Tax Act, 1961.
        This receipt will be reported in Form 10BD and Form 10BE will be issued to the donor.
    </div>
    <div class="signature-section">
        <div class="org-label">For Kings Equestrian Foundation</div>
        <div class="stamp-and-sign">
            <img src="${signBase64}" class="sign-img" />
            <img src="${stampBase64}" class="stamp-img" />
        </div>
    </div>
</div>
</body>
</html>`;
}

// --------------- DRIVE STORAGE ---------------

function getKingsFarmFolder() {
    const folderName = 'Kings Farm Receipts';
    const year       = new Date().getFullYear();

    let mainFolder = DriveApp.getFoldersByName(folderName);
    mainFolder = mainFolder.hasNext() ? mainFolder.next() : DriveApp.createFolder(folderName);

    const yearFolders = mainFolder.getFoldersByName(year.toString());
    return yearFolders.hasNext() ? yearFolders.next() : mainFolder.createFolder(year.toString());
}

function storeReceiptInDrive(receiptBlob, riderName, receiptNumber, referenceNumber) {
    try {
        const folder    = getKingsFarmFolder();
        const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
        const fileName  = `Receipt_${receiptNumber.replace(/\//g, '-')}_${riderName.replace(/\s+/g, '_')}_${timestamp}.pdf`;

        const file = folder.createFile(receiptBlob);
        file.setName(fileName);
        file.setDescription(`Receipt for ${riderName} | Reference: ${referenceNumber} | Receipt No: ${receiptNumber}`);

        Logger.log(`Receipt saved to Drive: ${fileName}`);

        return {
            fileId:    file.getId(),
            fileUrl:   file.getUrl(),
            fileName:  fileName,
            folderId:  folder.getId(),
            folderUrl: folder.getUrl()
        };
    } catch (error) {
        Logger.log('Error storing receipt in Drive: ' + error);
        return null;
    }
}