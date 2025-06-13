const SHEET_ID = "YOUR_SPREAD_SHEET_ID";
const FOLDER_ID = "YOUR_GOOGLE_DRIVE_ID";

function processEmails() {
  Logger.log("Starting email processing");
  
  try {
    const threads = getMatchingEmails();
    Logger.log(`Found ${threads.length} email threads to process`);
    
    if (threads.length === 0) {
      Logger.log("No matching emails found");
      return;
    }
    
    const sheet = SpreadsheetApp.openById(SHEET_ID).getActiveSheet();
    const folder = DriveApp.getFolderById(FOLDER_ID);
    const label = getOrCreateLabel("Processed");
    
    threads.forEach(thread => {
      processThread(thread, sheet, folder, label);
    });
    
    Logger.log("Email processing completed successfully");
  } catch (error) {
    Logger.log(`Error in processEmails: ${error.message}`);
  }
}

function processThread(thread, sheet, folder, label) {
  try {
    const messages = thread.getMessages();
    
    messages.forEach(message => {
      processMessage(message, sheet, folder);
    });
    
    thread.markRead();
    thread.addLabel(label);
  } catch (error) {
    Logger.log(`Error processing thread: ${error.message}`);
  }
}

function processMessage(message, sheet, folder) {
  try {
    const subject = message.getSubject();
    Logger.log(`Processing message with subject: ${subject}`);
    
    if (!subject.startsWith("Viable: Trial Document")) {
      Logger.log("Skipping non-matching subject");
      return;
    }
    
    // Process standard attachments
    const attachments = message.getAttachments();
    if (attachments.length === 0) {
      Logger.log("No standard attachments found");
    } else {
      Logger.log(`Found ${attachments.length} standard attachments`);
      attachments.forEach(attachment => {
        if (attachment.getContentType() === "application/pdf") {
          processPdfAttachment(attachment, sheet, folder);
        } else {
          Logger.log(`Skipping non-PDF attachment: ${attachment.getName()}`);
        }
      });
    }
    
    // Process inline images
    const inlineImages = message.getAttachments({ includeInlineImages: true }).filter(attachment => {
      const contentType = attachment.getContentType();
      return contentType.startsWith("image/") && !attachments.includes(attachment);
    });
    
    if (inlineImages.length === 0) {
      Logger.log("No inline images found");
    } else {
      Logger.log(`Found ${inlineImages.length} inline images`);
      inlineImages.forEach(image => {
        processInlineImage(image, sheet, folder);
      });
    }
    
  } catch (error) {
    Logger.log(`Error processing message: ${error.message}`);
  }
}

function convertImageToPdf(imageBlob, folder) {
  try {
    const tempDoc = DocumentApp.create(`TempDoc_${new Date().getTime()}`);
    const docId = tempDoc.getId();
    const doc = DocumentApp.openById(docId);
    const body = doc.getBody();
    
    body.appendImage(imageBlob);
    doc.saveAndClose();
    
    const pdfBlob = DriveApp.getFileById(docId).getAs("application/pdf");
    pdfBlob.setName(`ConvertedImage_${new Date().getTime()}.pdf`);
    
    const tempPdfFile = folder.createFile(pdfBlob);
    
    DriveApp.getFileById(docId).setTrashed(true);
    
    return tempPdfFile;
  } catch (error) {
    Logger.log(`Error converting image to PDF: ${error.message}`);
    return null;
  }
}

function processInlineImage(imageAttachment, sheet, folder) {
  let tempFile, docFile;
  try {
    Logger.log(`Processing inline image: ${imageAttachment.getName()}`);
    
    const imageBlob = imageAttachment.copyBlob();
    const pdfFile = convertImageToPdf(imageBlob, folder);
    if (!pdfFile) {
      Logger.log("Failed to convert inline image to PDF");
      return;
    }
    
    const pdfBlob = pdfFile.getBlob();
    
    tempFile = DriveApp.createFile(pdfBlob);
    
    const resource = {
      title: `OCR_${new Date().getTime()}`,
      parents: [{ id: folder.getId() }]
    };
    
    docFile = Drive.Files.insert(resource, pdfBlob, {
      ocr: true,
      ocrLanguage: "en",
      convert: true
    });
    
    Utilities.sleep(5000);
    
    const doc = DocumentApp.openById(docFile.id);
    const fullText = doc.getBody().getText();
    if (!fullText) {
      Logger.log("No text extracted from inline image PDF");
      return;
    }
    
    Logger.log(`Extracted text from inline image PDF:\n${fullText}`);
    
    const extractedData = extractInvoiceData(fullText);
    
    const formattedDate = formatDate(extractedData.date);
    const savedName = `${formattedDate}_${extractedData.vendor}_${extractedData.invoiceNumber}_Rs ${extractedData.amount}.pdf`;
    
    const driveFile = folder.createFile(pdfBlob).setName(savedName);
    
    Logger.log(`Saved inline image PDF to Drive: ${savedName}`);
    
    sheet.appendRow([
      new Date(),
      extractedData.date,
      extractedData.invoiceNumber,
      extractedData.amount,
      extractedData.vendor,
      driveFile.getUrl(),
      "application/pdf"
    ]);
    
  } catch (error) {
    Logger.log(`Error processing inline image PDF: ${error.message}`);
  } finally {
    if (tempFile) {
      try {
        tempFile.setTrashed(true);
      } catch (e) {
        Logger.log(`Error deleting temp file: ${e.message}`);
      }
    }
    if (docFile) {
      try {
        DriveApp.getFileById(docFile.id).setTrashed(true);
      } catch (e) {
        Logger.log(`Error deleting doc file: ${e.message}`);
      }
    }
    if (imageAttachment) {
      try {
        imageAttachment.setTrashed(true);
      } catch (e) {
        Logger.log(`Error deleting converted image file: ${e.message}`);
      }
    }
  }
}

function processPdfAttachment(attachment, sheet, folder) {
  let tempFile, docFile;
  try {
    Logger.log(`Processing PDF: ${attachment.getName()}`);
    
    const pdfBlob = attachment.copyBlob();
    
    tempFile = DriveApp.createFile(pdfBlob);
    
    const resource = {
      title: `OCR_${new Date().getTime()}`,
      parents: [{ id: folder.getId() }]
    };
    
    docFile = Drive.Files.insert(resource, pdfBlob, {
      ocr: true,
      ocrLanguage: "en",
      convert: true
    });
    
    Utilities.sleep(5000);
    
    const doc = DocumentApp.openById(docFile.id);
    const fullText = doc.getBody().getText();
    if (!fullText) {
      Logger.log("No text extracted from PDF");
      return;
    }
    
    Logger.log(`Extracted text from PDF:\n${fullText}`);
    
    const extractedData = extractInvoiceData(fullText);
    
    const formattedDate = formatDate(extractedData.date);
    const savedName = `${formattedDate}_${extractedData.vendor}_${extractedData.invoiceNumber}_Rs ${extractedData.amount}.pdf`;
    
    const driveFile = folder.createFile(pdfBlob).setName(savedName);
    
    Logger.log(`Saved PDF to Drive: ${savedName}`);
    
    sheet.appendRow([
      new Date(),
      extractedData.date,
      extractedData.invoiceNumber,
      extractedData.amount,
      extractedData.vendor,
      driveFile.getUrl(),
      "application/pdf"
    ]);
    
  } catch (error) {
    Logger.log(`Error processing PDF attachment: ${error.message}`);
  } finally {
    if (tempFile) {
      try {
        tempFile.setTrashed(true);
      } catch (e) {
        Logger.log(`Error deleting temp file: ${e.message}`);
      }
    }
    if (docFile) {
      try {
        DriveApp.getFileById(docFile.id).setTrashed(true);
      } catch (e) {
        Logger.log(`Error deleting doc file: ${e.message}`);
      }
    }
  }
}

function extractInvoiceData(text) {
  const data = {
    date: "N/A",
    vendor: "N/A",
    invoiceNumber: "N/A",
    amount: "N/A"
  };
  
  try {
    const cleanText = text.replace(/\s+/g, " ").toLowerCase();
    
    const datePatterns = [
      /(?:date|generated on|invoice date|challan generated on|bill date|date of issue):\s*(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})/i,
      /(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*[\s,.]\s*\d{1,2}[,\s.]\s*\d{4}/i,
      /(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})/
    ];
    
    for (const pattern of datePatterns) {
      const match = cleanText.match(pattern);
      if (match) {
        data.date = match[1];
        break;
      }
    }
    
    const vendorPatterns = [
      /^([A-Z][a-zA-Z\s&]+(?:llc|inc|ltd|corp|company|private limited|restaurant))/im, // Match at the start of the document
      /(?:from|name\(legal\)):\s*([^\n]+)/i,
      /<([^\s@]+@[^\s>]+)>/,
      /([A-Z][a-zA-Z\s&]+(?:llc|inc|ltd|corp|company|private limited|restaurant))/i // Fallback
    ];
    
    for (const pattern of vendorPatterns) {
      const match = text.match(pattern);
      if (match) {
        data.vendor = match[1].trim();
        break;
      }
    }
    
    const invoicePatterns = [
      /(?:invoice|bill|invoice no\.?|cpin)\s*#?\s*:?\s*([A-Z0-9\-]+)/i,
      /(?:number|no\.?)\s*:?\s*([A-Z0-9\-]+)/i,
      /(INV[A-Z0-9\-]+)/i
    ];
    
    for (const pattern of invoicePatterns) {
      const match = cleanText.match(pattern);
      if (match) {
        data.invoiceNumber = match[1];
        break;
      }
    }
    
    const amountPatterns = [
      /(?:total|amount due|balance due|net amount|final amount|invoice amount|total amount|food total)\s*:?\s*(?:rs\.?|inr|₹)?\s*([\d,]+\.?\d{0,2})/ig,
      /(?:rs\.?|inr|₹)\s*([\d,]+\.?\d{0,2})/ig,
      /([\d,]+\.?\d{0,2})\s*(?:total|amount due|balance due)/ig
    ];
    
    let amounts = [];
    for (const pattern of amountPatterns) {
      const matches = [...cleanText.matchAll(pattern)];
      matches.forEach(match => {
        const amountStr = match[1].replace(/,/g, "");
        const amount = parseFloat(amountStr);
        if (!isNaN(amount)) {
          amounts.push(amount);
        }
      });
    }
    
    if (amounts.length > 0) {
      const maxAmount = Math.max(...amounts);
      data.amount = maxAmount.toFixed(2);
    }
    
    return data;
    
  } catch (error) {
    Logger.log(`Error in extractInvoiceData: ${error.message}`);
    return data;
  }
}

function getMatchingEmails() {
  return GmailApp.search('subject:"Viable: Trial Document" label:inbox');
}

function getOrCreateLabel(name) {
  let label = GmailApp.getUserLabelByName(name);
  if (!label) {
    label = GmailApp.createLabel(name);
  }
  return label;
}

function formatDate(dateString) {
  if (!dateString || dateString === "N/A") return "N/A";
  
  const monthNames = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];
  const monthPattern = new RegExp(`(${monthNames.join("|")})[a-z]*[\\s,.]+(\\d{1,2})[\\s,.]+(\\d{4})`, "i");
  const monthMatch = dateString.match(monthPattern);
  if (monthMatch) {
    const month = (monthNames.indexOf(monthMatch[1].toLowerCase()) + 1).toString().padStart(2, "0");
    const day = monthMatch[2].padStart(2, "0");
    const year = monthMatch[3];
    return `${day}.${month}.${year}`;
  }
  
  const dateParts = dateString.split(/[\/\-\.]/);
  if (dateParts.length === 3) {
    const day = dateParts[0].padStart(2, "0");
    const month = dateParts[1].padStart(2, "0");
    const year = dateParts[2].length === 2 ? `20${dateParts[2]}` : dateParts[2];
    return `${day}.${month}.${year}`;
  }
  
  return "N/A";
}

function createTrigger() {
  const existingTriggers = ScriptApp.getProjectTriggers();
  existingTriggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === "processEmails") {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  ScriptApp.newTrigger("processEmails")
    .timeBased()
    .everyHours(3)
    .create();
    
  Logger.log("Created new trigger to run every 3 hours");
}
