/**
 * Work Order Generator - Google Apps Script (Updated)
 */

const SPREADSHEET_NAME = "YOUR SPREADSHEET NAME";
const CUSTOMER_DATABASE_SHEET_NAME = "YOUR DATABASE SHEET NAME";
const ORDER_DATA_SHEET_NAME = "YOUR ORDER DATASHEET NAME";
const PRODUCT_LINE_SHEET_NAME = "YOUR PRODUCT SHEET NAME";

const MASTER_SPREADSHEET_ID = "YOUR MASTER SPREADSHEET ID";
const MASTER_CUSTOMER_SHEET_NAME = "YOUR CUSTOMER MASTER SATASHEET NAME";
const PDF_FOLDER_ID = "YOUR PDF FOLDER ID";

function setupDriveAccess() {
  try {
    const testFolder = DriveApp.getRootFolder();
    Logger.log("✓ Drive access authorized");
    try {
      const pdfFolder = DriveApp.getFolderById(PDF_FOLDER_ID);
      Logger.log("✓ PDF folder accessible: " + pdfFolder.getName());
      return { success: true, message: "Drive access configured. PDF folder: " + pdfFolder.getName() };
    } catch (e) {
      const newFolder = DriveApp.createFolder("Work Order PDFs");
      return { success: true, message: "Created folder. Update PDF_FOLDER_ID to: " + newFolder.getId(), newFolderId: newFolder.getId() };
    }
  } catch (e) {
    return { success: false, message: "Drive auth needed: " + e.message };
  }
}

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('Work Order Generator')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

function getCustomerData() {
  try {
    // Read from: Client Database_ICPL_Main → Customers
    // A(1)=SL, B(2)=Company Name, C(3)=Address-1, D(4)=Address-2, E(5)=Address-3,
    // F(6)=GSTN, G(7)=Contact Person, H(8)=Contact No., I(9)=E-Mail, J(10)=Email CC
    const ss    = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
    const sheet = ss.getSheetByName(MASTER_CUSTOMER_SHEET_NAME);
    if (!sheet) return [];
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    const values = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
    return values
      .map(r => {
        // Combine Address-1, Address-2, Address-3 — skip empty parts
        const addrParts = [
          String(r[2]).trim(),  // C = Address-1
          String(r[3]).trim(),  // D = Address-2
          String(r[4]).trim()   // E = Address-3
        ].filter(p => p && p !== '');
        return {
          name:        String(r[1]).trim(),          // B = Company Name
          address:     addrParts.join(', '),          // C+D+E combined
          gst:         String(r[5]).trim(),           // F = GSTN
          contactName: String(r[6]).trim(),           // G = Contact Person
          contactNo:   String(r[7]).trim(),           // H = Contact No.
          email:       String(r[8]).trim(),           // I = E-Mail
          emailCC:     String(r[9]).trim()            // J = Email CC
        };
      })
      .filter(c => c.name);
  } catch (e) {
    Logger.log(e);
    return [];
  }
}

/**
 * GET PRODUCT LIST DATA
 * Fetches all product lists from "Work Order Maker" → "Product List" tab
 * Column mapping (from row 3):
 *   A = Stud Link Chain Grade,  B = Stud Link Chain Short Desc
 *   D = Anchor Grade,           E = Anchor Short Desc
 *   G = Chain/Slings Grade,     H = Chain/Slings Short Desc
 *   J = Elevator Conveyor Short Desc (no grade)
 *   L = Miscellaneous Short Desc (no grade)
 */
function getProductListData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Product List");
    if (!sheet) {
      Logger.log("Product List sheet not found");
      return { studLink: {grades:[], descs:[]}, anchor: {grades:[], descs:[]}, chainSlings: {grades:[], descs:[]}, elevator: {descs:[]}, misc: {descs:[]} };
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 3) return { studLink: {grades:[], descs:[]}, anchor: {grades:[], descs:[]}, chainSlings: {grades:[], descs:[]}, elevator: {descs:[]}, misc: {descs:[]} };

    // Read all 12 columns from row 3 downward (A through L)
    const numRows = lastRow - 2;
    const data = sheet.getRange(3, 1, numRows, 12).getValues();

    function col(data, colIndex) {
      return data.map(r => String(r[colIndex] || '').trim()).filter(Boolean);
    }

    return {
      studLink:   { grades: col(data, 0), descs: col(data, 1) },  // A, B
      anchor:     { grades: col(data, 3), descs: col(data, 4) },  // D, E
      chainSlings:{ grades: col(data, 6), descs: col(data, 7) },  // G, H
      elevator:   { descs: col(data, 9) },                         // J
      misc:       { descs: col(data, 11) }                         // L
    };
  } catch (e) {
    Logger.log("Error fetching product list: " + e);
    return { studLink: {grades:[], descs:[]}, anchor: {grades:[], descs:[]}, chainSlings: {grades:[], descs:[]}, elevator: {descs:[]}, misc: {descs:[]} };
  }
}

function generateWorkOrderNumber() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const orderSheet = ss.getSheetByName(ORDER_DATA_SHEET_NAME);
    if (!orderSheet) throw new Error("Order Data sheet not found.");
    const now = new Date();
    const year = now.getFullYear().toString().slice(-2);
    const month = (now.getMonth() + 1).toString().padStart(2, '0');
    const prefix = year + month;
    const lastRow = orderSheet.getLastRow();
    if (lastRow < 2) return prefix + "01";
    const workOrderNumbers = orderSheet.getRange(2, 6, lastRow - 1, 1).getValues();
    let maxSequence = 0;
    workOrderNumbers.forEach(row => {
      const woNumber = String(row[0]).trim();
      if (woNumber.startsWith(prefix) && woNumber.length === 6) {
        const sequence = parseInt(woNumber.slice(-2), 10);
        if (!isNaN(sequence) && sequence > maxSequence) maxSequence = sequence;
      }
    });
    return prefix + (maxSequence + 1).toString().padStart(2, '0');
  } catch (e) {
    Logger.log("Error generating WO number: " + e);
    const now = new Date();
    return Utilities.formatDate(now, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'yyMMddHHmmss');
  }
}

function getNextWorkOrderNumber() { return generateWorkOrderNumber(); }

function getAllWorkOrderNumbers() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const orderSheet = ss.getSheetByName(ORDER_DATA_SHEET_NAME);
    if (!orderSheet) return [];
    const lastRow = orderSheet.getLastRow();
    if (lastRow < 2) return [];
    const data = orderSheet.getRange(2, 1, lastRow - 1, 6).getValues();
    return data.filter(row => row[5]).map(row => ({ orderId: String(row[0]).trim(), workOrderNo: String(row[5]).trim() })).reverse();
  } catch (e) {
    Logger.log("Error getting WO numbers: " + e);
    return [];
  }
}

/**
 * GET WORK ORDER DETAILS
 * Column mapping (1-indexed):
 * 1=Order ID, 2=Timestamp, 3=Customer Name, 4=GST, 5=Address,
 * 6=WO No, 7=WO Date, 8=PO No, 9=PO Date, 10=LOI No, 11=LOI Date,
 * 12=Delivery Date, 13=Delivery Terms, 14=Delivery Terms Extra,
 * 15=Nominated Transporter, 16=Transporter Scope,
 * 17=Consignee Company, 18=Consignee Address, 19=Consignee GSTIN,
 * 20=Consignee Contact, 21=Consignee Phone,
 * 22=Payment Terms, 23=Advance %, 24=Guarantee Period, 25=Guarantee Clause,
 * 26=CD Applicable, 27=CD Value, 28=CD Per Week, 29=CD Per Month,
 * 30=PDI, 31=PDI Note, 32=Certifications,
 * 33=Docs Enclosed, 34=Approval Required, 35=Sample Qty,
 * 36=Docs With Material,
 * 37=Customer Contact Name, 38=Customer Contact No, 39=Customer Email, 40=Customer Email CC,
 * 41=BG Security %, 42=BG Security Date, 43=BG Performance %, 44=BG Performance Date,
 * 45=Additional Note
 */
function getWorkOrderDetails(workOrderNo) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const orderSheet = ss.getSheetByName(ORDER_DATA_SHEET_NAME);
    const lineSheet = ss.getSheetByName(PRODUCT_LINE_SHEET_NAME);
    if (!orderSheet || !lineSheet) return { success: false, message: "Required sheets not found" };

    const lastRow = orderSheet.getLastRow();
    if (lastRow < 2) return { success: false, message: "No orders found" };

    const orderData = orderSheet.getRange(2, 1, lastRow - 1, 45).getValues();
    let orderRow = null, orderRowIndex = -1;
    for (let i = 0; i < orderData.length; i++) {
      if (String(orderData[i][5]).trim() === workOrderNo) {
        orderRow = orderData[i];
        orderRowIndex = i + 2;
        break;
      }
    }
    if (!orderRow) return { success: false, message: "Work order not found" };

    const orderId = String(orderRow[0]).trim();

    const lineLastRow = lineSheet.getLastRow();
    let products = [];
    if (lineLastRow >= 2) {
      const lineData = lineSheet.getRange(2, 1, lineLastRow - 1, 10).getValues();
      products = lineData
        .filter(row => String(row[0]).trim() === orderId)
        .map(row => ({
          category: String(row[2]).trim(),
          size: String(row[3]).trim(),
          grade: String(row[4]).trim(),
          short_desc: String(row[5]).trim(),
          long_desc: String(row[6]).trim(),
          unit: String(row[7]).trim(),
          qty: String(row[8]).trim(),
          rate: String(row[9]).trim()
        }));
    }

    function fmtDate(v) {
      if (!v) return '';
      try {
        const d = new Date(v);
        if (isNaN(d.getTime())) return '';
        return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      } catch(e) { return ''; }
    }

    function parseArr(v) {
      if (!v) return [];
      return String(v).split(',').map(s => s.trim()).filter(Boolean);
    }

    return {
      success: true,
      orderId: orderId,
      orderData: {
        customerName:         String(orderRow[2]).trim(),
        customerGst:          String(orderRow[3]).trim(),
        customerAddress:      String(orderRow[4]).trim(),
        workOrderNo:          String(orderRow[5]).trim(),
        workOrderDate:        fmtDate(orderRow[6]),
        poNo:                 String(orderRow[7]).trim(),
        poDate:               fmtDate(orderRow[8]),
        loiNo:                String(orderRow[9]).trim(),
        loiDate:              fmtDate(orderRow[10]),
        deliveryDate:         fmtDate(orderRow[11]),
        deliveryTerms:        String(orderRow[12]).trim(),
        deliveryTermsExtra:   String(orderRow[13]).trim(),
        nominatedTransporter: String(orderRow[14]).trim(),
        transporterScope:     String(orderRow[15]).trim(),
        consigneeCompany:     String(orderRow[16]).trim(),
        consigneeAddr:        String(orderRow[17]).trim(),
        consigneeGstin:       String(orderRow[18]).trim(),
        consigneeContact:     String(orderRow[19]).trim(),
        consigneePhone:       String(orderRow[20]).trim(),
        paymentTerms:         String(orderRow[21]).trim(),
        advancePercent:       String(orderRow[22]).trim(),
        guaranteePeriod:      String(orderRow[23]).trim(),
        guaranteeClause:      String(orderRow[24]).trim(),
        cdApplicable:         String(orderRow[25]).trim() === 'TRUE' || String(orderRow[25]).trim() === 'true',
        cdValue:              String(orderRow[26]).trim(),
        cdPeriodWeek:         String(orderRow[27]).trim() === 'TRUE' || String(orderRow[27]).trim() === 'true',
        cdPeriodMonth:        String(orderRow[28]).trim() === 'TRUE' || String(orderRow[28]).trim() === 'true',
        pdi:                  String(orderRow[29]).trim(),
        pdiNote:              String(orderRow[30]).trim(),
        certifications:       parseArr(orderRow[31]),
        docsEnclosed:         parseArr(orderRow[32]),
        approvalReq:          parseArr(orderRow[33]),
        sampleQty:            String(orderRow[34]).trim(),
        docsWithMaterial:     parseArr(orderRow[35]),
        customerContactName:  String(orderRow[36] || '').trim(),
        customerContactNo:    String(orderRow[37] || '').trim(),
        customerEmail:        String(orderRow[38] || '').trim(),
        customerEmailCC:      String(orderRow[39] || '').trim(),
        bgSecurityPct:        String(orderRow[40] || '').trim(),
        bgSecurityDate:       fmtDate(orderRow[41]),
        bgPerformancePct:     String(orderRow[42] || '').trim(),
        bgPerformanceDate:    fmtDate(orderRow[43]),
        additionalNote:       String(orderRow[44] || '').trim()
      },
      products: products
    };
  } catch (e) {
    Logger.log("Error getting WO details: " + e);
    return { success: false, message: "Error: " + e.message };
  }
}

function deleteWorkOrder(orderId, workOrderNo) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const orderSheet = ss.getSheetByName(ORDER_DATA_SHEET_NAME);
    const lineSheet = ss.getSheetByName(PRODUCT_LINE_SHEET_NAME);
    if (!orderSheet || !lineSheet) return { success: false, message: "Required sheets not found" };

    const orderLastRow = orderSheet.getLastRow();
    if (orderLastRow >= 2) {
      const orderData = orderSheet.getRange(2, 1, orderLastRow - 1, 1).getValues();
      for (let i = orderData.length - 1; i >= 0; i--) {
        if (String(orderData[i][0]).trim() === orderId) { orderSheet.deleteRow(i + 2); break; }
      }
    }

    const lineLastRow = lineSheet.getLastRow();
    if (lineLastRow >= 2) {
      const lineData = lineSheet.getRange(2, 1, lineLastRow - 1, 1).getValues();
      for (let i = lineData.length - 1; i >= 0; i--) {
        if (String(lineData[i][0]).trim() === orderId) lineSheet.deleteRow(i + 2);
      }
    }
    return { success: true, message: "Work order deleted successfully" };
  } catch (e) {
    Logger.log("Error deleting WO: " + e);
    return { success: false, message: "Error: " + e.message };
  }
}

function addNewCustomerToDatabase(customerData) {
  try {
    const ss    = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
    const sheet = ss.getSheetByName(MASTER_CUSTOMER_SHEET_NAME);
    if (!sheet) return { success: false, message: "Customers sheet not found in Client Database_ICPL_Main." };

    // Column layout:
    // A(1)=SL, B(2)=Company Name, C(3)=Address-1, D(4)=Address-2, E(5)=Address-3,
    // F(6)=GSTN, G(7)=Contact Person, H(8)=Contact No., I(9)=E-Mail, J(10)=Email CC

    const lastRow    = sheet.getLastRow();
    const newRow     = lastRow + 1;
    const nextSL     = lastRow; // Row 2 = SL 1, Row 3 = SL 2, etc.

    sheet.getRange(newRow, 1).setValue(nextSL);                         // A = SL
    sheet.getRange(newRow, 2).setValue(customerData.name);              // B = Company Name
    sheet.getRange(newRow, 3).setValue(customerData.address);           // C = Address-1
    // D and E (Address-2, Address-3) left blank
    sheet.getRange(newRow, 6).setValue(customerData.gst);               // F = GSTN
    sheet.getRange(newRow, 7).setValue(customerData.contactName || ''); // G = Contact Person
    sheet.getRange(newRow, 8).setValue(customerData.contactNo   || ''); // H = Contact No.
    sheet.getRange(newRow, 9).setValue(customerData.email       || ''); // I = E-Mail
    sheet.getRange(newRow, 10).setValue(customerData.emailCC    || ''); // J = Email CC

    // Sort all data rows (Row 2 onwards) by Column B (Company Name) alphabetically
    const updatedLastRow = sheet.getLastRow();
    if (updatedLastRow > 2) {
      sheet.getRange(2, 1, updatedLastRow - 1, 10).sort({ column: 2, ascending: true });
    }

    // Re-number SL column (A) sequentially after sort
    for (let r = 2; r <= updatedLastRow; r++) {
      sheet.getRange(r, 1).setValue(r - 1);
    }

    return { success: true, message: `Customer "${customerData.name}" added and list sorted alphabetically.` };
  } catch (error) {
    Logger.log(error);
    return { success: false, message: "Error adding customer: " + error.message };
  }
}

/**
 * PROCESS WORK ORDER FORM
 */
function processForm(formObject) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const timestamp = new Date();

    const isRevision = formObject.isRevision === 'true';
    const oldOrderId = formObject.oldOrderId || '';
    const oldWorkOrderNo = formObject.oldWorkOrderNo || '';

    let orderId, workOrderNo;
    if (isRevision && oldOrderId && oldWorkOrderNo) {
      const deleteResult = deleteWorkOrder(oldOrderId, oldWorkOrderNo);
      if (!deleteResult.success) return { success: false, message: "Failed to delete old WO: " + deleteResult.message };
      orderId = oldOrderId;
      workOrderNo = oldWorkOrderNo;
    } else {
      orderId = 'WO-' + Utilities.formatDate(timestamp, ss.getSpreadsheetTimeZone(), 'yyyyMMddHHmmss');
      workOrderNo = generateWorkOrderNumber();
    }

    const orderSheet = ss.getSheetByName(ORDER_DATA_SHEET_NAME);
    if (!orderSheet) throw new Error("Order Data sheet not found.");

    const docsEnclosed = Array.isArray(formObject.docsEnclosed)
      ? formObject.docsEnclosed.join(', ')
      : (formObject.docsEnclosed || '');

    const approvalReq = Array.isArray(formObject.approvalReq)
      ? formObject.approvalReq.join(', ')
      : (formObject.approvalReq || '');

    const certifications = Array.isArray(formObject.certifications)
      ? formObject.certifications.join(', ')
      : (formObject.certifications || '');

    const docsWithMaterial = Array.isArray(formObject.docsWithMaterial)
      ? formObject.docsWithMaterial.join(', ')
      : (formObject.docsWithMaterial || '');

    // Add header row if sheet is empty
    if (orderSheet.getLastRow() === 0) {
      orderSheet.appendRow([
        "Order ID", "Timestamp", "Customer Name", "GST Number", "Customer Address",
        "Work Order No.", "Work Order Date",
        "PO No.", "PO Date", "LOI No.", "LOI Date",
        "Delivery Date", "Terms of Delivery", "Delivery Terms Extra",
        "Nominated Transporter", "Transporter Scope",
        "Consignee Company", "Consignee Address", "Consignee GSTIN",
        "Consignee Contact", "Consignee Phone",
        "Payment Terms", "Advance %", "Guarantee Period", "Guarantee Clause",
        "CD Applicable", "CD Value", "CD Per Week", "CD Per Month",
        "PDI", "PDI Note", "Certifications",
        "Docs Enclosed", "Approval Required", "Sample Qty",
        "Docs With Material",
        "Customer Contact Name", "Customer Contact No", "Customer Email", "Customer Email CC",
        "BG Security %", "BG Security Date", "BG Performance %", "BG Performance Date",
        "Additional Note"
      ]);
    }

    orderSheet.appendRow([
      orderId, timestamp,
      formObject.customerName       || '',
      formObject.customerGst        || '',
      formObject.customerAddress    || '',
      workOrderNo,
      formObject.workOrderDate      || '',
      formObject.poNo               || '',
      formObject.poDate             || '',
      formObject.loiNo              || '',
      formObject.loiDate            || '',
      formObject.deliveryDate       || '',
      formObject.deliveryTerms      || '',
      formObject.deliveryTermsExtra || '',
      formObject.nominatedTransporter || '',
      formObject.transporterScope   || '',
      formObject.consigneeCompany   || '',
      formObject.consigneeAddr      || '',
      formObject.consigneeGstin     || '',
      formObject.consigneeContact   || '',
      formObject.consigneePhone     || '',
      formObject.paymentTerms       || '',
      formObject.advancePercent     || '',
      formObject.guaranteePeriod    || '',
      formObject.guaranteeClause    || '',
      formObject.cdApplicable       || '',
      formObject.cdValue            || '',
      formObject.cdPeriod === 'Per Week'  || String(formObject.cdPeriod).includes('Per Week')  ? 'TRUE' : '',
      formObject.cdPeriod === 'Per Month' || String(formObject.cdPeriod).includes('Per Month') ? 'TRUE' : '',
      formObject.pdi                || '',
      formObject.pdiNote            || '',
      certifications,
      docsEnclosed,
      approvalReq,
      formObject.sampleQty          || '',
      docsWithMaterial,
      formObject.customerContactName || '',
      formObject.customerContactNo   || '',
      formObject.customerEmail       || '',
      formObject.customerEmailCC     || '',
      formObject.bgSecurityPct       || '',
      formObject.bgSecurityDate      || '',
      formObject.bgPerformancePct    || '',
      formObject.bgPerformanceDate   || '',
      formObject.additionalNote      || ''
    ]);

    // Product Line Items
    const lineSheet = ss.getSheetByName(PRODUCT_LINE_SHEET_NAME);
    if (!lineSheet) throw new Error("Product Line Items sheet not found.");
    if (lineSheet.getLastRow() === 0) {
      lineSheet.appendRow(["Order ID", "Product Index", "Category", "Size", "Grade", "Short Description", "Long Description", "Unit", "Quantity", "Rate"]);
    }

    const products = JSON.parse(formObject.allProductsJSON || "[]");
    const rows = products.map((p, i) => [orderId, i + 1, p.category, p.size, p.grade, p.short_desc, p.long_desc, p.unit, p.qty, p.rate]);
    if (rows.length) {
      lineSheet.getRange(lineSheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    }

    return {
      success: true,
      message: isRevision ? "Work Order Revised Successfully!" : "Work Order Generated Successfully!",
      orderId: orderId,
      workOrderNo: workOrderNo,
      isRevision: isRevision
    };
  } catch (e) {
    Logger.log(e);
    return { success: false, message: "Error: " + e.message };
  }
}

/**
 * GENERATE DRAFT PDF — Print Draft
 * Creates an Office Copy PDF from live form data.
 * Does NOT save to Drive or sheet — returns PDF as base64 so the
 * browser can open it directly as a data URL. Form stays open.
 */
function generateDraftPDF(formData) {
  try {
    function parseArr(v) {
      if (!v) return [];
      return String(v).split(',').map(s => s.trim()).filter(Boolean);
    }
    function fmtDateStr(v) {
      if (!v) return '';
      try {
        const d = new Date(v);
        if (isNaN(d.getTime())) return '';
        return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd-MMM-yyyy');
      } catch(e) { return v || ''; }
    }

    const orderData = {
      customerName:         formData.customerName         || '',
      customerGst:          formData.customerGst          || '',
      customerAddress:      formData.customerAddress      || '',
      customerContactName:  formData.customerContactName  || '',
      customerContactNo:    formData.customerContactNo    || '',
      customerEmail:        formData.customerEmail        || '',
      customerEmailCC:      formData.customerEmailCC      || '',
      workOrderNo:          formData.oldWorkOrderNo       || '(DRAFT)',
      workOrderDate:        fmtDateStr(formData.workOrderDate),
      poNo:                 formData.poNo                 || '',
      poDate:               fmtDateStr(formData.poDate),
      loiNo:                formData.loiNo                || '',
      loiDate:              fmtDateStr(formData.loiDate),
      deliveryDate:         fmtDateStr(formData.deliveryDate),
      deliveryTerms:        formData.deliveryTerms        || '',
      deliveryTermsExtra:   formData.deliveryTermsExtra   || '',
      nominatedTransporter: formData.nominatedTransporter || '',
      transporterScope:     formData.transporterScope     || '',
      consigneeCompany:     formData.consigneeCompany     || '',
      consigneeAddr:        formData.consigneeAddr        || '',
      consigneeGstin:       formData.consigneeGstin       || '',
      consigneeContact:     formData.consigneeContact     || '',
      consigneePhone:       formData.consigneePhone       || '',
      paymentTerms:         formData.paymentTerms         || '',
      advancePercent:       formData.advancePercent       || '',
      guaranteePeriod:      formData.guaranteePeriod      || '',
      guaranteeClause:      formData.guaranteeClause      || '',
      cdApplicable:         formData.cdApplicable === 'Applicable' || formData.cdApplicable === 'TRUE',
      cdValue:              formData.cdValue              || '',
      cdPeriodWeek:         String(formData.cdPeriod || '').includes('Per Week'),
      cdPeriodMonth:        String(formData.cdPeriod || '').includes('Per Month'),
      bgSecurityPct:        formData.bgSecurityPct        || '',
      bgSecurityDate:       fmtDateStr(formData.bgSecurityDate),
      bgPerformancePct:     formData.bgPerformancePct     || '',
      bgPerformanceDate:    fmtDateStr(formData.bgPerformanceDate),
      pdi:                  formData.pdi                  || '',
      pdiNote:              formData.pdiNote              || '',
      certifications:       parseArr(formData.certifications),
      docsEnclosed:         parseArr(formData.docsEnclosed),
      approvalReq:          parseArr(formData.approvalReq),
      sampleQty:            formData.sampleQty            || '',
      docsWithMaterial:     parseArr(formData.docsWithMaterial),
      additionalNote:       formData.additionalNote       || ''
    };

    let products = [];
    try { products = JSON.parse(formData.allProductsJSON || '[]'); } catch(e) {}

    // Generate Office Copy HTML with DRAFT watermark
    const htmlContent = createPDFHtmlContent(orderData, products, true);
    const draftHtml = htmlContent.replace(
      '<body>',
      `<body><div style="position:fixed;top:45%;left:50%;transform:translate(-50%,-50%) rotate(-35deg);font-size:90px;font-weight:900;color:rgba(200,30,30,0.12);pointer-events:none;white-space:nowrap;z-index:9999;font-family:Arial;">DRAFT</div>`
    );

    // Convert to PDF and return as base64 — no Drive access required
    const pdfBytes = Utilities.newBlob(draftHtml, 'text/html', 'draft.html')
      .getAs('application/pdf')
      .getBytes();

    const base64Pdf = Utilities.base64Encode(pdfBytes);

    return {
      success: true,
      base64Pdf: base64Pdf
    };
  } catch(e) {
    Logger.log("Error generating draft PDF: " + e);
    return { success: false, message: "Draft PDF error: " + e.message };
  }
}

/**
 * GENERATE WORK ORDER PDF — Print Approved
 * Combines Office Copy (with Rate) + Factory Copy (without Rate) into one PDF.
 * Saved to Drive as: WO_260301_CustomerName.pdf
 */
function generateWorkOrderPDF(orderId, workOrderNo, copyType) {
  if (copyType === undefined) copyType = 'approved';
  try {
    const details = getWorkOrderDetails(workOrderNo);
    if (!details.success) return { success: false, message: "Could not retrieve WO details: " + details.message };

    // Build combined HTML: Office Copy + Factory Copy
    const combinedHtml = createCombinedPDFHtmlContent(details.orderData, details.products);

    const pdfBlob = Utilities.newBlob(combinedHtml, 'text/html', 'workorder.html')
      .getAs('application/pdf')
      .setName('WO_' + workOrderNo + '_Approved_Copy.pdf');

    let pdfFile, folderName = "PDF Storage";
    try {
      const folder = DriveApp.getFolderById(PDF_FOLDER_ID);
      pdfFile = folder.createFile(pdfBlob);
      folderName = folder.getName();
    } catch (fe) {
      const folders = DriveApp.getFoldersByName("Work Order PDFs");
      if (folders.hasNext()) { pdfFile = folders.next().createFile(pdfBlob); folderName = "Work Order PDFs"; }
      else { pdfFile = DriveApp.createFile(pdfBlob); folderName = "My Drive (Root)"; }
    }

    // Also encode as base64 so the browser can open it directly
    const base64Pdf = Utilities.base64Encode(pdfBlob.getBytes());

    return {
      success: true,
      message: "PDF saved to " + folderName,
      pdfUrl: pdfFile.getUrl(),
      pdfName: pdfFile.getName(),
      base64Pdf: base64Pdf
    };
  } catch (e) {
    Logger.log("Error generating PDF: " + e);
    return { success: false, message: "Error generating PDF: " + e.message };
  }
}

/**
 * CREATE COMBINED PDF HTML CONTENT
 * Renders Factory Copy first, then a clear separator page, then HO Copy.
 * A single @page break-before forces each section onto its own page(s).
 */
function createCombinedPDFHtmlContent(orderData, products) {
  const officeCopyBody  = createPDFHtmlContent(orderData, products, true);   // Office Copy = full with rate
  const factoryCopyBody = createFactoryPDFHtmlContent(orderData, products);  // Factory Copy = selective, no rate

  // Strip outer html/head/body tags
  function extractBody(html) {
    const m = html.match(/<body[^>]*>([\s\S]*)<\/body>/i);
    return m ? m[1] : html;
  }
  function extractStyle(html) {
    const m = html.match(/<style[^>]*>([\s\S]*?)<\/style>/i);
    return m ? m[1] : '';
  }

  const sharedStyle    = extractStyle(officeCopyBody);
  const officeContent  = extractBody(officeCopyBody);
  const factoryContent = extractBody(factoryCopyBody);

  return `<!DOCTYPE html><html><head><meta charset="UTF-8">
  <style>
    ${sharedStyle}
    .office-section  { }
    .factory-section { page-break-before: always; break-before: page; }
  </style>
  </head><body>

  <!-- ══════════════ OFFICE COPY ══════════════ -->
  <div class="office-section">
    ${officeContent}
  </div>

  <!-- ══════════════ FACTORY COPY ══════════════ -->
  <div class="factory-section">
    ${factoryContent}
  </div>

  </body></html>`;
}

/**
 * CREATE PDF HTML CONTENT — HO Copy (full details, with Rate)
 */
function createPDFHtmlContent(orderData, products, includeRate) {
  if (includeRate === undefined) includeRate = true; // HO Copy always shows rate
  function fmtDate(dateStr) {
    if (!dateStr) return '';
    try {
      const d = new Date(dateStr);
      return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd-MMM-yyyy');
    } catch(e) { return dateStr; }
  }

  function row(label, value) {
    return `<div class="info-row"><div class="info-label">${label}</div><div class="info-value">${value || '—'}</div></div>`;
  }

  // Delivery terms combined
  const deliveryTermsCombined = [orderData.deliveryTerms, orderData.deliveryTermsExtra].filter(Boolean).join(' ');

  // CD Clause text
  let cdClauseText = orderData.cdApplicable ? 'Applicable' : 'Not Applicable';
  if (orderData.cdApplicable && orderData.cdValue) {
    const period = orderData.cdPeriodWeek ? 'Per Week' : (orderData.cdPeriodMonth ? 'Per Month' : '');
    cdClauseText = `Applicable — ${orderData.cdValue} ${period}`.trim();
  }

  let productRows = '';
  products.forEach((p, i) => {
    productRows += `<tr>
      <td style="text-align:center;">${i + 1}</td>
      <td>${p.category}</td>
      <td>${p.size}</td>
      <td>${p.grade}</td>
      <td>${p.short_desc}</td>
      <td>${p.long_desc}</td>
      <td style="text-align:center;">${p.unit}</td>
      <td style="text-align:right;">${p.qty}</td>
      ${includeRate ? `<td style="text-align:right;">${p.rate}</td>` : ''}
    </tr>`;
  });

  const rateHeader = includeRate ? '<th>Rate</th>' : '';

  return `<!DOCTYPE html><html><head><meta charset="UTF-8">
  <style>
    @page { size: A4; margin: 8mm 10mm; }
    * { box-sizing: border-box; }
    body {
      font-family: Arial, sans-serif;
      margin: 0;
      padding: 0;
      color: #222;
      font-size: 9.5px;
      line-height: 1.3;
    }

    /* ── HEADER ── */
    .header {
      text-align: center;
      border-bottom: 2.5px solid #0052CC;
      padding-bottom: 5px;
      margin-bottom: 6px;
    }
    .company-name { font-size: 16px; font-weight: bold; color: #0052CC; letter-spacing: 0.5px; }
    .document-title { font-size: 13px; font-weight: bold; margin-top: 2px; }
    .wo-meta { margin-top: 3px; font-size: 9.5px; }

    /* ── TWO-COLUMN LAYOUT for info sections ── */
    .two-col-layout {
      display: table;
      width: 100%;
      border-collapse: separate;
      border-spacing: 4px 0;
      margin-bottom: 5px;
    }
    .col-left, .col-right {
      display: table-cell;
      width: 50%;
      vertical-align: top;
    }

    /* ── SECTION BLOCK ── */
    .section { margin-bottom: 5px; }
    .section-full { margin-bottom: 5px; }
    .section-title {
      font-size: 9px;
      font-weight: bold;
      background: #0052CC;
      color: white;
      padding: 3px 6px;
      margin-bottom: 2px;
      text-transform: uppercase;
      letter-spacing: 0.3px;
    }

    /* ── INFO ROWS ── */
    .info-grid { display: table; width: 100%; border-collapse: collapse; }
    .info-row { display: table-row; }
    .info-label {
      display: table-cell;
      width: 42%;
      font-weight: bold;
      padding: 2px 5px;
      background: #f0f4ff;
      border: 0.5px solid #ccc;
      font-size: 8.5px;
      white-space: nowrap;
    }
    .info-value {
      display: table-cell;
      padding: 2px 5px;
      border: 0.5px solid #ccc;
      font-size: 8.5px;
      word-break: break-word;
    }

    /* ── PRODUCT TABLE ── */
    .product-table { width: 100%; border-collapse: collapse; margin-top: 3px; }
    .product-table th {
      background: #0052CC;
      color: white;
      padding: 3px 4px;
      text-align: left;
      border: 0.5px solid #aac;
      font-size: 8.5px;
      white-space: nowrap;
    }
    .product-table td {
      padding: 3px 4px;
      border: 0.5px solid #ccc;
      font-size: 8.5px;
      vertical-align: top;
    }
    .product-table tr:nth-child(even) td { background: #f9f9ff; }

    /* ── FOOTER ── */
    .footer {
      margin-top: 6px;
      padding-top: 4px;
      border-top: 1px solid #ccc;
      text-align: center;
      font-size: 8px;
      color: #888;
    }

    @media print {
      body { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
    }
  </style></head><body>

  <!-- HEADER -->
  <div class="header">
    <div class="company-name">INDIAN CHAIN PRIVATE LIMITED</div>
    <div class="document-title">WORK ORDER</div>
    <div><span style="display:inline-block;background:#0052CC;color:white;font-size:9px;font-weight:bold;padding:2px 8px;border-radius:3px;margin-top:3px;letter-spacing:0.5px;">OFFICE COPY</span></div>
    <div class="wo-meta">
      <strong>Work Order No:</strong> ${orderData.workOrderNo}
      &nbsp;&nbsp;|&nbsp;&nbsp;
      <strong>Date:</strong> ${fmtDate(orderData.workOrderDate)}
    </div>
  </div>

  <!-- ROW 1: Customer Info + Order Details -->
  <div class="two-col-layout">
    <div class="col-left">
      <div class="section">
        <div class="section-title">Customer Information</div>
        <div class="info-grid">
          ${row('Customer Name', orderData.customerName)}
          ${row('GST Number', orderData.customerGst)}
          ${row('Address', orderData.customerAddress)}
          ${orderData.customerContactName ? row('Contact Name', orderData.customerContactName) : ''}
          ${orderData.customerContactNo   ? row('Contact No.',  orderData.customerContactNo)   : ''}
          ${orderData.customerEmail       ? row('Email',        orderData.customerEmail)        : ''}
          ${orderData.customerEmailCC     ? row('Email CC',     orderData.customerEmailCC)      : ''}
        </div>
      </div>
    </div>
    <div class="col-right">
      <div class="section">
        <div class="section-title">Order Details</div>
        <div class="info-grid">
          ${row('PO No.', orderData.poNo)}
          ${row('PO Date', fmtDate(orderData.poDate))}
          ${row('LOI No.', orderData.loiNo)}
          ${row('LOI Date', fmtDate(orderData.loiDate))}
        </div>
      </div>
    </div>
  </div>

  <!-- ROW 2: Delivery Info + Consignee Info -->
  <div class="two-col-layout">
    <div class="col-left">
      <div class="section">
        <div class="section-title">Delivery Information</div>
        <div class="info-grid">
          ${row('Delivery Date', fmtDate(orderData.deliveryDate))}
          ${row('Terms of Delivery', deliveryTermsCombined)}
          ${row('Nominated Transporter', orderData.nominatedTransporter)}
          ${row('Transporter Scope', orderData.transporterScope)}
        </div>
      </div>
    </div>
    <div class="col-right">
      <div class="section">
        <div class="section-title">Consignee Information</div>
        <div class="info-grid">
          ${row('Company Name', orderData.consigneeCompany)}
          ${row('Address', orderData.consigneeAddr)}
          ${row('GSTIN', orderData.consigneeGstin)}
          ${row('Contact Person', orderData.consigneeContact)}
          ${row('Contact Number', orderData.consigneePhone)}
        </div>
      </div>
    </div>
  </div>

  <!-- ROW 3: Terms & Conditions + Inspection -->
  <div class="two-col-layout">
    <div class="col-left">
      <div class="section">
        <div class="section-title">Terms &amp; Conditions</div>
        <div class="info-grid">
          ${row('Payment Terms', orderData.paymentTerms + (orderData.advancePercent ? ' — Advance: ' + orderData.advancePercent + '%' : ''))}
          ${row('Guarantee Period', orderData.guaranteePeriod ? orderData.guaranteePeriod + ' months' : '')}
          ${row('Guarantee Clause', orderData.guaranteeClause)}
          ${(orderData.bgSecurityPct || orderData.bgSecurityDate || orderData.bgPerformancePct || orderData.bgPerformanceDate) ? row('Bank Guarantee',
            [
              orderData.bgSecurityPct    ? 'Security: ' + orderData.bgSecurityPct + '%'                              : '',
              orderData.bgSecurityDate   ? '(Valid till: ' + fmtDate(orderData.bgSecurityDate) + ')'                 : '',
              orderData.bgPerformancePct ? ' | Performance: ' + orderData.bgPerformancePct + '%'                     : '',
              orderData.bgPerformanceDate? '(Valid till: ' + fmtDate(orderData.bgPerformanceDate) + ')'              : ''
            ].filter(Boolean).join(' ')
          ) : ''}
          ${row('LD Clause', cdClauseText)}
        </div>
      </div>
    </div>
    <div class="col-right">
      <div class="section">
        <div class="section-title">Inspection — PDI &amp; Certification</div>
        <div class="info-grid">
          ${row('PDI', [orderData.pdi, orderData.pdiNote].filter(Boolean).join(' — '))}
          ${row('Certification', Array.isArray(orderData.certifications) ? orderData.certifications.join(', ') : orderData.certifications)}
        </div>
      </div>
      <div class="section" style="margin-top:4px;">
        <div class="section-title">Additional Information</div>
        <div class="info-grid">
          ${row('Documents Enclosed', Array.isArray(orderData.docsEnclosed) ? orderData.docsEnclosed.join(', ') : orderData.docsEnclosed)}
          ${row('Approval Required', (() => { const arr = Array.isArray(orderData.approvalReq) ? orderData.approvalReq : []; const s = orderData.sampleQty ? arr.map(a => a === 'Sample' ? 'Sample (' + orderData.sampleQty + ')' : a) : arr; return s.join(', '); })())}
          ${orderData.additionalNote ? row('Note', orderData.additionalNote) : ''}
        </div>
      </div>
    </div>
  </div>

  <!-- PRODUCT LINE ITEMS (full width) -->
  <div class="section-full">
    <div class="section-title">Product Line Items</div>
    <table class="product-table">
      <thead>
        <tr>
          <th style="width:4%">S.No</th>
          <th style="width:18%">Category</th>
          <th style="width:8%">Size</th>
          <th style="width:7%">Grade</th>
          <th style="width:15%">Short Description</th>
          <th style="width:${includeRate ? '26%' : '30%'}">Long Description</th>
          <th style="width:6%">Unit</th>
          <th style="width:6%">Qty</th>
          ${includeRate ? '<th style="width:10%">Rate</th>' : ''}
        </tr>
      </thead>
      <tbody>${productRows}</tbody>
    </table>
  </div>

  <!-- DOCUMENTS TO BE SENT WITH MATERIAL (full width) -->
  <div class="section-full" style="margin-top:5px;">
    <div class="section-title">Documents To Be Sent With Material</div>
    <div style="padding:4px 6px; font-size:8.5px; line-height:1.8;">
      ${(Array.isArray(orderData.docsWithMaterial) ? orderData.docsWithMaterial : []).map(d => `<span style="display:inline-block; margin-right:12px;">☑ ${d}</span>`).join('')}
    </div>
  </div>

  <!-- FOOTER -->
  <div class="footer">
    Generated on ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yyyy HH:mm:ss')}
    &nbsp;|&nbsp; This is a computer-generated document. No signature required.
  </div>
  </body></html>`;
}

/**
 * CREATE FACTORY COPY PDF HTML CONTENT
 * Selective sections only, Rate column hidden.
 * Sections printed:
 *   Customer Info: Name, GST, Address only
 *   Work Order Details: all
 *   Delivery Information: all
 *   Consignee Information: all
 *   Terms & Conditions: Payment Terms, Guarantee Period, Guarantee Clause, LD Clause only
 *   Inspection: all
 *   Product Details: all WITHOUT Rate
 *   Additional Information: all
 *   Documents To Be Sent With Material: all
 */
function createFactoryPDFHtmlContent(orderData, products) {
  function fmtDate(dateStr) {
    if (!dateStr) return '';
    try {
      const d = new Date(dateStr);
      return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd-MMM-yyyy');
    } catch(e) { return dateStr; }
  }

  function row(label, value) {
    return `<div class="info-row"><div class="info-label">${label}</div><div class="info-value">${value || '—'}</div></div>`;
  }

  const deliveryTermsCombined = [orderData.deliveryTerms, orderData.deliveryTermsExtra].filter(Boolean).join(' ');

  let cdClauseText = orderData.cdApplicable ? 'Applicable' : 'Not Applicable';
  if (orderData.cdApplicable && orderData.cdValue) {
    const period = orderData.cdPeriodWeek ? 'Per Week' : (orderData.cdPeriodMonth ? 'Per Month' : '');
    cdClauseText = `Applicable — ${orderData.cdValue} ${period}`.trim();
  }

  // Product rows — NO rate column
  let productRows = '';
  products.forEach((p, i) => {
    productRows += `<tr>
      <td style="text-align:center;">${i + 1}</td>
      <td>${p.category}</td>
      <td>${p.size}</td>
      <td>${p.grade}</td>
      <td>${p.short_desc}</td>
      <td>${p.long_desc}</td>
      <td style="text-align:center;">${p.unit}</td>
      <td style="text-align:right;">${p.qty}</td>
    </tr>`;
  });

  const sharedStyle = `
    @page { size: A4; margin: 8mm 10mm; }
    * { box-sizing: border-box; }
    body { font-family: Arial, sans-serif; margin: 0; padding: 0; color: #222; font-size: 9.5px; line-height: 1.3; }
    .header { text-align: center; border-bottom: 2.5px solid #0052CC; padding-bottom: 5px; margin-bottom: 6px; }
    .company-name { font-size: 16px; font-weight: bold; color: #0052CC; letter-spacing: 0.5px; }
    .document-title { font-size: 13px; font-weight: bold; margin-top: 2px; }
    .factory-badge { display:inline-block; background:#FF8B00; color:white; font-size:9px; font-weight:bold; padding:2px 8px; border-radius:3px; margin-top:3px; letter-spacing:0.5px; }
    .wo-meta { margin-top: 3px; font-size: 9.5px; }
    .two-col-layout { display: table; width: 100%; border-collapse: separate; border-spacing: 4px 0; margin-bottom: 5px; }
    .col-left, .col-right { display: table-cell; width: 50%; vertical-align: top; }
    .section { margin-bottom: 5px; }
    .section-full { margin-bottom: 5px; }
    .section-title { font-size: 9px; font-weight: bold; background: #0052CC; color: white; padding: 3px 6px; margin-bottom: 2px; text-transform: uppercase; letter-spacing: 0.3px; }
    .info-grid { display: table; width: 100%; border-collapse: collapse; }
    .info-row { display: table-row; }
    .info-label { display: table-cell; width: 42%; font-weight: bold; padding: 2px 5px; background: #f0f4ff; border: 0.5px solid #ccc; font-size: 8.5px; white-space: nowrap; }
    .info-value { display: table-cell; padding: 2px 5px; border: 0.5px solid #ccc; font-size: 8.5px; word-break: break-word; }
    .product-table { width: 100%; border-collapse: collapse; margin-top: 3px; }
    .product-table th { background: #0052CC; color: white; padding: 3px 4px; text-align: left; border: 0.5px solid #aac; font-size: 8.5px; white-space: nowrap; }
    .product-table td { padding: 3px 4px; border: 0.5px solid #ccc; font-size: 8.5px; vertical-align: top; }
    .product-table tr:nth-child(even) td { background: #f9f9ff; }
    .footer { margin-top: 6px; padding-top: 4px; border-top: 1px solid #ccc; text-align: center; font-size: 8px; color: #888; }
    @media print { body { -webkit-print-color-adjust: exact; print-color-adjust: exact; } }
  `;

  return `<!DOCTYPE html><html><head><meta charset="UTF-8"><style>${sharedStyle}</style></head><body>

  <!-- HEADER -->
  <div class="header">
    <div class="company-name">INDIAN CHAIN PRIVATE LIMITED</div>
    <div class="document-title">WORK ORDER</div>
    <div><span class="factory-badge">FACTORY COPY</span></div>
    <div class="wo-meta">
      <strong>Work Order No:</strong> ${orderData.workOrderNo}
      &nbsp;&nbsp;|&nbsp;&nbsp;
      <strong>Date:</strong> ${fmtDate(orderData.workOrderDate)}
    </div>
  </div>

  <!-- ROW 1: Customer Info (Name, GST, Address only) + Order Details -->
  <div class="two-col-layout">
    <div class="col-left">
      <div class="section">
        <div class="section-title">Customer Information</div>
        <div class="info-grid">
          ${row('Customer Name', orderData.customerName)}
          ${row('GST Number', orderData.customerGst)}
          ${row('Address', orderData.customerAddress)}
        </div>
      </div>
    </div>
    <div class="col-right">
      <div class="section">
        <div class="section-title">Order Details</div>
        <div class="info-grid">
          ${row('PO No.', orderData.poNo)}
          ${row('PO Date', fmtDate(orderData.poDate))}
          ${row('LOI No.', orderData.loiNo)}
          ${row('LOI Date', fmtDate(orderData.loiDate))}
        </div>
      </div>
    </div>
  </div>

  <!-- ROW 2: Delivery Info + Consignee Info -->
  <div class="two-col-layout">
    <div class="col-left">
      <div class="section">
        <div class="section-title">Delivery Information</div>
        <div class="info-grid">
          ${row('Delivery Date', fmtDate(orderData.deliveryDate))}
          ${row('Terms of Delivery', deliveryTermsCombined)}
          ${row('Nominated Transporter', orderData.nominatedTransporter)}
          ${row('Transporter Scope', orderData.transporterScope)}
        </div>
      </div>
    </div>
    <div class="col-right">
      <div class="section">
        <div class="section-title">Consignee Information</div>
        <div class="info-grid">
          ${row('Company Name', orderData.consigneeCompany)}
          ${row('Address', orderData.consigneeAddr)}
          ${row('GSTIN', orderData.consigneeGstin)}
          ${row('Contact Person', orderData.consigneeContact)}
          ${row('Contact Number', orderData.consigneePhone)}
        </div>
      </div>
    </div>
  </div>

  <!-- ROW 3: Terms & Conditions (selective) + Inspection -->
  <div class="two-col-layout">
    <div class="col-left">
      <div class="section">
        <div class="section-title">Terms &amp; Conditions</div>
        <div class="info-grid">
          ${row('Payment Terms', orderData.paymentTerms)}
          ${row('Guarantee Period', orderData.guaranteePeriod ? orderData.guaranteePeriod + ' months' : '')}
          ${row('Guarantee Clause', orderData.guaranteeClause)}
          ${row('LD Clause', cdClauseText)}
        </div>
      </div>
    </div>
    <div class="col-right">
      <div class="section">
        <div class="section-title">Inspection — PDI &amp; Certification</div>
        <div class="info-grid">
          ${row('PDI', [orderData.pdi, orderData.pdiNote].filter(Boolean).join(' — '))}
          ${row('Certification', Array.isArray(orderData.certifications) ? orderData.certifications.join(', ') : orderData.certifications)}
        </div>
      </div>
      <div class="section" style="margin-top:4px;">
        <div class="section-title">Additional Information</div>
        <div class="info-grid">
          ${row('Documents Enclosed', Array.isArray(orderData.docsEnclosed) ? orderData.docsEnclosed.join(', ') : orderData.docsEnclosed)}
          ${row('Approval Required', (() => { const arr = Array.isArray(orderData.approvalReq) ? orderData.approvalReq : []; const s = orderData.sampleQty ? arr.map(a => a === 'Sample' ? 'Sample (' + orderData.sampleQty + ')' : a) : arr; return s.join(', '); })())}
          ${orderData.additionalNote ? row('Note', orderData.additionalNote) : ''}
        </div>
      </div>
    </div>
  </div>

  <!-- PRODUCT LINE ITEMS — no Rate column -->
  <div class="section-full">
    <div class="section-title">Product Line Items</div>
    <table class="product-table">
      <thead>
        <tr>
          <th style="width:4%">S.No</th>
          <th style="width:20%">Category</th>
          <th style="width:9%">Size</th>
          <th style="width:8%">Grade</th>
          <th style="width:16%">Short Description</th>
          <th style="width:30%">Long Description</th>
          <th style="width:7%">Unit</th>
          <th style="width:6%">Qty</th>
        </tr>
      </thead>
      <tbody>${productRows}</tbody>
    </table>
  </div>

  <!-- DOCUMENTS TO BE SENT WITH MATERIAL -->
  <div class="section-full" style="margin-top:5px;">
    <div class="section-title">Documents To Be Sent With Material</div>
    <div style="padding:4px 6px; font-size:8.5px; line-height:1.8;">
      ${(Array.isArray(orderData.docsWithMaterial) ? orderData.docsWithMaterial : []).map(d => `<span style="display:inline-block; margin-right:12px;">☑ ${d}</span>`).join('')}
    </div>
  </div>

  <!-- FOOTER -->
  <div class="footer">
    Generated on ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yyyy HH:mm:ss')}
    &nbsp;|&nbsp; FACTORY COPY — This is a computer-generated document. No signature required.
  </div>
  </body></html>`;
}
