/**
 * Goods Receiving Management Automation
 * --------------------------------------
 * Portfolio-safe Google Apps Script example for automating goods receiving records.
 *
 * Technologies: Google Forms, Sheets, Docs, Drive, Apps Script Triggers.
 * No credentials or private document IDs should be hard-coded in this source.
 */

function createGoodsReceivingForm() {



  var form = FormApp.create("GOODS RECEIVING FORM");



  form.setDescription(

    "Form for recording received goods and raw materials."

  );



  form.setConfirmationMessage(

    "Goods receipt recorded successfully. Thank you."

  );



  // RECEIVING INFORMATION

  form.addSectionHeaderItem()

      .setTitle("A. RECEIVING INFORMATION");



  form.addDateItem()

      .setTitle("Goods Receipt Date")

      .setRequired(true);



  form.addTimeItem()

      .setTitle("Goods Receipt Time")

      .setRequired(true);



  form.addTextItem()

      .setTitle("Supplier / Partner Name")

      .setRequired(true);



  form.addTextItem()

      .setTitle("Delivery Note / Invoice Number")

      .setRequired(false);



  form.addTextItem()

      .setTitle("Receiving Officer")

      .setRequired(true);





  // ITEM DETAILS

  form.addSectionHeaderItem()

      .setTitle("B. GOODS DETAILS");



  form.addTextItem()

      .setTitle("Item Name")

      .setRequired(true);



  form.addTextItem()

      .setTitle("Quantity")

      .setRequired(true);



  form.addListItem()

      .setTitle("Unit")

      .setChoiceValues([

        "Kg",

        "Gram",

        "Liter",

        "Mililiter",

        "Buah",

        "Dus",

        "Karton",

        "Pack",

        "Ikat",

        "Lainnya"

      ])

      .setRequired(true);



  form.addMultipleChoiceItem()

      .setTitle("Item Condition")

      .setChoiceValues([

        "Good / Compliant",

        "Partially non-compliant",

        "Non-compliant / Rejected"

      ])

      .setRequired(true);



  form.addTextItem()

      .setTitle("Item Temperature (°C) - if applicable")

      .setRequired(false);



  form.addParagraphTextItem()

      .setTitle("Item Notes")

      .setRequired(false);





  // DOCUMENTATION

  form.addSectionHeaderItem()

      .setTitle("C. DOCUMENTATION");



  form.addParagraphTextItem()

      .setTitle("Additional Notes")

      .setRequired(false);





  Logger.log("FORM CREATED");

  Logger.log("Form URL: " + form.getPublishedUrl());

  Logger.log("Edit URL: " + form.getEditUrl());

}

function createReportTemplate() {



  // =========================================

  // CREATE ROOT FOLDER

  // =========================================



  var rootFolderName = "GOODS RECEIVING AUTOMATION SYSTEM";



  var folderList = DriveApp.getFoldersByName(rootFolderName);



  var folder;



  if (folderList.hasNext()) {

    folder = folderList.next();

  } else {

    folder = DriveApp.createFolder(rootFolderName);

  }





  // =========================================

  // CREATE GOOGLE DOCS

  // =========================================



  var doc = DocumentApp.create(

    "GOODS RECEIVING REPORT TEMPLATE"

  );



  var body = doc.getBody();





  // =========================================

  // TITLE

  // =========================================



  var titleParagraph = body.appendParagraph(

    "GOODS RECEIVING REPORT"

  );



  titleParagraph.setAlignment(DocumentApp.HorizontalAlignment.CENTER);



  titleParagraph.setBold(true);

  titleParagraph.setFontSize(16);





  var subtitleParagraph = body.appendParagraph(

    "GOODS RECEIVING SYSTEM"

  );



  subtitleParagraph.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  subtitleParagraph.setBold(true);

  subtitleParagraph.setFontSize(13);





  body.appendParagraph("");





  // =========================================

  // RECEIVING INFORMATION

  // =========================================



  var infoSection = body.appendParagraph(

    "A. RECEIVING INFORMATION"

  );



  infoSection.setBold(true);

  infoSection.setFontSize(12);





  body.appendTable([

    ["Menu Name", "{{MENU_NAME}}"],

    ["Goods Receipt Date", "{{DATE}}"],

    ["Supplier / Partner Name", "{{SUPPLIER_NAME}}"],

    ["Receiving Officer", "{{RECEIVING_OFFICER}}"]

  ]);





  body.appendParagraph("");





  // =========================================

  // GOODS DETAILS

  // =========================================



  var goodsSection = body.appendParagraph(

    "B. GOODS DETAILS"

  );



  goodsSection.setBold(true);

  goodsSection.setFontSize(12);





  body.appendTable([

    [

      "No",

      "Item Name",

      "Quantity",

      "Unit",

      "Condition",

      "Temperature"

    ],

    [

      "1",

      "{{ITEM_NAME}}",

      "{{QUANTITY}}",

      "{{UNIT}}",

      "{{CONDITION}}",

      "{{TEMPERATURE}}"

    ]

  ]);





  body.appendParagraph("");





  // =========================================

  // DOCUMENTATION

  // =========================================



  var documentationSection = body.appendParagraph(

    "C. DOCUMENTATION"

  );



  documentationSection.setBold(true);

  documentationSection.setFontSize(12);





  body.appendParagraph(

    "{{ITEM_PHOTOS}}"

  );





  body.appendParagraph("");





  // =========================================

  // NOTES

  // =========================================



  var notesSection = body.appendParagraph(

    "D. NOTES"

  );



  notesSection.setBold(true);

  notesSection.setFontSize(12);





  body.appendParagraph(

    "{{NOTES}}"

  );





  // =========================================

  // SAVE

  // =========================================



  doc.saveAndClose();





  // Move document to system folder



  var file = DriveApp.getFileById(doc.getId());



  folder.addFile(file);



  DriveApp.getRootFolder().removeFile(file);





  // =========================================

  // LOGGING

  // =========================================



  Logger.log(

    "REPORT TEMPLATE CREATED"

  );



  Logger.log(

    "Template URL: " + doc.getUrl()

  );



  Logger.log(

    "System Folder: " + folder.getUrl()

  );

}

function createDailyMenuForm() {



  var form = FormApp.create("DAILY MENU FORM");



  form.setDescription(

    "Form for recording the menu used for the day. " +

    "This form only needs to be completed once per day."

  );



  form.setConfirmationMessage(

    "Daily menu recorded successfully."

  );



  // ==============================

  // DAILY MENU DATA

  // ==============================



  form.addSectionHeaderItem()

      .setTitle("DAILY MENU");



  form.addDateItem()

      .setTitle("Menu Date")

      .setRequired(true);



  form.addTextItem()

      .setTitle("Menu Name")

      .setRequired(true);



  form.addParagraphTextItem()

      .setTitle("Menu Notes")

      .setRequired(false);





  // ==============================

  // CREATE MENU DATABASE

  // ==============================



  var spreadsheet = SpreadsheetApp.create(

    "DAILY MENU DATABASE"

  );



  form.setDestination(

    FormApp.DestinationType.SPREADSHEET,

    spreadsheet.getId()

  );





  // ==============================

  // LOGGING

  // ==============================



  Logger.log("=================================");

  Logger.log("FORM DAILY MENU BERHASIL DIBUAT");

  Logger.log("Form URL: " + form.getPublishedUrl());

  Logger.log("Edit URL: " + form.getEditUrl());

  Logger.log("Database: " + spreadsheet.getUrl());

  Logger.log("=================================");

}



function installAutomation() {



  // ==========================================

  // EXPECTED FILE NAMES

  // ==========================================



  var goodsSpreadsheetName =

    "GOODS RECEIVING FORM (Jawaban)";



  var menuSpreadsheetName =

    "DAILY MENU DATABASE";



  var templateName =

    "GOODS RECEIVING REPORT TEMPLATE";





  // ==========================================

  // FIND GOODS SPREADSHEET

  // ==========================================



  var goodsSpreadsheet =

    findLatestFile(goodsSpreadsheetName);



  if (!goodsSpreadsheet) {

    throw new Error(

      "Goods receiving database not found: " +

      goodsSpreadsheetName

    );

  }





  // ==========================================

  // FIND MENU DATABASE

  // ==========================================



  var menuSpreadsheet =

    findLatestFile(menuSpreadsheetName);



  if (!menuSpreadsheet) {

    throw new Error(

      "Daily menu database not found: " +

      menuSpreadsheetName

    );

  }





  // ==========================================

  // FIND REPORT TEMPLATE

  // ==========================================



  var template =

    findLatestFile(templateName);



  if (!template) {

    throw new Error(

      "Report template not found: " +

      templateName

    );

  }





  // ==========================================

  // STORE IDS IN SCRIPT PROPERTIES

  // ==========================================



  var properties =

    PropertiesService.getScriptProperties();



  properties.setProperties({



    GOODS_SPREADSHEET_ID:

      goodsSpreadsheet.getId(),



    MENU_SPREADSHEET_ID:

      menuSpreadsheet.getId(),



    REPORT_TEMPLATE_ID:

      template.getId()



  });





  // ==========================================

  // CHECK EXISTING TRIGGERS

  // ==========================================



  var triggers =

    ScriptApp.getProjectTriggers();



  for (var i = 0; i < triggers.length; i++) {



    if (

      triggers[i].getHandlerFunction() ===

      "onGoodsReceived"

    ) {



      ScriptApp.deleteTrigger(

        triggers[i]

      );

    }

  }





  // ==========================================

  // CREATE AUTOMATIC TRIGGER

  // ==========================================



  var goodsSpreadsheetAsli =

  SpreadsheetApp.openById(goodsSpreadsheet.getId());



ScriptApp.newTrigger("onGoodsReceived")

  .forSpreadsheet(goodsSpreadsheetAsli)

  .onFormSubmit()

  .create();





  Logger.log(

    "======================================"

  );



  Logger.log(

    "AUTOMATION INSTALLED SUCCESSFULLY"

  );



  Logger.log(

    "Goods Database: " +

    goodsSpreadsheet.getUrl()

  );



  Logger.log(

    "Menu Database: " +

    menuSpreadsheet.getUrl()

  );



  Logger.log(

    "Template: " +

    template.getUrl()

  );



  Logger.log(

    "Automatic trigger is active."

  );



  Logger.log(

    "======================================"

  );

}







// ==========================================

// FIND LATEST FILE

// ==========================================



function findLatestFile(headerNameFile) {



  var files =

    DriveApp.getFilesByName(headerNameFile);



  var latestFile = null;

  var latestTimestamp = 0;



  while (files.hasNext()) {



    var file = files.next();



    var waktu =

      file.getLastUpdated().getTime();



    if (waktu > latestTimestamp) {



      latestTimestamp = waktu;

      latestFile = file;



    }

  }



  return latestFile;

}







// ==========================================

// FORM SUBMISSION TRIGGER

// ==========================================



function onGoodsReceived(e) {



  var lock =

    LockService.getScriptLock();



  lock.waitLock(30000);



  try {



    Logger.log(

      "New goods received. Generating report..."

    );



    generateDailyReport(e);



  } finally {



    lock.releaseLock();



  }

}

function generateDailyReport(e) {



  var properties =

    PropertiesService.getScriptProperties();



  var barangId =

    properties.getProperty("GOODS_SPREADSHEET_ID");



  var menuId =

    properties.getProperty("MENU_SPREADSHEET_ID");



  var templateId =

    properties.getProperty("REPORT_TEMPLATE_ID");



  if (!barangId || !menuId || !templateId) {

    throw new Error(

      "System IDs are incomplete. Run installAutomation() first."

    );

  }



  // ==========================================

  // GOODS DATABASE

  // ==========================================



  var ssBarang =

    SpreadsheetApp.openById(barangId);



  var sheetBarang =

    ssBarang.getSheets()[0];



  var goodsData =

    sheetBarang.getDataRange().getValues();



  if (goodsData.length < 2) {

    throw new Error(

      "No goods data available."

    );

  }



  var goodsHeaders =

    goodsData[0];



  var dateColumn =

    findColumn(goodsHeaders, [

      "Goods Receipt Date",

      "Goods Receipt Date"

    ]);



  var supplierColumn =

    findColumn(goodsHeaders, [

      "Supplier Name",

      "Supplier Name"

    ]);



  var officerColumn =

    findColumn(goodsHeaders, [

      "Receiving Officer",

      "Receiving Officer"

    ]);



  var itemNameColumn =

    findColumn(goodsHeaders, [

      "Item Name",

      "Item Name"

    ]);



  var quantityColumn =

    findColumn(goodsHeaders, [

      "Quantity",

      "Quantity"

    ]);



  var unitColumn =

    findColumn(goodsHeaders, [

      "Unit",

      "Unit"

    ]);



  var conditionColumn =

    findColumn(goodsHeaders, [

      "Item Condition",

      "Item Condition"

    ]);



  var temperatureColumn =

    findColumn(goodsHeaders, [

      "Item Temperature (°C) - if applicable",

      "Item Temperature (°C)",

      "Item Temperature (°C) - if applicable",

      "Item Temperature (°C)"

    ]);



  var photoColumn =

    findColumn(goodsHeaders, [

      "UPLOAD ITEM PHOTOS",

      "Upload Item Photos"

    ]);



  var notesColumn =

    findColumn(goodsHeaders, [

      "Item Notes",

      "Item Notes"

    ]);



  // ==========================================

  // DETERMINE DATE FROM SUBMITTED ROW

  // ==========================================



  var rowNumber;



  if (e && e.range) {

    rowNumber =

      e.range.getRow();

  } else {

    rowNumber =

      goodsData.length;

  }



  var submittedRow =

    goodsData[rowNumber - 1];



  var receiptDateValue =

    submittedRow[dateColumn];



  var timezone =

    Session.getScriptTimeZone() ||

    "Asia/Jakarta";



  var dateKey =

    toDateKey(receiptDateValue, timezone);



  if (!dateKey) {

    throw new Error(

      "Goods receipt date could not be read."

    );

  }



  // ==========================================

  // GET ALL GOODS FOR THE DATE

  // ==========================================



  var allGoods = [];



  for (

    var i = 1;

    i < goodsData.length;

    i++

  ) {



    var rowData =

      goodsData[i];



    var key =

      toDateKey(

        rowData[dateColumn],

        timezone

      );



    if (key === dateKey) {



      allGoods.push({



        supplier:

          safeValue(rowData[supplierColumn]),



        petugas:

          safeValue(rowData[officerColumn]),



        headerNameBarang:

          safeValue(rowData[itemNameColumn]),



        jumlah:

          safeValue(rowData[quantityColumn]),



        satuan:

          safeValue(rowData[unitColumn]),



        kondisi:

          safeValue(rowData[conditionColumn]),



        suhu:

          safeValue(rowData[temperatureColumn]),



        foto:

          safeValue(rowData[photoColumn]),



        keterangan:

          safeValue(rowData[notesColumn])



      });



    }

  }



  if (allGoods.length === 0) {

    throw new Error(

      "No goods found for date " +

      dateKey

    );

  }



  // ==========================================

  // DATABASE DAILY MENU

  // ==========================================



  var ssMenu =

    SpreadsheetApp.openById(menuId);



  var sheetMenu =

    ssMenu.getSheets()[0];



  var dataMenu =

    sheetMenu.getDataRange().getValues();



  if (dataMenu.length < 2) {

    throw new Error(

      "Daily menu database is empty."

    );

  }



  var headerMenu =

    dataMenu[0];



  var dateColumnMenu =

    findColumn(headerMenu, [

      "Menu Date",

      "Menu Date"

    ]);



  var colNamaMenu =

    findColumn(headerMenu, [

      "Menu Name",

      "Menu Name"

    ]);



  var notesColumnMenu =

    findColumn(headerMenu, [

      "Menu Notes",

      "Menu Notes"

    ]);



  var menuName = "";

  var menuNotes = "";



  // GET MENU FOR THE DATE

  for (

    var m = dataMenu.length - 1;

    m >= 1;

    m--

  ) {



    var menuDateKey =

      toDateKey(

        dataMenu[m][dateColumnMenu],

        timezone

      );



    if (menuDateKey === dateKey) {



      menuName =

        safeValue(dataMenu[m][colNamaMenu]);



      menuNotes =

        safeValue(dataMenu[m][notesColumnMenu]);



      break;

    }

  }



  if (!menuName) {

    menuName =

      "MENU NOT ENTERED";



    Logger.log(

      "Warning: No menu found for date " +

      dateKey +

      " belum ditemukan."

    );

  }



  // ==========================================

  // REPORT FOLDER

  // ==========================================



  var templateFile =

    DriveApp.getFileById(templateId);



  var parents =

    templateFile.getParents();



  if (!parents.hasNext()) {

    throw new Error(

      "Template folder not found."

    );

  }



  var rootFolder =

    parents.next();



  var dateObject =

    createDateObject(

      dateKey

    );



  var tahun =

    Utilities.formatDate(

      dateObject,

      timezone,

      "yyyy"

    );



  var bulan =

    Utilities.formatDate(

      dateObject,

      timezone,

      "MM"

    );



  var headerNameBulan =

    Utilities.formatDate(

      dateObject,

      timezone,

      "MMMM"

    );



  var rootFolderNameBulan =

    bulan +

    " - " +

    headerNameBulan.toUpperCase();



  var yearFolder =

    getOrCreateFolder(

      rootFolder,

      tahun

    );



  var monthFolder =

    getOrCreateFolder(

      yearFolder,

      monthFolderName

    );



  var rootFolderNameTanggal =

    Utilities.formatDate(

      dateObject,

      timezone,

      "dd-MM-yyyy"

    );



  var dateFolder =

    getOrCreateFolder(

      monthFolder,

      dateFolderName

    );



  // ==========================================

  // REMOVE PREVIOUS REPORT FILES FOR DATE

  // ==========================================



  deleteFileIfExists(

    dateFolder,

    "GOODS RECEIVING REPORT - " +

    dateFolderName +

    ".pdf"

  );



  deleteFileIfExists(

    dateFolder,

    "GOODS RECEIVING REPORT - " +

    dateFolderName +

    ".docx"

  );



  // ==========================================

  // COPY TEMPLATE

  // ==========================================



  var reportName =

    "GOODS RECEIVING REPORT - " +

    dateFolderName;



  var reportCopy =

    templateFile.makeCopy(

      reportName,

      dateFolder

    );



  var doc =

    DocumentApp.openById(

      reportCopy.getId()

    );



  var body =

    doc.getBody();



  // ==========================================

  // GENERAL INFORMATION

  // ==========================================



  var suppliers =

    uniqueValues(

      allGoods.map(function(item) {

        return item.supplier;

      })

    );



  var receivingOfficers =

    uniqueValues(

      allGoods.map(function(item) {

        return item.petugas;

      })

    );



  var displayDate =

    Utilities.formatDate(

      dateObject,

      timezone,

      "dd/MM/yyyy"

    );



  replaceText(

    body,

    "{{MENU_NAME}}",

    menuName

  );



  replaceText(

    body,

    "{{DATE}}",

    displayDate

  );





  replaceText(

    body,

    "{{RECEIVING_OFFICER}}",

    receivingOfficers.join(", ")

  );



  // ==========================================

  // GOODS DETAILS

  // ==========================================



  var goodsTable =

    findGoodsTable(body);



  if (goodsTable) {



    // Hapus rowData contoh

    if (goodsTable.getNumRows() > 1) {

      goodsTable.removeRow(1);

    }



    for (

      var b = 0;

      b < allGoods.length;

      b++

    ) {



      var item =

        allGoods[b];



      var row =

        goodsTable.appendTableRow();



      row.appendTableCell(

        String(b + 1)

      );



      row.appendTableCell(

        item.headerNameBarang

      );



      row.appendTableCell(

        String(item.jumlah)

      );



      row.appendTableCell(

        item.satuan

      );



      row.appendTableCell(

        item.kondisi

      );



      row.appendTableCell(

        item.suhu

      );



      row.appendTableCell(

        item.supplier

      );

    }

  }



  // ==========================================

  // NOTES

  // ==========================================



  var allNotes =

    uniqueValues(

      allGoods

        .map(function(item) {

          return item.keterangan;

        })

        .filter(function(x) {

          return x !== "";

        })

    );



  var combinedNotes =

    allNotes.length > 0

      ? allNotes.join("\n")

      : menuNotes;



  replaceText(

    body,

    "{{NOTES}}",

    combinedNotes

  );



  // ==========================================

// GOODS RECEIVING DOCUMENTATION

// ==========================================



body.appendPageBreak();



var titleParagraphDokumentasi = body.appendParagraph(

  "GOODS RECEIVING DOCUMENTATION"

);



judulDokumentasi.setBold(true);

judulDokumentasi.setFontSize(13);

judulDokumentasi.setAlignment(

  DocumentApp.HorizontalAlignment.CENTER

);





// TANGGAL

var tanggalDokumentasi = body.appendParagraph(

  "Tanggal: " + Utilities.formatDate(

    dateObject,

    Session.getScriptTimeZone(),

    "dd MMMM yyyy"

  )

);

tanggalDokumentasi.setBold(true);

tanggalDokumentasi.setFontSize(10);

tanggalDokumentasi.setAlignment(

  DocumentApp.HorizontalAlignment.CENTER

);

tanggalDokumentasi.setSpacingAfter(18);





// FOTO

insertPhotoGallery(body, allGoods);



doc.saveAndClose();





  // ==========================================

  // EXPORT PDF

  // ==========================================



  var documentId =

    reportCopy.getId();



  var token =

    ScriptApp.getOAuthToken();



  var pdfUrl =

    "https://www\.googleapis.com/drive/v3/files/" +

    documentId +

    "/export?mimeType=" +

    encodeURIComponent(

      "application/pdf"

    );



  var pdfResponse =

    UrlFetchApp.fetch(

      pdfUrl,

      {

        headers: {

          Authorization:

            "Bearer " + token

        },

        muteHttpExceptions: true

      }

    );



  if (

    pdfResponse.getResponseCode() !== 200

  ) {



    throw new Error(

      "Gagal membuat PDF. Kode: " +

      pdfResponse.getResponseCode()

    );

  }



  var pdfBlob =

    pdfResponse

      .getBlob()

      .setName(

        reportName + ".pdf"

      );



  dateFolder.createFile(

    pdfBlob

  );



  // ==========================================

  // EXPORT WORD / DOCX

  // ==========================================



  var docxUrl =

    "https://www\.googleapis.com/drive/v3/files/" +

    documentId +

    "/export?mimeType=" +

    encodeURIComponent(

      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

    );



  var docxResponse =

    UrlFetchApp.fetch(

      docxUrl,

      {

        headers: {

          Authorization:

            "Bearer " + token

        },

        muteHttpExceptions: true

      }

    );



  if (

    docxResponse.getResponseCode() !== 200

  ) {



    throw new Error(

      "Gagal membuat Word. Kode: " +

      docxResponse.getResponseCode()

    );

  }



  var docxBlob =

    docxResponse

      .getBlob()

      .setName(

        reportName + ".docx"

      );



  dateFolder.createFile(

    docxBlob

  );



  // Hapus Google Docs result copy

  reportCopy.setTrashed(true);



  Logger.log(

    "======================================"

  );



  Logger.log(

    "DAILY REPORT GENERATED SUCCESSFULLY"

  );



  Logger.log(

    "Tanggal: " +

    displayDate

  );



  Logger.log(

    "Menu: " +

    menuName

  );



  Logger.log(

    "Item count: " +

    allGoods.length

  );



  Logger.log(

    "Folder: " +

    dateFolder.getUrl()

  );



  Logger.log(

    "======================================"

  );

}



// ==================================================

// FIND COLUMN BY HEADER NAME

// ==================================================



function findColumn(header, daftarNama) {



  for (var i = 0; i < header.length; i++) {



    var headerName =

      String(header[i])

        .trim()

        .toLowerCase();



    for (

      var j = 0;

      j < daftarNama.length;

      j++

    ) {



      if (

        headerName ===

        String(daftarNama[j])

          .trim()

          .toLowerCase()

      ) {



        return i;



      }

    }

  }



  throw new Error(

    "Column not found: " +

    daftarNama.join(" / ")

  );

}





// ==================================================

// CONVERT DATE TO KEY

// ==================================================



function toDateKey(

  value,

  timezone

) {



  if (!value) {

    return "";

  }



  if (

    Object.prototype.toString.call(value) ===

    "[object Date]" &&

    !isNaN(value.getTime())

  ) {



    return Utilities.formatDate(

      value,

      timezone,

      "yyyy-MM-dd"

    );

  }



  var teks =

    String(value).trim();



  // dd/MM/yyyy

  var match =

    teks.match(

      /^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$/

    );



  if (match) {



    return (

      match[3] +

      "-" +

      ("0" + match[2]).slice(-2) +

      "-" +

      ("0" + match[1]).slice(-2)

    );

  }



  var tanggal =

    new Date(teks);



  if (!isNaN(tanggal.getTime())) {



    return Utilities.formatDate(

      tanggal,

      timezone,

      "yyyy-MM-dd"

    );

  }



  return "";

}





// ==================================================

// CREATE DATE OBJECT

// ==================================================



function createDateObject(

  dateKey

) {



  var parts =

    dateKey.split("-");



  return new Date(

    Number(parts[0]),

    Number(parts[1]) - 1,

    Number(parts[2])

  );

}





// ==================================================

// SAFE VALUE

// ==================================================



function safeValue(value) {



  if (

    value === null ||

    value === undefined

  ) {



    return "";



  }



  return String(value).trim();

}





// ==================================================

// REPLACE TEXT

// ==================================================



function replaceText(

  body,

  placeholder,

  value

) {



  var pola =

    placeholder.replace(

      /[.*+?^${}()|[\]\\]/g,

      "\\\\$&"

    );



  body.replaceText(

    pola,

    String(value || "")

  );

}





// ==================================================

// CARI TABEL GOODS DETAILS

// ==================================================



function findGoodsTable(

  body

) {



  var tables =

    body.getTables();



  for (

    var i = 0;

    i < tables.length;

    i++

  ) {



    var teks =

      tables[i]

        .getText()

        .toLowerCase();



    if (

      teks.indexOf("headerName barang") !== -1 &&

      teks.indexOf("jumlah") !== -1 &&

      teks.indexOf("satuan") !== -1

    ) {



      return tables[i];



    }

  }



  return null;

}





// ==================================================

// GET UNIQUE VALUES

// ==================================================



function uniqueValues(

  array

) {



  var result = [];



  for (

    var i = 0;

    i < array.length;

    i++

  ) {



    var value =

      String(array[i] || "").trim();



    if (

      value !== "" &&

      result.indexOf(value) === -1

    ) {



      result.push(value);



    }

  }



  return result;

}





// ==================================================

// FIND OR CREATE FOLDER

// ==================================================



function getOrCreateFolder(

  parent,

  headerName

) {



  var folders =

    parent.getFoldersByName(headerName);



  if (folders.hasNext()) {

    return folders.next();

  }



  return parent.createFolder(headerName);

}





// ==================================================

// DELETE FILE IF EXISTS

// ==================================================



function deleteFileIfExists(

  folder,

  headerName

) {



  var files =

    folder.getFilesByName(headerName);



  while (files.hasNext()) {



    files.next().setTrashed(true);



  }

}





// ==================================================

// INSERT PHOTOS

// ==================================================

// ==================================================

// INSERT 3 X 3 PHOTO GALLERY

// ==================================================



function insertPhotoGallery(body, allGoods) {



  // ==========================================

  // KUMPULKAN FOTO + NAMA BAHAN

  // ==========================================



  var semuaFoto = [];



  for (var i = 0; i < allGoods.length; i++) {



    var photoText = allGoods[i].foto;



    if (!photoText) continue;



    var photoUrls = String(photoText)

      .split(/\s\*,\s\*|\n/)

      .filter(function(url) {

        return url.trim() !== "";

      });



    // SETIAP FOTO TETAP TERIKAT DENGAN NAMA BAHANNYA

    for (var j = 0; j < photoUrls.length; j++) {



      var photoFileId = extractFileId(photoUrls[j]);



      if (photoFileId) {



        semuaFoto.push({

          id: photoFileId,

          headerNameBarang: allGoods[i].headerNameBarang

        });



      }

    }

  }



  Logger.log(

    "Jumlah foto ditemukan: " + semuaFoto.length

  );





  // ==========================================

  // JIKA TIDAK ADA FOTO

  // ==========================================



  if (semuaFoto.length === 0) {



    body.appendParagraph(

      "No photos available."

    );



    return;

  }





  // ==========================================

  // FOTO 3 x 3

  // ==========================================



  var fotoPerHalaman = 15;





  for (

    var mulai = 0;

    mulai < semuaFoto.length;

    mulai += fotoPerHalaman

  ) {



    if (mulai > 0) {

      body.appendPageBreak();

    }





    var photoCount = Math.min(

      fotoPerHalaman,

      semuaFoto.length - mulai

    );





    var rowCount = Math.ceil(

      photoCount / 3

    );





    var tableData = [];





    for (

      var r = 0;

      r < rowCount;

      r++

    ) {



      tableData.push([

        "",

        "",

        ""

      ]);



    }





    var tabel = body.appendTable(

      tableData

    );





    // ==========================================

    // INSERT PHOTOS

    // ==========================================



    for (

      var n = 0;

      n < photoCount;

      n++

    ) {



      var rowData = Math.floor(n / 3);

      var kolom = n % 3;



      var cell = tabel.getCell(

        rowData,

        kolom

      );





      var photoData =

        semuaFoto[mulai + n];





      try {



        var file =

          DriveApp.getFileById(

            photoData.id

          );





        var paragraphFoto =

          cell.getChild(0).asParagraph();





        paragraphFoto.clear();





        var image =

          paragraphFoto.appendInlineImage(

            file.getBlob()

          );





        // Ukuran foto

        var maxWidth = 155;

        var maxHeight = 115;





        var width =

          image.getWidth();



        var height =

          image.getHeight();





        var skala =

          Math.min(

            maxWidth / width,

            maxHeight / height,

            1

          );





        image.setWidth(

          Math.round(width * skala)

        );





        image.setHeight(

          Math.round(height * skala)

        );





        paragraphFoto.setAlignment(

          DocumentApp.HorizontalAlignment.CENTER

        );





        // ======================================

        // NAMA BAHAN

        // ======================================



        var label =

          cell.appendParagraph(

            photoData.headerNameBarang

          );





        label.setAlignment(

          DocumentApp.HorizontalAlignment.CENTER

        );





        label.editAsText()

          .setFontSize(8)

          .setBold(true);





        Logger.log(

          "Photo inserted: " +

          photoData.headerNameBarang

        );





      } catch (error) {



        cell

          .getChild(0)

          .asParagraph()

          .setText(

            "Photo could not be loaded"

          );





        Logger.log(

          "Photo failed: " +

          error.message

        );



      }



    }





    // ==========================================

    // POSISI VERTIKAL CELL

    // ==========================================



    for (

      var rr = 0;

      rr < rowCount;

      rr++

    ) {



      for (

        var cc = 0;

        cc < 3;

        cc++

      ) {



        tabel

          .getCell(rr, cc)

          .setVerticalAlignment(

            DocumentApp.VerticalAlignment.CENTER

          );



      }



    }



  }





  Logger.log(

    "PHOTO GALLERY COMPLETE. Total: " +

    semuaFoto.length

  );



}





// ==================================================

// EXTRACT FILE ID FROM GOOGLE DRIVE URL

// ==================================================



function extractFileId(

  url

) {



  var teks =

    String(url || "");



  var match =

    teks.match(

      /[-\w]{25,}/

    );



  return match

    ? match[0]

    : "";

}





// ==================================================

// TEST LATEST SUBMISSION

// ==================================================



function testLatestReport() {



  var properties =

    PropertiesService.getScriptProperties();



  var barangId =

    properties.getProperty(

      "GOODS_SPREADSHEET_ID"

    );



  if (!barangId) {



    throw new Error(

      "Automation has not been installed."

    );



  }



  var ss =

    SpreadsheetApp.openById(

      barangId

    );



  var sheet =

    ss.getSheets()[0];



  var lastRow =

    sheet.getLastRow();



  if (lastRow < 2) {



    throw new Error(

      "No goods data available."

    );



  }



  var range =

    sheet.getRange(

      lastRow,

      1,

      1,

      sheet.getLastColumn()

    );



  Logger.log(

    "Using row " +

    lastRow +

    " for testing."

  );



  generateDailyReport({

    range: range

  });

}
