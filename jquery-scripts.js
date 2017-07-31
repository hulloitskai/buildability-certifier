const AppStorage = require('electron-store');
const Excel = require('exceljs');
const ipc = require('electron').ipcRenderer;
var appStorage = new AppStorage();

var preferencePaneActive = false;

function showPreferences(){
   $("#shadowcover").show().delay(50).animate({
      opacity: '0.6'
   },{
      duration: 400,
   });
   $("#preferences-pane").show().delay(300).animate({
      top: '0'
   },{
      easing: 'easeOutElastic',
      duration: 700,
      complete: function(){
         loadPreferences();
      }
   });
   $("#preferences-pane .body").delay(50).scrollTop(0).
   preferencePaneActive = true;
}

function hidePreferences(){
   $("#preferences-pane").animate({
      top: '1250'
   },{
      easing: 'easeInCirc',
      duration: 300,
      complete: function(){
         $(this).hide();
      }
   });

   $("#shadowcover").delay(200).animate({
      opacity: '0'
   },{
      complete: function(){
         $(this).hide();
      }
   });
   preferencePaneActive = false;
}


function loadPreferences(){
   if (appStorage.has('excel-file-directory')){
      $("#selected-file-directory p").text(appStorage.get('excel-file-directory'));
      initializeColumnOptions();
   }
   else{
      $("#selected-file-directory p").text("no file selected");
      $(".dropdown-form select").replaceWith("<select><option>no file selected</option></select>");
   }
}

function savePreferences(){
   appStorage.set('excel-file-directory', $("#selected-file-directory p").text());
   appStorage.set('name-column-selected', $("#name-column-form select").find(":selected").text());
   appStorage.set('email-column-selected', $("#email-column-form select").find(":selected").text());
   appStorage.set('award-column-selected', $("#award-column-form select").find(":selected").text());
   appStorage.set('identifier-column-selected', $("#identifier-column-form select").find(":selected").text());
   appStorage.set('completion-column-selected', $("#completion-column-form select").find(":selected").text());
}

function openFileSelector(){
   $("#selected-file-directory p").text("no file selected");
   $(".dropdown-form select").replaceWith("<select><option>no file selected</option></select>");
   ipc.send('open-file-dialog');
   ipc.on('selected-directory', function (event, path) {
      $("#selected-file-directory p").text(path);
      initializeColumnOptions();
   })
}

function dropdownOptionMaker(array){
   var htmlcontent = "<select>";
   array.forEach(function(item, index){
      htmlcontent += "<option>" + item + "</option>";
   })
   htmlcontent += "</select>";
   return htmlcontent;
}

function initializeColumnOptions(){
   var workbook = new Excel.Workbook();
   const excelFileDirectory = $("#selected-file-directory p").text();
   workbook.xlsx.readFile(excelFileDirectory).then(function(){
      var arrayOfColumns = [];
      workbook.eachSheet(function(worksheet, sheetId) {
         worksheet.getRow(1).eachCell(function(cell, colNumber){
            arrayOfColumns.push(cell.value);
         });
      });
      $(".dropdown-form select").replaceWith(dropdownOptionMaker(arrayOfColumns));
      if (appStorage.get('excel-file-directory') == $("#selected-file-directory p").text()){
         $("#name-column-form option").filter(function() {
            return $(this).text() === appStorage.get('name-column-selected');
         }).prop('selected', true);
         $("#email-column-form option").filter(function() {
            return $(this).text() === appStorage.get('email-column-selected');
         }).prop('selected', true);
         $("#identifier-column-form option").filter(function() {
            return $(this).text() === appStorage.get('identifier-column-selected');
         }).prop('selected', true);
         $("#award-column-form option").filter(function() {
            return $(this).text() === appStorage.get('award-column-selected');
         }).prop('selected', true);
         $("#completion-column-form option").filter(function() {
            return $(this).text() === appStorage.get('completion-column-selected');
         }).prop('selected', true);
      }
   });
}

function loadExcelNames(){
   var workbook = new Excel.Workbook();
   const excelFileDirectory = appStorage.get('excel-file-directory');
   workbook.xlsx.readFile(excelFileDirectory).then(function(){
      var sheetIdentity;
      var nameColumnNumber;
      workbook.eachSheet(function(worksheet, sheetId) {
         worksheet.getRow(1).eachCell(function(cell, colNumber){
            if (cell.value == appStorage.get('name-column-selected')){
               nameColumnNumber = colNumber;
               sheetIdentity = sheetId;
               return
            }
         });
      });
      var worksheetWithNames = workbook.getWorksheet(sheetIdentity);
      var arrayOfNames = [];
      worksheetWithNames.getColumn(nameColumnNumber).eachCell(function(cell, rowNumber){
         if (cell.value != appStorage.get('name-column-selected')){
            if (cell.value != null){
               arrayOfNames.push(cell.value);
            }
         }
      });
      awesomplete.list = arrayOfNames;
      globalNameList = arrayOfNames;
      awesomplete.evaluate();
   });
}

var awesomplete;
var globalNameList;

function certifyClients(){
   const clientNames = getSelectedClients();
   const workbook = new Excel.Workbook();
   workbook.xlsx.readFile(appStorage.get('excel-file-directory')).then(function(){
      var nameSheetNumber;
      var nameColumnNumber;
      var emailSheetNumber;
      var emailColumnNumber;
      var identifierSheetNumber;
      var identifierColumnNumber;
      var awardSheetNumber;
      var awardColumnNumber;

      workbook.eachSheet(function(worksheet, sheetId){
         worksheet.getRow(1).eachCell(function(cell, colNumber){
            if (cell.value == appStorage.get('name-column-selected')){
               nameColumnNumber = colNumber;
               nameSheetNumber = sheetId;
            }
            else if (cell.value == appStorage.get('email-column-selected')) {
               emailColumnNumber = colNumber;
               emailSheetNumber = sheetId;
            }
            else if (cell.value == appStorage.get('identifier-column-selected')) {
               identifierColumnNumber = colNumber;
               identifierSheetNumber = sheetId;
            }
            else if (cell.value == appStorage.get('award-column-selected')) {
               awardColumnNumber = colNumber;
               awardSheetNumber = sheetId;
            }
         });
      });

      const emailWorksheet = workbook.getWorksheet(emailSheetNumber);
      const nameWorksheet = workbook.getWorksheet(nameSheetNumber);
      const awardWorksheet = workbook.getWorksheet(awardSheetNumber);
      const identifierWorksheet = workbook.getWorksheet(identifierSheetNumber);
      var arrayOfCertificateDocuments = [];

      clientNames.forEach(function(clientName, arrayIndex){
         nameWorksheet.getColumn(nameColumnNumber).eachCell(function(cell, rowNumber){
            if ((cell.value != null) && (cell.value != appStorage.get('name-column-selected'))){
               if (cell.value == clientName){
                  var certificateDocumentSpecifications = {};
                  certificateDocumentSpecifications.name = cell.value;

                  if (emailWorksheet.getRow(rowNumber).getCell(emailColumnNumber).value != null){
                     var emailValue;
                     if(typeof emailWorksheet.getRow(rowNumber).getCell(emailColumnNumber).value != 'string'){
                        emailValue = emailWorksheet.getRow(rowNumber).getCell(emailColumnNumber).value.text;
                     }
                     else{
                        emailValue = emailWorksheet.getRow(rowNumber).getCell(emailColumnNumber).value;
                     }
                     certificateDocumentSpecifications.email = emailValue
                  }
                  else{
                     certificateDocumentSpecifications.email = "(email unavailable)";
                  }


                  if (awardWorksheet.getRow(rowNumber).getCell(awardColumnNumber).value != null){
                     certificateDocumentSpecifications.award =   awardWorksheet.getRow(rowNumber).getCell(awardColumnNumber).value;
                  }
                  else{
                     certificateDocumentSpecifications.award = "(award name unavailable)";
                  }

                  if (identifierWorksheet.getRow(rowNumber).getCell(identifierColumnNumber).value != null){
                     certificateDocumentSpecifications.identifier = identifierWorksheet.getRow(rowNumber).getCell(identifierColumnNumber).value;
                  }
                  else{
                     certificateDocumentSpecifications.identifier = "(certificate identifier unavailable)";
                  }

                  arrayOfCertificateDocuments.push(certificateDocumentSpecifications);
               }
            }
         });
      });

      arrayOfCertificateDocuments.forEach(function(specifications, index){
         makePDF(specifications.name, specifications.email, specifications.award, specifications.identifier);
      });
   });
}

const PDFDocument = require('pdfkit');
const moment = require('moment');
const fs = require('fs');

function makePDF(name, email, award, identifier){
   const date = new Date();
   var monthString;
   if ((date.getMonth() + 1) < 10){
      monthString = "0" + (date.getMonth() + 1).toString();
   }
   else{
      monthString = (date.getMonth() + 1).toString();
   }

   var dateString;
   if ((date.getDate() + 1) < 10){
      dateString = "0" + (date.getDate() + 1).toString();
   }
   else{
      dateString = (date.getDate() + 1).toString();
   }

   const fullIdentifier = date.getFullYear().toString() + monthString + dateString + "-" + identifier;

   const pdf = new PDFDocument({
      autoFirstPage: false
   });

   pdf.pipe(fs.createWriteStream("generated-content/" + fullIdentifier + ".pdf"));
   pdf.addPage({
    layout: "landscape"
   });

   pdf.registerFont('garamond', 'other_assets/pdf-generator-fonts/garamond.ttf');
   pdf.registerFont('century-gothic', 'other_assets/pdf-generator-fonts/century-gothic.ttf');

   pdf.image('other_assets/certificate-template/template-blank.jpg', 0, 0, {scale: 0.3275});
   pdf.font('garamond').fillColor("#414141").fontSize(36).text(name,13,238,{
      align: 'center'
   });
   pdf.fillColor("#414141").fontSize(22).text(award,13,369,{
      align: 'center'
   });
   pdf.fontSize(14).fillColor("#414141").text(moment().format('LL'), -265, 492,{
      align: 'center'
   });
   pdf.font('century-gothic').fillColor("#8C8C8C").fontSize(9).text(email,0,589.35,{
      align: 'right',
      width: 315,
      height: 50
   });
   pdf.fillColor("#8C8C8C").text(fullIdentifier,420,589.35,{
      align: 'left',
      width: 315,
      height: 50
   });
   pdf.end();
}

function searchAndRemoveFromArray(array, value){
   var indexOfValue;
   array.forEach(function(item, index){
      if (item == value){
         indexOfValue = index;
      }
   });
   array.splice(indexOfValue, 1);
}

function getSelectedClients(){
   var selectedClientNames = [];
   $(".name-selection .top #selected-client-names-list li").each(function(){
      selectedClientNames.push($(this).html());
   });
   return selectedClientNames;
}

function arrayContainsValue(array, value){
   var arrayDoesIndeedContainTheAforementionedValue = false;
   array.forEach(function(item, index){
      if (item == value){
         arrayDoesIndeedContainTheAforementionedValue = true;
         return;
      }
   });
   return arrayDoesIndeedContainTheAforementionedValue;
}
