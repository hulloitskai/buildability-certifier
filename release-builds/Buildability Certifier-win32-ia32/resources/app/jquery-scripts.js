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
   $(".preferences-pane").show().delay(300).animate({
      top: '0'
   },{
      easing: 'easeOutElastic',
      duration: 700,
      complete: function(){
         loadPreferences();
      }
   });
   $(".preferences-pane .body").delay(50).scrollTop(0).
   preferencePaneActive = true;
}

function hidePreferences(){
   $(".preferences-pane").animate({
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

var emailSendErrors;

function loadPreferences(){
   if (appStorage.has('excel-file-directory')){
      $("#selected-file-directory p").text(appStorage.get('excel-file-directory'));
      initializeColumnOptions();
   }
   else{
      $("#selected-file-directory p").text("no file selected");
      $(".dropdown-form select").replaceWith("<select><option>no file selected</option></select>");
   }

   if (appStorage.has('certificate-input-date')){
      $(".preferences-input-section #date-input-field").val(appStorage.get('certificate-input-date'));
   }
   if (appStorage.has('email-send-delay')){
      $("#email-send-delay input").val(appStorage.get('email-send-delay'));
   }
   if (appStorage.has('smtp-hostname')){
      $("#smtp-hostname input").val(appStorage.get('smtp-hostname'));
   }
   if (appStorage.has('sender-email-username')){
      $("#sender-email-username input").val(appStorage.get('sender-email-username'));
   }
   if (appStorage.has('sender-email-password')){
      $("#sender-email-password input").val(appStorage.get('sender-email-password'));
   }
   if (appStorage.has('backup-folder-size')){
      $("#backup-folder-size input").val(appStorage.get('backup-folder-size'));
   }
   if (appStorage.has('certificate-image-directory')){
      $("#certificate-image-directory p").text(appStorage.get('certificate-image-directory'));
   }
   if (appStorage.has('certificate-image-directory')){
      $("#certificate-image-directory p").text(appStorage.get('certificate-image-directory'));
   }
   if (appStorage.has('email-timeout')){
      $("#email-timeout input").val(appStorage.get('email-timeout'));
   }
   else{
      $("#certificate-image-directory p").text("no file selected");
   }
   if(appStorage.has('pdf-encryption-enabled')){
      $("#pdf-encryption-select option").filter(function() {
         return $(this).text() === appStorage.get('pdf-encryption-enabled');
      }).prop('selected', true);
   }
}

function savePreferences(){
   appStorage.set('excel-file-directory', $("#selected-file-directory p").text());
   appStorage.set('name-column-selected', $("#name-column-form select").find(":selected").text());
   appStorage.set('email-column-selected', $("#email-column-form select").find(":selected").text());
   appStorage.set('award-column-selected', $("#award-column-form select").find(":selected").text());
   appStorage.set('identifier-column-selected', $("#identifier-column-form select").find(":selected").text());
   appStorage.set('completion-column-selected', $("#completion-column-form select").find(":selected").text());
   if($(".preferences-input-section #date-input-field").val() != ''){
      appStorage.set('certificate-input-date', $(".preferences-input-section #date-input-field").val());
   }
   if($("#email-send-delay input").val() != ''){
      appStorage.set('email-send-delay', $("#email-send-delay input").val());
   }
   if($("#email-timeout input").val() != ''){
      appStorage.set('email-timeout', $("#email-timeout input").val());
   }
   if($("#sender-email-password input").val() != ''){
      appStorage.set('sender-email-password', $("#sender-email-password input").val());
   }
   if($("#sender-email-username input").val() != ''){
      appStorage.set('sender-email-username', $("#sender-email-username input").val());
   }
   if($("#smpt-hostname input").val() != ''){
      appStorage.set('smtp-hostname', $("#smtp-hostname input").val());
   }
   appStorage.set('backup-folder-size', $("#backup-folder-size input").val());
   appStorage.set('certificate-image-directory', $("#certificate-image-directory p").text());
   appStorage.set('pdf-encryption-enabled', $("#pdf-encryption-select").find(":selected").text());
}

function preferencesAreValid(){
   const name = $("#name-column-form select").find(":selected").text();
   const email = $("#email-column-form select").find(":selected").text();
   const award = $("#award-column-form select").find(":selected").text();
   const identifier = $("#identifier-column-form select").find(":selected").text();
   const completion = $("#completion-column-form select").find(":selected").text();

   if(moment($(".preferences-input-section #date-input-field").val()).isValid() == false){
      alert("Warning: Certificate date input is invalid. (Doesn't follow specified format).");
      return false;
   }
   else if (isNaN($("#backup-folder-size input").val())) {
      alert("Warning: Backup folder size is not a number.");
      return false;
   }
   else if (
      (name == email) ||
      (name == award) ||
      (name == identifier) ||
      (name == completion) ||
      (email == award) ||
      (email == identifier) ||
      (email == completion) ||
      (award == identifier) ||
      (award == completion) ||
      (identifier == completion)
   ){
      alert("Warning: Two of the selected columns are the same.");
      return false;
   }
   else{
      return true;
   }
}

function openExcelFileSelector(){
   $("#selected-file-directory p").text("no file selected");
   $(".dropdown-form select").replaceWith("<select><option>no file selected</option></select>");
   ipc.send('open-file-dialog');
   ipc.on('selected-directory', function (event, path) {
      $("#selected-file-directory p").text(path);
      initializeColumnOptions();
   })
}

function openCertificateImageFileSelector(){
   $("#certificate-image-directory p").text("no file selected");
   ipc.send('open-image-dialog');
   ipc.on('selected-image-directory', function (event, path) {
      $("#certificate-image-directory p").text(path);
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
   if ((excelFileDirectory != "no file selected") && (excelFileDirectory != '')){
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
         if (getStringFromCell(cell) != appStorage.get('name-column-selected')){
            if (getStringFromCell(cell) != null){
               arrayOfNames.push(getStringFromCell(cell));
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
   disableMainActionButton();
    $(".name-selection .top .lowbar input").click(function(event) {
       event.preventDefault();
    });

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
            if ((getStringFromCell(cell) != null) && (getStringFromCell(cell) != appStorage.get('name-column-selected'))){
               if (getStringFromCell(cell) == clientName){
                  var certificateDocumentSpecifications = {};
                  certificateDocumentSpecifications.name = getStringFromCell(cell);

                  if (emailWorksheet.getRow(rowNumber).getCell(emailColumnNumber).value != null){
                     certificateDocumentSpecifications.email = getStringFromCell(emailWorksheet.getRow(rowNumber).getCell(emailColumnNumber));
                  }
                  else{
                     certificateDocumentSpecifications.email = "(email unavailable)";
                  }

                  if (awardWorksheet.getRow(rowNumber).getCell(awardColumnNumber).value != null){
                     certificateDocumentSpecifications.award =   getStringFromCell(awardWorksheet.getRow(rowNumber).getCell(awardColumnNumber));
                  }
                  else{
                     certificateDocumentSpecifications.award = "(award name unavailable)";
                  }

                  if (identifierWorksheet.getRow(rowNumber).getCell(identifierColumnNumber).value != null){
                     certificateDocumentSpecifications.identifier = getStringFromCell(identifierWorksheet.getRow(rowNumber).getCell(identifierColumnNumber));
                  }
                  else{
                     certificateDocumentSpecifications.identifier = "(certificate identifier unavailable)";
                  }

                  arrayOfCertificateDocuments.push(certificateDocumentSpecifications);
               }
            }
         });
      });

      $("#bottombar").css("background-color", "#EACD81");
      $("#bottombar p").css("color", "#998C48");

      statusStage = "start";
      let loopCounter = 1;
      emailDoneCounter = 0;
      clientsThatErroredOut = [];
      emailSendErrors = [];
      anticipatedEmailCount = arrayOfCertificateDocuments.length;
      successfulEmailSendClients = [];
      certificateSpecifications = arrayOfCertificateDocuments;
      emailGrandmaster();
   });
}

function getStringFromCell(cell){
   if(typeof cell.value == 'object'){
      return cell.text;
   }
   else{
      return cell.value;
   }
}

var certificateSpecifications;
function emailGrandmaster(){
   let grandmasterLoopCounter = 1;
   let emailsPerBatch = 5;
   if(appStorage.has('email-send-delay') && (appStorage.get('email-send-delay') != '')){
      emailsPerBatch = parseInt(appStorage.get('email-send-delay'));
   }

   certificateSpecifications.forEach((specifications, index)=>{
      if (grandmasterLoopCounter <= emailsPerBatch){
         makePDF(specifications.name, specifications.email, specifications.award, specifications.identifier);
      }
      grandmasterLoopCounter += 1;
   });

   if (certificateSpecifications.length >= emailsPerBatch){
      certificateSpecifications.splice(0,emailsPerBatch);
   }
   else{
      certificateSpecifications = [];
   }
}

var statusStage;
const PDFDocument = require('pdfkit');
const moment = require('moment');
const fs = require('fs');
const qpdf = require('node-qpdf');

function pathlink(staticpath){
   return path.join(__dirname, staticpath);
}

function makePDF(name, email, award, identifier){
   console.log("Begin make-pdf for: " + identifier);
   const pdf = new PDFDocument({
      autoFirstPage: false
   });

   pdf.pipe(fs.createWriteStream(pathlink("generated-content/" + identifier + ".pdf")));
   pdf.addPage({
    layout: "landscape"
   });

   pdf.registerFont('garamond', pathlink('other_assets/pdf-generator-fonts/garamond.ttf'));
   pdf.registerFont('century-gothic', pathlink('./other_assets/pdf-generator-fonts/century-gothic.ttf'));

   let certificateImage = pathlink('other_assets/certificate-template/template-blank.jpg');
   if (appStorage.has('certificate-image-directory') && (appStorage.get('certificate-image-directory') != 'no file selected')){
      certificateImage = appStorage.get('certificate-image-directory');
   }

   pdf.image(certificateImage, 0, 0, {width: 792});
   pdf.font('garamond').fillColor("#414141").fontSize(36).text(name,13,238,{
      align: 'center'
   });
   pdf.fillColor("#414141").fontSize(22).text(award,13,369,{
      align: 'center'
   });

   const certificateDate = moment(appStorage.get('certificate-input-date'));
   pdf.fontSize(14).fillColor("#414141").text(certificateDate.format('LL'), -265, 492,{
      align: 'center'
   });

   pdf.font('century-gothic').fillColor("#8C8C8C").fontSize(9).text("",0,589.35,{
      align: 'right',
      width: 315,
      height: 50
   });

   const fullIdentifier = certificateDate.format("YYYYMMDD") + "-" + identifier;
   pdf.fillColor("#8C8C8C").text(fullIdentifier,420,589.35,{
      align: 'left',
      width: 315,
      height: 50
   });
   pdf.end();

   if (appStorage.has('pdf-encryption-enabled') && appStorage.get('pdf-encryption-enabled') == 'true'){
      qpdf.encrypt(pathlink('generated-content/' + identifier + '.pdf'), {
         keyLength: 256,
         password: 'buildabilitysecretkey',
         restrictions: {
           modify: 'none',
           print: 'full',
           extract: 'n',
           accessibility: 'y'
         }
      });
   }

   console.log("Done make-pdf for: " + identifier);
   sendCertificateEmail(name, email, identifier);
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

var clientsThatErroredOut;

emailjs = require('emailjs');
function sendCertificateEmail(clientName, emailAddress, identifier){
   console.log("Begin send-email for: " + identifier);

   let emailTimeout = 20000

   if (appStorage.has('email-timeout')){
      emailTimeout = parseInt(parseFloat(appStorage.get('email-timeout')) * 1000);
      console.log("Email timeout set to " + emailTimeout)
   }

   const server= emailjs.server.connect({
      user: appStorage.get('sender-email-username'),
      password: appStorage.get('sender-email-password'),
      host: appStorage.get('smtp-hostname'),
      ssl: true,
      port: 465,
      timeout: emailTimeout
   });

   var message	= {
      from: "<" + appStorage.get('sender-email-username') + ">",
      to: clientName + " <" + emailAddress + ">",
      subject:	"Congratulations on completing your course at Buildability!",
      text: "Congratulations on completing your course at buildability. \nPlease find your certificate attached to this email.",
      attachment:
      [
         {data:"<html><h3>Congratulations!</h3><br /><p>Please find your certificate attached to this email.</p></html>", alternative:true},
         {path: pathlink("generated-content/" + identifier + ".pdf"), type:"application/pdf", name: identifier +".pdf"}
      ]
   }

   server.send(message, function(err, message){
      emailDoneCounter += 1;
      if(err){
         console.log("send-email ERRORED for " + identifier);
         clientsThatErroredOut.push(clientName);
         $(".name-selection .top li").css("transition-duration","1s")
         $(".name-selection .top li").filter(function(){
            return $(this).text() === clientName;
         }).css("color","#E07F7F");
         window.setTimeout(function(){
            $(".name-selection .top li").filter(function(){
               return $(this).text() === clientName;
            }).css("opacity","0");
         }, 2500);
         window.setTimeout(function(){
            $(".name-selection .top li").filter(function(){
               return $(this).text() === clientName;
            }).remove();
         }, 3000);
         if (statusStage == "start"){
            statusStage = "end";
            $("#bottombar").css("background-color", "#E07F7F");
            $("#bottombar p").css("color", "#966565");
         }
         if(emailDoneCounter == anticipatedEmailCount){
            window.setTimeout(function(){
               $("#bottombar").css("background-color", "#37474F");
               $("#bottombar p").css("color", "#B0BEC5");
               $(".name-selection .top .lowbar input").prop('onclick',null).off('click');
               $("#select-all-client-names-cancel").hide().css("opacity","1");
            }, 2500);
            window.setTimeout(function(){
               $("#buildability-placeholder img").removeClass("transparent");
               $("#select-all-client-names-button").show().css("opacity","1");
               attemptRedoEmails();
            }, 3000);
         }
         if(arrayContainsValue(emailSendErrors, err.toString()) == false){
            alert(err);
            emailSendErrors.push(err.toString());
         }
      }
      else{
         console.log("email-send successful for " + identifier);
         $(".name-selection .top li").filter(function(){
            return $(this).text() === clientName;
         }).css("color","#BDEBC6");
         window.setTimeout(function(){
            $(".name-selection .top li").filter(function(){
               return $(this).text() === clientName;
            }).css("opacity","0");
         }, 2500);
         window.setTimeout(function(){
            $(".name-selection .top li").filter(function(){
               return $(this).text() === clientName;
            }).remove();
         }, 3000);
         if (statusStage == "start"){
            statusStage = "end";
            $("#bottombar").css("background-color", "#BDEBC6");
            $("#bottombar p").css("color", "#78BF86");
         }
         if(emailDoneCounter == anticipatedEmailCount){
            window.setTimeout(function(){
               $("#bottombar").css("background-color", "#37474F");
               $("#bottombar p").css("color", "#B0BEC5");
               $(".name-selection .top .lowbar input").prop('onclick',null).off('click');
               $("#select-all-client-names-cancel").hide().css("opacity","1");
            }, 2500);
            window.setTimeout(function(){
               $("#buildability-placeholder img").removeClass("transparent");
               $("#select-all-client-names-button").show().css("opacity","1");
               attemptRedoEmails();
            }, 3000);
         }
         successfulEmailSendClients.push(clientName);
      }

      if (fs.existsSync(pathlink('generated-content/' + identifier + '.pdf'))) {
         fs.unlink(pathlink('generated-content/' + identifier + '.pdf'), function(error){
            if(error){
               alert("A non-critical error occurred (pdf-cache didn't clear). Please notify the developer.")
               console.log("An error occued while attempting to remove PDF file: \n" + error);
            }
         });
      }

      if (emailDoneCounter == anticipatedEmailCount){
         if(successfulEmailSendClients != ''){
            console.log("Updating Excel files for " + successfulEmailSendClients);
            updateExcelCompletionStatus(successfulEmailSendClients);
         }
         else{
            console.log("Excel file will not be updated: No user emails sucessfully sent.")
         }
      }

      let emailsPerBatch = 5;
      if(appStorage.has('email-send-delay') && (appStorage.get('email-send-delay') != '')){
         emailsPerBatch = parseInt(appStorage.get('email-send-delay'));
      }

      if ((emailDoneCounter % emailsPerBatch == 0) && (certificateSpecifications.length > 0)){
         emailGrandmaster();
      }
   });
}

function attemptRedoEmails(){
   if((clientsThatErroredOut != '') && (emailDoneCounter == anticipatedEmailCount)){
      window.setTimeout(function(){
         if(confirm("It appears some emails encountered an error during the last run. Would you like to retry those emails?")){
            console.log("After errorying-out, retrying email-send for " + clientsThatErroredOut);
            clientsThatErroredOut.forEach((clientName, index)=>{
               $(".name-selection .top ol").append("<li>" + clientName + "</li>");
            });
            if (mainActivityButtonEnabled == false){
               $("#buildability-placeholder img").addClass("transparent");
               certifyClients();
            }
         }
      },300);
   }
}

const path = require('path');
var successfulEmailSendClients;

var emailDoneCounter;

var mainActivityButtonEnabled = false;

function enableMainActionButton(){
   mainActivityButtonEnabled = true;
   $(".name-selection .bottom button").removeClass("button-disabled");
   $(".name-selection .bottom button").click(function(){
      $(".name-selection .top li").css("color","#A3CEF2");
      $(".name-selection .top li").css("cursor","default");
      $(".name-selection .top li").css("transition-duration","0.5s");
      $(".name-selection .top li").prop('onclick',null).off('click');
      $("#select-all-client-names-cancel").css("opacity","0");
      $("#select-all-client-names-button").css("opacity","0");
      certifyClients();
   });
}

function disableMainActionButton(){
   if (mainActivityButtonEnabled == true){
      mainActivityButtonEnabled = false;
      $(".name-selection .bottom button").addClass("button-disabled");
      $(".name-selection .bottom button").prop('onclick',null).off('click');
   }
}

function updateExcelCompletionStatus(clientNames){
   backupExcelFile();
   const workbook = new Excel.Workbook();
   workbook.xlsx.readFile(appStorage.get('excel-file-directory')).then(function(){
      console.log("Excel file opened.");
      clientNames.forEach(function(clientName, index){
         var nameSheetID;
         var nameColumnNumber;
         var completionRowNumber;
         var completionSheetNumber;
         var completionColumnNumber;

         workbook.eachSheet(function(worksheet, sheetId){
            worksheet.getRow(1).eachCell(function(cell, colNumber){
               if (cell.value == appStorage.get('name-column-selected')){
                  nameColumnNumber = colNumber;
                  nameSheetID = worksheet.name;
                  worksheet.getColumn(nameColumnNumber).eachCell(function(cell, rowNumber){
                     if (getStringFromCell(cell) == clientName){
                        completionRowNumber = rowNumber;
                     }
                  })
               }
               else if (cell.value == appStorage.get('completion-column-selected')) {
                  completionColumnNumber = colNumber;
                  completionSheetNumber = sheetId;
               }
            });
         })
         var targetRow = workbook.getWorksheet(completionSheetNumber).getRow(completionRowNumber);
         targetRow.getCell(completionColumnNumber).value = "true";
         targetRow.commit();
         console.log("Updated Excel row for: " + clientName + ".");
      });
      console.log("Finished writing Excel file.");
      return workbook.xlsx.writeFile(appStorage.get('excel-file-directory'));
   });
}

var anticipatedEmailCount;

function checkProgramRunRequirements(){
   if(
      (appStorage.has('excel-file-directory') && (appStorage.get('excel-file-directory') != "no file selected"))
      &&
      (appStorage.has('certificate-input-date') && (appStorage.get('certificate-input-date') != ''))
      &&
      (appStorage.has('sender-email-username') && (appStorage.get('sender-email-username') != ''))
      &&
      (appStorage.has('sender-email-password') && (appStorage.get('sender-email-password') != ''))
      &&
      (appStorage.has('smtp-hostname') && (appStorage.get('smtp-hostname') != ''))
   ){
      return true;
   }
   else{
      return false;
   }
}

function backupExcelFile(){
   let files = fs.readdirSync(pathlink('userdata-backups'));
   let largestFileID = 0;
   files.forEach((file) => {

      //console.log("Filename: " + file);
      //console.log("x-pos" + file.search("x"));

      let fileID = parseInt(file.substring(0,file.search("x")));
      if (fileID > largestFileID){
         largestFileID = fileID;
      }
   });
   const newFileID = largestFileID += 1;
   fs.createReadStream(appStorage.get('excel-file-directory')).pipe(fs.createWriteStream(pathlink('userdata-backups/' + newFileID + '.xlsx')));

   let maxBackAmount = 25;
   if (appStorage.has('backup-folder-size') && (appStorage.get('backup-folder-size') != '')){
      maxBackAmount = appStorage.get('backup-folder-size');
   }

   if (newFileID > maxBackAmount){
      let itemsToSetBack = newFileID - maxBackAmount;
      console.log("Maximum Excel backup size reached. Removing " + itemsToSetBack + " items.")

      for(i = 1; i <= itemsToSetBack; i++){
         fs.unlinkSync(pathlink('userdata-backups/' + i + '.xlsx'));
         console.log("Deleting " + i + ".xlsx");
      }

      files = fs.readdirSync(pathlink('userdata-backups'));

      for(var i = itemsToSetBack + 1; i <= (newFileID); i ++) {
         fs.renameSync(pathlink('userdata-backups/' + i + '.xlsx'), pathlink('userdata-backups/' + (i - itemsToSetBack) + '.xlsx'))
      }
      console.log("Excel backup files renamed.");
   }

   console.log("Excel file backed up.");
}
