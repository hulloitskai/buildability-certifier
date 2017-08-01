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

   if (appStorage.has('certificate-input-date')){
      $(".preferences-input-section #date-input-field").val(appStorage.get('certificate-input-date'));
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
}

function preferencesAreValid(){
   if(moment($(".preferences-input-section #date-input-field").val()).isValid() == false){
      alert("Warning: Certificate date input is invalid. (Doesn't follow specified format).");
      return false;
   }
   else{
      return true;
   }
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

      $("#bottombar").css("background-color", "#EACD81");
      $("#bottombar p").css("color", "#998C48");

      statusStage = "start";
      arrayOfCertificateDocuments.forEach(function(specifications, index){
         makePDF(specifications.name, specifications.email, specifications.award, specifications.identifier);
      });
   });
}

var statusStage;

const PDFDocument = require('pdfkit');
const moment = require('moment');
const fs = require('fs');

function makePDF(name, email, award, identifier){
   console.log("Begin make-pdf for: " + identifier);
   const pdf = new PDFDocument({
      autoFirstPage: false
   });

   pdf.pipe(fs.createWriteStream("generated-content/" + identifier + ".pdf"));
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

emailjs = require('emailjs');
function sendCertificateEmail(clientName, emailAddress, identifier){
   console.log("Begin send-email for: " + identifier);

   const server= emailjs.server.connect({
      user: "testaccount@genuinebusiness.ca",
      password: "intracomuv1",
      host: "smtp.zoho.com",
      ssl: true,
      port: 465,
      timeout: 20000
   });

   var message	= {
      from: "<testaccount@genuinebusiness.ca>",
      to: clientName + " <" + emailAddress + ">",
      subject:	"Congratulations on completing your course at Buildability!",
      text: "Congratulations on completing your course at buildability. \nPlease find your certificate attached to this email.",
      attachment:
      [
         {data:"<html><h3>Congratulations!</h3><br /><p>Please find your certificate attached to this email.</p></html>", alternative:true},
         {path:"generated-content/" + identifier + ".pdf", type:"application/pdf", name: identifier +".pdf"}
      ]
   }

   server.send(message, function(err, message){
      if(err){
         console.log("send-email ERRORED for " + identifier);
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
            window.setTimeout(function(){
               $("#bottombar").css("background-color", "#37474F");
               $("#bottombar p").css("color", "#B0BEC5");
               $(".name-selection .top .lowbar input").prop('onclick',null).off('click');
               $("#select-all-client-names-cancel").hide().css("opacity","1");
            }, 2500);
            window.setTimeout(function(){
               $("#buildability-placeholder img").removeClass("transparent");
               $("#select-all-client-names-button").show().css("opacity","1");
            }, 3000);
            alert(err);
         }
      }
      else{
         console.log("email-send successful for " + identifier);
         $(".name-selection .top li").filter(function(){
            return $(this).text() === clientName;
         }).css("color","#BDEBC6");
         if (statusStage == "start"){
            statusStage = "end";
            $("#bottombar").css("background-color", "#BDEBC6");
            $("#bottombar p").css("color", "#78BF86");
            window.setTimeout(function(){
               $("#bottombar").css("background-color", "#37474F");
               $("#bottombar p").css("color", "#B0BEC5");
               $(".name-selection .top .lowbar input").prop('onclick',null).off('click');
               $("#select-all-client-names-cancel").hide().css("opacity","1");
            }, 2500);
            window.setTimeout(function(){
               $("#buildability-placeholder img").removeClass("transparent");
               $("#select-all-client-names-button").show().css("opacity","1");
            }, 3000);
         }
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
         updateExcelCompletionStatus(clientName);
      }
   });
}

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
   mainActivityButtonEnabled = false;
   $(".name-selection .bottom button").addClass("button-disabled");
   $(".name-selection .bottom button").prop('onclick',null).off('click');
}

function updateExcelCompletionStatus(clientName){
   const workbook = new Excel.Workbook();
   workbook.xlsx.readFile(appStorage.get('excel-file-directory')).then(function(){
      var nameSheetID;
      var nameColumnNumber;
      var completionRowNumber;
      var completionSheetNumber;;
      var completionColumnNumber;

      workbook.eachSheet(function(worksheet, sheetId){
         worksheet.getRow(1).eachCell(function(cell, colNumber){
            if (cell.value == appStorage.get('name-column-selected')){
               nameColumnNumber = colNumber;
               nameSheetID = worksheet.name;
               worksheet.getColumn(nameColumnNumber).eachCell(function(cell, rowNumber){
                  if (cell.value == clientName){
                     completionRowNumber = rowNumber;
                  }
               })
            }
            else if (cell.value == appStorage.get('completion-column-selected')) {
               completionColumnNumber = colNumber;
               completionSheetNumber = sheetId;
            }
         });
      });

      const targetRow = workbook.getWorksheet(completionSheetNumber).getRow(completionRowNumber)
      targetRow.getCell(completionColumnNumber).value = "true";
      targetRow.commit();
      return workbook.xlsx.writeFile(appStorage.get('excel-file-directory'));

   });
}
