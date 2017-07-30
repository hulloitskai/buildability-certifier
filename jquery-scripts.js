const AppStorage = require('electron-store');
const Excel = require('exceljs');
const ipc = require('electron').ipcRenderer;
var appStorage = new AppStorage();

var preferencePaneActive = false;
function showPreferences(){
   $("#shadowcover").show().animate({
      opacity: '0.6'
   });
   $("#preferences-pane").show().delay(200).animate({
      top: '0'
   },{
      easing: 'easeOutElastic',
      duration: 700
   });
   delay(100).loadPreferences();
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
         $("#completion-column-form option").filter(function() {
            return $(this).text() === appStorage.get('completion-column-selected');
         }).prop('selected', true);
      }
   });
}
