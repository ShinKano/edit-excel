const getFileList = require('./getFileList');
const readlineSync = require('readline-sync');
const XlsxPopulate = require('xlsx-populate');


function editExcel(){

  const fileList = getFileList();

  console.log("These are your .xlsx files to be edited...")
  console.log(fileList);

  if (readlineSync.keyInYN('Please close the files to be edited...OK?')) {
    const rowNum = readlineSync.questionInt('What is the ROW-NUMBER of the cell to be edited?');
    const colNum = readlineSync.questionInt('What is the COLUMN-NUMBER of the cell to be edited?');
    const changedValue = readlineSync.question('What is the VALUE to be put in the cell?');
    // 'Y' key was pressed.
    fileList.forEach(file => {
      // Load an existing workbook
      XlsxPopulate.fromFileAsync(file)
      .then(workbook => {
        // Get all sheets as an array
        const sheets = workbook.sheets();
    
        sheets.forEach((sheet, i) => {
          let targetCell = sheet.row(rowNum).cell(colNum);
          targetCell.value(changedValue);
        })
        
        //write and save.
        console.log(`***Done : ${file}***`)
        return workbook.toFileAsync(file);
      });
    
    })
    // Do something...
  } else {
    // Another key was pressed.
    console.log('Bye Bye...');
  }

}
module.exports = editExcel;






