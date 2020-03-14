#!/usr/bin/env node
/*
node.jsのスクリプトをコマンドラインツールとして動作させる
 */

const getFileList = require('./getFileList');
const readlineSync = require('readline-sync');
const XlsxPopulate = require('xlsx-populate');


function editExcel(){

  const fileList = getFileList();

  console.log("以下の.xlsxファイルが変更の対象です...")
  console.log(fileList);

  if (readlineSync.keyInYN('変更するファイルは閉じてください...よろしいです?')) {
    const rowNum = readlineSync.questionInt('変更するセルの【行番号】を数字で入力してください：');
    const colNum = readlineSync.questionInt('変更するセルの【行番号】を数字で入力してください：');
    const changedValue = readlineSync.question('セルに入力する値を入力してください：');
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
        console.log(`***完了しました🐒 : ${file}***`)
        return workbook.toFileAsync(file);
      });
    
    })
    // Do something...
  } else {
    // Another key was pressed.
    console.log('Bye Bye...');
  }

}
editExcel();
module.exports = editExcel;
