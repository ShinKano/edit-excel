const fs = require('fs');

// カレントディレクトリ内の.xlsxファイル全てを配列で取得
const getFileList = () => {

  const allFiles = fs.readdirSync('.');
  const xlsxFileList = allFiles.filter(function(file){
    return fs.statSync(file).isFile() && /.*\.xlsx$/.test(file); //絞り込み
  });

  return xlsxFileList;
}

module.exports = getFileList;


// fs.readdirSync('.', function(err, files){
//   if (err) throw err;
//   let fileList = files.filter(function(file){
//       return fs.statSync(file).isFile() && /.*\.xlsx$/.test(file); //絞り込み
//   })
//   console.log(fileList);
//   return fileList;
// });