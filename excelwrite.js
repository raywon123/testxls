var excel = require('exceljs');
var workbook = new excel.Workbook(); //creating workbook
var sheet = workbook.addWorksheet('MySheet'); //creating worksheet

var image = workbook.addImage({
    filename : './pngImage.png',
    extension : 'png'
}); // adding an image in workbook first
sheet.addImage(image, 'D10:G14'); // adding an image in the worksheet in particular place

var objArray =[{
    "id" : 0,
    "name" : "xxxx",
    "is_active" : "false"
},{
    "id" : 1,
    "name" : "yyyy",
    "is_active" : "true"
}]
sheet.addRow().values = Object.keys(objArray[0]);

objArray.forEach(function(item){
    var valueArray = [];
    valueArray = Object.values(item); // forming an array of values of single json in an array
    sheet.addRow().values = valueArray; // add the array as a row in sheet
})

workbook.xlsx.writeFile('./temp.xlsx').then(function() {
    console.log("file is written")
})

// -- sending to client
// var tempfile = require('tempfile');
// var tempFilePath = tempfile('.xlsx');
// console.log("tempFilePath : ", tempFilePath);
// workbook.xlsx.writeFile(tempFilePath).then(function() {
//     res.sendFile(tempFilePath, function(err){
//         console.log('---------- error downloading file: ', err);
//     });
//     console.log('file is written');
// });