const fs = require('fs');
const xlsx = require('xlsx');
const officegen = require('officegen');

// Load data from Excel file
// const excelFile = 'test_deck_builder.xlsx';
// console.log('Loading data from ' + excelFile);

// const workbook = xlsx.readFile(excelFile);
// console.log('Loaded ' + workbook.SheetNames.length + ' sheets');

// console.log('Sheet names:');
// workbook.SheetNames.forEach(sheetName => {
//   console.log(sheetName);
//   console.log(typeof sheetName);
// });

// const worksheet = workbook.Sheets[`Sheet1`];
// console.log(worksheet);

// const excelData = xlsx.utils.sheet_to_json(worksheet);
// console.log('Loaded ' + excelData.length + ' rows');

// Create a PowerPoint presentation
const pptx = officegen('pptx');
// console.log('Created PowerPoint presentation');

const slide1 = pptx.makeNewSlide();

slide1.addText('Hello, PowerPoint!', { x: 'c', y: '2%', cx: '90%', font_size: 36, bold: true });

// Create another slide
const slide2 = pptx.makeNewSlide();

// Add content to the second slide
slide2.addText('This is another slide', { x: 'c', y: '2%', cx: '90%', font_size: 36, bold: true });

// Save the PowerPoint presentation
const outputPptx = fs.createWriteStream('output_presentation.pptx');
pptx.generate(outputPptx);

outputPptx.on('finalize', function () {
    console.log('Presentation saved as output_presentation.pptx');
});

outputPptx.on('error', function (err) {
    console.log('Error creating presentation:', err);
});
