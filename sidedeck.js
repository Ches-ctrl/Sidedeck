const fs = require('fs');
const xlsx = require('xlsx');
const officegen = require('officegen');

const pptx = officegen('pptx');

const slide1 = pptx.makeNewSlide();
slide1.addText('Hello, PowerPoint!', { x: 'c', y: '2%', cx: '90%', font_size: 36, bold: true });

const slide2 = pptx.makeNewSlide();
slide2.addText('This is another slide', { x: 'c', y: '2%', cx: '90%', font_size: 36, bold: true });

const outputPptx = fs.createWriteStream('output_presentation2.pptx');
pptx.generate(outputPptx);

outputPptx.on('finalize', function () {
    console.log('Presentation saved as output_presentation2.pptx');
});

outputPptx.on('error', function (err) {
    console.log('Error creating presentation:', err);
});
