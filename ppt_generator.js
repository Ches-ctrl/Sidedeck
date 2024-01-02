const fs = require('fs');
const officegen = require('officegen');

const excelFilePath = 'test_deck_builder.xlsx';

const projectName = 'MyProject';
const version = 'v0';

// Create a PowerPoint document.
const pptx = officegen('pptx');

// Create a new PowerPoint slide.
const slide = pptx.makeNewSlide();

// Read the Excel file content and extract slide titles.
// You'll need to implement your logic for reading and extracting titles.

// For this example, let's assume you have an array of slide titles.
const slideTitles = ['Title 1', 'Title 2', 'Title 3'];

// Add the extracted slide titles to PowerPoint slides.
slideTitles.forEach((title) => {
  const pptxSlide = pptx.makeNewSlide();
  pptxSlide.addText(title, {
    x: 'c',
    y: 'c',
    font_face: 'Arial',
    font_size: 32,
  });
});

// Define the output file name.
const date = new Date().toISOString().split('T')[0];
const outputFileName = `${date}_${projectName}_${version}.pptx`;

// Create a writable stream to save the PowerPoint file.
const outputStream = fs.createWriteStream(outputFileName);

// Pipe the PowerPoint document to the output stream.
pptx.generate(outputStream);

outputStream.on('finish', () => {
  console.log(`PowerPoint presentation "${outputFileName}" has been created.`);
});

// Handle any errors during file generation.
outputStream.on('error', (err) => {
  console.error('Error generating PowerPoint presentation:', err);
});
