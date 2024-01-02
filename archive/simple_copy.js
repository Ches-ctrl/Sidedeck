const fs = require('fs');
const officegen = require('officegen');

// Create a new PowerPoint presentation (main presentation)
const mainPptx = officegen('pptx');

// Load the source (template) presentation from which you want to copy slides
const sourcePptx = officegen('pptx');
sourcePptx.load('path/to/source_presentation.pptx');

// Iterate through the slides in the source presentation and copy them to the main presentation
sourcePptx.slides.forEach((sourceSlide) => {
  // Copy the slide to the main presentation
  mainPptx.slides.push(sourceSlide);
});

// Define the output file name for the main presentation
const outputFileName = 'path/to/output_presentation.pptx';

// Create a writable stream to save the main presentation
const outputStream = fs.createWriteStream(outputFileName);

// Pipe the main presentation to the output stream
mainPptx.generate(outputStream);

outputStream.on('finish', () => {
  console.log('Main presentation with copied slides saved successfully.');
});

outputStream.on('error', (error) => {
  console.error('Error saving main presentation:', error);
});
