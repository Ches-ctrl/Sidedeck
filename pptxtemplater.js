const fs = require('fs');
const officegen = require('officegen');
const Docxtemplater = require('docxtemplater');

// Create a new PowerPoint presentation (main presentation) using officegen
const mainPresentation = officegen('pptx');

// Load the source (template) presentation using officegen
const sourcePresentation = officegen('pptx');
sourcePresentation.load('path/to/source_presentation.pptx');

// Define the output file name for the main presentation
const outputFileName = 'path/to/output_presentation.pptx';

// Create a writable stream to save the main presentation
const outputStream = fs.createWriteStream(outputFileName);

// Pipe the source presentation to the main presentation
sourcePresentation.generate(outputStream);

outputStream.on('finish', () => {
  console.log('Main presentation with copied slides saved successfully.');
});

outputStream.on('error', (error) => {
  console.error('Error saving main presentation:', error);
});
