const fs = require('fs');
const xlsx = require('xlsx');
const officegen = require('officegen');

// Load data from Excel file
const excelFile = 'test_deck_builder.xlsx';
const workbook = xlsx.readFile(excelFile);
const worksheet = workbook.Sheets['Sheet1'];
const excelData = xlsx.utils.sheet_to_json(worksheet);

// Create a PowerPoint presentation
const pptx = officegen('pptx');
const slides = pptx.slides;

// Iterate through Excel data and create slides
for (const row of excelData) {
    const slideTitle = row['Slide Title'];
    const templateName = row['Template Name'];

    // Add a slide with a title and content layout
    const slide = slides.addSlide();

    // Set slide title
    const title = slide.addText(slideTitle, { x: 'c', y: '2%', cx: '90%', font_size: 36, bold: true });

    // Add content (you can customize this part)
    const content = slide.addText('This is the content for ' + slideTitle, { x: 'c', y: '20%', cx: '80%', font_size: 20 });

    // Apply any additional formatting or template-specific customization here
}

// Save the PowerPoint presentation
const outputPptx = fs.createWriteStream('output_presentation.pptx');
pptx.generate(outputPptx);

outputPptx.on('finalize', function () {
    console.log('Presentation saved as output_presentation.pptx');
});

outputPptx.on('error', function (err) {
    console.log('Error creating presentation:', err);
});
