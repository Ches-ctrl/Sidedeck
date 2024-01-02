const officegen = require('officegen')
const fs = require('fs')

// Create an empty PowerPoint object:
let pptx = officegen('pptx')

// Let's add a title slide:

let slide = pptx.makeTitleSlide('Officegen', 'Example to a PowerPoint document')

// Pie chart slide example:

slide = pptx.makeNewSlide()
slide.name = 'Pie Chart slide'
slide.back = 'ffff00'
slide.addChart(
  {
    title: 'My production',
    renderType: 'pie',
    data:
	[
      {
        name: 'Oil',
        labels: ['Czech Republic', 'Ireland', 'Germany', 'Australia', 'Austria', 'UK', 'Belgium'],
        values: [301, 201, 165, 139, 128,  99, 60],
        colors: ['ff0000', '00ff00', '0000ff', 'ffff00', 'ff00ff', '00ffff', '000000']
      }
    ]
  }
)

// Let's generate the PowerPoint document into a file:

return new Promise((resolve, reject) => {
  let out = fs.createWriteStream('example.pptx')

  // This one catch only the officegen errors:
  pptx.on('error', function(err) {
    reject(err)
  })

  // Catch fs errors:
  out.on('error', function(err) {
    reject(err)
  })

  // End event after creating the PowerPoint file:
  out.on('close', function() {
    resolve()
  })

  // This async method is working like a pipe - it'll generate the pptx data and put it into the output stream:
  pptx.generate(out)
})
