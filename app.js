const SN_DASHBOARD = "Overview"

class App {
    constructor() {
        this.ws = SpreadsheetApp.openById('1nJ77vTuO7kM8OFqCXL3l3PPfcRcXdpGWrU7nz-G-xWo').getSheetByName(SN_DASHBOARD);
    }

    getDashboardData() {
        const data = this.ws.getDataRange().getValues();
        const headers = data.splice(0,1)[0];

        // create a json file
        return data.map(row => {
          const project = {}
          headers.forEach((key, i) => {
            let value = row[i];
            key = key.toLowerCase();

            if (key.includes(' ')) {
              key = key.split(' ').join('')
            }

            if (value instanceof Date) {
              value = value.toString()
            }

            project[key] = value
          })
          return project;
        })
    }
}

function getDashboardData() {
    const app = new App()
    const data = app.getDashboardData()
    return {projects: data}
}


function include(filename) {
    return HtmlService.createTemplateFromFile(filename).evaluate().getContent()
}

function doGet() {
    return HtmlService.createTemplateFromFile('templates/main')
        .evaluate()
        .setTitle("ARAC Project Dashboard")
        .addMetaTag("viewport", "width=device-width, initial-scale=1.0")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
}

//CUSTOM SHEET UIs
function report() {
  var html = HtmlService.createHtmlOutputFromFile('templates/report_dialog')
    .setWidth(400) // Set dialog width
    .setHeight(300); // Set dialog height
  SpreadsheetApp.getUi().showModalDialog(html, 'Project Report');
}

//CUSTOM FORMULAS
function countBlocks(countRange,colorRef) {
  var activeRange = SpreadsheetApp.getActiveRange()
  var activeSheet = SpreadsheetApp.getActiveSheet()

  var activeformula = activeRange.getFormula();
  var countRangeAddress = activeformula.match(/\((.*)\,/).pop().trim();
  var colorRefAddress = activeformula.match(/\,(.*)\)/).pop().trim();

  var backgrounds = activeSheet.getRange(countRangeAddress).getBackgrounds()
  var BackGround = activeSheet.getRange(colorRefAddress).getBackground();

  var countColorCells = 0;

  for (var i = 0; i < backgrounds.length; i++)

    for (var k = 0; k < backgrounds[i].length; k++)

      if ( backgrounds[i][k] == BackGround )
        countColorCells++;

  return countColorCells;

}
