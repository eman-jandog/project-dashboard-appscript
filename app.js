const WORKSHEET = PropertiesService.getScriptProperties().getProperty('dbsheet') 

class App {
    constructor() {
        this.ws = SpreadsheetApp.openById(WORKSHEET)
    }

    getProjectsData() {
      const sheets = this.ws.getSheets();
      const projects = [];
      for (const sheet of sheets) {
        const sheetName = sheet.getSheetName();
        const name = sheet.getRange('A1').getValue();
        const description = sheet.getRange('A2').getValue();
        const id = sheet.getRange('B4').getValue();
        const deadline = sheet.getRange('B5').getValue().toLocaleString('en-ph', { year: 'numeric', month: 'numeric', day: 'numeric'});
        const location = sheet.getRange('D4').getValue();
        const status = sheet.getRange('D5').getValue();

        const pt = sheet.getDataRange().getValues();
        const headers = [];
        for (const cell of pt[6]) {
          if (cell == "-") break
          else headers.push(cell);
        }

        const values = []
        for (let i=7; i<pt.length; i++) {
          const row = pt[i];
          if (!row[0]) continue;
          values.push(row.filter((item,i) => i < headers.length))
        }

        const sheetData = {
          id: id,
          name: name,
          description: description,
          deadline:  deadline,
          location: location,
          status: status,
          projectData: {
            headers: headers,
            values: values
          }
        }

        projects.push(sheetData);
      }

      return projects;
    }
}

function getDashboardData() {
    const app = new App();
    const data = app.getProjectsData();

    if (!data) throw new Error(`Resulting project data is empty.`);
    return {projects: data};
}


//Front-End Functions
function include(filename) {
    return HtmlService.createTemplateFromFile(filename).evaluate().getContent()
}

function doGet() {
    return HtmlService.createTemplateFromFile('templates/index')
        .evaluate()
        .setTitle("ARAC Project Dashboard")
        .addMetaTag("viewport", "width=device-width, initial-scale=1.0")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
}

//Custom Functions
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
