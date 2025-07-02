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

        if (sheetName.match(/draft/)) continue;

        const name = sheet.getRange('A1').getValue();
        const description = sheet.getRange('A2').getValue();
        const id = sheet.getRange('B4').getValue();
        const deadline = sheet.getRange('B5').getValue().toLocaleString('en-ph', { year: 'numeric', month: 'numeric', day: 'numeric'});
        const location = sheet.getRange('D4').getValue();
        const status = sheet.getRange('D5').getValue();

        if (!name || !description || !id || !location || !status) {
          throw new Error(`Missing information in sheet ${sheetName}`);
        }

        const pt = sheet.getDataRange().getValues();
        const headers = [];
        const values = [];
        
        const _headers = pt[6];
        for (let r=0; r<_headers.length; r++) {
          if (_headers[r] == "-") break;
          if (!_headers[r]) continue;
          
          headers.push(_headers[r]);
          let skip = 0;
          for (let i=7; i<pt.length; i++) {
            const row = pt[i];
            if (!row[0]) {
              skip++;
              continue;
            }
            values[i-7-skip] ? values[i-7-skip].push(row[r]) : values[i-7-skip] = [row[r]];
          }           
        }

        const sheetData = {
          id: id,
          name: name,
          description: description,
          deadline:  deadline,
          location: location,
          status: status,
          table: {
            headers: headers,
            values: values
          }
        }

        console.log(sheetData.table.values)
        projects.push(sheetData);
      }

      return projects;
    }
}

function getDashboardData() {
  const app = new App();
  try {
    const data = app.getProjectsData();
    if (!data) throw new Error(`Resulting project data is empty.`);
    return {projects: data};
  } catch (e) {
    throw new Error(e.message);
  }
}

//Front-End Functions
function include(filename) {
    return HtmlService.createTemplateFromFile(filename).evaluate().getContent()
}

// Default Appscript Functions
// When Page load get the html file
function doGet() {
    return HtmlService.createTemplateFromFile('templates/index')
        .evaluate()
        .setTitle("ARAC Project Dashboard")
        .addMetaTag("viewport", "width=device-width, initial-scale=1.0")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
}

//Custom UI Custom Menu
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('ARAC')
    .addItem('Report', 'report')
    .addToUi();
}

//Custom Functions
function report() {
  const deploymentId = PropertiesService.getScriptProperties().getProperty('deploymentId');
  const webAppUrl = `https://script.google.com/macros/s/${deploymentId}/exec`;
  const template = HtmlService.createTemplateFromFile('templates/report_dialog');
  template.webAppUrl = webAppUrl;

  const html = template.evaluate()
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