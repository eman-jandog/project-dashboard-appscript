const SN_DASHBOARD = "Dashboard"

class App {
    constructor() {
        this.ws = SpreadsheetApp.openById('1nJ77vTuO7kM8OFqCXL3l3PPfcRcXdpGWrU7nz-G-xWo').getSheetByName(SN_DASHBOARD)
    }

    getColumnName(i) {
        return this.ws.getRange(1, i + 1).getA1Notation().slice(0, -1)
    }

    getDashboarName() {
      return this.ws.getRange(1,2,1,1).getValues()[0];
    }

    getDashboardRefreshRate() {
      const rate = this.ws.getRange(1,4,1,1).getValues()[0];
      const duration = this.ws.getRange(2,4,1,1).getValues()[0];
      return { rate, duration }
    }

    createDatasets(values, colors, bgColors, datasetColumns) {
        datasetColumns = datasetColumns.split(",").map(v => v.trim().toUpperCase()).filter(v => v !== "")
        const [, ...headers] = values.shift()
        const datasets = []
        headers.forEach((h, i) => {
            const columnName = this.getColumnName(i + 1)
            if (datasetColumns.indexOf(columnName) !== -1) {
                const data = values.map(v => v[i+1]).filter(v => v !== '')
                const label = h
                if (bgColors[i] !== '#ffffff')
                  datasets.push({ label, data, backgroundColor: bgColors[i], borderColor: colors[i]})
                else datasets.push({ label, data})
            }
        })
        const labels = values.map(v => v[0]).filter(e => e !== "")
        return { labels, datasets }
    }

    createChartData([index, title, sheetName, datasetColumns]) {
        title = `${index}. ${title}`
        let type = "bar"
        const ss = SpreadsheetApp.getActiveSpreadsheet() //openByUrl(url)
        if (!ss) return null
        const ws = ss.getSheetByName(sheetName.trim())
        if (!ws) return null

        const dataRange = ws.getDataRange()
        const values = dataRange.getValues()
        const colors = dataRange.getFontColors()[0].slice(1)
        const bgColors = dataRange.getBackgrounds()[0].slice(1)

        const { labels, datasets } = this.createDatasets(values, colors, bgColors, datasetColumns)

        const chartData = {
            type,
            data: {
                labels,
                datasets,
            },
            options: {
              indexAxis: 'y',
              elements: {
                bar: {
                  borderWidth: 2,
                }
              },
              plugins: {
                title: {
                  display: title !== "",
                  text: title                    
                }
              }
            }
        }
        return chartData
    }

    getDashboardData() {
        const values = this.ws.getDataRange().getValues()
        values.splice(0,3)
        const charts = values.map(v => this.createChartData(v))
        return charts
    }
}

function include(filename) {
    return HtmlService.createTemplateFromFile(filename).evaluate().getContent()
}


function doGet() {
    return HtmlService.createTemplateFromFile('templates/index')
        .evaluate()
        .setTitle("Project Dashboard")
        .addMetaTag("viewport", "width=device-width, initial-scale=1.0")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
}

function getDashboardData() {
    const app = new App()
    const data = app.getDashboardData()
    const appname = app.getDashboarName();
    return JSON.stringify({ charts: data, appName: appname})
}

function getDashboardRefreshRate() {
    const app = new App()
    const { rate, duration }  = app.getDashboardRefreshRate();
    return JSON.stringify({ refresh: rate, transition: duration })
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
