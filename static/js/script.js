const app = document.getElementById("app")
app.innerHTML = `
<div class="text-muted">
<div class="spinner-grow spinner-grow-sm text-muted" role="status">
</div>
<span>Loading...</span>
</div>
`

let refreshMins = 10
let transitionMins = 60000
let active = 0;
let transitionStat = true;
let savePos = "0px";

let refreshInterval;
let tableContainer;
let items;
let dots;
let lengthItems;

getData()
load()


function reload() {
  document.getElementById("refresh-time").innerHTML = `
  <div class="text-muted">
  <div class="spinner-grow spinner-grow-sm text-muted" role="status">
  </div>
  <span>Refreshing...</span>
  </div>
  `
  load();
}


function getData(){
  google.script.run
  .withSuccessHandler( (data) => {
  const { refresh, transition } = JSON.parse(data)
  if (refresh) refreshMins = refresh[0]
  if (transition) transitionMins = transition[0] * 1000
  setInterval(reload, refreshMins * 60 * 1000)
  })
  .withFailureHandler(e => app.innerHTML = `<p class="error">${e.message}</p>`)
  .getDashboardRefreshRate()
}


function load(){
  google.script.run
  .withSuccessHandler( (data) => {
  const { charts, appName } = JSON.parse(data)
  createCharts(charts, appName)
  transition();
  Chart.defaults.font.size = 16;
  })
  .withFailureHandler(e => app.innerHTML = `<p class="error">${e.message}</p>`)
  .getDashboardData()
}


function createCharts(charts, appName){
  const chartDivs = charts.map(chartData => {
    const chartWrapper = document.createElement("DIV")
    chartWrapper.className = 'slider'

    const chartDiv = document.createElement("DIV")
    chartDiv.className = "ctn-chart"

    const chartCanvas = document.createElement("CANVAS")
    const ctx = chartCanvas.getContext('2d')
    const chartObj = new Chart(ctx, chartData)
    
    
    chartDiv.appendChild(chartCanvas)

    const panelsDiv = document.createElement("DIV")
    panelsDiv.className = "panels"

    const labels = chartData.data.labels
    const datasets = chartData.data.datasets[0].data
    labels.forEach((e,i) => {
      const panel = document.createElement("DIV")
      panel.className = "panel"
      panel.innerHTML = `<h1>${datasets[i] ? datasets[i] : "n/a"}</h1><p>${e}</p>`
      panelsDiv.appendChild(panel);
    })

    chartDiv.appendChild(panelsDiv) 
    chartWrapper.appendChild(chartDiv)
    return chartWrapper
  })
  
  const rowDiv = document.createElement("DIV")
  rowDiv.className = "row row-cols-1 table-container" //"row row-cols-1 row-cols-md-2 row-cols-lg-3"
  rowDiv.style.left = savePos;
  const navDiv = document.createElement("UL");
  navDiv.className = "dots";

  chartDivs.forEach((div,i) => {
    const navbutton = document.createElement("LI");
    if (i == active) navbutton.className = "active";
    navDiv.appendChild(navbutton);
    rowDiv.appendChild(div)
  })

  app.innerHTML = ""
  
  const updatedTime = new Date().toLocaleString()
  const title = document.createElement("DIV")
  title.className = "row-cols-1"
  title.innerHTML = `<div class="col"><h1 class="text-dark">${appName}</h1></div><div class="col text-muted fs-6" id="refresh-time">Refreshed: ${updatedTime} (refresh every ${refreshMins} mins)</div><div class="col text-muted fs-6" id="transition-status">Status: Running</div>`
  
  app.appendChild(title)
  app.appendChild(rowDiv)
  app.appendChild(navDiv)
}

function updateTransition(val) {
  let transition_status = document.getElementById('transition-status');
  transition_status.innerHTML =`Status: ${val}`;
  val=='Paused' ? transitionStat=false : transitionStat=true;
}

function next(){
    if (!transitionStat) updateTransition('Running');
    active = active + 1 <= lengthItems ? active + 1 : 0;
    reloadSlider();
}
function prev(){
    if (!transitionStat) updateTransition('Running');
    active = active - 1 >= 0 ? active - 1 : lengthItems;
    reloadSlider();
}
function pause(){
  updateTransition('Paused');
  clearInterval(refreshInterval);
}
function reloadSlider(){
  tableContainer.style.left = -items[active].offsetLeft + 'px';
  savePos = tableContainer.style.left;
  let last_active_dot = document.querySelector('.container-fluid .dots li.active');
  last_active_dot.classList.remove('active');
  dots[active].classList.add('active');

  if(refreshInterval) clearInterval(refreshInterval);
  refreshInterval = setInterval(()=> {next()}, transitionMins);
}
function transition() {
  tableContainer = document.querySelector('.container-fluid .table-container');
  items = document.querySelectorAll('.container-fluid .table-container .slider');
  dots = document.querySelectorAll('.container-fluid .dots li');

  lengthItems = items.length - 1;

  if(refreshInterval) clearInterval(refreshInterval);
  refreshInterval = setInterval(()=> {next()}, transitionMins);

  dots.forEach((li, key) => {
      li.addEventListener('click', ()=>{
          active = key;
          reloadSlider();
      })
  })
}
window.addEventListener('keydown', function(event) {
  if (event.key === 'ArrowRight') {
      next(); 
  } else if (event.key === 'ArrowLeft') {
      prev();
  } else if (event.key === ' ') {
      pause();
  }
});
window.onresize = function(event) {
  if(tableContainer)
    reloadSlider();
};