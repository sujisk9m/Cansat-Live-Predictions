// app.js
// Embedded Excel Data (simulated from Excel file)
// This data represents the Excel file content embedded in code
const EXCEL_DATA = [
  { timestamp: '10:00:00', speed: 15, ph: 7.2, salinity: 35.5, pressure: 101.3, turbidity: 2.1, energy: 4500, load: 100, leak: 'NO', depth: 5 },
  { timestamp: '10:00:05', speed: 18, ph: 7.3, salinity: 35.6, pressure: 101.5, turbidity: 2.3, energy: 4480, load: 98, leak: 'NO', depth: 6 },
  { timestamp: '10:00:10', speed: 22, ph: 7.1, salinity: 35.4, pressure: 101.8, turbidity: 2.5, energy: 4450, load: 95, leak: 'NO', depth: 7 },
  { timestamp: '10:00:15', speed: 25, ph: 7.4, salinity: 35.7, pressure: 102.1, turbidity: 2.7, energy: 4420, load: 92, leak: 'NO', depth: 8 },
  { timestamp: '10:00:20', speed: 28, ph: 7.2, salinity: 35.5, pressure: 102.4, turbidity: 2.9, energy: 4390, load: 90, leak: 'NO', depth: 9 },
  { timestamp: '10:00:25', speed: 32, ph: 7.3, salinity: 35.6, pressure: 102.7, turbidity: 3.1, energy: 4360, load: 88, leak: 'NO', depth: 10 },
  { timestamp: '10:00:30', speed: 35, ph: 7.1, salinity: 35.4, pressure: 103.0, turbidity: 3.3, energy: 4330, load: 85, leak: 'NO', depth: 11 },
  { timestamp: '10:00:35', speed: 38, ph: 7.4, salinity: 35.8, pressure: 103.3, turbidity: 3.5, energy: 4300, load: 83, leak: 'NO', depth: 12 },
  { timestamp: '10:00:40', speed: 42, ph: 7.2, salinity: 35.5, pressure: 103.6, turbidity: 3.7, energy: 4270, load: 80, leak: 'NO', depth: 13 },
  { timestamp: '10:00:45', speed: 45, ph: 7.3, salinity: 35.7, pressure: 103.9, turbidity: 3.9, energy: 4240, load: 78, leak: 'NO', depth: 14 },
  { timestamp: '10:00:50', speed: 48, ph: 7.1, salinity: 35.4, pressure: 104.2, turbidity: 4.1, energy: 4210, load: 75, leak: 'NO', depth: 15 },
  { timestamp: '10:00:55', speed: 52, ph: 7.4, salinity: 35.8, pressure: 104.5, turbidity: 4.3, energy: 4180, load: 73, leak: 'NO', depth: 16 },
  { timestamp: '10:01:00', speed: 55, ph: 7.2, salinity: 35.6, pressure: 104.8, turbidity: 4.5, energy: 4150, load: 70, leak: 'NO', depth: 17 },
  { timestamp: '10:01:05', speed: 58, ph: 7.3, salinity: 35.7, pressure: 105.1, turbidity: 4.7, energy: 4120, load: 68, leak: 'NO', depth: 18 },
  { timestamp: '10:01:10', speed: 62, ph: 7.1, salinity: 35.5, pressure: 105.4, turbidity: 4.9, energy: 4090, load: 65, leak: 'NO', depth: 19 },
  { timestamp: '10:01:15', speed: 65, ph: 7.4, salinity: 35.8, pressure: 105.7, turbidity: 5.1, energy: 4060, load: 63, leak: 'NO', depth: 20 },
  { timestamp: '10:01:20', speed: 68, ph: 7.2, salinity: 35.6, pressure: 106.0, turbidity: 5.3, energy: 4030, load: 60, leak: 'NO', depth: 21 },
  { timestamp: '10:01:25', speed: 72, ph: 7.3, salinity: 35.7, pressure: 106.3, turbidity: 5.5, energy: 4000, load: 58, leak: 'NO', depth: 22 },
  { timestamp: '10:01:30', speed: 75, ph: 7.1, salinity: 35.4, pressure: 106.6, turbidity: 5.7, energy: 3970, load: 55, leak: 'NO', depth: 23 },
  { timestamp: '10:01:35', speed: 78, ph: 7.4, salinity: 35.9, pressure: 106.9, turbidity: 5.9, energy: 3940, load: 53, leak: 'NO', depth: 24 },
  { timestamp: '10:01:40', speed: 82, ph: 7.2, salinity: 35.6, pressure: 107.2, turbidity: 6.1, energy: 3910, load: 50, leak: 'NO', depth: 25 },
  { timestamp: '10:01:45', speed: 85, ph: 7.3, salinity: 35.7, pressure: 107.5, turbidity: 6.3, energy: 3880, load: 48, leak: 'NO', depth: 26 },
  { timestamp: '10:01:50', speed: 88, ph: 7.1, salinity: 35.5, pressure: 107.8, turbidity: 6.5, energy: 3850, load: 45, leak: 'NO', depth: 27 },
  { timestamp: '10:01:55', speed: 92, ph: 7.4, salinity: 35.8, pressure: 108.1, turbidity: 6.7, energy: 3820, load: 43, leak: 'NO', depth: 28 },
  { timestamp: '10:02:00', speed: 95, ph: 7.2, salinity: 35.6, pressure: 108.4, turbidity: 6.9, energy: 3790, load: 40, leak: 'YES', depth: 29 },
  { timestamp: '10:02:05', speed: 98, ph: 7.3, salinity: 35.7, pressure: 108.7, turbidity: 7.1, energy: 3760, load: 38, leak: 'YES', depth: 30 },
  { timestamp: '10:02:10', speed: 100, ph: 7.1, salinity: 35.4, pressure: 109.0, turbidity: 7.3, energy: 3730, load: 35, leak: 'YES', depth: 31 },
  { timestamp: '10:02:15', speed: 98, ph: 7.4, salinity: 35.9, pressure: 109.3, turbidity: 7.5, energy: 3700, load: 33, leak: 'YES', depth: 32 },
  { timestamp: '10:02:20', speed: 95, ph: 7.2, salinity: 35.6, pressure: 109.6, turbidity: 7.7, energy: 3670, load: 30, leak: 'YES', depth: 33 },
  { timestamp: '10:02:25', speed: 92, ph: 7.3, salinity: 35.7, pressure: 109.9, turbidity: 7.9, energy: 3640, load: 28, leak: 'YES', depth: 34 },
  { timestamp: '10:02:30', speed: 88, ph: 7.1, salinity: 35.5, pressure: 110.2, turbidity: 8.1, energy: 3610, load: 25, leak: 'YES', depth: 35 },
  { timestamp: '10:02:35', speed: 85, ph: 7.4, salinity: 35.8, pressure: 110.5, turbidity: 8.3, energy: 3580, load: 23, leak: 'YES', depth: 36 },
  { timestamp: '10:02:40', speed: 82, ph: 7.2, salinity: 35.6, pressure: 110.8, turbidity: 8.5, energy: 3550, load: 20, leak: 'YES', depth: 37 },
  { timestamp: '10:02:45', speed: 78, ph: 7.3, salinity: 35.7, pressure: 111.1, turbidity: 8.7, energy: 3520, load: 18, leak: 'YES', depth: 38 },
  { timestamp: '10:02:50', speed: 75, ph: 7.1, salinity: 35.4, pressure: 111.4, turbidity: 8.9, energy: 3490, load: 15, leak: 'YES', depth: 39 },
  { timestamp: '10:02:55', speed: 72, ph: 7.4, salinity: 35.9, pressure: 111.7, turbidity: 9.1, energy: 3460, load: 13, leak: 'YES', depth: 40 },
  { timestamp: '10:03:00', speed: 68, ph: 7.2, salinity: 35.6, pressure: 112.0, turbidity: 9.3, energy: 3430, load: 10, leak: 'YES', depth: 41 },
  { timestamp: '10:03:05', speed: 65, ph: 7.3, salinity: 35.7, pressure: 112.3, turbidity: 9.5, energy: 3400, load: 8, leak: 'YES', depth: 42 },
  { timestamp: '10:03:10', speed: 62, ph: 7.1, salinity: 35.5, pressure: 112.6, turbidity: 9.7, energy: 3370, load: 5, leak: 'YES', depth: 43 },
  { timestamp: '10:03:15', speed: 58, ph: 7.4, salinity: 35.8, pressure: 112.9, turbidity: 9.9, energy: 3340, load: 3, leak: 'YES', depth: 44 },
  { timestamp: '10:03:20', speed: 55, ph: 7.2, salinity: 35.6, pressure: 113.2, turbidity: 10.1, energy: 3310, load: 0, leak: 'YES', depth: 45 }
];

const POLL_MS = 2000; // Update every 2 seconds for smoother animation
let currentDataIndex = 0;

// Charts
let speedChart, phTrendChart, salTrendChart, pressTrendChart, turbTrendChart;

function makeSemiDonut(ctx){
  return new Chart(ctx, {
    type: 'doughnut',
    data: {
      labels: ['Speed', 'Rest'],
      datasets: [{
        data: [0, 100],
        backgroundColor: [
          'rgba(74, 158, 255, 0.95)',
          'rgba(106, 13, 173, 0.7)'
        ],
        borderColor: [
          '#4a9eff',
          'rgba(139, 92, 246, 0.25)'
        ],
        borderWidth: 3,
        cutout: '75%'
      }]
    },
    options: {
      rotation: -90,
      circumference: 180,
      plugins: {
        legend: { display: false },
        tooltip: {
          enabled: true,
          backgroundColor: 'rgba(0, 0, 0, 0.9)',
          titleColor: '#4cc9f0',
          bodyColor: '#ffffff',
          borderColor: '#4cc9f0',
          borderWidth: 2,
          callbacks: {
            label: function(context) {
              return 'Speed: ' + context.parsed + '%';
            }
          }
        }
      },
      maintainAspectRatio: false
    }
  });
}

function makeLine(ctx, color, min, max){
  return new Chart(ctx, {
    type: 'line',
    data: {
      labels: [],
      datasets: [{
        data: [],
        tension: 0.5, // Smoother curves for realistic readings
        pointRadius: 0,
        pointHoverRadius: 3,
        borderWidth: 2.5,
        borderColor: color,
        backgroundColor: 'rgba(74, 158, 255, 0.18)',
        fill: true,
        borderCapStyle: 'round',
        borderJoinStyle: 'round'
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      animation: {
        duration: 750,
        easing: 'easeOutQuart'
      },
      plugins: {
        legend: { display: false },
        tooltip: {
          enabled: true,
          backgroundColor: 'rgba(0, 0, 0, 0.9)',
          titleColor: '#4cc9f0',
          bodyColor: '#ffffff',
          borderColor: '#4cc9f0',
          borderWidth: 2,
          padding: 10,
          displayColors: false,
          callbacks: {
            label: function(context) {
              return 'Value: ' + context.parsed.y.toFixed(2);
            }
          }
        }
      },
      scales: {
        x: {
          display: false,
          grid: { display: false }
        },
        y: {
          min: min,
          max: max,
          ticks: {
            color: '#a5b4fc',
            font: { size: 11, weight: 'bold' },
            stepSize: (max - min) / 5
          },
          grid: {
            color: 'rgba(74, 158, 255, 0.15)',
            lineWidth: 1,
            drawBorder: false
          }
        }
      },
      elements: {
        line: {
          borderColor: color,
          borderWidth: 2.5,
          tension: 0.5
        },
        point: {
          radius: 0,
          hoverRadius: 4,
          hoverBorderWidth: 2,
          hoverBackgroundColor: '#8b5cf6'
        }
      },
      interaction: {
        intersect: false,
        mode: 'index'
      }
    }
  });
}

function initCharts(){
  try { if(speedChart) speedChart.destroy(); } catch(e){}
  try { if(phTrendChart) phTrendChart.destroy(); } catch(e){}
  try { if(salTrendChart) salTrendChart.destroy(); } catch(e){}
  try { if(pressTrendChart) pressTrendChart.destroy(); } catch(e){}
  try { if(turbTrendChart) turbTrendChart.destroy(); } catch(e){}

  const speedGaugeEl = document.getElementById('speedGauge');
  const phTrendEl = document.getElementById('phTrend');
  const salTrendEl = document.getElementById('salTrend');
  const pressTrendEl = document.getElementById('pressTrend');
  const turbTrendEl = document.getElementById('turbTrend');

  if(speedGaugeEl) {
    speedChart = makeSemiDonut(speedGaugeEl.getContext('2d'));
  }
  if(phTrendEl) {
    phTrendChart = makeLine(phTrendEl.getContext('2d'), '#4a9eff', 6.5, 8.0);
  }
  if(salTrendEl) {
    salTrendChart = makeLine(salTrendEl.getContext('2d'), '#6ba3ff', 34.0, 37.0);
  }
  if(pressTrendEl) {
    pressTrendChart = makeLine(pressTrendEl.getContext('2d'), '#4a9eff', 100.0, 115.0);
  }
  if(turbTrendEl) {
    turbTrendChart = makeLine(turbTrendEl.getContext('2d'), '#2563eb', 0.0, 12.0);
  }
}

// Wait for DOM to be ready before initializing charts
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', initCharts);
} else {
  initCharts();
}

let currentRows = [];
let pollTimer = null;

// Get DOM elements with null checks
let phNow, salNow, pressNow, turbNow, energyNow, loadNow, speedVal, leakPill, depthNow;

function initDOMElements() {
  phNow = document.getElementById('phNow');
  salNow = document.getElementById('salNow');
  pressNow = document.getElementById('pressNow');
  turbNow = document.getElementById('turbNow');
  energyNow = document.getElementById('energyNow');
  loadNow = document.getElementById('loadNow');
  speedVal = document.getElementById('speedVal');
  leakPill = document.getElementById('leakPill');
  depthNow = document.getElementById('depthNow');
}

function normalize(o){
  o.speed = o.speed ? Number(o.speed) : 0;
  o.ph = o.ph ? Number(o.ph) : null;
  o.salinity = o.salinity ? Number(o.salinity) : null;
  o.turbidity = o.turbidity ? Number(o.turbidity) : null;
  o.pressure = o.pressure ? Number(o.pressure) : null;
  o.energy = o.energy ? Number(o.energy) : null;
  o.load = o.load ? Number(o.load) : null;
  o.depth = o.depth ? Number(o.depth) : null;

  if(o.leak !== undefined){
    const s = String(o.leak).toLowerCase();
    o.leak = (s === 'yes' || s === '1' || s === 'true' || s === 'detected');
  }

  if(!o.timestamp){
    const d = new Date();
    o.timestamp = d.toTimeString().split(" ")[0];
  }
}

function updateOverviewUI(latest){
  if(!latest) return;

  const ovPH = document.getElementById('ovPH');
  const ovSal = document.getElementById('ovSal');
  const ovPress = document.getElementById('ovPress');
  const ovLoad = document.getElementById('ovLoad');
  
  if(ovPH) ovPH.textContent = (latest.ph != null && latest.ph !== '') ? latest.ph.toFixed(1) : '--';
  if(ovSal) ovSal.textContent = (latest.salinity != null && latest.salinity !== '') ? latest.salinity.toFixed(1) : '--';
  if(ovPress) ovPress.textContent = (latest.pressure != null && latest.pressure !== '') ? latest.pressure.toFixed(1) : '--';
  if(ovLoad) ovLoad.textContent = (latest.load != null && latest.load !== '') ? latest.load : '--';
}

function applyRowsToUI(){
  const rows = currentRows;
  const rowCountEl = document.getElementById('rowCount');
  if(rowCountEl) rowCountEl.textContent = rows.length;
  if(!rows.length) return;

  const latest = rows[rows.length-1];

  if(phNow) phNow.textContent = (latest.ph != null && latest.ph !== '') ? latest.ph.toFixed(1) : "--";
  if(salNow) salNow.textContent = (latest.salinity != null && latest.salinity !== '') ? latest.salinity.toFixed(1) : "--";
  if(pressNow) pressNow.textContent = (latest.pressure != null && latest.pressure !== '') ? latest.pressure.toFixed(1) : "--";
  if(turbNow) turbNow.textContent = (latest.turbidity != null && latest.turbidity !== '') ? latest.turbidity.toFixed(1) : "--";
  if(energyNow) energyNow.textContent = (latest.energy != null && latest.energy !== '') ? latest.energy : "--";
  if(loadNow) loadNow.textContent = (latest.load != null && latest.load !== '') ? latest.load : "--";
  if(speedVal) speedVal.textContent = (latest.speed != null ? latest.speed : 0) + "%";
  if(depthNow) depthNow.textContent = (latest.depth != null && latest.depth !== '') ? latest.depth : "--";

  if(leakPill) {
    const leakText = latest.leak ? "YES" : "NO";
    leakPill.textContent = "Last: " + leakText;
    leakPill.style.color = latest.leak ? "#ff0000" : "#00ff00";
    leakPill.style.borderColor = latest.leak ? "#ff0000" : "#00ff00";
  }

  // Update speed gauge
  if(speedChart) {
    const speed = latest.speed != null ? Math.max(0, Math.min(100, latest.speed)) : 0;
    speedChart.data.datasets[0].data = [speed, 100 - speed];
    speedChart.update('none');
  }

  // Update trend charts with last 50 data points for realistic readings
  const slice = rows.slice(-50);
  
  // Update with smooth animation for realistic sensor readings
  if(phTrendChart && slice.length > 0) {
    const phData = slice.map(r => r.ph).filter(v => v !== null && v !== undefined);
    if(phData.length > 0) {
      phTrendChart.data.labels = Array(phData.length).fill('');
      phTrendChart.data.datasets[0].data = phData;
      phTrendChart.update('active');
    }
  }

  if(salTrendChart && slice.length > 0) {
    const salData = slice.map(r => r.salinity).filter(v => v !== null && v !== undefined);
    if(salData.length > 0) {
      salTrendChart.data.labels = Array(salData.length).fill('');
      salTrendChart.data.datasets[0].data = salData;
      salTrendChart.update('active');
    }
  }

  if(pressTrendChart && slice.length > 0) {
    const pressData = slice.map(r => r.pressure).filter(v => v !== null && v !== undefined);
    if(pressData.length > 0) {
      pressTrendChart.data.labels = Array(pressData.length).fill('');
      pressTrendChart.data.datasets[0].data = pressData;
      pressTrendChart.update('active');
    }
  }

  if(turbTrendChart && slice.length > 0) {
    const turbData = slice.map(r => r.turbidity).filter(v => v !== null && v !== undefined);
    if(turbData.length > 0) {
      turbTrendChart.data.labels = Array(turbData.length).fill('');
      turbTrendChart.data.datasets[0].data = turbData;
      turbTrendChart.update('active');
    }
  }

  // Update records table
  const tbody = document.querySelector('#recordsTable tbody');
  if(tbody) {
    tbody.innerHTML = '';
    const last10 = rows.slice(-10).reverse();
    last10.forEach(r=>{
      const tr = document.createElement('tr');
      const phVal = (r.ph != null && r.ph !== '') ? r.ph.toFixed(1) : '';
      const salVal = (r.salinity != null && r.salinity !== '') ? r.salinity.toFixed(1) : '';
      const pressVal = (r.pressure != null && r.pressure !== '') ? r.pressure.toFixed(1) : '';
      const turbVal = (r.turbidity != null && r.turbidity !== '') ? r.turbidity.toFixed(1) : '';
      tr.innerHTML = `<td>${r.timestamp||''}</td><td>${phVal}</td><td>${salVal}</td><td>${pressVal}</td><td>${turbVal}</td><td>${r.leak? 'YES':'NO'}</td>`;
      tbody.appendChild(tr);
    });
  }

  // Update load display
  const loadDisplayEl = document.getElementById('loadDisplay');
  if(loadDisplayEl) loadDisplayEl.textContent = (latest.load != null ? latest.load : 100) + " ccm";

  // Update simulation page
  updateSimulationPage(latest);
  
  // Update analytics page
  updateAnalyticsPage(rows);

  try { updateOverviewUI(latest); } catch(e){ console.warn(e); }
}

function loadExcelData(){
  // Simulate reading from Excel file (embedded data)
  // Add slight random variations to make it look more realistic
  const baseData = EXCEL_DATA[currentDataIndex % EXCEL_DATA.length];
  
  // Add small random variations with smooth transitions for realism
  const variation = (val, percent = 2) => {
    if (!val && val !== 0) return val;
    // Use a smoother random variation
    const randomFactor = (Math.random() - 0.5) * (percent / 50);
    // Add slight noise for sensor-like readings
    const noise = (Math.random() - 0.5) * (percent / 200);
    return val * (1 + randomFactor + noise);
  };

  // Generate timestamp based on current index
  const now = new Date();
  const timeStr = now.toTimeString().split(' ')[0];
  
  const row = {
    timestamp: timeStr,
    speed: Math.min(100, Math.max(0, Math.round(variation(baseData.speed, 3)))),
    ph: Math.max(6.0, Math.min(8.5, variation(baseData.ph, 1.5))),
    salinity: Math.max(33.0, Math.min(38.0, variation(baseData.salinity, 1.5))),
    pressure: Math.max(100.0, Math.min(115.0, variation(baseData.pressure, 1.5))),
    turbidity: Math.max(0.0, Math.min(12.0, variation(baseData.turbidity, 3))),
    energy: Math.max(3000, Math.round(variation(baseData.energy, 2))),
    load: Math.max(0, Math.min(100, Math.round(variation(baseData.load, 3)))),
    leak: baseData.leak,
    depth: Math.max(0, Math.round(variation(baseData.depth, 2)))
  };

  normalize(row);
  currentRows.push(row);
  
  // Keep only last 100 rows for performance and smooth scrolling
  if(currentRows.length > 100){
    currentRows = currentRows.slice(-100);
  }

  applyRowsToUI();
  currentDataIndex++;
}

function startPolling(){
  if(pollTimer) clearInterval(pollTimer);
  loadExcelData();
  pollTimer = setInterval(() => loadExcelData(), POLL_MS);
}

function stopPolling(){
  if(pollTimer) {
    clearInterval(pollTimer);
    pollTimer = null;
  }
}

// Analytics Chart
let analyticsChart = null;

function initAnalyticsChart() {
  const ctx = document.getElementById('analyticsChart');
  if(!ctx) {
    console.warn('Analytics chart canvas not found');
    return;
  }
  
  try {
    if(analyticsChart) {
      analyticsChart.destroy();
      analyticsChart = null;
    }
  } catch(e) {
    console.warn('Error destroying analytics chart:', e);
  }
  
  try {
    analyticsChart = new Chart(ctx.getContext('2d'), {
      type: 'line',
      data: {
        labels: [],
        datasets: [
          {
            label: 'pH',
            data: [],
            borderColor: '#4a9eff',
            backgroundColor: 'rgba(74, 158, 255, 0.15)',
            tension: 0.4,
            pointRadius: 0,
            pointHoverRadius: 4,
            borderWidth: 2
          },
          {
            label: 'Salinity',
            data: [],
            borderColor: '#6ba3ff',
            backgroundColor: 'rgba(107, 163, 255, 0.15)',
            tension: 0.4,
            pointRadius: 0,
            pointHoverRadius: 4,
            borderWidth: 2
          },
          {
            label: 'Pressure',
            data: [],
            borderColor: '#2563eb',
            backgroundColor: 'rgba(37, 99, 235, 0.18)',
            tension: 0.4,
            pointRadius: 0,
            pointHoverRadius: 4,
            borderWidth: 2
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        animation: {
          duration: 750,
          easing: 'easeOutQuart'
        },
        plugins: {
          legend: {
            display: true,
            position: 'top',
            labels: {
              color: '#6ba3ff',
              font: {
                size: 12
              },
              padding: 10,
              usePointStyle: true
            }
          },
          tooltip: {
            enabled: true,
            backgroundColor: 'rgba(0, 0, 0, 0.9)',
            titleColor: '#4a9eff',
            bodyColor: '#ffffff',
            borderColor: '#4a9eff',
            borderWidth: 1,
            padding: 10
          }
        },
        scales: {
          x: {
            display: true,
            ticks: { 
              color: '#a5b4fc',
              font: { size: 10 }
            },
            grid: { 
              color: 'rgba(74, 158, 255, 0.12)',
              display: true
            }
          },
          y: {
            display: true,
            ticks: { 
              color: '#a5b4fc',
              font: { size: 10 }
            },
            grid: { 
              color: 'rgba(74, 158, 255, 0.12)',
              display: true
            },
            beginAtZero: false
          }
        },
        interaction: {
          intersect: false,
          mode: 'index'
        }
      }
    });
    console.log('Analytics chart initialized successfully');
  } catch(e) {
    console.error('Error initializing analytics chart:', e);
    analyticsChart = null;
  }
}

function updateSimulationPage(latest) {
  if(!latest) return;
  
  const simDepth = document.getElementById('simDepthNow');
  const simSpeed = document.getElementById('simSpeedNow');
  const simDepthVal = document.getElementById('simDepth');
  const simPressure = document.getElementById('simPressure');
  const simTemp = document.getElementById('simTemp');
  const simVisibility = document.getElementById('simVisibility');
  const simBattery = document.getElementById('simBattery');
  const simSignal = document.getElementById('simSignal');
  
  if(simDepth) simDepth.textContent = latest.depth ?? '--';
  if(simSpeed) simSpeed.textContent = latest.speed ?? '--';
  if(simDepthVal) simDepthVal.textContent = (latest.depth ?? '--') + ' m';
  if(simPressure) simPressure.textContent = (latest.pressure ?? '--') + ' kPa';
  if(simTemp) simTemp.textContent = (15 + Math.random() * 5).toFixed(1) + ' °C';
  if(simVisibility) simVisibility.textContent = (10 + Math.random() * 20).toFixed(1) + ' m';
  if(simBattery) simBattery.textContent = Math.round((latest.energy ?? 4000) / 50) + ' %';
  if(simSignal) simSignal.textContent = (85 + Math.random() * 10).toFixed(0) + ' %';
}

function updateAnalyticsPage(rows) {
  if(!rows || rows.length === 0) return;
  
  // Calculate statistics
  const phValues = rows.map(r => r.ph).filter(v => v != null && v !== undefined);
  const salValues = rows.map(r => r.salinity).filter(v => v != null && v !== undefined);
  const pressValues = rows.map(r => r.pressure).filter(v => v != null && v !== undefined);
  const leakCount = rows.filter(r => r.leak).length;
  
  const avgPH = document.getElementById('avgPH');
  const avgSal = document.getElementById('avgSal');
  const maxPress = document.getElementById('maxPress');
  const minPress = document.getElementById('minPress');
  const totalRecords = document.getElementById('totalRecords');
  const leakEvents = document.getElementById('leakEvents');
  
  if(avgPH) avgPH.textContent = phValues.length > 0 ? (phValues.reduce((a,b) => a+b, 0) / phValues.length).toFixed(2) : '--';
  if(avgSal) avgSal.textContent = salValues.length > 0 ? (salValues.reduce((a,b) => a+b, 0) / salValues.length).toFixed(2) : '--';
  if(maxPress) maxPress.textContent = pressValues.length > 0 ? Math.max(...pressValues).toFixed(2) : '--';
  if(minPress) minPress.textContent = pressValues.length > 0 ? Math.min(...pressValues).toFixed(2) : '--';
  if(totalRecords) totalRecords.textContent = rows.length;
  if(leakEvents) leakEvents.textContent = leakCount;
  
  // Update analytics chart
  const ctx = document.getElementById('analyticsChart');
  if(ctx && analyticsChart) {
    try {
      const slice = rows.slice(-30);
      if(slice.length > 0) {
        const labels = slice.map((r, i) => i.toString());
        const phData = slice.map(r => r.ph).filter(v => v != null && v !== undefined);
        const salData = slice.map(r => r.salinity).filter(v => v != null && v !== undefined);
        const pressData = slice.map(r => r.pressure).filter(v => v != null && v !== undefined);
        
        // Normalize data lengths
        const minLength = Math.min(labels.length, phData.length, salData.length, pressData.length);
        
        analyticsChart.data.labels = labels.slice(0, minLength);
        analyticsChart.data.datasets[0].data = phData.slice(0, minLength);
        analyticsChart.data.datasets[1].data = salData.slice(0, minLength);
        analyticsChart.data.datasets[2].data = pressData.slice(0, minLength);
        analyticsChart.update('active');
      }
    } catch(e) {
      console.warn('Analytics chart update error:', e);
      // Reinitialize chart if there's an error
      if(analyticsChart) {
        try {
          analyticsChart.destroy();
        } catch(e2) {}
        initAnalyticsChart();
      }
    }
  } else if(ctx && !analyticsChart) {
    // Initialize chart if it doesn't exist
    initAnalyticsChart();
  }
  
  // Update analytics table
  const analyticsTbody = document.querySelector('#analyticsTable tbody');
  if(analyticsTbody) {
    analyticsTbody.innerHTML = '';
    const last20 = rows.slice(-20).reverse();
    last20.forEach(r => {
      const tr = document.createElement('tr');
      const phVal = (r.ph != null && r.ph !== '') ? r.ph.toFixed(1) : '';
      const salVal = (r.salinity != null && r.salinity !== '') ? r.salinity.toFixed(1) : '';
      const pressVal = (r.pressure != null && r.pressure !== '') ? r.pressure.toFixed(1) : '';
      const turbVal = (r.turbidity != null && r.turbidity !== '') ? r.turbidity.toFixed(1) : '';
      tr.innerHTML = `<td>${r.timestamp||''}</td><td>${phVal}</td><td>${salVal}</td><td>${pressVal}</td><td>${turbVal}</td><td>${r.leak? 'YES':'NO'}</td>`;
      analyticsTbody.appendChild(tr);
    });
  }
}

// Helper: download arbitrary content as a file
function downloadBlob(content, filename, mimeType) {
  try {
    const blob = new Blob([content], { type: mimeType });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  } catch (e) {
    console.error('Download error:', e);
    alert('Unable to start download on this browser.');
  }
}

// Export data entry point used by index.html inline script
window.exportDataFromApp = function(format) {
  const rows = currentRows || [];
  if (!rows.length) {
    alert('No data available to export yet.');
    return;
  }

  const ts = new Date().toISOString().replace(/[:.]/g, '-');

  if (format === 'csv') {
    const header = ['timestamp','ph','salinity','pressure','turbidity','leak'];
    const lines = [header.join(',')];
    rows.forEach(r => {
      lines.push([
        r.timestamp || '',
        r.ph ?? '',
        r.salinity ?? '',
        r.pressure ?? '',
        r.turbidity ?? '',
        r.leak ? 'YES' : 'NO'
      ].join(','));
    });
    const csv = lines.join('\n');
    downloadBlob(csv, `cansat_data_${ts}.csv`, 'text/csv;charset=utf-8;');
    return;
  }

  if (format === 'json') {
    const json = JSON.stringify(rows, null, 2);
    downloadBlob(json, `cansat_data_${ts}.json`, 'application/json;charset=utf-8;');
    return;
  }

  if (format === 'pdf') {
    if (window.jspdf && window.jspdf.jsPDF) {
      const { jsPDF } = window.jspdf;
      const doc = new jsPDF();

      doc.setFontSize(14);
      doc.text('Ocean CANSAT Data Export', 14, 18);
      doc.setFontSize(10);
      doc.text(`Generated: ${new Date().toLocaleString()}`, 14, 24);

      const header = ['Time','pH','Sal','Press','Turb','Leak'];
      let y = 32;
      doc.text(header.join(' | '), 14, y);
      y += 6;

      rows.slice(-40).forEach(r => {
        if (y > 280) {
          doc.addPage();
          y = 20;
        }
        const line = [
          r.timestamp || '',
          r.ph != null ? r.ph.toFixed(2) : '',
          r.salinity != null ? r.salinity.toFixed(2) : '',
          r.pressure != null ? r.pressure.toFixed(2) : '',
          r.turbidity != null ? r.turbidity.toFixed(2) : '',
          r.leak ? 'YES' : 'NO'
        ].join(' | ');
        doc.text(line, 14, y);
        y += 5;
      });

      doc.save(`cansat_data_${ts}.pdf`);
    } else {
      alert('PDF export library not loaded yet. Please try again in a moment.');
    }
    return;
  }

  alert('Unknown export format: ' + format);
};

function applyRowsToUI(){
  const rows = currentRows;
  const rowCountEl = document.getElementById('rowCount');
  if(rowCountEl) rowCountEl.textContent = rows.length;
  if(!rows.length) return;

  const latest = rows[rows.length-1];

  if(phNow) phNow.textContent = (latest.ph != null && latest.ph !== '') ? latest.ph.toFixed(1) : "--";
  if(salNow) salNow.textContent = (latest.salinity != null && latest.salinity !== '') ? latest.salinity.toFixed(1) : "--";
  if(pressNow) pressNow.textContent = (latest.pressure != null && latest.pressure !== '') ? latest.pressure.toFixed(1) : "--";
  if(turbNow) turbNow.textContent = (latest.turbidity != null && latest.turbidity !== '') ? latest.turbidity.toFixed(1) : "--";
  if(energyNow) energyNow.textContent = (latest.energy != null && latest.energy !== '') ? latest.energy : "--";
  if(loadNow) loadNow.textContent = (latest.load != null && latest.load !== '') ? latest.load : "--";
  if(speedVal) speedVal.textContent = (latest.speed != null ? latest.speed : 0) + "%";
  if(depthNow) depthNow.textContent = (latest.depth != null && latest.depth !== '') ? latest.depth : "--";

  if(leakPill) {
    const leakText = latest.leak ? "YES" : "NO";
    leakPill.textContent = "Last: " + leakText;
    leakPill.style.color = latest.leak ? "#ff0000" : "#00ff00";
    leakPill.style.borderColor = latest.leak ? "#ff0000" : "#00ff00";
  }

  // Update speed gauge
  if(speedChart) {
    const speed = latest.speed != null ? Math.max(0, Math.min(100, latest.speed)) : 0;
    speedChart.data.datasets[0].data = [speed, 100 - speed];
    speedChart.update('none');
  }

  // Update trend charts with last 50 data points for realistic readings
  const slice = rows.slice(-50);
  
  // Update with smooth animation for realistic sensor readings
  if(phTrendChart && slice.length > 0) {
    const phData = slice.map(r => r.ph).filter(v => v !== null && v !== undefined);
    if(phData.length > 0) {
      phTrendChart.data.labels = Array(phData.length).fill('');
      phTrendChart.data.datasets[0].data = phData;
      phTrendChart.update('active');
    }
  }

  if(salTrendChart && slice.length > 0) {
    const salData = slice.map(r => r.salinity).filter(v => v !== null && v !== undefined);
    if(salData.length > 0) {
      salTrendChart.data.labels = Array(salData.length).fill('');
      salTrendChart.data.datasets[0].data = salData;
      salTrendChart.update('active');
    }
  }

  if(pressTrendChart && slice.length > 0) {
    const pressData = slice.map(r => r.pressure).filter(v => v !== null && v !== undefined);
    if(pressData.length > 0) {
      pressTrendChart.data.labels = Array(pressData.length).fill('');
      pressTrendChart.data.datasets[0].data = pressData;
      pressTrendChart.update('active');
    }
  }

  if(turbTrendChart && slice.length > 0) {
    const turbData = slice.map(r => r.turbidity).filter(v => v !== null && v !== undefined);
    if(turbData.length > 0) {
      turbTrendChart.data.labels = Array(turbData.length).fill('');
      turbTrendChart.data.datasets[0].data = turbData;
      turbTrendChart.update('active');
    }
  }

  // Update records table
  const tbody = document.querySelector('#recordsTable tbody');
  if(tbody) {
    tbody.innerHTML = '';
    const last10 = rows.slice(-10).reverse();
    last10.forEach(r=>{
      const tr = document.createElement('tr');
      const phVal = (r.ph != null && r.ph !== '') ? r.ph.toFixed(1) : '';
      const salVal = (r.salinity != null && r.salinity !== '') ? r.salinity.toFixed(1) : '';
      const pressVal = (r.pressure != null && r.pressure !== '') ? r.pressure.toFixed(1) : '';
      const turbVal = (r.turbidity != null && r.turbidity !== '') ? r.turbidity.toFixed(1) : '';
      tr.innerHTML = `<td>${r.timestamp||''}</td><td>${phVal}</td><td>${salVal}</td><td>${pressVal}</td><td>${turbVal}</td><td>${r.leak? 'YES':'NO'}</td>`;
      tbody.appendChild(tr);
    });
  }

  // Update load display
  const loadDisplayEl = document.getElementById('loadDisplay');
  if(loadDisplayEl) loadDisplayEl.textContent = (latest.load != null ? latest.load : 100) + " ccm";

  // Update simulation page
  updateSimulationPage(latest);
  
  // Update analytics page
  updateAnalyticsPage(rows);

  try { updateOverviewUI(latest); } catch(e){ console.warn(e); }
}

// Make functions globally accessible for HTML script
window.initAnalyticsChart = initAnalyticsChart;
window.updateAnalyticsPage = updateAnalyticsPage;
window.currentRows = currentRows;

window.onload = () => {
  // Initialize DOM elements
  initDOMElements();
  
  // Initialize analytics chart after a short delay to ensure DOM is ready
  setTimeout(() => {
    initAnalyticsChart();
  }, 200);
  
  // Start automatically with embedded Excel data
  startPolling();
};
