// charts.js — inicializa todos os gráficos do dashboard

const COLORS = {
  navy:      '#1B5E38',
  navyDark:  '#0D3321',
  green:     '#5FAD41',
  lime:      '#8DC642',
  greenMid:  '#2E7D4F',
  greenLight:'rgba(141,198,66,0.2)',
  red:       '#C0504D',
  orange:    '#F79646',
  teal:      '#4BACC6',
};

Chart.defaults.font.family = 'Inter, Arial, sans-serif';
Chart.defaults.color = '#64748b';
Chart.defaults.animation = false;

function mkChart(id, config) {
  const el = document.getElementById(id);
  if (!el) return;
  if (el._chart) el._chart.destroy();
  el._chart = new Chart(el.getContext('2d'), config);
  return el._chart;
}

// =========================================================
// CAPACITY — Barras agrupadas por operação ao longo dos meses
// =========================================================
function buildCapacityChart() {
  const d = DADOS.capacity;
  mkChart('chart-capacity', {
    type: 'bar',
    data: {
      labels: d.meses,
      datasets: d.operacoes.map((op, i) => ({
        label: op,
        data: d.valores[i],
        backgroundColor: [COLORS.navy, COLORS.green, COLORS.lime, COLORS.greenMid][i],
        borderRadius: 6,
        borderSkipped: false,
      }))
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { position: 'top' } },
      scales: {
        x: { grid: { display: false } },
        y: { grid: { color: '#e8eef5' }, ticks: { stepSize: 2 }, title: { display:true, text:'Agentes' } }
      }
    }
  });

  // Totais — linha
  mkChart('chart-capacity-total', {
    type: 'line',
    data: {
      labels: d.meses,
      datasets: [{
        label: 'Total de Agentes',
        data: d.totais,
        borderColor: COLORS.navy,
        backgroundColor: 'rgba(27,94,56,0.12)',
        fill: true,
        tension: 0.4,
        pointRadius: 6,
        pointBackgroundColor: COLORS.navy,
        borderWidth: 3,
      }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        x: { grid: { display: false } },
        y: { grid: { color: '#e8eef5' }, min: 14,
          title: { display:true, text:'Total' } }
      }
    }
  });
}

// =========================================================
// FUNIL — Comparativo CPC% e Promessa% entre operações
// =========================================================
function buildFunilCharts(tabKey) {
  const rows = DADOS.funil[tabKey];
  const ops  = rows.map(r => r.op);
  const cpcPcts  = rows.map(r => +(r.cpc_pct * 100).toFixed(1));
  const promPcts = rows.map(r => +(r.prom_pct * 100).toFixed(1));

  mkChart('chart-funil-conv', {
    type: 'bar',
    data: {
      labels: ops,
      datasets: [
        { label: 'Taxa CPC/Alô (%)',     data: cpcPcts,  backgroundColor: COLORS.navy,  borderRadius: 6 },
        { label: 'Taxa Promessa/CPC (%)', data: promPcts, backgroundColor: COLORS.lime,   borderRadius: 6 },
      ]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { position: 'top' } },
      scales: {
        x: { grid: { display: false } },
        y: { grid: { color: '#e8eef5' }, title: { display:true, text:'%' } }
      }
    }
  });

  // Promessas absolutas
  mkChart('chart-funil-prom', {
    type: 'bar',
    data: {
      labels: ops,
      datasets: [{
        label: 'Promessas',
        data: rows.map(r => r.prom),
        backgroundColor: [COLORS.navy, COLORS.green, COLORS.lime, COLORS.greenMid],
        borderRadius: 8,
      }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        x: { grid: { display: false } },
        y: { grid: { color: '#e8eef5' }, title: { display:true, text:'Qtd Promessas' } }
      }
    }
  });
}

// =========================================================
// QUARTIL — Barras por turno
// =========================================================
function buildQuartilChart(turno) {
  const d = DADOS.quartil[turno];
  const id = turno === 'tarde' ? 'chart-quartil-tarde' : 'chart-quartil-manha';
  mkChart(id, {
    type: 'bar',
    data: {
      labels: d.quartis,
      datasets: [{
        label: 'Média de Promessas',
        data: d.mediapromessas,
        backgroundColor: turno === 'tarde'
          ? [COLORS.navy, '#2E7D4F', '#5FAD41', '#8DC642']
          : [COLORS.lime, '#8DC642', '#5FAD41', '#2E7D4F'],
        borderRadius: 8,
      }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false },
        datalabels: { display: false }
      },
      scales: {
        x: { grid: { display: false } },
        y: { grid: { color: '#e8eef5' }, title: { display:true, text:'Média Promessas' } }
      }
    }
  });
}

// =========================================================
// AÇÕES DIGITAIS
// =========================================================
function buildAcoesDigitaisCharts() {
  const wp = DADOS.acoesDigitais.whatsapp;
  const cb = DADOS.acoesDigitais.chatbot;

  // WhatsApp — barras agrupadas por mês
  mkChart('chart-wp', {
    type: 'bar',
    data: {
      labels: wp.operacoes,
      datasets: [
        { label: 'Março', data: wp.marco, backgroundColor: 'rgba(141,198,66,0.4)', borderRadius:4 },
        { label: 'Abril', data: wp.abril, backgroundColor: COLORS.green, borderRadius:4 },
        { label: 'Maio',  data: wp.maio,  backgroundColor: COLORS.navy, borderRadius:4 },
      ]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { position: 'top' } },
      scales: {
        x: { grid: { display: false } },
        y: { grid: { color: '#e8eef5' }, title: { display:true, text:'Envios WhatsApp' } }
      }
    }
  });

  // Chatbot — linhas
  mkChart('chart-chatbot', {
    type: 'line',
    data: {
      labels: ['Março', 'Abril', 'Maio'],
      datasets: cb.operacoes.map((op, i) => ({
        label: op,
        data: [cb.marco[i], cb.abril[i], cb.maio[i]],
        borderColor: [COLORS.navy, COLORS.lime, COLORS.green][i],
        backgroundColor: 'transparent',
        pointRadius: 6, pointBackgroundColor: [COLORS.navy, COLORS.lime, COLORS.green][i],
        borderWidth: 3, tension: 0.4,
      }))
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { position: 'top' } },
      scales: {
        x: { grid: { display: false } },
        y: { grid: { color: '#e8eef5' }, title: { display:true, text:'Envios Chatbot' } }
      }
    }
  });
}

// =========================================================
// INICIALIZA TUDO
// =========================================================
document.addEventListener('DOMContentLoaded', () => {
  buildCapacityChart();
  buildFunilCharts('abrilGeral');
  buildQuartilChart('tarde');
  buildQuartilChart('manha');
  buildAcoesDigitaisCharts();
});
