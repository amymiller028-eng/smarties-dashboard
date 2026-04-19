(() => {
  const state = { data: null, view: 'all', chart: null };

  async function loadData() {
    const res = await fetch('data.json?v=' + Date.now());
    if (!res.ok) throw new Error('Could not load data.json');
    state.data = await res.json();
  }

  function fmtDate(iso) {
    const d = new Date(iso + 'T00:00:00');
    return d.toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' });
  }

  function npsCaption(n) {
    if (n >= 70) return 'World-class satisfaction';
    if (n >= 50) return 'Excellent';
    if (n >= 30) return 'Good';
    if (n >= 0) return 'Needs attention';
    return 'Critical';
  }

  function render() {
    const d = state.data;
    const v = d.views[state.view];

    document.getElementById('lastUpdated').textContent = fmtDate(d.meta.lastUpdated);
    document.getElementById('footerDate').textContent = fmtDate(d.meta.lastUpdated);

    const npsEl = document.getElementById('npsValue');
    npsEl.textContent = v.nps;
    npsEl.classList.remove('fade-in'); void npsEl.offsetWidth; npsEl.classList.add('fade-in');
    document.getElementById('npsCaption').textContent = npsCaption(v.nps);
    const pct = ((v.nps + 100) / 200) * 100;
    document.getElementById('npsBar').style.width = pct + '%';

    document.getElementById('eiGrowth').textContent = v.eiDevelopmentAttributed;
    document.getElementById('eiConfidence').textContent = v.confidenceInEstimate;

    document.getElementById('participants').textContent = v.participants;
    document.getElementById('sessions').textContent = v.sessions;
    document.getElementById('clients').textContent = v.clients;

    document.querySelectorAll('[data-metric]').forEach(el => {
      const key = el.getAttribute('data-metric');
      el.textContent = v.topBox[key];
      const card = el.closest('.kpi-card');
      if (card) { card.classList.remove('fade-in'); void card.offsetWidth; card.classList.add('fade-in'); }
    });

    const total = v.modality.virtual + v.modality.inPerson;
    const vPct = total ? (v.modality.virtual / total) * 100 : 0;
    const iPct = total ? (v.modality.inPerson / total) * 100 : 0;
    document.getElementById('virtualFill').style.width = vPct + '%';
    document.getElementById('inPersonFill').style.width = iPct + '%';
    document.getElementById('virtualLabel').textContent = vPct > 8 ? `Virtual ${Math.round(vPct)}%` : '';
    document.getElementById('inPersonLabel').textContent = iPct > 8 ? `In person ${Math.round(iPct)}%` : '';

    renderTestimonials();
    renderTrend();
  }

  function renderTestimonials() {
    const container = document.getElementById('testimonials');
    const items = state.data.testimonials;
    const filtered = state.view === 'all'
      ? items
      : items.filter(t => (state.view === 'ttt' && t.program.includes('Train'))
                      || (state.view === 'private' && t.program.includes('Private')));
    const toShow = (filtered.length ? filtered : items).slice(0, 4);
    container.innerHTML = toShow.map(t => `
      <div class="testimonial fade-in">
        <div class="q">&ldquo;${escapeHtml(t.quote)}&rdquo;</div>
        <div class="src">— ${escapeHtml(t.program)}</div>
      </div>
    `).join('');
  }

  function renderTrend() {
    const ctx = document.getElementById('trendChart').getContext('2d');
    const labels = state.data.trend.map(t => t.label);
    const values = state.data.trend.map(t => t.nps);
    if (state.chart) state.chart.destroy();
    state.chart = new Chart(ctx, {
      type: 'line',
      data: {
        labels,
        datasets: [{
          label: 'NPS',
          data: values,
          borderColor: '#0ACC8B',
          backgroundColor: 'rgba(10,204,139,0.12)',
          pointBackgroundColor: '#002D61',
          pointBorderColor: '#fff',
          pointBorderWidth: 2,
          pointRadius: 6,
          pointHoverRadius: 8,
          fill: true,
          tension: 0.3,
          borderWidth: 3
        }]
      },
      options: {
        responsive: true,
        plugins: {
          legend: { display: false },
          tooltip: {
            backgroundColor: '#002D61',
            padding: 12,
            titleFont: { family: 'Montserrat', weight: '600' },
            bodyFont: { family: 'Montserrat' }
          }
        },
        scales: {
          y: {
            suggestedMin: 0,
            suggestedMax: 100,
            grid: { color: '#eef2f8' },
            ticks: { color: '#7a8699', font: { family: 'Montserrat' } }
          },
          x: {
            grid: { display: false },
            ticks: { color: '#7a8699', font: { family: 'Montserrat' } }
          }
        }
      }
    });
  }

  function escapeHtml(s) {
    return String(s).replace(/[&<>"']/g, c => ({
      '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
    }[c]));
  }

  function wireTabs() {
    document.querySelectorAll('.tab').forEach(btn => {
      btn.addEventListener('click', () => {
        document.querySelectorAll('.tab').forEach(t => {
          t.classList.remove('is-active');
          t.setAttribute('aria-selected', 'false');
        });
        btn.classList.add('is-active');
        btn.setAttribute('aria-selected', 'true');
        state.view = btn.dataset.view;
        render();
      });
    });
  }

  async function init() {
    try {
      await loadData();
      wireTabs();
      render();
    } catch (e) {
      document.body.insertAdjacentHTML('afterbegin',
        `<div style="background:#fb2056;color:#fff;padding:14px;text-align:center;font-family:Montserrat">
          Could not load data.json — make sure it sits next to index.html. (${e.message})
         </div>`);
    }
  }

  init();
})();
