(() => {
  const state = { data: null, view: 'all' };

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

  function escapeHtml(s) {
    return String(s).replace(/[&<>"']/g, c => ({
      '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
    }[c]));
  }

  function flashFadeIn(el) {
    if (!el) return;
    el.classList.remove('fade-in'); void el.offsetWidth; el.classList.add('fade-in');
  }

  function render() {
    const d = state.data;
    const v = d.views[state.view];
    if (!v) return;

    document.getElementById('lastUpdated').textContent = fmtDate(d.meta.lastUpdated);
    document.getElementById('footerDate').textContent = fmtDate(d.meta.lastUpdated);
    document.getElementById('viewLabel').textContent = v.label;

    if (v.type === 'refresher') {
      document.getElementById('standardView').hidden = true;
      document.getElementById('refresherView').hidden = false;
      renderRefresher(v);
    } else {
      document.getElementById('refresherView').hidden = true;
      document.getElementById('standardView').hidden = false;
      renderStandard(v);
    }
    renderTestimonials();
  }

  function renderStandard(v) {
    const npsEl = document.getElementById('npsValue');
    npsEl.textContent = v.nps;
    flashFadeIn(npsEl);
    document.getElementById('npsCaption').textContent = npsCaption(v.nps);
    const pct = ((v.nps + 100) / 200) * 100;
    document.getElementById('npsBar').style.width = pct + '%';

    document.getElementById('eiGrowth').textContent = v.eiDevelopmentAttributed;
    document.getElementById('eiConfidence').textContent = v.confidenceInEstimate;
    document.getElementById('participants').textContent = v.participants;
    document.getElementById('sessions').textContent = v.sessions;

    document.querySelectorAll('[data-metric]').forEach(el => {
      const key = el.getAttribute('data-metric');
      el.textContent = v.topBox[key];
      flashFadeIn(el.closest('.kpi-card'));
    });

    // Manager-expectations tile (only shown when applicable)
    const noManagerCard = document.getElementById('noManagerCard');
    if (typeof v.noManagerExpectationsPct === 'number') {
      noManagerCard.hidden = false;
      document.getElementById('noManagerPct').textContent = v.noManagerExpectationsPct;
      const n = v.managerExpectationsResponses;
      document.getElementById('noManagerN').textContent = n ? ` (n=${n})` : '';
      flashFadeIn(noManagerCard);
    } else {
      noManagerCard.hidden = true;
    }

    // Modality bar
    const total = v.modality.virtual + v.modality.inPerson;
    const vPct = total ? (v.modality.virtual / total) * 100 : 0;
    const iPct = total ? (v.modality.inPerson / total) * 100 : 0;
    document.getElementById('virtualFill').style.width = vPct + '%';
    document.getElementById('inPersonFill').style.width = iPct + '%';
    document.getElementById('virtualLabel').textContent = vPct > 8 ? `Virtual ${Math.round(vPct)}%` : '';
    document.getElementById('inPersonLabel').textContent = iPct > 8 ? `In person ${Math.round(iPct)}%` : '';
  }

  function renderRefresher(v) {
    document.getElementById('confBefore').textContent = v.confidenceBefore.toFixed(2);
    document.getElementById('confAfter').textContent = v.confidenceAfter.toFixed(2);
    flashFadeIn(document.getElementById('confBefore'));
    flashFadeIn(document.getElementById('confAfter'));
    const growth = v.confidenceGrowth;
    const growthEl = document.getElementById('confGrowth');
    growthEl.textContent = growth >= 0
      ? `+${growth.toFixed(2)} levels of growth on a 4-point scale`
      : `${growth.toFixed(2)} levels`;
    document.getElementById('confScale').textContent = v.confidenceScale || '';
    document.getElementById('pctValuable').textContent = v.pctRatedValuable;
    document.getElementById('refParticipants').textContent = v.participants;
    document.getElementById('refSessions').textContent = v.sessions;
    document.getElementById('pctMovedUp').textContent = v.pctMovedUpInConfidence;
  }

  function renderTestimonials() {
    const container = document.getElementById('testimonials');
    if (!container) return;
    const items = state.data.testimonials || [];
    let pool = items;
    if (state.view !== 'all') {
      // Match testimonials whose view matches the current view, or
      // for summary views, anything inside that family.
      const family = state.view.split('-')[0]; // 'ttt', 'private', etc.
      pool = items.filter(t =>
        t.view === state.view || (state.view.endsWith('-summary') && t.view.startsWith(family))
      );
    }
    const toShow = pool.length ? pool : items;
    const facBit = (t) => t.facilitator ? ` &middot; <span class="src-fac">${escapeHtml(t.facilitator)}</span>` : '';
    container.innerHTML = toShow.map(t => `
      <div class="testimonial fade-in">
        <div class="q">&ldquo;${escapeHtml(t.quote)}&rdquo;</div>
        <div class="src">— ${escapeHtml(t.program)}${facBit(t)}</div>
      </div>
    `).join('') || '<div class="testimonial"><div class="q" style="font-style:normal;color:#7a8699">No quotes yet for this view.</div></div>';
  }

  function setActiveTab(view) {
    state.view = view;

    const primaryTabs = document.querySelectorAll('.primary-tabs .tab');
    let activeGroup = null;

    primaryTabs.forEach(btn => {
      const isActive = btn.dataset.view === view
        || (btn.dataset.group && view.startsWith(btn.dataset.group + '-'))
        || (btn.dataset.group === 'ttt' && view.startsWith('ttt-'))
        || (btn.dataset.group === 'private' && view.startsWith('private-'));
      btn.classList.toggle('is-active', isActive);
      btn.setAttribute('aria-selected', isActive ? 'true' : 'false');
      if (isActive && btn.dataset.group) activeGroup = btn.dataset.group;
    });

    document.querySelectorAll('.sub-tabs').forEach(group => {
      const show = group.dataset.group === activeGroup;
      group.hidden = !show;
      if (show) {
        group.querySelectorAll('.sub-tab').forEach(st => {
          st.classList.toggle('is-active', st.dataset.view === view);
        });
      }
    });
  }

  function wireTabs() {
    document.querySelectorAll('.primary-tabs .tab').forEach(btn => {
      btn.addEventListener('click', () => {
        setActiveTab(btn.dataset.view);
        render();
      });
    });
    document.querySelectorAll('.sub-tab').forEach(btn => {
      btn.addEventListener('click', () => {
        setActiveTab(btn.dataset.view);
        render();
      });
    });
  }

  async function init() {
    try {
      await loadData();
      wireTabs();
      setActiveTab('all');
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
