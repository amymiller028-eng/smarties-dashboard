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
      renderShareSnippets(v, 'shareGridRefresher');
    } else {
      document.getElementById('refresherView').hidden = true;
      document.getElementById('standardView').hidden = false;
      renderStandard(v);
      renderShareSnippets(v, 'shareGrid');
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

  function bestTestimonialForView() {
    const items = state.data.testimonials || [];
    let pool = items;
    if (state.view !== 'all') {
      const family = state.view.split('-')[0];
      pool = items.filter(t =>
        t.view === state.view || (state.view.endsWith('-summary') && t.view.startsWith(family))
      );
    }
    if (!pool.length) pool = items;
    // Prefer shorter, punchier quotes (under 180 chars) for shareability
    const short = pool.filter(t => t.quote.length <= 180);
    const chosen = (short.length ? short : pool);
    return chosen[Math.floor(Math.random() * chosen.length)];
  }

  function buildStandardSnippets(v) {
    const label = v.label;
    const tb = v.topBox || {};
    const snippets = [];

    snippets.push({
      channel: 'linkedin',
      chip: 'LinkedIn post',
      text:
`${v.nps} NPS. ${tb.applyOnJob}% say they'll apply it on the job. ${v.participants} participants across ${v.clients} client companies can't be wrong.

Our ${label} delivers emotional intelligence training that actually sticks.

#EmotionalIntelligence #LeadershipDevelopment #TalentSmartEQ`
    });

    snippets.push({
      channel: 'email',
      chip: 'Client email',
      text:
`Quick proof point: in our latest ${label} data, ${tb.worthwhileInvestment}% of participants called it a worthwhile investment of their time, and ${tb.applyOnJob}% said they'll apply what they learned on the job. Happy to walk you through what that looks like for a team like yours.`
    });

    snippets.push({
      channel: 'pitch',
      chip: 'Pitch / proposal line',
      text:
`Participants credit ${v.eiDevelopmentAttributed}% of their emotional intelligence growth to our ${label} — with ${v.confidenceInEstimate}% average confidence in that estimate. (n=${v.participants} participants, ${v.sessions} sessions.)`
    });

    const quote = bestTestimonialForView();
    if (quote) {
      snippets.push({
        channel: 'quote',
        chip: 'Shareable quote',
        text:
`"${quote.quote}"

— ${quote.program} participant`
      });
    }

    return snippets;
  }

  function buildRefresherSnippets(v) {
    const before = v.confidenceBefore.toFixed(2);
    const after = v.confidenceAfter.toFixed(2);
    const growth = v.confidenceGrowth.toFixed(2);
    return [
      {
        channel: 'linkedin',
        chip: 'LinkedIn post',
        text:
`Before: ${before}. After: ${after}. That's a +${growth} jump in facilitator confidence after a single Refresher session.

${v.pctMovedUpInConfidence}% of facilitators improved. ${v.pctRatedValuable}% rated the session valuable.

Certification doesn't end at Level 2. #TrainTheTrainer #TalentSmartEQ`
      },
      {
        channel: 'email',
        chip: 'Client email',
        text:
`Our certified facilitators don't just stay sharp — they get sharper. After our latest Refresher, ${v.pctMovedUpInConfidence}% of facilitators moved up on our 4-point confidence scale, with ${v.pctRatedValuable}% rating the session "very" or "extremely" valuable.`
      },
      {
        channel: 'pitch',
        chip: 'Pitch / proposal line',
        text:
`Facilitator confidence rose from ${before} to ${after} on a 4-point scale (+${growth}) after one Refresher session — a measurable, ongoing investment in delivery quality. (n=${v.participants} facilitators.)`
      }
    ];
  }

  function renderShareSnippets(v, targetId) {
    const grid = document.getElementById(targetId);
    if (!grid) return;
    const snippets = v.type === 'refresher' ? buildRefresherSnippets(v) : buildStandardSnippets(v);
    grid.innerHTML = snippets.map(s => `
      <div class="share-card fade-in" data-channel="${s.channel}">
        <div class="share-head-row">
          <span class="share-chip">${escapeHtml(s.chip)}</span>
        </div>
        <p class="share-text">${escapeHtml(s.text)}</p>
        <div class="share-actions">
          <button class="copy-btn" type="button" data-copy="${encodeURIComponent(s.text)}">Copy</button>
        </div>
      </div>
    `).join('');
    grid.querySelectorAll('.copy-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        const text = decodeURIComponent(btn.getAttribute('data-copy'));
        copyToClipboard(text).then(() => {
          btn.classList.add('is-copied');
          btn.textContent = 'Copied!';
          setTimeout(() => {
            btn.classList.remove('is-copied');
            btn.textContent = 'Copy';
          }, 1800);
        });
      });
    });
  }

  function copyToClipboard(text) {
    if (navigator.clipboard && navigator.clipboard.writeText) {
      return navigator.clipboard.writeText(text);
    }
    return new Promise(resolve => {
      const ta = document.createElement('textarea');
      ta.value = text; ta.style.position = 'fixed'; ta.style.opacity = '0';
      document.body.appendChild(ta);
      ta.select();
      try { document.execCommand('copy'); } catch (_) {}
      document.body.removeChild(ta);
      resolve();
    });
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
