(() => {
  // `period` is the selected year ('2025' | '2026' | 'all'). It defaults to
  // whatever the data says, so the year control can be added without changing
  // what anyone sees today.
  const state = { data: null, view: 'all', period: null, band: 'All' };

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

  // "Jane, SYF" / "Jane" / "" depending on what's on file for the quote.
  function personLabel(t) {
    const name = (t.name || '').trim();
    const company = (t.company || '').trim();
    if (name && company) return `${name}, ${company}`;
    return name || company || '';
  }

  // Leading "<name> · " for a quote's source line (empty when no name on file).
  function personBit(t) {
    const label = personLabel(t);
    return label ? `<span class="src-person">${escapeHtml(label)}</span> &middot; ` : '';
  }

  // Quotes carry the year they came from. Show only the selected period —
  // entries with no `period` (older data.json) always pass.
  function testimonialsForPeriod() {
    const all = state.data.testimonials || [];
    if (state.period === 'all') return all;
    return all.filter(t => !t.period || t.period === state.period);
  }

  function flashFadeIn(el) {
    if (!el) return;
    el.classList.remove('fade-in'); void el.offsetWidth; el.classList.add('fade-in');
  }

  // Views that vary by year live in viewsByPeriod. Refresher and LTF don't —
  // they're their own instruments on their own timelines — so they fall back
  // to the flat `views` map.
  function viewFor(period, key) {
    const byPeriod = state.data.viewsByPeriod && state.data.viewsByPeriod[period];
    return (byPeriod && byPeriod[key]) || null;
  }

  function isPeriodic(key) {
    const vbp = state.data.viewsByPeriod || {};
    return Object.keys(vbp).some(p => vbp[p][key]);
  }

  function periodHasData(period, key) {
    const v = viewFor(period, key);
    return !!v && v.participants > 0;
  }

  function availablePeriods() {
    return ((state.data.meta && state.data.meta.periods) || []).concat('all');
  }

  function renderPeriodControls() {
    const bar = document.getElementById('periodBar');
    const host = document.getElementById('periodControls');
    if (!bar || !host) return;

    if (!isPeriodic(state.view)) { bar.hidden = true; return; }
    bar.hidden = false;

    host.innerHTML = availablePeriods().map(p => {
      const has = periodHasData(p, state.view);
      const label = p === 'all' ? 'All time' : p;
      return `<button type="button" class="period-btn${p === state.period ? ' is-active' : ''}"` +
             ` data-period="${p}"${has ? '' : ' disabled title="No data for this program in ' + label + '"'}>` +
             `${label}</button>`;
    }).join('');

    host.querySelectorAll('.period-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        state.period = btn.dataset.period;
        render();
      });
    });
  }

  function render() {
    const d = state.data;

    // Selecting a program with nothing in the current year would render an
    // empty page — fall back to All time rather than showing zeros.
    if (isPeriodic(state.view) && !periodHasData(state.period, state.view)) {
      const fallback = availablePeriods().find(p => periodHasData(p, state.view));
      if (fallback) state.period = fallback;
    }

    const v = viewFor(state.period, state.view) || d.views[state.view];
    if (!v) return;
    renderPeriodControls();

    document.getElementById('lastUpdated').textContent = fmtDate(d.meta.lastUpdated);
    document.getElementById('footerDate').textContent = fmtDate(d.meta.lastUpdated);
    document.getElementById('viewLabel').textContent =
      isPeriodic(state.view)
        ? `${v.label} · ${state.period === 'all' ? 'All time' : state.period}`
        : v.label;

    const panes = {
      standard: document.getElementById('standardView'),
      ltf: document.getElementById('ltfView'),
      learner: document.getElementById('learnerView'),
      refresher: document.getElementById('refresherView')
    };
    const which = panes[v.type] ? v.type : 'standard';
    Object.entries(panes).forEach(([k, el]) => { if (el) el.hidden = k !== which; });

    if (which === 'refresher') {
      renderRefresher(v);
      renderShareSnippets(v, 'shareGridRefresher');
      renderTestimonials('testimonials');
    } else if (which === 'ltf') {
      renderLtf(v);
      renderShareSnippets(v, 'shareGridLtf');
      renderTestimonials('ltfTestimonials');
    } else if (which === 'learner') {
      renderLearner(v);
    } else {
      renderStandard(v);
      renderShareSnippets(v, 'shareGrid');
      renderTestimonials('testimonials');
    }
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

    renderTalkTracks(v);

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

    renderManagerSplit(v);

    renderModality(v.modality, {
      virtualFill: 'virtualFill', inPersonFill: 'inPersonFill',
      virtualLabel: 'virtualLabel', inPersonLabel: 'inPersonLabel'
    });
  }

  // Only appears where the program actually asks the manager question.
  function renderManagerSplit(v) {
    const panel = document.getElementById('mgrSplit');
    if (!panel) return;
    const s = v.managerSplit;
    if (!s) { panel.hidden = true; return; }
    panel.hidden = false;

    document.getElementById('mgrWithNps').textContent = s.withExpectations.nps;
    document.getElementById('mgrWithN').textContent = s.withExpectations.n;
    document.getElementById('mgrWithoutNps').textContent = s.without.nps;
    document.getElementById('mgrWithoutN').textContent = s.without.n;
    document.getElementById('mgrGap').textContent = (s.gap >= 0 ? '+' : '') + s.gap;

    const total = s.withExpectations.n + s.without.n;
    const share = total ? Math.round(s.without.n / total * 100) : 0;
    document.getElementById('mgrNote').textContent =
      `${share}% of participants said no expectations were set. This is a correlation, not proof of ` +
      `cause — but it points at the pre-work conversation rather than the program itself.`;
    flashFadeIn(panel);
  }

  // The EQ-growth card's "how to use this" popover: explains the two numbers
  // and gives one conservative, client-ready line (attribution × confidence).
  function talkTrack(key, v) {
    if (key !== 'ei') return '';
    const ei = v.eiDevelopmentAttributed;
    const conf = v.confidenceInEstimate;
    const adj = Math.round(ei * conf / 100);
    return (
      `<span class="pop-body">How do you put a number on something as soft as emotional intelligence? You ask participants two things — how much of their growth they credit to the program (<strong>${ei}%</strong>), and how confident they are in that estimate (<strong>${conf}%</strong>) — then multiply the two for a <em>confidence-adjusted</em> figure of about <strong>${adj}%</strong>.<br><br>` +
      `This isn’t something we invented. Multiplying an estimate by its confidence level is a standard step in the <strong>Phillips ROI Methodology</strong> — the framework L&amp;D and HR teams across the industry (Fortune 500s, government, universities) use to credibly isolate the impact of training. It’s how the field puts defensible numbers on “soft” skills.</span>` +
      `<span class="pop-key">Bottom line: lead with “<strong>${ei}% of their EQ growth is credited to this program.</strong>” Most clients are happy with that. Keep the ${adj}% in your back pocket for anyone who wants the conservative math.</span>`
    );
  }

  function renderTalkTracks(v) {
    document.querySelectorAll('#standardView .info-pop[data-track]').forEach(el => {
      el.innerHTML = talkTrack(el.getAttribute('data-track'), v);
    });
  }

  // Shared by the standard and LTF views — both carry a { virtual, inPerson } count.
  function renderModality(modality, ids) {
    const total = modality.virtual + modality.inPerson;
    const vPct = total ? (modality.virtual / total) * 100 : 0;
    const iPct = total ? (modality.inPerson / total) * 100 : 0;
    document.getElementById(ids.virtualFill).style.width = vPct + '%';
    document.getElementById(ids.inPersonFill).style.width = iPct + '%';
    document.getElementById(ids.virtualLabel).textContent = vPct > 8 ? `Virtual ${Math.round(vPct)}%` : '';
    document.getElementById(ids.inPersonLabel).textContent = iPct > 8 ? `In person ${Math.round(iPct)}%` : '';
  }

  function renderLtf(v) {
    const npsEl = document.getElementById('ltfNps');
    npsEl.textContent = v.nps;
    flashFadeIn(npsEl);
    document.getElementById('ltfNpsCaption').textContent = npsCaption(v.nps);
    document.getElementById('ltfNpsBar').style.width = ((v.nps + 100) / 200) * 100 + '%';

    document.getElementById('ltfAdvocacy').textContent = v.advocacy;
    document.getElementById('ltfAdvocacyStrong').textContent = v.advocacyStrongly;
    document.getElementById('ltfParticipants').textContent = v.participants;
    document.getElementById('ltfSessions').textContent = v.sessions;
    document.getElementById('ltfClients').textContent = v.clients;

    const total = v.statements.length;
    document.getElementById('ltfStatementsSub').textContent =
      `${v.unanimous} of ${total} statements drew agreement from every participant. ` +
      `Bars show the share who picked "strongly agree" — the top of the scale.`;

    document.getElementById('ltfStatements').innerHTML = v.statements.map(s => `
      <div class="statement-row">
        <div class="statement-text">${escapeHtml(s.text)}</div>
        <div class="statement-bar"><span style="width:${s.strongly}%"></span></div>
        <div class="statement-nums">
          <b>${s.strongly}%</b>strongly agree
          <em>${s.top2}% agreed &middot; avg ${s.mean.toFixed(2)}/5</em>
        </div>
      </div>`).join('');

    renderModality(v.modality, {
      virtualFill: 'ltfVirtualFill', inPersonFill: 'ltfInPersonFill',
      virtualLabel: 'ltfVirtualLabel', inPersonLabel: 'ltfInPersonLabel'
    });
  }

  // --- 90-day impact -------------------------------------------------------
  // Positions on the slope track use the real 1-5 scale, so a +0.32 shift looks
  // like what it is. The colored segment carries the change.
  const SCALE_MIN = 1, SCALE_MAX = 5;
  const scalePct = v => ((v - SCALE_MIN) / (SCALE_MAX - SCALE_MIN)) * 100;

  function rankRows(items, klass) {
    if (!items || !items.length) return '<div class="dist-row"><div class="dist-label" style="color:#7a8699">No responses yet.</div></div>';
    const max = Math.max.apply(null, items.map(i => i.count));
    return items.map(i => `
      <div class="${klass}-row">
        <div class="${klass}-label">${escapeHtml(i.label)}</div>
        <div class="${klass}-bar"><span style="width:${(i.count / max) * 100}%"></span></div>
        <div class="${klass}-value">${i.pct !== undefined ? i.pct + '%' : i.count}</div>
      </div>`).join('');
  }

  function renderBandControls(v) {
    const host = document.getElementById('bandControls');
    if (!host) return;
    if (!state.band || !v.byBand[state.band]) state.band = 'All';
    host.innerHTML = v.bands.map(b =>
      `<button type="button" class="period-btn${b === state.band ? ' is-active' : ''}" data-band="${escapeHtml(b)}">` +
      `${escapeHtml(b)} <span style="opacity:.55">(${v.byBand[b].respondents})</span></button>`
    ).join('');
    host.querySelectorAll('.period-btn').forEach(btn => {
      btn.addEventListener('click', () => { state.band = btn.dataset.band; render(); });
    });
  }

  function renderLearner(v) {
    renderBandControls(v);
    const d = v.byBand[state.band] || v.byBand.All;

    const before = document.getElementById('lrnBefore');
    before.textContent = d.beforeAvg.toFixed(2);
    document.getElementById('lrnAfter').textContent = d.afterAvg.toFixed(2);
    flashFadeIn(before);

    const shift = d.afterAvg - d.beforeAvg;
    document.getElementById('lrnShift').textContent =
      `${shift >= 0 ? '+' : ''}${shift.toFixed(2)} across ${d.behaviors.length} leadership behaviors, on a 5-point scale`;
    document.getElementById('lrnFootnote').textContent =
      `${d.improvedAnyPct}% improved on at least one behavior · ${d.pairedRespondents} people answered both halves` +
      (d.pairedRespondents < 25 ? ' — a small sample, read as directional' : '');

    document.getElementById('lrnApplying').textContent = d.appliesOftenPct;
    document.getElementById('lrnN').textContent = d.respondents;
    document.getElementById('lrnPaired').textContent = d.pairedRespondents;
    document.getElementById('lrnNps').textContent = d.nps === null ? '—' : d.nps;

    document.getElementById('lrnSlopeSub').textContent =
      'Each row is one behavior, rated before the program and again 90 days later. ' +
      'Sorted by how much it moved.';

    document.getElementById('lrnSlopes').innerHTML = d.behaviors.map(b => {
      const a = scalePct(Math.min(b.before, b.after));
      const w = Math.abs(scalePct(b.after) - scalePct(b.before));
      return `
      <div class="slope-row">
        <div class="slope-text">${escapeHtml(b.text)}${b.reversed ? '<span class="rev" title="Agreeing is the negative answer, so this is inverted">inverted</span>' : ''}</div>
        <div class="slope-track">
          <span class="slope-seg" style="left:${a}%; width:${w}%"></span>
          <span class="slope-dot before" style="left:${scalePct(b.before)}%" title="Before: ${b.before}"></span>
          <span class="slope-dot after" style="left:${scalePct(b.after)}%" title="After: ${b.after}"></span>
        </div>
        <div class="slope-change">${b.change >= 0 ? '+' : ''}${b.change.toFixed(2)}</div>
        <div class="slope-improved">${b.improvedPct}% up</div>
      </div>`;
    }).join('') + `
      <div class="slope-legend">
        <span><i class="legend-dot before"></i>Before training</span>
        <span><i class="legend-dot after"></i>90 days later</span>
        <span style="color:#7a8699">Scale runs 1 (strongly disagree) to 5 (strongly agree)</span>
      </div>`;

    document.getElementById('lrnApplies').innerHTML = rankRows(d.applies, 'dist');
    document.getElementById('lrnManager').innerHTML = rankRows(d.managerReinforcement, 'dist');

    const block = name => (d.blocks.find(b => b.label === name) || {}).items || [];
    document.getElementById('lrnOutcomes').innerHTML = rankRows(block('Business outcomes improved'), 'rank');
    document.getElementById('lrnAreas').innerHTML = rankRows(block('Where it helped most'), 'rank');
    document.getElementById('lrnBarriers').innerHTML = rankRows(block('Barriers to applying it'), 'rank');
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

  function renderTestimonials(targetId) {
    const container = document.getElementById(targetId || 'testimonials');
    if (!container) return;
    const items = testimonialsForPeriod();
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
        <div class="src">— ${personBit(t)}${escapeHtml(t.program)}${facBit(t)}</div>
      </div>
    `).join('') || '<div class="testimonial"><div class="q" style="font-style:normal;color:#7a8699">No quotes yet for this view.</div></div>';
  }

  function bestTestimonialForView() {
    const items = testimonialsForPeriod();
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
      const who = personLabel(quote) || `${quote.program} participant`;
      snippets.push({
        channel: 'quote',
        chip: 'Shareable quote',
        text:
`"${quote.quote}"

— ${who}`
      });
    }

    return snippets;
  }

  // LTF has no EQ-attribution question, so its snippets lead on unanimity
  // and advocacy instead of the confidence-adjusted growth figure.
  function buildLtfSnippets(v) {
    const total = v.statements.length;
    const top = v.statements[0];
    const snippets = [
      {
        channel: 'linkedin',
        chip: 'LinkedIn post',
        text:
`${v.nps} NPS. Zero detractors. ${v.advocacy}% would put their own team through it.

${v.unanimous} of our ${total} end-of-session statements drew agreement from every single participant in Leading Through Friction.

Friction isn't the problem. It's the opportunity.

#LeadershipDevelopment #EmotionalIntelligence #TalentSmartEQ`
      },
      {
        channel: 'email',
        chip: 'Client email',
        text:
`Quick proof point from Leading Through Friction: ${v.advocacy}% of participants said they'd want their team or other leaders in their organization to go through it. ${top.strongly}% strongly agreed with "${top.text}." NPS came in at ${v.nps} with no detractors. Happy to walk you through what that looks like for your leaders.`
      },
      {
        channel: 'pitch',
        chip: 'Pitch / proposal line',
        text:
`Leading Through Friction scored an NPS of ${v.nps} with zero detractors, and every participant agreed on ${v.unanimous} of ${total} impact statements — including that they'd send their own team. (n=${v.participants} leaders, ${v.sessions} session${v.sessions === 1 ? '' : 's'}.)`
      }
    ];

    const quote = bestTestimonialForView();
    if (quote) {
      const who = personLabel(quote) || `${quote.program} participant`;
      snippets.push({
        channel: 'quote', chip: 'Shareable quote',
        text: `"${quote.quote}"\n\n— ${who}`
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
    const snippets = v.type === 'refresher' ? buildRefresherSnippets(v)
                   : v.type === 'ltf' ? buildLtfSnippets(v)
                   : buildStandardSnippets(v);
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

  // Hide any tab whose view isn't in data.json. LTF's Public and Train the
  // Trainer sub-tabs stay hidden until those sheets exist, then appear on
  // their own with no code change.
  function pruneNav() {
    const have = state.data.views || {};
    document.querySelectorAll('.tab, .sub-tab').forEach(btn => {
      btn.hidden = !have[btn.dataset.view];
    });
  }

  async function init() {
    try {
      await loadData();
      state.period = (state.data.meta && state.data.meta.defaultPeriod) || 'all';
      pruneNav();
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
