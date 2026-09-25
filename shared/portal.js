/* ═══════════════════════════════════════════════
   HOLIDAY HEROES CONNECT — Shared Portal JS
   ═══════════════════════════════════════════════ */

(function () {
  'use strict';

  // ─── Wait for config, then overlay with API data ───
  window.addEventListener('DOMContentLoaded', () => {
    const staticCfg = window.PORTAL_CONFIG;

    // Client ID: from URL param first, then static config
    const urlClient = new URLSearchParams(window.location.search).get('client');
    const clientId  = urlClient || staticCfg?.client?.id;

    if (!clientId) { console.error('No client ID in URL or PORTAL_CONFIG'); return; }

    // If server pre-loaded the config (SSR), use it directly — fast & reliable
    if (window._portalServerCfg) {
      _runPortal(_buildCfgFromApi(clientId, window._portalServerCfg));
      return;
    }

    const apiBase = window.BACKEND_URL_OVERRIDE || '';
    fetch(apiBase + '/api/portal-config/' + clientId)
      .then(r => r.ok ? r.json() : null)
      .then(api => {
        if (api) {
          const cfg = staticCfg
            ? _mergeApiCfg(staticCfg, api)
            : _buildCfgFromApi(clientId, api);
          _runPortal(cfg);
        } else if (staticCfg) {
          _runPortal(staticCfg);
        } else {
          console.error('No portal config found for client: ' + clientId);
        }
      })
      .catch(() => { if (staticCfg) _runPortal(staticCfg); });
  });

  function _buildCfgFromApi(clientId, api) {
    const branding = api.branding || {};
    return {
      client: {
        id:          clientId,
        name:        branding.name        || clientId,
        fullName:    branding.fullName    || '',
        accentColor: branding.accentColor || '#0b8f82',
        defaultLang: branding.defaultLang || 'en',
      },
      languages:      ['en', 'tr'],
      welcomeTitle:   api.welcomeTitle   || '',
      welcomeMessage: api.welcomeMessage || '',
      emailTemplates: [],
      videos:    _mapApiVideos(api.videos   || []),
      scenarios: api.scenarios || [],
      meetings:  _mapApiMeetings(api.meetings || []),
      bugsLink:    api.bugsLink    || '',
      contactLink: api.contactLink || '',
    };
  }

  function _mergeApiCfg(base, api) {
    const out = Object.assign({}, base);
    // Branding from admin
    if (api.branding) {
      out.client = Object.assign({}, base.client, {
        name:        api.branding.name        || base.client.name,
        fullName:    api.branding.fullName    || base.client.fullName,
        accentColor: api.branding.accentColor || base.client.accentColor,
        defaultLang: api.branding.defaultLang || base.client.defaultLang,
      });
    }
    // Welcome text
    if (api.welcomeTitle)   out.welcomeTitle   = api.welcomeTitle;
    if (api.welcomeMessage) out.welcomeMessage = api.welcomeMessage;
    // Videos
    if (api.videos && api.videos.length) out.videos = _mapApiVideos(api.videos);
    // Scenarios
    if (api.scenarios && api.scenarios.length) out.scenarios = api.scenarios;
    // Meetings
    if (api.meetings && api.meetings.length) out.meetings = _mapApiMeetings(api.meetings);
    // Links
    if (api.bugsLink)    out.bugsLink    = api.bugsLink;
    if (api.contactLink) out.contactLink = api.contactLink;
    return out;
  }

  function _mapApiVideos(videos) {
    return videos.map(v => ({
      id:          v.id,
      category:    { en: 'Training', tr: 'Eğitim' },
      title:       { en: v.title || '', tr: v.title || '' },
      duration:    '',
      src:         v.src || v.uploadedFile || '',
      description: v.description || '',
    }));
  }

  function _mapApiMeetings(meetings) {
    return meetings.map(m => ({
      id:           m.id,
      date:         m.date || '',
      dateLabel:    m.dateLabel || { en: m.date || '', tr: m.date || '' },
      title:        m.title    || { en: '', tr: '' },
      participants: m.participants || 0,
      file:         m.file || '',
    }));
  }

  function _runPortal(cfg) {
    applyClientConfig(cfg);
    buildNav(cfg);
    buildVideos(cfg);
    buildScenarios(cfg);
    buildMeetings(cfg);
    buildTranslation(cfg);
    _applyDynamicLinks(cfg);
    initLang(cfg);
    initNav();
  }

  function _applyDynamicLinks(cfg) {
    if (cfg.bugsLink)    document.querySelectorAll('.js-bugs-link').forEach(el => { el.href = cfg.bugsLink; });
    if (cfg.contactLink) document.querySelectorAll('.js-contact-link').forEach(el => { el.href = cfg.contactLink; });
  }

  // ─── Apply client branding ─────────────────────
  function applyClientConfig(cfg) {
    const root = document.documentElement;
    if (cfg.client.accentColor) {
      root.style.setProperty('--accent', cfg.client.accentColor);
      root.style.setProperty('--accent-light', hexToLight(cfg.client.accentColor));
    }
    // Titles
    setText('[data-client-name]', cfg.client.name);
    setText('[data-client-full]', cfg.client.fullName);
    // Welcome banner (API-driven)
    if (cfg.welcomeTitle)   setText('.welcome-title', cfg.welcomeTitle);
    if (cfg.welcomeMessage) setText('.welcome-sub',   cfg.welcomeMessage);
  }

  function hexToLight(hex) {
    // Returns a very light tint of the accent color for backgrounds
    const r = parseInt(hex.slice(1,3),16);
    const g = parseInt(hex.slice(3,5),16);
    const b = parseInt(hex.slice(5,7),16);
    return `rgba(${r},${g},${b},0.1)`;
  }

  // ─── Navigation ────────────────────────────────
  function buildNav(cfg) {
    const container = document.getElementById('nav-items');
    if (!container) return;

    const items = [
      { id:'videos',      icon:'🎬', en:'Training Videos',  tr:'Eğitim Videoları',  badge: cfg.videos.length },
      { id:'scenarios',   icon:'✅', en:'Test Scenarios',   tr:'Test Senaryoları',  badge: null },
      { id:'translation', icon:'📸', en:'UI Translation',   tr:'Arayüz Çevirisi',   badge: null },
      { id:'meetings',    icon:'📅', en:'Meeting Recaps',   tr:'Toplantı Özetleri', badge: (cfg.meetings||[]).length || null },
      { id:'feedback',    icon:'💬', en:'Feedback',         tr:'Geri Bildirim',      badge: null },
    ];

    items.forEach((item, i) => {
      const btn = document.createElement('button');
      btn.className = 'nav-item';
      btn.dataset.target = item.id;
      btn.setAttribute('data-en', item.en);
      btn.setAttribute('data-tr', item.tr);
      btn.innerHTML = `
        <span class="nav-icon">${item.icon}</span>
        <span class="nav-label">${item.en}</span>
        ${item.badge ? `<span class="nav-badge">${item.badge}</span>` : ''}
      `;
      container.appendChild(btn);
    });
  }

  function initNav() {
    document.addEventListener('click', e => {
      const btn = e.target.closest('.nav-item[data-target]');
      if (!btn) return;
      const target = btn.dataset.target;

      document.querySelectorAll('.nav-item').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');

      document.querySelectorAll('.page-section').forEach(s => s.classList.remove('active'));
      const section = document.getElementById('section-' + target);
      if (section) section.classList.add('active');

      const label = btn.querySelector('.nav-label');
      setText('[data-section-title]', label ? label.textContent : '');

      // Close sidebar on mobile after navigation
      closeMobileSidebar();
      // Show back button on mobile
      document.getElementById('topbar-back-btn')?.classList.add('active');
    });

    // Clicking the sidebar brand goes back to the overview
    document.querySelectorAll('.sidebar-brand').forEach(el => {
      el.style.cursor = 'pointer';
      el.addEventListener('click', () => {
        window.goHome && window.goHome();
        closeMobileSidebar();
      });
    });

    // Hamburger button
    const hamburger = document.getElementById('hamburger-btn');
    const overlay   = document.getElementById('sidebar-overlay');
    const sidebar   = document.querySelector('.portal-sidebar');

    if (hamburger) {
      hamburger.addEventListener('click', () => {
        sidebar.classList.toggle('mobile-open');
        overlay.classList.toggle('mobile-open');
      });
    }
    if (overlay) {
      overlay.addEventListener('click', closeMobileSidebar);
    }
  }

  function closeMobileSidebar() {
    document.querySelector('.portal-sidebar')?.classList.remove('mobile-open');
    document.getElementById('sidebar-overlay')?.classList.remove('mobile-open');
  }

  window.goHome = function() {
    document.querySelectorAll('.nav-item').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.page-section').forEach(s => s.classList.remove('active'));
    const overview = document.getElementById('section-overview');
    if (overview) overview.classList.add('active');
    setText('[data-section-title]', '');
    // Hide back button
    document.getElementById('topbar-back-btn')?.classList.remove('active');
  };

  // ─── Videos ────────────────────────────────────
  function buildVideos(cfg) {
    const grid = document.getElementById('video-grid');
    if (!grid) return;

    grid.style.cssText = 'display:flex;flex-direction:column;gap:10px;';

    cfg.videos.forEach((v, idx) => {
      const hasVideo = !!v.src;
      const num = idx + 1;

      const card = document.createElement('div');
      card.className = 'recap-card';
      card.style.cursor = hasVideo ? 'pointer' : 'default';
      card.dataset.en = v.title.en;
      card.dataset.tr = v.title.tr;

      const meta = [
        `<span data-en="${v.category.en}" data-tr="${v.category.tr}">${v.category.en}</span>`,
        v.duration ? `⏱ ${v.duration}` : '',
        !hasVideo ? '<em>Coming soon</em>' : '',
      ].filter(Boolean).join(' &nbsp;·&nbsp; ');

      card.innerHTML = `
        <div class="recap-date-block">
          <div class="recap-date-day">${num}</div>
        </div>
        <div class="recap-info">
          <div class="recap-title" data-en="${v.title.en}" data-tr="${v.title.tr}">${v.title.en}</div>
          ${v.description ? `<div class="recap-desc">${v.description}</div>` : ''}
          <div class="recap-meta">${meta}</div>
        </div>
        <div class="recap-arrow" style="${hasVideo ? '' : 'opacity:0.3;'}">▶</div>
      `;

      if (hasVideo) {
        card.addEventListener('click', () => openVideoModal(v));
      }

      grid.appendChild(card);
    });
  }

  function openVideoModal(v) {
    const modal = document.getElementById('video-modal');
    const title = document.getElementById('modal-video-title');
    const frame = document.getElementById('modal-iframe');
    const lang  = getCurrentLang();

    title.textContent = v.title[lang] || v.title.en;
    frame.src = v.src;
    modal.classList.add('open');
    document.body.style.overflow = 'hidden';
  }

  window.openVideoModal = openVideoModal;

  window.closeVideoModal = function () {
    const modal = document.getElementById('video-modal');
    const frame = document.getElementById('modal-iframe');
    modal.classList.remove('open');
    frame.src = '';
    document.body.style.overflow = '';
  };

  // ─── Email Templates ───────────────────────────
  const EMAIL_CAT_LABELS = {
    'platform':           { en: 'Platform Onboarding — Video Automation', tr: 'Platform Tanıtımı — Video Otomasyonu' },
    'first-booking':      { en: 'First Booking Encouragement',            tr: 'İlk Rezervasyon Teşviki' },
    'after-registration': { en: 'After Registration',                     tr: 'Kayıt Sonrası' },
    'after-booking':      { en: 'After Booking',                          tr: 'Rezervasyon Sonrası' },
  };

  function getEmailCategory(tpl) {
    if (tpl.category) return tpl.category;
    const num = parseInt((tpl.id || '').replace('email-', ''));
    if (num >= 1 && num <= 7) return 'platform';
    if (num >= 8) return 'first-booking';
    return 'other';
  }

  function buildEmailCard(tpl, container) {
    const div = document.createElement('div');
    div.className = 'card email-card';

    // PDF or per-language PNG screenshot
    let mediaSection = '';
    if (tpl.screenshot) {
      mediaSection = `<div class="email-pdf-wrap">
        <iframe src="${tpl.screenshot}" title="${tpl.title.en}"></iframe>
      </div>`;
    } else if (tpl.screenshotEn || tpl.screenshotTr) {
      const en = tpl.screenshotEn || tpl.screenshotTr;
      const tr = tpl.screenshotTr || tpl.screenshotEn;
      mediaSection = `<div class="email-pdf-wrap">
        <img id="otp-img-${tpl.id}" src="${en}" data-en="${en}" data-tr="${tr}" style="width:100%;border-radius:6px;" alt="${tpl.title.en}">
      </div>`;
    }

    const hasBody = tpl.body && (tpl.body.en || tpl.body.tr);

    div.innerHTML = `
      <div class="card-head">
        <div class="card-icon">📧</div>
        <div>
          <div class="card-title" data-en="${tpl.title.en}" data-tr="${tpl.title.tr}">${tpl.title.en}</div>
          <div class="card-sub">
            <strong data-en="Subject:" data-tr="Konu:">Subject:</strong>
            <span data-en="${tpl.subject.en}" data-tr="${tpl.subject.tr}"> ${tpl.subject.en}</span>
          </div>
        </div>
      </div>
      <div class="email-layout">
        ${mediaSection}
        ${hasBody ? `<div style="display:flex; flex-direction:column;">
          <div class="email-preview" id="preview-${tpl.id}" data-en="${encodeText(tpl.body.en)}" data-tr="${encodeText(tpl.body.tr)}">${tpl.body.en}</div>
          <div class="email-actions">
            <button class="btn btn-ghost btn-sm" onclick="toggleEmail('${tpl.id}')">
              <span data-en="Expand" data-tr="Genişlet" id="expand-label-${tpl.id}">Expand</span>
            </button>
            <button class="btn btn-primary btn-sm" onclick="copyEmail('${tpl.id}')">
              <span id="copy-label-${tpl.id}" data-en="Copy template" data-tr="Şablonu kopyala">Copy template</span>
            </button>
          </div>
        </div>` : (tpl.note ? `<div style="display:flex;align-items:center;padding:16px;background:#f0fdf9;border-radius:10px;gap:10px;align-self:center;">
          <span style="font-size:18px;">ℹ️</span>
          <span style="font-size:13px;color:#0b6e63;" data-en="${tpl.note.en}" data-tr="${tpl.note.tr}">${tpl.note.en}</span>
        </div>` : '')}
      </div>
    `;
    container.appendChild(div);
  }

  function buildEmailTemplates(cfg) {
    const container = document.getElementById('email-list');
    if (!container) return;

    // Group by category preserving order of first appearance
    const order = [];
    const groups = {};
    cfg.emailTemplates.forEach(tpl => {
      const cat = getEmailCategory(tpl);
      if (!groups[cat]) { groups[cat] = []; order.push(cat); }
      groups[cat].push(tpl);
    });

    order.forEach(cat => {
      const label = EMAIL_CAT_LABELS[cat];
      if (label) {
        const divider = document.createElement('div');
        divider.className = 'divider-label';
        divider.setAttribute('data-en', label.en);
        divider.setAttribute('data-tr', label.tr);
        divider.textContent = label.en;
        container.appendChild(divider);
      }
      groups[cat].forEach(tpl => buildEmailCard(tpl, container));
    });
  }

  function encodeText(str) {
    return str.replace(/"/g, '&quot;').replace(/\n/g, '&#10;');
  }

  window.toggleEmail = function (id) {
    const el    = document.getElementById('preview-' + id);
    const label = document.getElementById('expand-label-' + id);
    const lang  = getCurrentLang();
    const isExp = el.classList.toggle('expanded');
    label.textContent = isExp
      ? (lang === 'tr' ? 'Daralt' : 'Collapse')
      : (lang === 'tr' ? 'Genişlet' : 'Expand');
  };

  window.copyEmail = function (id) {
    const el   = document.getElementById('preview-' + id);
    const lang = getCurrentLang();
    const text = el.getAttribute('data-' + lang) || el.textContent;
    const decoded = text.replace(/&#10;/g, '\n').replace(/&quot;/g, '"');
    navigator.clipboard.writeText(decoded).then(() => {
      const label = document.getElementById('copy-label-' + id);
      label.textContent = '✓ Copied!';
      label.classList.add('copy-success');
      setTimeout(() => {
        label.textContent = lang === 'tr' ? 'Şablonu kopyala' : 'Copy template';
        label.classList.remove('copy-success');
      }, 2000);
    });
  };

  // ─── Test Scenarios ────────────────────────────
  function buildScenarios(cfg) {
    const container = document.getElementById('scenarios-list');
    if (!container) return;

    let total = 0, done = 0;
    const storageKey = 'portal_scenarios_' + cfg.client.id;
    const saved = JSON.parse(localStorage.getItem(storageKey) || '{}');

    cfg.scenarios.forEach((group, gi) => {
      const div = document.createElement('div');
      div.className = 'card';
      div.style.marginBottom = '16px';

      div.innerHTML = `
        <div class="card-head">
          <div class="card-icon">📋</div>
          <div>
            <div class="card-title" data-en="${group.group.en}" data-tr="${group.group.tr}">${group.group.en}</div>
            <div class="card-sub" id="group-progress-${gi}"></div>
          </div>
        </div>
        <div id="group-items-${gi}"></div>
      `;

      const itemsContainer = div.querySelector(`#group-items-${gi}`);
      const groupItems = group.items;
      let groupDone = 0;

      groupItems.forEach((item, ii) => {
        const key  = `${gi}_${ii}`;
        const isDone = !!saved[key];
        if (isDone) { groupDone++; done++; }
        total++;

        const row = document.createElement('div');
        row.className = 'accordion';

        row.innerHTML = `
          <div class="check-item" style="padding: 12px 22px;">
            <div class="check-box${isDone ? ' done' : ''}" data-key="${key}" onclick="toggleCheck(this, '${cfg.client.id}')"></div>
            <div class="check-text${isDone ? ' done-text' : ''}" data-en="${item.en}" data-tr="${item.tr}">${item.en}</div>
          </div>
        `;

        itemsContainer.appendChild(row);
      });

      // Group progress
      updateGroupProgress(gi, groupDone, groupItems.length);
      container.appendChild(div);
    });

    updateOverallProgress(done, total);
  }

  function updateGroupProgress(gi, done, total) {
    const el = document.getElementById('group-progress-' + gi);
    if (el) el.textContent = `${done}/${total} completed`;
  }

  function updateOverallProgress(done, total) {
    const el   = document.getElementById('overall-progress');
    const fill = document.getElementById('progress-fill');
    if (el)   el.textContent = `${done} of ${total} checks completed`;
    if (fill) fill.style.width = `${Math.round(done/total*100)}%`;
  }

  window.toggleCheck = function (box, clientId) {
    const key      = box.dataset.key;
    const storageKey = 'portal_scenarios_' + clientId;
    const saved    = JSON.parse(localStorage.getItem(storageKey) || '{}');
    const isDone   = box.classList.toggle('done');
    const textEl   = box.nextElementSibling;

    if (isDone)  { saved[key] = true;  if (textEl) textEl.classList.add('done-text'); }
    else         { delete saved[key];  if (textEl) textEl.classList.remove('done-text'); }

    localStorage.setItem(storageKey, JSON.stringify(saved));

    // Recount
    const allBoxes = document.querySelectorAll('.check-box');
    let total = allBoxes.length, done = 0;
    allBoxes.forEach(b => { if (b.classList.contains('done')) done++; });
    updateOverallProgress(done, total);
  };

  // ─── Meeting Recaps ────────────────────────────
  function buildMeetings(cfg) {
    const list = document.getElementById('meetings-list');
    if (!list) return;

    cfg.meetings.forEach(m => {
      const d     = new Date(m.date);
      const day   = d.getDate();
      const month = d.toLocaleString('en', { month: 'short' }).toUpperCase();
      const card  = document.createElement('div');
      card.className = 'recap-card';
      card.style.cursor = 'pointer';
      card.innerHTML = `
        <div class="recap-date-block">
          <div class="recap-date-day">${day}</div>
          <div class="recap-date-month">${month}</div>
        </div>
        <div class="recap-info">
          <div class="recap-title" data-en="${m.title.en}" data-tr="${m.title.tr}">${m.title.en}</div>
          <div class="recap-meta">
            <span data-en="${m.dateLabel.en}" data-tr="${m.dateLabel.tr}">${m.dateLabel.en}</span>
            &nbsp;·&nbsp; ${m.participants} participants
          </div>
        </div>
        <div class="recap-arrow">→</div>
      `;
      card.addEventListener('click', () => openMeetingModal(m.file));
      list.appendChild(card);
    });
  }

  function openMeetingModal(file) {
    const modal  = document.getElementById('meeting-modal');
    const iframe = document.getElementById('meeting-iframe');
    if (!modal || !iframe) return;
    iframe.src = file;
    modal.style.display = 'flex';
    document.body.style.overflow = 'hidden';
  }

  window.closeMeetingModal = function () {
    const modal  = document.getElementById('meeting-modal');
    const iframe = document.getElementById('meeting-iframe');
    if (!modal) return;
    modal.style.display = 'none';
    iframe.src = '';
    document.body.style.overflow = '';
  };

  // ─── Screenshot Translation ────────────────────
  function buildTranslation(cfg) {
    // Hotspots defined inline in HTML via data-* attributes
    // This function handles the popup logic
    const storageKey = 'portal_translations_' + cfg.client.id;
    window._translations = JSON.parse(localStorage.getItem(storageKey) || '{}');

    document.addEventListener('click', e => {
      // Close popup on outside click
      const popup = document.querySelector('.translation-popup.open');
      if (popup && !popup.contains(e.target) && !e.target.closest('.hotspot')) {
        popup.classList.remove('open');
        document.querySelectorAll('.hotspot').forEach(h => h.classList.remove('active'));
      }
    });
  }

  window.openHotspot = function (el) {
    const key   = el.dataset.key;
    const orig  = el.dataset.orig;
    const label = el.dataset.label;
    const popup = document.getElementById('popup-' + key);
    const stored = window._translations[key] || '';

    // Close others
    document.querySelectorAll('.translation-popup.open').forEach(p => p.classList.remove('open'));
    document.querySelectorAll('.hotspot.active').forEach(h => h.classList.remove('active'));

    if (popup) {
      popup.querySelector('textarea').value = stored;
      popup.classList.add('open');
      el.classList.add('active');
      popup.querySelector('textarea').focus();
    }
  };

  window.saveTranslation = function (key, clientId) {
    const popup = document.getElementById('popup-' + key);
    const value = popup.querySelector('textarea').value.trim();
    const storageKey = 'portal_translations_' + clientId;

    window._translations[key] = value;
    localStorage.setItem(storageKey, JSON.stringify(window._translations));
    popup.classList.remove('open');
    document.querySelectorAll('.hotspot').forEach(h => h.classList.remove('active'));
    renderTranslationPanel();
  };

  window.renderTranslationPanel = function () {
    const panel = document.getElementById('translation-rows');
    if (!panel) return;
    panel.innerHTML = '';

    const entries = Object.entries(window._translations || {}).filter(([,v]) => v);
    if (!entries.length) {
      panel.innerHTML = '<div style="padding:16px 22px;font-size:13px;color:var(--ink-muted)" data-en="No translations saved yet." data-tr="Henüz çeviri kaydedilmedi.">No translations saved yet.</div>';
      return;
    }

    entries.forEach(([key, value]) => {
      const hs = document.querySelector(`.hotspot[data-key="${key}"]`);
      const orig  = hs ? hs.dataset.orig  : key;
      const label = hs ? hs.dataset.label : key;
      const row = document.createElement('div');
      row.className = 'translation-row';
      row.innerHTML = `<div class="tr-key">${label}</div><div class="tr-orig">${orig}</div><div class="tr-value">${value}</div>`;
      panel.appendChild(row);
    });
  };

  window.exportTranslations = function (clientId) {
    const data = JSON.parse(localStorage.getItem('portal_translations_' + clientId) || '{}');
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `translations_${clientId}_${new Date().toISOString().slice(0,10)}.json`;
    a.click();
  };

  // ─── Language System ───────────────────────────
  function getCurrentLang() {
    return document.documentElement.lang || 'en';
  }

  function initLang(cfg) {
    const defaultLang = cfg.client.defaultLang || 'en';
    setLang(defaultLang);

    document.querySelectorAll('.lang-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        const lang = btn.dataset.lang;
        if (lang) setLang(lang);
      });
    });
  }

  window.setLang = function (lang) {
    document.documentElement.lang = lang;

    document.querySelectorAll('.lang-btn').forEach(b => {
      b.classList.toggle('active', b.dataset.lang === lang);
    });

    document.querySelectorAll('[data-en]').forEach(el => {
      const val = el.getAttribute('data-' + lang);
      if (val !== null) el.textContent = val;
    });

    // Email previews use encoded text
    document.querySelectorAll('.email-preview[data-en]').forEach(el => {
      const val = el.getAttribute('data-' + lang);
      if (val !== null) el.textContent = val.replace(/&#10;/g, '\n').replace(/&quot;/g, '"');
    });

    // Per-language email images (OTP etc.)
    document.querySelectorAll('img[data-en][data-tr]').forEach(img => {
      const src = img.getAttribute('data-' + lang);
      if (src) img.src = src;
    });
  };

  // ─── Accordion ─────────────────────────────────
  window.toggleAccordion = function (btn) {
    btn.closest('.accordion').classList.toggle('open');
  };

  // ─── Helpers ───────────────────────────────────
  function setText(selector, text) {
    document.querySelectorAll(selector).forEach(el => { el.textContent = text; });
  }

})();
