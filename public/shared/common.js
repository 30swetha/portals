/* ==========================================================================
   Mobileum Caller ID Platform - Shared Common Utilities
   ========================================================================== */

const $ = id => document.getElementById(id);

const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({
  '&': '&amp;',
  '<': '&lt;',
  '>': '&gt;',
  '"': '&quot;',
  "'": '&#39;'
}[c]));

const peso = n => '₱' + Number(n || 0).toLocaleString('en-PH', {
  minimumFractionDigits: 2,
  maximumFractionDigits: 2
});

const initials = n => String(n || '?')
  .split(/\s+/)
  .filter(Boolean)
  .slice(0, 2)
  .map(w => w[0].toUpperCase())
  .join('');

async function api(url, opt) {
  const r = await fetch(url, opt);
  let j = null;
  try { j = await r.json(); } catch (e) {}
  if (!r.ok) {
    const e = new Error((j && j.error) || r.statusText);
    e.data = j;
    throw e;
  }
  return j;
}

// Mobileum signature dots loader HTML
function dotsLoader() {
  return '<div class="mobileum-dots" title="Loading..."><span></span><span></span><span></span></div>';
}

// Toast notification helper
function showToast(title, message, type = 'info', duration = 4000) {
  let container = document.getElementById('toastContainer');
  if (!container) {
    container = document.createElement('div');
    container.id = 'toastContainer';
    container.className = 'toast-container';
    document.body.appendChild(container);
  }

  const icons = {
    success: '✓',
    error: '✕',
    info: 'ℹ',
    warning: '⚠'
  };

  const toast = document.createElement('div');
  toast.className = `toast ${type}`;
  toast.innerHTML = `
    <div class="toast-icon">${icons[type] || 'ℹ'}</div>
    <div class="toast-content">
      <div class="toast-title">${esc(title)}</div>
      ${message ? `<div class="toast-message">${esc(message)}</div>` : ''}
    </div>
    <button class="toast-close" aria-label="Close">&times;</button>
  `;

  toast.querySelector('.toast-close').onclick = () => {
    toast.classList.remove('show');
    setTimeout(() => toast.remove(), 200);
  };

  container.appendChild(toast);
  // Trigger animation frame
  requestAnimationFrame(() => {
    toast.classList.add('show');
  });

  if (duration > 0) {
    setTimeout(() => {
      if (toast.parentElement) {
        toast.classList.remove('show');
        setTimeout(() => toast.remove(), 250);
      }
    }, duration);
  }
  return toast;
}

// Modern enterprise "what the customer will see" card
function callerCard(o) {
  const lg = o.logoUrl
    ? `<img src="${esc(o.logoUrl)}" alt="${esc(o.brandName || 'Logo')}" onerror="this.onerror=null;this.src='/shared/assets/mobileum-logo.png';">`
    : esc(initials(o.brandName));
  const isVerified = o.verified !== false;
  return `
    <div class="cid">
      <div class="lg">${lg}</div>
      <div style="flex:1;min-width:0;">
        <div class="nm">${esc(o.brandName)}</div>
        <div class="sub">${esc(o.department || '')}${o.department && o.display ? ' &middot; ' : ''}${esc(o.display || '')}</div>
        ${o.tagline ? `<div class="sub" style="font-style:italic;opacity:0.85;">"${esc(o.tagline)}"</div>` : ''}
        ${isVerified ? '<div class="vf"><svg width="12" height="12" viewBox="0 0 24 24" fill="currentColor"><path d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41z"/></svg> Verified business</div>' : '<div style="color:#F59E0B;font-size:11.5px;margin-top:3px;">⚠ Unverified preview</div>'}
      </div>
    </div>
  `;
}

// Shows where the data is saved (the shared Excel workbook)
function dbStatus(el, s) {
  if (!el) return;
  const name = (s.database || '').split(/[\\/]/).pop() || 'CallerID-Database.xlsx';
  el.innerHTML = s.pending > 0
    ? `<span class="err"><strong>${s.pending}</strong> change(s) waiting to be written into <b>${esc(name)}</b> (close it in Excel if open)</span>`
    : `<span style="color:var(--status-approved)">●</span> Shared Excel database: <b>${esc(name)}</b>` +
      (s.lastFlush ? ` &middot; last saved ${new Date(s.lastFlush).toLocaleTimeString('en-PH', { timeZone: 'Asia/Manila' })}` : '');
}

// In-app non-blocking confirmation modal dialog
function showConfirmModal(opts) {
  const {
    title = 'Confirm Removal',
    message = 'Are you sure you want to proceed?',
    confirmText = 'Remove',
    confirmClass = 'no',
    onConfirm
  } = opts || {};

  let modal = document.getElementById('mobileumConfirmModal');
  if (!modal) {
    modal = document.createElement('div');
    modal.id = 'mobileumConfirmModal';
    modal.className = 'confirm-modal-backdrop';
    document.body.appendChild(modal);
  }

  modal.innerHTML = `
    <div class="confirm-modal-box">
      <div class="confirm-modal-header">
        <h3 class="confirm-modal-title">${esc(title)}</h3>
        <button type="button" class="confirm-modal-close" id="btnConfirmClose" title="Cancel">&times;</button>
      </div>
      <div class="confirm-modal-body">
        <p style="margin:0;font-size:13.5px;color:var(--deep-navy);line-height:1.5;">${esc(message)}</p>
      </div>
      <div class="confirm-modal-footer">
        <button type="button" class="btn sm" id="btnConfirmCancel">Cancel</button>
        <button type="button" class="btn sm ${esc(confirmClass)}" id="btnConfirmOk">${esc(confirmText)}</button>
      </div>
    </div>
  `;

  modal.classList.add('open');

  const close = () => {
    modal.classList.remove('open');
  };

  modal.querySelector('#btnConfirmClose').onclick = close;
  modal.querySelector('#btnConfirmCancel').onclick = close;
  modal.onclick = e => {
    if (e.target === modal) close();
  };

  modal.querySelector('#btnConfirmOk').onclick = () => {
    close();
    if (typeof onConfirm === 'function') {
      try { onConfirm(); } catch (err) { console.error(err); }
    }
  };
}

// Cross-port shared removal state for packages, numbers, and requests
function getSharedList(key) {
  let list = [];
  try {
    const match = document.cookie.match(new RegExp('(?:^|;\\s*)' + key + '=([^;]*)'));
    if (match && match[1]) {
      list = JSON.parse(decodeURIComponent(match[1]));
    }
  } catch (e) {}

  if (!Array.isArray(list) || !list.length) {
    try {
      const stored = localStorage.getItem(key);
      if (stored) list = JSON.parse(stored);
    } catch (e) {}
  }
  return Array.isArray(list) ? list : [];
}

function addToSharedList(key, item) {
  const current = getSharedList(key);
  if (!current.includes(item)) {
    current.push(item);
  }
  const str = encodeURIComponent(JSON.stringify(current));
  document.cookie = `${key}=${str}; path=/; max-age=2592000; SameSite=Lax`;
  try { document.cookie = `${key}=${str}; path=/; domain=localhost; max-age=2592000; SameSite=Lax`; } catch (e) {}
  try { localStorage.setItem(key, JSON.stringify(current)); } catch (e) {}
  return current;
}

function removeFromSharedList(key, item) {
  let current = getSharedList(key);
  current = current.filter(x => x !== item);
  const str = encodeURIComponent(JSON.stringify(current));
  document.cookie = `${key}=${str}; path=/; max-age=2592000; SameSite=Lax`;
  try { document.cookie = `${key}=${str}; path=/; domain=localhost; max-age=2592000; SameSite=Lax`; } catch (e) {}
  try { localStorage.setItem(key, JSON.stringify(current)); } catch (e) {}
  return current;
}

