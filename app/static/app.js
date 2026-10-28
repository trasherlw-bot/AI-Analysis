'use strict';
const $ = s => document.querySelector(s);
const esc = s => String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
let pollTimer = null;
let toastTimer;

function toast(msg) {
  const t = $('#toast'); t.textContent = msg; t.classList.add('show');
  clearTimeout(toastTimer); toastTimer = setTimeout(() => t.classList.remove('show'), 3000);
}

async function api(method, path, body, raw) {
  const r = await fetch('api/' + path, {
    method, credentials: 'same-origin',
    headers: body && !raw ? { 'Content-Type': 'application/json' } : {},
    body: raw ? body : body ? JSON.stringify(body) : undefined,
  });
  if (r.status === 401) { showLogin(); throw new Error('auth'); }
  const j = await r.json().catch(() => ({}));
  if (!r.ok) throw new Error(typeof j.detail === 'string' ? j.detail : 'Something went wrong');
  return j;
}
async function run(fn) { try { return await fn(); } catch (e) { if (e.message !== 'auth') toast(e.message === 'Failed to fetch' ? 'No connection' : e.message); } }

const fmtSize = n => n > 1048576 ? (n / 1048576).toFixed(1) + ' MB' : Math.max(1, Math.round(n / 1024)) + ' KB';
const fmtTime = s => new Date(s * 1000).toLocaleString();

function showLogin() { $('#app').hidden = true; $('#login').hidden = false; stopPolling(); }
function showApp() { $('#login').hidden = true; $('#app').hidden = false; refresh(); }

async function boot() {
  try {
    const r = await fetch('api/me', { credentials: 'same-origin' });
    r.ok ? showApp() : showLogin();
  } catch { showLogin(); $('#login-err').textContent = 'Cannot reach the server.'; }
}

$('#login-form').addEventListener('submit', async e => {
  e.preventDefault(); $('#login-err').textContent = '';
  try { await api('POST', 'login', { password: $('#pw').value }); $('#pw').value = ''; showApp(); }
  catch (err) { if (err.message !== 'auth') $('#login-err').textContent = err.message; }
});
$('#logout').addEventListener('click', async () => { await run(() => api('POST', 'logout')); showLogin(); });

async function refresh() {
  const d = await run(() => api('GET', 'files')); if (!d) return;
  $('#warn').hidden = d.configured;
  $('#names-warn').hidden = d.names_masked;
  $('#files').innerHTML = d.files.length ? d.files.map(f => `
    <div class="file">
      <div class="name">${esc(f.name)}<div class="meta">${fmtSize(f.size)} · added ${fmtTime(f.mtime)}${f.result_time ? ' · processed ' + fmtTime(f.result_time) : ''}</div></div>
      <div class="acts">
        <button class="btn primary" data-a="run" data-n="${esc(f.name)}">Run</button>
        ${f.result_time ? `<a class="btn" href="api/result/${encodeURIComponent(f.name)}">Download result</a>` : ''}
        <button class="btn ghost" data-a="del" data-n="${esc(f.name)}">Delete</button>
      </div>
    </div>`).join('') : '<p class="muted small">No files yet. Upload one above.</p>';
  await pollJob();
}

async function pollJob() {
  const j = await run(() => api('GET', 'job')); if (!j) return;
  const st = $('#state'); st.textContent = j.state; st.className = 'pill ' + j.state;
  $('#log').textContent = j.log || (j.state === 'idle' ? 'Nothing running.' : '…');
  const log = $('#log'); log.scrollTop = log.scrollHeight;
  $('#stop').hidden = j.state !== 'running';
  document.querySelectorAll('[data-a=run]').forEach(b => b.disabled = j.state === 'running');
  if (j.state === 'running') { if (!pollTimer) pollTimer = setInterval(pollJob, 2000); }
  else if (pollTimer) { stopPolling(); refresh(); }
}
function stopPolling() { clearInterval(pollTimer); pollTimer = null; }

$('#files').addEventListener('click', async e => {
  const b = e.target.closest('[data-a]'); if (!b) return;
  const name = b.dataset.n;
  if (b.dataset.a === 'run') {
    await run(() => api('POST', 'run', { name, fresh: $('#fresh').checked }));
    await pollJob();
  } else if (b.dataset.a === 'del') {
    if (!confirm('Delete "' + name + '" and its saved results?')) return;
    await run(() => api('DELETE', 'files/' + encodeURIComponent(name)));
    refresh();
  }
});
$('#stop').addEventListener('click', async () => { await run(() => api('POST', 'stop')); pollJob(); });

$('#upload').addEventListener('click', async () => {
  const f = $('#file').files[0];
  if (!f) { toast('Choose an .xlsx file first'); return; }
  $('#upload-msg').textContent = 'Uploading…';
  const r = await run(() => api('PUT', 'files/' + encodeURIComponent(f.name), f, true));
  $('#upload-msg').textContent = r ? 'Uploaded: ' + r.name : '';
  if (r) { $('#file').value = ''; refresh(); }
});

boot();
