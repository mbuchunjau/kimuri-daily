// Shared helpers for both pages.
(function () {
  var PIN_KEY = 'kimuri_pin';

  function getPin() { try { return localStorage.getItem(PIN_KEY) || ''; } catch (e) { return ''; } }
  function setPin(v) { try { localStorage.setItem(PIN_KEY, v); } catch (e) {} }
  function clearPin() { try { localStorage.removeItem(PIN_KEY); } catch (e) {} }

  async function call(action, data, pin) {
    var res = await fetch('/api/gas', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ action: action, pin: pin, data: data || {} })
    });
    var json;
    try { json = await res.json(); } catch (e) { throw new Error('Unexpected reply from the server.'); }
    if (res.status === 401) { var err = new Error('Wrong PIN.'); err.code = 401; throw err; }
    if (!json.ok) throw new Error(json.error || 'Something went wrong.');
    return json.data;
  }

  async function api(action, data) {
    try {
      return await call(action, data, getPin());
    } catch (e) {
      if (e.code === 401) { clearPin(); showLock(function () { location.reload(); }); }
      throw e;
    }
  }

  function showLock(onOk) {
    var el = document.getElementById('lock');
    if (!el) return;
    el.classList.add('on');
    var input = document.getElementById('pin');
    var err = document.getElementById('lockErr');
    var btn = document.getElementById('unlock');
    input.value = ''; err.textContent = '';
    setTimeout(function () { input.focus(); }, 50);
    function submit() {
      var v = input.value.trim();
      if (!v) return;
      btn.disabled = true; err.textContent = '';
      call('ping', {}, v).then(function () {
        setPin(v); el.classList.remove('on'); btn.disabled = false; onOk();
      }).catch(function (e) {
        btn.disabled = false;
        err.textContent = e.code === 401 ? 'Wrong PIN. Try again.' : e.message;
      });
    }
    btn.onclick = submit;
    input.onkeydown = function (ev) { if (ev.key === 'Enter') submit(); };
  }

  // Unlock (if needed), then run the page's init function.
  function ready(init) {
    var pin = getPin();
    if (!pin) return showLock(init);
    call('ping', {}, pin).then(init).catch(function (e) {
      if (e.code === 401) { clearPin(); showLock(init); }
      else { init(); }
    });
  }

  function esc(s) {
    return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; });
  }
  function fmt(n) { return Math.round(n).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ','); }
  function todayIso() {
    var d = new Date();
    return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
  }
  function niceDate(iso) {
    var p = iso.split('-');
    var d = new Date(+p[0], +p[1] - 1, +p[2]);
    return d.toLocaleDateString('en-GB', { weekday: 'short', day: 'numeric', month: 'short', year: 'numeric' });
  }
  function plain(text) { return String(text).replace(/\*/g, ''); }

  window.Kimuri = { api: api, ready: ready, esc: esc, fmt: fmt, todayIso: todayIso, niceDate: niceDate, plain: plain };
})();
