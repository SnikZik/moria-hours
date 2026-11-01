(function () {
  var $ = function (id) { return document.getElementById(id); };
  var el = { status: $('status'), clock: $('clock'), detail: $('detail'), toast: $('toast'), start: $('start'), pause: $('pause'), end: $('end') };

  // The key arrives once in the install link. Stored so the app keeps working if iOS drops the query.
  var key = new URLSearchParams(location.search).get('k');
  try {
    if (key) localStorage.setItem('hours_key', key);
    else key = localStorage.getItem('hours_key');
  } catch (e) {}

  var current = null;
  var clockSkew = 0;
  var busy = false;

  function call(method, action) {
    var url = window.HOURS_API;
    var options = { redirect: 'follow' };
    if (method === 'GET') {
      url += '?key=' + encodeURIComponent(key) + '&t=' + Date.now();
    } else {
      options.method = 'POST';
      options.headers = { 'Content-Type': 'text/plain;charset=utf-8' };
      options.body = JSON.stringify({ key: key, action: action });
    }
    return fetch(url, options).then(function (r) { return r.json(); });
  }

  function pad(n) { return (n < 10 ? '0' : '') + n; }

  function duration(ms) {
    var s = Math.max(0, Math.floor(ms / 1000));
    return Math.floor(s / 3600) + ':' + pad(Math.floor(s / 60) % 60) + ':' + pad(s % 60);
  }

  function hoursText(minutes) {
    return Math.floor(minutes / 60) + ':' + pad(minutes % 60);
  }

  function tick() {
    if (!current) return;
    var now = Date.now() + clockSkew;
    if (current.state === 'working' && current.startedAt) {
      el.clock.textContent = duration(now - current.startedAt - current.breakMinutes * 60000);
    } else if (current.state === 'break') {
      el.clock.textContent = duration(now - current.breakStartedAt);
    }
  }

  function render(s) {
    current = s;
    clockSkew = s.now - Date.now();
    var isOff = s.state === 'off';
    var isBreak = s.state === 'break';

    el.start.disabled = !isOff;
    el.pause.disabled = isOff;
    el.end.disabled = isOff;
    el.pause.textContent = isBreak ? 'חזרתי מהפסקה' : 'יצאתי להפסקה';

    if (isBreak) {
      el.status.textContent = 'בהפסקה';
      el.detail.textContent = 'התחלת ב ' + s.startLabel;
    } else if (!isOff) {
      el.status.textContent = 'בעבודה';
      el.detail.textContent = 'התחלת ב ' + s.startLabel + (s.breakMinutes ? ', הפסקות ' + s.breakMinutes + ' דקות' : '');
    } else if (s.last) {
      el.status.textContent = 'סיימת להיום';
      el.clock.textContent = hoursText(s.last.netMinutes);
      el.detail.textContent = s.last.start + ' עד ' + s.last.end + (s.last.breakMinutes ? ', הפסקות ' + s.last.breakMinutes + ' דקות' : '');
    } else {
      el.status.textContent = 'לא בעבודה';
      el.clock.textContent = '0:00';
      el.detail.textContent = '';
    }
    tick();
  }

  function toast(text, isError) {
    el.toast.textContent = text;
    el.toast.className = 'show' + (isError ? ' err' : '');
    clearTimeout(toast.timer);
    if (!isError) toast.timer = setTimeout(function () { el.toast.className = ''; el.toast.textContent = ''; }, 3500);
  }

  function showMessage(text) {
    $('note-text').textContent = text;
    $('note').hidden = false;
    $('note-close').focus();
  }

  function setBusy(on) {
    busy = on;
    document.body.classList.toggle('busy', on);
  }

  function press(action, doneText) {
    if (busy) return;
    setBusy(true);
    toast('שולח');
    call('POST', action).then(function (s) {
      if (s.error === 'unauthorized') throw new Error('unauthorized');
      if (s.state) render(s);
      if (s.message) showMessage(s.message);
      if (s.error === 'invalid_state') toast('המצב כבר התעדכן, בדקי שזה נכון');
      else if (s.error) toast('לא נשמר, נסי שוב', true);
      else toast(doneText + ' ' + new Date(s.now).toLocaleTimeString('he-IL', { hour: '2-digit', minute: '2-digit', timeZone: 'Asia/Jerusalem' }));
    }).catch(function (e) {
      toast(e.message === 'unauthorized' ? 'הקישור לא תקין, פני לשניר' : 'לא נשמר, בדקי אינטרנט ונסי שוב', true);
    }).then(function () { setBusy(false); });
  }

  function refresh() {
    if (busy || !key) return;
    call('GET').then(function (s) {
      if (s.error === 'unauthorized') { el.status.textContent = 'הקישור לא תקין'; return; }
      if (s.state) { render(s); if (el.toast.classList.contains('err')) { el.toast.className = ''; el.toast.textContent = ''; } }
    }).catch(function () {
      if (!current) el.status.textContent = 'אין חיבור';
      toast('אין חיבור לאינטרנט. שום דבר לא נשמר עד שיחזור', true);
    });
  }

  $('note-close').addEventListener('click', function () { $('note').hidden = true; });
  el.start.addEventListener('click', function () { press('start', 'נרשמה התחלה'); });
  el.pause.addEventListener('click', function () {
    if (current && current.state === 'break') press('break_end', 'נרשמה חזרה');
    else press('break_start', 'נרשמה הפסקה');
  });
  el.end.addEventListener('click', function () {
    if (window.confirm('לסיים את יום העבודה?')) press('end', 'נרשם סיום');
  });

  if (!key) {
    el.status.textContent = 'חסר קישור התקנה';
    el.detail.textContent = 'פני לשניר לקבלת הקישור';
    return;
  }

  setInterval(tick, 1000);
  document.addEventListener('visibilitychange', function () { if (!document.hidden) refresh(); });
  refresh();
  if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js').catch(function () {});
})();
