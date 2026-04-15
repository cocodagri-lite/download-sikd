var XLSX_LOADED = false;

function loadSheetJS(callback) {
  var s = document.createElement('script');
  s.src = 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js';
  s.onload = function() { XLSX_LOADED = true; callback(); };
  s.onerror = function() {
    var s2 = document.createElement('script');
    s2.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
    s2.onload  = function() { XLSX_LOADED = true; callback(); };
    s2.onerror = function() {
      document.getElementById('loading').innerHTML =
        '<p style="color:#a32d2d;padding:1rem">Gagal memuat library. Periksa koneksi lalu refresh.</p>';
    };
    document.head.appendChild(s2);
  };
  document.head.appendChild(s);
}

loadSheetJS(initApp);

function initApp() {
  document.getElementById('loading').style.display = 'none';
  document.getElementById('mainapp').style.display = '';

  var PRESETS = {
    apbd: { path: 'csv_apbd', type: 'apbd', usePeriode: true  },
    tkdd: { path: 'csv',      type: 'tkdd', usePeriode: false }
  };

  var mode   = 'apbd';
  var fmt    = 'xlsx';
  var params = { path:'csv_apbd', type:'apbd', usePeriode:true, periode:'3', tahun:'2025' };

  // Wire buttons
  document.getElementById('btn-apbd').addEventListener('click',    function(){ setMode('apbd'); });
  document.getElementById('btn-tkdd').addEventListener('click',    function(){ setMode('tkdd'); });
  document.getElementById('pill-xls').addEventListener('click',   function(){ setFmt('xls'); });
  document.getElementById('pill-xlsx').addEventListener('click',  function(){ setFmt('xlsx'); });
  document.getElementById('applyBtn').addEventListener('click',   applyParams);
  document.getElementById('dlBtn').addEventListener('click',      startAll);
  document.getElementById('retryBtn').addEventListener('click',   retryFailed);
  document.getElementById('speriode').addEventListener('change',  updatePreview);
  document.getElementById('stahun').addEventListener('change',    updatePreview);
  document.getElementById('provFilter').addEventListener('change', buildTable);
  document.getElementById('popupClose').addEventListener('click', closePopup);
  document.getElementById('popupRetryBtn').addEventListener('click', function(){
    closePopup();
    setTimeout(retryFailed, 300);
  });
  document.getElementById('popupOverlay').addEventListener('click', function(e){
    if (e.target === document.getElementById('popupOverlay')) closePopup();
  });

  function setMode(m) {
    mode = m;
    document.getElementById('btn-apbd').className = (m === 'apbd') ? 'on' : '';
    document.getElementById('btn-tkdd').className = (m === 'tkdd') ? 'on' : '';
    var pr  = PRESETS[m];
    var col = document.getElementById('colperiode');
    var sel = document.getElementById('speriode');
    if (pr.usePeriode) { col.classList.remove('off'); sel.disabled = false; }
    else               { col.classList.add('off');    sel.disabled = true;  }
    updatePreview();
    document.getElementById('oknote').style.display = 'none';
  }

  function setFmt(f) {
    fmt = f;
    document.getElementById('pill-xls').className  = 'pill' + (f === 'xls'  ? ' pxls'  : '');
    document.getElementById('pill-xlsx').className = 'pill' + (f === 'xlsx' ? ' pxlsx' : '');
  }

  function updatePreview() {
    var pr = PRESETS[mode];
    document.getElementById('pvpath').textContent  = pr.path;
    document.getElementById('pvtype').textContent  = pr.type;
    document.getElementById('pvtahun').textContent = document.getElementById('stahun').value;
    var pw = document.getElementById('pvpwrap');
    if (pr.usePeriode) {
      pw.style.display = 'inline';
      document.getElementById('pvperiode').textContent = document.getElementById('speriode').value;
    } else {
      pw.style.display = 'none';
    }
  }

  function applyParams() {
    var pr = PRESETS[mode];
    params = {
      path: pr.path, type: pr.type, usePeriode: pr.usePeriode,
      periode: document.getElementById('speriode').value,
      tahun:   document.getElementById('stahun').value
    };
    document.getElementById('oknote').style.display = 'inline';
    buildTable();
  }

  function buildUrl(prov, pemda) {
    var u = 'https://djpk.kemenkeu.go.id/portal/' + params.path + '?type=' + params.type;
    if (params.usePeriode) u += '&periode=' + params.periode;
    u += '&tahun=' + params.tahun + '&provinsi=' + prov + '&pemda=' + pemda;
    return u;
  }

  function convertToXlsx(arrayBuffer) {
    var data = new Uint8Array(arrayBuffer);
    var wb   = XLSX.read(data, { type: 'array' });
    var out  = XLSX.write(wb, { bookType: 'xlsx', type: 'array', compression: true });
    return new Blob([out], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
  }

  // Province dropdown
  var provFilter = document.getElementById('provFilter');
  var seen = {};
  for (var ci = 0; ci < COMPACT.length; ci++) {
    var prov = COMPACT[ci][3];
    if (!seen[prov]) {
      seen[prov] = true;
      var opt = document.createElement('option');
      opt.value = prov;
      opt.textContent = prov;
      provFilter.appendChild(opt);
    }
  }

  var tbody       = document.getElementById('tbody');
  var retryBtn    = document.getElementById('retryBtn');
  var failCountEl = document.getElementById('failCount');
  var filtered    = [];
  var statusArr   = [];
  var isRunning   = false;

  function setStatus(i, cls, txt) {
    var el = document.getElementById('s' + i);
    if (el) { el.className = 'badge ' + cls; el.textContent = txt; }
    statusArr[i] = cls;
  }

  function buildTable() {
    if (isRunning) return;
    var sel = provFilter.value;
    filtered = sel
      ? COMPACT.filter(function(x){ return x[3] === sel; })
      : COMPACT.slice();

    tbody.innerHTML = '';
    statusArr = [];
    for (var k = 0; k < filtered.length; k++) statusArr.push('bwait');

    var pi  = params.usePeriode ? ' | periode=' + params.periode : '';
    var ext = fmt === 'xlsx' ? '.xlsx' : '.xls';

    document.getElementById('subinfo').textContent =
      filtered.length + ' file | ' + mode.toUpperCase() + ' | tahun=' + params.tahun + pi;
    document.getElementById('dlBtn').textContent =
      '\u2193 Download Semua (' + filtered.length + ' file)';
    document.getElementById('stxt').textContent = '';
    document.getElementById('ptxt').textContent = '';
    document.getElementById('pgwrap').style.display = 'none';
    document.getElementById('pgbar').style.width = '0%';
    retryBtn.style.display = 'none';

    for (var i = 0; i < filtered.length; i++) {
      (function(idx) {
        // COMPACT[idx] = [prov_code, pemda_code, nama_daerah, nama_provinsi]
        var row = filtered[idx];

        var tr = document.createElement('tr');

        // No.
        var tdNo = document.createElement('td');
        tdNo.className = 'center';
        tdNo.textContent = idx + 1;

        // Provinsi (row[3])
        var tdProv = document.createElement('td');
        tdProv.textContent = row[3];

        // Nama Daerah (row[2])
        var tdNama = document.createElement('td');
        tdNama.style.fontWeight = '500';
        tdNama.textContent = row[2];

        // Nama File (row[2] + ext)
        var tdFile = document.createElement('td');
        tdFile.style.color = 'var(--muted)';
        tdFile.style.fontFamily = 'monospace';
        tdFile.style.fontSize = '11px';
        tdFile.textContent = row[2] + ext;

        // Status
        var tdSt  = document.createElement('td');
        var badge = document.createElement('span');
        badge.id = 's' + idx;
        badge.className   = 'badge bwait';
        badge.textContent = 'Menunggu';
        tdSt.appendChild(badge);

        // Aksi
        var tdAct = document.createElement('td');
        var btn   = document.createElement('button');
        btn.className   = 'bone';
        btn.textContent = '\u2193 Unduh';
        btn.addEventListener('click', function(){ doOne(idx); });
        tdAct.appendChild(btn);

        tr.appendChild(tdNo);
        tr.appendChild(tdProv);
        tr.appendChild(tdNama);
        tr.appendChild(tdFile);
        tr.appendChild(tdSt);
        tr.appendChild(tdAct);
        tbody.appendChild(tr);
      })(i);
    }
  }

  buildTable();

  function doDownload(i) {
    var row  = filtered[i];
    var prov = row[0], pemda = row[1], nama = row[2];
    setStatus(i, 'bdl', 'Mengunduh...');

    return fetch(buildUrl(prov, pemda))
      .then(function(r) {
        if (!r.ok) throw new Error('HTTP ' + r.status);
        return r.arrayBuffer();
      })
      .then(function(buf) {
        var blob, ext;
        if (fmt === 'xlsx') {
          setStatus(i, 'bconv', 'Mengonversi...');
          blob = convertToXlsx(buf);
          ext  = '.xlsx';
        } else {
          blob = new Blob([buf]);
          ext  = '.xls';
        }
        var a = document.createElement('a');
        a.href     = URL.createObjectURL(blob);
        a.download = nama + ext;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        setTimeout(function(){ URL.revokeObjectURL(a.href); }, 8000);
        setStatus(i, 'bok', '\u2713 Selesai');
        return true;
      })
      .catch(function() {
        setStatus(i, 'berr', '\u2717 Gagal');
        return false;
      });
  }

  function doOne(i) {
    if (isRunning) return;
    document.getElementById('s' + i).scrollIntoView({ block: 'nearest' });
    doDownload(i).then(updateStats);
  }

  function delay(ms) {
    return new Promise(function(resolve){ setTimeout(resolve, ms); });
  }

  // Popup
  function showPopup(ok, fail, total) {
    var allOk = fail === 0;
    var noOk  = ok === 0;
    document.getElementById('popupIcon').textContent =
      allOk ? '\u2705' : (noOk ? '\u274C' : '\u26A0\uFE0F');
    document.getElementById('popupTitle').textContent =
      allOk ? 'Download Selesai!' : (noOk ? 'Download Gagal' : 'Selesai dengan Catatan');
    document.getElementById('popupTitle').style.color =
      allOk ? 'var(--ok)' : (noOk ? 'var(--err)' : 'var(--warn)');
    document.getElementById('popupTotal').textContent = total;
    document.getElementById('popupOk').textContent    = ok;
    document.getElementById('popupFail').textContent  = fail;
    document.getElementById('popupOkBox').style.background  = 'var(--okbg)';
    document.getElementById('popupOkBox').style.color       = 'var(--ok)';
    document.getElementById('popupFailBox').style.background = fail > 0 ? 'var(--errbg)' : 'var(--bg)';
    document.getElementById('popupFailBox').style.color      = fail > 0 ? 'var(--err)'   : 'var(--muted)';
    document.getElementById('popupRetryWrap').style.display  = fail > 0 ? 'block' : 'none';
    var ov = document.getElementById('popupOverlay');
    ov.style.display = 'flex';
    setTimeout(function() {
      ov.style.opacity = '1';
      document.getElementById('popupBox').style.transform = 'translateY(0)';
    }, 10);
  }

  function closePopup() {
    var ov = document.getElementById('popupOverlay');
    ov.style.opacity = '0';
    document.getElementById('popupBox').style.transform = 'translateY(20px)';
    setTimeout(function(){ ov.style.display = 'none'; }, 250);
  }

  function startAll() {
    if (isRunning) return;
    isRunning = true;
    var btn = document.getElementById('dlBtn');
    btn.disabled = true;
    for (var i = 0; i < filtered.length; i++) {
      if (statusArr[i] !== 'bok') setStatus(i, 'bwait', 'Menunggu');
    }
    retryBtn.style.display = 'none';
    document.getElementById('pgwrap').style.display = 'block';

    var ok = 0, fail = 0, idx = 0;

    function next() {
      if (idx >= filtered.length) {
        document.getElementById('ptxt').textContent = '\u2713 Semua selesai!';
        document.getElementById('pgbar').style.width = '100%';
        btn.textContent = '\u21BA Ulangi Semua';
        btn.disabled = false;
        isRunning = false;
        updateStats();
        showPopup(ok, fail, filtered.length);
        return;
      }
      var i = idx++;
      if (statusArr[i] === 'bok') { ok++; next(); return; }
      document.getElementById('s' + i).scrollIntoView({ block: 'nearest' });
      document.getElementById('ptxt').textContent = i + ' / ' + filtered.length;
      document.getElementById('pgbar').style.width =
        Math.round(i / filtered.length * 100) + '%';
      doDownload(i).then(function(s) {
        if (s) ok++; else fail++;
        document.getElementById('stxt').textContent =
          'Selesai: ' + ok + ' | Gagal: ' + fail;
        delay(700).then(next);
      });
    }
    next();
  }

  function retryFailed() {
    if (isRunning) return;
    isRunning = true;
    document.getElementById('dlBtn').disabled = true;
    retryBtn.disabled = true;
    var failIdx = [];
    for (var i = 0; i < statusArr.length; i++) {
      if (statusArr[i] === 'berr') failIdx.push(i);
    }
    var ok = 0, fail = 0, k = 0;

    function next() {
      if (k >= failIdx.length) {
        document.getElementById('ptxt').textContent =
          '\u2713 Retry selesai! Berhasil: ' + ok + ', Gagal: ' + fail;
        document.getElementById('dlBtn').disabled = false;
        retryBtn.disabled = false;
        isRunning = false;
        updateStats();
        showPopup(ok, fail, failIdx.length);
        return;
      }
      var i = failIdx[k++];
      document.getElementById('s' + i).scrollIntoView({ block: 'nearest' });
      document.getElementById('ptxt').textContent = 'Retry ' + k + '/' + failIdx.length;
      doDownload(i).then(function(s) {
        if (s) ok++; else fail++;
        delay(700).then(next);
      });
    }
    next();
  }

  function updateStats() {
    var f = 0, o = 0;
    for (var i = 0; i < statusArr.length; i++) {
      if (statusArr[i] === 'berr') f++;
      if (statusArr[i] === 'bok')  o++;
    }
    document.getElementById('stxt').textContent = 'Selesai: ' + o + ' | Gagal: ' + f;
    if (f > 0) {
      retryBtn.style.display = 'inline-block';
      failCountEl.textContent = f;
    } else {
      retryBtn.style.display = 'none';
    }
  }

} // end initApp
