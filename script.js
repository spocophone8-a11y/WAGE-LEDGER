(function(){
  const DAY_MS = 86400000;
  const SUPPORTS_FS = 'showOpenFilePicker' in window;

  // Local-calendar date helpers (deliberately not UTC-based — see the
  // streak-tracker fix this app's date handling follows the same pattern).
  function isoDate(d){
    const y = d.getFullYear();
    const m = String(d.getMonth()+1).padStart(2,'0');
    const day = String(d.getDate()).padStart(2,'0');
    return y+'-'+m+'-'+day;
  }
  function todayISO(){ const d = new Date(); d.setHours(0,0,0,0); return isoDate(d); }
  function fmtDate(iso){
    const [y,m,d] = iso.split('-').map(Number);
    return new Date(y,m-1,d).toLocaleDateString(undefined,{ day:'2-digit', month:'short', year:'numeric' });
  }
  function num(v){ const n = parseFloat(v); return isNaN(n) ? 0 : n; }
  function money(n){ return (Math.round(n*100)/100).toLocaleString(); }

  let labourers = [];   // [{name, mobile, notes}]
  let entries = [];     // [{date, name, work, wage, paid, debtGiven, debtRepaid, notes}]
  let openCards = new Set();

  let fileHandle = null;
  let connectedName = null;

  const statusEl = document.getElementById('statusMsg');
  const gate = document.getElementById('gate');
  const toolbarConnected = document.getElementById('toolbarConnected');
  const toolbarFallback = document.getElementById('toolbarFallback');
  const connLabel = document.getElementById('connLabel');
  const footerNote = document.getElementById('footerNote');
  const addLabourerPanel = document.getElementById('addLabourerPanel');
  const newName = document.getElementById('newName');
  const newMobile = document.getElementById('newMobile');
  const newNotes = document.getElementById('newNotes');
  const addLabourerBtn = document.getElementById('addLabourerBtn');
  const labourerList = document.getElementById('labourerList');
  const searchBox = document.getElementById('searchBox');
  let searchQuery = '';
  const totalsPanel = document.getElementById('totalsPanel');
  const totLabourers = document.getElementById('totLabourers');
  const totWage = document.getElementById('totWage');
  const totPaid = document.getElementById('totPaid');
  const totDue = document.getElementById('totDue');
  const totDebt = document.getElementById('totDebt');

  function setStatus(msg, warn){
    statusEl.textContent = msg || '';
    statusEl.classList.toggle('warn', !!warn);
  }
  function isConnected(){ return SUPPORTS_FS ? !!fileHandle : connectedName !== null; }

  function refreshChrome(){
    const connected = isConnected();
    gate.classList.toggle('hidden', connected);
    toolbarConnected.classList.toggle('hidden', !(connected && SUPPORTS_FS));
    toolbarFallback.classList.toggle('hidden', SUPPORTS_FS);
    if(connected && SUPPORTS_FS) connLabel.textContent = 'Connected — ' + connectedName;
    [newName,newMobile,newNotes,addLabourerBtn].forEach(el => el.disabled = !connected);
    searchBox.disabled = !connected;
    footerNote.innerHTML = SUPPORTS_FS
      ? '<strong>Live-connected:</strong> every labourer, work entry, payment, or debt change writes straight back to the workbook on your computer — the Labourers and Entries tabs, with live totals, stay in sync. Open the file in Excel any time to see the current state.'
      : '<strong>Note:</strong> this browser doesn\'t support connecting directly to a local file, so use Load / Download instead — Chrome or Edge on this same page would let you connect live.';
  }
  refreshChrome();

  function labourerStats(name){
    const rows = entries.filter(e => e.name === name);
    const totalWage = rows.reduce((s,e)=>s+e.wage,0);
    const totalPaid = rows.reduce((s,e)=>s+e.paid,0);
    const totalDue = totalWage - totalPaid;
    const debtGiven = rows.reduce((s,e)=>s+e.debtGiven,0);
    const debtRepaid = rows.reduce((s,e)=>s+e.debtRepaid,0);
    const netDebt = debtGiven - debtRepaid;
    const days = new Set(rows.map(e => e.date)).size; // unique dates worked, not entry count
    return { totalWage, totalPaid, totalDue, netDebt, days };
  }

  function renderTotals(){
    const connected = isConnected();
    totalsPanel.classList.toggle('hidden', !connected || labourers.length === 0);
    if(!connected || labourers.length === 0) return;
    // Grand total across every labourer, regardless of the current search filter.
    const wage = entries.reduce((s,e)=>s+e.wage,0);
    const paid = entries.reduce((s,e)=>s+e.paid,0);
    const debtGiven = entries.reduce((s,e)=>s+e.debtGiven,0);
    const debtRepaid = entries.reduce((s,e)=>s+e.debtRepaid,0);
    totLabourers.textContent = labourers.length;
    totWage.textContent = '₹' + money(wage);
    totPaid.textContent = '₹' + money(paid);
    totDue.textContent = '₹' + money(wage - paid);
    totDebt.textContent = '₹' + money(debtGiven - debtRepaid);
  }

  async function persist(){
    if(SUPPORTS_FS && fileHandle){
      try{
        const wb = buildWorkbook();
        const wbout = XLSX.write(wb, {bookType:'xlsx', type:'array'});
        const writable = await fileHandle.createWritable();
        await writable.write(wbout);
        await writable.close();
        setStatus('Saved to ' + connectedName + ' just now.');
      }catch(err){
        setStatus('Could not save to the file: ' + err.message, true);
      }
    } else if(!SUPPORTS_FS){
      setStatus(connectedName ? 'Remember to click "Download wage_ledger.xlsx" to keep this change.' : 'Load a file first, or use Download to start one.');
    }
  }

  function render(){
    const connected = isConnected();
    renderTotals();
    labourerList.innerHTML = '';
    if(!connected){
      return;
    }
    if(labourers.length === 0){
      labourerList.innerHTML = '<div class="empty">No labourers on the ledger yet. Add one above.</div>';
      return;
    }
    const q = searchQuery.trim().toLowerCase();
    const visible = q
      ? labourers.filter(l => l.name.toLowerCase().includes(q) || (l.mobile||'').toLowerCase().includes(q))
      : labourers;
    if(visible.length === 0){
      labourerList.innerHTML = '<div class="empty">No labourer matches "' + searchQuery.trim() + '".</div>';
      return;
    }
    visible.forEach(lab => {
      const stats = labourerStats(lab.name);
      const card = document.createElement('div');
      card.className = 'labourer-card' + (openCards.has(lab.name) ? ' open' : '');

      const head = document.createElement('div');
      head.className = 'card-head';
      head.innerHTML = `
        <div class="id-block">
          <span class="chevron">▸</span>
          <div class="name-mobile">
            <h3></h3>
            <a class="mobile-link"></a>
            <div class="notes-text"></div>
          </div>
        </div>
        <div class="stat-strip">
          <div class="stat"><span class="label">Days</span><span class="val"></span></div>
          <div class="stat paid"><span class="label">Paid</span><span class="val"></span></div>
          <div class="stat due"><span class="label">Due</span><span class="val"></span></div>
          <div class="stat debt"><span class="label">Debt</span><span class="val"></span></div>
        </div>
      `;
      head.querySelector('h3').textContent = lab.name;
      const mLink = head.querySelector('.mobile-link');
      if(lab.mobile){ mLink.href = 'tel:' + lab.mobile; mLink.textContent = lab.mobile; }
      else { mLink.removeAttribute('href'); mLink.textContent = 'no mobile number'; }
      head.querySelector('.notes-text').textContent = lab.notes || '';
      const vals = head.querySelectorAll('.stat .val');
      vals[0].textContent = stats.days;
      vals[1].textContent = '₹' + money(stats.totalPaid);
      vals[2].textContent = '₹' + money(stats.totalDue);
      vals[3].textContent = '₹' + money(stats.netDebt);
      head.addEventListener('click', () => {
        if(openCards.has(lab.name)) openCards.delete(lab.name);
        else openCards.add(lab.name);
        render();
      });
      card.appendChild(head);

      const body = document.createElement('div');
      body.className = 'card-body';

      const form = document.createElement('div');
      form.className = 'entry-form';
      form.innerHTML = `
        <div><label>Date</label><input type="date" class="f-date"></div>
        <div><label>Work done</label><input type="text" class="f-work" placeholder="e.g. brick laying"></div>
        <div><label>Wage ₹</label><input type="number" class="f-wage" min="0" step="0.01"></div>
        <div><label>Paid ₹</label><input type="number" class="f-paid" min="0" step="0.01"></div>
        <div><label>Debt given ₹</label><input type="number" class="f-debtgiven" min="0" step="0.01"></div>
        <div><label>Debt repaid ₹</label><input type="number" class="f-debtrepaid" min="0" step="0.01"></div>
        <div><label>Notes</label><input type="text" class="f-notes" placeholder="optional"></div>
        <button class="f-add">+ Add entry</button>
      `;
      const fDate = form.querySelector('.f-date');
      fDate.value = todayISO();
      form.querySelector('.f-add').addEventListener('click', async () => {
        const date = fDate.value || todayISO();
        const work = form.querySelector('.f-work').value.trim();
        const wage = num(form.querySelector('.f-wage').value);
        const paid = num(form.querySelector('.f-paid').value);
        const debtGiven = num(form.querySelector('.f-debtgiven').value);
        const debtRepaid = num(form.querySelector('.f-debtrepaid').value);
        const notes = form.querySelector('.f-notes').value.trim();
        if(!work && wage === 0 && paid === 0 && debtGiven === 0 && debtRepaid === 0){
          return; // nothing meaningful entered
        }
        entries.push({ date, name: lab.name, work, wage, paid, debtGiven, debtRepaid, notes });
        openCards.add(lab.name);
        render();
        await persist();
      });
      body.appendChild(form);

      const entHead = document.createElement('div');
      entHead.className = 'entries-head';
      entHead.innerHTML = '<span>Date</span><span>Work</span><span style="text-align:right">Wage</span><span style="text-align:right">Paid</span><span style="text-align:right">Due</span><span>Debt</span><span></span>';
      body.appendChild(entHead);

      const rows = entries
        .map((e, idx) => ({ e, idx }))
        .filter(x => x.e.name === lab.name)
        .sort((a,b) => b.e.date.localeCompare(a.e.date));

      if(rows.length === 0){
        const none = document.createElement('div');
        none.className = 'no-entries';
        none.textContent = 'No work entries logged yet for ' + lab.name + '.';
        body.appendChild(none);
      } else {
        rows.forEach(({e, idx}) => {
          const row = document.createElement('div');
          row.className = 'entry-row';
          const due = e.wage - e.paid;
          const debtBits = [];
          if(e.debtGiven) debtBits.push('+₹'+money(e.debtGiven)+' given');
          if(e.debtRepaid) debtBits.push('-₹'+money(e.debtRepaid)+' repaid');
          row.innerHTML = `
            <span>${fmtDate(e.date)}</span>
            <span>${e.work || '—'}${e.notes ? ' <span style="color:var(--graphite);font-size:11px;">('+e.notes+')</span>' : ''}</span>
            <span style="text-align:right">₹${money(e.wage)}</span>
            <span style="text-align:right">₹${money(e.paid)}</span>
            <span style="text-align:right" class="${due>0?'e-due positive':''}">₹${money(due)}</span>
            <span class="e-debt">${debtBits.join(', ') || '—'}</span>
            <button class="del-btn" title="Remove this entry">×</button>
          `;
          row.querySelector('.del-btn').addEventListener('click', async (ev) => {
            ev.stopPropagation();
            entries.splice(idx, 1);
            render();
            await persist();
          });
          body.appendChild(row);
        });
      }

      const cardActions = document.createElement('div');
      cardActions.className = 'card-actions';
      const delLab = document.createElement('button');
      delLab.className = 'del-btn';
      delLab.textContent = 'Remove ' + lab.name + ' and their entries';
      delLab.addEventListener('click', async (ev) => {
        ev.stopPropagation();
        if(!confirm('Remove ' + lab.name + ' and all their logged entries? This cannot be undone.')) return;
        labourers = labourers.filter(l => l.name !== lab.name);
        entries = entries.filter(e => e.name !== lab.name);
        openCards.delete(lab.name);
        render();
        await persist();
      });
      cardActions.appendChild(delLab);
      body.appendChild(cardActions);

      card.appendChild(body);
      labourerList.appendChild(card);
    });
  }

  addLabourerBtn.addEventListener('click', async () => {
    const name = newName.value.trim();
    const mobile = newMobile.value.trim();
    const notes = newNotes.value.trim();
    if(!name) return;
    if(labourers.some(l => l.name.toLowerCase() === name.toLowerCase())){
      setStatus('A labourer named "' + name + '" is already on the ledger — use a distinguishing name (e.g. add a surname).', true);
      return;
    }
    labourers.push({ name, mobile, notes });
    newName.value = ''; newMobile.value = ''; newNotes.value = '';
    openCards.add(name);
    render();
    await persist();
  });
  [newName,newMobile,newNotes].forEach(el => el.addEventListener('keydown', e => { if(e.key==='Enter') addLabourerBtn.click(); }));

  searchBox.addEventListener('input', () => {
    searchQuery = searchBox.value;
    render();
  });

  // ---------- Parse an existing workbook ----------
  function parseWorkbookBuffer(buf){
    const wb = XLSX.read(buf, { type:'array', cellDates:true });
    const labSheetName = wb.SheetNames.find(n => n.toLowerCase() === 'labourers') || wb.SheetNames[0];
    const entSheetName = wb.SheetNames.find(n => n.toLowerCase() === 'entries');

    const labAoa = XLSX.utils.sheet_to_json(wb.Sheets[labSheetName], { header:1, raw:true });
    const labHeaderIdx = labAoa.findIndex(r => r && r[0] === 'Name');
    if(labHeaderIdx === -1) throw new Error('No "Name" header row found on the Labourers sheet.');
    const newLabourers = [];
    for(let r = labHeaderIdx+1; r < labAoa.length; r++){
      const row = labAoa[r];
      if(!row || !row[0]) continue;
      newLabourers.push({ name: String(row[0]).trim(), mobile: row[1] ? String(row[1]).trim() : '', notes: row[2] ? String(row[2]).trim() : '' });
    }

    const newEntries = [];
    if(entSheetName){
      const entAoa = XLSX.utils.sheet_to_json(wb.Sheets[entSheetName], { header:1, raw:true });
      const entHeaderIdx = entAoa.findIndex(r => r && r[0] === 'Date');
      if(entHeaderIdx !== -1){
        for(let r = entHeaderIdx+1; r < entAoa.length; r++){
          const row = entAoa[r];
          if(!row || !row[0] || !row[1]) continue;
          let dateVal = row[0];
          let d = (dateVal instanceof Date) ? dateVal : new Date(dateVal);
          if(isNaN(d)) continue;
          // Excel/SheetJS store dates via UTC components — read them back the same way.
          const iso = d.getUTCFullYear() + '-' + String(d.getUTCMonth()+1).padStart(2,'0') + '-' + String(d.getUTCDate()).padStart(2,'0');
          newEntries.push({
            date: iso,
            name: String(row[1]).trim(),
            work: row[2] ? String(row[2]).trim() : '',
            wage: num(row[3]),
            paid: num(row[4]),
            debtGiven: num(row[6]),
            debtRepaid: num(row[7]),
            notes: row[8] ? String(row[8]).trim() : ''
          });
        }
      }
    }
    labourers = newLabourers;
    entries = newEntries;
    openCards = new Set();
  }

  // ---------- File System Access API ----------
  document.getElementById('connectBtn').addEventListener('click', async () => {
    try{
      const [handle] = await window.showOpenFilePicker({
        types: [{ description:'Excel Workbook', accept:{'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':['.xlsx']} }],
        excludeAcceptAllOption:false, multiple:false
      });
      const file = await handle.getFile();
      const buf = await file.arrayBuffer();
      parseWorkbookBuffer(buf); searchQuery = ''; searchBox.value = '';
      fileHandle = handle; connectedName = file.name;
      setStatus('Connected to ' + connectedName + '. Changes save straight back to it.');
      refreshChrome(); render();
    }catch(err){
      if(err.name !== 'AbortError') setStatus('Could not connect: ' + err.message, true);
    }
  });

  document.getElementById('newBtn').addEventListener('click', async () => {
    try{
      const handle = await window.showSaveFilePicker({
        suggestedName: 'wage_ledger.xlsx',
        types: [{ description:'Excel Workbook', accept:{'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':['.xlsx']} }]
      });
      fileHandle = handle; connectedName = handle.name || 'wage_ledger.xlsx';
      labourers = []; entries = []; openCards = new Set(); searchQuery = ''; searchBox.value = '';
      refreshChrome(); render();
      await persist();
      setStatus('Created ' + connectedName + ' and connected to it.');
    }catch(err){
      if(err.name !== 'AbortError') setStatus('Could not create the file: ' + err.message, true);
    }
  });

  document.getElementById('switchBtn').addEventListener('click', () => {
    fileHandle = null; connectedName = null;
    labourers = []; entries = []; openCards = new Set(); searchQuery = ''; searchBox.value = '';
    setStatus('');
    refreshChrome(); render();
  });

  // ---------- Fallback (Firefox/Safari) ----------
  if(!SUPPORTS_FS){
    document.getElementById('fileInput').addEventListener('change', async (e) => {
      const file = e.target.files[0];
      if(!file) return;
      try{
        const buf = await file.arrayBuffer();
        parseWorkbookBuffer(buf); searchQuery = ''; searchBox.value = '';
        connectedName = file.name;
        setStatus('Loaded ' + file.name + '. Click "Download wage_ledger.xlsx" after making changes to save them.');
        refreshChrome(); render();
      }catch(err){
        setStatus('Could not read that file: ' + err.message, true);
      }
      e.target.value = '';
    });
    document.getElementById('downloadBtn').addEventListener('click', () => {
      if(!isConnected()){ connectedName = 'wage_ledger.xlsx'; refreshChrome(); render(); }
      const wb = buildWorkbook();
      XLSX.writeFile(wb, 'wage_ledger.xlsx');
      setStatus('Downloaded wage_ledger.xlsx — replace your local copy with this one.');
    });
  }

  // ---------- Build workbook: Labourers (summary) + Entries (log) ----------
  function buildWorkbook(){
    const wb = XLSX.utils.book_new();
    const setCell = (ws,row,col,val,formula) => {
      const addr = XLSX.utils.encode_cell({r:row-1,c:col-1});
      if(formula) ws[addr] = { t:'n', f: formula };
      else if(val === null || val === undefined) ws[addr] = { t:'s', v:'' };
      else ws[addr] = (typeof val === 'string') ? { t:'s', v: val } : { t:'n', v: val };
    };

    // ---- Entries sheet ----
    const entAOA = [];
    entAOA.push(['Wage Ledger — Entries']);
    entAOA.push(["One row per day's work. Due Amount is calculated automatically. Generated from the Wage Ledger web page."]);
    entAOA.push([]);
    entAOA.push(['Date','Labourer Name','Work Done','Wage Amount','Paid Amount','Due Amount','Debt Given','Debt Repaid','Notes']);
    const entHeaderRow = 4;
    const entFirstDataRow = entHeaderRow + 1;
    const sortedEntries = entries.slice().sort((a,b) => a.date.localeCompare(b.date));
    sortedEntries.forEach(e => {
      const [y,m,d] = e.date.split('-').map(Number);
      const excelDate = new Date(Date.UTC(y, m-1, d));
      entAOA.push([excelDate, e.name, e.work, e.wage, e.paid, '', e.debtGiven, e.debtRepaid, e.notes]);
    });
    const wsEnt = XLSX.utils.aoa_to_sheet(entAOA);
    for(let i=0;i<sortedEntries.length;i++){
      const r = entFirstDataRow + i;
      const dAddr = XLSX.utils.encode_cell({r:r-1,c:0});
      if(wsEnt[dAddr]) wsEnt[dAddr].z = 'ddd dd-mmm-yyyy';
      setCell(wsEnt, r, 6, null, 'D'+r+'-E'+r); // Due = Wage - Paid
    }
    const entLastDataRow = entFirstDataRow + Math.max(sortedEntries.length,1) - 1;
    wsEnt['!cols'] = [{wch:14},{wch:18},{wch:22},{wch:12},{wch:12},{wch:12},{wch:12},{wch:13},{wch:24}];
    XLSX.utils.book_append_sheet(wb, wsEnt, 'Entries');

    // ---- Labourers sheet ----
    const labAOA = [];
    labAOA.push(['Wage Ledger — Labourers']);
    labAOA.push(['Totals are calculated automatically from the Entries sheet.']);
    labAOA.push([]);
    labAOA.push(['Name','Mobile Number','Notes','Total Wage','Total Paid','Total Due','Debt Given','Debt Repaid','Net Debt']);
    const wsLab = XLSX.utils.aoa_to_sheet(labAOA);
    const labHeaderRow = 4;
    labourers.forEach((lab, i) => {
      const r = labHeaderRow + 1 + i;
      setCell(wsLab, r, 1, lab.name);
      setCell(wsLab, r, 2, lab.mobile || '');
      setCell(wsLab, r, 3, lab.notes || '');
      setCell(wsLab, r, 4, null, "SUMIF(Entries!$B:$B,A"+r+",Entries!$D:$D)");
      setCell(wsLab, r, 5, null, "SUMIF(Entries!$B:$B,A"+r+",Entries!$E:$E)");
      setCell(wsLab, r, 6, null, "D"+r+"-E"+r);
      setCell(wsLab, r, 7, null, "SUMIF(Entries!$B:$B,A"+r+",Entries!$G:$G)");
      setCell(wsLab, r, 8, null, "SUMIF(Entries!$B:$B,A"+r+",Entries!$H:$H)");
      setCell(wsLab, r, 9, null, "G"+r+"-H"+r);
    });
    // aoa_to_sheet only knew about the header rows above — the labourer data
    // rows were added afterward via direct cell writes, which does NOT extend
    // the sheet's used-range ("!ref"). Without this, Excel/SheetJS only
    // serializes the original header-only range and every labourer row is
    // silently dropped from the saved file. Extend it explicitly.
    const labLastRowIndex0 = 3 + labourers.length; // 0-based index of the last used row
    wsLab['!ref'] = XLSX.utils.encode_range({ s:{r:0,c:0}, e:{r: labLastRowIndex0, c: 8} });
    wsLab['!cols'] = [{wch:18},{wch:15},{wch:24},{wch:12},{wch:12},{wch:12},{wch:12},{wch:13},{wch:11}];
    XLSX.utils.book_append_sheet(wb, wsLab, 'Labourers');
    // Put Labourers first so it opens as the summary view
    wb.SheetNames = ['Labourers','Entries'];

    return wb;
  }

  render();
})();
