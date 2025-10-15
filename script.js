// Main script for NAB Path Detector
class PathDetector {
  constructor(){
    // UI refs
    this.fileInput=document.getElementById('fileInput');
    this.chooseBtn=document.getElementById('chooseFilesBtn');
    this.dropzone=document.getElementById('dropzone');
    this.selectedList=document.getElementById('selectedList');
    this.scanBtn=document.getElementById('scanBtn');
    this.resetBtn=document.getElementById('resetBtn');
    this.exportBtn=document.getElementById('exportBtn');
    this.progressPanel=document.getElementById('progressPanel');
    this.progressBar=document.getElementById('progressBar');
    this.progressPct=document.getElementById('progressPct');
    this.currentFileLabel=document.getElementById('currentFile');
    this.elapsedLabel=document.getElementById('elapsed');
    this.detectionCountLabel=document.getElementById('detectionCount');
    this.resultsSection=document.getElementById('results');
    this.nasTable=document.querySelector('#nasTable tbody');
    this.otherTable=document.querySelector('#otherTable tbody');
    this.nasCountLabel=document.getElementById('nasCount');
    this.otherCountLabel=document.getElementById('otherCount');
    this.protectedList=document.getElementById('protectedList');
    this.protectedCountLabel=document.getElementById('protectedCount');
    this.acknowledgeCheckbox=document.getElementById('acknowledgeGuidance');

    // state
    this.files=[];
    this.results=[];
    this.protectedFiles=[];
    this.totalDetections=0;
    this.startTime=null;
    this.timer=null;

    // regexes
  this.NAS_REGEX = /\\\\\s*(aur\.national\.com\.au|dfs|filesrv|corp)[\\\/][^\s"'<>]+/gi;
  this.OTHER_REGEX = /(?:[A-Z]:\\[^\s"'<>]+|\\\\\s*[A-Za-z0-9._-]+\\[^\s"'<>]+)/g;

    // event wiring
    this.initEvents();
  }

  initEvents(){
    this.chooseBtn.addEventListener('click',()=>this.fileInput.click());
    this.fileInput.addEventListener('change',(e)=>this.handleFiles(e.target.files));

    // drag and drop
    ['dragenter','dragover'].forEach(ev=>{
      this.dropzone.addEventListener(ev,(e)=>{e.preventDefault();this.dropzone.classList.add('drag');});
    });
    ['dragleave','drop'].forEach(ev=>{
      this.dropzone.addEventListener(ev,(e)=>{e.preventDefault();this.dropzone.classList.remove('drag');});
    });
    this.dropzone.addEventListener('drop',(e)=>{ const dt=e.dataTransfer; if(dt && dt.files) this.handleFiles(dt.files);});

    // checkbox
    this.acknowledgeCheckbox.addEventListener('change',()=>this.updateScanButton());
    this.updateScanButton();

    // controls
    this.scanBtn.addEventListener('click',()=>this.startScan());
    this.resetBtn.addEventListener('click',()=>this.reset());
    document.getElementById('exportCsvTop').addEventListener('click',()=>this.exportCSV());
    document.getElementById('exportCsvBottom').addEventListener('click',()=>this.exportCSV());
    document.getElementById('scanAgainTop').addEventListener('click',()=>this.reset());
    document.getElementById('scanAgainBottom').addEventListener('click',()=>this.reset());
    // status elements
    this.statusMessage = document.getElementById('statusMessage');
    this.errorDetails = document.getElementById('errorDetails');
    this.errorDetailsBody = document.getElementById('errorDetailsBody');
    this.scanningOverlay = document.getElementById('scanningOverlay');
    this.loadingText = document.getElementById('loadingText');
  }

  updateScanButton(){
    const enabled = this.acknowledgeCheckbox.checked && this.files.length>0;
    this.scanBtn.disabled = !enabled;
    this.resetBtn.disabled = false; // always allow reset
  }

  handleFiles(fileList){
    const arr = Array.from(fileList).slice(0,10);
    const allowedExt = ['.xlsm','.xlsb','.docm','.pptm'];
    this.files = [];
    const display=[];
    arr.forEach(f=>{
      const lower = f.name.toLowerCase();
      const extOk = allowedExt.some(e=>lower.endsWith(e));
      if(!extOk) return;
      if(f.size>100*1024*1024){
        display.push(`<div class="file-warning">${f.name} — skipped (over 100MB)</div>`);
        return;
      }
      this.files.push(f);
      display.push(`<div>${f.name} — ${(f.size/1024/1024).toFixed(1)} MB</div>`);
    });
    if(this.files.length===0 && display.length===0) display.push('<div>No supported files selected.</div>');
    this.selectedList.innerHTML = display.join('');
    this.updateScanButton();
  }

  async startScan(){
    if(!this.acknowledgeCheckbox.checked) {
      alert('Please acknowledge that you understand by checking the box below before starting the scan.');
      return;
    }
    if(this.files.length===0) return alert('No files selected.');

    // reset results
    this.results=[]; this.protectedFiles=[]; this.totalDetections=0;
    this.nasTable.innerHTML=''; this.otherTable.innerHTML=''; this.protectedList.innerHTML='';
    this.nasCountLabel.textContent='(0)'; this.otherCountLabel.textContent='(0)'; this.protectedCountLabel.textContent='(0)';

    this.progressPanel.hidden=false; this.resultsSection.hidden=true;
    this.startTime=Date.now(); this.updateElapsed(); this.timer=setInterval(()=>this.updateElapsed(),500);

    try{
      for(let i=0;i<this.files.length;i++){
        const file=this.files[i];
        this.currentFileLabel.textContent = `${file.name} (${i+1}/${this.files.length})`;
        const pct = Math.round((i/this.files.length)*100);
        this.setProgress(pct);
        // show scanning overlay/status
        this.showStatus(`Scanning ${file.name} (${i+1}/${this.files.length})...`, 'info');
        if(this.loadingText) this.loadingText.textContent = `Scanning file ${i+1} of ${this.files.length}: ${file.name}`;
        if(this.scanningOverlay) this.scanningOverlay.classList.remove('hidden');
        try{
          await this.scanFile(file);
        }catch(err){
          console.error('scan error',file.name,err);
          this.protectedFiles.push({file:file.name, status:'Protected', reason: err && err.message ? err.message : 'Unreadable or protected'});
        }
        // update detections and progress display
        this.detectionCountLabel.textContent = this.totalDetections;
      }
    }catch(fatal){
      console.error('Fatal scanning error',fatal);
      this.showStatus('❌ Scan failed due to unexpected error. Please refresh and try again.', 'error');
      if(this.scanningOverlay) this.scanningOverlay.classList.add('hidden');
      clearInterval(this.timer);
      this.progressPanel.hidden=true; this.resultsSection.hidden=false;
      this.protectedFiles.push({file:'<scan-failed>', status:'Failed', reason: fatal && fatal.message ? fatal.message : 'Unexpected error'});
      this.populateResults();
      return;
    }

    this.setProgress(100);
    clearInterval(this.timer); this.updateElapsed();
    if(this.scanningOverlay) this.scanningOverlay.classList.add('hidden');
    this.progressPanel.hidden=true; this.resultsSection.hidden=false;
    this.populateResults();

    // show final status banner
    if(this.protectedFiles.length>0){
      this.showStatus(`⚠️ Scan partially completed – ${this.files.length - this.protectedFiles.length} of ${this.files.length} files scanned successfully. ${this.totalDetections} detections found.`, 'warning');
    }else{
      this.showStatus(`✅ Scan Completed – ${this.files.length} files scanned, ${this.totalDetections} detections found.`, 'success');
    }
  }

  setProgress(pct){
    this.progressBar.style.width = pct+"%";
    this.progressPct.textContent = pct+"%";
  }

  updateElapsed(){
    if(!this.startTime) return; const s = Math.round((Date.now()-this.startTime)/1000); this.elapsedLabel.textContent = `${s}s`;
  }

  async scanFile(file){
    const name=file.name;
    let isZip=true;
    let zip=null;
    try{
      zip = await JSZip.loadAsync(file);
    }catch(e){
      isZip=false;
    }

    if(!isZip){
      // try raw buffer scan (best-effort for binary formats)
      try{
        const buf = await file.arrayBuffer();
        const txt = new TextDecoder('latin1').decode(buf);
        const norm = this.normalizeText(txt);
        console.debug(`Normalized length for ${name}: ${norm.length}`);
        this.searchTextForPaths(norm,name,'Unknown');
      }catch(e){
        this.protectedFiles.push({file:name, reason:'Protected or Encrypted File – cannot scan'});
      }
      return;
    }

    // Determine application type
    let app = 'Unknown';
    if (zip.file('xl/workbook.xml') || zip.folder('xl')) {
    app = 'Excel';
    } else if (zip.file('word/document.xml') || zip.folder('word')) {
    app = 'Word';
    } else if (zip.folder('ppt')) {
    app = 'PowerPoint';
    }


    // scan vbaProject.bin (if present)
    const vbaPath = (app==='Excel'?'xl/vbaProject.bin':app==='Word'?'word/vbaProject.bin':app==='PowerPoint'?'ppt/vbaProject.bin':'vbaProject.bin');
    if(zip.file(vbaPath)){
      try{
        const arr = await zip.file(vbaPath).async('arraybuffer');
        const vbatxt = await window.VBAExtractorHelper.extractTextFromVBA(arr);
        console.debug('VBA text first 300 chars for', name, vbatxt.slice(0,300));
        console.debug('NAS regex test:', this.NAS_REGEX.test(vbatxt));
        const norm = this.normalizeText(vbatxt);
        console.debug(`Normalized length for ${name} (vba): ${norm.length}`);
        console.debug('Normalized preview:', norm.slice(0,500));
        this.searchTextForPaths(norm,name, 'VBA / vbaProject.bin');
      }catch(e){
        console.warn('vba read fail',e);
      }
    }

    // Excel connections
    if(app==='Excel'){
      if(zip.file('xl/connections.xml')){
        const t = await zip.file('xl/connections.xml').async('string');
        const norm = this.normalizeText(t);
        console.debug(`Normalized length for ${name} (connections): ${norm.length}`);
        this.searchTextForPaths(norm,name,'xl/connections.xml');
      }
      // external links
      if (zip.folder('xl/externalLinks')) {
        for (const [relativePath, fileEntry] of Object.entries(zip.folder('xl/externalLinks').files)) {
          if (!fileEntry.dir) {
            const t = await fileEntry.async('string');
            const norm = this.normalizeText(t);
            console.debug(`Normalized length for ${name} (xl/externalLinks/${relativePath}): ${norm.length}`);
            console.debug('Normalized preview:', norm.slice(0,500));
            this.searchTextForPaths(norm,name,'xl/externalLinks/'+relativePath);
          }
        }
      }
      // rels
      if(zip.folder('xl/_rels')){
        for (const [relativePath, fileEntry] of Object.entries(zip.folder('xl/_rels').files)) {
          if(!fileEntry.dir){
            const t=await fileEntry.async('string');
            const norm = this.normalizeText(t);
            console.debug(`Normalized length for ${name} (xl/_rels/${relativePath}): ${norm.length}`);
            console.debug('Normalized preview:', norm.slice(0,500));
            this.searchTextForPaths(norm,name,'xl/_rels/'+relativePath);
          }
        }
      }
      // customXml item*.data
      if(zip.folder('customXml')){
        zip.folder('customXml').forEach(async (relativePath,fileEntry)=>{
          if(!fileEntry.dir && /item.*\.data$/i.test(relativePath)){
              const t=await fileEntry.async('string');
              const norm = this.normalizeText(t);
              console.debug(`Normalized length for ${name} (customXml/${relativePath}): ${norm.length}`);
              this.searchTextForPaths(norm,name,'customXml/'+relativePath,'Power Query present');
            }
        });
      }
    }

    // Word
    if(app==='Word'){
      if(zip.file('word/document.xml')){
        const t=await zip.file('word/document.xml').async('string');
        const norm = this.normalizeText(t);
        console.debug(`Normalized length for ${name} (word/document.xml): ${norm.length}`);
        this.searchTextForPaths(norm,name,'word/document.xml');
      }
      if(zip.file('word/_rels/document.xml.rels')){
        const t=await zip.file('word/_rels/document.xml.rels').async('string');
        const norm = this.normalizeText(t);
        console.debug(`Normalized length for ${name} (word/_rels/document.xml.rels): ${norm.length}`);
        this.searchTextForPaths(norm,name,'word/_rels/document.xml.rels');
      }
    }

    // PowerPoint
    if(app==='PowerPoint'){
      if(zip.folder('ppt/slides/_rels')){
        for (const [relativePath, fileEntry] of Object.entries(zip.folder('ppt/slides/_rels').files)) {
          if(!fileEntry.dir){
            const t=await fileEntry.async('string');
            const norm = this.normalizeText(t);
            console.debug(`Normalized length for ${name} (ppt/slides/_rels/${relativePath}): ${norm.length}`);
            console.debug('Normalized preview:', norm.slice(0,500));
            this.searchTextForPaths(norm,name,'ppt/slides/_rels/'+relativePath);
          }
        }
      }
    }
    
    // Log total detections and filtered matches for this file
    console.info(`File ${name}: found ${this.results.length} total detections so far (${this.filteredCount} matches filtered as system/internal noise)`);
  }

  // Normalize text: remove hidden unicode and directional marks, normalize slashes, handle invisible spaces
  normalizeText(s){
    if(!s) return '';
    // remove hidden unicode, BOM and directional marks
    s = s.replace(/[\u200B-\u200F\uFEFF\u202A-\u202E]/g,'');
    // normalize forward slashes to backslashes
    s = s.replace(/\//g,'\\');
    // normalize multiple backslashes
    s = s.replace(/\\+\s*\\+/g,'\\\\');
    // remove any non-printable ASCII except common punctuation
    s = s.replace(/[^\x09\x0A\x0D\x20-\x7E]/g,'');
    return s;
  }

  isHumanPath(p){
    if(!p) return false;
    const clean = p.trim();

    // Exclude internal schemas and Office structure
    if(/schemas\.openxmlformats\.org/i.test(clean)) return false;
    if(/content-types/i.test(clean)) return false;
    if(/relationships/i.test(clean)) return false;
    if(/docProps|_rels|xl\/_rels|ppt\/_rels|word\/_rels/i.test(clean)) return false;

    // Exclude binary garbage (non-printable characters)
    if(/[\x00-\x08\x0E-\x1F]/.test(clean)) return false;

    // Exclude trivial system temp references
    if(/^C:\\Temp\\?$/i.test(clean)) return false;
    if(/^C:\\Windows/i.test(clean)) return false;

    // Require reasonable length and at least one subfolder
    if(clean.length < 10) return false;
    if(!/[\\/].*[\\/]/.test(clean)) return false;

    return true;
  }

  searchTextForPaths(text, fileName, section='unknown', hint){
    if(!text) return;
    // Initialize filtered count if not exists
    if(typeof this.filteredCount === 'undefined') this.filteredCount = 0;
    
    // run both regexes
    const nasMatches = [...(text.matchAll(this.NAS_REGEX))];
    const otherMatches = [...(text.matchAll(this.OTHER_REGEX))];

    const recordMatch=(match,category)=>{
      const matched=match[0];
      // Filter out system/internal paths
      if(!this.isHumanPath(matched)) {
        this.filteredCount++;
        return;
      }
      const context = this.getContextAround(text, match.index);
      const usage = this.inferUsage(context);
      const impact = this.usageImpact(usage);
      const line = this.getLineNumber(text, match.index);
      const rec={file:fileName,application:this.guessAppFromSection(section),section:section,line:line,path:matched,category:category,possibleUsage:usage,impact:impact};
      this.results.push(rec);
      this.totalDetections++;
    };

    nasMatches.forEach(m=>recordMatch(m,'NAS Shared Drive'));
    otherMatches.forEach(m=>{
      // avoid duplicates that are NAS
      const isNas = m[0].match(this.NAS_REGEX);
      if(isNas) return;
      recordMatch(m,'Other Hard-Coded Path');
    });
  }

  getContextAround(text, idx, radius=80){
    try{
      const start = Math.max(0, idx - radius);
      const end = Math.min(text.length, idx + radius);
      return text.substring(start,end);
    }catch(e){return ''}
  }

  inferUsage(line){
    const l=line || '';
    if(/Workbooks\.Open|FileSystemObject|CreateObject\(\"Scripting\.FileSystemObject\"\)/i.test(l)) return 'Workbook Open Macro';
    if(/ConnectionString|Data Source|Power Query|ODBC|OLEDB|SqlClient/i.test(l)) return 'Power Query / Data Connection';
    if(/Hyperlink|TargetMode=\"External\"/i.test(l)) return 'Static Hyperlink';
    if(/Dir\(|Kill\(|Name\s+\w+/i.test(l)) return 'Folder Operation Macro';
    return 'Generic File Reference';
  }

  usageImpact(usage){
    if(usage==='Workbook Open Macro' || usage==='Folder Operation Macro') return 'High';
    if(usage==='Generic File Reference' || usage==='Power Query / Data Connection') return 'Medium';
    if(usage==='Static Hyperlink') return 'Low';
    return 'Medium';
  }

  getLineNumber(text, idx){
    if(!text) return 0;
    const pre = text.substring(0, idx);
    return pre.split(/\r?\n/).length;
  }

  guessAppFromSection(section){
    if(!section) return 'Unknown';
    if(section.indexOf('xl/')===0 || section==='VBA / vbaProject.bin' && section.toLowerCase().indexOf('xl')>-1) return 'Excel';
    if(section.indexOf('word/')===0) return 'Word';
    if(section.indexOf('ppt/')===0) return 'PowerPoint';
    return 'Generic';
  }

  populateResults(){
    // populate tables and protected list
    this.nasTable.innerHTML=''; this.otherTable.innerHTML=''; this.protectedList.innerHTML='';
    let nasCount=0, otherCount=0;
    this.results.forEach(r=>{
      const tr = document.createElement('tr');
      tr.innerHTML = `<td>${r.file}</td><td>${r.line}</td><td class="path-cell">${this.highlight(r.path)}</td><td>${r.possibleUsage}</td><td>${r.impact}</td>`;
      if(r.category==='NAS Shared Drive'){ this.nasTable.appendChild(tr); nasCount++; }
      else { this.otherTable.appendChild(tr); otherCount++; }
    });
    this.nasCountLabel.textContent = `(${nasCount})`;
    this.otherCountLabel.textContent = `(${otherCount})`;
    this.totalDetections = nasCount + otherCount;
    this.detectionCountLabel.textContent = this.totalDetections;

    this.protectedFiles.forEach(p=>{
      const div = document.createElement('div'); div.textContent = `${p.file} — ${p.reason || p.status || 'Protected or unreadable'}`; this.protectedList.appendChild(div);
    });
    this.protectedCountLabel.textContent = `(${this.protectedFiles.length})`;

    // enable export
    document.getElementById('exportCsvTop').disabled=false;
    document.getElementById('exportCsvBottom').disabled=false;
  }

  showStatus(message, type='info'){
    if(!this.statusMessage) return;
    this.statusMessage.textContent = message;
    this.statusMessage.className = `status-message ${type}`;
    this.statusMessage.classList.remove('hidden');
  }

  hideStatus(){
    if(!this.statusMessage) return;
    this.statusMessage.classList.add('hidden');
  }

  highlight(text){
    const escaped = this.escapeHtml(text);
    return escaped.replace(/(\\\\(aur\.national\.com\.au|dfs|filesrv|corp)[\\\/][^\s"'<>]+)/gi, '<mark>$1</mark>').replace(/(\\\\[A-Za-z0-9._-]+\\[^\s"'<>]+|[A-Z]:\\[^\s"'<>]+)/g, '<mark>$1</mark>');
  }

  escapeHtml(s){
    return (s+'').replace(/[&<>\"']/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":"&#39;"}[c];});
  }

  exportCSV(){
    if(this.results.length===0){ alert('No results to export'); return; }
    const header=['file','line','path','category','possibleUsage','impact'];
    const rows = [header.join(',')];
    this.results.forEach(r=>{
      const row=[r.file,r.line,`"${r.path.replace(/"/g,'""')}"`,r.category,r.possibleUsage,r.impact];
      rows.push(row.join(','));
    });
    const csv = rows.join('\r\n');
    const blob = new Blob([csv],{type:'text/csv;charset=utf-8;'});
    saveAs(blob,'nab-path-detection-report.csv');
  }

  reset(){
    // reset UI to initial
    this.files=[]; this.results=[]; this.protectedFiles=[]; this.totalDetections=0;
    this.selectedList.innerHTML=''; this.nasTable.innerHTML=''; this.otherTable.innerHTML=''; this.protectedList.innerHTML='';
    this.resultsSection.hidden=true; this.progressPanel.hidden=true; this.setProgress(0);
    this.updateScanButton();
  }
}

// Initialize on DOM ready
window.addEventListener('DOMContentLoaded',()=>{
  window.pathDetector = new PathDetector();
});
