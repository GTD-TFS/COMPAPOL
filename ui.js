/* =============== UI (render + interacciones) =============== */
/* Requiere: core.js (utilidades, estado, helpers) */
/* Opcional: io.js (import/export) para los botones de importar/guardar */

/* ---------- Utiles locales de UI ---------- */
function computeTituloFicha(f){
  const titulo=(f.nombre||f.apellidos)
    ? `${(f.nombre||'').trim()} ${(f.apellidos||'').trim().split(/\s+/)[0]||''}`.trim()
    : `Filiación`;
  return titulo || 'Filiación';
}

/* ---------- Botones de formato activos ---------- */
function updateFmtButtons(){
  try{
    const b = document.queryCommandState && document.queryCommandState('bold');
    const i = document.queryCommandState && document.queryCommandState('italic');
    const u = document.queryCommandState && document.queryCommandState('underline');
    $('#boldBtn')?.classList.toggle('active', !!b);
    $('#italicBtn')?.classList.toggle('active', !!i);
    $('#underBtn')?.classList.toggle('active', !!u);
  }catch{}
}

/* ---------- Coletillas (modal) ---------- */
function renderColetillas(){
  const cont=$("#coletillasList"); if(!cont) return;
  cont.innerHTML="";
  COLETILLAS.forEach(c=>{
    const el=document.createElement('div'); el.className="coletilla";
    el.innerHTML = `<div class="label">${c.label}</div>
                    <div><button class="btn secondary tiny">➜ Insertar</button></div>`;
    el.querySelector('button').onclick=()=>{
      insertHTMLAtCursor(escapeHtml(c.text));
      closeColetillas();
      editorFocus();
    };
    cont.appendChild(el);
  });
}
function openColetillas(){
  renderColetillas(); // Renderiza solo al abrir
  const m = $('#coletillasModal'); if(!m) return;
  m.classList.add('show');
  m.setAttribute('aria-hidden','false');
}
function closeColetillas(){
  const m = $('#coletillasModal'); if(!m) return;
  m.classList.remove('show');
  m.setAttribute('aria-hidden','true');
  const list = $('#coletillasList'); if(list) list.innerHTML="";
}

/* ---------- Render de filiaciones ---------- */
function renderFiliaciones(){
  const cont=$("#filiaciones"); if(!cont) return;
  cont.innerHTML="";
  state.filiaciones.forEach((f,i)=>{
    const det=document.createElement('details'); det.className="f-item";
    if(typeof openedIndex==='number' && openedIndex===i) det.open=true;

    const titulo = computeTituloFicha(f);
    const parts=[]; const tdoc=getTipoDocShown(f); if(tdoc) parts.push(tdoc); if(f.dni) parts.push(f.dni);
    const docmeta = parts.join(' · ');

    det.innerHTML=`
      <summary>
        <span class="tag">#${String(f.fixedId).padStart(2,'0')}</span>
        <span class="title">${escapeHtml(titulo)}</span>
        <span class="summary-right">
          ${docmeta?`<span class="docmeta">${escapeHtml(docmeta)}</span>`:""}
          <button class="btn success tiny" data-include="${f.fixedId}" title="Al texto">Al texto →</button>
          <span class="caret"></span>
        </span>
      </summary>
      <div class="details-body">
        <div class="row">
          <div class="col">
            <label>Condición <span class="req">*</span></label>
            <select data-k="condSel" data-i="${i}">
              <option value="" ${f.condSel===""?'selected':''}></option>
              <option value="Perjudicado" ${f.condSel==="Perjudicado"?'selected':''}>Perjudicado</option>
              <option value="Testigo" ${f.condSel==="Testigo"?'selected':''}>Testigo</option>
              <option value="Detenido" ${f.condSel==="Detenido"?'selected':''}>Detenido</option>
              <option value="Requirente" ${f.condSel==="Requirente"?'selected':''}>Requirente</option>
              <option value="Otro" ${f.condSel==="Otro"?'selected':''}>Otro</option>
            </select>
          </div>
          <div class="col" ${f.condSel==="Otro"?'':"style='display:none'"}><label>Condición (otro)</label><input data-k="condOtro" data-i="${i}" value="${f.condOtro||''}" /></div>

          <div class="col"><label>Nombre</label><input data-k="nombre" data-i="${i}" value="${f.nombre||''}" /></div>
          <div class="col"><label>Apellidos</label><input data-k="apellidos" data-i="${i}" value="${f.apellidos||''}" /></div>

          <div class="col">
            <label>Tipo de documento</label>
            <select data-k="tipoSel" data-i="${i}">
              <option value="" ${f.tipoSel===""?'selected':''}></option>
              <option value="DNI" ${f.tipoSel==="DNI"?'selected':''}>DNI</option>
              <option value="NIE" ${f.tipoSel==="NIE"?'selected':''}>NIE</option>
              <option value="Pasaporte" ${f.tipoSel==="Pasaporte"?'selected':''}>Pasaporte</option>
              <option value="Indocumentado/a" ${f.tipoSel==="Indocumentado/a"?'selected':''}>Indocumentado/a</option>
              <option value="Otro" ${f.tipoSel==="Otro"?'selected':''}>Otro documento</option>
            </select>
          </div>
          <div class="col" ${f.tipoSel==="Otro"?'':"style='display:none'"}><label>Otro documento</label><input data-k="otroDoc" data-i="${i}" value="${f.otroDoc||''}" /></div>
          <div class="col"><label>Nº Documento</label><input data-k="dni" data-i="${i}" value="${f.dni||''}" /></div>

          <div class="col"><label>Fecha de nacimiento</label><input data-k="fechaNac" data-i="${i}" value="${f.fechaNac||''}" inputmode="numeric" /></div>
          <div class="col"><label>Lugar de nacimiento</label><input data-k="lugarNac" data-i="${i}" value="${f.lugarNac||''}" /></div>

          <div class="col"><label>Nombre de los Padres</label><input data-k="padres" data-i="${i}" value="${f.padres||''}" /></div>
          <div class="col"><label>Domicilio</label><input data-k="domicilio" data-i="${i}" value="${f.domicilio||''}" /></div>
          <div class="col"><label>Teléfono</label><input data-k="telefono" data-i="${i}" value="${f.telefono||''}" inputmode="tel" /></div>
        </div>
        <div class="btn-row" style="margin-top:8px; justify-content:flex-end">
          <button class="btn secondary tiny" data-xlsx="${i}" title="Descargar XLSX" ${!isCondValid(f)?'disabled':''}>⬇️ XLSX</button>
          <button class="btn secondary tiny" data-json="${i}" title="Descargar JSON">⬇️ JSON</button>
          <button class="btn danger tiny" data-del="${i}" title="Eliminar">🗑️</button>
          <button class="btn success tiny" data-include="${f.fixedId}" title="Al texto">Al texto →</button>
        </div>
      </div>`;
    cont.appendChild(det);
  });

  $('#emptyHint') && ($('#emptyHint').style.display = state.filiaciones.length ? 'none':'block');

  // Inputs/selects (normalización + UI reactiva)
  cont.querySelectorAll('input[data-k], select[data-k]').forEach(inp=>{
    const onInput = e=>{
      const i=+e.target.dataset.i, k=e.target.dataset.k;
      const f=state.filiaciones[i];
      if(k==="fechaNac"){
        let v=e.target.value.replace(/\D/g,'').slice(0,8);
        let out=""; if(v.length>=2){ out+=v.slice(0,2)+"/"; } else { out+=v; }
        if(v.length>=4){ out+=v.slice(2,4)+"/"; } else if(v.length>2){ out+=v.slice(2); }
        if(v.length>4){ out+=v.slice(4); }
        e.target.value=out; f[k]=out; saveDebounced(); return;
      }
      f[k]=e.target.value;

      if(k==="tipoSel"){
        const det=e.target.closest('details');
        const col = det.querySelector('input[data-k="otroDoc"]')?.closest('.col');
        if(col){ col.style.display = e.target.value==="Otro" ? "" : "none"; }
      }
      if(k==="condSel"){
        const det=e.target.closest('details');
        const col = det.querySelector('input[data-k="condOtro"]')?.closest('.col');
        if(col){ col.style.display = e.target.value==="Otro" ? "" : "none"; }
      }

      saveDebounced();

      if(["nombre","apellidos","tipoSel","otroDoc","dni"].includes(k)){
        if(f.tipoSel==="Indocumentado" || f.tipoSel==="Indocumentada") f.tipoSel="Indocumentado/a";
        f.tipoDoc = mapIndocumentadoAny(normTipoDocLabel(f.tipoSel, f.otroDoc));
        const det=e.target.closest('details'); if(det){
          const tEl=det.querySelector('.title');
          if(tEl) tEl.textContent = computeTituloFicha(f);
          const meta=det.querySelector('.docmeta');
          const parts=[]; const tdoc=getTipoDocShown(f); if(tdoc) parts.push(tdoc); if(f.dni) parts.push(f.dni);
          if(meta){ meta.textContent = parts.join(' · '); }
        }
      }
    };
    const onBlur = e=>{
      const i=+e.target.dataset.i, k=e.target.dataset.k;
      const f=state.filiaciones[i];
      switch(k){
        case "nombre":   f.nombre   = normNombre(f.nombre); break;
        case "apellidos":f.apellidos= normApellidos(f.apellidos); break;
        case "tipoSel":  f.tipoDoc  = mapIndocumentadoAny(normTipoDocLabel(f.tipoSel, f.otroDoc)); break;
        case "otroDoc":  f.otroDoc  = titleCaseEs(f.otroDoc); f.tipoDoc = mapIndocumentadoAny(normTipoDocLabel(f.tipoSel, f.otroDoc)); break;
        case "dni":      f.dni      = normNumDoc(f.dni); break;
        case "padres":   f.padres   = normPadres(f.padres); break;
        case "domicilio":f.domicilio= normDomicilio(f.domicilio); break;
        case "lugarNac": f.lugarNac = normLugarNac(f.lugarNac); break;
        case "condOtro": f.condOtro = titleCaseEs(f.condOtro); break;
      }
      e.target.value = f[k] || "";
      save();
      if(["nombre","apellidos","tipoSel","otroDoc","dni"].includes(k)){
        const det=e.target.closest('details'); if(det){
          const tEl=det.querySelector('.title'); if(tEl) tEl.textContent=computeTituloFicha(f);
          const meta=det.querySelector('.docmeta');
          const parts=[]; const tdoc=getTipoDocShown(f); if(tdoc) parts.push(tdoc); if(f.dni) parts.push(f.dni);
          if(meta){ meta.textContent = parts.join(' · '); }
        }
      }
    };
    inp.addEventListener('input', onInput, {passive:true});
    inp.addEventListener('blur', onBlur, {passive:true});
  });

  // Acciones
  cont.querySelectorAll('button[data-del]').forEach(b=>b.onclick=()=>{
    const idx=+b.dataset.del; state.filiaciones.splice(idx,1); save(); openedIndex=-1; renderFiliaciones();
  });
  cont.querySelectorAll('button[data-xlsx]').forEach(b=>b.onclick=()=>{
    const idx=+b.dataset.xlsx; const f=state.filiaciones[idx];
    if(!isCondValid(f)){ alert(`Falta rellenar la condición en la filiación #${f.fixedId}.`); return; }
    const xlsx=makeXLSXFromFiliacion(f);
    download(fileNameForFiliacion(f), xlsx, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  });
  cont.querySelectorAll('button[data-json]').forEach(b=>b.onclick=()=>{
    const idx=+b.dataset.json; const f=state.filiaciones[idx];
    const pretty = JSON.stringify(f, null, 2);
    download(`filiacion_${f.fixedId}.json`, toUTF8(pretty), "application/json");
  });
  cont.querySelectorAll('button[data-include]').forEach(b=>b.onclick=()=>{
    includeFiliacionById(+b.dataset.include);
  });
}

/* ---------- Editor: atajos y formato ---------- */
let capNext = false;

function getTextNodeAndOffsetForRange(r){
  let node = r.startContainer;
  let offset = r.startOffset;
  if(node.nodeType === Node.TEXT_NODE){
    return {node, offset};
  }
  if(node.nodeType === Node.ELEMENT_NODE){
    if(node.childNodes && node.childNodes.length && offset>0){
      node = node.childNodes[offset-1];
      while(node && node.lastChild) node = node.lastChild;
      if(node && node.nodeType===Node.TEXT_NODE){
        return {node, offset: node.data.length};
      }
    }
    let walk = r.startContainer;
    while(walk && walk !== editorEl()){
      if(walk.previousSibling){
        walk = walk.previousSibling;
        while(walk && walk.lastChild) walk = walk.lastChild;
        if(walk && walk.nodeType===Node.TEXT_NODE){
          return {node:walk, offset: walk.data.length};
        }
      }else{
        walk = walk.parentNode;
      }
    }
  }
  return {node:null, offset:0};
}

function handleShortcutsAtCaret(){
  const sel = window.getSelection();
  if(!sel || !sel.rangeCount) return false;
  const r = sel.getRangeAt(0);
  if(!editorEl().contains(r.startContainer)) return false;

  const {node, offset} = getTextNodeAndOffsetForRange(r);
  if(!node) return false;

  const textBefore = node.data.slice(0, offset);

  // ccc / CCC -> Coletillas
  const mC = textBefore.match(/ccc$/i);
  if(mC){
    const rr = document.createRange();
    rr.setStart(node, Math.max(0, offset-3));
    rr.setEnd(node, offset);
    rr.deleteContents();
    openColetillas();
    return true;
  }

  // fN / FN (sin paréntesis)
  const m = textBefore.match(/\b[fF](\d+)$/);
  if(m){
    const id = parseInt(m[1],10);
    const rr = document.createRange();
    rr.setStart(node, offset - m[0].length);
    rr.setEnd(node, offset);
    rr.deleteContents();
    includeFiliacionById(id);
    return true;
  }
  return false;
}

/* ---------- Eventos globales de UI ---------- */
function bindUI(){
  // Doc editor
  const ed = $('#doc');
  if(ed){
    ed.addEventListener('beforeinput', (e)=>{
      if(e.inputType==='insertText' && e.data === '.') capNext = true;
      if(e.inputType==='insertText' && capNext && typeof e.data === 'string' && e.data.length===1){
        const ch = e.data;
        if(/[a-záéíóúüñ]/i.test(ch)){
          e.preventDefault();
          document.execCommand('insertText', false, ch.toLocaleUpperCase('es-ES'));
          capNext = false;
        }
      }
    });
    ed.addEventListener('input', ()=>{
      handleShortcutsAtCaret();
      state.doc=getDocHTML(); saveDocDebounced();
      saveEditorSelection();
      updateFmtButtons();
    });
    ed.addEventListener('keydown', (e)=>{
      if(e.key === 'Enter'){
        e.preventDefault();
        insertHTMLAtCursor('<br>-- &nbsp;');
        capNext = true;
        state.doc=getDocHTML(); saveDocDebounced();
        return;
      }
      if(e.key === '.'){ capNext = true; }
    });
    ed.addEventListener('mouseup', ()=>{ saveEditorSelection(); updateFmtButtons(); });
    ed.addEventListener('keyup',   ()=>{ saveEditorSelection(); updateFmtButtons(); });
    ed.addEventListener('focus', ()=>{ $$('#filiaciones details[open]').forEach(d=>d.open=false); });
  }

  // Fmt buttons
  $('#boldBtn') && ($('#boldBtn').onclick = ()=>{ editorFocus(); document.execCommand('bold', false, null); state.doc=getDocHTML(); saveDocDebounced(); updateFmtButtons(); });
  $('#italicBtn') && ($('#italicBtn').onclick = ()=>{ editorFocus(); document.execCommand('italic', false, null); state.doc=getDocHTML(); saveDocDebounced(); updateFmtButtons(); });
  $('#underBtn') && ($('#underBtn').onclick = ()=>{ editorFocus(); document.execCommand('underline', false, null); state.doc=getDocHTML(); saveDocDebounced(); updateFmtButtons(); });

  // Coletillas
  $('#openColetillasBtn') && ($('#openColetillasBtn').onclick=openColetillas);
  $('#closeColetillasBtn') && ($('#closeColetillasBtn').onclick=closeColetillas);
  $('#coletillasModal') && $('#coletillasModal').addEventListener('click', e=>{ if(e.target.id==='coletillasModal'){ closeColetillas(); } });

  // Título
  $('#titulo') && $('#titulo').addEventListener('input', ()=>{ state.titulo=$('#titulo').value; saveDocDebounced(); });
  $('#titulo') && $('#titulo').addEventListener('blur',  ()=>{ state.titulo = titleCaseEs($('#titulo').value||""); $('#titulo').value = state.titulo; saveDocDebounced(); });

  // Refrescar total
  $('#refreshTopBtn') && ($('#refreshTopBtn').onclick = ()=> wipeAllAndRender());
}

/* ---------- Exponer funciones de UI a global (si hace falta) ---------- */
window.renderFiliaciones = renderFiliaciones;
window.openColetillas  = openColetillas;
window.closeColetillas = closeColetillas;
window.updateFmtButtons = updateFmtButtons;

/* ---------- Init UI ---------- */
(function uiInit(){
  renderFiliaciones();
  updateFmtButtons();
})();
