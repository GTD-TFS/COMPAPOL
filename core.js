// ============ CORE: utilidades, estado, normalizadores, ZIP/ODT/XLSX ============

(function(){
  const $ = s=>document.querySelector(s);

  // ----- Estado -----
  const LS_KEY = "gestor_partes_comparecencias_mobile_v3";
  const state = { filiaciones:[], titulo:"", doc:"", nextId:1 };
  let openedIndex = -1;

  function save(){ try{ localStorage.setItem(LS_KEY, JSON.stringify(state)); }catch{} }
  function hardResetStorage(){ try{ localStorage.removeItem(LS_KEY); }catch{} }
  function load(){ try{ Object.assign(state, JSON.parse(localStorage.getItem(LS_KEY)||"{}")); }catch{} }

  // ----- Utils -----
  function debounce(fn, wait=800){ let t; return (...a)=>{ clearTimeout(t); t=setTimeout(()=>fn(...a), wait); }; }
  const saveDebounced = debounce(()=>save(), 800);
  const saveDocDebounced = debounce(()=>save(), 1200);

  function toUTF8(s){ return new TextEncoder().encode(s); }
  function fromUTF8(u8){ return new TextDecoder().decode(u8); }

  function todayISO(){ const d=new Date(), pad=n=>String(n).padStart(2,'0'); return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`; }

  function escapeXml(s){ return String(s??"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&apos;"); }
  function escapeHtml(s){ return (s??"").replace(/[&<>"']/g, m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m])); }
  function stripHTMLExceptBIU(html){
    const tmp=document.createElement('div'); tmp.innerHTML=html||"";
    const walker=document.createTreeWalker(tmp, NodeFilter.SHOW_ELEMENT, null);
    const allowed=new Set(['B','STRONG','I','EM','U','BR','DIV','P']);
    const toRemove=[];
    while(walker.nextNode()){
      const el=walker.currentNode;
      if(!allowed.has(el.tagName)){
        while(el.firstChild) el.parentNode.insertBefore(el.firstChild, el);
        toRemove.push(el);
      }
    }
    toRemove.forEach(n=>n.remove());
    return tmp.innerHTML;
  }

  // Descarga compatible iOS
  function download(filename, data, type="application/octet-stream"){
    const blob = new Blob([data], {type});
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = filename;

    const isIOS = /iPad|iPhone|iPod/.test(navigator.userAgent) || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);
    if(isIOS){
      const reader = new FileReader();
      reader.onloadend = function () {
        const dataUrl = reader.result;
        const win = window.open(dataUrl, '_blank');
        if(!win){ alert("El navegador ha bloqueado la descarga. Permite popups para esta página."); }
        setTimeout(()=>URL.revokeObjectURL(url), 0);
      };
      reader.readAsDataURL(blob);
      return;
    }

    document.body.appendChild(a);
    a.click();
    setTimeout(()=>{ URL.revokeObjectURL(url); a.remove(); },0);
  }

  // ----- Normalizadores -----
  const SMALL = new Set(["y","e","de","del","la","las","los","el","al","a","en"]);
  const titleCaseEs = (s)=>{
    s = (s||"").trim().toLowerCase();
    if(!s) return "";
    return s.split(/\s+/).map((w,i)=>{
      if(i>0 && SMALL.has(w)) return w;
      return w.charAt(0).toUpperCase() + w.slice(1);
    }).join(" ");
  };
  const mapIndocumentadoAny = (s)=>{
    const t=(s||"").trim().toLowerCase();
    if(/^indocumentad[oa]\/?a?$/.test(t) || t==="indocumentado" || t==="indocumentada" || t==="indocumentado/a" ) return "Indocumentado/a";
    return s||"";
  };
  function normTipoDocLabel(sel, otro){
    switch(sel){
      case "DNI": return "DNI";
      case "NIE": return "NIE";
      case "Pasaporte": return "Pasaporte";
      case "Indocumentado/a": return "Indocumentado/a";
      case "Otro": return titleCaseEs(otro||"");
      default: return "";
    }
  }
  const normNumDoc = s => (s||"").toUpperCase();
  const normNombre = s => titleCaseEs(s);
  const normApellidos = s => (s||"").toUpperCase();
  const normPadres = s => titleCaseEs(s);
  const normDomicilio = s => titleCaseEs(s);
  const normLugarNac = s => titleCaseEs(s);

  function excelSerialToDMYString(n){
    const num = Number(n);
    if(!isFinite(num)) return "";
    const base = new Date(Date.UTC(1899,11,30));
    const ms = num * 86400000;
    const d = new Date(base.getTime()+ms);
    const dd = String(d.getUTCDate()).padStart(2,'0');
    const mm = String(d.getUTCMonth()+1).padStart(2,'0');
    const yy = d.getUTCFullYear();
    if(yy<1900 || yy>2400) return "";
    return `${dd}/${mm}/${yy}`;
  }

  function normalizeFiliacion(f){
    f.nombre    = normNombre(f.nombre);
    f.apellidos = normApellidos(f.apellidos);
    f.tipoSel   = f.tipoSel || "";
    f.otroDoc   = f.otroDoc || "";
    if(f.tipoSel==="Indocumentado" || f.tipoSel==="Indocumentada") f.tipoSel="Indocumentado/a";
    f.tipoDoc   = mapIndocumentadoAny(normTipoDocLabel(f.tipoSel, f.otroDoc));
    f.dni       = normNumDoc(f.dni);
    f.padres    = normPadres(f.padres);
    f.domicilio = normDomicilio(f.domicilio);
    f.fechaNac  = f.fechaNac || "";
    f.lugarNac  = normLugarNac(f.lugarNac || "");
    f.condSel   = f.condSel || "";
    f.condOtro  = f.condOtro || "";
    if(typeof f.fixedId !== 'number' || !isFinite(f.fixedId)){ f.fixedId = state.nextId++; }
    return f;
  }
  function nuevaFiliacion(){ return normalizeFiliacion({
    nombre:"", apellidos:"", tipoSel:"", otroDoc:"", tipoDoc:"", dni:"", padres:"",
    domicilio:"", telefono:"", fechaNac:"", lugarNac:"", condSel:"", condOtro:"", fixedId: state.nextId
  }); }

  // ----- ZIP/ODT/XLSX -----
  function crc32(u8){ let c=~0>>>0; for(let i=0;i<u8.length;i++){ c=(c>>>8)^CRC_TABLE[(c^u8[i])&0xFF]; } return (~c)>>>0; }
  const CRC_TABLE=(()=>{const t=new Uint32Array(256); for(let n=0;n<256;n++){let c=n; for(let k=0;k<8;k++){ c=(c&1)?(0xEDB88320^(c>>>1)):(c>>>1);} t[n]=c>>>0;} return t;})();
  function dosDateTime(d=new Date()){ const time=((d.getHours()<<11)|(d.getMinutes()<<5)|(Math.floor(d.getSeconds()/2)))&0xFFFF; const date=(((d.getFullYear()-1980)<<9)|((d.getMonth()+1)<<5)|d.getDate())&0xFFFF; return {time,date}; }
  class ZipWriter{
    constructor(){ this.entries=[]; this.parts=[]; this.offset=0; }
    addFile(name, data){
      const nameBytes = toUTF8(name);
      const content = (typeof data==="string")? toUTF8(data) : (data instanceof Uint8Array? data : new Uint8Array(data));
      const {time,date}=dosDateTime(); const crc=crc32(content);
      const local=new Uint8Array(30+nameBytes.length); const v=new DataView(local.buffer);
      v.setUint32(0,0x04034b50,true); v.setUint16(4,20,true); v.setUint16(6,0,true); v.setUint16(8,0,true);
      v.setUint16(10,time,true); v.setUint16(12,date,true); v.setUint32(14,crc,true);
      v.setUint32(18,content.length,true); v.setUint32(22,content.length,true);
      v.setUint16(26,nameBytes.length,true); v.setUint16(28,0,true);
      local.set(nameBytes,30);
      const offsetHere=this.offset; this.parts.push(local,content); this.offset+=local.length+content.length;
      this.entries.push({nameBytes, crc, size:content.length, time, date, offset:offsetHere});
    }
    finalize(){
      const centralParts=[]; let centralSize=0;
      for(const e of this.entries){
        const c=new Uint8Array(46+e.nameBytes.length); const v=new DataView(c.buffer);
        v.setUint32(0,0x02014b50,true); v.setUint16(4,20,true); v.setUint16(6,20,true);
        v.setUint16(8,0,true); v.setUint16(10,0,true); v.setUint16(12,e.time,true); v.setUint16(14,e.date,true);
        v.setUint32(16,e.crc,true); v.setUint32(20,e.size,true); v.setUint32(24,e.size,true);
        v.setUint16(28,e.nameBytes.length,true); v.setUint16(30,0,true); v.setUint16(32,0,true);
        v.setUint16(34,0,true); v.setUint16(36,0,true); v.setUint32(38,0,true); v.setUint32(42,e.offset,true);
        c.set(e.nameBytes,46); centralParts.push(c); centralSize+=c.length;
      }
      const centralOffset=this.offset; this.parts.push(...centralParts); this.offset+=centralSize;
      const end=new Uint8Array(22); const ve=new DataView(end.buffer);
      ve.setUint32(0,0x06054b50,true); ve.setUint16(4,0,true); ve.setUint16(6,0,true);
      ve.setUint16(8,this.entries.length,true); ve.setUint16(10,this.entries.length,true);
      ve.setUint32(12,centralSize,true); ve.setUint32(16,centralOffset,true); ve.setUint16(20,0,true);
      this.parts.push(end);
      let total=0; for(const p of this.parts) total+=p.length;
      const out=new Uint8Array(total); let off=0; for(const p of this.parts){ out.set(p,off); off+=p.length; }
      return out;
    }
  }

  function htmlToODTParagraphs(html){
    const tmp=document.createElement('div'); tmp.innerHTML = stripHTMLExceptBIU(html||"");
    const out=[];
    function nodeToText(node){
      if(node.nodeType===Node.TEXT_NODE) return escapeXml(node.nodeValue||"");
      if(node.nodeType!==Node.ELEMENT_NODE) return "";
      const tag=node.tagName;
      if(tag==="BR") return "<text:line-break/>";
      const inner=Array.from(node.childNodes).map(nodeToText).join("");
      if(tag==="B" || tag==="STRONG") return `<text:span text:style-name="B">${inner}</text:span>`;
      if(tag==="I" || tag==="EM")     return `<text:span text:style-name="I">${inner}</text:span>`;
      if(tag==="U")                   return `<text:span text:style-name="U">${inner}</text:span>`;
      return inner;
    }
    const blocks = Array.from(tmp.childNodes);
    if(!blocks.length){ out.push("<text:p/>"); return out.join(""); }
    for(const n of blocks){
      if(n.nodeType===Node.ELEMENT_NODE && (n.tagName==="DIV" || n.tagName==="P")){
        const inner=Array.from(n.childNodes).map(nodeToText).join("");
        out.push(`<text:p>${inner}</text:p>`);
      }else if(n.nodeType===Node.ELEMENT_NODE && n.tagName==="BR"){
        out.push("<text:p/>");
      }else{
        const t=nodeToText(n);
        if(t) out.push(`<text:p>${t}</text:p>`);
      }
    }
    return out.join("");
  }
  function makeODTFromHTML(title, html){
    const paras = htmlToODTParagraphs(html);
    const content =
`<?xml version="1.0" encoding="UTF-8"?>
<office:document-content
 xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
 xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
 xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
 xmlns:fo="urn:oasis:names:tc:opendocument:xsl-fo-compatible:1.0"
 office:version="1.2">
 <office:automatic-styles>
  <style:style style:name="B" style:family="text"><style:text-properties fo:font-weight="bold" style:font-weight-asian="bold" style:font-weight-complex="bold"/></style:style>
  <style:style style:name="I" style:family="text"><style:text-properties fo:font-style="italic" style:font-style-asian="italic" style:font-style-complex="italic"/></style:style>
  <style:style style:name="U" style:family="text"><style:text-properties style:text-underline-type="single" style:text-underline-style="solid"/></style:style>
 </office:automatic-styles>
 <office:body><office:text>
  <text:h text:outline-level="1">${escapeXml(title||'Documento')}</text:h>
  ${paras}
 </office:text></office:body>
</office:document-content>`;
    const manifest =
`<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
 <manifest:file-entry manifest:media-type="application/vnd.oasis.opendocument.text" manifest:full-path="/"/>
 <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="content.xml"/>
</manifest:manifest>`;
    const z=new ZipWriter();
    z.addFile("mimetype","application/vnd.oasis.opendocument.text");
    z.addFile("content.xml", content);
    z.addFile("META-INF/manifest.xml", manifest);
    return z.finalize();
  }

  const XLSX_LABELS = [
    "Nombre","Apellidos","Tipo de documento","Nº Documento","Sexo","Nacionalidad","Nombre de los Padres",
    "Fecha de nacimiento","Lugar de nacimiento","Domicilio","Teléfono","Delito","C.P. Agentes","Diligencias",
    "Instructor","Lugar del hecho","Lugar de la detención","Hora del hecho","Hora de la detención",
    "Breve resumen de los hechos","Indicios por los que se detiene","Abogado","Comunicarse con","Informar de detención",
    "Intérprete","Médico","Consulado","Indicativo","Fecha de generación","Condición"
  ];
  function valueForLabel(f, label){
    switch(label){
      case "Nombre": return f.nombre||"";
      case "Apellidos": return f.apellidos||"";
      case "Tipo de documento": return f.tipoDoc||"";
      case "Nº Documento": return f.dni||"";
      case "Nombre de los Padres": return f.padres||"";
      case "Fecha de nacimiento": return f.fechaNac||"";
      case "Lugar de nacimiento": return f.lugarNac||"";
      case "Domicilio": return f.domicilio||"";
      case "Teléfono": return f.telefono||"";
      case "Fecha de generación": return "";
      case "Condición": return (f.condSel==="Otro") ? (f.condOtro||"") : (f.condSel||"");
      default: return "";
    }
  }
  function sheetXMLForFiliacion(f){
    const rows = XLSX_LABELS.map((lab, idx)=>{
      const r = idx+1;
      const a = `<c r="A${r}" t="inlineStr"><is><t>${escapeXml(lab)}</t></is></c>`;
      const b = `<c r="B${r}" t="inlineStr"><is><t>${escapeXml(valueForLabel(f, lab))}</t></is></c>`;
      return `<row r="${r}">${a}${b}</row>`;
    }).join("");
    return `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:B${XLSX_LABELS.length}"/>
  <sheetData>${rows}</sheetData>
</worksheet>`;
  }
  function workbookXML(){
    return `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Resumen" sheetId="1" r:id="rId1"/></sheets>
</workbook>`;
  }
  function workbookRelsXML(){
    return `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`;
  }
  function rootRelsXML(){
    return `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`;
  }
  function contentTypesXML(){
    return `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`;
  }
  function corePropsXML(){
    const now = new Date().toISOString();
    return `<?xml version="1.0" encoding="UTF-8"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
 xmlns:dc="http://purl.org/dc/elements/1.1/"
 xmlns:dcterms="http://purl.org/dc/terms/"
 xmlns:dcmitype="http://purl.org/dc/dcmitype/"
 xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:creator>Editor Filiaciones (Móvil)</dc:creator>
  <cp:lastModifiedBy>Editor Filiaciones (Móvil)</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>
</cp:coreProperties>`;
  }
  function appPropsXML(){
    return `<?xml version="1.0" encoding="UTF-8"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
 xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Microsoft Excel</Application>
</Properties>`;
  }
  function makeXLSXFromFiliacion(f){
    const z=new ZipWriter();
    z.addFile("[Content_Types].xml", contentTypesXML());
    z.addFile("_rels/.rels", rootRelsXML());
    z.addFile("docProps/core.xml", corePropsXML());
    z.addFile("docProps/app.xml", appPropsXML());
    z.addFile("xl/workbook.xml", workbookXML());
    z.addFile("xl/_rels/workbook.xml.rels", workbookRelsXML());
    z.addFile("xl/worksheets/sheet1.xml", sheetXMLForFiliacion(f));
    return z.finalize();
  }

  function fileBaseFromTitle(){
    const t = (state.titulo||"Proyecto").trim();
    const date = todayISO();
    let base = `${t} ${date}`;
    base = base.replace(/\s+/g,' ').trim();
    const safe = base.normalize("NFKD").replace(/[\u0300-\u036f]/g,"").replace(/[^\w\- ]+/g,"").replace(/\s+/g,"_");
    return safe.length>50 ? safe.slice(0,50) : safe;
  }
  function fileNameForFiliacion(f){
    const nombre=(f.nombre||"").trim();
    const ap1=(f.apellidos||"").trim().split(/\s+/)[0]||"";
    let base=[nombre, ap1].filter(Boolean).join(" ").trim();
    if(!base) base=`filiacion_${f.fixedId}`;
    base=base.replace(/[\\/:*?"<>|]/g,"_");
    return `${base}.xlsx`;
  }

  // ----- Export API -----
  window.Core = {
    $, LS_KEY, state, save, load, hardResetStorage,
    saveDebounced, saveDocDebounced,
    toUTF8, fromUTF8, todayISO, escapeXml, escapeHtml, stripHTMLExceptBIU, download,
    titleCaseEs, mapIndocumentadoAny, normTipoDocLabel, excelSerialToDMYString,
    normNumDoc, normNombre, normApellidos, normPadres, normDomicilio, normLugarNac,
    normalizeFiliacion, nuevaFiliacion,
    ZipWriter, makeXLSXFromFiliacion, fileBaseFromTitle, fileNameForFiliacion,
    openedIndexRef: { get value(){ return openedIndex; }, set value(v){ openedIndex=v; } }
  };
})();
