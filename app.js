/* ===== Utilidades ===== */
const $ = s=>document.querySelector(s);
const $$ = s=>Array.from(document.querySelectorAll(s));
function debounce(fn, wait=800){ let t; return (...a)=>{ clearTimeout(t); t=setTimeout(()=>fn(...a), wait); }; }
const saveDebounced = debounce(()=>save(), 800);
const saveDocDebounced = debounce(()=>save(), 1200);
function download(filename, data, type="application/octet-stream"){
  const blob = new Blob([data], {type}); const url = URL.createObjectURL(blob);
  const a = document.createElement('a'); a.href=url; a.download=filename; document.body.appendChild(a); a.click();
  setTimeout(()=>{URL.revokeObjectURL(url); a.remove();},0);
}
function toUTF8(s){ return new TextEncoder().encode(s); }
function fromUTF8(u8){ return new TextDecoder().decode(u8); }
function todayISO(){ const d=new Date(), pad=n=>String(n).padStart(2,'0'); return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`; }
function escapeXml(s){ return String(s??"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&apos;"); }
function escapeHtml(s){ return (s??"").replace(/[&<>"']/g, m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m])); }
function stripHTMLExceptBIU(html){
  const tmp=document.createElement('div'); tmp.innerHTML=html||"";
  const walker=document.createTreeWalker(tmp, NodeFilter.SHOW_ELEMENT, null);
  const allowed=new Set(['B','STRONG','I','EM','U','BR']);
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

/* ===== Estado ===== */
const LS_KEY="gestor_partes_comparecencias_mobile_v1";
const state={filiaciones:[], titulo:"", doc:""}; // doc = HTML editor
let openedIndex = -1;
function save(){ try{ localStorage.setItem(LS_KEY, JSON.stringify(state)); }catch{} }
function load(){ try{ Object.assign(state, JSON.parse(localStorage.getItem(LS_KEY)||"{}")); }catch{} }

/* ===== Normalizadores ===== */
const SMALL = new Set(["y","e","de","del","la","las","los","el","al","a","en"]);
function titleCaseEs(s){
  s = (s||"").trim().toLowerCase();
  if(!s) return "";
  return s.split(/\s+/).map((w,i)=>{
    if(i>0 && SMALL.has(w)) return w;
    return w.charAt(0).toUpperCase() + w.slice(1);
  }).join(" ");
}
function mapIndocumentadoAny(s){
  const t=(s||"").trim().toLowerCase();
  if(/^indocumentad[oa]\/?a?$/.test(t) || t==="indocumentado" || t==="indocumentada" || t==="indocumentado/a" ) return "Indocumentado/a";
  return s||"";
}
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

// Excel serial date -> dd/mm/yyyy
function excelSerialToDMYString(n){
  const num = Number(n);
  if(!isFinite(num)) return "";
  const base = new Date(Date.UTC(1899,11,30)); // Excel 1900 bug
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
  return f;
}
function nuevaFiliacion(){ return normalizeFiliacion({ 
  nombre:"", apellidos:"", tipoSel:"", otroDoc:"", tipoDoc:"", dni:"", padres:"",
  domicilio:"", telefono:"", fechaNac:"", lugarNac:"", condSel:"", condOtro:""
}); }

/* ===== Coletillas (solo etiqueta visible) ===== */
const COLETILLAS = [
  { label:"Info derechos", text:"Resulta conveniente hacer constar que se ha informado a las partes de sus derechos y obligaciones." },
  { label:"Advertencia plazo", text:"Se advierte a la persona interesada de que la falta de respuesta en el plazo conferido podrá entenderse como desistimiento." },
  { label:"Unión de escrito", text:"Queda unido a las actuaciones el escrito presentado, dándose por reproducido su contenido a los efectos oportunos." },
  { label:"Notificación y recursos", text:"Notifíquese a las partes personadas, con indicación de los recursos que procedan." }
];

/* ===== ZIP Writer (XLSX/ODT) ===== */
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

/* ===== ODT (preserva B/I/U) ===== */
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

/* ===== XLSX (plantilla “Resumen” + Condición al final) ===== */
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
  <dc:creator>Editor Filiaciones</dc:creator>
  <cp:lastModifiedBy>Editor Filiaciones</cp:lastModifiedBy>
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

/* ===== Helpers ===== */
function fileBaseFromTitle(){
  const t = (state.titulo||"Proyecto").trim();
  const date = todayISO();
  let base = `${t} ${date}`;
  base = base.replace(/\s+/g,' ').trim();
  const safe = base.normalize("NFKD").replace(/[\u0300-\u036f]/g,"").replace(/[^\w\- ]+/g,"").replace(/\s+/g,"_");
  return safe.length>50 ? safe.slice(0,50) : safe;
}
function fileNameForFiliacion(f, idx){
  const nombre=(f.nombre||"").trim();
  const ap1=(f.apellidos||"").trim().split(/\s+/)[0]||"";
  let base=[nombre, ap1].filter(Boolean).join(" ").trim();
  if(!base) base=`filiacion_${idx+1}`;
  base=base.replace(/[\\/:*?"<>|]/g,"_");
  return `${base}.xlsx`;
}

/* ===== Importar XLSX ===== */
class ZipReader{
  constructor(u8){ this.u8=u8; this.dv=new DataView(u8.buffer); }
  async readText(path){ const data=await this.readFile(path); return fromUTF8(data); }
  async exists(path){ try{ await this.readFile(path); return true; }catch{ return false; } }
  async readFile(path){
    const eocd=this.findEOCD(); if(!eocd) throw new Error("EOCD no encontrado");
    const cdOff=eocd.cdOffset, cdSize=eocd.cdSize; const cdEnd=cdOff+cdSize;
    let p=cdOff;
    while(p<cdEnd){
      const sig=this.dv.getUint32(p,true); if(sig!==0x02014b50) break;
      const compMethod=this.dv.getUint16(p+10,true);
      const csize=this.dv.getUint32(p+20,true);
      const nameLen=this.dv.getUint16(p+28,true);
      const extraLen=this.dv.getUint16(p+30,true);
      const commLen=this.dv.getUint16(p+32,true);
      const lhoff=this.dv.getUint32(p+42,true);
      const name=fromUTF8(this.u8.slice(p+46,p+46+nameLen));
      const next=p+46+nameLen+extraLen+commLen;
      if(name===path){
        const sig2=this.dv.getUint32(lhoff,true); if(sig2!==0x04034b50) throw new Error("Local header inválido");
        const nlen=this.dv.getUint16(lhoff+26,true);
        const xlen=this.dv.getUint16(lhoff+28,true);
        const dataStart=lhoff+30+nlen+xlen;
        const comp=this.u8.slice(dataStart,dataStart+csize);
        if(compMethod===0){ return comp; }
        if(compMethod===8){
          if(typeof DecompressionStream==="undefined") throw new Error("Navegador sin 'DecompressionStream' (deflate). Usa Chrome/Edge actual.");
          const ds=new DecompressionStream('deflate-raw');
          const ab=await new Response(new Blob([comp]).stream().pipeThrough(ds)).arrayBuffer();
          return new Uint8Array(ab);
        }
        throw new Error("Compresión no soportada: "+compMethod);
      }
      p=next;
    }
    throw new Error("Archivo no encontrado en ZIP: "+path);
  }
  findEOCD(){
    const u8=this.u8; const start=Math.max(0,u8.length-0xFFFF);
    for(let i=u8.length-22;i>=start;i--){
      if(this.dv.getUint32(i,true)===0x06054b50){
        const cdSize=this.dv.getUint32(i+12,true);
        const cdOffset=this.dv.getUint32(i+16,true);
        return {cdSize,cdOffset,offset:i};
      }
    }
    return null;
  }
}
async function importXlsxFiles(fileList){
  let lastIdx = -1;
  for(const file of fileList){
    try{
      const buf=new Uint8Array(await file.arrayBuffer());
      const zip=new ZipReader(buf);
      const wbXml=await zip.readText("xl/workbook.xml");
      const wb=new DOMParser().parseFromString(wbXml,"application/xml");
      const sheets=Array.from(wb.getElementsByTagName("sheet"));
      const resumen=sheets.find(s=>(s.getAttribute("name")||"").toLowerCase()==="resumen");
      if(!resumen) throw new Error("Hoja 'Resumen' no encontrada.");
      const rid=resumen.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships","id")||resumen.getAttribute("r:id");
      const relsXml=await zip.readText("xl/_rels/workbook.xml.rels");
      const rel=Array.from(new DOMParser().parseFromString(relsXml,"application/xml").getElementsByTagName("Relationship")).find(r=>r.getAttribute("Id")===rid);
      let target=rel?.getAttribute("Target")||"worksheets/sheet1.xml";
      if(!target.startsWith("worksheets/")) target="worksheets/sheet1.xml";
      const sheetXml=await zip.readText("xl/"+target);
      let shared=[];
      if(await zip.exists("xl/sharedStrings.xml")){
        const sXml=await zip.readText("xl/sharedStrings.xml");
        const sDoc=new DOMParser().parseFromString(sXml,"application/xml");
        shared=Array.from(sDoc.getElementsByTagName("si")).map(si=>Array.from(si.getElementsByTagName("t")).map(t=>t.textContent||"").join(""));
      }
      const ws=new DOMParser().parseFromString(sheetXml,"application/xml");
      const cells=Array.from(ws.getElementsByTagName("c"));
      const byRow={};
      function readCell(c){
        const t=c.getAttribute("t"); const v=c.getElementsByTagName("v")[0];
        if(t==="s"){ const idx=v?parseInt(v.textContent||"0",10):0; return shared[idx]??""; }
        if(t==="inlineStr"){ const is=c.getElementsByTagName("is")[0]; const tnode=is?.getElementsByTagName("t")[0]; return tnode?tnode.textContent||"" : ""; }
        const val = v? v.textContent||"" :"";
        if(/^\d+(\.\d+)?$/.test(val)){
          const asDate = excelSerialToDMYString(val);
          return asDate || val;
        }
        return val;
      }
      cells.forEach(c=>{
        const ref=c.getAttribute("r")||""; const m=ref.match(/^([A-Z]+)(\d+)$/); if(!m) return;
        const col=m[1]; const row=parseInt(m[2],10); const val=readCell(c);
        byRow[row]=byRow[row]||{}; if(col==="A") byRow[row].A=val; else if(col==="B") byRow[row].B=val;
      });
      const map={}; Object.values(byRow).forEach(r=>{ if(r?.A) map[r.A.trim()]=(r.B||"").trim(); });

      let tipoRaw = map["Tipo de documento"] || "";
      tipoRaw = mapIndocumentadoAny(tipoRaw);
      const tipoSel = (["DNI","NIE","PASAPORTE"].includes((tipoRaw||"").toUpperCase())) ? (tipoRaw.toUpperCase()==="PASAPORTE"?"Pasaporte":tipoRaw.toUpperCase())
                    : (tipoRaw==="Indocumentado/a" ? "Indocumentado/a" : (tipoRaw ? "Otro" : ""));

      const f = normalizeFiliacion({
        nombre    : map["Nombre"] || "",
        apellidos : map["Apellidos"] || "",
        tipoSel   : tipoSel,
        otroDoc   : (tipoSel==="Otro") ? tipoRaw : "",
        tipoDoc   : "",
        dni       : map["Nº Documento"] || "",
        fechaNac  : map["Fecha de nacimiento"] || "",
        lugarNac  : map["Lugar de nacimiento"] || "",
        padres    : map["Nombre de los Padres"] || "",
        domicilio : map["Domicilio"] || "",
        telefono  : map["Teléfono"] || "",
        condSel   : "", condOtro: ""
      });

      state.filiaciones.push(f); lastIdx = state.filiaciones.length-1;
    }catch(err){ alert(`No se pudo importar “${file.name}”: ${err.message}`); }
  }
  if(lastIdx>=0){ save(); openedIndex = lastIdx; renderFiliaciones(); }
}

/* ===== UI Filiaciones ===== */
function getTipoDocShown(f){
  return f.tipoSel==="Otro" ? (f.otroDoc?f.otroDoc:f.tipoDoc) : f.tipoSel || f.tipoDoc || "";
}
function isCondValid(f){ return f.condSel && (f.condSel!=="Otro" || (f.condOtro && f.condOtro.trim()!=="")); }
function computeTituloFicha(f,i){
  const titulo=(f.nombre||f.apellidos) ? `${(f.nombre||'').trim()} ${(f.apellidos||'').trim().split(/\s+/)[0]||''}`.trim() : `Filiación ${i+1}`;
  const tdoc=getTipoDocShown(f);
  const parts=[]; if(tdoc) parts.push(tdoc); if(f.dni) parts.push(f.dni);
  return {titulo:titulo||`Filiación ${i+1}`, subtitulo:parts.length?`(${parts.join(' · ')})`:""};
}
function renderFiliaciones(){
  const cont=$("#filiaciones"); cont.innerHTML="";
  state.filiaciones.forEach((f,i)=>{
    const det=document.createElement('details'); det.className="f-item"; if(openedIndex===i) det.open=true;
    const info=computeTituloFicha(f,i);
    det.innerHTML=`
      <summary>
        <span class="tag">#${String(i+1).padStart(2,'0')}</span>
        <span class="title">${info.titulo}</span>
        ${info.subtitulo?`<span class="sub">${info.subtitulo}</span>`:""}
        <span class="caret"></span>
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
        <div class="btn-row" style="margin-top:8px">
          <button class="btn ghost tiny" data-up="${i}" title="Subir">↑</button>
          <button class="btn ghost tiny" data-down="${i}" title="Bajar">↓</button>
          <button class="btn secondary tiny" data-add="${i}" title="Añadir al documento" ${!isCondValid(f)?'disabled':''}>➜ Añadir</button>
          <button class="btn secondary tiny" data-xlsx="${i}" title="Descargar XLSX" ${!isCondValid(f)?'disabled':''}>⬇️ XLSX</button>
          <button class="btn danger tiny" data-del="${i}" title="Eliminar">🗑️</button>
        </div>
      </div>`;
    cont.appendChild(det);
  });
  $('#emptyHint').style.display = state.filiaciones.length ? 'none':'block';

  // inputs/selects (edición + normalización)
  cont.querySelectorAll('input[data-k], select[data-k]').forEach(inp=>{
    const handlerInput = e=>{
      const i=+e.target.dataset.i, k=e.target.dataset.k;
      if(k==="fechaNac"){
        let v=e.target.value.replace(/\D/g,'').slice(0,8);
        let out=""; if(v.length>=2){ out+=v.slice(0,2)+"/"; } else { out+=v; }
        if(v.length>=4){ out+=v.slice(2,4)+"/"; } else if(v.length>2){ out+=v.slice(2); }
        if(v.length>4){ out+=v.slice(4); }
        e.target.value=out; state.filiaciones[i][k]=out; saveDebounced(); return;
      }
      state.filiaciones[i][k]=e.target.value;
      if(k==="tipoSel"){
        const det=e.target.closest('details');
        const otherCol = det.querySelector('input[data-k="otroDoc"]')?.closest('.col');
        if(otherCol){ otherCol.style.display = e.target.value==="Otro" ? "" : "none"; }
      }
      if(k==="condSel"){
        const det=e.target.closest('details');
        const otherCol = det.querySelector('input[data-k="condOtro"]')?.closest('.col');
        if(otherCol){ otherCol.style.display = e.target.value==="Otro" ? "" : "none"; }
      }
      saveDebounced();
      if(["nombre","apellidos","tipoSel","otroDoc","dni"].includes(k)){
        const f=state.filiaciones[i];
        if(f.tipoSel==="Indocumentado" || f.tipoSel==="Indocumentada") f.tipoSel="Indocumentado/a";
        f.tipoDoc = mapIndocumentadoAny(normTipoDocLabel(f.tipoSel, f.otroDoc));
        const det=e.target.closest('details'); if(det){
          const info=computeTituloFicha(f, i);
          const tEl=det.querySelector('.title'); const sEl=det.querySelector('.sub');
          if(tEl) tEl.textContent=info.titulo;
          if(info.subtitulo){ if(sEl) sEl.textContent=info.subtitulo; else det.querySelector('summary').insertAdjacentHTML('beforeend', `<span class="sub">${info.subtitulo}</span>`); }
          else if(sEl){ sEl.remove(); }
        }
      }
      if(k==="condSel" || k==="condOtro"){
        const f=state.filiaciones[i]; const valid=isCondValid(f);
        const det=e.target.closest('details'); det.querySelector('[data-add]').disabled=!valid; det.querySelector('[data-xlsx]').disabled=!valid;
      }
    };
    const handlerBlur = e=>{
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
          const info=computeTituloFicha(f, i);
          det.querySelector('.title').textContent=info.titulo;
          const sEl=det.querySelector('.sub');
          if(info.subtitulo){ if(sEl) sEl.textContent=info.subtitulo; else det.querySelector('summary').insertAdjacentHTML('beforeend', `<span class="sub">${info.subtitulo}</span>`); }
          else if(sEl){ sEl.remove(); }
        }
      }
    };
    inp.addEventListener('input', handlerInput, {passive:true});
    inp.addEventListener('blur', handlerBlur, {passive:true});
  });

  // acciones
  cont.querySelectorAll('button[data-del]').forEach(b=>b.onclick=()=>{ const idx=+b.dataset.del; state.filiaciones.splice(idx,1); save(); openedIndex=-1; renderFiliaciones(); });
  cont.querySelectorAll('button[data-up]').forEach(b=>b.onclick=()=>{ const i=+b.dataset.up; if(i<=0)return; [state.filiaciones[i-1],state.filiaciones[i]]=[state.filiaciones[i],state.filiaciones[i-1]]; save(); openedIndex=i-1; renderFiliaciones();});
  cont.querySelectorAll('button[data-down]').forEach(b=>b.onclick=()=>{ const i=+b.dataset.down; if(i>=state.filiaciones.length-1)return; [state.filiaciones[i+1],state.filiaciones[i]]=[state.filiaciones[i],state.filiaciones[i+1]]; save(); openedIndex=i+1; renderFiliaciones();});
  cont.querySelectorAll('button[data-add]').forEach(b=>b.onclick=()=>{ 
    const f=state.filiaciones[+b.dataset.add]; 
    const txt = buildColetillaFromFiliacion(f);
    editorFocus();
    document.execCommand('insertHTML', false, `<b>${escapeHtml(txt)}</b>`); // en negrita, sin punto/salto
    state.doc = $('#doc').innerHTML; saveDocDebounced();
    $('#doc').focus();
  });
  cont.querySelectorAll('button[data-xlsx]').forEach(b=>b.onclick=()=>{ const idx=+b.dataset.xlsx; const f=state.filiaciones[idx]; const xlsx=makeXLSXFromFiliacion(f); download(fileNameForFiliacion(f, idx), xlsx, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"); });
}

/* ===== Coletilla por ficha (orden y reglas) ===== */
function buildColetillaFromFiliacion(f){
  const segs=[];
  const nom=[f.nombre,f.apellidos].filter(Boolean).join(' ').trim(); if(nom) segs.push(nom);

  const tdoc = getTipoDocShown(f);
  const isIndoc = /^Indocumentad[oa]\/?a?$/.test(tdoc||"") || (tdoc==="Indocumentado/a");
  if(isIndoc){ segs.push("Indocumentado/a"); }
  else{
    if(tdoc && f.dni){ segs.push(`con ${tdoc} número ${f.dni}`); }
    else if(tdoc){ segs.push(`con ${tdoc}`); }
    else if(f.dni){ segs.push(`con número ${f.dni}`); }
  }

  if(f.lugarNac || f.fechaNac){
    let born = "nacido/a";
    if(f.lugarNac) born += ` en ${f.lugarNac}`;
    if(f.fechaNac) born += ` el día ${f.fechaNac}`; // sin coma
    segs.push(born);
  }

  if(f.padres){ segs.push(`hijo/a de ${f.padres}`); }
  if(f.domicilio){ segs.push(`con domicilio en ${f.domicilio}`); }
  if(f.telefono){ segs.push(`teléfono ${f.telefono}`); }

  return segs.join(", ");
}

/* ===== Documento & proyecto ===== */
function editorEl(){ return $('#doc'); }
function editorFocus(){ editorEl().focus(); }
function getDocHTML(){ return editorEl().innerHTML; }
function setDocHTML(html){ editorEl().innerHTML = html||""; }
function exportODT(){
  const base=fileBaseFromTitle(); const title=state.titulo||"Documento";
  const html = getDocHTML(); const odt=makeODTFromHTML(title, html);
  download(`${base}.odt`, odt, "application/vnd.oasis.opendocument.text");
}
$('#printBtn').onclick=()=>window.print();
$('#saveProjectBtn').onclick=()=>{ const base=fileBaseFromTitle(); const json=toUTF8(JSON.stringify({...state, doc:getDocHTML()}, null, 2)); download(`${base}.json`, json, "application/json"); };
$('#loadProjectBtn').onclick=()=>$('#loadProjectInput').click();
$('#loadProjectInput').addEventListener('change', async (e)=>{
  const file=e.target.files?.[0]; if(!file) return;
  try{
    const data=JSON.parse(await file.text());
    if(!data || !Array.isArray(data.filiaciones)) throw new Error("Formato inválido");
    state.filiaciones=data.filiaciones.map(ff=>normalizeFiliacion(Object.assign({tipoSel:"", otroDoc:"", condSel:"", condOtro:""}, ff)));
    state.titulo=data.titulo||""; state.doc=data.doc||"";
    $('#titulo').value=state.titulo; setDocHTML(state.doc);
    save(); openedIndex = state.filiaciones.length? state.filiaciones.length-1 : -1; renderFiliaciones();
    alert("Proyecto cargado correctamente.");
  }catch(err){ alert("No se pudo cargar el proyecto: " + err.message); }
  finally{ e.target.value=""; }
});
$('#titulo').addEventListener('blur', ()=>{ state.titulo = titleCaseEs($('#titulo').value||""); $('#titulo').value = state.titulo; saveDocDebounced(); });
$('#titulo').addEventListener('input', ()=>{ state.titulo=$('#titulo').value; saveDocDebounced(); }, {passive:true});

/* Mayúscula tras punto y tras “-- ” (con fallback para móviles) */
let capNext = false;
$('#doc').addEventListener('beforeinput', (e)=>{
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
$('#doc').addEventListener('input', (e)=>{
  // Fallback: si el evento input trae el punto como último carácter
  if(e.inputType==='insertText' && e.data === '.') capNext = true;
  state.doc=getDocHTML(); saveDocDebounced();
});
$('#doc').addEventListener('keydown', (e)=>{
  if(e.key === 'Enter'){
    e.preventDefault();
    document.execCommand('insertHTML', false, '<br>-- &nbsp;');
    capNext = true;
    state.doc=getDocHTML(); saveDocDebounced();
    return;
  }
  if(e.key === '.'){ capNext = true; }
});
$('#doc').addEventListener('focus', ()=>{ $$('#filiaciones details[open]').forEach(d=>d.open=false); });

// Formato B/I/U
function cmd(name){ editorFocus(); document.execCommand(name, false, null); state.doc=getDocHTML(); saveDocDebounced(); }
$('#boldBtn').onclick=()=>cmd('bold'); $('#italicBtn').onclick=()=>cmd('italic'); $('#underBtn').onclick=()=>cmd('underline');

/* Refrescar (borra también filiaciones) */
function refreshAll(){ renderFiliaciones(); renderColetillas(); }
$('#refreshTopBtn').onclick=()=>{ if(confirm("Esto refrescará y borrará TODAS las filiaciones. ¿Continuar?")){ state.filiaciones=[]; openedIndex=-1; save(); refreshAll(); $('#emptyHint').style.display='block'; } };

/* Coletillas (modal) */
function renderColetillas(){
  const cont=$("#coletillasList"); cont.innerHTML="";
  COLETILLAS.forEach(c=>{
    const el=document.createElement('div'); el.className="coletilla";
    el.innerHTML = `<div class="label">${c.label}</div>
                    <div><button class="btn secondary tiny">➜ Insertar</button></div>`;
    el.querySelector('button').onclick=()=>{ 
      editorFocus(); document.execCommand('insertText', false, c.text);
      state.doc = $('#doc').innerHTML; saveDocDebounced();
      closeColetillas(); editorFocus();
    };
    cont.appendChild(el);
  });
}
function openColetillas(){ $('#coletillasModal').classList.add('show'); $('#coletillasModal').setAttribute('aria-hidden','false'); renderColetillas(); }
function closeColetillas(){ $('#coletillasModal').classList.remove('show'); $('#coletillasModal').setAttribute('aria-hidden','true'); refreshAll(); }
$('#openColetillasBtn').onclick=openColetillas;
$('#closeColetillasBtn').onclick=closeColetillas;
$('#coletillasModal').addEventListener('click', e=>{ if(e.target.id==='coletillasModal'){ closeColetillas(); } });
window.addEventListener('keydown', e=>{ if(e.key==='Escape') closeColetillas(); });

/* Importar XLSX */
$('#importXlsxBtn').onclick=()=>$('#importXlsxInput').click();
$('#importXlsxInput').addEventListener('change', async (e)=>{ const files=e.target.files; if(!files?.length) return; await importXlsxFiles(files); e.target.value=""; });

/* Añadir ficha nueva */
$('#addFBtn').onclick=()=>{ state.filiaciones.push(nuevaFiliacion()); save(); openedIndex = state.filiaciones.length-1; renderFiliaciones(); };

/* Export ODT con Ctrl/Cmd + 1 */
window.addEventListener('keydown', (e)=>{ if((e.ctrlKey||e.metaKey) && e.key === '1'){ e.preventDefault(); exportODT(); } });

/* Init */
(function init(){
  load();
  $('#titulo').value=state.titulo||"";
  setDocHTML(state.doc||"");
  openedIndex = -1;
  renderFiliaciones();
  $('#emptyHint').style.display = state.filiaciones.length ? 'none':'block';
})();
