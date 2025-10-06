// ============ IO: import/export (XLSX, JSON, proyecto), ZIP reader ============

(function(){
  const { $, state, save, normalizeFiliacion,
          mapIndocumentadoAny, normTipoDocLabel, excelSerialToDMYString,
          makeXLSXFromFiliacion, download, fileBaseFromTitle, fileNameForFiliacion,
          toUTF8 } = window.Core;

  // --------- ZIP Reader (deflate / deflate-raw con fallback) ----------
  class ZipReader{
    constructor(u8){ this.u8=u8; this.dv=new DataView(u8.buffer); }
    async readText(path){ const data=await this.readFile(path); return new TextDecoder().decode(data); }
    async exists(path){ try{ await this.readFile(path); return true; }catch{ return false; } }

    async inflateRawOrZlib(comp){
      if(typeof DecompressionStream === 'undefined'){
        throw new Error("Este navegador no soporta descompresión ZIP (DecompressionStream).");
      }
      try{
        const dsRaw=new DecompressionStream('deflate-raw');
        const ab=await new Response(new Blob([comp]).stream().pipeThrough(dsRaw)).arrayBuffer();
        return new Uint8Array(ab);
      }catch(_){
        const ds=new DecompressionStream('deflate');
        const ab=await new Response(new Blob([comp]).stream().pipeThrough(ds)).arrayBuffer();
        return new Uint8Array(ab);
      }
    }

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
        const name=new TextDecoder().decode(this.u8.slice(p+46,p+46+nameLen));
        const next=p+46+nameLen+extraLen+commLen;
        if(name===path){
          const sig2=this.dv.getUint32(lhoff,true); if(sig2!==0x04034b50) throw new Error("Local header inválido");
          const nlen=this.dv.getUint16(lhoff+26,true);
          const xlen=this.dv.getUint16(lhoff+28,true);
          const dataStart=lhoff+30+nlen+xlen;
          const comp=this.u8.slice(dataStart,dataStart+csize);
          if(compMethod===0){ return comp; }
          if(compMethod===8){
            return await this.inflateRawOrZlib(comp);
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

  // -------- Import XLSX --------
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
          if(t==="inlineStr"){
            const is=c.getElementsByTagName("is")[0];
            const tnode=is?.getElementsByTagName("t")[0];
            return tnode?tnode.textContent||"" : "";
          }
          const val = v? v.textContent||"" : "";
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
        const tipoSel = (["DNI","NIE","PASAPORTE"].includes((tipoRaw||"").toUpperCase()))
          ? (tipoRaw.toUpperCase()==="PASAPORTE" ? "Pasaporte" : tipoRaw.toUpperCase())
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

        state.filiaciones.push(f);
        lastIdx = state.filiaciones.length-1;
      }catch(err){
        alert(`No se pudo importar “${file.name}”: ${err.message}`);
      }
    }
    if(lastIdx>=0){
      save();
      window.UI?.renderFiliaciones?.();
    }
  }

  // -------- Import JSON --------
  async function importJsonFiles(fileList){
    for(const file of fileList){
      try{
        const data = JSON.parse(await file.text());
        const list = Array.isArray(data) ? data : [data];
        for(const raw of list){
          const f = normalizeFiliacion(raw);
          state.filiaciones.push(f);
        }
      }catch(err){
        alert(`No se pudo importar “${file.name}”: ${err.message}`);
      }
    }
    save();
    window.UI?.renderFiliaciones?.();
  }

  // -------- Exportaciones de una ficha --------
  function exportFiliacionXlsxByIndex(i){
    const f = state.filiaciones[i];
    const xlsx = makeXLSXFromFiliacion(f);
    download(fileNameForFiliacion(f), xlsx, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  }
  function exportFiliacionJsonByIndex(i){
    const f = state.filiaciones[i];
    const pretty = JSON.stringify(f, null, 2);
    download(`filiacion_${f.fixedId||i+1}.json`, toUTF8(pretty), "application/json");
  }

  // -------- Guardar/Cargar proyecto --------
  function saveProject(){
    const base=fileBaseFromTitle();
    const json=toUTF8(JSON.stringify({...state, doc: window.Core.getDocHTML()}, null, 2));
    download(`${base}.json`, json, "application/json");
  }
  async function loadProjectFromFile(file){
    const data=JSON.parse(await file.text());
    if(!data || !Array.isArray(data.filiaciones)) throw new Error("Formato inválido");
    // recomponer nextId si falta o no es válido
    if(typeof data.nextId !== 'number' || !isFinite(data.nextId)){
      const maxId = Math.max(0, ...data.filiaciones.map(ff=>ff.fixedId||0));
      data.nextId = maxId + 1;
    }
    state.filiaciones = data.filiaciones.map(ff=>normalizeFiliacion(Object.assign({tipoSel:"", otroDoc:"", condSel:"", condOtro:""}, ff)));
    state.titulo = data.titulo || "";
    state.doc    = data.doc || "";
    state.nextId = data.nextId;
    save();
  }

  // -------- Export ODT (si lo usas con algún botón) --------
  function exportODT(){
    const base=fileBaseFromTitle(); const title=state.titulo||"Documento";
    const html = window.Core.getDocHTML(); const odt=window.Core.makeODTFromHTML(title, html);
    download(`${base}.odt`, odt, "application/vnd.oasis.opendocument.text");
  }

  // Exponer IO
  window.IO = {
    importXlsxFiles, importJsonFiles,
    exportFiliacionXlsxByIndex, exportFiliacionJsonByIndex,
    saveProject, loadProjectFromFile, exportODT
  };

  // Wire de botones existentes en HTML (móvil)
  window.addEventListener('DOMContentLoaded', ()=>{
    // Import XLSX
    $('#importXlsxBtn')?.addEventListener('click', ()=>$('#importXlsxInput').click());
    $('#importXlsxInput')?.addEventListener('change', async (e)=>{
      const files=e.target.files; if(!files?.length) return;
      await importXlsxFiles(files);
      e.target.value="";
    });

    // Import JSON
    $('#importJsonBtn')?.addEventListener('click', ()=>$('#importJsonInput').click());
    $('#importJsonInput')?.addEventListener('change', async (e)=>{
      const files=e.target.files; if(!files?.length) return;
      await importJsonFiles(files);
      e.target.value="";
    });

    // Guardar/Cargar proyecto
    $('#saveProjectBtn')?.addEventListener('click', ()=> saveProject());
    $('#loadProjectBtn')?.addEventListener('click', ()=> $('#loadProjectInput').click());
    $('#loadProjectInput')?.addEventListener('change', async (e)=>{
      const file=e.target.files?.[0]; if(!file) return;
      try{
        await loadProjectFromFile(file);
        window.UI?.afterProjectLoaded?.();
        alert("Proyecto cargado correctamente.");
      }catch(err){
        alert("No se pudo cargar el proyecto: " + err.message);
      }finally{
        e.target.value="";
      }
    });

    // Imprimir
    $('#printBtn')?.addEventListener('click', ()=>window.print());
  });

})();
