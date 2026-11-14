(function(){
  function init(){
    const root=document.querySelector('#detraccionesMasivasView');
    const card=root?.querySelector('.det-paste-card');
    if(!card||card.dataset.excelUnified==='1')return;
    card.dataset.excelUnified='1';
    document.getElementById('det-excel-tools')?.remove();

    const head=card.querySelector('.det-card-head h2');
    const desc=card.querySelector('.det-card-head p');
    if(head)head.textContent='3. Carga masiva desde Excel';
    if(desc)desc.textContent='Elige una sola forma de importar: pega directamente desde Excel o descarga y completa la plantilla. Ambas opciones usan las mismas validaciones.';
    const body=card.querySelector('.det-card-body');
    const guide=body.querySelector('#det-paste-guide');
    const pasteLabel=body.querySelector('label[for="det-paste-input"]');
    const pasteInput=body.querySelector('#det-paste-input');
    const pasteActions=body.querySelector('.det-paste-actions');
    const pasteMessage=body.querySelector('#det-paste-message');
    const note=body.querySelector('.det-note');
    if(!guide||!pasteInput||!pasteActions||!pasteMessage)return;

    const style=document.createElement('style');
    style.textContent='.det-import-tabs{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:14px}.det-import-tab{border:1px solid #d6e0e8;background:#f5f8fb;color:#24445f;border-radius:10px;padding:10px 14px;font-weight:800;cursor:pointer}.det-import-tab[aria-selected="true"]{background:#163a5f;color:#fff;border-color:#163a5f}.det-import-panel{border:1px solid #e0e7ee;border-radius:13px;padding:15px;background:#fff}.det-import-panel[hidden]{display:none}.det-template-row{display:flex;gap:9px;flex-wrap:wrap;align-items:center}.det-template-help{font-size:11px;color:#667085;line-height:1.55;margin:10px 0}.det-file-label{display:inline-flex;align-items:center;gap:7px;border:1px solid #163a5f;background:#163a5f;color:#fff;border-radius:11px;padding:11px 14px;font-weight:900;cursor:pointer}.det-file-label input{display:none}.det-role-template{font-size:12px;font-weight:800;color:#163a5f;margin-bottom:10px}';
    document.head.appendChild(style);

    const tabs=document.createElement('div');
    tabs.className='det-import-tabs';
    tabs.innerHTML='<button type="button" class="det-import-tab" data-panel="paste" aria-selected="true">Pegar desde Excel</button><button type="button" class="det-import-tab" data-panel="file" aria-selected="false">Subir plantilla Excel</button>';
    const pastePanel=document.createElement('div');pastePanel.className='det-import-panel';pastePanel.dataset.panel='paste';
    const filePanel=document.createElement('div');filePanel.className='det-import-panel';filePanel.dataset.panel='file';filePanel.hidden=true;
    [guide,pasteLabel,pasteInput,pasteActions,pasteMessage].forEach(n=>pastePanel.appendChild(n));
    const roleText=document.createElement('div');roleText.className='det-role-template';roleText.id='det-excel-role-text';
    const fileActions=document.createElement('div');fileActions.className='det-template-row';
    fileActions.innerHTML='<button type="button" class="det-btn secondary" id="det-excel-template">⬇ Descargar plantilla Excel</button><label class="det-file-label">＋ Seleccionar plantilla completada<input id="det-excel-file" type="file" accept=".xlsx,.xls,.csv"></label>';
    const fileMsg=document.createElement('div');fileMsg.id='det-excel-status';fileMsg.className='det-validation';fileMsg.style.display='none';fileMsg.style.marginTop='12px';
    const fileHelp=document.createElement('p');fileHelp.className='det-template-help';fileHelp.textContent='La plantilla cambia según el rol seleccionado. No modifiques los encabezados. Al subirla, ALVAPEN valida todas las filas antes de incorporarlas al lote; si una fila tiene errores, no se importa ninguna.';
    filePanel.append(roleText,fileActions,fileHelp,fileMsg);
    body.insertBefore(tabs,body.firstChild);body.insertBefore(pastePanel,tabs.nextSibling);body.insertBefore(filePanel,pastePanel.nextSibling);
    if(note)note.textContent='La carga masiva solo incorpora operaciones al lote. Después revisa las operaciones, valida RUC y estructura, y finalmente genera el TXT para Pago Masivo SUNAT.';

    function role(){return root.querySelector('input[name="det-role"]:checked')?.value||'adquiriente';}
    function headers(){
      return role()==='proveedor'
        ? ['Tipo documento adquiriente','Número documento adquiriente','Nombre/Razón social (solo código 040 y si no es RUC)','Tipo operación','Código bien/servicio','Monto depósito','Periodo AAAAMM','Tipo comprobante','Serie','Número']
        : ['RUC proveedor','Cuenta detracciones proveedor','Tipo operación','Código bien/servicio','Monto depósito','Periodo AAAAMM','Tipo comprobante','Serie','Número'];
    }
    function updateRole(){
      const r=role();
      roleText.textContent=r==='proveedor'
        ? 'Rol seleccionado: PROVEEDOR · 10 columnas. La cuenta del depositante se toma de “1. Datos del lote”.'
        : 'Rol seleccionado: ADQUIRIENTE · 9 columnas. Incluye la cuenta de detracciones de cada proveedor.';
      const items=Array.from(guide.children);
      items.forEach((el,i)=>{const show=r==='proveedor'?i>=2:i<2;el.style.display=show?'':'none';});
    }
    function switchPanel(name){
      tabs.querySelectorAll('button').forEach(b=>b.setAttribute('aria-selected',String(b.dataset.panel===name)));
      pastePanel.hidden=name!=='paste';filePanel.hidden=name!=='file';
    }
    tabs.addEventListener('click',e=>{const b=e.target.closest('button[data-panel]');if(b)switchPanel(b.dataset.panel);});
    root.querySelectorAll('input[name="det-role"]').forEach(el=>el.addEventListener('change',updateRole));
    updateRole();

    function xlsx(){
      if(window.XLSX)return Promise.resolve(window.XLSX);
      if(window.__alvapenXlsxPromise)return window.__alvapenXlsxPromise;
      window.__alvapenXlsxPromise=new Promise((resolve,reject)=>{
        const s=document.createElement('script');s.src='https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js';
        s.onload=()=>window.XLSX?resolve(window.XLSX):reject(new Error('No se pudo inicializar el lector Excel.'));
        s.onerror=()=>reject(new Error('No se pudo cargar el lector Excel. Verifica tu conexión.'));
        document.head.appendChild(s);
      });return window.__alvapenXlsxPromise;
    }
    function setStatus(kind,msg){fileMsg.className='det-validation '+kind;fileMsg.style.display='block';fileMsg.textContent=msg;}
    document.getElementById('det-excel-template').addEventListener('click',async()=>{
      try{
        const X=await xlsx(),h=headers(),provider=role()==='proveedor';
        const wb=X.utils.book_new();
        const ws=X.utils.aoa_to_sheet([h]);
        ws['!cols']=h.map(v=>({wch:Math.max(18,Math.min(48,v.length+5))}));
        X.utils.book_append_sheet(wb,ws,'Operaciones');
        const instructions=provider?[
          ['PLANTILLA DE DETRACCIONES – ROL PROVEEDOR'],
          ['Una fila por operación. No cambies los encabezados.'],
          ['Tipo documento adquiriente: 1 para DNI o 6 para RUC.'],
          ['Número documento adquiriente: número de documento del adquiriente.'],
          ['Nombre/Razón social: completar solo cuando corresponda al código 040 y el documento no sea RUC.'],
          ['Tipo operación: código permitido por el generador.'],
          ['Código bien/servicio: código SUNAT de tres dígitos.'],
          ['Monto depósito: importe positivo, sin S/ ni separador de miles.'],
          ['Periodo: AAAAMM. Tipo comprobante: 01 factura, 03 boleta o 09 guía, sujeto a las reglas del generador.'],
          ['Serie: cuatro caracteres. Número: hasta ocho dígitos.'],
          ['La cuenta de detracciones del proveedor se registra una sola vez en Datos del lote.']
        ]:[
          ['PLANTILLA DE DETRACCIONES – ROL ADQUIRIENTE'],
          ['Una fila por operación. No cambies los encabezados.'],
          ['RUC proveedor: RUC de 11 dígitos del proveedor.'],
          ['Cuenta detracciones proveedor: cuenta de 11 dígitos correspondiente a ese proveedor.'],
          ['Tipo operación: código permitido por el generador.'],
          ['Código bien/servicio: código SUNAT de tres dígitos.'],
          ['Monto depósito: importe positivo, sin S/ ni separador de miles.'],
          ['Periodo: AAAAMM. Tipo comprobante: 01 factura, 03 boleta o 09 guía, sujeto a las reglas del generador.'],
          ['Serie: cuatro caracteres. Número: hasta ocho dígitos.']
        ];
        const wi=X.utils.aoa_to_sheet(instructions);wi['!cols']=[{wch:115}];X.utils.book_append_sheet(wb,wi,'Instrucciones');
        X.writeFile(wb,provider?'Plantilla_Detracciones_Proveedor.xlsx':'Plantilla_Detracciones_Adquiriente.xlsx');
        setStatus('ok','Plantilla descargada para el rol '+(provider?'Proveedor':'Adquiriente')+'. Completa la hoja Operaciones y súbela aquí.');
      }catch(err){setStatus('err',err.message||String(err));}
    });
    document.getElementById('det-excel-file').addEventListener('change',async ev=>{
      const file=ev.target.files?.[0];if(!file)return;
      try{
        const X=await xlsx(),data=await file.arrayBuffer(),wb=X.read(data,{type:'array'});
        const ws=wb.Sheets['Operaciones']||wb.Sheets[wb.SheetNames[0]];
        if(!ws)throw new Error('No se encontró la hoja Operaciones.');
        const rows=X.utils.sheet_to_json(ws,{header:1,defval:'',raw:false});
        const expected=headers();
        if(!rows.length)throw new Error('La plantilla está vacía.');
        const incoming=rows[0].map(v=>String(v??'').trim());
        if(expected.some((h,i)=>incoming[i]!==h)||incoming.length!==expected.length)throw new Error('La plantilla no corresponde al rol actual o sus encabezados fueron modificados. Descarga la plantilla del rol '+(role()==='proveedor'?'Proveedor':'Adquiriente')+'.');
        const dataRows=rows.slice(1).filter(row=>row.some(v=>String(v??'').trim()!==''));
        if(!dataRows.length)throw new Error('No hay operaciones para importar.');
        const tsv=dataRows.map(row=>expected.map((_,i)=>String(row[i]??'').replace(/\t|\r?\n/g,' ')).join('\t')).join('\n');
        pasteInput.value=tsv;
        switchPanel('paste');
        if(typeof window.detImportPasted!=='function')throw new Error('No está disponible el validador de carga masiva.');
        window.detImportPasted();
        const msg=document.getElementById('det-paste-message');
        if(msg&&msg.classList.contains('ok'))setStatus('ok','Archivo procesado. '+msg.innerText.replace(/\s+/g,' ').trim());
        else setStatus('err','El archivo se leyó, pero la validación encontró observaciones. Revisa el detalle en “Pegar desde Excel”. No se incorporaron filas si hubo errores.');
      }catch(err){setStatus('err',err.message||String(err));}
      finally{ev.target.value='';}
    });
  }
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();