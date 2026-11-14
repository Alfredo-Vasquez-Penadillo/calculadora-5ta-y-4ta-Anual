(function(){
  const section = document.querySelector('#detraccionesMasivasView .det-operations-card');
  if (!section || document.getElementById('det-excel-tools')) return;
  const card = document.createElement('section');
  card.className = 'det-card';
  card.id = 'det-excel-tools';
  card.style.gridColumn = '1 / -1';
  card.innerHTML = '<div class="det-card-head"><h2>3. Carga masiva desde Excel</h2><p>Descarga la plantilla, pega tus operaciones y vuelve a cargar el archivo para incorporarlas al lote.</p></div>'+
    '<div class="det-card-body"><div class="det-excel-actions">'+
    '<button class="det-btn secondary" type="button" id="det-excel-template">⬇ Descargar plantilla Excel</button>'+
    '<label class="det-btn primary" for="det-excel-file" style="display:inline-flex;align-items:center;cursor:pointer">＋ Subir plantilla completada</label>'+
    '<input id="det-excel-file" type="file" accept=".xlsx,.xls,.csv" style="display:none">'+
    '</div><p class="det-note">Completa una fila por operación. No modifiques los encabezados. Los datos generales del lote (RUC, razón social, lote y cuenta del proveedor cuando corresponda) se completan previamente en “1. Datos del lote”.</p>'+
    '<div id="det-excel-status" class="det-validation" style="display:none"></div></div>';
  section.parentNode.insertBefore(card, section);
  const $ = id => document.getElementById(id);
  const headers = ['documento_contraparte','nombre_adquiriente','tipo_documento','cuenta_detracciones_proveedor','tipo_operacion','codigo_bien_servicio','monto_deposito','periodo_tributario','tipo_comprobante','serie_comprobante','numero_comprobante'];
  const example = ['20123456789','','6','00012345678','01','037',1500,'202611','01','F001','123'];
  function status(type, html){const e=$('det-excel-status');e.className='det-validation '+type;e.style.display='block';e.innerHTML=html;}
  function loadXLSX(){
    if(window.XLSX)return Promise.resolve(window.XLSX);
    if(window.__alvapenXlsxPromise)return window.__alvapenXlsxPromise;
    window.__alvapenXlsxPromise=new Promise((resolve,reject)=>{
      const s=document.createElement('script');s.src='https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js';
      s.onload=()=>window.XLSX?resolve(window.XLSX):reject(new Error('No se pudo inicializar el lector Excel.'));
      s.onerror=()=>reject(new Error('No se pudo cargar el componente Excel. Verifica tu conexión e inténtalo nuevamente.'));
      document.head.appendChild(s);
    });
    return window.__alvapenXlsxPromise;
  }
  $('det-excel-template').addEventListener('click',async()=>{
    try{
      const X=await loadXLSX();
      const wb=X.utils.book_new();
      const ws=X.utils.aoa_to_sheet([headers,example]);
      ws['!cols']=[{wch:23},{wch:32},{wch:17},{wch:30},{wch:18},{wch:22},{wch:18},{wch:20},{wch:18},{wch:18},{wch:22}];
      X.utils.book_append_sheet(wb,ws,'Operaciones');
      const info=X.utils.aoa_to_sheet([
        ['INSTRUCCIONES – ALVAPEN / DETRACCIONES'],
        ['Una fila representa una operación. No cambies ni elimines los encabezados.'],
        ['documento_contraparte: RUC del proveedor si el rol es Adquiriente; DNI o RUC del adquiriente si el rol es Proveedor.'],
        ['nombre_adquiriente: solo se utiliza para rol Proveedor, código 040 y documento distinto de RUC. En los demás casos dejar vacío.'],
        ['tipo_documento: 6 para RUC; 1 para DNI.'],
        ['cuenta_detracciones_proveedor: 11 dígitos. En rol Adquiriente, cuenta de cada proveedor; en rol Proveedor, puedes dejarla vacía porque se toma del encabezado del lote.'],
        ['tipo_operacion: códigos permitidos en la herramienta: 01, 02, 03, 04, 05 o 09.'],
        ['codigo_bien_servicio: código SUNAT de 3 dígitos, por ejemplo 037. Para operación 09 usar 099.'],
        ['monto_deposito: número positivo, sin símbolo de moneda.'],
        ['periodo_tributario: formato AAAAMM, por ejemplo 202611.'],
        ['tipo_comprobante: 01 factura; 03 boleta; 09 guía de remisión.'],
        ['serie_comprobante: 4 caracteres, por ejemplo F001, B001, E001 o G001.'],
        ['numero_comprobante: número de hasta 8 dígitos.'],
        ['Antes de cargar, selecciona el rol correcto y completa los datos generales del lote en ALVAPEN. La importación vuelve a validar cada fila con las reglas del generador.']
      ]);
      info['!cols']=[{wch:120}];X.utils.book_append_sheet(wb,info,'Instrucciones');
      X.writeFile(wb,'Plantilla_Detracciones_ALVAPEN.xlsx');
      status('ok','Plantilla Excel descargada. Completa la hoja “Operaciones” y luego súbela aquí.');
    }catch(e){status('err',e.message);}
  });
  $('det-excel-file').addEventListener('change',async ev=>{
    const file=ev.target.files&&ev.target.files[0];if(!file)return;
    try{
      const X=await loadXLSX();
      const data=await file.arrayBuffer();
      const wb=X.read(data,{type:'array',cellDates:false});
      const ws=wb.Sheets['Operaciones']||wb.Sheets[wb.SheetNames[0]];
      if(!ws)throw new Error('El archivo no contiene una hoja de operaciones.');
      const rows=X.utils.sheet_to_json(ws,{defval:'',raw:false});
      if(!rows.length)throw new Error('La plantilla no contiene filas para importar.');
      const missing=headers.filter(h=>!Object.prototype.hasOwnProperty.call(rows[0],h));
      if(missing.length)throw new Error('No se reconocen los encabezados: '+missing.join(', ')+'. Descarga la plantilla oficial y no modifiques sus encabezados.');
      let added=0;const errors=[];
      rows.forEach((r,i)=>{
        const line=i+2;
        if(headers.every(h=>String(r[h]??'').trim()===''))return;
        const role=document.querySelector('input[name="det-role"]:checked')?.value||'adquiriente';
        const set=(id,v)=>{const el=$(id);if(el)el.value=String(v??'').trim();};
        set('det-counterparty-doc',r.documento_contraparte);
        set('det-counterparty-name',r.nombre_adquiriente);
        set('det-doc-type',r.tipo_documento||'6');
        if(role==='adquiriente')set('det-account',r.cuenta_detracciones_proveedor);
        set('det-operation',String(r.tipo_operacion||'').padStart(2,'0'));
        set('det-service',String(r.codigo_bien_servicio||'').padStart(3,'0'));
        set('det-amount',String(r.monto_deposito).replace(',','.'));
        set('det-period',String(r.periodo_tributario).replace(/\.0$/,''));
        set('det-comp-type',String(r.tipo_comprobante||'').padStart(2,'0'));
        set('det-series',r.serie_comprobante);
        set('det-number',String(r.numero_comprobante).replace(/\.0$/,'').padStart(8,'0'));
        if(typeof window.detUpdateCompTypes==='function')window.detUpdateCompTypes();
        const before=Number($('det-count')?.textContent||0);
        if(typeof window.detAddOperation!=='function')throw new Error('No se encontró el generador de operaciones.');
        window.detAddOperation();
        const after=Number($('det-count')?.textContent||0);
        if(after>before)added++;else{
          const msg=$('det-form-message')?.innerText||'No se pudo validar la operación.';
          errors.push('Fila '+line+': '+msg.replace(/\s+/g,' ').trim());
        }
      });
      const safe=errors.map(x=>'<li>'+x.replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]))+'</li>').join('');
      if(errors.length){
        status('err','<strong>Importación terminada: '+added+' operación(es) agregada(s), '+errors.length+' fila(s) observada(s).</strong><ul>'+safe+'</ul><button type="button" class="det-btn secondary" id="det-excel-errors" style="margin-top:8px">Descargar reporte de errores</button>');
        $('det-excel-errors').addEventListener('click',async()=>{
          const X=await loadXLSX();const out=X.utils.aoa_to_sheet([['Observación'],...errors.map(x=>[x])]);const b=X.utils.book_new();X.utils.book_append_sheet(b,out,'Errores');X.writeFile(b,'Errores_Importacion_Detracciones.xlsx');
        });
      }else status('ok','<strong>Importación correcta.</strong> Se agregaron '+added+' operación(es). Revisa el lote y ejecuta “Validar estructura” antes de generar el TXT.');
    }catch(e){status('err',String(e.message||e));}
    finally{ev.target.value='';}
  });
})();