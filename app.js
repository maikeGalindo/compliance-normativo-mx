var KEY="cg_v4",S={data:[],depts:[],af:[],sel:-1,mat:"",mod:"dep",q:"",f:"ALL",na:false,ed:{},ro:false,on:false,upd:"",por:"",menu:false,page:0,pageSize:30,auditSel:{}};
var ST=["Pendiente","Parcial","Cumple","No cumple","N/A"];
function $(i){return document.getElementById(i)}
function esc(t){return String(t==null?"":t).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;")}
function load(){try{var o=JSON.parse(localStorage.getItem(KEY)||"{}");["data","depts","af","ed","upd"].forEach(function(k){if(o[k])S[k]=o[k]});S.ro=!!o.ro;S.on=!!o.on;S.por=localStorage.getItem("cg_por")||""}catch(e){}}
function save(){try{localStorage.setItem(KEY,JSON.stringify({data:S.data,depts:S.depts,af:S.af,ed:S.ed,upd:S.upd,ro:S.ro,on:S.on}))}catch(e){}}
function toast(m){var d=document.createElement("div");d.className="toast";d.textContent=m;document.body.appendChild(d);setTimeout(function(){d.remove()},2800)}
function loadXLSX(cb){if(window.XLSX)return cb();var s=document.createElement("script");s.src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";s.onload=cb;document.head.appendChild(s)}
function parse(buf){loadXLSX(function(){var wb=XLSX.read(buf,{type:"array",cellDates:true}),W=wb.Sheets;
  if(!W.CATALOGO||!W.CONTROL||!W.CONFIG){toast("El archivo no parece la guía (faltan hojas)");return}
  var rd=function(n,r0){return W[n]?XLSX.utils.sheet_to_json(W[n],{header:1,defval:"",range:r0}):[]};
  S.ro=false;S.on=false;S.upd="";S.ed={};build(rd("CATALOGO",4),rd("CONTROL",4),rd("CONFIG",2),rd("AUD_FISCAL",4));S.sel=-1;S.mat="";S.mod="dep";S.page=0;save();render();toast(S.data.length+" obligaciones cargadas")})}
function build(cat,ctl,cfg,af){var dp=[],out=[],k;
  for(k=0;k<9;k++){var nm=cfg[k]&&cfg[k][10];if(nm)dp.push({i:k,n:String(nm)})}
  cat.forEach(function(r,i){if(!r[0])return;var c=ctl[i]||[],rl={};for(k=0;k<9;k++){var v=r[12+k];if(v==="Sí"||v==="Apoyo")rl[k]=v}
    out.push({id:String(r[0]),area:r[1],obl:r[2],fund:r[3],evid:r[4],per:r[5],cond:r[7],car:r[8]||"Legal",rl:rl,ap:c[4]==="No"?"No":"Sí",est:ST.indexOf(c[6])>-1?c[6]:"Pendiente",resp:c[5]||"",fecha:(c[8]||"").toString().slice(0,10),enlace:c[12]||"",obs:c[13]||""})});
  S.data=out;S.depts=dp;
  S.af=(af||[]).filter(function(r){return r[0]}).map(function(r){return {id:String(r[0]),bl:r[1],obl:r[2],pide:r[3],fund:r[4],per:r[5],area:r[6],fmt:r[7],est:r[8]||"Pendiente",resp:r[9]||"",fecha:(r[10]||"").toString().slice(0,10),enlace:r[11]||"",obs:r[12]||"",ap:"Sí",af:1}});
  if(!dp.some(function(d){return d.i===S.sel}))S.sel=-1}

function getStatus(it){
  var id = it.id || "";
  if(id.indexOf("NOM035")===0) return "nuevo";
  // Mark some 2026 updates as modificado
  if(id.indexOf("RH-00")===0 || id.indexOf("PLD-")===0 || id.indexOf("FIS-")===0) return "mod";
  // Check if marked obsolete in notes or by user
  if(S.ed[id] && S.ed[id].obsolete) return "obs";
  return "";
}
function statusBadge(st){
  if(st==="nuevo") return '<span class="badge nuevo">Nuevo</span>';
  if(st==="mod") return '<span class="badge mod">Modificado</span>';
  if(st==="obs") return '<span class="badge obs">Obsoleto</span>';
  return "";
}

function get(it){var e=S.ed[it.id]||{};return {est:e.est||it.est,resp:e.resp!=null?e.resp:it.resp,fecha:e.fecha!=null?e.fecha:it.fecha,enlace:e.enlace!=null?e.enlace:it.enlace,obs:e.obs!=null?e.obs:it.obs,validated:!!e.validated}}
function eff(it){return it.ap==="No"?"N/A":get(it).est}
function side(){var h='<div class="brand"><div class="logo">⚖️</div><div><b>Cumplimiento</b><small>Normativo MX</small></div></div><div class="lbl">Departamentos</div>';
  h+='<button class="dep'+(S.sel<0&&S.mod==="dep"?" on":"")+'" data-ac="sel" data-v="-1"><span>Todos</span><span class="n">'+S.data.length+'</span></button>';
  S.depts.forEach(function(d){var n=S.data.filter(function(i){return i.rl[d.i]==="Sí"}).length;h+='<button class="dep'+(S.sel===d.i&&S.mod==="dep"?" on":"")+'" data-ac="sel" data-v="'+d.i+'"><span>'+esc(d.n)+'</span><span class="n">'+n+'</span></button>'});
  h+='<div class="sub"><button class="dep'+(S.mat==="__a"?" on":"")+'" data-ac="mat" data-v="__a">Apoyo a otras áreas</button></div>';
  var hasNew = S.data.some(function(i){return getStatus(i)==="nuevo"||getStatus(i)==="mod"});
  h+='<div class="lbl">Módulos</div><button class="dep'+(S.mod==="aud"?" on":"")+'" data-ac="mod" data-v="aud">Auditoría'+(hasNew?' <span class="mod-badge">nuevo</span>':'')+'</button><button class="dep'+(S.mod==="fis"?" on":"")+'" data-ac="mod" data-v="fis">Fiscal'+(hasNew?' <span class="mod-badge">nuevo</span>':'')+'</button>';
  h+='<div class="gear"><button class="gb" data-ac="menu">⚙️</button><span class="rol">'+(S.ro?"Solo lectura":"Edición")+'</span></div>';
  if(S.menu)h+='<div class="menu"><button data-ac="reload">Actualizar catálogo</button><button data-ac="ld">Cargar Excel</button><button data-ac="theme">Tema</button><button data-ac="ro">'+(S.ro?"Habilitar edición":"Solo lectura")+'</button></div>';
  $("side").innerHTML=h}
function render(){
  var init=document.getElementById("initial");if(init)init.remove();
  side();var head=$("head"),w=$("wrap");
  if(!S.data.length){head.innerHTML='<div><h1>Panel de cumplimiento normativo</h1><p>Consulta lo que aplica a cada departamento y da seguimiento a la evidencia.</p></div>';
    w.innerHTML='<div class="empty"><h2>Cargando catálogo...</h2><p>Si no carga en unos segundos, toca el botón.</p><button class="btn p" data-ac="reload">Actualizar catálogo</button><br><button class="btn" data-ac="ld">Cargar Excel manual</button></div>';return}
  var list;
  if(S.mod==="aud"){
    // Auditoría interna: show all obligations with selection capability
    list = S.data.slice();
  } else if(S.mod==="fis"){
    // Fiscal: obligations typically reviewed by authorities/firms
    var fiscalKeywords = ["fiscal","impuesto","isr","iva","sat","declaracion","cfdi","comprobante","contable","dictamen","retencion","nomina","imss","infonavit","isn","ieps","contabilidad electronica","balanza","catalogo de cuentas","diot","pagos provisionales","enteros","retenciones","ptu","aguinaldo","prima vacacional","subsidio","asimilados","regimen","rfc","curp","e.firma","opinion de cumplimiento","buro","32-d","art 69","art 69-b","efos","edos","devolucion","compensacion","acreditamiento"];
    list = S.data.filter(function(i){
      var t = (i.obl+" "+i.fund+" "+i.area).toLowerCase();
      return fiscalKeywords.some(function(k){return t.indexOf(k)>-1});
    });
  } else if(S.mat==="__a"){
    // Apoyo a otras áreas: show items marked as Apoyo for the selected dept (or any if Todos)
    list = S.data.filter(function(i){
      if(S.sel<0) return Object.values(i.rl).some(function(v){return v==="Apoyo"});
      return i.rl[S.sel]==="Apoyo";
    });
  } else {
    list = S.data.filter(function(i){var r=S.sel<0?"Sí":i.rl[S.sel];return r==="Sí"});
  }
  if(S.f!=="ALL")list=list.filter(function(i){return eff(i)===S.f});
  if(S.q){var q=S.q.toLowerCase();list=list.filter(function(i){return (i.obl+" "+i.id+" "+i.fund).toLowerCase().indexOf(q)>-1})}
  var total=list.length,start=S.page*S.pageSize,page=list.slice(start,start+S.pageSize),pages=Math.ceil(total/S.pageSize)||1;
  var validated=S.data.filter(function(i){return get(i).validated}).length;
  var valPct=S.data.length?Math.round(validated/S.data.length*100):0;
  var valClass=valPct>=80?"ok":(valPct>=40?"":"bad");
  head.innerHTML='<div style="display:flex;align-items:center;gap:16px;flex-wrap:wrap"><div><h1>Panel de cumplimiento normativo</h1><p>'+total+' obligaciones</p></div><div class="val-summary"><span class="val-dot '+valClass+'"></span><span>'+validated+' / '+S.data.length+' validadas ('+valPct+'%)</span></div></div>';
  var pager=pages>1?'<div class="pager"><button data-ac="pg" data-v="'+(S.page-1)+'"'+(S.page<=0?' disabled':'')+'>←</button><span>'+(S.page+1)+' / '+pages+'</span><button data-ac="pg" data-v="'+(S.page+1)+'"'+(S.page>=pages-1?' disabled':'')+'>→</button></div>':'';
  var auditBar="";
  if(S.mod==="aud"){
    var selCount=Object.keys(S.auditSel||{}).length;
    auditBar='<div class="audit-bar"><span style="font-size:13px;font-weight:600">Auditoría interna</span><button data-ac="audselall">Seleccionar visibles</button><button data-ac="audclear">Limpiar</button><button class="p" data-ac="audreport"'+(selCount?" ":" disabled")+'>Generar reporte ('+selCount+')</button></div>';
  }
  w.innerHTML=auditBar+'<div class="row"><input class="search" id="q" placeholder="Buscar obligación..." value="'+esc(S.q)+'"><div class="seg"><button data-ac="f" data-v="ALL" class="'+(S.f==="ALL"?"on":"")+'">Todas</button><button data-ac="f" data-v="Pendiente" class="'+(S.f==="Pendiente"?"on":"")+'">Pendientes</button><button data-ac="f" data-v="Cumple" class="'+(S.f==="Cumple"?"on":"")+'">Cumple</button></div></div>'+pager+'<div class="grid">'+page.map(card).join("")+'</div>'+pager;
  var qi=$("q");if(qi){qi.oninput=function(){S.q=this.value;S.page=0;render()}}
}
function card(it){var g=get(it),s=eff(it);
  var audChk = (S.mod==="aud") ? '<label class="audit-sel" onclick="event.stopPropagation()"><input type="checkbox" data-ac="audsel" data-id="'+esc(it.id)+'"'+((S.auditSel||{})[it.id]?" checked":"")+'> Incluir</label>' : "";
  var valMark = g.validated ? '<span class="val-dot ok" title="Validada" style="margin-left:6px"></span>' : "";
  var st=getStatus(it);
  return '<div class="card" data-ac="detail" data-v="'+esc(it.id)+'"><div class="top"><div><span class="code">'+esc(it.id)+'</span><h3>'+esc(it.obl)+statusBadge(st)+valMark+'</h3></div><span class="b">'+esc(s)+'</span></div><div class="fund"><small>Fundamento</small>'+esc(it.fund)+'</div><div class="meta"><span>'+(g.resp||"Sin responsable")+'</span><span>'+(g.fecha||"Sin fecha")+'</span>'+audChk+'</div></div>'}

function openDetail(id){
  var it = S.data.find(function(i){return i.id===id});
  if(!it) return;
  var g = get(it);
  var html = '<div class="mask" id="mask"><div class="modal"><div class="mh"><div><span class="code">'+esc(it.id)+'</span><h3>'+esc(it.obl)+'</h3></div><button class="close" data-ac="close">×</button></div><div class="mb">';
  html += '<div><h4>Fundamento</h4><p>'+esc(it.fund||"—")+'</p></div>';
  html += '<div><h4>Evidencia / Cómo materializar el cumplimiento</h4><p>'+esc(it.evid||"No especificada en el catálogo")+'</p></div>';
  html += '<label class="val-box'+(g.validated?" checked":"")+'"><input type="checkbox" data-ac="val" data-id="'+esc(it.id)+'"'+(g.validated?" checked":"")+'> <span>'+(g.validated?"Evidencia validada — Cumple":"Marcar como validada (cumple con la evidencia)")+'</span></label>';
  html += '<div><h4>Periodicidad</h4><p>'+esc(it.per||"—")+'</p></div>';
  if(it.cond) html += '<div><h4>Condición / Notas</h4><p>'+esc(it.cond)+'</p></div>';
  html += '<div><h4>Estado actual</h4><p>'+esc(g.est||it.est)+'</p></div>';
  html += '<div><button class="btn" data-ac="toggleobs" data-id="'+esc(it.id)+'" style="font-size:12px">'+(getStatus(it)==="obs"?"Quitar obsoleto":"Marcar como obsoleto")+'</button></div>';
  html += '<div><h4>Responsable</h4><p>'+esc(g.resp||"Sin asignar")+'</p></div>';
  html += '<div><h4>Fecha límite</h4><p>'+esc(g.fecha||"Sin fecha")+'</p></div>';
  if(g.enlace) html += '<div><h4>Enlace</h4><p><a href="'+esc(g.enlace)+'" target="_blank">'+esc(g.enlace)+'</a></p></div>';
  if(g.obs) html += '<div><h4>Observaciones</h4><p>'+esc(g.obs)+'</p></div>';
  html += '</div></div></div>';
  document.body.insertAdjacentHTML("beforeend", html);
}

document.addEventListener("click",function(e){
  var b=e.target.closest("[data-ac]");if(!b)return;
  var ac=b.dataset.ac,v=b.dataset.v;
  if(ac==="ld"){$("file").click()}
  else if(ac==="detail"){openDetail(v)}
  else if(ac==="val"){var id=b.dataset.id;S.ed[id]=S.ed[id]||{};S.ed[id].validated=b.checked;save();render();var m=document.getElementById("mask");if(m)m.remove();openDetail(id)}
  else if(ac==="toggleobs"){var id=b.dataset.id;S.ed[id]=S.ed[id]||{};S.ed[id].obsolete=!S.ed[id].obsolete;save();render();var m=document.getElementById("mask");if(m)m.remove();openDetail(id)}
  else if(ac==="close"){var m=document.getElementById("mask");if(m)m.remove()}
  else if(ac==="reload"){S.data=[];localStorage.removeItem(KEY);S.page=0;autoLoad();toast("Actualizando...")}
  else if(ac==="sel"){S.sel=+v;S.mat="";S.mod="dep";S.page=0;render()}
  else if(ac==="mat"){S.mat=v;S.page=0;render()}
  else if(ac==="mod"){S.mod=v;S.mat="";S.page=0;render()}
  else if(ac==="menu"){S.menu=!S.menu;render()}
  else if(ac==="f"){S.f=v;S.page=0;render()}
  else if(ac==="pg"){S.page=+v;render()}
  else if(ac==="theme"){var r=document.documentElement;r.dataset.theme=r.dataset.theme==="dark"?"":"dark"}
  else if(ac==="ro"){S.ro=!S.ro;save();render()}
  else if(ac==="audsel"){var id=b.dataset.id;S.auditSel=S.auditSel||{};if(b.checked)S.auditSel[id]=1;else delete S.auditSel[id];save();render()}
  else if(ac==="audselall"){S.auditSel=S.auditSel||{};var list=S.data;if(S.q){var q=S.q.toLowerCase();list=list.filter(function(i){return (i.obl+" "+i.id).toLowerCase().indexOf(q)>-1})}list.slice(S.page*S.pageSize,(S.page+1)*S.pageSize).forEach(function(i){S.auditSel[i.id]=1});save();render()}
  else if(ac==="audclear"){S.auditSel={};save();render()}
  else if(ac==="audsave"){toast("Auditoría guardada");var m=document.getElementById("mask");if(m)m.remove()}
  else if(ac==="audreport"){generateAuditReport()}
});

function generateAuditReport(){
  var ids=Object.keys(S.auditSel||{});
  if(!ids.length){toast("Selecciona obligaciones para auditar");return}
  var items=S.data.filter(function(i){return S.auditSel[i.id]});
  S.currentAudit = S.currentAudit || {};
  var html='<div class="mask" id="mask"><div class="modal" style="max-width:860px"><div class="mh"><div><h3>Auditoría Interna</h3><p style="margin:4px 0 0;font-size:12px;color:var(--muted)">'+items.length+' obligaciones — '+new Date().toLocaleDateString("es-MX")+'</p></div><button class="close" data-ac="close">×</button></div><div class="mb" id="auditBody">';
  items.forEach(function(it){
    var g=get(it);
    var a = S.currentAudit[it.id] || {};
    html+='<div class="audit-item" data-id="'+esc(it.id)+'"><h4>'+esc(it.id)+' — '+esc(it.obl)+'</h4>';
    html+='<div style="font-size:12px;color:var(--muted);margin-bottom:6px">Periodicidad: '+esc(it.per||"—")+' · Evidencia requerida: '+esc(it.evid||"—")+'</div>';
    html+='<div class="audit-fields">';
    html+='<div><label>Hallazgo</label><textarea data-field="hallazgo" data-id="'+esc(it.id)+'">'+esc(a.hallazgo||"")+'</textarea></div>';
    html+='<div><label>Evidencia entregada</label><input type="text" data-field="evidencia" data-id="'+esc(it.id)+'" value="'+esc(a.evidencia||"")+'" placeholder="Descripción o nombre de archivo"></div>';
    html+='<div><label>Responsable</label><input type="text" data-field="resp" data-id="'+esc(it.id)+'" value="'+esc(a.resp||g.resp||"")+'"></div>';
    html+='<div><label>Fecha límite</label><input type="date" data-field="fecha" data-id="'+esc(it.id)+'" value="'+esc(a.fecha||g.fecha||"")+'"></div>';
    html+='<div><label>Estado</label><select data-field="estado" data-id="'+esc(it.id)+'"><option'+(a.estado==="Pendiente"?" selected":"")+'>Pendiente</option><option'+(a.estado==="Cumple"?" selected":"")+'>Cumple</option><option'+(a.estado==="No cumple"?" selected":"")+'>No cumple</option><option'+(a.estado==="Parcial"?" selected":"")+'>Parcial</option></select></div>';
    html+='<div><label>Archivo (nombre)</label><input type="file" data-field="file" data-id="'+esc(it.id)+'" style="font-size:11px"></div>';
    html+='</div></div>';
  });
  html+='</div><div style="padding:12px 18px;border-top:1px solid var(--line);display:flex;justify-content:space-between;gap:8px"><button class="btn" data-ac="audsave">Guardar auditoría</button><button class="btn p" onclick="window.print()">Imprimir / PDF</button></div></div></div>';
  document.body.insertAdjacentHTML("beforeend", html);
  // Bind field changes
  document.querySelectorAll("#auditBody [data-field]").forEach(function(el){
    el.addEventListener("change", function(){
      var id=this.dataset.id, field=this.dataset.field;
      S.currentAudit[id]=S.currentAudit[id]||{};
      if(field==="file"){S.currentAudit[id].fileName=this.files[0]?this.files[0].name:""}
      else S.currentAudit[id][field]=this.value;
    });
  });
}

$("file").addEventListener("change",function(){var f=this.files[0];if(!f)return;var r=new FileReader();r.onload=function(){parse(new Uint8Array(r.result))};r.readAsArrayBuffer(f);this.value=""});
var CATALOGO_URL="https://raw.githubusercontent.com/maikeGalindo/compliance-normativo-mx/main/catalogo.json?t="+Date.now();
function autoLoad(){
  if(S.data.length)return;
  fetch(CATALOGO_URL)
    .then(function(r){if(!r.ok)throw new Error("HTTP "+r.status);return r.json()})
    .then(function(json){
      S.ro=false;S.on=false;S.upd="";S.ed={};
      build(json.catalogo.slice(4),json.control.slice(4),json.config.slice(2),(json.aud_fiscal||[]).slice(4));
      S.sel=-1;S.mat="";S.mod="dep";S.page=0;save();render();
      toast("Catálogo cargado ("+S.data.length+")");
    })
    .catch(function(e){
      var w=$("wrap");
      if(w)w.innerHTML='<div class="empty"><h2>No se pudo cargar</h2><p>'+esc(String(e))+'</p><button class="btn p" data-ac="reload">Reintentar</button></div>';
    });
}
load();render();autoLoad();
