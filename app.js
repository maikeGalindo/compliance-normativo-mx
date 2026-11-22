var KEY="cg_v4",S={data:[],depts:[],af:[],sel:-1,mat:"",mod:"dep",q:"",f:"ALL",na:false,ed:{},ro:false,on:false,upd:"",por:"",menu:false,page:0,pageSize:30};
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
function get(it){var e=S.ed[it.id]||{};return {est:e.est||it.est,resp:e.resp!=null?e.resp:it.resp,fecha:e.fecha!=null?e.fecha:it.fecha,enlace:e.enlace!=null?e.enlace:it.enlace,obs:e.obs!=null?e.obs:it.obs}}
function eff(it){return it.ap==="No"?"N/A":get(it).est}
function side(){var h='<div class="brand"><div class="logo">⚖️</div><div><b>Cumplimiento</b><small>Normativo MX</small></div></div><div class="lbl">Departamentos</div>';
  h+='<button class="dep'+(S.sel<0&&S.mod==="dep"?" on":"")+'" data-ac="sel" data-v="-1"><span>Todos</span><span class="n">'+S.data.length+'</span></button>';
  S.depts.forEach(function(d){var n=S.data.filter(function(i){return i.rl[d.i]==="Sí"}).length;h+='<button class="dep'+(S.sel===d.i&&S.mod==="dep"?" on":"")+'" data-ac="sel" data-v="'+d.i+'"><span>'+esc(d.n)+'</span><span class="n">'+n+'</span></button>'});
  h+='<div class="sub"><button class="dep'+(S.mat==="__a"?" on":"")+'" data-ac="mat" data-v="__a">Apoyo a otras áreas</button></div>';
  h+='<div class="lbl">Módulos</div><button class="dep'+(S.mod==="aud"?" on":"")+'" data-ac="mod" data-v="aud">Auditoría</button><button class="dep'+(S.mod==="fis"?" on":"")+'" data-ac="mod" data-v="fis">Fiscal</button>';
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
    list = S.af.filter(function(i){return (i.bl||"").toLowerCase().indexOf("aud")>-1 || (i.area||"").toLowerCase().indexOf("aud")>-1});
    if(!list.length) list = S.data.filter(function(i){return (i.area||"").toLowerCase().indexOf("aud")>-1 || (i.obl||"").toLowerCase().indexOf("auditor")>-1});
  } else if(S.mod==="fis"){
    list = S.af.filter(function(i){return (i.bl||"").toLowerCase().indexOf("fisc")>-1 || (i.area||"").toLowerCase().indexOf("fisc")>-1});
    if(!list.length) list = S.data.filter(function(i){return (i.area||"").toLowerCase().indexOf("fisc")>-1 || (i.obl||"").toLowerCase().indexOf("fiscal")>-1});
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
  head.innerHTML='<div><h1>Panel de cumplimiento normativo</h1><p>'+total+' obligaciones</p></div>';
  var pager=pages>1?'<div class="pager"><button data-ac="pg" data-v="'+(S.page-1)+'"'+(S.page<=0?' disabled':'')+'>←</button><span>'+(S.page+1)+' / '+pages+'</span><button data-ac="pg" data-v="'+(S.page+1)+'"'+(S.page>=pages-1?' disabled':'')+'>→</button></div>':'';
  w.innerHTML='<div class="row"><input class="search" id="q" placeholder="Buscar obligación..." value="'+esc(S.q)+'"><div class="seg"><button data-ac="f" data-v="ALL" class="'+(S.f==="ALL"?"on":"")+'">Todas</button><button data-ac="f" data-v="Pendiente" class="'+(S.f==="Pendiente"?"on":"")+'">Pendientes</button><button data-ac="f" data-v="Cumple" class="'+(S.f==="Cumple"?"on":"")+'">Cumple</button></div></div>'+pager+'<div class="grid">'+page.map(card).join("")+'</div>'+pager;
  var qi=$("q");if(qi){qi.oninput=function(){S.q=this.value;S.page=0;render()}}
}
function card(it){var g=get(it),s=eff(it);
  return '<div class="card" data-ac="detail" data-v="'+esc(it.id)+'"><div class="top"><div><span class="code">'+esc(it.id)+'</span><h3>'+esc(it.obl)+'</h3></div><span class="b">'+esc(s)+'</span></div><div class="fund"><small>Fundamento</small>'+esc(it.fund)+'</div><div class="meta"><span>'+(g.resp||"Sin responsable")+'</span><span>'+(g.fecha||"Sin fecha")+'</span></div></div>'}

function openDetail(id){
  var it = S.data.find(function(i){return i.id===id});
  if(!it) return;
  var g = get(it);
  var html = '<div class="mask" id="mask"><div class="modal"><div class="mh"><div><span class="code">'+esc(it.id)+'</span><h3>'+esc(it.obl)+'</h3></div><button class="close" data-ac="close">×</button></div><div class="mb">';
  html += '<div><h4>Fundamento</h4><p>'+esc(it.fund||"—")+'</p></div>';
  html += '<div><h4>Evidencia / Cómo materializar el cumplimiento</h4><p>'+esc(it.evid||"No especificada en el catálogo")+'</p></div>';
  html += '<div><h4>Periodicidad</h4><p>'+esc(it.per||"—")+'</p></div>';
  if(it.cond) html += '<div><h4>Condición / Notas</h4><p>'+esc(it.cond)+'</p></div>';
  html += '<div><h4>Estado actual</h4><p>'+esc(g.est||it.est)+'</p></div>';
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
});
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
