/* ================================================================
   Monitor WIN — lógica do dashboard
   (extraído de dashboard_win.html; servido pelo server em /app.js)
   ================================================================ */

/* ============ NÍVEIS: carregados do servidor (/niveis) ============ */
let NIV = null;               // preenchido via fetch
const API = location.protocol.startsWith("http") ? "" : "http://127.0.0.1:8001";
const WS_URL = (location.protocol === "https:" ? "wss://" : "ws://") +
               (location.host || "127.0.0.1:8001") + "/ws";

const S = { ultimo:null, abertura:null, maxima:null, minima:null,
            volume:null, agrC:null, agrV:null, vwap:null, fecAnt:null,
            planoAtiv:{compra:null, venda:null},  // hora REAL do servidor
            ticks:[], lastAlert:{}, soundOn:true, sim:null, ws:null };
const M = { sp500:null, dxy:null, dolar:null, di:null };   // último pacote macro

const $ = id => document.getElementById(id);
const fmt = v => v==null ? "—" : Math.round(v).toLocaleString("pt-BR");

async function loadNiveis(){
  try{
    const r = await fetch(API + "/niveis");
    NIV = await r.json();
  }catch(e){
    // fallback local caso o server esteja fora
    NIV = { contrato:"WIN", resistencias:[175110,175810,176510],
            suportes:[174410,173710,173010],
            alvos_compra:[175810,176510,177210],
            alvos_venda:[173710,173010,172310],
            zona_decisiva:{min:174400,max:175100},
            ladder:{min:172310,max:177210} };
  }
  $("contrato").textContent = `${NIV.contrato||"WIN"} · ${NIV.data_pregao||""} · ${NIV.fonte||""}`;
  LAD = { min: NIV.ladder.min, max: NIV.ladder.max };
  buildLadder(); buildPlano(); render();
  loadPerfil();
}

/* ============ PERFIL DE VOLUME (régua por acúmulo) ============ */
let PERFIL = null;
async function loadPerfil(){
  if(!NIV || !NIV.data_pregao) return;
  try{
    const r = await fetch(API + "/perfil/" + NIV.data_pregao);
    if(!r.ok) throw 0;
    const p = await r.json();
    PERFIL = (p.perfil && p.perfil.length) ? p : null;
  }catch(e){ PERFIL = null; }   // server antigo (sem /perfil) ou sem dados
  buildLadder();
}
setInterval(loadPerfil, 120000);   // perfil evolui ao longo do pregão

/* ============ RÉGUA ============ */
/* Range efetivo da régua: parte do ladder do JSON e SÓ expande quando o
   preço sai dele (nunca encolhe no dia) — os níveis não ficam dançando. */
let LAD = { min: 0, max: 1 };
function pctInLadder(p){
  return Math.min(100, Math.max(0, (p-LAD.min)/(LAD.max-LAD.min)*100));
}
function ensureLadderRange(p){
  if(p==null || !NIV) return;
  const margin = Math.max(200, (LAD.max-LAD.min)*0.03);
  let changed = false;
  if(p > LAD.max){ LAD.max = Math.ceil((p+margin)/5)*5; changed = true; }
  if(p < LAD.min){ LAD.min = Math.floor((p-margin)/5)*5; changed = true; }
  if(changed) buildLadder();
}
function buildLadder(){
  const el = $("ladder");
  el.innerHTML = '<div class="rail"></div>';
  const z = document.createElement("div");
  z.className = "zone";
  z.style.bottom = pctInLadder(NIV.zona_decisiva.min) + "%";
  z.style.height = (pctInLadder(NIV.zona_decisiva.max) - pctInLadder(NIV.zona_decisiva.min)) + "%";
  el.appendChild(z);
  // Faixas de calor: intensidade ∝ volume negociado em cada faixa de preço
  if(PERFIL){
    const maxPct = Math.max(...PERFIL.perfil.map(b=>b.pct));
    PERFIL.perfil.forEach(b=>{
      const lo = pctInLadder(b.preco - PERFIL.bucket/2);
      const hi = pctInLadder(b.preco + PERFIL.bucket/2);
      if(hi <= lo) return;
      const d = document.createElement("div");
      d.className = "vol-band";
      d.style.bottom = lo + "%";
      d.style.height = (hi-lo) + "%";
      d.style.opacity = (0.05 + 0.32 * b.pct / maxPct).toFixed(3);
      el.appendChild(d);
    });
  }
  const add = (price, cls, label) => {
    const d = document.createElement("div");
    d.className = "lvl " + cls;
    d.style.bottom = pctInLadder(price) + "%";
    d.innerHTML = `<b>${fmt(price)}</b><span class="tick"></span><span>${label}</span>`;
    el.appendChild(d);
  };
  NIV.resistencias.forEach((r,i)=>add(r,"res","R"+(i+1)));
  NIV.suportes.forEach((s,i)=>add(s,"sup","S"+(i+1)));
  const lastAC = NIV.alvos_compra[NIV.alvos_compra.length-1];
  const lastAV = NIV.alvos_venda[NIV.alvos_venda.length-1];
  if(!NIV.resistencias.includes(lastAC)) add(lastAC,"res","ALVO C");
  if(!NIV.suportes.includes(lastAV))     add(lastAV,"sup","ALVO V");
  if(PERFIL && PERFIL.poc != null) add(PERFIL.poc,"poc","POC");
  const m = document.createElement("div");
  m.className = "price-marker"; m.id = "marker"; m.style.bottom = "50%";
  m.innerHTML = `<div class="lineflag"></div><div class="tag" id="markerTag">—</div>`;
  el.appendChild(m);
}
function buildPlano(){
  if(!NIV) return;
  // Alvo atingido = máxima/mínima do dia tocou o nível (não usa o último,
  // que perderia toques rápidos). Invalidação derivada da zona decisiva.
  const hitC = a => S.maxima!=null && S.maxima >= a;
  const hitV = a => S.minima!=null && S.minima <= a;
  const alvos = (list, hit) =>
    list.map(a => hit(a) ? `<s>${fmt(a)}</s>✅` : fmt(a)).join(" / ");
  // Gatilho ativado: a hora vem do SERVIDOR (rompimento real dos snapshots,
  // sobrevive a F5 e a reinício); só mostra ✅ quando o servidor confirma.
  const gat = ts => ts
    ? ` <span class="hitflag">✅ ativado ${ts.slice(0,5)}</span>` : "";
  $("plano").innerHTML = `
    <div class="dist"><span class="up">Compra acima</span><b>${fmt(NIV.zona_decisiva.max)} + vol${gat(S.planoAtiv.compra)}</b></div>
    <div class="dist"><span class="up">Alvos</span><b>${alvos(NIV.alvos_compra, hitC)}</b></div>
    <div class="dist"><span class="up">Invalida</span><b>abaixo de ${fmt(NIV.zona_decisiva.min)}</b></div>
    <div class="dist"><span class="down">Venda abaixo</span><b>${fmt(NIV.zona_decisiva.min)} + vol${gat(S.planoAtiv.venda)}</b></div>
    <div class="dist"><span class="down">Alvos</span><b>${alvos(NIV.alvos_venda, hitV)}</b></div>
    <div class="dist"><span class="down">Invalida</span><b>acima de ${fmt(NIV.zona_decisiva.max)}</b></div>`;
}

/* ============ SOM ============ */
let audioCtx = null;
function beep(freq){
  if(!S.soundOn) return;
  try{
    audioCtx = audioCtx || new (window.AudioContext||window.webkitAudioContext)();
    const o = audioCtx.createOscillator(), g = audioCtx.createGain();
    o.frequency.value = freq;
    g.gain.setValueAtTime(.25, audioCtx.currentTime);
    g.gain.exponentialRampToValueAtTime(.001, audioCtx.currentTime + .5);
    o.connect(g).connect(audioCtx.destination);
    o.start(); o.stop(audioCtx.currentTime + .5);
  }catch(e){}
}

/* ============ ALERTAS ============ */
const COOLDOWN = 60000;
function pushAlert(txt, cls){
  const box = $("alerts");
  const d = document.createElement("div");
  d.className = "alert-item " + cls;
  d.innerHTML = `<time>${new Date().toLocaleTimeString("pt-BR")}</time>${txt}`;
  box.prepend(d);
  while(box.children.length > 40) box.lastChild.remove();
  $("priceCard").classList.remove("flash");
  void $("priceCard").offsetWidth;
  $("priceCard").classList.add("flash");
}
function checkCrossings(prev, cur){
  if(prev==null || !NIV) return;
  const now = Date.now();
  const fire = (lvl, dir, cls, freq) => {
    const key = dir + lvl;
    if(now - (S.lastAlert[key]||0) < COOLDOWN) return;
    S.lastAlert[key] = now;
    pushAlert(`${dir==="up"?"⬆ ROMPEU":"⬇ PERDEU"} ${fmt(lvl)}`, cls);
    beep(freq);
  };
  NIV.resistencias.concat([NIV.zona_decisiva.max]).forEach(l=>{
    if(prev < l && cur >= l) fire(l,"up","res",880);
  });
  NIV.suportes.concat([NIV.zona_decisiva.min]).forEach(l=>{
    if(prev > l && cur <= l) fire(l,"down","sup",440);
  });
}

/* ============ BIAS / DISTÂNCIAS / SPARK ============ */
function updateBias(){
  if(!NIV) return;
  const el = $("bias"), p = S.ultimo;
  if(p==null) return;
  const delta = (S.agrC!=null && S.agrV!=null) ? S.agrC - S.agrV : null;
  const dTxt = delta==null ? "" : ` · Δ ${delta>=0?"+":""}${fmt(delta)}`;
  if(p > NIV.zona_decisiva.max){
    el.className = "bias buy";
    el.textContent = `🐂 VIÉS COMPRADOR — acima de ${fmt(NIV.zona_decisiva.max)}${dTxt}`;
  }else if(p < NIV.zona_decisiva.min){
    el.className = "bias sell";
    el.textContent = `🐻 VIÉS VENDEDOR — abaixo de ${fmt(NIV.zona_decisiva.min)}${dTxt}`;
  }else{
    el.className = "bias neutral";
    el.textContent = `⚖ ZONA DECISIVA (${fmt(NIV.zona_decisiva.min)}–${fmt(NIV.zona_decisiva.max)}) — aguardar rompimento${dTxt}`;
  }
}
function updateDist(){
  if(!NIV) return;
  const p = S.ultimo; if(p==null) return;
  // Nível "atingido" = máxima/mínima do dia tocou (mesmo critério do plano):
  // riscado + ✅, como no monitor PETR4.
  const hitUp = l => S.maxima!=null && S.maxima >= l;
  const hitDn = l => S.minima!=null && S.minima <= l;
  const rows = [];
  // Regua de precos: R3 no topo ... S3 embaixo (zona decisiva no meio)
  for(let i = NIV.resistencias.length-1; i >= 0; i--)
    rows.push({label:`R${i+1} ${fmt(NIV.resistencias[i])}`, d:NIV.resistencias[i]-p, cls:"up", hit:hitUp(NIV.resistencias[i])});
  // ZD↑/ZD↓ são o gatilho do plano (não alvo) — nunca riscam aqui, só os
  // alvos R/S de fato atingidos ficam riscados na lista de distância.
  rows.push({label:`ZD↑ ${fmt(NIV.zona_decisiva.max)}`, d:NIV.zona_decisiva.max-p, cls:"amber", hit:false});
  rows.push({label:`ZD↓ ${fmt(NIV.zona_decisiva.min)}`, d:NIV.zona_decisiva.min-p, cls:"amber", hit:false});
  NIV.suportes.forEach((s,i)=>rows.push({label:`S${i+1} ${fmt(s)}`, d:s-p, cls:"down", hit:hitDn(s)}));
  let nearest = 0;
  rows.forEach((r,i)=>{ if(Math.abs(r.d) < Math.abs(rows[nearest].d)) nearest = i; });
  $("distList").innerHTML = rows.map((r,i)=>
    `<div class="dist${i===nearest?" near":""}"><span class="${r.cls}">${r.hit?`<s>${r.label}</s> <span class="hitflag">✅</span>`:r.label}</span><b>${r.d>=0?"+":""}${fmt(r.d)} pts</b></div>`
  ).join("");
}
function drawSpark(){
  const c = $("spark"), ctx = c.getContext("2d");
  ctx.clearRect(0,0,c.width,c.height);
  const t = S.ticks; if(t.length < 2) return;
  const min = Math.min(...t), max = Math.max(...t), pad = 12;
  const x = i => pad + i*(c.width-2*pad)/(t.length-1);
  const y = v => max===min ? c.height/2 : pad + (max-v)*(c.height-2*pad)/(max-min);
  if(NIV){
    [NIV.zona_decisiva.min, NIV.zona_decisiva.max].forEach(l=>{
      if(l>=min && l<=max){
        ctx.strokeStyle="rgba(240,180,41,.5)"; ctx.setLineDash([5,5]);
        ctx.beginPath(); ctx.moveTo(pad,y(l)); ctx.lineTo(c.width-pad,y(l)); ctx.stroke();
        ctx.setLineDash([]);
      }
    });
  }
  ctx.strokeStyle = t[t.length-1] >= t[0] ? "#2bd576" : "#ff4d5e";
  ctx.lineWidth = 2; ctx.beginPath();
  t.forEach((v,i)=> i ? ctx.lineTo(x(i),y(v)) : ctx.moveTo(x(i),y(v)));
  ctx.stroke();
}

/* ============ MACRO (S&P 500 · Dólar · DI) ============ */
const fmt2 = (v,d=2) => v==null ? "—" : v.toFixed(d).replace(".",",");
function renderMacro(){
  if(!M.sp500 && !M.dxy && !M.dolar && !M.di) return;
  // Impacto no IBOV: S&P 500↑ (risk-on global) favorece o índice;
  // Dólar↓ e Juros↓ favorecem. Faixa morta evita ruído virar seta.
  const imp = [];   // +1 favorável, -1 contrário, 0 neutro
  const rows = [];
  const seta = s => s>0 ? `<span class="seta up">▲</span>`
              : s<0 ? `<span class="seta down">▼</span>`
              : `<span class="seta flat">◆</span>`;
  if(M.sp500){
    const v = M.sp500.var_pct, s = Math.abs(v)<0.1 ? 0 : (v>0?1:-1);
    imp.push(s);
    rows.push(`<div class="macro-row"><span class="nome">🇺🇸 S&P 500<small>ES=F · Yahoo</small></span>
      <b class="${v>0?"up":v<0?"down":"flat"}">${fmt2(M.sp500.preco)} ${v>=0?"+":""}${fmt2(v)}%</b>${seta(s)}</div>`);
  }
  if(M.dxy){
    // DXY (dólar global): sobe → risk-off p/ emergentes → contra o IBOV.
    const v = M.dxy.var_pct, s = Math.abs(v)<0.1 ? 0 : (v<0?1:-1);
    imp.push(s);
    rows.push(`<div class="macro-row"><span class="nome">🌐 DXY<small>Dólar global · ICE</small></span>
      <b class="${v>0?"down":v<0?"up":"flat"}">${fmt2(M.dxy.preco)} ${v>=0?"+":""}${fmt2(v)}%</b>${seta(s)}</div>`);
  }
  if(M.dolar){
    const v = M.dolar.var_pct, s = Math.abs(v)<0.05 ? 0 : (v<0?1:-1);
    imp.push(s);
    rows.push(`<div class="macro-row"><span class="nome">💵 Dólar<small>${M.dolar.fonte||"USD/BRL · Yahoo"}</small></span>
      <b class="${v>0?"up":v<0?"down":"flat"}">${fmt2(M.dolar.preco,3)} ${v>=0?"+":""}${fmt2(v)}%</b>${seta(s)}</div>`);
  }
  if(M.di){
    const v = M.di.var_bps, s = v==null||Math.abs(v)<1 ? 0 : (v<0?1:-1);
    imp.push(s);
    const stale = M.di.desatualizado ? " ⏸" : "";
    rows.push(`<div class="macro-row"><span class="nome">📈 Juros DI<small>${M.di.ticker||"DI1"} · Profit RTD${stale}</small></span>
      <b class="${v>0?"down":v<0?"up":"flat"}">${fmt2(M.di.taxa,3)}% ${v==null?"":(v>=0?"+":"")+fmt2(v,1)+" bps"}</b>${seta(s)}</div>`);
  }
  $("macroRows").innerHTML = rows.join("");
  // Selo de alinhamento (estilo monitor PETR4): direção do vento macro
  // (maioria das setas) x direção do WIN (var vs FEC). Concordando,
  // "Alinhados COMPRA/VENDA"; discordando, "Divergentes".
  const fav = imp.filter(s=>s>0).length, con = imp.filter(s=>s<0).length;
  const base = S.fecAnt ?? S.abertura;
  const winVar = (S.ultimo!=null && base) ? (S.ultimo/base-1)*100 : null;
  const winTxt = winVar==null ? "" :
    ` · WIN ${winVar>=0?"+":""}${winVar.toFixed(2).replace(".",",")}%`;
  const macroDir = fav>con ? 1 : con>fav ? -1 : 0;
  const winDir = winVar==null ? 0 : winVar>0.05 ? 1 : winVar<-0.05 ? -1 : 0;
  let cls, txt;
  if(macroDir===0 || winDir===0){
    cls="misto"; txt=`◆ Vento macro ${macroDir>0?"a favor":macroDir<0?"contra":"misto"} — ${fav} a favor · ${con} contra`;
  }else if(macroDir===winDir){
    cls = macroDir>0 ? "fav" : "contra";
    txt = `✔ Alinhados ${macroDir>0?"COMPRA":"VENDA"} — macro ${fav>con?fav:con}/${imp.length} ${macroDir>0?"a favor":"contra"}`;
  }else{
    cls="misto";
    txt=`⚠ Divergentes — macro ${macroDir>0?"a favor":"contra"}, WIN na contramão`;
  }
  $("macroBanner").innerHTML = `<div class="align-banner ${cls}">${txt}${winTxt}</div>`;
}
async function loadMacroInicial(){
  try{
    const r = await fetch(API + "/macro");
    const d = await r.json();
    if(d.evento === "macro"){ M.sp500=d.sp500; M.dxy=d.dxy; M.dolar=d.dolar; M.di=d.di; renderMacro(); }
  }catch(e){}
}
async function loadPlanoAtivInicial(){
  try{
    const r = await fetch(API + "/plano_ativacao");
    const d = await r.json();
    S.planoAtiv = {compra:d.compra ?? null, venda:d.venda ?? null};
    if(NIV) buildPlano();
  }catch(e){}
}

/* ============ RENDER ============ */
function render(){
  const p = S.ultimo; if(p==null) return;
  $("px").textContent = fmt(p);
  if(S.abertura){
    const d = p - S.abertura, pct = (d/S.abertura*100).toFixed(2).replace(".",",");
    const el = $("pxVar");
    el.textContent = `${d>=0?"+":""}${fmt(d)} pts (${pct}%)`;
    el.className = "var " + (d>0?"up":d<0?"down":"flat");
  }
  $("stAbert").textContent = fmt(S.abertura);
  $("stMax").textContent = fmt(S.maxima);
  $("stMin").textContent = fmt(S.minima);
  $("stVol").textContent = fmt(S.volume);
  $("stAgrC").textContent = fmt(S.agrC);
  $("stAgrV").textContent = fmt(S.agrV);
  const delta = (S.agrC!=null&&S.agrV!=null)? S.agrC-S.agrV : null;
  const sd = $("stDelta");
  sd.textContent = delta==null?"—":(delta>=0?"+":"")+fmt(delta);
  sd.className = "v " + (delta>0?"up":delta<0?"down":"");
  $("stVwap").textContent = fmt(S.vwap);
  if(NIV && $("marker")){
    ensureLadderRange(p);
    $("marker").style.bottom = pctInLadder(p) + "%";
    $("markerTag").textContent = fmt(p);
  }
  $("lastUpdate").textContent = new Date().toLocaleTimeString("pt-BR");
  updateBias(); updateDist(); buildPlano(); drawSpark(); renderMacro();
  updateOpPnl();
}

/* ============ INGESTÃO ============ */
function onTick(data){
  if(data.evento === "niveis_atualizados"){ loadNiveis(); return; }
  if(data.evento === "operacoes_atualizadas"){ loadOps(); return; }
  if(data.evento === "blue_chips"){ renderBlueChips(data); return; }
  if(data.evento === "plano_ativacao"){
    S.planoAtiv = {compra:data.compra ?? null, venda:data.venda ?? null};
    if(NIV) buildPlano();
    return;
  }
  if(data.evento === "macro"){
    M.sp500 = data.sp500 ?? M.sp500;
    M.dxy   = data.dxy   ?? M.dxy;
    M.dolar = data.dolar ?? M.dolar;
    M.di    = data.di    ?? M.di;
    renderMacro(); return;
  }
  if(data.evento === "confluencia"){
    pushAlert("⭐ " + data.msg, "conf");
    beep(1320); setTimeout(()=>beep(1320), 250);   // beep duplo agudo
    return;
  }
  if(data.evento === "divergencia"){
    pushAlert("⚠ " + data.msg, "dive");
    beep(300);                                      // beep grave de aviso
    return;
  }
  const prev = S.ultimo;
  const g = k => (data[k]!==undefined && data[k]!==null) ? Number(data[k]) : null;
  S.ultimo   = g("ultimo")     ?? S.ultimo;
  S.abertura = g("abertura")   ?? S.abertura;
  S.maxima   = g("maxima")     ?? S.maxima;
  S.minima   = g("minima")     ?? S.minima;
  S.volume   = g("volume")     ?? S.volume;
  S.agrC     = g("agr_compra") ?? S.agrC;
  S.agrV     = g("agr_venda")  ?? S.agrV;
  S.vwap     = g("vwap")       ?? S.vwap;
  S.fecAnt   = g("fec_ant")    ?? S.fecAnt;
  if(S.ultimo!=null){
    S.ticks.push(S.ultimo);
    if(S.ticks.length > 300) S.ticks.shift();
    checkCrossings(prev, S.ultimo);
  }
  render();
}

/* ============ BLUE CHIPS ============ */
function renderBlueChips(data){
  const body = $("bcBody");
  body.innerHTML = data.ativos.map(a => {
    const cls = a.fluxo==="compra"?"up":a.fluxo==="venda"?"down":"flat";
    const varTxt = a.var_pct==null ? "—" :
      `${a.var_pct>=0?"+":""}${a.var_pct.toFixed(2).replace(".",",")}%`;
    const barW = Math.max(2, Math.min(40, a.dominancia));
    return `<tr>
      <td><b>${a.ticker}</b></td>
      <td class="num">${a.ultimo!=null ? a.ultimo.toFixed(2).replace(".",",") : "—"}</td>
      <td class="num ${cls}">${varTxt}</td>
      <td class="num" style="color:var(--dim)">${a.peso_ibov!=null ? a.peso_ibov.toFixed(1).replace(".",",")+"%" : "—"}</td>
      <td class="num ${cls}" style="white-space:nowrap">${a.fluxo}
        <span style="display:inline-block;width:${barW}px;height:6px;background:currentColor;opacity:.6;margin-left:4px;vertical-align:middle;border-radius:2px"></span>
      </td>
    </tr>`;
  }).join("");
  const total = data.ativos.length;
  const viesHtml = data.vies==="compra"
    ? `<span class="up">▲ VIÉS DE COMPRA</span>`
    : data.vies==="venda"
    ? `<span class="down">▼ VIÉS DE VENDA</span>`
    : `<span class="flat">◆ VIÉS MISTO</span>`;
  const n = data.vies==="compra" ? data.positivas : data.vies==="venda" ? data.negativas : Math.max(data.positivas,data.negativas);
  $("bcBias").innerHTML =
    `<span>Sinal agregado blue chips → WIN</span><b>${viesHtml} (${n}/${total})</b>`;
}
async function loadBlueChipsInicial(){
  try{
    const r = await fetch(API + "/blue_chips");
    const d = await r.json();
    if(d.ativos) renderBlueChips(d);
  }catch(e){}
}

/* ============ REGISTRO DO TRADER (fase 1 do aprendizado) ============ */
/* O trader clica ao ENTRAR (toda entrada, ganhe ou perca) e ao SAIR.
   Preco e contexto sao capturados pelo SERVIDOR no instante do clique —
   aqui so dispara o pedido e desenha o estado. */
let OPS = { aberta: null, operacoes: [] };

async function loadOps(){
  try{
    const r = await fetch(API + "/operacoes");
    const d = await r.json();
    OPS = { aberta: d.aberta || null, operacoes: d.operacoes || [] };
  }catch(e){ OPS = { aberta: null, operacoes: [] }; }
  renderOps();
}

async function opAbrir(tipo){
  if(S.sim){
    pushAlert("Simulador ativo — registro de operações desabilitado (dados seriam falsos)", "dive");
    return;
  }
  try{
    const r = await fetch(API + "/operacoes", {
      method: "POST", headers: {"Content-Type": "application/json"},
      body: JSON.stringify({tipo, motivo: $("opMotivo").value,
                            nota: ($("opNota") ? $("opNota").value : "")})
    });
    const d = await r.json();
    if(!r.ok){ pushAlert("⚠ " + (d.erro || "falha ao registrar"), "dive"); return; }
    pushAlert(`📝 ${tipo.toUpperCase()} registrada @ ${fmt(d.preco_entrada)} (${d.motivo})`,
              tipo === "compra" ? "res" : "sup");
  }catch(e){ pushAlert("⚠ servidor fora do ar — operação NÃO registrada", "dive"); }
  loadOps();
}

async function opFechar(){
  if(S.sim) return;
  try{
    const r = await fetch(API + "/operacoes/fechar", {
      method: "POST", headers: {"Content-Type": "application/json"},
      body: JSON.stringify({nota: ($("opNota") ? $("opNota").value : "")})
    });
    const d = await r.json();
    if(!r.ok){ pushAlert("⚠ " + (d.erro || "falha ao fechar"), "dive"); return; }
    const pts = d.resultado_pts;
    pushAlert(`📝 Fechou ${d.tipo} ${fmt(d.preco_entrada)} → ${fmt(d.preco_saida)} = ` +
              `${pts >= 0 ? "+" : ""}${fmt(pts)} pts`, pts >= 0 ? "res" : "sup");
  }catch(e){ pushAlert("⚠ servidor fora do ar — saída NÃO registrada", "dive"); }
  loadOps();
}

function renderOps(){
  const box = $("opControls");
  if(OPS.aberta){
    const o = OPS.aberta;
    box.innerHTML = `
      <div class="op-open ${o.tipo}">
        <span>${o.tipo === "compra" ? "▲ COMPRA" : "▼ VENDA"} @ ${fmt(o.preco_entrada)}
          <span class="flat">· ${o.ts.slice(0,5)} · ${o.motivo}</span></span>
        <span class="pnl" id="opPnl">—</span>
        <button class="btn" id="btnOpSair">✖ SAÍ</button>
      </div>
      <input class="op-nota" id="opNota" maxlength="500"
             placeholder="por que vai sair? alvo/stop/sentimento (opcional)">`;
    $("btnOpSair").onclick = opFechar;
    updateOpPnl();
  }else{
    box.innerHTML = `
      <div class="op-controls">
        <select id="opMotivo" title="Motivo da entrada (enriquece o dataset)">
          <option value="rompimento">rompimento</option>
          <option value="pullback">pullback</option>
          <option value="fluxo">fluxo</option>
          <option value="reversão">reversão</option>
          <option value="vwap">vwap</option>
          <option value="outro">outro</option>
        </select>
        <button class="btn buy" id="btnOpCompra">▲ COMPREI</button>
        <button class="btn sell" id="btnOpVenda">▼ VENDI</button>
      </div>
      <input class="op-nota" id="opNota" maxlength="500"
             placeholder="por que está entrando? (opcional, vale ouro)">`;
    $("btnOpCompra").onclick = () => opAbrir("compra");
    $("btnOpVenda").onclick  = () => opAbrir("venda");
  }
  const fechadas = OPS.operacoes.filter(o => o.resultado_pts != null);
  $("opList").innerHTML = fechadas.length
    ? fechadas.map(o => {
        const pts = o.resultado_pts;
        const cls = pts > 0 ? "up" : pts < 0 ? "down" : "flat";
        const notas = [o.nota_entrada && `entrada: ${o.nota_entrada}`,
                       o.nota_saida   && `saída: ${o.nota_saida}`]
                      .filter(Boolean).join("\n");
        const notaAttr = notas
          ? ` class="op-item tem-nota" title="${notas.replace(/"/g,"&quot;")}"`
          : ` class="op-item"`;
        return `<div${notaAttr}>
          <span class="flat">${o.ts.slice(0,5)}</span>
          <span class="${o.tipo === "compra" ? "up" : "down"}">${o.tipo === "compra" ? "▲" : "▼"} ${fmt(o.preco_entrada)}→${fmt(o.preco_saida)}</span>
          <span class="flat">${o.motivo}${notas ? " 🗒" : ""}</span>
          <span class="res ${cls}">${pts >= 0 ? "+" : ""}${fmt(pts)}</span>
        </div>`;
      }).join("")
    : `<div class="op-vazio">Nenhuma operação registrada hoje. Registre TODAS as entradas — as erradas ensinam tanto quanto as certas.</div>`;
}

function updateOpPnl(){
  /* P&L ao vivo da operacao aberta (informativo; o oficial e o do server) */
  const el = $("opPnl");
  if(!el || !OPS.aberta || S.ultimo == null) return;
  const sinal = OPS.aberta.tipo === "compra" ? 1 : -1;
  const pts = (S.ultimo - OPS.aberta.preco_entrada) * sinal;
  el.textContent = `${pts >= 0 ? "+" : ""}${fmt(pts)} pts`;
  el.className = "pnl " + (pts > 0 ? "up" : pts < 0 ? "down" : "flat");
}

/* ============ WEBSOCKET ============ */
function setConn(state, txt){
  $("connDot").className = "dot " + (state==="on"?"on":state==="sim"?"sim":"");
  $("connTxt").textContent = txt;
}
function connectWS(){
  try{
    S.ws = new WebSocket(WS_URL);
    S.ws.onopen = () => setConn("on","conectado (WS)");
    S.ws.onmessage = e => { try{ onTick(JSON.parse(e.data)); }catch(_){} };
    S.ws.onclose = () => { setConn("off","reconectando…"); setTimeout(connectWS, 3000); };
    S.ws.onerror = () => S.ws.close();
  }catch(e){ setTimeout(connectWS, 3000); }
}

/* ============ SIMULADOR ============ */
function toggleSim(){
  if(S.sim){ clearInterval(S.sim); S.sim=null; $("btnSim").textContent="▶ Simulador"; setConn("off","desconectado"); return; }
  let px = 174600, vol=0, ac=0, av=0;
  const abert = 174800;
  S.abertura = abert; S.maxima = px; S.minima = px;
  S.planoAtiv = {compra:null, venda:null};   // sem servidor no modo sim
  setConn("sim","SIMULADOR");
  $("btnSim").textContent = "⏹ Parar sim.";
  S.sim = setInterval(()=>{
    const drift = (Math.random()-.5)*60 + (Math.random()<.05 ? (Math.random()-.5)*250 : 0);
    px = Math.round((px + drift)/5)*5;
    vol += Math.round(Math.random()*80);
    if(Math.random() > .48) ac += Math.round(Math.random()*50);
    else av += Math.round(Math.random()*50);
    S.maxima = Math.max(S.maxima??px, px);
    S.minima = Math.min(S.minima??px, px);
    // No sim não há servidor: carimba a ativação localmente (hora atual).
    const hhmm = new Date().toTimeString().slice(0,5);
    if(NIV && S.maxima>=NIV.zona_decisiva.max && !S.planoAtiv.compra) S.planoAtiv.compra = hhmm;
    if(NIV && S.minima<=NIV.zona_decisiva.min && !S.planoAtiv.venda)  S.planoAtiv.venda  = hhmm;
    onTick({ultimo:px, abertura:abert, maxima:S.maxima, minima:S.minima,
            volume:vol, agr_compra:ac, agr_venda:av, vwap:abert + (ac-av)/10});
  }, 900);
}

$("btnSim").onclick = toggleSim;
$("btnSound").onclick = () => {
  S.soundOn = !S.soundOn;
  $("btnSound").textContent = "🔔 Som: " + (S.soundOn?"ON":"OFF");
};

loadNiveis();
loadBlueChipsInicial();
loadMacroInicial();
loadPlanoAtivInicial();
loadOps();
connectWS();
