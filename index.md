<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>BetTrack Pro</title>
<link href="https://fonts.googleapis.com/css2?family=Space+Mono:wght@400;700&family=Syne:wght@400;600;700;800&display=swap" rel="stylesheet">
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
*{margin:0;padding:0;box-sizing:border-box}
:root{
  --bg:#0a0a0f;--surface:#111118;--surface2:#1a1a24;--surface3:#22222f;
  --border:#2a2a3a;--accent:#7c6dfa;--accent2:#fa6d7c;
  --green:#4dfa9a;--red:#fa4d6d;--gold:#fac84d;--blue:#4db8fa;
  --text:#f0f0ff;--muted:#6b6b88;
  --font-d:'Syne',sans-serif;--font-m:'Space Mono',monospace;
}
body{background:var(--bg);color:var(--text);font-family:var(--font-d);min-height:100vh;overflow-x:hidden}
body::before{content:'';position:fixed;inset:0;background-image:repeating-linear-gradient(0deg,transparent,transparent 40px,rgba(124,109,250,0.025) 40px,rgba(124,109,250,0.025) 41px),repeating-linear-gradient(90deg,transparent,transparent 40px,rgba(124,109,250,0.025) 40px,rgba(124,109,250,0.025) 41px);pointer-events:none;z-index:0}
.orb{position:fixed;border-radius:50%;filter:blur(90px);pointer-events:none;z-index:0}
.orb1{width:500px;height:500px;background:rgba(124,109,250,0.1);top:-150px;right:-150px}
.orb2{width:350px;height:350px;background:rgba(250,109,124,0.07);bottom:50px;left:-100px}

/* NAV */
nav{position:sticky;top:0;z-index:200;background:rgba(10,10,15,0.9);backdrop-filter:blur(20px);border-bottom:1px solid var(--border);padding:0 1.5rem;display:flex;align-items:center;justify-content:space-between;height:56px}
.logo{font-size:1.1rem;font-weight:800;letter-spacing:-0.02em}
.logo span{color:var(--accent)}
.nav-tabs{display:flex;gap:0.25rem}
.nav-tab{background:none;border:none;color:var(--muted);font-family:var(--font-d);font-size:0.78rem;font-weight:600;padding:0.4rem 0.9rem;border-radius:6px;cursor:pointer;transition:all 0.2s;letter-spacing:0.02em}
.nav-tab:hover{color:var(--text);background:var(--surface2)}
.nav-tab.active{color:var(--text);background:var(--surface2);border:1px solid var(--border)}
.nav-right{display:flex;align-items:center;gap:0.75rem}
.nav-badge{font-family:var(--font-m);font-size:0.72rem;padding:0.25rem 0.6rem;border-radius:20px;font-weight:700}
.badge-green{background:rgba(77,250,154,0.15);color:var(--green);border:1px solid rgba(77,250,154,0.2)}
.badge-red{background:rgba(250,77,109,0.15);color:var(--red);border:1px solid rgba(250,77,109,0.2)}
.badge-gold{background:rgba(250,200,77,0.15);color:var(--gold);border:1px solid rgba(250,200,77,0.2)}

/* MAIN */
.container{position:relative;z-index:1;max-width:1300px;margin:0 auto;padding:1.5rem}
.page{display:none}.page.active{display:block}

/* METRICS GRID */
.metrics-grid{display:grid;grid-template-columns:repeat(5,1fr);gap:0.75rem;margin-bottom:1.5rem}
.metric-card{background:var(--surface);border:1px solid var(--border);border-radius:12px;padding:1rem 1.1rem;position:relative;overflow:hidden;cursor:default}
.metric-card::before{content:'';position:absolute;top:0;left:0;right:0;height:2px}
.mc-purple::before{background:var(--accent)}
.mc-green::before{background:var(--green)}
.mc-gold::before{background:var(--gold)}
.mc-red::before{background:var(--red)}
.mc-blue::before{background:var(--blue)}
.metric-label{font-size:0.58rem;letter-spacing:0.14em;text-transform:uppercase;color:var(--muted);font-family:var(--font-m);margin-bottom:0.4rem}
.metric-value{font-family:var(--font-m);font-size:1.4rem;font-weight:700;line-height:1}
.metric-sub{font-size:0.62rem;color:var(--muted);margin-top:0.35rem;font-family:var(--font-m)}

/* PANELS */
.panel{background:var(--surface);border:1px solid var(--border);border-radius:14px;overflow:hidden}
.panel-header{padding:1rem 1.25rem;border-bottom:1px solid var(--border);display:flex;align-items:center;justify-content:space-between;gap:1rem}
.panel-title{font-size:0.72rem;font-weight:700;letter-spacing:0.1em;text-transform:uppercase;color:var(--muted);font-family:var(--font-m)}
.panel-body{padding:1.25rem}

/* LAYOUT */
.two-col{display:grid;grid-template-columns:1fr 360px;gap:1.25rem;align-items:start}
.stack{display:flex;flex-direction:column;gap:1rem}

/* FORM */
.form-row{display:grid;gap:0.65rem;margin-bottom:0.65rem}
.fr2{grid-template-columns:1fr 1fr}
.fr3{grid-template-columns:1fr 1fr 1fr}
.fr4{grid-template-columns:1fr 1fr 1fr 1fr}
.field label{display:block;font-size:0.58rem;letter-spacing:0.12em;text-transform:uppercase;color:var(--muted);font-family:var(--font-m);margin-bottom:0.3rem}
.field input,.field select{background:var(--surface2);border:1px solid var(--border);border-radius:8px;padding:0.5rem 0.7rem;color:var(--text);font-family:var(--font-m);font-size:0.78rem;outline:none;transition:border-color 0.2s;width:100%}
.field input:focus,.field select:focus{border-color:var(--accent)}
option{background:#1a1a24}

/* ITEMS */
.item-row{display:grid;grid-template-columns:18px 1fr 1fr 1fr 70px 28px;gap:0.4rem;align-items:center;padding:0.55rem 0.7rem;background:var(--surface2);border-radius:7px;border:1px solid var(--border);margin-bottom:0.4rem}
.item-num{font-family:var(--font-m);font-size:0.6rem;color:var(--muted);text-align:center}
.ii{background:transparent;border:none;color:var(--text);font-family:var(--font-m);font-size:0.76rem;outline:none;width:100%}
.ii::placeholder{color:var(--muted)}
.ic{background:transparent;border:none;border-left:1px solid var(--border);color:var(--gold);font-family:var(--font-m);font-size:0.76rem;font-weight:700;outline:none;width:100%;padding-left:0.5rem}
.ic::placeholder{color:var(--muted)}
.btn-rm{background:none;border:none;color:var(--muted);cursor:pointer;font-size:0.85rem;padding:0.1rem;transition:color 0.2s}
.btn-rm:hover{color:var(--red)}
.btn-add-item{background:none;border:1px dashed var(--border);border-radius:7px;color:var(--muted);font-family:var(--font-m);font-size:0.7rem;padding:0.45rem;width:100%;cursor:pointer;transition:all 0.2s;margin-bottom:0.75rem}
.btn-add-item:hover{border-color:var(--accent);color:var(--accent)}
.cuota-preview{background:var(--surface2);border:1px solid var(--border);border-radius:8px;padding:0.65rem 1rem;display:flex;align-items:center;justify-content:space-between;margin-bottom:0.75rem}
.cp-label{font-size:0.6rem;letter-spacing:0.12em;text-transform:uppercase;color:var(--muted);font-family:var(--font-m)}
.cp-value{font-family:var(--font-m);font-size:1.2rem;font-weight:700;color:var(--gold)}

/* BUTTONS */
.btn-primary{background:var(--accent);color:#fff;border:none;border-radius:9px;padding:0.7rem 1.25rem;font-family:var(--font-d);font-size:0.85rem;font-weight:700;cursor:pointer;width:100%;transition:all 0.2s}
.btn-primary:hover{filter:brightness(1.15);transform:translateY(-1px)}
.btn-sm{background:var(--surface2);border:1px solid var(--border);color:var(--text);border-radius:6px;padding:0.3rem 0.7rem;font-family:var(--font-m);font-size:0.65rem;cursor:pointer;transition:all 0.2s;white-space:nowrap}
.btn-sm:hover{border-color:var(--accent);color:var(--accent)}
.btn-export{background:rgba(77,184,250,0.1);border:1px solid rgba(77,184,250,0.3);color:var(--blue);border-radius:6px;padding:0.35rem 0.8rem;font-family:var(--font-m);font-size:0.65rem;cursor:pointer;transition:all 0.2s;display:flex;align-items:center;gap:0.4rem}
.btn-export:hover{background:rgba(77,184,250,0.2)}
.btn-danger{background:rgba(250,77,109,0.1);border:1px solid rgba(250,77,109,0.3);color:var(--red);border-radius:6px;padding:0.3rem 0.65rem;font-family:var(--font-m);font-size:0.62rem;cursor:pointer;transition:all 0.2s}
.btn-danger:hover{background:rgba(250,77,109,0.2)}
.btn-won{background:rgba(77,250,154,0.12);border:1px solid rgba(77,250,154,0.25);color:var(--green);border-radius:6px;padding:0.28rem 0.6rem;font-family:var(--font-m);font-size:0.62rem;cursor:pointer;transition:all 0.2s}
.btn-won:hover{background:rgba(77,250,154,0.22)}
.btn-lost{background:rgba(250,77,109,0.12);border:1px solid rgba(250,77,109,0.25);color:var(--red);border-radius:6px;padding:0.28rem 0.6rem;font-family:var(--font-m);font-size:0.62rem;cursor:pointer;transition:all 0.2s}
.btn-lost:hover{background:rgba(250,77,109,0.22)}

/* HISTORY */
.bet-card{background:var(--surface2);border:1px solid var(--border);border-radius:11px;overflow:hidden;margin-bottom:0.6rem;transition:border-color 0.2s}
.bet-card.won{border-left:3px solid var(--green)}
.bet-card.lost{border-left:3px solid var(--red)}
.bet-card.pending{border-left:3px solid rgba(107,107,136,0.4)}
.bet-card-hdr{padding:0.65rem 0.9rem;display:flex;align-items:center;justify-content:space-between;cursor:pointer;user-select:none}
.bc-left{display:flex;align-items:center;gap:0.65rem}
.status-dot{width:7px;height:7px;border-radius:50%;flex-shrink:0}
.sd-won{background:var(--green)}
.sd-lost{background:var(--red)}
.sd-pending{background:var(--muted)}
.bc-meta .bc-date{font-size:0.58rem;letter-spacing:0.06em;color:var(--muted);font-family:var(--font-m)}
.bc-meta .bc-name{font-size:0.82rem;font-weight:600}
.bc-right{display:flex;align-items:center;gap:0.65rem}
.cuota-badge{background:rgba(250,200,77,0.1);color:var(--gold);font-family:var(--font-m);font-size:0.65rem;font-weight:700;padding:0.15rem 0.45rem;border-radius:4px}
.bc-profit{font-family:var(--font-m);font-size:0.82rem;font-weight:700}
.pos{color:var(--green)}.neg{color:var(--red)}.pen{color:var(--muted)}
.bet-items-body{display:none;padding:0.5rem 0.9rem 0.75rem;border-top:1px solid var(--border)}
.bet-card.open .bet-items-body{display:block}
.bi-row{display:flex;align-items:center;justify-content:space-between;padding:0.35rem 0;border-bottom:1px solid var(--border)}
.bi-row:last-child{border-bottom:none}
.bi-teams{font-size:0.73rem}
.bi-pick{font-size:0.62rem;color:var(--muted);font-family:var(--font-m)}
.bi-right{display:flex;align-items:center;gap:0.4rem}
.bi-cuota{font-family:var(--font-m);font-size:0.68rem;color:var(--gold)}
.res-badge{font-size:0.6rem;font-weight:700;padding:0.1rem 0.35rem;border-radius:3px}
.res-si{background:rgba(77,250,154,0.15);color:var(--green)}
.res-no{background:rgba(250,77,109,0.15);color:var(--red)}
.res-pen{background:rgba(107,107,136,0.15);color:var(--muted)}
.bc-actions{display:flex;gap:0.4rem;margin-top:0.6rem;flex-wrap:wrap;align-items:center}

/* FILTERS */
.filters-bar{display:flex;align-items:center;gap:0.5rem;flex-wrap:wrap;margin-bottom:1rem}
.filter-btn{background:var(--surface2);border:1px solid var(--border);color:var(--muted);border-radius:6px;padding:0.3rem 0.7rem;font-family:var(--font-m);font-size:0.65rem;cursor:pointer;transition:all 0.2s}
.filter-btn.active,.filter-btn:hover{border-color:var(--accent);color:var(--accent);background:rgba(124,109,250,0.08)}
.filter-sep{width:1px;height:20px;background:var(--border);margin:0 0.25rem}
.search-input{background:var(--surface2);border:1px solid var(--border);border-radius:6px;padding:0.3rem 0.7rem;color:var(--text);font-family:var(--font-m);font-size:0.68rem;outline:none;transition:border-color 0.2s;width:160px}
.search-input:focus{border-color:var(--accent)}

/* ANALYTICS */
.analytics-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:1rem;margin-bottom:1.25rem}
.chart-wrap{background:var(--surface);border:1px solid var(--border);border-radius:14px;padding:1.1rem}
.chart-title{font-size:0.65rem;letter-spacing:0.12em;text-transform:uppercase;color:var(--muted);font-family:var(--font-m);margin-bottom:0.9rem}

/* MONTHLY TABLE */
.data-table{width:100%;border-collapse:collapse;font-family:var(--font-m);font-size:0.72rem}
.data-table th{font-size:0.58rem;letter-spacing:0.1em;text-transform:uppercase;color:var(--muted);padding:0.5rem 0.75rem;text-align:left;border-bottom:1px solid var(--border)}
.data-table td{padding:0.55rem 0.75rem;border-bottom:1px solid rgba(42,42,58,0.5)}
.data-table tr:last-child td{border-bottom:none}
.data-table tr:hover td{background:var(--surface2)}
.tfoot-row td{font-weight:700;border-top:1px solid var(--border);color:var(--text)}

/* PERIOD SELECTOR */
.period-tabs{display:flex;gap:0.25rem;background:var(--surface2);padding:0.25rem;border-radius:8px;border:1px solid var(--border)}
.period-tab{background:none;border:none;color:var(--muted);font-family:var(--font-m);font-size:0.65rem;padding:0.3rem 0.75rem;border-radius:6px;cursor:pointer;transition:all 0.2s}
.period-tab.active{background:var(--accent);color:#fff}

/* CHART BARS */
.bar-chart{display:flex;align-items:flex-end;gap:3px;height:100px;width:100%}
.bar-wrap{flex:1;display:flex;flex-direction:column;align-items:center;gap:3px;justify-content:flex-end;height:100%}
.bar{width:100%;border-radius:3px 3px 0 0;min-height:2px;transition:opacity 0.2s;position:relative}
.bar.pos{background:var(--green)}
.bar.neg{background:var(--red)}
.bar.zero{background:var(--muted)}
.bar-lbl{font-family:var(--font-m);font-size:0.5rem;color:var(--muted);text-align:center;white-space:nowrap}

/* DONUT */
canvas{display:block}

/* DIVIDER */
.divider{display:flex;align-items:center;gap:0.75rem;margin:0.75rem 0 0.65rem}
.divider span{font-size:0.58rem;letter-spacing:0.12em;text-transform:uppercase;color:var(--muted);font-family:var(--font-m);white-space:nowrap}
.divider::before,.divider::after{content:'';flex:1;height:1px;background:var(--border)}

/* STREAK */
.streak-dots{display:flex;gap:0.35rem;flex-wrap:wrap}
.sdot{width:22px;height:22px;border-radius:5px;display:flex;align-items:center;justify-content:center;font-size:0.55rem;font-weight:700;font-family:var(--font-m)}
.sdot.w{background:rgba(77,250,154,0.18);color:var(--green)}
.sdot.l{background:rgba(250,77,109,0.18);color:var(--red)}
.sdot.p{background:rgba(107,107,136,0.18);color:var(--muted)}

/* TOAST */
.toast{position:fixed;bottom:1.5rem;right:1.5rem;background:var(--surface);border:1px solid var(--accent);border-radius:9px;padding:0.65rem 1.1rem;font-size:0.75rem;font-family:var(--font-m);color:var(--text);z-index:999;transform:translateY(16px);opacity:0;transition:all 0.3s;pointer-events:none;max-width:280px}
.toast.show{transform:translateY(0);opacity:1}

/* EMPTY */
.empty{text-align:center;padding:2.5rem 1rem;color:var(--muted)}
.empty-icon{font-size:2rem;margin-bottom:0.6rem;opacity:0.3}
.empty-txt{font-size:0.72rem;font-family:var(--font-m)}

/* GROUPED HISTORY */
.day-group{margin-bottom:1.25rem}
.day-label{font-size:0.6rem;letter-spacing:0.12em;text-transform:uppercase;color:var(--muted);font-family:var(--font-m);padding:0.35rem 0;border-bottom:1px solid var(--border);margin-bottom:0.6rem;display:flex;justify-content:space-between;align-items:center}
.day-balance{font-weight:700}

@keyframes fadeIn{from{opacity:0;transform:translateY(-6px)}to{opacity:1;transform:translateY(0)}}
.bet-card{animation:fadeIn 0.2s ease}

/* RESPONSIVE */
@media(max-width:900px){
  .two-col{grid-template-columns:1fr}
  .metrics-grid{grid-template-columns:repeat(3,1fr)}
  .analytics-grid{grid-template-columns:1fr}
}
</style>
</head>
<body>
<div class="orb orb1"></div>
<div class="orb orb2"></div>

<nav>
  <div class="logo">BET<span>TRACK</span> <span style="font-size:0.55rem;color:var(--muted);font-weight:400">PRO</span></div>
  <div class="nav-tabs">
    <button class="nav-tab active" onclick="switchPage('dashboard')">Dashboard</button>
    <button class="nav-tab" onclick="switchPage('historial')">Historial</button>
    <button class="nav-tab" onclick="switchPage('analytics')">Análisis</button>
    <button class="nav-tab" onclick="switchPage('mensual')">Mensual</button>
  </div>
  <div class="nav-right">
    <span class="nav-badge badge-gold" id="nav-roi">ROI 0%</span>
    <span class="nav-badge" id="nav-balance" style="background:rgba(77,250,154,0.12);color:var(--green);border:1px solid rgba(77,250,154,0.2)">S/ 0.00</span>
  </div>
</nav>

<!-- ═══ PAGE: DASHBOARD ═══ -->
<div class="container page active" id="page-dashboard">
  <div class="metrics-grid">
    <div class="metric-card mc-purple">
      <div class="metric-label">Total apostado</div>
      <div class="metric-value" id="m-total">S/ 0</div>
      <div class="metric-sub" id="m-bets">0 combinadas</div>
    </div>
    <div class="metric-card mc-green">
      <div class="metric-label">Ganancia neta</div>
      <div class="metric-value" id="m-profit">S/ 0</div>
      <div class="metric-sub" id="m-wonlost">0G · 0P · 0⏳</div>
    </div>
    <div class="metric-card mc-gold">
      <div class="metric-label">Win rate</div>
      <div class="metric-value" id="m-wr">—</div>
      <div class="metric-sub" id="m-wrn">de apuestas resueltas</div>
    </div>
    <div class="metric-card mc-red">
      <div class="metric-label">Cuota media</div>
      <div class="metric-value" id="m-avgodds">—</div>
      <div class="metric-sub" id="m-maxodds">Mayor: —</div>
    </div>
    <div class="metric-card mc-blue">
      <div class="metric-label">Mes actual</div>
      <div class="metric-value" id="m-month">S/ 0</div>
      <div class="metric-sub" id="m-monthsub">0 apuestas este mes</div>
    </div>
  </div>

  <div class="two-col">
    <!-- FORM -->
    <div>
      <div class="panel" style="margin-bottom:1.25rem">
        <div class="panel-header">
          <span class="panel-title">Nueva combinada</span>
          <span class="nav-badge badge-gold" id="fc-count">1 SEL.</span>
        </div>
        <div class="panel-body">
          <div class="form-row fr2">
            <div class="field"><label>Nombre / Liga</label><input type="text" id="f-name" placeholder="ej. Champions jornada 5"></div>
            <div class="field"><label>Fecha</label><input type="date" id="f-date"></div>
          </div>
          <div class="divider"><span>Selecciones</span></div>
          <div id="items-container"></div>
          <button class="btn-add-item" onclick="addItem()">+ Agregar selección</button>
          <div class="cuota-preview">
            <span class="cp-label">Cuota combinada total</span>
            <span class="cp-value" id="total-cuota">1.00×</span>
          </div>
          <div class="form-row fr3" style="margin-bottom:0.75rem">
            <div class="field"><label>Monto (S/)</label><input type="number" id="f-monto" placeholder="25.00" step="0.01" min="0" oninput="updateCuota()"></div>
            <div class="field"><label>Ganancia potencial</label><input type="text" id="f-pot" readonly style="color:var(--gold);cursor:default"></div>
            <div class="field"><label>Estado</label>
              <select id="f-estado">
                <option value="pending">⏳ Pendiente</option>
                <option value="won">✅ Ganada</option>
                <option value="lost">❌ Perdida</option>
              </select>
            </div>
          </div>
          <button class="btn-primary" onclick="saveBet()">Registrar combinada →</button>
        </div>
      </div>

      <!-- Recent bets on dashboard -->
      <div class="panel">
        <div class="panel-header">
          <span class="panel-title">Últimas apuestas</span>
          <button class="btn-sm" onclick="switchPage('historial')">Ver todas →</button>
        </div>
        <div class="panel-body" style="padding:0.75rem">
          <div id="dash-recent"></div>
        </div>
      </div>
    </div>

    <!-- SIDEBAR -->
    <div class="stack">
      <div class="panel">
        <div class="panel-header"><span class="panel-title">Rendimiento (últimas 10)</span></div>
        <div class="panel-body" style="padding:0.9rem">
          <div class="bar-chart" id="dash-chart" style="height:90px"></div>
        </div>
      </div>
      <div class="panel">
        <div class="panel-header"><span class="panel-title">Racha reciente</span></div>
        <div class="panel-body"><div class="streak-dots" id="dash-streak"></div></div>
      </div>
      <div class="panel">
        <div class="panel-header"><span class="panel-title">Resolver pendientes</span></div>
        <div class="panel-body"><div id="dash-pending"></div></div>
      </div>
      <div class="panel">
        <div class="panel-header"><span class="panel-title">Mejores ganancias</span></div>
        <div class="panel-body"><div id="dash-top"></div></div>
      </div>
    </div>
  </div>
</div>

<!-- ═══ PAGE: HISTORIAL ═══ -->
<div class="container page" id="page-historial">
  <div class="panel">
    <div class="panel-header">
      <span class="panel-title">Historial completo</span>
      <div style="display:flex;gap:0.5rem">
        <button class="btn-export" onclick="exportExcel()">⬇ Excel</button>
        <button class="btn-export" onclick="exportCSV()" style="background:rgba(77,250,154,0.08);border-color:rgba(77,250,154,0.25);color:var(--green)">⬇ CSV</button>
      </div>
    </div>
    <div class="panel-body">
      <div class="filters-bar">
        <button class="filter-btn active" onclick="setFilter('all',this)">Todas</button>
        <button class="filter-btn" onclick="setFilter('won',this)">✅ Ganadas</button>
        <button class="filter-btn" onclick="setFilter('lost',this)">❌ Perdidas</button>
        <button class="filter-btn" onclick="setFilter('pending',this)">⏳ Pendientes</button>
        <div class="filter-sep"></div>
        <button class="filter-btn active" data-period="all" onclick="setPeriod('all',this)">Todo</button>
        <button class="filter-btn" data-period="today" onclick="setPeriod('today',this)">Hoy</button>
        <button class="filter-btn" data-period="week" onclick="setPeriod('week',this)">Semana</button>
        <button class="filter-btn" data-period="month" onclick="setPeriod('month',this)">Mes</button>
        <div class="filter-sep"></div>
        <input type="text" class="search-input" placeholder="Buscar..." oninput="setSearch(this.value)" id="hist-search">
        <div style="display:flex;gap:0.35rem;margin-left:auto">
          <button class="filter-btn active" data-group="group" onclick="setGroupMode('group',this)">Agrupar</button>
          <button class="filter-btn" data-group="list" onclick="setGroupMode('list',this)">Lista</button>
        </div>
      </div>
      <div id="hist-list"></div>
    </div>
  </div>
</div>

<!-- ═══ PAGE: ANALYTICS ═══ -->
<div class="container page" id="page-analytics">
  <div class="metrics-grid" style="grid-template-columns:repeat(4,1fr)">
    <div class="metric-card mc-purple">
      <div class="metric-label">Total invertido</div>
      <div class="metric-value" id="an-total">S/ 0</div>
      <div class="metric-sub" id="an-avg">Promedio: —</div>
    </div>
    <div class="metric-card mc-green">
      <div class="metric-label">Profit total</div>
      <div class="metric-value" id="an-profit">S/ 0</div>
      <div class="metric-sub" id="an-roi">ROI: 0%</div>
    </div>
    <div class="metric-card mc-gold">
      <div class="metric-label">Mayor ganancia</div>
      <div class="metric-value" id="an-best">—</div>
      <div class="metric-sub" id="an-best-name">—</div>
    </div>
    <div class="metric-card mc-red">
      <div class="metric-label">Mayor pérdida</div>
      <div class="metric-value" id="an-worst">—</div>
      <div class="metric-sub" id="an-worst-name">—</div>
    </div>
  </div>

  <div style="display:grid;grid-template-columns:2fr 1fr;gap:1rem;margin-bottom:1rem">
    <div class="chart-wrap">
      <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:0.9rem">
        <div class="chart-title">Profit acumulado por periodo</div>
        <div class="period-tabs">
          <button class="period-tab active" onclick="setChartPeriod('weekly',this)">Semanas</button>
          <button class="period-tab" onclick="setChartPeriod('monthly',this)">Meses</button>
        </div>
      </div>
      <div class="bar-chart" id="an-chart" style="height:130px;gap:4px"></div>
    </div>
    <div class="chart-wrap">
      <div class="chart-title">Distribución resultados</div>
      <canvas id="donut-chart" width="200" height="200" style="margin:0 auto;display:block;max-width:180px;max-height:180px"></canvas>
      <div id="donut-legend" style="display:flex;gap:1rem;justify-content:center;margin-top:0.75rem;font-family:var(--font-m);font-size:0.62rem"></div>
    </div>
  </div>

  <div style="display:grid;grid-template-columns:1fr 1fr;gap:1rem">
    <div class="chart-wrap">
      <div class="chart-title">Racha más larga ganando / perdiendo</div>
      <div id="an-streaks" style="font-family:var(--font-m);font-size:0.78rem"></div>
    </div>
    <div class="chart-wrap">
      <div class="chart-title">Balance por día de la semana</div>
      <div class="bar-chart" id="an-weekday" style="height:100px"></div>
    </div>
  </div>
</div>

<!-- ═══ PAGE: MENSUAL ═══ -->
<div class="container page" id="page-mensual">
  <div class="panel">
    <div class="panel-header">
      <span class="panel-title">Balance mensual</span>
      <button class="btn-export" onclick="exportExcelMonthly()">⬇ Excel mensual</button>
    </div>
    <div class="panel-body">
      <div id="monthly-table-wrap"></div>
    </div>
  </div>
  <div style="height:1rem"></div>
  <div class="panel">
    <div class="panel-header">
      <span class="panel-title">Balance semanal</span>
    </div>
    <div class="panel-body">
      <div id="weekly-table-wrap"></div>
    </div>
  </div>
</div>

<div class="toast" id="toast"></div>

<script>
// ═══════════════════════════════════════════════════
// STATE
// ═══════════════════════════════════════════════════
let bets = JSON.parse(localStorage.getItem('btpro_bets')||'[]');
let formItems = [{local:'',visita:'',apuesta:'',cuota:''}];
let histFilter = 'all';
let histPeriod = 'all';
let histSearch = '';
let histGroupMode = 'group';
let chartPeriod = 'weekly';
let currentPage = 'dashboard';

// ═══════════════════════════════════════════════════
// NAVIGATION
// ═══════════════════════════════════════════════════
function switchPage(p){
  currentPage = p;
  document.querySelectorAll('.page').forEach(el=>el.classList.remove('active'));
  document.getElementById('page-'+p).classList.add('active');
  document.querySelectorAll('.nav-tab').forEach(el=>el.classList.remove('active'));
  document.querySelectorAll('.nav-tab').forEach(el=>{
    if(el.textContent.toLowerCase().includes(p==='dashboard'?'dash':p==='historial'?'hist':p==='analytics'?'an':'mens')) el.classList.add('active');
  });
  // simpler: index
  const tabs=['dashboard','historial','analytics','mensual'];
  const navTabs=document.querySelectorAll('.nav-tab');
  navTabs.forEach((t,i)=>t.classList.toggle('active',tabs[i]===p));
  renderPage(p);
}
function renderPage(p){
  if(p==='dashboard'){renderDash();}
  else if(p==='historial'){renderHist();}
  else if(p==='analytics'){renderAnalytics();}
  else if(p==='mensual'){renderMensual();}
}

// ═══════════════════════════════════════════════════
// FORM
// ═══════════════════════════════════════════════════
function initForm(){
  const d=new Date();
  document.getElementById('f-date').value=d.toISOString().split('T')[0];
  formItems=[{local:'',visita:'',apuesta:'',cuota:''}];
  renderFormItems();
}
function addItem(){
  if(formItems.length>=8){showToast('Máximo 8 selecciones');return;}
  formItems.push({local:'',visita:'',apuesta:'',cuota:''});
  renderFormItems();
}
function removeItem(i){
  if(formItems.length<=1){showToast('Mínimo 1 selección');return;}
  formItems.splice(i,1);
  renderFormItems();
}
function renderFormItems(){
  const c=document.getElementById('items-container');
  c.innerHTML='';
  formItems.forEach((it,i)=>{
    const row=document.createElement('div');
    row.className='item-row';
    row.innerHTML=`<span class="item-num">${i+1}</span>
      <input class="ii" placeholder="Local" value="${it.local||''}" oninput="formItems[${i}].local=this.value">
      <input class="ii" placeholder="Visita" value="${it.visita||''}" oninput="formItems[${i}].visita=this.value">
      <input class="ii" placeholder="1 / X2 / BTTS..." value="${it.apuesta||''}" oninput="formItems[${i}].apuesta=this.value">
      <input class="ic" type="number" placeholder="1.00" step="0.01" min="1" value="${it.cuota||''}" oninput="formItems[${i}].cuota=this.value;updateCuota()">
      <button class="btn-rm" onclick="removeItem(${i})">✕</button>`;
    c.appendChild(row);
  });
  document.getElementById('fc-count').textContent=formItems.length+' SEL.';
  updateCuota();
}
function updateCuota(){
  let prod=1;
  formItems.forEach(it=>{const v=parseFloat(it.cuota);if(v&&v>1)prod*=v;});
  document.getElementById('total-cuota').textContent=prod.toFixed(2)+'×';
  const monto=parseFloat(document.getElementById('f-monto').value)||0;
  document.getElementById('f-pot').value=monto>0&&prod>1?'S/ '+((monto*prod)-monto).toFixed(2):'—';
}
function saveBet(){
  const name=document.getElementById('f-name').value.trim()||'Combinada';
  const date=document.getElementById('f-date').value;
  const monto=parseFloat(document.getElementById('f-monto').value);
  const estado=document.getElementById('f-estado').value;
  if(!date){showToast('Ingresa la fecha');return;}
  if(!monto||monto<=0){showToast('Ingresa el monto');return;}
  const valid=formItems.filter(it=>it.local||it.visita||it.apuesta||it.cuota);
  if(valid.length===0){showToast('Agrega al menos 1 selección');return;}
  let cuotaTotal=1;
  valid.forEach(it=>{const v=parseFloat(it.cuota);if(v&&v>1)cuotaTotal*=v;});
  cuotaTotal=parseFloat(cuotaTotal.toFixed(4));
  const profit=estado==='won'?parseFloat(((monto*cuotaTotal)-monto).toFixed(2)):estado==='lost'?-monto:null;
  const bet={id:Date.now(),name,date,monto,cuotaTotal,estado,profit,
    items:valid.map(it=>({...it,resultado:estado==='pending'?'pending':estado==='won'?'SI':'NO'}))};
  bets.unshift(bet);
  save();
  renderAll();
  initForm();
  showToast('✓ Combinada registrada');
}
function resolveBet(id,estado){
  const b=bets.find(x=>x.id===id);if(!b)return;
  b.estado=estado;
  b.profit=estado==='won'?parseFloat(((b.monto*b.cuotaTotal)-b.monto).toFixed(2)):-b.monto;
  b.items=b.items.map(it=>({...it,resultado:estado==='won'?'SI':'NO'}));
  save();renderAll();
  showToast(estado==='won'?'✅ Ganada':'❌ Perdida');
}
function deleteBet(id){
  if(!confirm('¿Eliminar esta apuesta?'))return;
  bets=bets.filter(b=>b.id!==id);
  save();renderAll();showToast('Eliminada');
}
function save(){localStorage.setItem('btpro_bets',JSON.stringify(bets));}

// ═══════════════════════════════════════════════════
// FILTERS
// ═══════════════════════════════════════════════════
function setFilter(f,el){
  histFilter=f;
  document.querySelectorAll('.filters-bar .filter-btn:not([data-period]):not([data-group])').forEach(b=>b.classList.remove('active'));
  el.classList.add('active');
  renderHist();
}
function setPeriod(p,el){
  histPeriod=p;
  document.querySelectorAll('[data-period]').forEach(b=>b.classList.remove('active'));
  el.classList.add('active');
  renderHist();
}
function setSearch(v){histSearch=v.toLowerCase();renderHist();}
function setGroupMode(m,el){
  histGroupMode=m;
  document.querySelectorAll('[data-group]').forEach(b=>b.classList.remove('active'));
  el.classList.add('active');
  renderHist();
}
function getFilteredBets(){
  let res=[...bets];
  if(histFilter!=='all')res=res.filter(b=>b.estado===histFilter);
  if(histPeriod!=='all'){
    const now=new Date();
    res=res.filter(b=>{
      const bd=new Date(b.date+'T12:00:00');
      if(histPeriod==='today'){
        return bd.toDateString()===now.toDateString();
      } else if(histPeriod==='week'){
        const wStart=new Date(now);wStart.setDate(now.getDate()-now.getDay());wStart.setHours(0,0,0,0);
        return bd>=wStart;
      } else if(histPeriod==='month'){
        return bd.getMonth()===now.getMonth()&&bd.getFullYear()===now.getFullYear();
      }
    });
  }
  if(histSearch)res=res.filter(b=>b.name.toLowerCase().includes(histSearch)||b.items.some(it=>(it.local+''+it.visita+''+it.apuesta).toLowerCase().includes(histSearch)));
  return res;
}

// ═══════════════════════════════════════════════════
// RENDER BET CARD
// ═══════════════════════════════════════════════════
function betCardHTML(bet){
  const sc=bet.estado==='won'?'won':bet.estado==='lost'?'lost':'pending';
  const dc=bet.estado==='won'?'sd-won':bet.estado==='lost'?'sd-lost':'sd-pending';
  const profTxt=bet.profit!==null?(bet.profit>=0?'+':'')+'S/ '+Math.abs(bet.profit).toFixed(2):'⏳';
  const profC=bet.profit===null?'pen':bet.profit>=0?'pos':'neg';
  const dt=bet.date?new Date(bet.date+'T12:00:00').toLocaleDateString('es-PE',{weekday:'short',day:'2-digit',month:'short'}):'';
  return `<div class="bet-card ${sc}" id="bc-${bet.id}">
    <div class="bet-card-hdr" onclick="toggleBC(${bet.id})">
      <div class="bc-left">
        <div class="status-dot ${dc}"></div>
        <div class="bc-meta">
          <div class="bc-date">${dt} · ${bet.items.length} sel. · S/ ${bet.monto.toFixed(2)}</div>
          <div class="bc-name">${bet.name}</div>
        </div>
      </div>
      <div class="bc-right">
        <span class="cuota-badge">${bet.cuotaTotal.toFixed(2)}×</span>
        <span class="bc-profit ${profC}">${profTxt}</span>
        <span style="color:var(--muted);font-size:0.7rem">▾</span>
      </div>
    </div>
    <div class="bet-items-body">
      <div style="font-size:0.6rem;color:var(--muted);font-family:var(--font-m);margin-bottom:0.5rem">
        Apostado: S/ ${bet.monto.toFixed(2)} · Pot. ganancia: S/ ${((bet.monto*bet.cuotaTotal)-bet.monto).toFixed(2)} · Retorno total: S/ ${(bet.monto*bet.cuotaTotal).toFixed(2)}
      </div>
      ${bet.items.map(it=>`
        <div class="bi-row">
          <div>
            <div class="bi-teams">${it.local||'—'} vs ${it.visita||'—'}</div>
            <div class="bi-pick">${it.apuesta||'—'}</div>
          </div>
          <div class="bi-right">
            <span class="bi-cuota">${parseFloat(it.cuota||1).toFixed(2)}×</span>
            <span class="res-badge ${it.resultado==='SI'?'res-si':it.resultado==='NO'?'res-no':'res-pen'}">${it.resultado==='SI'?'SI':it.resultado==='NO'?'NO':'—'}</span>
          </div>
        </div>`).join('')}
      <div class="bc-actions">
        ${bet.estado==='pending'?`
          <button class="btn-won" onclick="resolveBet(${bet.id},'won')">✅ Ganada</button>
          <button class="btn-lost" onclick="resolveBet(${bet.id},'lost')">❌ Perdida</button>`:''}
        <button class="btn-danger" onclick="deleteBet(${bet.id})">✕ Eliminar</button>
      </div>
    </div>
  </div>`;
}
function toggleBC(id){
  const el=document.getElementById('bc-'+id);
  if(el)el.classList.toggle('open');
}

// ═══════════════════════════════════════════════════
// RENDER DASHBOARD
// ═══════════════════════════════════════════════════
function renderDash(){
  const total=bets.reduce((s,b)=>s+b.monto,0);
  const won=bets.filter(b=>b.estado==='won');
  const lost=bets.filter(b=>b.estado==='lost');
  const pending=bets.filter(b=>b.estado==='pending');
  const profit=bets.reduce((s,b)=>s+(b.profit||0),0);
  const resolved=won.length+lost.length;
  const wr=resolved>0?(won.length/resolved*100):null;
  const cuotas=bets.map(b=>b.cuotaTotal).filter(c=>c>1);
  const avgOdds=cuotas.length?cuotas.reduce((a,b)=>a+b,0)/cuotas.length:null;
  const maxOdds=cuotas.length?Math.max(...cuotas):null;
  const now=new Date();
  const monthBets=bets.filter(b=>{const d=new Date(b.date+'T12:00:00');return d.getMonth()===now.getMonth()&&d.getFullYear()===now.getFullYear();});
  const monthProfit=monthBets.reduce((s,b)=>s+(b.profit||0),0);
  const roi=total>0?profit/total*100:0;

  document.getElementById('m-total').textContent='S/ '+total.toFixed(2);
  document.getElementById('m-bets').textContent=bets.length+' combinada'+(bets.length!==1?'s':'');
  const pe=document.getElementById('m-profit');
  pe.textContent=(profit>=0?'+':'')+'S/ '+profit.toFixed(2);
  pe.style.color=profit>0?'var(--green)':profit<0?'var(--red)':'var(--text)';
  document.getElementById('m-wonlost').textContent=`${won.length}G · ${lost.length}P · ${pending.length}⏳`;
  document.getElementById('m-wr').textContent=wr!==null?wr.toFixed(0)+'%':'—';
  document.getElementById('m-wrn').textContent=`de ${resolved} apuesta${resolved!==1?'s':''} resuelta${resolved!==1?'s':''}`;
  document.getElementById('m-avgodds').textContent=avgOdds?avgOdds.toFixed(2)+'×':'—';
  document.getElementById('m-maxodds').textContent='Mayor: '+(maxOdds?maxOdds.toFixed(2)+'×':'—');
  const me=document.getElementById('m-month');
  me.textContent=(monthProfit>=0?'+':'')+'S/ '+monthProfit.toFixed(2);
  me.style.color=monthProfit>0?'var(--green)':monthProfit<0?'var(--red)':'var(--text)';
  document.getElementById('m-monthsub').textContent=monthBets.length+' apuesta'+(monthBets.length!==1?'s':'')+' este mes';

  const nb=document.getElementById('nav-balance');
  nb.textContent=(profit>=0?'+':'')+'S/ '+profit.toFixed(2);
  nb.style.color=profit>=0?'var(--green)':'var(--red)';
  const nr=document.getElementById('nav-roi');
  nr.textContent='ROI '+roi.toFixed(1)+'%';
  nr.style.color=roi>0?'var(--green)':roi<0?'var(--red)':'var(--muted)';

  // Recent 5
  const rc=document.getElementById('dash-recent');
  if(bets.length===0){rc.innerHTML=emptyHTML();} else {
    rc.innerHTML=bets.slice(0,5).map(betCardHTML).join('');
  }
  // Chart
  renderBarChart('dash-chart', bets.filter(b=>b.estado!=='pending').slice(0,10).reverse().map(b=>({v:b.profit||0,l:b.date?new Date(b.date+'T12:00:00').toLocaleDateString('es-PE',{day:'2-digit',month:'2-digit'}):''})), 80);
  // Streak
  const sr=document.getElementById('dash-streak');
  const sl=bets.slice(0,15);
  sr.innerHTML=sl.length===0?'<span style="font-size:0.7rem;color:var(--muted);font-family:var(--font-m)">Sin datos</span>':
    sl.map(b=>`<div class="sdot ${b.estado==='won'?'w':b.estado==='lost'?'l':'p'}" title="${b.name}">${b.estado==='won'?'W':b.estado==='lost'?'L':'?'}</div>`).join('');
  // Pending
  const pd=document.getElementById('dash-pending');
  const plist=bets.filter(b=>b.estado==='pending').slice(0,5);
  pd.innerHTML=plist.length===0?'<span style="font-size:0.7rem;color:var(--muted);font-family:var(--font-m)">Sin pendientes</span>':
    plist.map(b=>`<div style="display:flex;align-items:center;justify-content:space-between;padding:0.45rem 0;border-bottom:1px solid var(--border)">
      <div>
        <div style="font-size:0.76rem;font-weight:600">${b.name}</div>
        <div style="font-size:0.6rem;color:var(--muted);font-family:var(--font-m)">${b.cuotaTotal.toFixed(2)}× · S/ ${b.monto.toFixed(2)}</div>
      </div>
      <div style="display:flex;gap:0.3rem">
        <button class="btn-won" onclick="resolveBet(${b.id},'won')">W</button>
        <button class="btn-lost" onclick="resolveBet(${b.id},'lost')">L</button>
      </div>
    </div>`).join('');
  // Top
  const tp=document.getElementById('dash-top');
  const tlist=bets.filter(b=>b.estado==='won').sort((a,b)=>b.profit-a.profit).slice(0,3);
  tp.innerHTML=tlist.length===0?'<span style="font-size:0.7rem;color:var(--muted);font-family:var(--font-m)">Sin ganancias aún</span>':
    tlist.map((b,i)=>`<div style="display:flex;align-items:center;justify-content:space-between;padding:0.45rem 0;border-bottom:1px solid var(--border)">
      <div style="display:flex;align-items:center;gap:0.5rem">
        <span style="font-family:var(--font-m);font-size:0.6rem;color:var(--muted)">#${i+1}</span>
        <div>
          <div style="font-size:0.76rem;font-weight:600">${b.name}</div>
          <div style="font-size:0.6rem;color:var(--muted);font-family:var(--font-m)">${b.cuotaTotal.toFixed(2)}×</div>
        </div>
      </div>
      <span style="font-family:var(--font-m);font-size:0.8rem;font-weight:700;color:var(--green)">+S/ ${b.profit.toFixed(2)}</span>
    </div>`).join('');
}

// ═══════════════════════════════════════════════════
// RENDER HISTORIAL
// ═══════════════════════════════════════════════════
function renderHist(){
  const filtered=getFilteredBets();
  const container=document.getElementById('hist-list');
  if(filtered.length===0){container.innerHTML=emptyHTML();return;}
  if(histGroupMode==='list'){
    container.innerHTML=filtered.map(betCardHTML).join('');
    return;
  }
  // Group by day
  const groups={};
  filtered.forEach(b=>{
    const dk=b.date||'Sin fecha';
    if(!groups[dk])groups[dk]=[];
    groups[dk].push(b);
  });
  const sortedKeys=Object.keys(groups).sort((a,b)=>b.localeCompare(a));
  let html='';
  sortedKeys.forEach(dk=>{
    const grp=groups[dk];
    const dayProfit=grp.reduce((s,b)=>s+(b.profit||0),0);
    const dayDate=dk!=='Sin fecha'?new Date(dk+'T12:00:00').toLocaleDateString('es-PE',{weekday:'long',day:'2-digit',month:'long',year:'numeric'}):'Sin fecha';
    const dpC=dayProfit>0?'pos':dayProfit<0?'neg':'pen';
    html+=`<div class="day-group">
      <div class="day-label">
        <span>${dayDate}</span>
        <span class="day-balance ${dpC}">${dayProfit>=0?'+':''}S/ ${dayProfit.toFixed(2)}</span>
      </div>
      ${grp.map(betCardHTML).join('')}
    </div>`;
  });
  container.innerHTML=html;
}

// ═══════════════════════════════════════════════════
// RENDER ANALYTICS
// ═══════════════════════════════════════════════════
function setChartPeriod(p,el){
  chartPeriod=p;
  document.querySelectorAll('.period-tab').forEach(t=>t.classList.remove('active'));
  el.classList.add('active');
  renderAnalyticsChart();
}
function renderAnalytics(){
  const total=bets.reduce((s,b)=>s+b.monto,0);
  const profit=bets.reduce((s,b)=>s+(b.profit||0),0);
  const roi=total>0?profit/total*100:0;
  const avg=bets.length?total/bets.length:0;
  const won=bets.filter(b=>b.estado==='won');
  const lost=bets.filter(b=>b.estado==='lost');
  const pending=bets.filter(b=>b.estado==='pending');
  const wonProfit=won.length?won.reduce((s,b)=>s+b.profit,0):null;
  const lostBest=lost.length?lost.slice().sort((a,b)=>a.profit-b.profit)[0]:null;

  document.getElementById('an-total').textContent='S/ '+total.toFixed(2);
  document.getElementById('an-avg').textContent='Promedio: S/ '+avg.toFixed(2);
  const ape=document.getElementById('an-profit');
  ape.textContent=(profit>=0?'+':'')+'S/ '+profit.toFixed(2);
  ape.style.color=profit>=0?'var(--green)':'var(--red)';
  document.getElementById('an-roi').textContent='ROI: '+(roi>=0?'+':'')+roi.toFixed(2)+'%';
  // best
  if(won.length){
    const best=won.slice().sort((a,b)=>b.profit-a.profit)[0];
    document.getElementById('an-best').textContent='+S/ '+best.profit.toFixed(2);
    document.getElementById('an-best').style.color='var(--green)';
    document.getElementById('an-best-name').textContent=best.name;
  }
  if(lost.length){
    const worst=lost.slice().sort((a,b)=>a.profit-b.profit)[0];
    document.getElementById('an-worst').textContent='S/ '+Math.abs(worst.profit).toFixed(2);
    document.getElementById('an-worst').style.color='var(--red)';
    document.getElementById('an-worst-name').textContent=worst.name;
  }
  // Donut
  renderDonut([won.length,lost.length,pending.length],['Ganadas','Perdidas','Pendientes'],['#4dfa9a','#fa4d6d','#6b6b88']);
  // Streaks
  const res=bets.filter(b=>b.estado!=='pending').map(b=>b.estado).reverse();
  let maxW=0,maxL=0,curW=0,curL=0;
  res.forEach(s=>{
    if(s==='won'){curW++;curL=0;}else{curL++;curW=0;}
    maxW=Math.max(maxW,curW);maxL=Math.max(maxL,curL);
  });
  document.getElementById('an-streaks').innerHTML=`
    <div style="display:flex;gap:1.5rem">
      <div>
        <div style="font-size:0.6rem;color:var(--muted);font-family:var(--font-m);margin-bottom:0.25rem">Racha ganando máx.</div>
        <div style="font-size:2rem;font-weight:700;color:var(--green);font-family:var(--font-m)">${maxW}</div>
      </div>
      <div>
        <div style="font-size:0.6rem;color:var(--muted);font-family:var(--font-m);margin-bottom:0.25rem">Racha perdiendo máx.</div>
        <div style="font-size:2rem;font-weight:700;color:var(--red);font-family:var(--font-m)">${maxL}</div>
      </div>
      <div>
        <div style="font-size:0.6rem;color:var(--muted);font-family:var(--font-m);margin-bottom:0.25rem">Apuestas resueltas</div>
        <div style="font-size:2rem;font-weight:700;color:var(--accent);font-family:var(--font-m)">${res.length}</div>
      </div>
    </div>`;
  // Weekday chart
  const days=['Dom','Lun','Mar','Mié','Jue','Vie','Sáb'];
  const wdProfit=[0,0,0,0,0,0,0];
  bets.filter(b=>b.profit!==null).forEach(b=>{
    const d=new Date(b.date+'T12:00:00').getDay();
    wdProfit[d]+=b.profit;
  });
  renderBarChart('an-weekday',days.map((l,i)=>({v:wdProfit[i],l})),90);
  renderAnalyticsChart();
}
function renderAnalyticsChart(){
  const resolved=bets.filter(b=>b.estado!=='pending'&&b.date);
  if(resolved.length===0){document.getElementById('an-chart').innerHTML='<span style="font-size:0.7rem;color:var(--muted);font-family:var(--font-m);margin:auto">Sin datos</span>';return;}
  let groups={};
  resolved.forEach(b=>{
    let key;
    const d=new Date(b.date+'T12:00:00');
    if(chartPeriod==='monthly'){
      key=`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`;
    } else {
      const mon=new Date(d);mon.setDate(d.getDate()-d.getDay());
      key=mon.toISOString().split('T')[0];
    }
    if(!groups[key])groups[key]={profit:0,label:''};
    groups[key].profit+=b.profit||0;
  });
  const sorted=Object.keys(groups).sort();
  const data=sorted.map(k=>{
    let l;
    if(chartPeriod==='monthly'){const[y,m]=k.split('-');l=new Date(+y,+m-1).toLocaleDateString('es-PE',{month:'short',year:'2-digit'});}
    else{const d=new Date(k+'T12:00:00');l=d.toLocaleDateString('es-PE',{day:'2-digit',month:'2-digit'});}
    return{v:parseFloat(groups[k].profit.toFixed(2)),l};
  });
  renderBarChart('an-chart',data,120);
}
function renderBarChart(id,data,maxH=90){
  const c=document.getElementById(id);
  if(!data||data.length===0){c.innerHTML='<span style="font-size:0.65rem;color:var(--muted);font-family:var(--font-m);margin:auto">Sin datos</span>';return;}
  const maxAbs=Math.max(...data.map(d=>Math.abs(d.v)),0.01);
  c.innerHTML=data.map(d=>{
    const h=Math.max(3,Math.round(Math.abs(d.v)/maxAbs*maxH));
    const cls=d.v>0?'pos':d.v<0?'neg':'zero';
    return`<div class="bar-wrap" title="${d.l}: ${d.v>=0?'+':''}${d.v.toFixed(2)}">
      <div class="bar ${cls}" style="height:${h}px"></div>
      <div class="bar-lbl">${d.l}</div>
    </div>`;
  }).join('');
}
function renderDonut(data,labels,colors){
  const canvas=document.getElementById('donut-chart');
  const ctx=canvas.getContext('2d');
  const total=data.reduce((a,b)=>a+b,0);
  const cx=100,cy=100,r=75,inner=45;
  ctx.clearRect(0,0,200,200);
  if(total===0){
    ctx.fillStyle='#2a2a3a';ctx.beginPath();ctx.arc(cx,cy,r,0,Math.PI*2);ctx.fill();
    ctx.fillStyle='#0a0a0f';ctx.beginPath();ctx.arc(cx,cy,inner,0,Math.PI*2);ctx.fill();
    document.getElementById('donut-legend').innerHTML='<span style="color:var(--muted)">Sin datos</span>';
    return;
  }
  let startAngle=-Math.PI/2;
  data.forEach((v,i)=>{
    const slice=v/total*Math.PI*2;
    ctx.beginPath();ctx.moveTo(cx,cy);
    ctx.arc(cx,cy,r,startAngle,startAngle+slice);
    ctx.closePath();ctx.fillStyle=colors[i];ctx.fill();
    startAngle+=slice;
  });
  ctx.fillStyle='#111118';ctx.beginPath();ctx.arc(cx,cy,inner,0,Math.PI*2);ctx.fill();
  ctx.fillStyle='#f0f0ff';ctx.font='bold 18px Space Mono,monospace';ctx.textAlign='center';ctx.textBaseline='middle';
  ctx.fillText(total,cx,cy-8);
  ctx.fillStyle='#6b6b88';ctx.font='10px Space Mono,monospace';
  ctx.fillText('total',cx,cy+10);
  document.getElementById('donut-legend').innerHTML=data.map((v,i)=>`
    <div style="display:flex;align-items:center;gap:0.3rem">
      <div style="width:8px;height:8px;border-radius:2px;background:${colors[i]}"></div>
      <span style="color:var(--muted)">${labels[i]}: <strong style="color:var(--text)">${v}</strong></span>
    </div>`).join('');
}

// ═══════════════════════════════════════════════════
// RENDER MENSUAL
// ═══════════════════════════════════════════════════
function renderMensual(){
  renderMonthlyTable();
  renderWeeklyTable();
}
function getMonthlyData(){
  const groups={};
  bets.forEach(b=>{
    if(!b.date)return;
    const d=new Date(b.date+'T12:00:00');
    const k=`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`;
    if(!groups[k]){groups[k]={key:k,label:'',total:0,profit:0,won:0,lost:0,pending:0,bets:0};}
    const g=groups[k];
    g.label=d.toLocaleDateString('es-PE',{month:'long',year:'numeric'});
    g.total+=b.monto;g.profit+=(b.profit||0);
    g.bets++;
    if(b.estado==='won')g.won++;else if(b.estado==='lost')g.lost++;else g.pending++;
  });
  return Object.values(groups).sort((a,b)=>b.key.localeCompare(a.key));
}
function getWeeklyData(){
  const groups={};
  bets.forEach(b=>{
    if(!b.date)return;
    const d=new Date(b.date+'T12:00:00');
    const mon=new Date(d);mon.setDate(d.getDate()-d.getDay());
    const sun=new Date(mon);sun.setDate(mon.getDate()+6);
    const k=mon.toISOString().split('T')[0];
    if(!groups[k]){
      const label=`${mon.toLocaleDateString('es-PE',{day:'2-digit',month:'short'})} – ${sun.toLocaleDateString('es-PE',{day:'2-digit',month:'short',year:'numeric'})}`;
      groups[k]={key:k,label,total:0,profit:0,won:0,lost:0,pending:0,bets:0};
    }
    const g=groups[k];
    g.total+=b.monto;g.profit+=(b.profit||0);g.bets++;
    if(b.estado==='won')g.won++;else if(b.estado==='lost')g.lost++;else g.pending++;
  });
  return Object.values(groups).sort((a,b)=>b.key.localeCompare(a.key));
}
function tableHTML(rows,period){
  if(rows.length===0)return'<div class="empty"><div class="empty-txt">Sin datos para mostrar</div></div>';
  const totTotal=rows.reduce((s,r)=>s+r.total,0);
  const totProfit=rows.reduce((s,r)=>s+r.profit,0);
  const totWon=rows.reduce((s,r)=>s+r.won,0);
  const totLost=rows.reduce((s,r)=>s+r.lost,0);
  const roi=totTotal>0?totProfit/totTotal*100:0;
  return`<div style="overflow-x:auto"><table class="data-table">
    <thead><tr>
      <th>${period==='monthly'?'Mes':'Semana'}</th>
      <th>Apuestas</th><th>Ganadas</th><th>Perdidas</th><th>Pendientes</th>
      <th>Win%</th><th>Invertido</th><th>Profit</th><th>ROI</th>
    </tr></thead>
    <tbody>
      ${rows.map(r=>{
        const wr=r.won+r.lost>0?(r.won/(r.won+r.lost)*100):null;
        const roi=r.total>0?r.profit/r.total*100:0;
        const pc=r.profit>0?'var(--green)':r.profit<0?'var(--red)':'var(--muted)';
        const rc=roi>0?'var(--green)':roi<0?'var(--red)':'var(--muted)';
        return`<tr>
          <td style="font-weight:600;text-transform:capitalize">${r.label}</td>
          <td>${r.bets}</td>
          <td style="color:var(--green)">${r.won}</td>
          <td style="color:var(--red)">${r.lost}</td>
          <td style="color:var(--muted)">${r.pending}</td>
          <td>${wr!==null?wr.toFixed(0)+'%':'—'}</td>
          <td style="font-family:var(--font-m)">S/ ${r.total.toFixed(2)}</td>
          <td style="font-family:var(--font-m);font-weight:700;color:${pc}">${r.profit>=0?'+':''}S/ ${r.profit.toFixed(2)}</td>
          <td style="font-family:var(--font-m);color:${rc}">${roi>=0?'+':''}${roi.toFixed(1)}%</td>
        </tr>`;
      }).join('')}
    </tbody>
    <tfoot><tr class="tfoot-row">
      <td>TOTAL</td><td>${rows.reduce((s,r)=>s+r.bets,0)}</td>
      <td style="color:var(--green)">${totWon}</td>
      <td style="color:var(--red)">${totLost}</td>
      <td style="color:var(--muted)">${rows.reduce((s,r)=>s+r.pending,0)}</td>
      <td>${totWon+totLost>0?(totWon/(totWon+totLost)*100).toFixed(0)+'%':'—'}</td>
      <td style="font-family:var(--font-m)">S/ ${totTotal.toFixed(2)}</td>
      <td style="font-family:var(--font-m);font-weight:700;color:${totProfit>=0?'var(--green)':'var(--red)'}">${totProfit>=0?'+':''}S/ ${totProfit.toFixed(2)}</td>
      <td style="font-family:var(--font-m);color:${roi>=0?'var(--green)':'var(--red)'}">${roi>=0?'+':''}${roi.toFixed(1)}%</td>
    </tr></tfoot>
  </table></div>`;
}
function renderMonthlyTable(){document.getElementById('monthly-table-wrap').innerHTML=tableHTML(getMonthlyData(),'monthly');}
function renderWeeklyTable(){document.getElementById('weekly-table-wrap').innerHTML=tableHTML(getWeeklyData(),'weekly');}

// ═══════════════════════════════════════════════════
// EXPORT
// ═══════════════════════════════════════════════════
function exportExcel(){
  if(!bets.length){showToast('Sin apuestas para exportar');return;}
  const wb=XLSX.utils.book_new();
  // Sheet 1: Historial detallado
  const rows=[['Fecha','Nombre','Cuota Total','Monto (S/)','Ganancia (S/)','Estado','# Sels']];
  bets.forEach(b=>{
    rows.push([b.date,b.name,b.cuotaTotal,b.monto,b.profit!==null?b.profit:'',(b.estado==='won'?'Ganada':b.estado==='lost'?'Perdida':'Pendiente'),b.items.length]);
  });
  const ws1=XLSX.utils.aoa_to_sheet(rows);
  ws1['!cols']=[{wch:12},{wch:25},{wch:12},{wch:12},{wch:14},{wch:12},{wch:8}];
  XLSX.utils.book_append_sheet(wb,ws1,'Historial');
  // Sheet 2: Detalle selecciones
  const rows2=[['Fecha','Combinada','Sel.','Local','Visita','Apuesta','Cuota','Resultado']];
  bets.forEach(b=>{
    b.items.forEach((it,i)=>{
      rows2.push([b.date,b.name,i+1,it.local,it.visita,it.apuesta,parseFloat(it.cuota)||1,it.resultado==='SI'?'SI':it.resultado==='NO'?'NO':'Pendiente']);
    });
  });
  const ws2=XLSX.utils.aoa_to_sheet(rows2);
  ws2['!cols']=[{wch:12},{wch:25},{wch:5},{wch:18},{wch:18},{wch:15},{wch:8},{wch:12}];
  XLSX.utils.book_append_sheet(wb,ws2,'Selecciones');
  // Sheet 3: Mensual
  const mdata=getMonthlyData();
  const rows3=[['Mes','Apuestas','Ganadas','Perdidas','Pendientes','Win%','Invertido (S/)','Profit (S/)','ROI%']];
  mdata.forEach(r=>{const wr=r.won+r.lost>0?+(r.won/(r.won+r.lost)*100).toFixed(1):null;const roi=r.total>0?+(r.profit/r.total*100).toFixed(2):0;rows3.push([r.label,r.bets,r.won,r.lost,r.pending,wr,+r.total.toFixed(2),+r.profit.toFixed(2),roi]);});
  const ws3=XLSX.utils.aoa_to_sheet(rows3);
  ws3['!cols']=[{wch:20},{wch:10},{wch:10},{wch:10},{wch:12},{wch:8},{wch:14},{wch:14},{wch:8}];
  XLSX.utils.book_append_sheet(wb,ws3,'Mensual');
  // Sheet 4: Semanal
  const wdata=getWeeklyData();
  const rows4=[['Semana','Apuestas','Ganadas','Perdidas','Pendientes','Win%','Invertido (S/)','Profit (S/)','ROI%']];
  wdata.forEach(r=>{const wr=r.won+r.lost>0?+(r.won/(r.won+r.lost)*100).toFixed(1):null;const roi=r.total>0?+(r.profit/r.total*100).toFixed(2):0;rows4.push([r.label,r.bets,r.won,r.lost,r.pending,wr,+r.total.toFixed(2),+r.profit.toFixed(2),roi]);});
  const ws4=XLSX.utils.aoa_to_sheet(rows4);
  ws4['!cols']=[{wch:28},{wch:10},{wch:10},{wch:10},{wch:12},{wch:8},{wch:14},{wch:14},{wch:8}];
  XLSX.utils.book_append_sheet(wb,ws4,'Semanal');
  XLSX.writeFile(wb,'BetTrack_'+new Date().toISOString().split('T')[0]+'.xlsx');
  showToast('✓ Excel descargado (4 hojas)');
}
function exportExcelMonthly(){
  if(!bets.length){showToast('Sin datos');return;}
  const wb=XLSX.utils.book_new();
  const mdata=getMonthlyData();
  const rows=[['Mes','Apuestas','Ganadas','Perdidas','Pendientes','Win%','Invertido (S/)','Profit (S/)','ROI%']];
  mdata.forEach(r=>{const wr=r.won+r.lost>0?+(r.won/(r.won+r.lost)*100).toFixed(1):null;const roi=r.total>0?+(r.profit/r.total*100).toFixed(2):0;rows.push([r.label,r.bets,r.won,r.lost,r.pending,wr,+r.total.toFixed(2),+r.profit.toFixed(2),roi]);});
  const ws=XLSX.utils.aoa_to_sheet(rows);
  ws['!cols']=[{wch:20},{wch:10},{wch:10},{wch:10},{wch:12},{wch:8},{wch:14},{wch:14},{wch:8}];
  XLSX.utils.book_append_sheet(wb,ws,'Balance Mensual');
  XLSX.writeFile(wb,'BetTrack_Mensual_'+new Date().toISOString().split('T')[0]+'.xlsx');
  showToast('✓ Excel mensual descargado');
}
function exportCSV(){
  if(!bets.length){showToast('Sin apuestas');return;}
  const rows=[['Fecha','Nombre','Cuota Total','Monto','Profit','Estado','Selecciones']];
  bets.forEach(b=>rows.push([b.date,b.name,b.cuotaTotal,b.monto,b.profit??'',b.estado,b.items.length]));
  const csv=rows.map(r=>r.map(v=>'"'+String(v).replace(/"/g,'""')+'"').join(',')).join('\n');
  const blob=new Blob(['\ufeff'+csv],{type:'text/csv;charset=utf-8'});
  const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download='BetTrack_'+new Date().toISOString().split('T')[0]+'.csv';a.click();
  showToast('✓ CSV descargado');
}

// ═══════════════════════════════════════════════════
// UTILS
// ═══════════════════════════════════════════════════
function emptyHTML(){return'<div class="empty"><div class="empty-icon">◎</div><div class="empty-txt">Sin apuestas aquí todavía.</div></div>';}
function showToast(msg){const t=document.getElementById('toast');t.textContent=msg;t.classList.add('show');setTimeout(()=>t.classList.remove('show'),2500);}
function renderAll(){renderPage(currentPage);}

// ═══════════════════════════════════════════════════
// INIT
// ═══════════════════════════════════════════════════
initForm();
renderDash();
</script>
</body>
</html>
