<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Dashboard — Marcas, CNPJ e Responsáveis</title>
<script src="https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js"></script>
<style>
  *{box-sizing:border-box;margin:0;padding:0}
  html,body{height:100%}
  body{font-family:'Segoe UI',system-ui,-apple-system,Roboto,sans-serif;background:#f5f7fb;color:#1b2237;font-size:14px;-webkit-font-smoothing:antialiased}
  .app{display:grid;grid-template-columns:260px 1fr;min-height:100vh}
  @media(max-width:860px){.app{grid-template-columns:1fr}}
  .sidebar{background:#0f1736;color:#e8ecf8;padding:26px 20px;display:flex;flex-direction:column;gap:22px;position:sticky;top:0;height:100vh;overflow-y:auto}
  @media(max-width:860px){.sidebar{position:static;height:auto}}
  .brand{display:flex;align-items:center;gap:12px}
  .brand .logo{width:42px;height:42px;border-radius:12px;background:linear-gradient(135deg,#6366f1,#8b5cf6);display:grid;place-items:center;font-size:20px;box-shadow:0 8px 20px -8px rgba(139,92,246,.8)}
  .brand .name{font-size:15px;font-weight:700}
  .brand .sub{font-size:11.5px;color:#8a93b8;margin-top:2px}
  .nav-label{font-size:10.5px;font-weight:700;letter-spacing:1.2px;color:#5e6794;text-transform:uppercase;margin-bottom:8px}
  .side-actions{display:flex;flex-direction:column;gap:8px}
  .side-btn{display:flex;align-items:center;gap:12px;padding:12px 14px;border-radius:11px;border:none;cursor:pointer;font-family:inherit;font-size:13.5px;font-weight:600;background:transparent;color:#c2c9e2;text-align:left;transition:.15s}
  .side-btn:hover{background:rgba(255,255,255,.07);color:#fff}
  .side-btn .ic{width:30px;height:30px;border-radius:9px;display:grid;place-items:center;background:rgba(255,255,255,.09);font-size:15px;flex-shrink:0}
  .side-btn.primary{background:linear-gradient(135deg,#6366f1,#8b5cf6);color:#fff;box-shadow:0 10px 22px -12px rgba(139,92,246,.9)}
  .side-btn.primary:hover{transform:translateY(-1px)}
  .side-btn.primary .ic{background:rgba(255,255,255,.2)}
  .side-btn.danger{color:#fca5a5}
  .side-btn.danger:hover{background:rgba(239,68,68,.15);color:#fff}
  .side-divider{height:1px;background:rgba(255,255,255,.08);margin:2px 0}
  .side-info{background:rgba(255,255,255,.05);border-radius:12px;padding:14px;font-size:12px;color:#8a93b8;line-height:1.6;border:1px solid rgba(255,255,255,.06)}
  .side-info strong{color:#d7ddf5;display:block;font-size:12.5px;margin-bottom:4px}
  .status-box{display:flex;align-items:center;gap:8px;background:rgba(16,185,129,.12);border:1px solid rgba(16,185,129,.35);border-radius:10px;padding:10px 12px;font-size:11.5px;color:#6ee7b7;font-weight:600}
  .status-box .dot{width:8px;height:8px;border-radius:50%;background:#10b981;box-shadow:0 0 0 4px rgba(16,185,129,.25);flex-shrink:0}
  .status-box.warn{background:rgba(245,158,11,.12);border-color:rgba(245,158,11,.35);color:#fcd34d}
  .status-box.warn .dot{background:#f59e0b;box-shadow:0 0 0 4px rgba(245,158,11,.25)}
  .status-box.err{background:rgba(239,68,68,.12);border-color:rgba(239,68,68,.35);color:#fca5a5}
  .status-box.err .dot{background:#ef4444;box-shadow:0 0 0 4px rgba(239,68,68,.25)}
  .side-footer{margin-top:auto;font-size:11px;color:#5e6794;text-align:center;line-height:1.6}
  .main{padding:28px 32px 40px;display:flex;flex-direction:column;gap:22px;max-width:1600px}
  @media(max-width:860px){.main{padding:20px 16px 32px}}
  .page-header{display:flex;flex-wrap:wrap;gap:14px;align-items:flex-end;justify-content:space-between}
  .page-header h1{font-size:24px;font-weight:700;letter-spacing:-.5px;color:#0f1736}
  .page-header p{font-size:13.5px;color:#6b7492;margin-top:4px}
  .page-header .badge{font-size:11.5px;font-weight:700;color:#4338ca;background:#eef0ff;padding:6px 12px;border-radius:20px}
  .stats{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:14px}
  .stat{background:#fff;border-radius:14px;padding:18px 18px 16px;border:1px solid #e9edf5;display:flex;flex-direction:column;gap:6px;position:relative;overflow:hidden;transition:.2s}
  .stat:hover{transform:translateY(-2px);box-shadow:0 12px 26px -18px rgba(30,40,90,.4);border-color:#dde3ee}
  .stat::before{content:'';position:absolute;left:0;top:0;bottom:0;width:4px;background:var(--c,#4f46e5)}
  .stat .label{font-size:10.5px;font-weight:700;color:#7b84a2;text-transform:uppercase;letter-spacing:.7px}
  .stat .value{font-size:28px;font-weight:700;color:#0f1736;line-height:1;letter-spacing:-.5px}
  .stat .sub{font-size:11.5px;color:#9aa2bd}
  .section{background:#fff;border:1px solid #e9edf5;border-radius:16px;padding:20px 22px}
  .section-head{display:flex;align-items:center;justify-content:space-between;gap:12px;margin-bottom:16px;flex-wrap:wrap}
  .section-head h2{font-size:15.5px;font-weight:700;color:#0f1736;display:flex;align-items:center;gap:8px}
  .section-head h2 .dot{width:8px;height:8px;border-radius:50%;background:#6366f1}
  .section-head small{font-size:12px;color:#8b93ad;font-weight:600}
  .charts-grid{display:grid;grid-template-columns:1.4fr 1fr;gap:20px}
  @media(max-width:1000px){.charts-grid{grid-template-columns:1fr}}
  .bars{display:flex;flex-direction:column;gap:12px;max-height:340px;overflow-y:auto;padding-right:6px}
  .bars::-webkit-scrollbar{width:8px}.bars::-webkit-scrollbar-thumb{background:#dde3ee;border-radius:8px}
  .bar-row{display:grid;grid-template-columns:150px 1fr 44px;align-items:center;gap:12px;font-size:13px}
  .bar-name{white-space:nowrap;overflow:hidden;text-overflow:ellipsis;color:#3a4160;font-weight:500}
  .bar-track{background:#f0f3fa;border-radius:8px;height:22px;overflow:hidden}
  .bar-fill{height:100%;border-radius:8px;transition:width .6s cubic-bezier(.4,0,.2,1)}
  .bar-val{text-align:right;font-weight:700;color:#2c3350;font-variant-numeric:tabular-nums}
  .empty-msg{color:#9aa2bd;font-size:13px;text-align:center;padding:30px 0}
  .toolbar{display:flex;gap:10px;flex-wrap:wrap;align-items:center;margin-bottom:16px}
  .toolbar input,.toolbar select{padding:10px 13px;border:1px solid #e0e5ef;border-radius:10px;font-size:13px;font-family:inherit;background:#fafbfe;color:#1b2237;outline:none;transition:.15s}
  .toolbar input:focus,.toolbar select:focus{border-color:#6366f1;box-shadow:0 0 0 3px rgba(99,102,241,.13);background:#fff}
  .toolbar .search{flex:1;min-width:220px}
  .clear-btn{padding:10px 14px;border:1px solid #e0e5ef;background:#fff;border-radius:10px;font-family:inherit;font-size:13px;font-weight:600;color:#5b637f;cursor:pointer;transition:.15s}
  .clear-btn:hover{background:#f5f7fb;color:#1b2237}
  .count-pill{margin-left:auto;font-size:12.5px;font-weight:700;color:#4338ca;background:#eef0ff;padding:7px 14px;border-radius:20px;white-space:nowrap}
  .table-scroll{max-height:600px;overflow:auto;border:1px solid #e9edf5;border-radius:12px}
  .table-scroll::-webkit-scrollbar{width:10px;height:10px}
  .table-scroll::-webkit-scrollbar-thumb{background:#dde3ee;border-radius:8px;border:2px solid #fff}
  table{width:100%;border-collapse:separate;border-spacing:0;font-size:13.5px}
  thead th{position:sticky;top:0;background:#f7f9fd;z-index:3;text-align:left;padding:13px 14px;font-size:10.5px;letter-spacing:.8px;text-transform:uppercase;color:#7b84a2;font-weight:700;border-bottom:1px solid #e9edf5;white-space:nowrap}
  tbody td{padding:9px 14px;border-bottom:1px solid #f2f5fa;vertical-align:middle}
  tbody tr:hover{background:#fafbff}
  tbody tr:last-child td{border-bottom:none}
  .num{color:#b0b7cc;font-variant-numeric:tabular-nums;font-size:12px;font-weight:600}
  .marca{font-weight:600;color:#0f1736}
  .cnpj{font-variant-numeric:tabular-nums;color:#5b637f;font-size:13px;white-space:nowrap}
  .center{text-align:center}
  .tag{display:inline-block;padding:4px 11px;border-radius:20px;font-size:11.5px;font-weight:700;white-space:nowrap}
  .tag.informatica{background:#e0f2fe;color:#0369a1}
  .tag.area{background:#ede9fe;color:#6d28d9}
  .tag.coffe{background:#fef3c7;color:#b45309}
  .tag.outro{background:#eef1f8;color:#4b5573}
  .cell-input{width:100%;padding:8px 11px;border:1px solid transparent;background:#f5f7fc;border-radius:8px;font-size:13.5px;font-family:inherit;color:#0f1736;outline:none;transition:.15s}
  .cell-input:hover{background:#eef1fa}
  .cell-input:focus{background:#fff;border-color:#6366f1;box-shadow:0 0 0 3px rgba(99,102,241,.14)}
  .cell-input::placeholder{color:#adb5cb;font-weight:400}
  select.cell-input{cursor:pointer;appearance:none;background-image:url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 24 24' fill='none' stroke='%237b84a2' stroke-width='3'><polyline points='6 9 12 15 18 9'/></svg>");background-repeat:no-repeat;background-position:right 10px center;padding-right:28px}
  .icon-btn{border:none;background:transparent;color:#c2c8da;cursor:pointer;font-size:14px;padding:6px 9px;border-radius:8px;transition:.15s;line-height:1}
  .icon-btn:hover{background:#fee2e2;color:#dc2626}
  .overlay{position:fixed;inset:0;background:rgba(15,23,54,.55);backdrop-filter:blur(4px);display:none;align-items:center;justify-content:center;z-index:50;padding:16px}
  .overlay.show{display:flex}
  .modal{background:#fff;border-radius:18px;width:100%;max-width:480px;padding:26px;box-shadow:0 34px 70px -22px rgba(0,0,0,.55);animation:pop .22s cubic-bezier(.34,1.4,.64,1);max-height:92vh;overflow-y:auto}
  @keyframes pop{from{transform:scale(.95);opacity:0}to{transform:scale(1);opacity:1}}
  .modal h3{font-size:18px;font-weight:700;color:#0f1736;margin-bottom:4px}
  .modal .sub{font-size:13px;color:#8b93ad;margin-bottom:20px;line-height:1.5}
  .modal .warn-box{background:#fef2f2;border:1px solid #fecaca;border-radius:10px;padding:12px 14px;font-size:12.5px;color:#991b1b;line-height:1.5;margin-bottom:16px}
  .field{display:flex;flex-direction:column;gap:7px;margin-bottom:14px}
  .field label{font-size:11px;font-weight:700;color:#7b84a2;text-transform:uppercase;letter-spacing:.6px}
  .field input,.field select{padding:11px 13px;border:1px solid #e0e5ef;border-radius:10px;font-size:14px;font-family:inherit;outline:none;background:#fafbfe;transition:.15s;color:#0f1736}
  .field input:focus,.field select:focus{border-color:#6366f1;box-shadow:0 0 0 3px rgba(99,102,241,.13);background:#fff}
  .modal-actions{display:flex;gap:10px;justify-content:flex-end;margin-top:24px}
  .modal-actions .btn{padding:11px 20px;border-radius:10px;font-size:13.5px;font-weight:600;cursor:pointer;font-family:inherit;border:none;transition:.15s}
  .btn-cancel{background:#f1f4fa;color:#4b5573}.btn-cancel:hover{background:#e6ebf6}
  .btn-save{background:#6366f1;color:#fff}.btn-save:hover{background:#4f46e5;transform:translateY(-1px)}
  .btn-danger{background:#dc2626;color:#fff}.btn-danger:hover{background:#b91c1c;transform:translateY(-1px)}
  .imp-info{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:14px}
  .chip{background:#eef0ff;color:#4338ca;font-size:12px;font-weight:700;padding:7px 13px;border-radius:20px}
  .chip.warn{background:#fef3c7;color:#b45309}
  .chip.ok{background:#dcfce7;color:#15803d}
  .imp-preview{border:1px solid #e9edf5;border-radius:12px;max-height:240px;overflow:auto;margin-bottom:18px}
  .imp-preview table{font-size:12.5px}
  .imp-preview thead th{padding:9px 11px;font-size:10px}
  .imp-preview tbody td{padding:7px 11px}
  .mode-tabs{display:flex;gap:6px;margin-bottom:8px;background:#f1f4fa;padding:5px;border-radius:12px}
  .mode-tabs button{flex:1;border:none;background:transparent;padding:11px 12px;font-size:13px;font-weight:600;border-radius:9px;cursor:pointer;font-family:inherit;color:#5b637f;transition:.15s}
  .mode-tabs button.active{background:#fff;color:#4338ca;box-shadow:0 3px 10px -5px rgba(30,40,90,.35)}
  .drop-hint{position:fixed;inset:0;background:rgba(99,102,241,.9);color:#fff;display:none;align-items:center;justify-content:center;flex-direction:column;gap:14px;z-index:80;font-size:24px;font-weight:700;pointer-events:none}
  .drop-hint.show{display:flex}.drop-hint small{font-size:14px;font-weight:400;opacity:.9}.drop-hint .emoji{font-size:56px}
  .toast{position:fixed;bottom:26px;left:50%;transform:translate(-50%,80px);background:#0f1736;color:#fff;padding:13px 24px;border-radius:12px;font-size:13.5px;font-weight:600;opacity:0;transition:.3s;pointer-events:none;z-index:99;box-shadow:0 16px 34px -14px rgba(0,0,0,.7)}
  .toast.show{opacity:1;transform:translate(-50%,0)}
  .sync-badge{position:fixed;top:16px;right:16px;background:#fff;border:1px solid #e9edf5;border-radius:20px;padding:8px 14px;font-size:12px;font-weight:600;color:#4338ca;box-shadow:0 8px 20px -12px rgba(30,40,90,.4);z-index:70;display:flex;align-items:center;gap:8px}
  .sync-badge .pulse{width:8px;height:8px;border-radius:50%;background:#10b981;animation:pulse 2s infinite}
  .sync-badge.loading .pulse{background:#f59e0b}
  .sync-badge.error .pulse{background:#ef4444}
  @keyframes pulse{0%,100%{opacity:1}50%{opacity:.35}}
  @media(max-width:640px){.bar-row{grid-template-columns:100px 1fr 36px;font-size:12px}.toolbar .search{width:100%}.page-header h1{font-size:20px}.stat .value{font-size:24px}}
</style>
</head>
<body>

<div class="app">

  <!-- ============ SIDEBAR ============ -->
  <aside class="sidebar">
    <div class="brand">
      <div class="logo">📊</div>
      <div><div class="name">Dashboard</div><div class="sub">Marcas &amp; Responsáveis</div></div>
    </div>

    <div class="side-divider"></div>

    <div>
      <div class="nav-label">Ações principais</div>
      <div class="side-actions">
        <button class="side-btn primary" id="btnImport">
          <span class="ic">📥</span><span>Importar planilha</span>
        </button>
        <button class="side-btn" id="btnAdd">
          <span class="ic">➕</span><span>Adicionar marca</span>
        </button>
      </div>
    </div>

    <div>
      <div class="nav-label">Dados</div>
      <div class="side-actions">
        <button class="side-btn" id="btnExport">
          <span class="ic">⬇</span><span>Exportar CSV</span>
        </button>
        <button class="side-btn" id="btnRefresh">
          <span class="ic">🔄</span><span>Sincronizar agora</span>
        </button>
        <button class="side-btn danger" id="btnPurge">
          <span class="ic">🗑</span><span>Apagar tudo</span>
        </button>
      </div>
    </div>

    <div id="statusBox" class="status-box"><span class="dot"></span><span>Conectando…</span></div>

    <div class="side-info">
      <strong>☁️ Sincronizado na nuvem</strong>
      Todos os dados ficam no Cloudflare D1. O que você alterar aqui aparece nas outras máquinas em até 10 segundos.
    </div>

    <div class="side-footer">
      Cloudflare Workers + D1
    </div>
  </aside>

  <!-- ============ MAIN ============ -->
  <main class="main">

    <header class="page-header">
      <div>
        <h1>Visão geral</h1>
        <p>Acompanhe e edite os responsáveis por cada marca. Tudo sincronizado.</p>
      </div>
      <span class="badge" id="headerBadge">carregando…</span>
    </header>

    <section class="stats" id="stats"></section>

    <div class="charts-grid">
      <div class="section">
        <div class="section-head">
          <h2><span class="dot"></span> Marcas por responsável</h2>
          <small id="chartTotal"></small>
        </div>
        <div class="bars" id="chartResp"></div>
      </div>
      <div class="section">
        <div class="section-head">
          <h2><span class="dot" style="background:#0ea5e9"></span> Distribuição por setor</h2>
        </div>
        <div class="bars" id="chartSetor"></div>
      </div>
    </div>

    <section class="section">
      <div class="section-head">
        <h2><span class="dot" style="background:#10b981"></span> Lista de marcas</h2>
        <small>Clique no campo “Responsável” para editar — salva automaticamente</small>
      </div>

      <div class="toolbar">
        <input class="search" id="fSearch" type="text" placeholder="🔍 Buscar por marca, CNPJ ou responsável...">
        <select id="fSetor"><option value="">Todos os setores</option></select>
        <select id="fResp"><option value="">Todos os responsáveis</option></select>
        <select id="fSit">
          <option value="">Todas as situações</option>
          <option value="__empty">Sem situação</option>
          <option value="Pendente">Pendente</option>
          <option value="Em andamento">Em andamento</option>
          <option value="Concluído">Concluído</option>
        </select>
        <button class="clear-btn" id="btnClear">Limpar</button>
        <span class="count-pill" id="countPill">0 marcas</span>
      </div>

      <div class="table-scroll">
        <table>
          <thead>
            <tr>
              <th style="width:52px">#</th>
              <th style="width:210px">Marca</th>
              <th style="width:180px">CNPJ</th>
              <th style="width:130px">Setor</th>
              <th style="width:230px">Responsável ✏️</th>
              <th style="width:170px">Situação</th>
              <th style="width:52px" class="center">Ação</th>
            </tr>
          </thead>
          <tbody id="tbody"></tbody>
        </table>
      </div>
    </section>

  </main>
</div>

<div class="sync-badge" id="syncBadge">
  <span class="pulse"></span><span id="syncText">conectando…</span>
</div>

<datalist id="resp-list"></datalist>
<datalist id="setor-list"></datalist>
<input type="file" id="fileInput" accept=".xlsx,.xls,.csv" style="display:none">

<div class="drop-hint" id="dropHint">
  <div class="emoji">📥</div>
  <div>Solte a planilha para importar</div>
  <small>Formatos aceitos: .xlsx · .xls · .csv</small>
</div>

<!-- ============================================================
     MODAL: ADICIONAR MARCA
     ============================================================ -->
<div class="overlay" id="overlay">
  <div class="modal">
    <h3>Adicionar nova marca</h3>
    <p class="sub">Preencha os campos abaixo. Somente a Marca é obrigatória.</p>

    <div class="field">
      <label for="mMarca">Marca *</label>
      <input id="mMarca" type="text" placeholder="Ex.: Samsung">
    </div>
    <div class="field">
      <label for="mCnpj">CNPJ</label>
      <input id="mCnpj" type="text" placeholder="00.000.000/0000-00" inputmode="numeric">
    </div>
    <div class="field">
      <label for="mSetor">Setor</label>
      <input id="mSetor" type="text" list="setor-list" placeholder="Ex.: Informática">
    </div>
    <div class="field">
      <label for="mResp">Responsável</label>
      <input id="mResp" type="text" list="resp-list" placeholder="Digite ou selecione...">
    </div>
    <div class="field">
      <label for="mSit">Situação</label>
      <select id="mSit">
        <option value="">—</option>
        <option value="Pendente">Pendente</option>
        <option value="Em andamento">Em andamento</option>
        <option value="Concluído">Concluído</option>
      </select>
    </div>

    <div class="modal-actions">
      <button class="btn btn-cancel" id="btnCancel">Cancelar</button>
      <button class="btn btn-save" id="btnSave">Adicionar marca</button>
    </div>
  </div>
</div>

<!-- ============================================================
     MODAL: IMPORTAR PLANILHA
     ============================================================ -->
<div class="overlay" id="importOverlay">
  <div class="modal" style="max-width:740px">
    <h3>📥 Confirmar importação</h3>
    <p class="sub">Arquivo: <strong id="impFileName">—</strong></p>

    <div class="imp-info" id="impInfo"></div>
    <div class="imp-preview" id="impPreview"></div>

    <div class="field" style="margin-bottom:8px">
      <label>Como deseja importar?</label>
    </div>
    <div class="mode-tabs" id="impModeTabs">
      <button type="button" data-mode="append" class="active">➕ Adicionar aos existentes</button>
      <button type="button" data-mode="replace">🔄 Substituir todos os dados</button>
    </div>

    <div class="modal-actions">
      <button class="btn btn-cancel" id="impCancel">Cancelar</button>
      <button class="btn btn-save" id="impConfirm">Importar dados</button>
    </div>
  </div>
</div>

<!-- ============================================================
     MODAL: APAGAR TUDO (com senha)
     ============================================================ -->
<div class="overlay" id="purgeOverlay">
  <div class="modal" style="max-width:440px">
    <h3>🗑 Apagar todos os dados</h3>
    <p class="sub">Informe a senha para confirmar a exclusão.</p>

    <div class="warn-box">
      ⚠️ Esta ação é <b>permanente</b> e afeta <b>todas as máquinas</b> conectadas ao dashboard. Não há como desfazer.
    </div>

    <div class="field">
      <label for="purgePwd">Senha</label>
      <input id="purgePwd" type="password" placeholder="Digite a senha..." autocomplete="off">
    </div>

    <p id="purgeError"
       style="color:#dc2626;font-size:12.5px;font-weight:600;display:none;margin-bottom:8px"></p>

    <div class="modal-actions">
      <button class="btn btn-cancel" id="purgeCancel">Cancelar</button>
      <button class="btn btn-danger" id="purgeConfirm">Apagar tudo</button>
    </div>
  </div>
</div>

<div class="toast" id="toast"></div>

<script>
/* =========================================================
   SEED — cole aqui a MESMA constante RAW dos arquivos anteriores
   Formato: "Marca|CNPJ|Setor" por linha
   ========================================================= */
const RAW = `
/* ==== COLE AQUI A CONST RAW ==== */
`;

/* =========================================================
   CONFIG / HELPERS
   ========================================================= */
const API = '/api';
const POLL_MS = 10000;
const SITUACOES = ['', 'Pendente', 'Em andamento', 'Concluído'];
const PALETTE   = ['#6366f1','#0ea5e9','#10b981','#f59e0b','#ef4444','#8b5cf6',
                   '#ec4899','#14b8a6','#f97316','#4f46e5','#84cc16','#06b6d4'];

let rows = [];
let lastSyncJson = '';

const $   = s => document.querySelector(s);
const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const fmtCnpj = v => {
  if(!v) return '';
  const d = String(v).replace(/\.0+$/,'').replace(/\D/g,'');
  return d.length === 14
    ? d.replace(/^(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})$/, '$1.$2.$3/$4-$5')
    : String(v);
};
const cleanCnpj = v => {
  if(v == null) return '';
  let s = String(v).trim();
  if(!s) return '';
  if(/^-?\d+(\.\d+)?[eE][+-]?\d+$/.test(s)){
    const n = Number(s); if(isFinite(n)) s = n.toFixed(0);
  }
  return s.replace(/\.0+$/, '');
};
const tagSetor = s => {
  const t = (s || '').trim();
  const l = t.toLowerCase();
  let c = 'outro';
  if(l.includes('inform')) c = 'informatica';
  else if(l.includes('coffe') || l.includes('break')) c = 'coffe';
  else if(l.includes('área') || l.includes('area')) c = 'area';
  return `<span class="tag ${c}">${esc(t || '—')}</span>`;
};
const toast = m => {
  const t = $('#toast');
  t.textContent = m;
  t.classList.add('show');
  clearTimeout(t._tm);
  t._tm = setTimeout(() => t.classList.remove('show'), 2400);
};
const normHeader = h => String(h ?? '')
  .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
  .toLowerCase().trim().replace(/[^a-z0-9]/g, '');

/* =========================================================
   API
   ========================================================= */
async function apiGet(path){
  const r = await fetch(API + path);
  if(!r.ok) throw new Error('HTTP ' + r.status);
  return r.json();
}
async function apiSend(method, path, body){
  const r = await fetch(API + path, {
    method,
    headers: { 'Content-Type': 'application/json' },
    body: body ? JSON.stringify(body) : undefined
  });
  if(!r.ok){ const t = await r.text(); throw new Error('HTTP ' + r.status + ' ' + t); }
  return r.json();
}
const apiPost = (p, b) => apiSend('POST', p, b);
const apiPut  = (p, b) => apiSend('PUT', p, b);
const apiDelete = p => apiSend('DELETE', p);

/* =========================================================
   STATUS
   ========================================================= */
function setStatus(kind, text){
  const box = $('#statusBox');
  box.className = 'status-box' + (kind === 'ok' ? '' : ' ' + kind);
  box.innerHTML = `<span class="dot"></span><span>${text}</span>`;

  const badge = $('#syncBadge');
  badge.className = 'sync-badge' + (kind === 'err' ? ' error' : kind === 'warn' ? ' loading' : '');
  $('#syncText').textContent = text;
}

/* =========================================================
   CARREGAR + SEED
   ========================================================= */
async function loadRows(){
  try{
    const data = await apiGet('/rows');
    if(!Array.isArray(data)) throw new Error('resposta inválida');

    // Banco vazio → faz seed com RAW (apenas uma vez)
    if(data.length === 0 && RAW.trim().length > 0){
      await seedDatabase();
      return loadRows();
    }

    rows = data;
    lastSyncJson = JSON.stringify(rows);
    setStatus('ok', 'sincronizado');
    refreshAll();
  }catch(err){
    console.error(err);
    setStatus('err', 'sem conexão com a API');
    toast('Não consegui carregar os dados: ' + err.message);
  }
}

async function seedDatabase(){
  const seed = RAW.replace(/\r/g,'').trim().split('\n').filter(Boolean).map(l => {
    const p = l.split('|');
    return { marca: p[0]||'', cnpj: p[1]||'', setor: p[2]||'', responsavel: '', situacao: '' };
  });
  if(!seed.length) return;
  await apiPost('/rows/bulk', { rows: seed, replace: true });
  toast(`Base inicial criada com ${seed.length} registros`);
}

/* =========================================================
   POLLING
   ========================================================= */
function isEditingCell(){
  const el = document.activeElement;
  return el && el.closest && el.closest('#tbody');
}

async function pollUpdates(){
  if(isEditingCell()) return;
  try{
    const data = await apiGet('/rows');
    const json = JSON.stringify(data);
    if(json !== lastSyncJson){
      rows = data;
      lastSyncJson = json;
      refreshAll();
    }
    setStatus('ok', 'sincronizado');
  }catch(err){
    setStatus('err', 'offline');
  }
}
setInterval(pollUpdates, POLL_MS);

/* =========================================================
   FILTROS
   ========================================================= */
const filters = { search:'', setor:'', resp:'', sit:'' };

function filtered(){
  const q = filters.search.toLowerCase().trim();
  return rows.filter(r => {
    if(filters.setor && (r.setor||'') !== filters.setor) return false;
    if(filters.resp){
      const resp = (r.responsavel||'').trim();
      if(filters.resp === '__empty'){ if(resp) return false; }
      else if(resp !== filters.resp) return false;
    }
    if(filters.sit){
      const s = r.situacao || '';
      if(filters.sit === '__empty'){ if(s) return false; }
      else if(s !== filters.sit) return false;
    }
    if(q){
      const blob = `${r.marca} ${r.cnpj} ${r.setor} ${r.responsavel}`.toLowerCase();
      if(!blob.includes(q)) return false;
    }
    return true;
  });
}

/* =========================================================
   KPIs / GRÁFICOS
   ========================================================= */
function renderStats(){
  const total = rows.length;
  const comResp = rows.filter(r => (r.responsavel||'').trim()).length;
  const semResp = total - comResp;
  const distintos = new Set(rows.map(r => (r.responsavel||'').trim()).filter(Boolean)).size;
  const pct = total ? Math.round(comResp / total * 100) : 0;

  $('#stats').innerHTML = `
    <div class="stat" style="--c:#6366f1">
      <span class="label">Total de marcas</span><span class="value">${total}</span>
      <span class="sub">registros na base</span>
    </div>
    <div class="stat" style="--c:#10b981">
      <span class="label">Com responsável</span><span class="value">${comResp}</span>
      <span class="sub">${pct}% atribuído</span>
    </div>
    <div class="stat" style="--c:#f59e0b">
      <span class="label">Sem responsável</span><span class="value">${semResp}</span>
      <span class="sub">aguardando atribuição</span>
    </div>
    <div class="stat" style="--c:#0ea5e9">
      <span class="label">Responsáveis</span><span class="value">${distintos}</span>
      <span class="sub">pessoas/equipes</span>
    </div>`;

  $('#headerBadge').textContent = `${total} marcas`;
}

function renderChartResp(){
  const counts = {};
  rows.forEach(r => {
    const k = (r.responsavel||'').trim() || 'Sem responsável';
    counts[k] = (counts[k] || 0) + 1;
  });
  const entries = Object.entries(counts).sort((a,b) => b[1]-a[1] || a[0].localeCompare(b[0],'pt-BR'));
  const box = $('#chartResp');
  if(!entries.length){ box.innerHTML = '<p class="empty-msg">Nenhum dado disponível.</p>'; return; }
  const max = Math.max(...entries.map(e => e[1]));
  $('#chartTotal').textContent = entries.length + ' responsável(is)';
  box.innerHTML = entries.map(([nome, qtd], i) => {
    const sem = nome === 'Sem responsável';
    const color = sem ? '#cbd5e1' : PALETTE[i % PALETTE.length];
    const w = max ? (qtd / max * 100) : 0;
    return `<div class="bar-row">
      <div class="bar-name" title="${esc(nome)}">${esc(nome)}</div>
      <div class="bar-track"><div class="bar-fill" style="width:${w}%;background:${color}"></div></div>
      <div class="bar-val">${qtd}</div>
    </div>`;
  }).join('');
}

function renderChartSetor(){
  const counts = {};
  rows.forEach(r => {
    const k = (r.setor||'').trim() || 'Não informado';
    counts[k] = (counts[k] || 0) + 1;
  });
  const entries = Object.entries(counts).sort((a,b) => b[1]-a[1]);
  const box = $('#chartSetor');
  if(!entries.length){ box.innerHTML = '<p class="empty-msg">Nenhum dado disponível.</p>'; return; }
  const max = Math.max(...entries.map(e => e[1]));
  box.innerHTML = entries.map(([nome, qtd], i) => {
    const low = nome.toLowerCase();
    let color = PALETTE[(i + 3) % PALETTE.length];
    if(low.includes('inform')) color = '#0ea5e9';
    else if(low.includes('coffe') || low.includes('break')) color = '#f59e0b';
    else if(low.includes('área') || low.includes('area')) color = '#8b5cf6';
    const w = max ? (qtd / max * 100) : 0;
    return `<div class="bar-row">
      <div class="bar-name" title="${esc(nome)}">${esc(nome)}</div>
      <div class="bar-track"><div class="bar-fill" style="width:${w}%;background:${color}"></div></div>
      <div class="bar-val">${qtd}</div>
    </div>`;
  }).join('');
}

/* =========================================================
   DATALISTS + FILTROS
   ========================================================= */
const getResponsaveis = () => [...new Set(rows.map(r => (r.responsavel||'').trim()).filter(Boolean))]
  .sort((a,b) => a.localeCompare(b, 'pt-BR'));
const getSetores = () => [...new Set(rows.map(r => (r.setor||'').trim()).filter(Boolean))]
  .sort((a,b) => a.localeCompare(b, 'pt-BR'));

function renderDatalist(){
  $('#resp-list').innerHTML = getResponsaveis().map(r => `<option value="${esc(r)}"></option>`).join('');
  $('#setor-list').innerHTML = getSetores().map(s => `<option value="${esc(s)}"></option>`).join('');
}

function renderFilterOptions(){
  const s1 = $('#fSetor'); const c1 = s1.value;
  s1.innerHTML = '<option value="">Todos os setores</option>' +
    getSetores().map(s => `<option value="${esc(s)}">${esc(s)}</option>`).join('');
  s1.value = c1;

  const s2 = $('#fResp'); const c2 = s2.value;
  const semQtd = rows.filter(r => !(r.responsavel||'').trim()).length;
  s2.innerHTML = '<option value="">Todos os responsáveis</option>' +
    (semQtd ? `<option value="__empty">— Sem responsável (${semQtd})</option>` : '') +
    getResponsaveis().map(r => {
      const q = rows.filter(x => (x.responsavel||'').trim() === r).length;
      return `<option value="${esc(r)}">${esc(r)} (${q})</option>`;
    }).join('');
  s2.value = c2;

  $('#fSit').value = filters.sit;
  $('#fSetor').value = filters.setor;
  $('#fResp').value  = filters.resp;
}

/* =========================================================
   TABELA
   ========================================================= */
function renderTable(){
  const list = filtered();
  $('#countPill').textContent = `${list.length} de ${rows.length} marcas`;

  if(!list.length){
    $('#tbody').innerHTML = `<tr><td colspan="7">
      <p class="empty-msg">Nenhuma marca encontrada com os filtros atuais.</p>
    </td></tr>`;
    return;
  }

  $('#tbody').innerHTML = list.map((r, i) => `
    <tr>
      <td class="num">${i + 1}</td>
      <td class="marca">${esc(r.marca)}</td>
      <td class="cnpj">${esc(fmtCnpj(r.cnpj)) || '<span style="color:#c7ccdd">—</span>'}</td>
      <td>${tagSetor(r.setor)}</td>
      <td>
        <input class="cell-input" type="text" list="resp-list"
               data-id="${r.id}" value="${esc(r.responsavel)}"
               placeholder="Digite o responsável...">
      </td>
      <td>
        <select class="cell-input sit" data-id="${r.id}">
          ${SITUACOES.map(s =>
            `<option value="${esc(s)}"${r.situacao === s ? ' selected' : ''}>${s || '—'}</option>`
          ).join('')}
        </select>
      </td>
      <td class="center">
        <button class="icon-btn" data-del="${r.id}" title="Excluir">✕</button>
      </td>
    </tr>`).join('');
}

function refreshAll(){
  renderStats();
  renderChartResp();
  renderChartSetor();
  renderDatalist();
  renderFilterOptions();
  renderTable();
}

/* =========================================================
   EDIÇÃO INLINE — salva via API
   ========================================================= */
$('#tbody').addEventListener('change', async e => {
  const t = e.target;
  const id = t.dataset.id;
  if(!id) return;

  const row = rows.find(r => String(r.id) === String(id));
  if(!row) return;

  const original = { ...row };
  if(t.tagName === 'INPUT') row.responsavel = t.value.trim();
  else if(t.tagName === 'SELECT') row.situacao = t.value;

  try{
    setStatus('warn', 'salvando…');
    const updated = await apiPut('/rows/' + id, {
      marca: row.marca, cnpj: row.cnpj, setor: row.setor,
      responsavel: row.responsavel, situacao: row.situacao
    });
    Object.assign(row, updated);
    lastSyncJson = JSON.stringify(rows);
    renderStats();
    renderChartResp();
    renderDatalist();
    renderFilterOptions();
    setStatus('ok', 'sincronizado');
  }catch(err){
    console.error(err);
    Object.assign(row, original);
    if(t.tagName === 'INPUT') t.value = original.responsavel || '';
    if(t.tagName === 'SELECT') t.value = original.situacao || '';
    setStatus('err', 'erro ao salvar');
    toast('Não foi possível salvar: ' + err.message);
  }
});

/* =========================================================
   EXCLUIR 1 REGISTRO
   ========================================================= */
$('#tbody').addEventListener('click', async e => {
  const btn = e.target.closest('[data-del]');
  if(!btn) return;
  const id = btn.dataset.del;
  const row = rows.find(r => String(r.id) === String(id));
  if(!row) return;
  if(!confirm(`Excluir a marca "${row.marca}"?`)) return;
  try{
    setStatus('warn', 'excluindo…');
    await apiDelete('/rows/' + id);
    rows = rows.filter(r => String(r.id) !== String(id));
    lastSyncJson = JSON.stringify(rows);
    refreshAll();
    setStatus('ok', 'sincronizado');
    toast('Marca removida');
  }catch(err){
    setStatus('err', 'erro ao excluir');
    toast('Erro: ' + err.message);
  }
});

/* =========================================================
   FILTROS — EVENTOS
   ========================================================= */
$('#fSearch').addEventListener('input', e => { filters.search = e.target.value; renderTable(); });
$('#fSetor').addEventListener('change', e => { filters.setor = e.target.value; renderTable(); });
$('#fResp').addEventListener('change',  e => { filters.resp  = e.target.value; renderTable(); });
$('#fSit').addEventListener('change',   e => { filters.sit   = e.target.value; renderTable(); });
$('#btnClear').addEventListener('click', () => {
  filters.search = filters.setor = filters.resp = filters.sit = '';
  $('#fSearch').value = '';
  $('#fSetor').value = '';
  $('#fResp').value = '';
  $('#fSit').value = '';
  renderTable();
});

/* =========================================================
   MODAL ADICIONAR
   ========================================================= */
const overlay = $('#overlay');
function openModal(){
  $('#mMarca').value = '';
  $('#mCnpj').value  = '';
  $('#mSetor').value = '';
  $('#mResp').value  = '';
  $('#mSit').value   = '';
  overlay.classList.add('show');
  setTimeout(() => $('#mMarca').focus(), 60);
}
function closeModal(){ overlay.classList.remove('show'); }

$('#btnAdd').addEventListener('click', openModal);
$('#btnCancel').addEventListener('click', closeModal);
overlay.addEventListener('click', e => { if(e.target === overlay) closeModal(); });

$('#mCnpj').addEventListener('input', e => {
  const v = e.target.value.replace(/\D/g, '').slice(0, 14);
  let out = v;
  if(v.length > 12)      out = `${v.slice(0,2)}.${v.slice(2,5)}.${v.slice(5,8)}/${v.slice(8,12)}-${v.slice(12)}`;
  else if(v.length > 8)  out = `${v.slice(0,2)}.${v.slice(2,5)}.${v.slice(5,8)}/${v.slice(8)}`;
  else if(v.length > 5)  out = `${v.slice(0,2)}.${v.slice(2,5)}.${v.slice(5)}`;
  else if(v.length > 2)  out = `${v.slice(0,2)}.${v.slice(2)}`;
  e.target.value = out;
});

$('#btnSave').addEventListener('click', async () => {
  const marca = $('#mMarca').value.trim();
  if(!marca){ $('#mMarca').focus(); toast('Informe o nome da marca'); return; }
  try{
    setStatus('warn', 'adicionando…');
    const created = await apiPost('/rows', {
      marca,
      cnpj: $('#mCnpj').value.trim(),
      setor: $('#mSetor').value.trim(),
      responsavel: $('#mResp').value.trim(),
      situacao: $('#mSit').value
    });
    rows.push(created);
    lastSyncJson = JSON.stringify(rows);
    refreshAll();
    closeModal();
    setStatus('ok', 'sincronizado');
    toast(`Marca "${marca}" adicionada`);
  }catch(err){
    setStatus('err', 'erro ao adicionar');
    toast('Erro: ' + err.message);
  }
});

/* =========================================================
   IMPORTAR PLANILHA
   ========================================================= */
const importOverlay = $('#importOverlay');
let pendingImport = null;
let importMode = 'append';

function detectColumns(headerRow){
  const headers = (headerRow || []).map(h => String(h ?? ''));
  const find = a => headers.findIndex(h => a.includes(normHeader(h)));
  return {
    marca:       find(['marca','brand','nome','produto','empresa']),
    cnpj:        find(['cnpj','cnpjcpf','documento','cpfcnpj','cnpjdoc']),
    setor:       find(['setor','segmento','categoria','area']),
    responsavel: find(['responsavel','owner','encarregado','resp']),
    situacao:    find(['situacao','status','estado'])
  };
}

function parseGrid(grid){
  if(!Array.isArray(grid) || !grid.length) return [];
  let headerIdx = -1, cols = null;
  for(let i = 0; i < Math.min(grid.length, 10); i++){
    const c = detectColumns(grid[i] || []);
    if(c.marca !== -1){ headerIdx = i; cols = c; break; }
  }
  if(headerIdx === -1){
    cols = { marca:1, cnpj:2, setor:3, responsavel:4, situacao:5 };
  }
  const out = [];
  for(let i = headerIdx + 1; i < grid.length; i++){
    const row = grid[i] || [];
    const marca = String(row[cols.marca] ?? '').trim();
    if(!marca) continue;
    out.push({
      marca,
      cnpj:        cols.cnpj >= 0 ? cleanCnpj(row[cols.cnpj]) : '',
      setor:       cols.setor >= 0 ? String(row[cols.setor] ?? '').trim() : '',
      responsavel: cols.responsavel >= 0 ? String(row[cols.responsavel] ?? '').trim() : '',
      situacao:    cols.situacao >= 0 ? String(row[cols.situacao] ?? '').trim() : ''
    });
  }
  return out;
}

function parseCSV(text){
  text = text.replace(/^\uFEFF/, '');
  const rows = [];
  let cur = [], f = '', q = false;
  for(let i = 0; i < text.length; i++){
    const c = text[i];
    if(q){
      if(c === '"' && text[i+1] === '"'){ f += '"'; i++; }
      else if(c === '"'){ q = false; }
      else f += c;
    } else {
      if(c === '"') q = true;
      else if(c === ';' || c === ','){ cur.push(f); f = ''; }
      else if(c === '\n'){ cur.push(f); rows.push(cur); cur = []; f = ''; }
      else if(c !== '\r') f += c;
    }
  }
  if(f !== '' || cur.length){ cur.push(f); rows.push(cur); }
  return rows;
}

async function readFileToGrid(file){
  const isCSV = /\.(csv|txt)$/i.test(file.name);
  if(typeof XLSX !== 'undefined'){
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type:'array', cellText:true, cellDates:false });
    const ws = wb.Sheets[wb.SheetNames[0]];
    return XLSX.utils.sheet_to_json(ws, { header:1, raw:false, defval:'' });
  }
  if(isCSV) return parseCSV(await file.text());
  throw new Error('Biblioteca XLSX não disponível.');
}

function renderImportPreview(){
  if(!pendingImport) return;
  const total = pendingImport.length;
  const comCnpj = pendingImport.filter(r => r.cnpj).length;
  const comResp = pendingImport.filter(r => r.responsavel).length;

  $('#impInfo').innerHTML = `
    <span class="chip ok">${total} registros</span>
    <span class="chip">${comCnpj} com CNPJ</span>
    <span class="chip${comResp === 0 ? ' warn' : ''}">${comResp} com responsável</span>`;

  const head = `<thead><tr>
    <th style="width:36px">#</th><th>Marca</th><th>CNPJ</th><th>Setor</th>
    <th>Responsável</th><th>Situação</th>
  </tr></thead>`;

  const body = pendingImport.slice(0, 8).map((r, i) => `
    <tr>
      <td class="num">${i+1}</td>
      <td class="marca">${esc(r.marca)}</td>
      <td class="cnpj">${esc(fmtCnpj(r.cnpj)) || '—'}</td>
      <td>${esc(r.setor || '—')}</td>
      <td>${esc(r.responsavel || '—')}</td>
      <td>${esc(r.situacao || '—')}</td>
    </tr>`).join('');

  const mais = pendingImport.length > 8
    ? `<tr><td colspan="6" style="text-align:center;color:#9aa2bd;padding:9px;font-size:12px">
        … e mais ${pendingImport.length - 8} registro(s)
      </td></tr>` : '';

  $('#impPreview').innerHTML = `<table>${head}<tbody>${body}${mais}</tbody></table>`;
}

async function handleFile(file){
  if(!file) return;
  try{
    const grid = await readFileToGrid(file);
    const parsed = parseGrid(grid);
    if(!parsed.length){ toast('Nenhum registro válido encontrado'); return; }
    pendingImport = parsed;
    $('#impFileName').textContent = file.name;
    setImportMode('append');
    renderImportPreview();
    importOverlay.classList.add('show');
  }catch(err){
    console.error(err);
    toast('Erro ao ler o arquivo: ' + err.message);
  }
}

function setImportMode(mode){
  importMode = mode;
  document.querySelectorAll('#impModeTabs button')
    .forEach(b => b.classList.toggle('active', b.dataset.mode === mode));
}
document.querySelectorAll('#impModeTabs button')
  .forEach(b => b.addEventListener('click', () => setImportMode(b.dataset.mode)));

$('#btnImport').addEventListener('click', () => {
  const i = $('#fileInput'); i.value = ''; i.click();
});
$('#fileInput').addEventListener('change', e => {
  const f = e.target.files && e.target.files[0];
  if(f) handleFile(f);
});

$('#impCancel').addEventListener('click', () => {
  importOverlay.classList.remove('show');
  pendingImport = null;
});
importOverlay.addEventListener('click', e => {
  if(e.target === importOverlay){
    importOverlay.classList.remove('show');
    pendingImport = null;
  }
});

$('#impConfirm').addEventListener('click', async () => {
  if(!pendingImport || !pendingImport.length) return;
  try{
    setStatus('warn', 'importando…');
    const res = await apiPost('/rows/bulk', {
      rows: pendingImport,
      replace: importMode === 'replace'
    });
    importOverlay.classList.remove('show');
    await loadRows();
    toast(`${res.inserted} registros importados`);
  }catch(err){
    setStatus('err', 'erro ao importar');
    toast('Erro: ' + err.message);
  }
});

/* Drag & drop */
const dropHint = $('#dropHint');
let dragDepth = 0;
document.addEventListener('dragenter', e => {
  if(e.dataTransfer && [...(e.dataTransfer.types || [])].includes('Files')){
    dragDepth++;
    dropHint.classList.add('show');
  }
});
document.addEventListener('dragleave', () => {
  dragDepth = Math.max(0, dragDepth - 1);
  if(dragDepth === 0) dropHint.classList.remove('show');
});
document.addEventListener('dragover', e => {
  if(e.dataTransfer) e.dataTransfer.dropEffect = 'copy';
  e.preventDefault();
});
document.addEventListener('drop', e => {
  e.preventDefault();
  dragDepth = 0;
  dropHint.classList.remove('show');
  const f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
  if(f) handleFile(f);
});

/* =========================================================
   APAGAR TUDO — com senha (validada no servidor)
   ========================================================= */
const purgeOverlay = $('#purgeOverlay');
const purgePwd     = $('#purgePwd');
const purgeError   = $('#purgeError');

function openPurge(){
  purgePwd.value = '';
  purgeError.style.display = 'none';
  purgeOverlay.classList.add('show');
  setTimeout(() => purgePwd.focus(), 60);
}
function closePurge(){ purgeOverlay.classList.remove('show'); }

$('#btnPurge').addEventListener('click', openPurge);
$('#purgeCancel').addEventListener('click', closePurge);
purgeOverlay.addEventListener('click', e => {
  if(e.target === purgeOverlay) closePurge();
});

purgePwd.addEventListener('keydown', e => {
  if(e.key === 'Enter'){ e.preventDefault(); $('#purgeConfirm').click(); }
});

$('#purgeConfirm').addEventListener('click', async () => {
  const senha = purgePwd.value;
  if(!senha){
    purgeError.textContent = 'Digite a senha.';
    purgeError.style.display = 'block';
    return;
  }

  try{
    setStatus('warn', 'apagando…');
    const res = await fetch(API + '/rows/purge', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ password: senha })
    });

    if(res.status === 401){
      purgeError.textContent = 'Senha incorreta.';
      purgeError.style.display = 'block';
      purgePwd.select();
      setStatus('ok', 'sincronizado');
      return;
    }
    if(!res.ok){
      const t = await res.text();
      throw new Error('HTTP ' + res.status + ' ' + t);
    }

    const data = await res.json();
    rows = [];
    lastSyncJson = JSON.stringify(rows);
    refreshAll();
    closePurge();
    setStatus('ok', 'sincronizado');
    toast(`Base apagada (${data.deleted ?? 0} registros removidos)`);
  }catch(err){
    console.error(err);
    setStatus('err', 'erro ao apagar');
    purgeError.textContent = 'Erro: ' + err.message;
    purgeError.style.display = 'block';
  }
});

/* =========================================================
   ESC fecha todos os modais
   ========================================================= */
document.addEventListener('keydown', e => {
  if(e.key === 'Escape'){
    closeModal();
    importOverlay.classList.remove('show');
    if(purgeOverlay) purgeOverlay.classList.remove('show');
  }
});

/* =========================================================
   EXPORTAR CSV / SINCRONIZAR
   ========================================================= */
$('#btnExport').addEventListener('click', () => {
  const linhas = filtered();
  const head = ['#','MARCA','CNPJ','SETOR','RESPONSAVEL','SITUACAO'];
  const csv = [head.join(';')].concat(
    linhas.map((r, i) => [
      i + 1,
      `"${(r.marca||'').replace(/"/g,'""')}"`,
      `"${fmtCnpj(r.cnpj)}"`,
      `"${(r.setor||'').replace(/"/g,'""')}"`,
      `"${(r.responsavel||'').replace(/"/g,'""')}"`,
      `"${(r.situacao||'').replace(/"/g,'""')}"`
    ].join(';'))
  ).join('\n');

  const blob = new Blob(['\uFEFF' + csv], { type:'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `marcas_responsaveis_${new Date().toISOString().slice(0,10)}.csv`;
  a.click();
  URL.revokeObjectURL(url);
  toast(`${linhas.length} registros exportados`);
});

$('#btnRefresh').addEventListener('click', async () => {
  setStatus('warn', 'sincronizando…');
  await loadRows();
  toast('Sincronizado');
});

/* =========================================================
   INIT
   ========================================================= */
loadRows();
</script>
</body>
</html>
