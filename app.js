/* =========================================================
   Entregas Picolé — app pessoal (dados salvos no navegador)
   ========================================================= */

const STORAGE_KEY = 'picole_data_v1';
const SYNC_CONFIG_KEY = 'picole_sync_config';
const LAST_SYNC_KEY = 'picole_last_sync';
const SHEETJS_URL = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
const FIREBASE_APP_URL = 'https://www.gstatic.com/firebasejs/12.15.0/firebase-app-compat.js';
const FIREBASE_FIRESTORE_URL = 'https://www.gstatic.com/firebasejs/12.15.0/firebase-firestore-compat.js';

const FERIADOS_PADRAO = [
  ['2026-01-01', 'Confraternização Universal'],
  ['2026-02-16', 'Carnaval (segunda-feira)'],
  ['2026-02-17', 'Carnaval (terça-feira)'],
  ['2026-04-03', 'Sexta-feira Santa'],
  ['2026-04-21', 'Tiradentes'],
  ['2026-05-01', 'Dia do Trabalho'],
  ['2026-06-04', 'Corpus Christi'],
  ['2026-09-07', 'Independência do Brasil'],
  ['2026-10-12', 'Nossa Senhora Aparecida'],
  ['2026-11-02', 'Finados'],
  ['2026-11-15', 'Proclamação da República'],
  ['2026-11-20', 'Dia da Consciência Negra'],
  ['2026-12-25', 'Natal'],
];

function uid(prefix) {
  return prefix + '_' + Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
}

/* ---------------- Data store ---------------- */
let DB = null;

function defaultDB() {
  return {
    clientes: [],
    feriados: FERIADOS_PADRAO.map(([data, descricao]) => ({ id: uid('f'), data, descricao })),
    entregas: [],
  };
}

function loadDB() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) {
      DB = JSON.parse(raw);
      // garante que todas as chaves existam mesmo em bases antigas
      if (!DB.clientes) DB.clientes = [];
      if (!DB.feriados) DB.feriados = [];
      if (!DB.entregas) DB.entregas = [];
      return;
    }
  } catch (e) {
    console.error('Falha ao ler dados salvos', e);
  }
  DB = defaultDB();
  saveDB();
}

function saveDB() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(DB));
  agendarPushNuvem();
}

/* ---------------- Sincronização na nuvem (opcional) ---------------- */
let SYNC_CONFIG = null; // { firebaseConfig: {...}, syncCode: '...' } ou null
let firebaseReady = null;
let pushTimer = null;

function carregarConfigSync() {
  try {
    const raw = localStorage.getItem(SYNC_CONFIG_KEY);
    SYNC_CONFIG = raw ? JSON.parse(raw) : null;
  } catch (e) {
    SYNC_CONFIG = null;
  }
}

function ensureFirebase() {
  if (!SYNC_CONFIG) return Promise.reject(new Error('Sincronização não configurada'));
  if (firebaseReady) return firebaseReady;
  firebaseReady = new Promise((resolve, reject) => {
    const s1 = document.createElement('script');
    s1.src = FIREBASE_APP_URL;
    s1.onload = () => {
      const s2 = document.createElement('script');
      s2.src = FIREBASE_FIRESTORE_URL;
      s2.onload = () => {
        try {
          if (!firebase.apps || firebase.apps.length === 0) {
            firebase.initializeApp(SYNC_CONFIG.firebaseConfig);
          }
          resolve();
        } catch (err) {
          reject(err);
        }
      };
      s2.onerror = () => reject(new Error('Falha ao carregar Firestore (sem internet?)'));
      document.head.appendChild(s2);
    };
    s1.onerror = () => reject(new Error('Falha ao carregar Firebase (sem internet?)'));
    document.head.appendChild(s1);
  });
  return firebaseReady;
}

function parseFirebaseConfigText(text) {
  const get = (key) => {
    const m = text.match(new RegExp(key + '\\s*:\\s*["\']([^"\']*)["\']'));
    return m ? m[1] : '';
  };
  return {
    apiKey: get('apiKey'),
    authDomain: get('authDomain'),
    projectId: get('projectId'),
    storageBucket: get('storageBucket'),
    messagingSenderId: get('messagingSenderId'),
    appId: get('appId'),
  };
}

function gerarCodigoSync() {
  const chars = 'ABCDEFGHJKMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789';
  let s = '';
  for (let i = 0; i < 28; i++) s += chars[Math.floor(Math.random() * chars.length)];
  return s;
}

function agendarPushNuvem() {
  if (!SYNC_CONFIG) return;
  clearTimeout(pushTimer);
  pushTimer = setTimeout(() => { pushParaNuvem(); }, 1500);
}

async function pushParaNuvem() {
  // usado apos uma edicao local: sempre manda o que esta aqui pra nuvem,
  // sem comparar com o que ja esta la (a edicao que a pessoa acabou de
  // fazer neste aparelho tem prioridade sobre o que estiver na nuvem).
  if (!SYNC_CONFIG) return;
  try {
    await ensureFirebase();
    const now = Date.now();
    DB._updatedAt = now;
    const ref = firebase.firestore().collection('sincronizacoes').doc(SYNC_CONFIG.syncCode);
    await ref.set({ json: JSON.stringify(DB), updatedAt: now });
    localStorage.setItem(LAST_SYNC_KEY, String(now));
    if (typeof renderSyncStatus === 'function') renderSyncStatus();
  } catch (err) {
    console.error('push para nuvem falhou', err);
  }
}

async function sincronizar(silencioso) {
  // usado ao abrir o app e no botao "Sincronizar agora": compara com a
  // nuvem e traz dados mais novos de outro aparelho, se houver.
  if (!SYNC_CONFIG) return;
  if (!silencioso) toast('Sincronizando…');
  try {
    await ensureFirebase();
    const ref = firebase.firestore().collection('sincronizacoes').doc(SYNC_CONFIG.syncCode);
    const snap = await ref.get();
    const cloudData = snap.exists ? snap.data() : null;
    const localUpdatedAt = DB._updatedAt || 0;

    if (cloudData && cloudData.updatedAt > localUpdatedAt) {
      // nuvem tem versao mais nova (outro aparelho sincronizou depois) -> traz pra local
      DB = JSON.parse(cloudData.json);
      localStorage.setItem(STORAGE_KEY, JSON.stringify(DB));
      if (typeof refreshAll === 'function') refreshAll();
      if (!silencioso) toast('Dados atualizados a partir da nuvem.');
    } else {
      // local esta em dia ou mais novo -> manda pra nuvem
      const now = Date.now();
      DB._updatedAt = now;
      await ref.set({ json: JSON.stringify(DB), updatedAt: now });
      localStorage.setItem(STORAGE_KEY, JSON.stringify(DB));
      if (!silencioso) toast(cloudData ? 'Dados enviados para a nuvem.' : 'Primeira sincronização concluída.');
    }
    localStorage.setItem(LAST_SYNC_KEY, String(Date.now()));
  } catch (err) {
    console.error('sincronizar falhou', err);
    if (!silencioso) toast('Não foi possível sincronizar. Verifique sua internet.');
  }
  if (typeof renderSyncStatus === 'function') renderSyncStatus();
}

function sincronizarAgora() {
  sincronizar(false);
}

/* ---------------- Utilidades de data ---------------- */
function parseDate(s) {
  // 'YYYY-MM-DD' -> Date local (meio-dia, evita problema de fuso na virada do dia)
  if (!s) return null;
  const [y, m, d] = s.split('-').map(Number);
  return new Date(y, m - 1, d, 12, 0, 0);
}
function dateToKey(d) {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}
function todayStr() {
  return dateToKey(new Date());
}
function fmtDateBR(s) {
  if (!s) return '';
  const d = parseDate(s);
  return d.toLocaleDateString('pt-BR');
}
const MESES_ABREV = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez'];
function fmtMesAno(s) {
  const d = parseDate(s);
  return `${MESES_ABREV[d.getMonth()]}/${d.getFullYear()}`;
}
function fmtMoney(v) {
  if (v === null || v === undefined || isNaN(v)) v = 0;
  return v.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
}
function addDaysStr(s, n) {
  const d = parseDate(s);
  d.setDate(d.getDate() + n);
  return dateToKey(d);
}

/* ---------------- Dias úteis (domingo e feriados não contam; sábado conta) ---------------- */
function holidaySet() {
  return new Set(DB.feriados.map(f => f.data));
}
function isValidDay(date, hset) {
  if (date.getDay() === 0) return false; // domingo
  if (hset.has(dateToKey(date))) return false;
  return true;
}
function networkDays(startKey, endKey, hset) {
  // conta dias válidos de start a end, inclusive (equivalente a NETWORKDAYS.INTL "0000001")
  let d = parseDate(startKey);
  const end = parseDate(endKey);
  let count = 0;
  while (d <= end) {
    if (isValidDay(d, hset)) count++;
    d.setDate(d.getDate() + 1);
  }
  return count;
}
function addWorkdays(startKey, n, hset) {
  let d = parseDate(startKey);
  let remaining = Math.round(n);
  const forward = remaining >= 0;
  while (remaining !== 0) {
    d.setDate(d.getDate() + (forward ? 1 : -1));
    if (isValidDay(d, hset)) remaining += forward ? -1 : 1;
  }
  return dateToKey(d);
}

/* =========================================================
   Cálculo do Painel (mesmo método da planilha):
   taxa de consumo (qtd/dias úteis) média dos últimos 4
   intervalos entre as últimas 5 entregas; fallback para
   média de tempo simples; fallback final: intervalo padrão.
   ========================================================= */
function entregasDoCliente(nome) {
  return DB.entregas
    .filter(e => e.cliente === nome)
    .slice()
    .sort((a, b) => (a.data < b.data ? 1 : a.data > b.data ? -1 : 0)); // desc
}

function calcPainelCliente(cliente) {
  const hset = holidaySet();
  const lista = entregasDoCliente(cliente.nome).slice(0, 5);

  if (lista.length === 0) {
    return {
      ultimaEntrega: null, ultimaQtd: null, intervaloEstimado: cliente.intervaloPadrao || 7,
      proximaEntrega: null, diasAte: null, status: 'Sem histórico',
    };
  }

  const datas = lista.map(e => e.data);
  const qtds = lista.map(e => Number(e.quantidade) || 0);

  const intervalos = [];
  const taxas = [];
  for (let i = 0; i < datas.length - 1; i++) {
    const maisRecente = datas[i];
    const maisAntiga = datas[i + 1];
    const qtdAntiga = qtds[i + 1];
    const inicio = addDaysStr(maisAntiga, 1);
    const intervalo = networkDays(inicio, maisRecente, hset);
    intervalos.push(intervalo);
    if (qtdAntiga > 0 && intervalo > 0) taxas.push(qtdAntiga / intervalo);
  }

  let intervaloEstimado = null;
  if (taxas.length > 0) {
    const taxaMedia = taxas.reduce((a, b) => a + b, 0) / taxas.length;
    if (qtds[0] > 0 && taxaMedia > 0) intervaloEstimado = qtds[0] / taxaMedia;
  }
  if (intervaloEstimado === null) {
    if (intervalos.length > 0) {
      intervaloEstimado = intervalos.reduce((a, b) => a + b, 0) / intervalos.length;
    } else {
      intervaloEstimado = cliente.intervaloPadrao || 7;
    }
  }

  const ultimaEntrega = datas[0];
  const proximaEntrega = addWorkdays(ultimaEntrega, intervaloEstimado, hset);
  const diasAte = Math.round((parseDate(proximaEntrega) - parseDate(todayStr())) / 86400000);

  let status;
  if (diasAte < 0) status = 'Atrasado';
  else if (diasAte <= 2) status = 'Em breve';
  else status = 'No prazo';

  return { ultimaEntrega, ultimaQtd: qtds[0], intervaloEstimado, proximaEntrega, diasAte, status };
}

/* ---------------- Pagamentos pendentes ---------------- */
function valorTotalEntrega(e) {
  const q = Number(e.quantidade) || 0;
  const vu = Number(e.valorUnitario) || 0;
  return q * vu;
}
function getPendencias() {
  return DB.entregas
    .filter(e => !e.dataPagamento && valorTotalEntrega(e) > 0)
    .map(e => ({ id: e.id, cliente: e.cliente, data: e.data, valor: valorTotalEntrega(e) }))
    .sort((a, b) => (a.data < b.data ? -1 : a.data > b.data ? 1 : 0));
}
function getTotalPendente() {
  return getPendencias().reduce((s, p) => s + p.valor, 0);
}

/* ---------------- Resumo mensal ---------------- */
function getResumoMensal() {
  const clientesAtivos = DB.clientes.slice().sort((a, b) => a.nome.localeCompare(b.nome));
  if (DB.entregas.length === 0 || clientesAtivos.length === 0) {
    return { meses: [], linhas: [], totais: [] };
  }
  const datasOrdenadas = DB.entregas.map(e => e.data).sort();
  let inicio = parseDate(datasOrdenadas[0]);
  inicio = new Date(inicio.getFullYear(), inicio.getMonth(), 1, 12);
  const hoje = new Date();
  let fimRef = new Date(Math.max(
    parseDate(datasOrdenadas[datasOrdenadas.length - 1]).getTime(),
    hoje.getTime()
  ));
  fimRef = new Date(fimRef.getFullYear(), fimRef.getMonth(), 1, 12);

  const meses = [];
  let cursor = new Date(inicio);
  while (cursor <= fimRef) {
    meses.push(dateToKey(cursor));
    cursor.setMonth(cursor.getMonth() + 1);
  }

  const linhas = clientesAtivos.map(c => {
    const porMes = meses.map(mesKey => {
      const inicioMes = mesKey;
      const fimMesDate = parseDate(mesKey);
      fimMesDate.setMonth(fimMesDate.getMonth() + 1);
      const fimMes = dateToKey(fimMesDate);
      return DB.entregas
        .filter(e => e.cliente === c.nome && e.data >= inicioMes && e.data < fimMes)
        .reduce((s, e) => s + (Number(e.quantidade) || 0), 0);
    });
    const total = porMes.reduce((a, b) => a + b, 0);
    return { nome: c.nome, porMes, total };
  });

  const totais = meses.map((_, i) => linhas.reduce((s, l) => s + l.porMes[i], 0));
  const totalGeral = totais.reduce((a, b) => a + b, 0);

  return { meses, linhas, totais, totalGeral };
}

/* =========================================================
   Navegação
   ========================================================= */
const TITULOS = {
  painel: ['Painel', 'Próximas entregas previstas'],
  entregas: ['Entregas', 'Histórico de entregas'],
  pagamentos: ['Pagamentos', 'Pendências a receber'],
  mais: ['Mais', 'Cadastros e dados'],
  clientes: ['Lojas', 'Cadastro de clientes'],
  resumo: ['Resumo mensal', 'Picolés entregues por mês'],
  feriados: ['Feriados', 'Dias ignorados no cálculo'],
};

function showScreen(name) {
  document.querySelectorAll('.screen').forEach(s => s.classList.remove('active'));
  document.getElementById('screen-' + name).classList.add('active');

  document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  const tabMap = { painel: 'painel', entregas: 'entregas', pagamentos: 'pagamentos' };
  const activeTab = tabMap[name] || 'mais';
  const tabEl = document.querySelector(`.tab[data-tab="${activeTab}"]`);
  if (tabEl) tabEl.classList.add('active');

  const [title, sub] = TITULOS[name] || [name, ''];
  document.getElementById('topbar-title').textContent = title;
  document.getElementById('topbar-sub').textContent = sub;

  document.getElementById('fab-entrega').style.display = name === 'entregas' ? 'flex' : 'none';
  document.getElementById('fab-cliente').style.display = name === 'clientes' ? 'flex' : 'none';
  document.getElementById('fab-feriado').style.display = name === 'feriados' ? 'flex' : 'none';

  renderScreen(name);
}

function renderScreen(name) {
  if (name === 'painel') renderPainel();
  else if (name === 'entregas') renderEntregas();
  else if (name === 'pagamentos') renderPagamentos();
  else if (name === 'clientes') renderClientes();
  else if (name === 'resumo') renderResumo();
  else if (name === 'feriados') renderFeriados();
}

function refreshAll() {
  renderPainel(); renderEntregas(); renderPagamentos();
  renderClientes(); renderResumo(); renderFeriados(); renderSyncStatus();
}

/* ---------------- Toast ---------------- */
let toastTimer = null;
function toast(msg) {
  const el = document.getElementById('toast');
  el.textContent = msg;
  el.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => el.classList.remove('show'), 2200);
}

/* =========================================================
   PAINEL
   ========================================================= */
function badgeClass(status) {
  if (status === 'Atrasado') return 'red';
  if (status === 'Em breve') return 'amber';
  if (status === 'No prazo') return 'green';
  return 'grey';
}

function renderPainel() {
  const el = document.getElementById('painel-list');
  const ativos = DB.clientes.filter(c => c.ativo !== false);
  if (ativos.length === 0) {
    el.innerHTML = emptyState('🍭', 'Nenhuma loja cadastrada ainda', 'Cadastre suas lojas em Mais → Lojas para começar a ver as previsões aqui.');
    return;
  }
  const linhas = ativos.map(c => ({ c, r: calcPainelCliente(c) }));
  linhas.sort((a, b) => {
    if (a.r.diasAte === null && b.r.diasAte === null) return a.c.nome.localeCompare(b.c.nome);
    if (a.r.diasAte === null) return 1;
    if (b.r.diasAte === null) return -1;
    return a.r.diasAte - b.r.diasAte;
  });

  el.innerHTML = linhas.map(({ c, r }) => {
    const proxima = r.proximaEntrega ? fmtDateBR(r.proximaEntrega) : '—';
    const ultima = r.ultimaEntrega ? fmtDateBR(r.ultimaEntrega) : 'sem entregas';
    const diasTxt = r.diasAte === null ? '' :
      r.diasAte === 0 ? 'hoje' :
      r.diasAte > 0 ? `em ${r.diasAte} dia${r.diasAte > 1 ? 's' : ''}` :
      `${Math.abs(r.diasAte)} dia${Math.abs(r.diasAte) > 1 ? 's' : ''} atrás`;
    return `
      <div class="card" onclick="openClienteDetalhe('${c.id}')" style="cursor:pointer;">
        <div class="card-row">
          <div>
            <div class="card-title">${escapeHtml(c.nome)}</div>
            <div class="card-sub">Última entrega: ${ultima}${r.ultimaQtd ? ' · ' + r.ultimaQtd + ' un.' : ''}</div>
          </div>
          <span class="badge ${badgeClass(r.status)}">${r.status}</span>
        </div>
        <div class="card-sub" style="margin-top:8px;font-size:13px;color:var(--ink);">
          Próxima entrega prevista: <strong>${proxima}</strong> ${diasTxt ? '(' + diasTxt + ')' : ''}
        </div>
      </div>`;
  }).join('');
}

function openClienteDetalhe(clienteId) {
  const cliente = DB.clientes.find(c => c.id === clienteId);
  if (!cliente) return;
  const lista = entregasDoCliente(cliente.nome); // ja vem ordenada da mais recente pra mais antiga

  const totalQtd = lista.reduce((s, e) => s + (Number(e.quantidade) || 0), 0);

  const linhasHtml = lista.length === 0
    ? emptyState('🚚', 'Nenhuma entrega registrada', 'Ainda não há entregas lançadas para essa loja.')
    : lista.map(e => {
        const valor = valorTotalEntrega(e);
        const pago = !!e.dataPagamento;
        return `
          <div class="card" onclick="closeModal(); openEntregaForm('${e.id}')" style="cursor:pointer;">
            <div class="card-row">
              <div>
                <div class="card-title">${fmtDateBR(e.data)}</div>
                <div class="card-sub">${e.quantidade} un.${valor > 0 ? ' · ' + fmtMoney(valor) : ''}</div>
              </div>
              <span class="badge ${pago ? 'green' : 'amber'}">${pago ? 'Pago' : 'Pendente'}</span>
            </div>
          </div>`;
      }).join('');

  const html = `
    <div class="modal-header">
      <h2>${escapeHtml(cliente.nome)}</h2>
      <button class="close-x" onclick="closeModal()">✕</button>
    </div>
    ${lista.length > 0 ? `<div class="hint" style="margin-bottom:12px;">${lista.length} entrega${lista.length > 1 ? 's' : ''} registrada${lista.length > 1 ? 's' : ''} · ${totalQtd} un. no total. Toque numa entrega para editar.</div>` : ''}
    ${linhasHtml}
  `;
  openModal(html);
}

/* =========================================================
   ENTREGAS
   ========================================================= */
function renderEntregas() {
  const el = document.getElementById('entregas-list');
  const lista = DB.entregas.slice().sort((a, b) => (a.data < b.data ? 1 : a.data > b.data ? -1 : 0));
  if (lista.length === 0) {
    el.innerHTML = emptyState('🚚', 'Nenhuma entrega registrada', 'Toque no botão + para lançar a primeira entrega.');
    return;
  }
  el.innerHTML = lista.map(e => {
    const valorTotal = valorTotalEntrega(e);
    const pago = !!e.dataPagamento;
    return `
      <div class="card" onclick="openEntregaForm('${e.id}')" style="cursor:pointer;">
        <div class="card-row">
          <div>
            <div class="card-title">${escapeHtml(e.cliente)}</div>
            <div class="card-sub">${fmtDateBR(e.data)} · ${e.quantidade} un. ${valorTotal > 0 ? '· ' + fmtMoney(valorTotal) : ''}</div>
          </div>
          <span class="badge ${pago ? 'green' : 'amber'}">${pago ? 'Pago' : 'Pendente'}</span>
        </div>
      </div>`;
  }).join('');
}

function openEntregaForm(id) {
  const editando = !!id;
  const e = editando ? DB.entregas.find(x => x.id === id) : null;
  const ativos = DB.clientes.filter(c => c.ativo !== false);
  const options = ativos.map(c =>
    `<option value="${escapeHtml(c.nome)}" ${e && e.cliente === c.nome ? 'selected' : ''}>${escapeHtml(c.nome)}</option>`
  ).join('');

  if (ativos.length === 0) {
    toast('Cadastre uma loja antes de lançar uma entrega.');
    return;
  }

  const html = `
    <div class="modal-header">
      <h2>${editando ? 'Editar entrega' : 'Nova entrega'}</h2>
      <button class="close-x" onclick="closeModal()">✕</button>
    </div>
    <div class="field">
      <label>Loja</label>
      <select id="f-cliente">${options}</select>
    </div>
    <div class="row2">
      <div class="field">
        <label>Data da entrega</label>
        <input type="date" id="f-data" value="${e ? e.data : todayStr()}">
      </div>
      <div class="field">
        <label>Quantidade</label>
        <input type="number" id="f-qtd" inputmode="numeric" value="${e ? e.quantidade : ''}" placeholder="0">
      </div>
    </div>
    <div class="field">
      <label>Observações da entrega</label>
      <textarea id="f-obs" placeholder="Opcional">${e ? escapeHtml(e.observacoes || '') : ''}</textarea>
    </div>
    <div class="field">
      <label>Valor unitário (por picolé)</label>
      <input type="number" id="f-valor-unit" inputmode="decimal" step="0.01" value="${e ? e.valorUnitario : ''}" placeholder="0,00">
    </div>
    <div class="field">
      <label>Data do pagamento</label>
      <input type="date" id="f-data-pagto" value="${e && e.dataPagamento ? e.dataPagamento : ''}">
      <div class="hint">Em branco = pendente. Preenchida = pago.</div>
    </div>
    <div class="field">
      <label>Observações do pagamento</label>
      <textarea id="f-obs-pagto" placeholder="Opcional">${e ? escapeHtml(e.obsPagamento || '') : ''}</textarea>
    </div>
    <div class="formbtns">
      <button class="btn ghost block" onclick="closeModal()">Cancelar</button>
      <button class="btn primary block" onclick="saveEntrega(${editando ? `'${id}'` : 'null'})">Salvar</button>
    </div>
    ${editando ? `<button class="danger-link" onclick="deleteEntrega('${id}')">Excluir esta entrega</button>` : ''}
  `;
  openModal(html);
}

function saveEntrega(id) {
  const cliente = document.getElementById('f-cliente').value;
  const data = document.getElementById('f-data').value;
  const quantidade = Number(document.getElementById('f-qtd').value) || 0;
  const observacoes = document.getElementById('f-obs').value.trim();
  const valorUnitario = Number(document.getElementById('f-valor-unit').value) || 0;
  const dataPagamento = document.getElementById('f-data-pagto').value || null;
  const obsPagamento = document.getElementById('f-obs-pagto').value.trim();

  if (!data || !cliente || quantidade <= 0) {
    toast('Preencha loja, data e uma quantidade maior que zero.');
    return;
  }

  if (id) {
    if (!confirm('Salvar as alterações feitas nesta entrega?')) {
      return;
    }
    const e = DB.entregas.find(x => x.id === id);
    Object.assign(e, { cliente, data, quantidade, observacoes, valorUnitario, dataPagamento, obsPagamento });
  } else {
    DB.entregas.push({ id: uid('e'), cliente, data, quantidade, observacoes, valorUnitario, dataPagamento, obsPagamento });
  }
  saveDB();
  closeModal();
  refreshAll();
  toast('Entrega salva.');
}

function deleteEntrega(id) {
  if (!confirm('Excluir esta entrega? Essa ação não pode ser desfeita.')) return;
  DB.entregas = DB.entregas.filter(x => x.id !== id);
  saveDB();
  closeModal();
  refreshAll();
  toast('Entrega excluída.');
}

/* =========================================================
   PAGAMENTOS
   ========================================================= */
function renderPagamentos() {
  document.getElementById('total-pendente').textContent = fmtMoney(getTotalPendente());
  const el = document.getElementById('pagamentos-list');
  const pend = getPendencias();
  if (pend.length === 0) {
    el.innerHTML = emptyState('✅', 'Nada pendente', 'Todas as entregas com valor lançado já foram pagas.');
    return;
  }
  el.innerHTML = pend.map(p => `
    <div class="card">
      <div class="card-row">
        <div>
          <div class="card-title">${escapeHtml(p.cliente)}</div>
          <div class="card-sub">Entrega de ${fmtDateBR(p.data)}</div>
        </div>
        <div style="text-align:right;">
          <div class="card-title" style="color:var(--coral-dark);">${fmtMoney(p.valor)}</div>
        </div>
      </div>
      <button class="btn ghost small" style="margin-top:10px;" onclick="abrirConfirmacaoPagamento('${p.id}')">Marcar como pago</button>
    </div>
  `).join('');
}

function abrirConfirmacaoPagamento(entregaId) {
  const e = DB.entregas.find(x => x.id === entregaId);
  if (!e) return;
  const html = `
    <div class="modal-header">
      <h2>Confirmar pagamento</h2>
      <button class="close-x" onclick="closeModal()">✕</button>
    </div>
    <div class="hint" style="margin-bottom:14px;">${escapeHtml(e.cliente)} · Entrega de ${fmtDateBR(e.data)} · ${fmtMoney(valorTotalEntrega(e))}</div>
    <div class="field">
      <label>Data do pagamento</label>
      <input type="date" id="f-confirma-data-pagto" value="${todayStr()}">
    </div>
    <div class="field">
      <label>Observações do pagamento</label>
      <textarea id="f-confirma-obs-pagto" placeholder="Opcional">${escapeHtml(e.obsPagamento || '')}</textarea>
    </div>
    <div class="formbtns">
      <button class="btn ghost block" onclick="closeModal()">Cancelar</button>
      <button class="btn primary block" onclick="confirmarPagamento('${entregaId}')">Confirmar pagamento</button>
    </div>
  `;
  openModal(html);
}

function confirmarPagamento(id) {
  const e = DB.entregas.find(x => x.id === id);
  if (!e) return;
  const data = document.getElementById('f-confirma-data-pagto').value;
  if (!data) { toast('Escolha a data do pagamento.'); return; }
  e.dataPagamento = data;
  e.obsPagamento = document.getElementById('f-confirma-obs-pagto').value.trim();
  saveDB();
  closeModal();
  refreshAll();
  toast('Pagamento confirmado.');
}

function abrirHistoricoPagos() {
  const pagos = DB.entregas
    .filter(e => e.dataPagamento)
    .slice()
    .sort((a, b) => (a.dataPagamento < b.dataPagamento ? 1 : a.dataPagamento > b.dataPagamento ? -1 : 0));

  const linhasHtml = pagos.length === 0
    ? emptyState('🧾', 'Nenhum pagamento registrado', 'Assim que marcar uma entrega como paga, ela aparece aqui.')
    : pagos.map(e => `
        <div class="card" onclick="closeModal(); openEntregaForm('${e.id}')" style="cursor:pointer;">
          <div class="card-row">
            <div>
              <div class="card-title">${escapeHtml(e.cliente)}</div>
              <div class="card-sub">Entrega de ${fmtDateBR(e.data)} · Pago em ${fmtDateBR(e.dataPagamento)}</div>
            </div>
            <div class="card-title" style="color:var(--teal-700);">${fmtMoney(valorTotalEntrega(e))}</div>
          </div>
        </div>
      `).join('');

  const html = `
    <div class="modal-header">
      <h2>Pagamentos realizados</h2>
      <button class="close-x" onclick="closeModal()">✕</button>
    </div>
    ${pagos.length > 0 ? `<div class="hint" style="margin-bottom:12px;">Do mais recente para o mais antigo. Toque para editar.</div>` : ''}
    ${linhasHtml}
  `;
  openModal(html);
}

/* ---------------- Exportar pendências como imagem ---------------- */
async function exportarPagamentosImagem() {
  const pend = getPendencias();
  if (pend.length === 0) {
    toast('Nada pendente para exportar.');
    return;
  }

  const WIDTH = 800;
  const PAD = 32;
  const ROW_H = 66;
  const HEADER_H = 158;
  const FOOTER_H = 46;
  const height = HEADER_H + pend.length * ROW_H + FOOTER_H;

  const scale = 2;
  const canvas = document.createElement('canvas');
  canvas.width = WIDTH * scale;
  canvas.height = height * scale;
  const ctx = canvas.getContext('2d');
  ctx.scale(scale, scale);

  // fundo
  ctx.fillStyle = '#FAFAF8';
  ctx.fillRect(0, 0, WIDTH, height);

  // faixa do cabecalho
  const grad = ctx.createLinearGradient(0, 0, WIDTH, 0);
  grad.addColorStop(0, '#0E7C7B');
  grad.addColorStop(1, '#0A4F4E');
  ctx.fillStyle = grad;
  ctx.fillRect(0, 0, WIDTH, HEADER_H);

  ctx.fillStyle = '#FFFFFF';
  ctx.font = '700 25px Roboto, Arial, sans-serif';
  ctx.fillText('Pagamentos pendentes a receber', PAD, 46);

  ctx.font = '400 14px Roboto, Arial, sans-serif';
  ctx.globalAlpha = 0.85;
  ctx.fillText(`Gerado em ${new Date().toLocaleDateString('pt-BR')} · Entregas Picolés`, PAD, 70);
  ctx.globalAlpha = 1;

  ctx.font = '400 13px Roboto, Arial, sans-serif';
  ctx.globalAlpha = 0.8;
  ctx.fillText('TOTAL PENDENTE', PAD, 106);
  ctx.globalAlpha = 1;
  ctx.font = '800 32px Roboto, Arial, sans-serif';
  ctx.fillText(fmtMoney(getTotalPendente()), PAD, 138);

  // linhas
  let y = HEADER_H;
  pend.forEach((p, i) => {
    if (i % 2 === 1) {
      ctx.fillStyle = '#F0F4F4';
      ctx.fillRect(0, y, WIDTH, ROW_H);
    }
    ctx.fillStyle = '#1F2933';
    ctx.font = '700 19px Roboto, Arial, sans-serif';
    ctx.fillText(p.cliente, PAD, y + 28);

    ctx.fillStyle = '#5C6B73';
    ctx.font = '400 14px Roboto, Arial, sans-serif';
    ctx.fillText(`Entrega de ${fmtDateBR(p.data)}`, PAD, y + 48);

    ctx.fillStyle = '#E5502F';
    ctx.font = '700 20px Roboto, Arial, sans-serif';
    const valorTxt = fmtMoney(p.valor);
    const w = ctx.measureText(valorTxt).width;
    ctx.fillText(valorTxt, WIDTH - PAD - w, y + 38);

    ctx.strokeStyle = '#E7ECEC';
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(PAD, y + ROW_H - 1);
    ctx.lineTo(WIDTH - PAD, y + ROW_H - 1);
    ctx.stroke();

    y += ROW_H;
  });

  // rodape
  ctx.fillStyle = '#5C6B73';
  ctx.font = '400 12.5px Roboto, Arial, sans-serif';
  ctx.fillText(`${pend.length} entrega${pend.length > 1 ? 's' : ''} pendente${pend.length > 1 ? 's' : ''}`, PAD, y + 28);

  canvas.toBlob(async (blob) => {
    if (!blob) { toast('Não foi possível gerar a imagem.'); return; }
    const nomeArquivo = `Pagamentos_Pendentes_${todayStr()}.png`;
    const file = new File([blob], nomeArquivo, { type: 'image/png' });

    if (navigator.canShare && navigator.canShare({ files: [file] })) {
      try {
        await navigator.share({ files: [file], title: 'Pagamentos pendentes' });
        return;
      } catch (err) {
        if (err && err.name === 'AbortError') return; // usuario cancelou o compartilhamento
        // se falhar por outro motivo, cai no download abaixo
      }
    }
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = nomeArquivo;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1500);
    toast('Imagem salva.');
  }, 'image/png');
}

/* =========================================================
   CLIENTES (Lojas)
   ========================================================= */
function renderClientes() {
  const el = document.getElementById('clientes-list');
  const lista = DB.clientes.slice().sort((a, b) => a.nome.localeCompare(b.nome));
  if (lista.length === 0) {
    el.innerHTML = emptyState('🏬', 'Nenhuma loja cadastrada', 'Toque no botão + para cadastrar a primeira loja.');
    return;
  }
  el.innerHTML = lista.map(c => `
    <div class="card" onclick="openClienteForm('${c.id}')" style="cursor:pointer;">
      <div class="card-row">
        <div>
          <div class="card-title">${escapeHtml(c.nome)}</div>
          <div class="card-sub">${escapeHtml(c.endereco || 'Sem endereço cadastrado')}</div>
        </div>
        <span class="badge ${c.ativo === false ? 'grey' : 'green'}">${c.ativo === false ? 'Inativa' : 'Ativa'}</span>
      </div>
    </div>
  `).join('');
}

function openClienteForm(id) {
  const editando = !!id;
  const c = editando ? DB.clientes.find(x => x.id === id) : null;
  const html = `
    <div class="modal-header">
      <h2>${editando ? 'Editar loja' : 'Nova loja'}</h2>
      <button class="close-x" onclick="closeModal()">✕</button>
    </div>
    <div class="field">
      <label>Nome da loja</label>
      <input type="text" id="f-nome" value="${c ? escapeHtml(c.nome) : ''}" placeholder="Ex: Padaria do Zé">
    </div>
    <div class="field">
      <label>Endereço</label>
      <input type="text" id="f-endereco" value="${c ? escapeHtml(c.endereco || '') : ''}" placeholder="Opcional">
    </div>
    <div class="field">
      <label>Contato</label>
      <input type="text" id="f-contato" value="${c ? escapeHtml(c.contato || '') : ''}" placeholder="Opcional">
    </div>
    <div class="field">
      <label>Intervalo padrão (dias)</label>
      <input type="number" id="f-intervalo" value="${c ? c.intervaloPadrao : 7}">
      <div class="hint">Usado só até a loja acumular pelo menos 2 entregas registradas.</div>
    </div>
    <div class="field">
      <label>Status</label>
      <select id="f-ativo">
        <option value="true" ${!c || c.ativo !== false ? 'selected' : ''}>Ativa</option>
        <option value="false" ${c && c.ativo === false ? 'selected' : ''}>Inativa</option>
      </select>
    </div>
    <div class="formbtns">
      <button class="btn ghost block" onclick="closeModal()">Cancelar</button>
      <button class="btn primary block" onclick="saveCliente(${editando ? `'${id}'` : 'null'})">Salvar</button>
    </div>
    ${editando ? `<button class="danger-link" onclick="deleteCliente('${id}')">Excluir esta loja</button>` : ''}
  `;
  openModal(html);
}

function saveCliente(id) {
  const nome = document.getElementById('f-nome').value.trim();
  const endereco = document.getElementById('f-endereco').value.trim();
  const contato = document.getElementById('f-contato').value.trim();
  const intervaloPadrao = Number(document.getElementById('f-intervalo').value) || 7;
  const ativo = document.getElementById('f-ativo').value === 'true';

  if (!nome) { toast('Digite o nome da loja.'); return; }

  const nomeDuplicado = DB.clientes.some(c => c.nome === nome && c.id !== id);
  if (nomeDuplicado) { toast('Já existe uma loja com esse nome.'); return; }

  if (id) {
    const antigo = DB.clientes.find(x => x.id === id);
    const nomeAntigo = antigo.nome;
    Object.assign(antigo, { nome, endereco, contato, intervaloPadrao, ativo });
    if (nomeAntigo !== nome) {
      DB.entregas.forEach(e => { if (e.cliente === nomeAntigo) e.cliente = nome; });
    }
  } else {
    DB.clientes.push({ id: uid('c'), nome, endereco, contato, intervaloPadrao, ativo });
  }
  saveDB();
  closeModal();
  refreshAll();
  toast('Loja salva.');
}

function deleteCliente(id) {
  const c = DB.clientes.find(x => x.id === id);
  const temEntregas = DB.entregas.some(e => e.cliente === c.nome);
  const msg = temEntregas
    ? 'Essa loja tem entregas registradas. Elas serão mantidas no histórico, mas a loja some dos cadastros. Continuar?'
    : 'Excluir esta loja?';
  if (!confirm(msg)) return;
  DB.clientes = DB.clientes.filter(x => x.id !== id);
  saveDB();
  closeModal();
  refreshAll();
  toast('Loja excluída.');
}

/* =========================================================
   RESUMO MENSAL
   ========================================================= */
function renderResumo() {
  const el = document.getElementById('resumo-container');
  const { meses, linhas, totais, totalGeral } = getResumoMensal();
  if (meses.length === 0) {
    el.innerHTML = emptyState('📊', 'Sem dados ainda', 'Assim que houver entregas registradas, o resumo mensal aparece aqui.');
    return;
  }
  const headCols = meses.map(m => `<th>${fmtMesAno(m)}</th>`).join('');
  const bodyRows = linhas.map(l => `
    <tr>
      <td>${escapeHtml(l.nome)}</td>
      ${l.porMes.map(v => `<td>${v || 0}</td>`).join('')}
      <td><strong>${l.total}</strong></td>
    </tr>
  `).join('');
  const totalRow = `
    <tr class="total-row">
      <td>Total geral</td>
      ${totais.map(v => `<td>${v}</td>`).join('')}
      <td>${totalGeral}</td>
    </tr>
  `;
  el.innerHTML = `
    <div class="table-wrap">
      <table class="resumo">
        <thead><tr><th>Loja</th>${headCols}<th>Total</th></tr></thead>
        <tbody>${totalRow}${bodyRows}</tbody>
      </table>
    </div>
  `;
}

/* =========================================================
   FERIADOS
   ========================================================= */
function renderFeriados() {
  const el = document.getElementById('feriados-list');
  const lista = DB.feriados.slice().sort((a, b) => a.data.localeCompare(b.data));
  if (lista.length === 0) {
    el.innerHTML = emptyState('📅', 'Nenhum feriado cadastrado', 'Toque no botão + para adicionar.');
    return;
  }
  el.innerHTML = lista.map(f => `
    <div class="card" onclick="openFeriadoForm('${f.id}')" style="cursor:pointer;">
      <div class="card-row">
        <div>
          <div class="card-title">${fmtDateBR(f.data)}</div>
          <div class="card-sub">${escapeHtml(f.descricao)}</div>
        </div>
      </div>
    </div>
  `).join('');
}

function openFeriadoForm(id) {
  const editando = !!id;
  const f = editando ? DB.feriados.find(x => x.id === id) : null;
  const html = `
    <div class="modal-header">
      <h2>${editando ? 'Editar feriado' : 'Novo feriado'}</h2>
      <button class="close-x" onclick="closeModal()">✕</button>
    </div>
    <div class="field">
      <label>Data</label>
      <input type="date" id="f-data-feriado" value="${f ? f.data : ''}">
    </div>
    <div class="field">
      <label>Descrição</label>
      <input type="text" id="f-desc-feriado" value="${f ? escapeHtml(f.descricao) : ''}" placeholder="Ex: Aniversário da cidade">
    </div>
    <div class="formbtns">
      <button class="btn ghost block" onclick="closeModal()">Cancelar</button>
      <button class="btn primary block" onclick="saveFeriado(${editando ? `'${id}'` : 'null'})">Salvar</button>
    </div>
    ${editando ? `<button class="danger-link" onclick="deleteFeriado('${id}')">Excluir este feriado</button>` : ''}
  `;
  openModal(html);
}

function saveFeriado(id) {
  const data = document.getElementById('f-data-feriado').value;
  const descricao = document.getElementById('f-desc-feriado').value.trim();
  if (!data || !descricao) { toast('Preencha data e descrição.'); return; }
  if (id) {
    Object.assign(DB.feriados.find(x => x.id === id), { data, descricao });
  } else {
    DB.feriados.push({ id: uid('f'), data, descricao });
  }
  saveDB();
  closeModal();
  refreshAll();
  toast('Feriado salvo.');
}

function deleteFeriado(id) {
  if (!confirm('Excluir este feriado?')) return;
  DB.feriados = DB.feriados.filter(x => x.id !== id);
  saveDB();
  closeModal();
  refreshAll();
  toast('Feriado excluído.');
}

/* =========================================================
   Helpers de UI
   ========================================================= */
function emptyState(icone, titulo, texto) {
  return `<div class="empty"><div class="stick">${icone}</div><strong>${titulo}</strong><div style="margin-top:6px;">${texto}</div></div>`;
}
function escapeHtml(s) {
  if (s === null || s === undefined) return '';
  return String(s).replace(/[&<>"']/g, m => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[m]));
}
function openModal(html) {
  document.getElementById('modal-content').innerHTML = html;
  document.getElementById('modal-backdrop').classList.add('active');
}
function closeModal() {
  document.getElementById('modal-backdrop').classList.remove('active');
}
document.getElementById('modal-backdrop').addEventListener('click', (e) => {
  if (e.target.id === 'modal-backdrop') closeModal();
});

function confirmarLimpeza() {
  if (!confirm('Isso apaga TODOS os dados salvos neste app (lojas, entregas, feriados). Não pode ser desfeito. Continuar?')) return;
  localStorage.removeItem(STORAGE_KEY);
  loadDB();
  refreshAll();
  toast('Dados apagados.');
}

/* ---------------- UI de sincronização ---------------- */
function renderSyncStatus() {
  const el = document.getElementById('sync-status-card');
  if (!el) return;
  if (!SYNC_CONFIG) {
    el.innerHTML = `
      <div class="settings-item" onclick="abrirConfigSync()" style="cursor:pointer;">
        <div><div class="t">Sincronização e backup na nuvem</div><div class="d">Desativada — toque para configurar</div></div>
        <div>›</div>
      </div>`;
    return;
  }
  const lastSync = localStorage.getItem(LAST_SYNC_KEY);
  const lastTxt = lastSync ? new Date(Number(lastSync)).toLocaleString('pt-BR') : 'ainda não';
  el.innerHTML = `
    <div class="settings-item">
      <div><div class="t">Sincronização ativa</div><div class="d">Última sincronização: ${lastTxt}</div></div>
    </div>
    <button class="btn ghost block" style="margin:4px 0 8px;" onclick="sincronizarAgora()">🔄 Sincronizar agora</button>
    <button class="btn ghost block" style="margin-bottom:8px;" onclick="mostrarCodigoGerado('${SYNC_CONFIG.syncCode}')">Ver código de sincronização</button>
    <button class="danger-link" onclick="desconectarSync()">Desconectar sincronização deste aparelho</button>
  `;
}

function abrirConfigSync() {
  const atual = SYNC_CONFIG;
  const html = `
    <div class="modal-header">
      <h2>Sincronização na nuvem</h2>
      <button class="close-x" onclick="closeModal()">✕</button>
    </div>
    <div class="hint" style="margin-bottom:14px;">
      Cole abaixo a configuração do seu projeto Firebase (o bloco com apiKey, projectId etc. que aparece no console do Firebase, em Configurações do projeto → Seus apps).
    </div>
    <div class="field">
      <label>Configuração do Firebase</label>
      <textarea id="f-firebase-config" placeholder="Cole aqui o firebaseConfig" style="min-height:110px;">${atual ? escapeHtml(JSON.stringify(atual.firebaseConfig, null, 2)) : ''}</textarea>
    </div>
    <div class="field">
      <label>Código de sincronização</label>
      <input type="text" id="f-sync-code" value="${atual ? atual.syncCode : ''}" placeholder="Deixe em branco para criar um novo">
      <div class="hint">Deixe em branco se este for o primeiro aparelho (um código novo é criado). Se já tiver um código de outro aparelho, cole ele aqui.</div>
    </div>
    <div class="formbtns">
      <button class="btn ghost block" onclick="closeModal()">Cancelar</button>
      <button class="btn primary block" onclick="salvarConfigSyncUI()">Salvar</button>
    </div>
  `;
  openModal(html);
}

function salvarConfigSyncUI() {
  const text = document.getElementById('f-firebase-config').value;
  const parsed = parseFirebaseConfigText(text);
  if (!parsed.apiKey || !parsed.projectId) {
    toast('Não consegui entender essa configuração. Confira se colou o bloco certo.');
    return;
  }
  let code = document.getElementById('f-sync-code').value.trim();
  const novoCodigo = !code;
  if (!code) code = gerarCodigoSync();

  SYNC_CONFIG = { firebaseConfig: parsed, syncCode: code };
  localStorage.setItem(SYNC_CONFIG_KEY, JSON.stringify(SYNC_CONFIG));
  firebaseReady = null; // forca reinicializar com a config nova
  closeModal();
  renderSyncStatus();

  if (novoCodigo) {
    toast('Sincronização configurada!');
    pushParaNuvem();
    setTimeout(() => mostrarCodigoGerado(code), 300);
  } else {
    sincronizar(false);
  }
}

function mostrarCodigoGerado(code) {
  const html = `
    <div class="modal-header">
      <h2>Seu código de sincronização</h2>
      <button class="close-x" onclick="closeModal()">✕</button>
    </div>
    <div class="hint" style="margin-bottom:10px;">Guarde este código. Use-o para conectar outros aparelhos aos mesmos dados: em Mais → Sincronização, cole a mesma configuração do Firebase e este código.</div>
    <div class="card" style="text-align:center;">
      <div style="font-size:20px;font-weight:800;letter-spacing:.03em;font-family:monospace;color:var(--teal-900);word-break:break-all;">${code}</div>
    </div>
    <button class="btn primary block" style="margin-top:14px;" onclick="closeModal()">Entendi</button>
  `;
  openModal(html);
}

function desconectarSync() {
  if (!confirm('Isso para de sincronizar este aparelho (os dados salvos aqui continuam, só param de se atualizar com a nuvem). Continuar?')) return;
  SYNC_CONFIG = null;
  localStorage.removeItem(SYNC_CONFIG_KEY);
  closeModal();
  renderSyncStatus();
  toast('Sincronização desconectada.');
}

/* =========================================================
   Importar / Exportar planilha (.xlsx) — via SheetJS (CDN)
   ========================================================= */
function ensureSheetJS() {
  if (window.XLSX) return Promise.resolve();
  return new Promise((resolve, reject) => {
    const s = document.createElement('script');
    s.src = SHEETJS_URL;
    s.onload = () => resolve();
    s.onerror = () => reject(new Error('Falha ao carregar biblioteca (sem internet?)'));
    document.head.appendChild(s);
  });
}

function triggerImport() {
  if (!confirm('Importar substitui TODOS os dados atuais deste app pelos dados da planilha. Continuar?')) return;
  document.getElementById('file-import').click();
}

document.getElementById('file-import').addEventListener('change', async (ev) => {
  const file = ev.target.files[0];
  ev.target.value = '';
  if (!file) return;
  toast('Lendo planilha…');
  try {
    await ensureSheetJS();
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: 'array', cellDates: true });
    importarWorkbook(wb);
  } catch (err) {
    console.error(err);
    toast('Não foi possível importar. Verifique sua internet e o arquivo.');
  }
});

function excelDateToKey(v) {
  if (v instanceof Date && !isNaN(v)) {
    return `${v.getUTCFullYear()}-${String(v.getUTCMonth() + 1).padStart(2, '0')}-${String(v.getUTCDate()).padStart(2, '0')}`;
  }
  if (typeof v === 'number' && v > 0) {
    const ms = Math.round((v - 25569) * 86400 * 1000);
    const d = new Date(ms);
    return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, '0')}-${String(d.getUTCDate()).padStart(2, '0')}`;
  }
  if (typeof v === 'string' && v.trim()) {
    const m = v.trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (m) return `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`;
  }
  return null;
}

function importarWorkbook(wb) {
  const novo = { clientes: [], feriados: [], entregas: [] };

  const shClientes = wb.Sheets['Clientes'];
  if (shClientes) {
    XLSX.utils.sheet_to_json(shClientes, { defval: '' }).forEach(r => {
      const nome = String(r['Nome da Loja'] || '').trim();
      if (!nome) return;
      const ativoRaw = String(r['Ativo'] || '').trim().toLowerCase();
      const ativo = !['não', 'nao', 'false', 'inativo'].includes(ativoRaw);
      novo.clientes.push({
        id: uid('c'), nome,
        endereco: String(r['Endereço'] || r['Endereco'] || ''),
        contato: String(r['Contato'] || ''),
        intervaloPadrao: Number(r['Intervalo Padrao Dias']) || 7,
        ativo,
      });
    });
  }

  const shFeriados = wb.Sheets['Feriados'];
  if (shFeriados) {
    XLSX.utils.sheet_to_json(shFeriados, { defval: '' }).forEach(r => {
      const dataKey = excelDateToKey(r['Data']);
      if (!dataKey) return;
      novo.feriados.push({ id: uid('f'), data: dataKey, descricao: String(r['Descrição'] || r['Descricao'] || '') });
    });
  }
  if (novo.feriados.length === 0) novo.feriados = FERIADOS_PADRAO.map(([data, descricao]) => ({ id: uid('f'), data, descricao }));

  let importadas = 0;
  const shEntregas = wb.Sheets['Entregas'];
  if (shEntregas) {
    XLSX.utils.sheet_to_json(shEntregas, { defval: '' }).forEach(r => {
      const dataKey = excelDateToKey(r['Data']);
      const cliente = String(r['Cliente'] || '').trim();
      const quantidade = Number(r['Quantidade']) || 0;
      if (!dataKey || !cliente || quantidade <= 0) return;
      novo.entregas.push({
        id: uid('e'), data: dataKey, cliente, quantidade,
        observacoes: String(r['Observações'] || r['Observacoes'] || ''),
        valorUnitario: Number(r['Valor Unitario'] || r['Valor Unitário']) || 0,
        dataPagamento: excelDateToKey(r['Data Pagamento']),
        obsPagamento: String(r['Observações de Pagamento'] || r['Observacoes de Pagamento'] || ''),
      });
      importadas++;
    });
  }

  DB = novo;
  saveDB();
  refreshAll();
  showScreen('painel');
  toast(`Importado: ${novo.clientes.length} loja(s), ${importadas} entrega(s).`);
}

async function exportarPlanilha() {
  toast('Preparando planilha…');
  try {
    await ensureSheetJS();
  } catch (err) {
    toast('Não foi possível exportar. Verifique sua internet.');
    return;
  }

  const wb = XLSX.utils.book_new();

  const clientesRows = DB.clientes.map((c, i) => ({
    'ID': i + 1,
    'Nome da Loja': c.nome,
    'Endereço': c.endereco || '',
    'Contato': c.contato || '',
    'Intervalo Padrao Dias': c.intervaloPadrao,
    'Ativo': c.ativo === false ? 'Não' : 'Sim',
  }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(clientesRows), 'Clientes');

  const feriadosRows = DB.feriados.slice().sort((a, b) => a.data.localeCompare(b.data)).map(f => ({
    'Data': fmtDateBR(f.data),
    'Descrição': f.descricao,
  }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(feriadosRows), 'Feriados');

  const entregasRows = DB.entregas.slice().sort((a, b) => a.data.localeCompare(b.data)).map(e => ({
    'Data': fmtDateBR(e.data),
    'Cliente': e.cliente,
    'Quantidade': e.quantidade,
    'Observações': e.observacoes || '',
    'Valor Unitario': e.valorUnitario || 0,
    'Valor Total': valorTotalEntrega(e),
    'Data Pagamento': e.dataPagamento ? fmtDateBR(e.dataPagamento) : '',
    'Observações de Pagamento': e.obsPagamento || '',
  }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(entregasRows), 'Entregas');

  const painelRows = DB.clientes.filter(c => c.ativo !== false).map(c => {
    const r = calcPainelCliente(c);
    return {
      'Cliente': c.nome,
      'Última Entrega': r.ultimaEntrega ? fmtDateBR(r.ultimaEntrega) : '',
      'Última Quantidade': r.ultimaQtd || '',
      'Intervalo Estimado (dias)': r.intervaloEstimado ? Number(r.intervaloEstimado.toFixed(1)) : '',
      'Próxima Entrega Estimada': r.proximaEntrega ? fmtDateBR(r.proximaEntrega) : 'Sem entregas registradas',
      'Dias até a Entrega': r.diasAte === null ? '' : r.diasAte,
      'Status': r.status,
    };
  });
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(painelRows), 'Painel');

  const { meses, linhas, totais, totalGeral } = getResumoMensal();
  if (meses.length > 0) {
    const header = { 'Loja': 'Total Geral' };
    meses.forEach((m, i) => { header[fmtMesAno(m)] = totais[i]; });
    header['Total'] = totalGeral;
    const resumoRows = [header];
    linhas.forEach(l => {
      const row = { 'Loja': l.nome };
      meses.forEach((m, i) => { row[fmtMesAno(m)] = l.porMes[i]; });
      row['Total'] = l.total;
      resumoRows.push(row);
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(resumoRows), 'Resumo Mensal');
  }

  const pend = getPendencias();
  const pagRows = [{ 'Loja': 'TOTAL PENDENTE', 'Data da Entrega': '', 'Valor a Receber': getTotalPendente() }];
  pend.forEach(p => pagRows.push({ 'Loja': p.cliente, 'Data da Entrega': fmtDateBR(p.data), 'Valor a Receber': p.valor }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(pagRows), 'Pagamentos');

  XLSX.writeFile(wb, `Entregas_Picoles_${todayStr()}.xlsx`);
  toast('Planilha exportada.');
}

/* =========================================================
   Inicialização
   ========================================================= */
carregarConfigSync();
loadDB();
refreshAll();
if (SYNC_CONFIG) {
  sincronizar(true);
}
