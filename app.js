/* =========================================================
   Entregas Picolé — app pessoal (dados salvos no navegador)
   ========================================================= */

const STORAGE_KEY = 'picole_data_v1';
const SHEETJS_URL = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';

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
  renderClientes(); renderResumo(); renderFeriados();
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
  const ordem = { 'Atrasado': 0, 'Em breve': 1, 'No prazo': 2, 'Sem histórico': 3 };
  linhas.sort((a, b) => (ordem[a.r.status] - ordem[b.r.status]) || a.c.nome.localeCompare(b.c.nome));

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
      <button class="btn ghost small" style="margin-top:10px;" onclick="marcarComoPago('${p.id}')">Marcar como pago hoje</button>
    </div>
  `).join('');
}

function marcarComoPago(id) {
  const e = DB.entregas.find(x => x.id === id);
  if (!e) return;
  e.dataPagamento = todayStr();
  saveDB();
  refreshAll();
  toast('Pagamento registrado.');
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
  ctx.fillText(`Gerado em ${new Date().toLocaleDateString('pt-BR')} · Entregas Picolé`, PAD, 70);
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
loadDB();
refreshAll();
