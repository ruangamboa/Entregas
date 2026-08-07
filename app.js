/* =========================================================
   Entregas Picolé — app pessoal (dados salvos no navegador)
   ========================================================= */

const STORAGE_KEY = 'picole_data_v1';
const LAST_SYNC_KEY = 'picole_last_sync';
const WORKSPACE_ID = 'principal'; // workspace fixo — sem mais codigo de sincronizacao
const SHEETJS_URL = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
const FIREBASE_APP_URL = 'https://www.gstatic.com/firebasejs/12.15.0/firebase-app-compat.js';
const FIREBASE_FIRESTORE_URL = 'https://www.gstatic.com/firebasejs/12.15.0/firebase-firestore-compat.js';
const FIREBASE_AUTH_URL = 'https://www.gstatic.com/firebasejs/12.15.0/firebase-auth-compat.js';

// Configuração do Firebase deste app — ja vem pronta, ninguem precisa colar
// nada em nenhum aparelho.
const FIREBASE_CONFIG = {
  apiKey: "AIzaSyDoQkfipxKJmpipiMUFS9BPKias-y24w7I",
  authDomain: "entregas-picoles.firebaseapp.com",
  projectId: "entregas-picoles",
  storageBucket: "entregas-picoles.firebasestorage.app",
  messagingSenderId: "1062049635898",
  appId: "1:1062049635898:web:dc3a93626d37631a6a1f03",
};

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

/* =========================================================
   Tipos de cliente / venda
   - personalizada: fluxo original (já existia antes)
   - comum: entregas comuns, com pagamento parcelado e vencimento
   - direta: vendas diretas, fluxo igual ao personalizada
   ========================================================= */
const TIPOS_CLIENTE = ['personalizada', 'comum', 'direta'];
const TIPO_LABEL = {
  personalizada: 'Entregas Personalizadas',
  comum: 'Entregas Comuns',
  direta: 'Vendas Diretas',
};
const TIPO_LABEL_CURTO = {
  personalizada: 'Personalizadas',
  comum: 'Comuns',
  direta: 'Diretas',
};
let TIPO_ATUAL = localStorage.getItem('picole_tipo_atual') || 'personalizada';
if (!TIPOS_CLIENTE.includes(TIPO_ATUAL)) TIPO_ATUAL = 'personalizada';

function normalizarTexto(s) {
  return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim().toLowerCase();
}
function normalizarTipoImportado(valor) {
  const v = normalizarTexto(valor);
  if (!v) return null; // célula vazia — quem chama decide o que fazer (ex: manter o tipo que já tinha)
  if (v === 'comum' || v === 'comuns') return 'comum';
  if (v === 'direta' || v === 'diretas') return 'direta';
  if (v === 'personalizada' || v === 'personalizadas') return 'personalizada';
  return TIPOS_CLIENTE.includes(v) ? v : null; // texto não reconhecido — também cai no fallback
}

function setTipoAtual(tipo) {
  if (!TIPOS_CLIENTE.includes(tipo)) return;
  TIPO_ATUAL = tipo;
  localStorage.setItem('picole_tipo_atual', tipo);
  const nomeTela = document.querySelector('.screen.active').id.replace('screen-', '');
  refreshAll();
  showScreen(nomeTela);
}

function tipoSelectorHtml() {
  return `<div class="pill-select">` + TIPOS_CLIENTE.map(t =>
    `<button class="${t === TIPO_ATUAL ? 'active' : ''}" onclick="setTipoAtual('${t}')">${TIPO_LABEL_CURTO[t]}</button>`
  ).join('') + `</div>`;
}

/* ---------------- Data store ---------------- */
let DB = null;

function defaultDB() {
  return {
    clientes: [],
    feriados: FERIADOS_PADRAO.map(([data, descricao]) => ({ id: uid('f'), data, descricao })),
    entregas: [],
    motoristas: [],
    categoriasFinanceiras: [],
    lancamentos: [],
  };
}

function migrarTipos() {
  // dados antigos (antes dos tipos de cliente) viram "personalizada"
  DB.clientes.forEach(c => { if (!c.tipo || !TIPOS_CLIENTE.includes(c.tipo)) c.tipo = 'personalizada'; });
  DB.entregas.forEach(e => { if (!e.tipo || !TIPOS_CLIENTE.includes(e.tipo)) e.tipo = 'personalizada'; });
  // corrige categorias financeiras salvas com tipo "saida" em vez de "despesa" (bug da criação rápida)
  if (DB.categoriasFinanceiras) {
    DB.categoriasFinanceiras.forEach(c => { if (c.tipo === 'saida') c.tipo = 'despesa'; });
  }
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
      if (!DB.motoristas) DB.motoristas = [];
      if (!DB.categoriasFinanceiras) DB.categoriasFinanceiras = [];
      if (!DB.lancamentos) DB.lancamentos = [];
      migrarTipos();
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

/* =========================================================
   Sincronização na nuvem (opcional) — Firestore
   - clientes/feriados: um "blob" só, gerenciado pelo admin
   - entregas: um documento por entrega, mesclados por updatedAt
     (permite colaboradores lançarem entregas sem mexer em clientes)
   - usuarios: papel de cada e-mail (admin | colaborador)
   ========================================================= */
/* =========================================================
   Sincronização na nuvem — Firestore, workspace único
   - login com Google OU e-mail/senha (Firebase Auth)
   - sem "codigo de sincronizacao": o acesso e definido so pelo papel
     atribuido ao e-mail da pessoa em usuarios/{email}
   - clientes/feriados: um "blob" so, gerenciado pelo admin
   - entregas: um documento por entrega, mesclados por updatedAt
   - papeis: admin (tudo) | colaborador (entregas/pagamentos) |
     convidado (so leitura, nao escreve nada)
   ========================================================= */
let AUTH_USER = null;   // { email, nome } quando logado, ou null
let MEU_PAPEL = null;   // 'admin' | 'colaborador' | 'convidado' | null
let firebaseReady = null;
let pushTimer = null;
let lastCloudPushAt = 0;

function ensureFirebase() {
  if (firebaseReady) return firebaseReady;
  firebaseReady = new Promise((resolve, reject) => {
    const s1 = document.createElement('script');
    s1.src = FIREBASE_APP_URL;
    s1.onload = () => {
      const s2 = document.createElement('script');
      s2.src = FIREBASE_FIRESTORE_URL;
      s2.onload = () => {
        const s3 = document.createElement('script');
        s3.src = FIREBASE_AUTH_URL;
        s3.onload = () => {
          try {
            if (!firebase.apps || firebase.apps.length === 0) {
              firebase.initializeApp(FIREBASE_CONFIG);
              firebase.auth().onAuthStateChanged((user) => {
                AUTH_USER = user ? { email: user.email, nome: user.displayName || user.email } : null;
                atualizarPapelEDataAposLogin();
              });
              firebase.auth().getRedirectResult().catch((err) => {
                console.error('login redirect falhou', err);
                toast('Erro no login: ' + (err && err.code ? err.code : 'desconhecido'));
              });
            }
            resolve();
          } catch (err) {
            reject(err);
          }
        };
        s3.onerror = () => reject(new Error('Falha ao carregar Firebase Auth (sem internet?)'));
        document.head.appendChild(s3);
      };
      s2.onerror = () => reject(new Error('Falha ao carregar Firestore (sem internet?)'));
      document.head.appendChild(s2);
    };
    s1.onerror = () => reject(new Error('Falha ao carregar Firebase (sem internet?)'));
    document.head.appendChild(s1);
  });
  return firebaseReady;
}

// tenta iniciar a sessao ja salva (se a pessoa tiver logado antes),
// silenciosamente, assim que o app abre.
function tentarRetomarLogin() {
  ensureFirebase().catch((err) => console.error('firebase init falhou', err));
}

/* ---------------- login ---------------- */
async function loginComGoogle() {
  try {
    await ensureFirebase();
    const provider = new firebase.auth.GoogleAuthProvider();
    try {
      await firebase.auth().signInWithPopup(provider);
      toast('Login feito!');
    } catch (popupErr) {
      const bloqueado = popupErr && ['auth/popup-blocked', 'auth/cancelled-popup-request',
        'auth/operation-not-supported-in-this-environment'].includes(popupErr.code);
      if (bloqueado) {
        toast('Abrindo tela de login…');
        await firebase.auth().signInWithRedirect(provider);
      } else if (popupErr && popupErr.code === 'auth/popup-closed-by-user') {
        // fechou de proposito, sem mensagem
      } else {
        throw popupErr;
      }
    }
  } catch (err) {
    console.error(err);
    toast('Erro no login: ' + (err && err.code ? err.code : 'desconhecido'));
  }
}

async function loginComEmailSenha() {
  const email = document.getElementById('f-login-email').value.trim();
  const senha = document.getElementById('f-login-senha').value;
  if (!email || !senha) { toast('Preencha e-mail e senha.'); return; }
  try {
    await ensureFirebase();
    await firebase.auth().signInWithEmailAndPassword(email, senha);
    toast('Login feito!');
  } catch (err) {
    console.error(err);
    if (err && err.code === 'auth/user-not-found') {
      toast('Não existe conta com esse e-mail. Toque em "Criar conta" se você foi convidado.');
    } else if (err && err.code === 'auth/wrong-password') {
      toast('Senha incorreta.');
    } else {
      toast('Erro no login: ' + (err && err.code ? err.code : 'desconhecido'));
    }
  }
}

async function criarContaEmailSenha() {
  const email = document.getElementById('f-login-email').value.trim();
  const senha = document.getElementById('f-login-senha').value;
  if (!email || !senha) { toast('Preencha e-mail e senha.'); return; }
  if (senha.length < 6) { toast('A senha precisa ter pelo menos 6 caracteres.'); return; }
  try {
    await ensureFirebase();
    await firebase.auth().createUserWithEmailAndPassword(email, senha);
    toast('Conta criada! Já entrou automaticamente.');
  } catch (err) {
    console.error(err);
    if (err && err.code === 'auth/email-already-in-use') {
      toast('Já existe conta com esse e-mail. Use "Entrar".');
    } else {
      toast('Erro ao criar conta: ' + (err && err.code ? err.code : 'desconhecido'));
    }
  }
}

function logout() {
  if (!firebaseReady) return;
  ensureFirebase().then(() => firebase.auth().signOut()).catch(() => {});
  AUTH_USER = null;
  MEU_PAPEL = null;
  renderSyncStatus();
}

async function atualizarPapelEDataAposLogin() {
  if (!AUTH_USER) { MEU_PAPEL = null; renderSyncStatus(); return; }
  try {
    const doc = await firebase.firestore()
      .collection('sincronizacoes').doc(WORKSPACE_ID)
      .collection('usuarios').doc(AUTH_USER.email).get();
    MEU_PAPEL = doc.exists ? doc.data().papel : null;
    if (MEU_PAPEL) {
      await sincronizar(true);
    }
    if (AREA_ATUAL === 'financeiro' && !podeVerResumoEFinanceiro()) {
      trocarArea('entregas');
    }
  } catch (err) {
    console.error('falha ao checar papel', err);
    toast('Erro ao conectar com a nuvem: ' + (err && err.code ? err.code : (err && err.message) || 'desconhecido'));
  }
  renderSyncStatus();
}

/* ---------------- gestão de usuários (só admin) ---------------- */
async function listarUsuarios() {
  await ensureFirebase();
  const snap = await firebase.firestore()
    .collection('sincronizacoes').doc(WORKSPACE_ID)
    .collection('usuarios').get();
  return snap.docs.map(d => d.data());
}

async function adicionarUsuario(email, papel) {
  await ensureFirebase();
  await firebase.firestore()
    .collection('sincronizacoes').doc(WORKSPACE_ID)
    .collection('usuarios').doc(email)
    .set({ email, papel, adicionadoEm: Date.now() });
}

async function removerUsuarioCloud(email) {
  await ensureFirebase();
  await firebase.firestore()
    .collection('sincronizacoes').doc(WORKSPACE_ID)
    .collection('usuarios').doc(email)
    .delete();
}

// Matriz de permissões (4 papéis: admin, gerente, colaborador, convidado).
// Sem login = app local, uso livre (equivalente a admin).
function podeGerenciarUsuarios() {
  return !AUTH_USER || MEU_PAPEL === 'admin';
}
function podeImportar() {
  return !AUTH_USER || MEU_PAPEL === 'admin';
}
function podeExportar() {
  return !AUTH_USER || MEU_PAPEL === 'admin' || MEU_PAPEL === 'gerente';
}
function podeGerenciar() {
  // lojas e feriados: admin e gerente
  return !AUTH_USER || MEU_PAPEL === 'admin' || MEU_PAPEL === 'gerente';
}
function podeEditarEntregas() {
  // admin, gerente e colaborador podem; convidado (so leitura) nao pode.
  return !AUTH_USER || MEU_PAPEL === 'admin' || MEU_PAPEL === 'gerente' || MEU_PAPEL === 'colaborador';
}
function podeConfirmarPagamento() {
  // colaborador NAO pode confirmar pagamentos, so admin e gerente
  return !AUTH_USER || MEU_PAPEL === 'admin' || MEU_PAPEL === 'gerente';
}
function podeVerResumoEFinanceiro() {
  // colaborador e convidado nao tem acesso (nem leitura) ao resumo
  // mensal nem a area financeira
  return !AUTH_USER || MEU_PAPEL === 'admin' || MEU_PAPEL === 'gerente';
}

/* ---------------- push/pull ---------------- */
function agendarPushNuvem() {
  if (!AUTH_USER || !MEU_PAPEL) return;
  clearTimeout(pushTimer);
  pushTimer = setTimeout(() => { pushParaNuvem(); }, 1500);
}

async function pushParaNuvem() {
  if (!AUTH_USER || !MEU_PAPEL) return;
  try {
    await ensureFirebase();
    const now = Date.now();
    const ref = firebase.firestore().collection('sincronizacoes').doc(WORKSPACE_ID);

    if (podeGerenciar()) {
      await ref.set({
        clientesJson: JSON.stringify(DB.clientes),
        feriadosJson: JSON.stringify(DB.feriados),
        motoristasJson: JSON.stringify(DB.motoristas),
        updatedAt: now,
      }, { merge: true });
      await ref.collection('financeiroConfig').doc('categorias').set({
        json: JSON.stringify(DB.categoriasFinanceiras),
        updatedAt: now,
      });
    }

    if (podeEditarEntregas()) {
      const alteradas = DB.entregas.filter(e => (e.updatedAt || 0) > lastCloudPushAt);
      if (alteradas.length > 0) {
        const batch = firebase.firestore().batch();
        alteradas.forEach(e => {
          batch.set(ref.collection('entregas').doc(e.id), e);
        });
        await batch.commit();
      }
    }

    if (podeGerenciar()) {
      const lancAlterados = DB.lancamentos.filter(l => (l.updatedAt || 0) > lastCloudPushAt);
      if (lancAlterados.length > 0) {
        const batch = firebase.firestore().batch();
        lancAlterados.forEach(l => {
          batch.set(ref.collection('lancamentos').doc(l.id), l);
        });
        await batch.commit();
      }
    }

    lastCloudPushAt = now;
    localStorage.setItem('picole_last_push_at', String(now));
    localStorage.setItem(LAST_SYNC_KEY, String(now));
    if (typeof renderSyncStatus === 'function') renderSyncStatus();
  } catch (err) {
    console.error('push para nuvem falhou', err);
  }
}

function mesclarEntregasPorId(locais, daNuvem) {
  const mapa = new Map();
  locais.forEach(e => mapa.set(e.id, e));
  daNuvem.forEach(e => {
    const local = mapa.get(e.id);
    if (!local || (e.updatedAt || 0) > (local.updatedAt || 0)) {
      mapa.set(e.id, e);
    }
  });
  return Array.from(mapa.values());
}

async function sincronizar(silencioso) {
  if (!AUTH_USER || !MEU_PAPEL) return;
  if (!silencioso) toast('Sincronizando…');
  try {
    await ensureFirebase();
    const ref = firebase.firestore().collection('sincronizacoes').doc(WORKSPACE_ID);
    const snap = await ref.get();
    const cloudData = snap.exists ? snap.data() : null;

    if (cloudData) {
      if (cloudData.clientesJson) DB.clientes = JSON.parse(cloudData.clientesJson);
      if (cloudData.feriadosJson) DB.feriados = JSON.parse(cloudData.feriadosJson);
      if (cloudData.motoristasJson) DB.motoristas = JSON.parse(cloudData.motoristasJson);
    }

    if (podeVerResumoEFinanceiro()) {
      const catDoc = await ref.collection('financeiroConfig').doc('categorias').get();
      if (catDoc.exists && catDoc.data().json) DB.categoriasFinanceiras = JSON.parse(catDoc.data().json);
    }

    const entregasSnap = await ref.collection('entregas').get();
    const daNuvem = entregasSnap.docs.map(d => d.data());
    DB.entregas = mesclarEntregasPorId(DB.entregas, daNuvem);

    if (podeVerResumoEFinanceiro()) {
      const lancSnap = await ref.collection('lancamentos').get();
      const lancNuvem = lancSnap.docs.map(d => d.data());
      DB.lancamentos = mesclarEntregasPorId(DB.lancamentos, lancNuvem);
    }

    migrarTipos(); // dados vindos da nuvem podem ser de antes dos tipos de cliente existirem

    localStorage.setItem(STORAGE_KEY, JSON.stringify(DB));
    if (typeof refreshAll === 'function') refreshAll();

    await pushParaNuvem();
    if (!silencioso) toast('Sincronização concluída.');
  } catch (err) {
    console.error('sincronizar falhou', err);
    if (!silencioso) toast('Não foi possível sincronizar. Verifique sua internet.');
  }
  if (typeof renderSyncStatus === 'function') renderSyncStatus();
}

function sincronizarAgora() {
  sincronizar(false);
}

/* ---------------- UI de sincronização ---------------- */
function renderSyncStatus() {
  const el = document.getElementById('sync-status-card');
  if (!el) return;

  if (!AUTH_USER) {
    el.innerHTML = `
      <div class="settings-item">
        <div><div class="t">Sincronização e backup na nuvem</div><div class="d">Entre para ativar e usar em mais de um aparelho</div></div>
      </div>
      <button class="btn primary block" style="margin:8px 0;" onclick="loginComGoogle()">Entrar com Google</button>
      <div class="hint" style="margin:6px 0 8px;text-align:center;">ou com e-mail e senha</div>
      <div class="field">
        <input type="email" id="f-login-email" placeholder="E-mail">
      </div>
      <div class="field">
        <input type="password" id="f-login-senha" placeholder="Senha">
      </div>
      <div class="formbtns" style="margin-bottom:4px;">
        <button class="btn ghost block" onclick="criarContaEmailSenha()">Criar conta</button>
        <button class="btn ghost block" onclick="loginComEmailSenha()">Entrar</button>
      </div>
      <div class="hint">Criar conta só funciona se um administrador já convidou seu e-mail.</div>
    `;
    return;
  }

  const lastSync = localStorage.getItem(LAST_SYNC_KEY);
  const lastTxt = lastSync ? new Date(Number(lastSync)).toLocaleString('pt-BR') : 'ainda não';
  const papelTxt = MEU_PAPEL === 'admin' ? 'Administrador' : MEU_PAPEL === 'gerente' ? 'Gerente' :
    MEU_PAPEL === 'colaborador' ? 'Colaborador' :
    MEU_PAPEL === 'convidado' ? 'Convidado (somente leitura)' : 'Sem acesso atribuído ainda';

  el.innerHTML = `
    <div class="settings-item">
      <div><div class="t">${escapeHtml(AUTH_USER.nome)}</div><div class="d">${escapeHtml(AUTH_USER.email)} · ${papelTxt}</div></div>
    </div>
    ${MEU_PAPEL ? `
      <div class="settings-item">
        <div><div class="t">Sincronização ativa</div><div class="d">Última sincronização: ${lastTxt}</div></div>
      </div>
      <button class="btn ghost block" style="margin:4px 0 8px;" onclick="sincronizarAgora()">🔄 Sincronizar agora</button>
      ${MEU_PAPEL === 'admin' ? `<button class="btn ghost block" style="margin-bottom:8px;" onclick="abrirGestaoUsuarios()">👥 Gerenciar usuários</button>` : ''}
    ` : `
      <div class="hint" style="margin:8px 0;">Sua conta ainda não tem acesso liberado. Peça para um administrador te convidar pelo e-mail ${escapeHtml(AUTH_USER.email)}.</div>
    `}
    <button class="btn ghost block" onclick="logout()">Sair da conta</button>
  `;
}

/* ---------------- gestão de usuários (UI, só admin) ---------------- */
async function abrirGestaoUsuarios() {
  openModal(`
    <div class="modal-header">
      <h2>Usuários</h2>
      <button class="close-x" onclick="closeModal()">✕</button>
    </div>
    <div class="hint" style="margin-bottom:12px;">Carregando…</div>
  `);
  try {
    const usuarios = await listarUsuarios();
    const nomePapel = (p) => p === 'admin' ? 'Administrador' : p === 'gerente' ? 'Gerente' :
      p === 'colaborador' ? 'Colaborador' : 'Convidado (somente leitura)';
    const linhas = usuarios.length === 0
      ? emptyState('👥', 'Ninguém adicionado ainda', 'Use o botão abaixo para convidar alguém pelo e-mail do Google.')
      : usuarios.map(u => `
          <div class="card">
            <div class="card-row">
              <div>
                <div class="card-title">${escapeHtml(u.email)}</div>
                <div class="card-sub">${nomePapel(u.papel)}</div>
              </div>
              <button class="btn ghost small" onclick="confirmarRemoverUsuario('${escapeHtml(u.email)}')">Remover</button>
            </div>
          </div>
        `).join('');

    document.getElementById('modal-content').innerHTML = `
      <div class="modal-header">
        <h2>Usuários</h2>
        <button class="close-x" onclick="closeModal()">✕</button>
      </div>
      <div class="hint" style="margin-bottom:12px;">Administrador: acesso total. Gerente: tudo igual ao admin, exceto gerenciar usuários e importar planilha. Colaborador: lança/edita entregas, mas não confirma pagamentos nem vê resumo mensal/financeiro. Convidado: só visualiza o que o colaborador vê, sem editar nada.</div>
      <button class="btn primary block" style="margin-bottom:14px;" onclick="abrirAdicionarUsuario()">+ Convidar por e-mail</button>
      ${linhas}
    `;
  } catch (err) {
    console.error(err);
    document.getElementById('modal-content').innerHTML = `
      <div class="modal-header"><h2>Usuários</h2><button class="close-x" onclick="closeModal()">✕</button></div>
      <div class="hint">Não foi possível carregar a lista. Verifique sua internet.</div>
    `;
  }
}

function abrirAdicionarUsuario() {
  const html = `
    <div class="modal-header">
      <h2>Convidar usuário</h2>
      <button class="close-x" onclick="closeModal()">✕</button>
    </div>
    <div class="field">
      <label>E-mail da pessoa (Google ou o que ela for usar)</label>
      <input type="email" id="f-novo-usuario-email" placeholder="pessoa@gmail.com">
    </div>
    <div class="field">
      <label>Nível de acesso</label>
      <select id="f-novo-usuario-papel">
        <option value="colaborador">Colaborador — lança/edita entregas</option>
        <option value="convidado">Convidado — só visualiza, não altera nada</option>
        <option value="gerente">Gerente — tudo do admin, exceto usuários e importar</option>
        <option value="admin">Administrador — acesso total</option>
      </select>
    </div>
    <div class="hint" style="margin-bottom:14px;">A pessoa precisa entrar no app com esse e-mail (Google, ou criando uma senha) — não precisa de mais nada.</div>
    <div class="formbtns">
      <button class="btn ghost block" onclick="abrirGestaoUsuarios()">Voltar</button>
      <button class="btn primary block" onclick="salvarNovoUsuario()">Convidar</button>
    </div>
  `;
  openModal(html);
}

async function salvarNovoUsuario() {
  const email = document.getElementById('f-novo-usuario-email').value.trim().toLowerCase();
  const papel = document.getElementById('f-novo-usuario-papel').value;
  if (!email || !email.includes('@')) { toast('Digite um e-mail válido.'); return; }
  try {
    await adicionarUsuario(email, papel);
    toast('Usuário convidado.');
    abrirGestaoUsuarios();
  } catch (err) {
    console.error(err);
    toast('Não foi possível convidar. Verifique sua internet.');
  }
}

function confirmarRemoverUsuario(email) {
  if (!confirm(`Remover o acesso de ${email}?`)) return;
  removerUsuarioCloud(email)
    .then(() => { toast('Usuário removido.'); abrirGestaoUsuarios(); })
    .catch((err) => { console.error(err); toast('Não foi possível remover.'); });
}

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
  const completa = (s && s.length === 7) ? s + '-01' : s;
  const d = parseDate(completa);
  if (!d || isNaN(d.getTime())) return '—';
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
function entregasAtivas(tipo) {
  return DB.entregas.filter(e => !e.deletado && (tipo === undefined || e.tipo === tipo));
}

function clientesDoTipo(tipo) {
  return DB.clientes.filter(c => tipo === undefined || c.tipo === tipo);
}

function entregasDoCliente(nome, tipo) {
  return entregasAtivas(tipo)
    .filter(e => e.cliente === nome)
    .slice()
    .sort((a, b) => (a.data < b.data ? 1 : a.data > b.data ? -1 : 0)); // desc
}

function calcPainelCliente(cliente) {
  const hset = holidaySet();
  const lista = entregasDoCliente(cliente.nome, cliente.tipo).slice(0, 5);

  if (lista.length === 0) {
    return {
      ultimaEntrega: null, ultimaQtd: null, intervaloEstimado: cliente.intervaloPadrao || 7,
      proximaEntrega: null, diasAte: null, status: 'Sem histórico',
    };
  }

  const datas = lista.map(e => e.data);
  const qtds = lista.map(e => Number(e.quantidade) || 0);

  const intervalos = [];
  const amostras = []; // { taxa, peso } — média ponderada pelos dias úteis de cada intervalo
  for (let i = 0; i < datas.length - 1; i++) {
    const maisRecente = datas[i];
    const maisAntiga = datas[i + 1];
    const qtdAntiga = qtds[i + 1];
    const inicio = addDaysStr(maisAntiga, 1);
    const intervalo = networkDays(inicio, maisRecente, hset);
    intervalos.push(intervalo);
    if (qtdAntiga > 0 && intervalo > 0) {
      const taxa = qtdAntiga / intervalo;
      const peso = intervalo * (i === 0 ? 2 : 1); // último intervalo pesa o dobro
      amostras.push({ taxa, peso });
    }
  }

  let intervaloEstimado = null;
  if (amostras.length > 0) {
    const somaPesos = amostras.reduce((s, a) => s + a.peso, 0);
    const taxaMedia = somaPesos > 0 ? amostras.reduce((s, a) => s + a.taxa * a.peso, 0) / somaPesos : null;
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
  if (Array.isArray(e.parcelas) && e.parcelas.length > 0) {
    return e.parcelas.reduce((s, p) => s + (Number(p.valor) || 0), 0);
  }
  const q = Number(e.quantidade) || 0;
  const vu = Number(e.valorUnitario) || 0;
  return q * vu;
}
function entregaStatusPagamento(e) {
  // 'pago' | 'parcial' | 'pendente'
  if (Array.isArray(e.parcelas) && e.parcelas.length > 0) {
    const pagas = e.parcelas.filter(p => !!p.dataPagamento).length;
    if (pagas === 0) return 'pendente';
    if (pagas === e.parcelas.length) return 'pago';
    return 'parcial';
  }
  return e.dataPagamento ? 'pago' : 'pendente';
}
function getPendencias(tipo) {
  return entregasAtivas(tipo)
    .filter(e => !e.dataPagamento && valorTotalEntrega(e) > 0)
    .map(e => ({ id: e.id, cliente: e.cliente, data: e.data, valor: valorTotalEntrega(e) }))
    .sort((a, b) => (a.data < b.data ? -1 : a.data > b.data ? 1 : 0));
}
function getTotalPendente(tipo) {
  return getPendencias(tipo).reduce((s, p) => s + p.valor, 0);
}

/* ---------------- Pagamentos parcelados (Entregas Comuns) ---------------- */
function getParcelasComuns() {
  // achata todas as parcelas de entregas do tipo "comum" em linhas individuais
  const hoje = todayStr();
  const linhas = [];
  entregasAtivas('comum').forEach(e => {
    (e.parcelas || []).forEach(p => {
      if (p.dataPagamento) return; // só pendentes aqui
      const vencida = !!p.vencimento && p.vencimento < hoje;
      linhas.push({
        entregaId: e.id, parcelaId: p.id, cliente: e.cliente, dataEntrega: e.data,
        quantidade: e.quantidade, valor: Number(p.valor) || 0, vencimento: p.vencimento,
        status: vencida ? 'vencida' : 'a_vencer',
      });
    });
  });
  return linhas;
}
function getResumoPagamentosComunsPorCliente() {
  const linhas = getParcelasComuns();
  const mapa = new Map();
  linhas.forEach(l => {
    if (!mapa.has(l.cliente)) mapa.set(l.cliente, { cliente: l.cliente, vencidas: [], aVencer: [] });
    const g = mapa.get(l.cliente);
    if (l.status === 'vencida') g.vencidas.push(l); else g.aVencer.push(l);
  });
  return Array.from(mapa.values()).sort((a, b) => a.cliente.localeCompare(b.cliente));
}

/* ---------------- Resumo mensal ---------------- */
function getResumoMensal(tipo) {
  const clientesAtivos = clientesDoTipo(tipo).slice().sort((a, b) => a.nome.localeCompare(b.nome));
  const entregas = entregasAtivas(tipo);
  if (entregas.length === 0 || clientesAtivos.length === 0) {
    return { meses: [], linhas: [], totais: [] };
  }
  const datasOrdenadas = entregas.map(e => e.data).sort();
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
      return entregas
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

/* ---------------- Resumo Geral (juntando os 3 tipos) ---------------- */
function getResumoGeralMensal() {
  const entregas = entregasAtivas(); // todos os tipos
  if (entregas.length === 0) return { meses: [], linhas: [], totalGeral: { personalizada: 0, comum: 0, direta: 0, total: 0 } };

  const datasOrdenadas = entregas.map(e => e.data).sort();
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

  const linhas = meses.map(mesKey => {
    const fimMesDate = parseDate(mesKey);
    fimMesDate.setMonth(fimMesDate.getMonth() + 1);
    const fimMes = dateToKey(fimMesDate);
    const doMes = entregas.filter(e => e.data >= mesKey && e.data < fimMes);
    const porTipo = { personalizada: 0, comum: 0, direta: 0 };
    doMes.forEach(e => { porTipo[e.tipo] = (porTipo[e.tipo] || 0) + (Number(e.quantidade) || 0); });
    const total = porTipo.personalizada + porTipo.comum + porTipo.direta;
    return { mes: mesKey, ...porTipo, total };
  });

  const totalGeral = {
    personalizada: linhas.reduce((s, l) => s + l.personalizada, 0),
    comum: linhas.reduce((s, l) => s + l.comum, 0),
    direta: linhas.reduce((s, l) => s + l.direta, 0),
    total: linhas.reduce((s, l) => s + l.total, 0),
  };

  return { meses, linhas, totalGeral };
}

function renderResumoGeral() {
  const el = document.getElementById('resumo-geral-container');
  const { meses, linhas, totalGeral } = getResumoGeralMensal();
  if (meses.length === 0) {
    el.innerHTML = emptyState('📊', 'Sem dados ainda', 'Assim que houver entregas registradas em qualquer um dos tipos, o resumo geral aparece aqui.');
    return;
  }
  const bodyRows = linhas.slice().reverse().map(l => `
    <tr>
      <td>${fmtMesAno(l.mes)}</td>
      <td>${l.personalizada}</td>
      <td>${l.comum}</td>
      <td>${l.direta}</td>
      <td><strong>${l.total}</strong></td>
    </tr>
  `).join('');
  const totalRow = `
    <tr class="total-row">
      <td>Total geral</td>
      <td>${totalGeral.personalizada}</td>
      <td>${totalGeral.comum}</td>
      <td>${totalGeral.direta}</td>
      <td>${totalGeral.total}</td>
    </tr>
  `;
  el.innerHTML = `
    <div class="hint" style="margin-bottom:12px;">Quantidade de picolés entregues/vendidos por mês, somando os 3 tipos.</div>
    <div class="table-wrap">
      <table class="resumo">
        <thead><tr><th>Mês</th><th>${TIPO_LABEL_CURTO.personalizada}</th><th>${TIPO_LABEL_CURTO.comum}</th><th>${TIPO_LABEL_CURTO.direta}</th><th>Total</th></tr></thead>
        <tbody>${totalRow}${bodyRows}</tbody>
      </table>
    </div>
  `;
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
  'resumo-geral': ['Resumo geral', 'Personalizadas + Comuns + Diretas'],
  feriados: ['Feriados', 'Dias ignorados no cálculo'],
  motoristas: ['Motoristas', 'Cadastro de motoristas'],
  'fin-lancamentos': ['Lançamentos', 'Entradas e saídas recentes'],
  'fin-relatorio': ['Relatório', 'Entradas e saídas por período'],
  'fin-resumo': ['Resumo mensal', 'Financeiro por mês'],
  'fin-mais': ['Mais', 'Configurações do financeiro'],
  'fin-categorias': ['Categorias', 'Tipos de entrada e despesa'],
};

const AREAS = {
  entregas: ['painel', 'entregas', 'pagamentos', 'mais', 'clientes', 'resumo', 'resumo-geral', 'feriados', 'motoristas'],
  financeiro: ['fin-lancamentos', 'fin-relatorio', 'fin-resumo', 'fin-mais', 'fin-categorias'],
};
const TAB_MAP_POR_AREA = {
  entregas: { painel: 'painel', entregas: 'entregas', pagamentos: 'pagamentos' },
  financeiro: { 'fin-lancamentos': 'fin-lancamentos', 'fin-relatorio': 'fin-relatorio', 'fin-resumo': 'fin-resumo' },
};
const TAB_PADRAO_POR_AREA = { entregas: null, financeiro: null };
const TELA_INICIAL_POR_AREA = { entregas: 'painel', financeiro: 'fin-lancamentos' };

let AREA_ATUAL = localStorage.getItem('picole_area_atual') || 'entregas';
if (!AREAS[AREA_ATUAL]) AREA_ATUAL = 'entregas';

function areaDaTela(name) {
  return AREAS.financeiro.includes(name) ? 'financeiro' : 'entregas';
}

function abrirMaisDaAreaAtual() {
  showScreen(AREA_ATUAL === 'financeiro' ? 'fin-mais' : 'mais');
}

function abrirSeletorAreas() {
  const html = `
    <div class="modal-header">
      <h2>Áreas</h2>
      <button class="close-x" onclick="closeModal()">✕</button>
    </div>
    <div class="card" onclick="closeModal(); trocarArea('entregas')" style="cursor:pointer;margin-bottom:10px;">
      <div class="card-row">
        <div><div class="card-title">🚚 Entregas</div><div class="card-sub">Painel, entregas, pagamentos e cadastros</div></div>
        <div>›</div>
      </div>
    </div>
    ${podeVerResumoEFinanceiro() ? `
    <div class="card" onclick="closeModal(); trocarArea('financeiro')" style="cursor:pointer;">
      <div class="card-row">
        <div><div class="card-title">💰 Financeiro</div><div class="card-sub">Lançamentos, relatórios e resumo mensal</div></div>
        <div>›</div>
      </div>
    </div>` : ''}
  `;
  openModal(html);
}

function trocarArea(area) {
  if (!AREAS[area]) return;
  AREA_ATUAL = area;
  localStorage.setItem('picole_area_atual', area);
  showScreen(TELA_INICIAL_POR_AREA[area]);
}

function atualizarTabbarPorArea() {
  document.querySelectorAll('.tab.area-entregas').forEach(t => t.style.display = AREA_ATUAL === 'entregas' ? '' : 'none');
  document.querySelectorAll('.tab.area-financeiro').forEach(t => t.style.display = AREA_ATUAL === 'financeiro' ? '' : 'none');
}

function showScreen(name) {
  if ((name === 'resumo' || name === 'resumo-geral') && !podeVerResumoEFinanceiro()) {
    toast('Você não tem acesso ao Resumo Mensal.');
    name = 'mais';
  }
  if (areaDaTela(name) === 'financeiro' && !podeVerResumoEFinanceiro()) {
    toast('Você não tem acesso à área financeira.');
    name = 'painel';
  }

  AREA_ATUAL = areaDaTela(name);
  localStorage.setItem('picole_area_atual', AREA_ATUAL);
  atualizarTabbarPorArea();

  document.querySelectorAll('.screen').forEach(s => s.classList.remove('active'));
  document.getElementById('screen-' + name).classList.add('active');

  document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  const tabMap = TAB_MAP_POR_AREA[AREA_ATUAL] || {};
  const activeTab = tabMap[name] || TAB_PADRAO_POR_AREA[AREA_ATUAL];
  const tabEl = document.querySelector(`.tab[data-tab="${activeTab}"]`);
  if (tabEl) tabEl.classList.add('active');

  const [title, sub] = TITULOS[name] || [name, ''];
  const TELAS_COM_TIPO = ['painel', 'entregas', 'pagamentos', 'clientes'];
  document.getElementById('topbar-title').textContent = title;
  document.getElementById('topbar-sub').textContent = TELAS_COM_TIPO.includes(name)
    ? TIPO_LABEL[TIPO_ATUAL] : sub;

  document.getElementById('fab-entrega').style.display = (name === 'entregas' && podeEditarEntregas()) ? 'flex' : 'none';
  document.getElementById('fab-cliente').style.display = (name === 'clientes' && podeGerenciar()) ? 'flex' : 'none';
  document.getElementById('fab-feriado').style.display = (name === 'feriados' && podeGerenciar()) ? 'flex' : 'none';
  document.getElementById('fab-motorista').style.display = (name === 'motoristas' && podeGerenciar()) ? 'flex' : 'none';
  document.getElementById('fab-lancamento').style.display = (name === 'fin-lancamentos' && podeGerenciar()) ? 'flex' : 'none';
  document.getElementById('fab-categoria').style.display = (name === 'fin-categorias' && podeGerenciar()) ? 'flex' : 'none';

  renderScreen(name);
}

function renderScreen(name) {
  if (name === 'painel') renderPainel();
  else if (name === 'entregas') renderEntregas();
  else if (name === 'pagamentos') renderPagamentos();
  else if (name === 'clientes') renderClientes();
  else if (name === 'resumo') renderResumo();
  else if (name === 'resumo-geral') renderResumoGeral();
  else if (name === 'feriados') renderFeriados();
  else if (name === 'motoristas') renderMotoristas();
  else if (name === 'fin-lancamentos') renderFinLancamentos();
  else if (name === 'fin-relatorio') renderFinRelatorio();
  else if (name === 'fin-resumo') renderFinResumoMensal();
  else if (name === 'fin-categorias') renderCategoriasFinanceiras();
}

function refreshAll() {
  renderPainel(); renderEntregas(); renderPagamentos();
  renderClientes(); renderResumo(); renderResumoGeral(); renderFeriados(); renderMotoristas(); renderSyncStatus();
  if (podeVerResumoEFinanceiro()) {
    renderFinLancamentos();
    if (document.getElementById('screen-fin-relatorio').classList.contains('active')) renderFinRelatorioResultado();
    if (document.getElementById('screen-fin-resumo').classList.contains('active')) renderFinResumoMensal();
    renderCategoriasFinanceiras();
  }
  atualizarVisibilidadePermissoes();
}

function atualizarVisibilidadePermissoes() {
  const itemResumo = document.getElementById('item-resumo');
  if (itemResumo) itemResumo.style.display = podeVerResumoEFinanceiro() ? '' : 'none';
  const itemResumoGeral = document.getElementById('item-resumo-geral');
  if (itemResumoGeral) itemResumoGeral.style.display = podeVerResumoEFinanceiro() ? '' : 'none';
  const itemImportar = document.getElementById('item-importar');
  if (itemImportar) itemImportar.style.display = podeImportar() ? '' : 'none';
  const itemExportar = document.getElementById('item-exportar');
  if (itemExportar) itemExportar.style.display = podeExportar() ? '' : 'none';
  const itemRestaurar = document.getElementById('item-restaurar-backup');
  if (itemRestaurar) itemRestaurar.style.display = podeImportar() ? '' : 'none';
  const itemDesfazer = document.getElementById('item-desfazer-restauracao');
  if (itemDesfazer) itemDesfazer.style.display = podeImportar() ? '' : 'none';
  const itemFinCategorias = document.getElementById('item-fin-categorias');
  if (itemFinCategorias) itemFinCategorias.style.display = podeGerenciar() ? '' : 'none';
  atualizarTabbarPorArea();
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
  const ativos = clientesDoTipo(TIPO_ATUAL).filter(c => c.ativo !== false);
  if (ativos.length === 0) {
    el.innerHTML = tipoSelectorHtml() + emptyState('🍭', 'Nenhuma loja cadastrada ainda', 'Cadastre suas lojas em Mais → Lojas para começar a ver as previsões aqui.');
    return;
  }
  const linhas = ativos.map(c => ({ c, r: calcPainelCliente(c) }));
  linhas.sort((a, b) => {
    if (a.r.diasAte === null && b.r.diasAte === null) return a.c.nome.localeCompare(b.c.nome);
    if (a.r.diasAte === null) return 1;
    if (b.r.diasAte === null) return -1;
    return a.r.diasAte - b.r.diasAte;
  });

  el.innerHTML = tipoSelectorHtml() + linhas.map(({ c, r }) => {
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
  const lista = entregasDoCliente(cliente.nome, cliente.tipo); // ja vem ordenada da mais recente pra mais antiga

  const totalQtd = lista.reduce((s, e) => s + (Number(e.quantidade) || 0), 0);

  const linhasHtml = lista.length === 0
    ? emptyState('🚚', 'Nenhuma entrega registrada', 'Ainda não há entregas lançadas para essa loja.')
    : lista.map(e => {
        const valor = valorTotalEntrega(e);
        const status = entregaStatusPagamento(e);
        const badgeCls = status === 'pago' ? 'green' : status === 'parcial' ? 'amber' : 'amber';
        const badgeTxt = status === 'pago' ? 'Pago' : status === 'parcial' ? 'Parcial' : 'Pendente';
        return `
          <div class="card" onclick="closeModal(); openEntregaForm('${e.id}')" style="cursor:pointer;">
            <div class="card-row">
              <div>
                <div class="card-title">${fmtDateBR(e.data)}</div>
                <div class="card-sub">${e.quantidade} un.${valor > 0 ? ' · ' + fmtMoney(valor) : ''}</div>
              </div>
              <span class="badge ${badgeCls}">${badgeTxt}</span>
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
  const lista = entregasAtivas(TIPO_ATUAL).slice().sort((a, b) => (a.data < b.data ? 1 : a.data > b.data ? -1 : 0));
  if (lista.length === 0) {
    el.innerHTML = tipoSelectorHtml() + emptyState('🚚', 'Nenhuma entrega registrada', 'Toque no botão + para lançar a primeira entrega.');
    return;
  }
  el.innerHTML = tipoSelectorHtml() + lista.map(e => {
    const valorTotal = valorTotalEntrega(e);
    const status = entregaStatusPagamento(e);
    const badgeCls = status === 'pago' ? 'green' : 'amber';
    const badgeTxt = status === 'pago' ? 'Pago' : status === 'parcial' ? 'Parcial' : 'Pendente';
    return `
      <div class="card" onclick="openEntregaForm('${e.id}')" style="cursor:pointer;">
        <div class="card-row">
          <div>
            <div class="card-title">${escapeHtml(e.cliente)}</div>
            <div class="card-sub">${fmtDateBR(e.data)} · ${e.quantidade} un. ${valorTotal > 0 ? '· ' + fmtMoney(valorTotal) : ''}${e.motorista ? ' · 🚐 ' + escapeHtml(e.motorista) : ''}</div>
          </div>
          <span class="badge ${badgeCls}">${badgeTxt}</span>
        </div>
      </div>`;
  }).join('');
}

function openEntregaForm(id) {
  if (!podeEditarEntregas()) { toast('Seu acesso é somente leitura.'); return; }
  const editando = !!id;
  const e = editando ? DB.entregas.find(x => x.id === id) : null;
  const tipo = editando ? e.tipo : TIPO_ATUAL;
  const ativos = clientesDoTipo(tipo).filter(c => c.ativo !== false);
  const options = ativos.map(c =>
    `<option value="${escapeHtml(c.nome)}" ${e && e.cliente === c.nome ? 'selected' : ''}>${escapeHtml(c.nome)}</option>`
  ).join('');

  if (ativos.length === 0) {
    toast(`Cadastre uma loja em "${TIPO_LABEL[tipo]}" antes de lançar uma entrega.`);
    return;
  }

  if (tipo === 'comum') {
    openEntregaFormComum(e, tipo, options);
    return;
  }

  const valorInicial = editando ? e.valorUnitario : (ativos[0] ? (ativos[0].valorPadrao || '') : '');
  const qtdInicial = editando ? e.quantidade : (ativos[0] ? (ativos[0].quantidadePadrao || '') : '');

  const html = `
    <div class="modal-header">
      <h2>${editando ? 'Editar entrega' : 'Nova entrega'}</h2>
      <button class="close-x" onclick="closeModal()">✕</button>
    </div>
    <div class="hint" style="margin-bottom:10px;">${TIPO_LABEL[tipo]}</div>
    <div class="field">
      <label>Loja</label>
      <select id="f-cliente" ${editando ? '' : 'onchange="atualizarPadroesLojaAuto()"'}>${options}</select>
    </div>
    <div class="row2">
      <div class="field">
        <label>Data da entrega</label>
        <input type="date" id="f-data" value="${e ? e.data : todayStr()}">
      </div>
      <div class="field">
        <label>Quantidade</label>
        <input type="number" id="f-qtd" inputmode="numeric" value="${qtdInicial}" placeholder="0">
      </div>
    </div>
    <div class="field">
      <label>Motorista</label>
      <select id="f-motorista">${motoristaOptionsHtml(e ? e.motorista : '')}</select>
    </div>
    <div class="field">
      <label>Observações da entrega</label>
      <textarea id="f-obs" placeholder="Opcional">${e ? escapeHtml(e.observacoes || '') : ''}</textarea>
    </div>
    <div class="field">
      <label>Valor unitário (por picolé)</label>
      <input type="number" id="f-valor-unit" inputmode="decimal" step="0.01" value="${valorInicial}" placeholder="0,00">
      <div class="hint">${editando ? '' : 'Preenchido automaticamente com o valor padrão da loja, se houver. Pode ajustar.'}</div>
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

function motoristaOptionsHtml(selecionado) {
  const branco = `<option value="" ${!selecionado ? 'selected' : ''}>Sem motorista definido</option>`;
  const opcoes = motoristasAtivos().slice().sort((a, b) => a.nome.localeCompare(b.nome)).map(m =>
    `<option value="${escapeHtml(m.nome)}" ${selecionado === m.nome ? 'selected' : ''}>${escapeHtml(m.nome)}</option>`
  ).join('');
  return branco + opcoes;
}

function atualizarPadroesLojaAuto() {
  const cliente = document.getElementById('f-cliente').value;
  const tipo = TIPO_ATUAL;
  const c = DB.clientes.find(x => x.nome === cliente && x.tipo === tipo);
  const elValor = document.getElementById('f-valor-unit');
  const elQtd = document.getElementById('f-qtd');
  if (elValor && c && c.valorPadrao) elValor.value = c.valorPadrao;
  if (elQtd && c && c.quantidadePadrao) elQtd.value = c.quantidadePadrao;
}

function saveEntrega(id) {
  const cliente = document.getElementById('f-cliente').value;
  const data = document.getElementById('f-data').value;
  const quantidade = Number(document.getElementById('f-qtd').value) || 0;
  const motorista = document.getElementById('f-motorista').value || null;
  const observacoes = document.getElementById('f-obs').value.trim();
  const valorUnitario = Number(document.getElementById('f-valor-unit').value) || 0;
  const dataPagamento = document.getElementById('f-data-pagto').value || null;
  const obsPagamento = document.getElementById('f-obs-pagto').value.trim();

  if (!data || !cliente || quantidade <= 0) {
    toast('Preencha loja, data e uma quantidade maior que zero.');
    return;
  }

  let entregaSalva;
  if (id) {
    if (!confirm('Salvar as alterações feitas nesta entrega?')) {
      return;
    }
    const e = DB.entregas.find(x => x.id === id);
    Object.assign(e, { cliente, data, quantidade, motorista, observacoes, valorUnitario, dataPagamento, obsPagamento, updatedAt: Date.now() });
    entregaSalva = e;
  } else {
    entregaSalva = { id: uid('e'), tipo: TIPO_ATUAL, cliente, data, quantidade, motorista, observacoes, valorUnitario, dataPagamento, obsPagamento, updatedAt: Date.now(), deletado: false };
    DB.entregas.push(entregaSalva);
  }
  sincronizarLancamentoAutoEntrega(entregaSalva);
  saveDB();
  closeModal();
  refreshAll();
  toast('Entrega salva.');
}

/* ---------------- Formulário de Entregas Comuns (parcelado) ---------------- */
function getValorPadraoCliente(nome, tipo) {
  const c = DB.clientes.find(x => x.nome === nome && x.tipo === tipo);
  return c && c.valorPadrao ? Number(c.valorPadrao) : 0;
}

function gerarParcelasPadrao(n, valorTotal, dataBase, periodicidadeDias) {
  n = Math.max(1, Math.round(n) || 1);
  const periodo = Math.max(1, Math.round(periodicidadeDias) || 30);
  const valorParcela = Math.round((valorTotal / n) * 100) / 100;
  const parcelas = [];
  for (let i = 0; i < n; i++) {
    parcelas.push({
      id: uid('p'),
      valor: i === n - 1 ? Math.round((valorTotal - valorParcela * (n - 1)) * 100) / 100 : valorParcela,
      vencimento: addDaysStr(dataBase, periodo * (i + 1)),
      dataPagamento: null,
    });
  }
  return parcelas;
}

function parcelaRowHtml(p, idx) {
  return `
    <div class="card" style="padding:10px 12px;margin-bottom:8px;" data-parcela-row data-parcela-id="${p.id}">
      <div class="row2">
        <div class="field" style="margin-bottom:6px;">
          <label>Parcela ${idx + 1} — valor</label>
          <input type="number" step="0.01" class="f-parcela-valor" value="${p.valor}">
        </div>
        <div class="field" style="margin-bottom:6px;">
          <label>Vencimento</label>
          <input type="date" class="f-parcela-vencimento" value="${p.vencimento || ''}">
        </div>
      </div>
      <div class="field" style="margin-bottom:0;">
        <label>Data do pagamento (em branco = pendente)</label>
        <input type="date" class="f-parcela-pagamento" value="${p.dataPagamento || ''}">
      </div>
    </div>`;
}

function renderParcelasContainer(parcelas) {
  document.getElementById('parcelas-container').innerHTML = parcelas.map((p, i) => parcelaRowHtml(p, i)).join('');
}

function recalcularParcelas() {
  const valorTotal = Number(document.getElementById('f-valor-total').value) || 0;
  const n = Number(document.getElementById('f-num-parcelas').value) || 1;
  const periodicidade = Number(document.getElementById('f-periodicidade').value) || 30;
  const dataBase = document.getElementById('f-data').value || todayStr();
  localStorage.setItem('picole_periodicidade_dias', String(periodicidade));
  renderParcelasContainer(gerarParcelasPadrao(n, valorTotal, dataBase, periodicidade));
}

function atualizarValorTotalAutoComum() {
  const el = document.getElementById('f-valor-total');
  if (!el || el.dataset.editadoManual === '1') return;
  const cliente = document.getElementById('f-cliente').value;
  const qtd = Number(document.getElementById('f-qtd').value) || 0;
  const valorPadrao = getValorPadraoCliente(cliente, 'comum');
  el.value = valorPadrao ? Math.round(qtd * valorPadrao * 100) / 100 : el.value;
}

function atualizarPadroesComumAuto() {
  const cliente = document.getElementById('f-cliente').value;
  const c = DB.clientes.find(x => x.nome === cliente && x.tipo === 'comum');
  const elQtd = document.getElementById('f-qtd');
  if (elQtd && c && c.quantidadePadrao) elQtd.value = c.quantidadePadrao;
  atualizarValorTotalAutoComum();
}

function coletarParcelasDoForm() {
  return Array.from(document.querySelectorAll('[data-parcela-row]')).map(row => ({
    id: row.dataset.parcelaId,
    valor: Number(row.querySelector('.f-parcela-valor').value) || 0,
    vencimento: row.querySelector('.f-parcela-vencimento').value || null,
    dataPagamento: row.querySelector('.f-parcela-pagamento').value || null,
  }));
}

function openEntregaFormComum(e, tipo, options) {
  const editando = !!e;
  const periodicidadePadrao = Number(localStorage.getItem('picole_periodicidade_dias')) || 30;
  const parcelasIniciais = editando && Array.isArray(e.parcelas) && e.parcelas.length > 0
    ? e.parcelas
    : gerarParcelasPadrao(1, editando ? valorTotalEntrega(e) : 0, editando ? e.data : todayStr(), periodicidadePadrao);

  const html = `
    <div class="modal-header">
      <h2>${editando ? 'Editar entrega' : 'Nova entrega'}</h2>
      <button class="close-x" onclick="closeModal()">✕</button>
    </div>
    <div class="hint" style="margin-bottom:10px;">${TIPO_LABEL[tipo]}</div>
    <div class="field">
      <label>Loja</label>
      <select id="f-cliente" ${editando ? '' : 'onchange="atualizarPadroesComumAuto()"'}>${options}</select>
    </div>
    <div class="row2">
      <div class="field">
        <label>Data da entrega</label>
        <input type="date" id="f-data" value="${editando ? e.data : todayStr()}">
      </div>
      <div class="field">
        <label>Quantidade</label>
        <input type="number" id="f-qtd" inputmode="numeric" value="${editando ? e.quantidade : ''}" placeholder="0" ${editando ? '' : 'oninput="atualizarValorTotalAutoComum()"'}>
      </div>
    </div>
    <div class="field">
      <label>Motorista</label>
      <select id="f-motorista">${motoristaOptionsHtml(e ? e.motorista : '')}</select>
    </div>
    <div class="field">
      <label>Observações da entrega</label>
      <textarea id="f-obs" placeholder="Opcional">${editando ? escapeHtml(e.observacoes || '') : ''}</textarea>
    </div>
    <div class="section-title" style="margin-top:6px;">Pagamento parcelado</div>
    <div class="row2">
      <div class="field">
        <label>Valor total da entrega</label>
        <input type="number" id="f-valor-total" step="0.01" value="${editando ? valorTotalEntrega(e) : ''}" placeholder="0,00" oninput="this.dataset.editadoManual='1'">
      </div>
      <div class="field">
        <label>Nº de parcelas</label>
        <input type="number" id="f-num-parcelas" min="1" value="${parcelasIniciais.length}">
      </div>
    </div>
    <div class="field">
      <label>Vencer a cada quantos dias</label>
      <input type="number" id="f-periodicidade" min="1" value="${periodicidadePadrao}">
      <div class="hint">Ex: 30 = mensal, 7 = semanal. Usado ao gerar as parcelas abaixo.</div>
    </div>
    <button type="button" class="btn ghost block" style="margin-bottom:10px;" onclick="recalcularParcelas()">🔁 Gerar/recalcular parcelas</button>
    <div class="hint" style="margin-bottom:8px;">Gerar recria as parcelas abaixo (divide o valor e espaça os vencimentos pela periodicidade escolhida). Depois de gerar, você pode ajustar valor, vencimento e marcar cada parcela como paga individualmente.</div>
    <div id="parcelas-container">${parcelasIniciais.map((p, i) => parcelaRowHtml(p, i)).join('')}</div>
    <div class="formbtns" style="margin-top:10px;">
      <button class="btn ghost block" onclick="closeModal()">Cancelar</button>
      <button class="btn primary block" onclick="saveEntregaComum(${editando ? `'${e.id}'` : 'null'})">Salvar</button>
    </div>
    ${editando ? `<button class="danger-link" onclick="deleteEntrega('${e.id}')">Excluir esta entrega</button>` : ''}
  `;
  openModal(html);
}

function saveEntregaComum(id) {
  const cliente = document.getElementById('f-cliente').value;
  const data = document.getElementById('f-data').value;
  const quantidade = Number(document.getElementById('f-qtd').value) || 0;
  const motorista = document.getElementById('f-motorista').value || null;
  const observacoes = document.getElementById('f-obs').value.trim();
  const parcelas = coletarParcelasDoForm();

  if (!data || !cliente || quantidade <= 0) {
    toast('Preencha loja, data e uma quantidade maior que zero.');
    return;
  }
  if (parcelas.length === 0) {
    toast('Gere ao menos uma parcela de pagamento.');
    return;
  }
  if (parcelas.some(p => !p.vencimento)) {
    toast('Preencha a data de vencimento de todas as parcelas.');
    return;
  }

  let entregaSalva;
  if (id) {
    if (!confirm('Salvar as alterações feitas nesta entrega?')) return;
    const e = DB.entregas.find(x => x.id === id);
    Object.assign(e, { cliente, data, quantidade, motorista, observacoes, parcelas, updatedAt: Date.now() });
    entregaSalva = e;
  } else {
    entregaSalva = { id: uid('e'), tipo: 'comum', cliente, data, quantidade, motorista, observacoes, parcelas, updatedAt: Date.now(), deletado: false };
    DB.entregas.push(entregaSalva);
  }
  sincronizarParcelasComumFinanceiro(entregaSalva);
  saveDB();
  closeModal();
  refreshAll();
  toast('Entrega salva.');
}

function deleteEntrega(id) {
  if (!confirm('Excluir esta entrega? Essa ação não pode ser desfeita.')) return;
  const e = DB.entregas.find(x => x.id === id);
  if (e) {
    // exclusao "suave": mantem o registro marcado como apagado, em vez de
    // remove-lo de verdade, para que a exclusao tambem se propague na
    // proxima sincronizacao com outros aparelhos.
    e.deletado = true;
    e.updatedAt = Date.now();
    removerLancamentosDaEntrega(id);
  }
  saveDB();
  closeModal();
  refreshAll();
  toast('Entrega excluída.');
}

/* =========================================================
   PAGAMENTOS
   ========================================================= */
function renderPagamentos() {
  const el = document.getElementById('pagamentos-list');
  const kpiWrap = document.getElementById('pagamentos-kpi-wrap');

  if (TIPO_ATUAL === 'comum') {
    kpiWrap.style.display = 'none';
    document.getElementById('btn-exportar-imagem').style.display = 'none';
    document.getElementById('btn-historico-pagos').style.display = '';
    renderPagamentosComuns(el);
    return;
  }

  kpiWrap.style.display = '';
  document.getElementById('btn-exportar-imagem').style.display = '';
  document.getElementById('btn-historico-pagos').style.display = '';
  document.getElementById('total-pendente').textContent = fmtMoney(getTotalPendente(TIPO_ATUAL));
  const pend = getPendencias(TIPO_ATUAL);
  if (pend.length === 0) {
    el.innerHTML = tipoSelectorHtml() + emptyState('✅', 'Nada pendente', 'Todas as entregas com valor lançado já foram pagas.');
    return;
  }
  el.innerHTML = tipoSelectorHtml() + pend.map(p => `
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
      ${podeConfirmarPagamento() ? `<button class="btn ghost small" style="margin-top:10px;" onclick="abrirConfirmacaoPagamento('${p.id}')">Marcar como pago</button>` : ''}
    </div>
  `).join('');
}

function renderPagamentosComuns(el) {
  const grupos = getResumoPagamentosComunsPorCliente();
  const totalVencidas = grupos.reduce((s, g) => s + g.vencidas.reduce((a, l) => a + l.valor, 0), 0);
  const totalAVencer = grupos.reduce((s, g) => s + g.aVencer.reduce((a, l) => a + l.valor, 0), 0);

  const kpis = `
    <div class="row2" style="margin-bottom:14px;">
      <div class="kpi" style="background:linear-gradient(135deg,#E5502F,#9A3324);">
        <div class="label">Vencidas</div>
        <div class="value">${fmtMoney(totalVencidas)}</div>
      </div>
      <div class="kpi">
        <div class="label">A vencer</div>
        <div class="value">${fmtMoney(totalAVencer)}</div>
      </div>
    </div>`;

  if (grupos.length === 0) {
    el.innerHTML = tipoSelectorHtml() + emptyState('✅', 'Nada pendente', 'Todas as parcelas já foram pagas.');
    return;
  }

  el.innerHTML = tipoSelectorHtml() + kpis + grupos.map(g => {
    const totalCliente = [...g.vencidas, ...g.aVencer].reduce((s, l) => s + l.valor, 0);
    return `
      <div class="card" onclick="abrirDetalhePagamentosComuns('${escapeHtml(g.cliente)}')" style="cursor:pointer;">
        <div class="card-row">
          <div>
            <div class="card-title">${escapeHtml(g.cliente)}</div>
            <div class="card-sub">
              ${g.vencidas.length > 0 ? `<span style="color:var(--red-ink);font-weight:700;">${g.vencidas.length} vencida${g.vencidas.length > 1 ? 's' : ''}</span>` : ''}
              ${g.vencidas.length > 0 && g.aVencer.length > 0 ? ' · ' : ''}
              ${g.aVencer.length > 0 ? `${g.aVencer.length} a vencer` : ''}
            </div>
          </div>
          <div class="card-title">${fmtMoney(totalCliente)}</div>
        </div>
      </div>`;
  }).join('');
}

function abrirDetalhePagamentosComuns(clienteNome) {
  const grupos = getResumoPagamentosComunsPorCliente();
  const g = grupos.find(x => x.cliente === clienteNome);
  if (!g) return;
  const linha = (l) => `
    <div class="card">
      <div class="card-row">
        <div>
          <div class="card-title">${fmtDateBR(l.dataEntrega)}</div>
          <div class="card-sub">${l.quantidade} un. · vencimento ${fmtDateBR(l.vencimento)}</div>
        </div>
        <div style="text-align:right;">
          <div class="card-title" style="color:${l.status === 'vencida' ? 'var(--red-ink)' : 'var(--ink)'};">${fmtMoney(l.valor)}</div>
          <span class="badge ${l.status === 'vencida' ? 'red' : 'amber'}">${l.status === 'vencida' ? 'Vencida' : 'A vencer'}</span>
        </div>
      </div>
      ${podeConfirmarPagamento() ? `<button class="btn ghost small" style="margin-top:10px;" onclick="abrirConfirmacaoPagamentoParcela('${l.entregaId}','${l.parcelaId}')">Marcar parcela como paga</button>` : ''}
    </div>`;

  const html = `
    <div class="modal-header">
      <h2>${escapeHtml(clienteNome)}</h2>
      <button class="close-x" onclick="closeModal()">✕</button>
    </div>
    ${g.vencidas.length > 0 ? `
      <button class="btn primary block" style="margin-bottom:10px;" onclick="exportarPagamentosComunsImagem('${escapeHtml(clienteNome)}', true)">📤 Exportar imagem só das vencidas</button>
    ` : ''}
    <button class="btn ghost block" style="margin-bottom:14px;" onclick="exportarPagamentosComunsImagem('${escapeHtml(clienteNome)}', false)">📤 Exportar imagem de todas as pendentes</button>
    ${g.vencidas.length > 0 ? `<div class="section-title">Vencidas</div>${g.vencidas.map(linha).join('')}` : ''}
    ${g.aVencer.length > 0 ? `<div class="section-title">A vencer</div>${g.aVencer.map(linha).join('')}` : ''}
  `;
  openModal(html);
}

function abrirConfirmacaoPagamentoParcela(entregaId, parcelaId) {
  if (!podeConfirmarPagamento()) { toast('Você não tem permissão para confirmar pagamentos.'); return; }
  const e = DB.entregas.find(x => x.id === entregaId);
  if (!e) return;
  const p = (e.parcelas || []).find(x => x.id === parcelaId);
  if (!p) return;
  const html = `
    <div class="modal-header">
      <h2>Confirmar pagamento</h2>
      <button class="close-x" onclick="closeModal()">✕</button>
    </div>
    <div class="hint" style="margin-bottom:14px;">${escapeHtml(e.cliente)} · Entrega de ${fmtDateBR(e.data)} · ${fmtMoney(p.valor)}</div>
    <div class="field">
      <label>Data do pagamento</label>
      <input type="date" id="f-confirma-data-pagto" value="${todayStr()}">
    </div>
    <div class="formbtns">
      <button class="btn ghost block" onclick="closeModal()">Cancelar</button>
      <button class="btn primary block" onclick="confirmarPagamentoParcela('${entregaId}','${parcelaId}')">Confirmar pagamento</button>
    </div>
  `;
  openModal(html);
}

function confirmarPagamentoParcela(entregaId, parcelaId) {
  const e = DB.entregas.find(x => x.id === entregaId);
  if (!e) return;
  const p = (e.parcelas || []).find(x => x.id === parcelaId);
  if (!p) return;
  const data = document.getElementById('f-confirma-data-pagto').value;
  if (!data) { toast('Escolha a data do pagamento.'); return; }
  p.dataPagamento = data;
  e.updatedAt = Date.now();
  sincronizarParcelasComumFinanceiro(e);
  saveDB();
  closeModal();
  refreshAll();
  toast('Pagamento da parcela confirmado.');
}

function abrirConfirmacaoPagamento(entregaId) {
  if (!podeConfirmarPagamento()) { toast('Você não tem permissão para confirmar pagamentos.'); return; }
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
  e.updatedAt = Date.now();
  sincronizarLancamentoAutoEntrega(e);
  saveDB();
  closeModal();
  refreshAll();
  toast('Pagamento confirmado.');
}

function dataPagamentoEfetiva(e) {
  if (Array.isArray(e.parcelas) && e.parcelas.length > 0) {
    return e.parcelas.map(p => p.dataPagamento).filter(Boolean).sort().slice(-1)[0] || null;
  }
  return e.dataPagamento || null;
}

function abrirHistoricoPagos() {
  const pagos = entregasAtivas(TIPO_ATUAL)
    .filter(e => entregaStatusPagamento(e) === 'pago')
    .slice()
    .sort((a, b) => {
      const da = dataPagamentoEfetiva(a) || '';
      const db = dataPagamentoEfetiva(b) || '';
      return da < db ? 1 : da > db ? -1 : 0;
    });

  const linhasHtml = pagos.length === 0
    ? emptyState('🧾', 'Nenhum pagamento registrado', 'Assim que marcar uma entrega como paga, ela aparece aqui.')
    : pagos.map(e => {
        const dataPag = dataPagamentoEfetiva(e);
        return `
        <div class="card" onclick="closeModal(); openEntregaForm('${e.id}')" style="cursor:pointer;">
          <div class="card-row">
            <div>
              <div class="card-title">${escapeHtml(e.cliente)}</div>
              <div class="card-sub">Entrega de ${fmtDateBR(e.data)}${dataPag ? ' · Pago em ' + fmtDateBR(dataPag) : ''}</div>
            </div>
            <div class="card-title" style="color:var(--teal-700);">${fmtMoney(valorTotalEntrega(e))}</div>
          </div>
        </div>`;
      }).join('');

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
  const pend = getPendencias(TIPO_ATUAL);
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
  ctx.fillText(fmtMoney(getTotalPendente(TIPO_ATUAL)), PAD, 138);

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

/* ---------------- Exportar pendências de Entregas Comuns como imagem (por cliente) ---------------- */
async function exportarPagamentosComunsImagem(clienteNome, apenasVencidas) {
  const grupos = getResumoPagamentosComunsPorCliente();
  const g = grupos.find(x => x.cliente === clienteNome);
  if (!g) { toast('Nada pendente para exportar.'); return; }
  const linhas = apenasVencidas ? g.vencidas : [...g.vencidas, ...g.aVencer];
  if (linhas.length === 0) {
    toast('Nada para exportar nessa opção.');
    return;
  }
  linhas.sort((a, b) => (a.vencimento < b.vencimento ? -1 : a.vencimento > b.vencimento ? 1 : 0));

  const titulo = apenasVencidas ? 'Parcelas vencidas' : 'Parcelas pendentes (vencidas e a vencer)';
  const total = linhas.reduce((s, l) => s + l.valor, 0);

  const WIDTH = 800;
  const PAD = 32;
  const ROW_H = 74;
  const HEADER_H = 172;
  const FOOTER_H = 46;
  const height = HEADER_H + linhas.length * ROW_H + FOOTER_H;

  const scale = 2;
  const canvas = document.createElement('canvas');
  canvas.width = WIDTH * scale;
  canvas.height = height * scale;
  const ctx = canvas.getContext('2d');
  ctx.scale(scale, scale);

  ctx.fillStyle = '#FAFAF8';
  ctx.fillRect(0, 0, WIDTH, height);

  const grad = ctx.createLinearGradient(0, 0, WIDTH, 0);
  grad.addColorStop(0, '#0E7C7B');
  grad.addColorStop(1, '#0A4F4E');
  ctx.fillStyle = grad;
  ctx.fillRect(0, 0, WIDTH, HEADER_H);

  ctx.fillStyle = '#FFFFFF';
  ctx.font = '700 22px Roboto, Arial, sans-serif';
  ctx.fillText(escapeHtml(clienteNome), PAD, 40);

  ctx.font = '600 15px Roboto, Arial, sans-serif';
  ctx.globalAlpha = 0.9;
  ctx.fillText(titulo, PAD, 64);
  ctx.globalAlpha = 1;

  ctx.font = '400 13px Roboto, Arial, sans-serif';
  ctx.globalAlpha = 0.85;
  ctx.fillText(`Gerado em ${new Date().toLocaleDateString('pt-BR')} · Entregas Comuns`, PAD, 84);
  ctx.globalAlpha = 1;

  ctx.font = '400 13px Roboto, Arial, sans-serif';
  ctx.globalAlpha = 0.8;
  ctx.fillText('TOTAL', PAD, 120);
  ctx.globalAlpha = 1;
  ctx.font = '800 32px Roboto, Arial, sans-serif';
  ctx.fillText(fmtMoney(total), PAD, 152);

  let y = HEADER_H;
  linhas.forEach((l, i) => {
    if (i % 2 === 1) {
      ctx.fillStyle = '#F0F4F4';
      ctx.fillRect(0, y, WIDTH, ROW_H);
    }
    ctx.fillStyle = '#1F2933';
    ctx.font = '700 18px Roboto, Arial, sans-serif';
    ctx.fillText(`Entrega ${fmtDateBR(l.dataEntrega)} · ${l.quantidade} un.`, PAD, y + 26);

    ctx.fillStyle = '#5C6B73';
    ctx.font = '400 13.5px Roboto, Arial, sans-serif';
    ctx.fillText(`Vencimento: ${fmtDateBR(l.vencimento)}`, PAD, y + 46);

    const statusTxt = l.status === 'vencida' ? 'VENCIDA' : 'A VENCER';
    ctx.fillStyle = l.status === 'vencida' ? '#E5502F' : '#8A6210';
    ctx.font = '700 12px Roboto, Arial, sans-serif';
    ctx.fillText(statusTxt, PAD, y + 64);

    ctx.fillStyle = l.status === 'vencida' ? '#E5502F' : '#1F2933';
    ctx.font = '700 20px Roboto, Arial, sans-serif';
    const valorTxt = fmtMoney(l.valor);
    const w = ctx.measureText(valorTxt).width;
    ctx.fillText(valorTxt, WIDTH - PAD - w, y + 40);

    ctx.strokeStyle = '#E7ECEC';
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(PAD, y + ROW_H - 1);
    ctx.lineTo(WIDTH - PAD, y + ROW_H - 1);
    ctx.stroke();

    y += ROW_H;
  });

  ctx.fillStyle = '#5C6B73';
  ctx.font = '400 12.5px Roboto, Arial, sans-serif';
  ctx.fillText(`${linhas.length} parcela${linhas.length > 1 ? 's' : ''}`, PAD, y + 28);

  canvas.toBlob(async (blob) => {
    if (!blob) { toast('Não foi possível gerar a imagem.'); return; }
    const sufixo = apenasVencidas ? 'Vencidas' : 'Pendentes';
    const nomeArquivo = `Pagamentos_${sufixo}_${clienteNome.replace(/[^\w]+/g, '_')}_${todayStr()}.png`;
    const file = new File([blob], nomeArquivo, { type: 'image/png' });

    if (navigator.canShare && navigator.canShare({ files: [file] })) {
      try {
        await navigator.share({ files: [file], title: titulo });
        return;
      } catch (err) {
        if (err && err.name === 'AbortError') return;
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
  const lista = clientesDoTipo(TIPO_ATUAL).slice().sort((a, b) => a.nome.localeCompare(b.nome));
  if (lista.length === 0) {
    el.innerHTML = tipoSelectorHtml() + emptyState('🏬', 'Nenhuma loja cadastrada', 'Toque no botão + para cadastrar a primeira loja.');
    return;
  }
  el.innerHTML = tipoSelectorHtml() + lista.map(c => `
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
  if (!podeGerenciar()) { toast('Só administradores podem editar lojas.'); return; }
  const editando = !!id;
  const c = editando ? DB.clientes.find(x => x.id === id) : null;
  const tipo = editando ? c.tipo : TIPO_ATUAL;
  const html = `
    <div class="modal-header">
      <h2>${editando ? 'Editar loja' : 'Nova loja'}</h2>
      <button class="close-x" onclick="closeModal()">✕</button>
    </div>
    <div class="hint" style="margin-bottom:10px;">${TIPO_LABEL[tipo]}</div>
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
      <label>Valor padrão do picolé para essa loja</label>
      <input type="number" id="f-valor-padrao" step="0.01" value="${c && c.valorPadrao ? c.valorPadrao : ''}" placeholder="0,00">
      <div class="hint">Opcional. Preenche automaticamente o valor unitário ao lançar uma nova entrega.</div>
    </div>
    <div class="field">
      <label>Quantidade padrão entregue nessa loja</label>
      <input type="number" id="f-qtd-padrao" value="${c && c.quantidadePadrao ? c.quantidadePadrao : ''}" placeholder="0">
      <div class="hint">Opcional. Preenche automaticamente a quantidade ao lançar uma nova entrega.</div>
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
  const valorPadrao = Number(document.getElementById('f-valor-padrao').value) || 0;
  const quantidadePadrao = Number(document.getElementById('f-qtd-padrao').value) || 0;
  const ativo = document.getElementById('f-ativo').value === 'true';

  if (!nome) { toast('Digite o nome da loja.'); return; }

  const tipo = id ? DB.clientes.find(x => x.id === id).tipo : TIPO_ATUAL;
  const nomeDuplicado = DB.clientes.some(c => c.nome === nome && c.tipo === tipo && c.id !== id);
  if (nomeDuplicado) { toast('Já existe uma loja com esse nome nesse tipo.'); return; }

  if (id) {
    const antigo = DB.clientes.find(x => x.id === id);
    const nomeAntigo = antigo.nome;
    Object.assign(antigo, { nome, endereco, contato, intervaloPadrao, valorPadrao, quantidadePadrao, ativo });
    if (nomeAntigo !== nome) {
      DB.entregas.forEach(e => { if (e.cliente === nomeAntigo && e.tipo === tipo) e.cliente = nome; });
    }
  } else {
    DB.clientes.push({ id: uid('c'), tipo, nome, endereco, contato, intervaloPadrao, valorPadrao, quantidadePadrao, ativo });
  }
  saveDB();
  closeModal();
  refreshAll();
  toast('Loja salva.');
}

function deleteCliente(id) {
  const c = DB.clientes.find(x => x.id === id);
  const temEntregas = entregasAtivas(c.tipo).some(e => e.cliente === c.nome);
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
  const { meses, linhas, totais, totalGeral } = getResumoMensal(TIPO_ATUAL);
  if (meses.length === 0) {
    el.innerHTML = tipoSelectorHtml() + emptyState('📊', 'Sem dados ainda', 'Assim que houver entregas registradas, o resumo mensal aparece aqui.');
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
  el.innerHTML = tipoSelectorHtml() + `
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
  if (!podeGerenciar()) { toast('Só administradores podem editar feriados.'); return; }
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
   MOTORISTAS
   ========================================================= */
function motoristasAtivos() {
  return DB.motoristas.filter(m => m.ativo !== false);
}

function renderMotoristas() {
  const el = document.getElementById('motoristas-list');
  const lista = DB.motoristas.slice().sort((a, b) => a.nome.localeCompare(b.nome));
  if (lista.length === 0) {
    el.innerHTML = emptyState('🚐', 'Nenhum motorista cadastrado', 'Toque no botão + para adicionar.');
    return;
  }
  el.innerHTML = lista.map(m => `
    <div class="card" onclick="openMotoristaForm('${m.id}')" style="cursor:pointer;">
      <div class="card-row">
        <div>
          <div class="card-title">${escapeHtml(m.nome)}</div>
        </div>
        <span class="badge ${m.ativo === false ? 'grey' : 'green'}">${m.ativo === false ? 'Inativo' : 'Ativo'}</span>
      </div>
    </div>
  `).join('');
}

function openMotoristaForm(id) {
  if (!podeGerenciar()) { toast('Só administradores podem editar motoristas.'); return; }
  const editando = !!id;
  const m = editando ? DB.motoristas.find(x => x.id === id) : null;
  const html = `
    <div class="modal-header">
      <h2>${editando ? 'Editar motorista' : 'Novo motorista'}</h2>
      <button class="close-x" onclick="closeModal()">✕</button>
    </div>
    <div class="field">
      <label>Nome do motorista</label>
      <input type="text" id="f-nome-motorista" value="${m ? escapeHtml(m.nome) : ''}" placeholder="Ex: João">
    </div>
    <div class="field">
      <label>Status</label>
      <select id="f-ativo-motorista">
        <option value="true" ${!m || m.ativo !== false ? 'selected' : ''}>Ativo</option>
        <option value="false" ${m && m.ativo === false ? 'selected' : ''}>Inativo</option>
      </select>
    </div>
    <div class="formbtns">
      <button class="btn ghost block" onclick="closeModal()">Cancelar</button>
      <button class="btn primary block" onclick="saveMotorista(${editando ? `'${id}'` : 'null'})">Salvar</button>
    </div>
    ${editando ? `<button class="danger-link" onclick="deleteMotorista('${id}')">Excluir este motorista</button>` : ''}
  `;
  openModal(html);
}

function saveMotorista(id) {
  const nome = document.getElementById('f-nome-motorista').value.trim();
  const ativo = document.getElementById('f-ativo-motorista').value === 'true';
  if (!nome) { toast('Digite o nome do motorista.'); return; }
  const nomeDuplicado = DB.motoristas.some(m => m.nome === nome && m.id !== id);
  if (nomeDuplicado) { toast('Já existe um motorista com esse nome.'); return; }
  if (id) {
    Object.assign(DB.motoristas.find(x => x.id === id), { nome, ativo });
  } else {
    DB.motoristas.push({ id: uid('m'), nome, ativo });
  }
  saveDB();
  closeModal();
  refreshAll();
  toast('Motorista salvo.');
}

function deleteMotorista(id) {
  if (!confirm('Excluir este motorista?')) return;
  DB.motoristas = DB.motoristas.filter(x => x.id !== id);
  saveDB();
  closeModal();
  refreshAll();
  toast('Motorista excluído.');
}

/* =========================================================
   FINANCEIRO
   ========================================================= */
const CATEGORIA_VENDA = {
  personalizada: 'Vendas de Entregas Personalizadas',
  comum: 'Vendas de Entregas Comuns',
  direta: 'Vendas de Vendas Diretas',
};

function garantirCategoriasVenda() {
  TIPOS_CLIENTE.forEach(t => {
    const nome = CATEGORIA_VENDA[t];
    if (!DB.categoriasFinanceiras.some(c => c.tipo === 'entrada' && c.nome === nome)) {
      DB.categoriasFinanceiras.push({ id: uid('cf'), tipo: 'entrada', nome, ativo: true, sistema: true });
    }
  });
}

function categoriasFinanceirasDoTipo(tipoMov) {
  garantirCategoriasVenda();
  return DB.categoriasFinanceiras.filter(c => c.tipo === (tipoMov === 'entrada' ? 'entrada' : 'despesa') && c.ativo !== false)
    .slice().sort((a, b) => a.nome.localeCompare(b.nome));
}

function lancamentosAtivos() {
  return DB.lancamentos.filter(l => !l.deletado);
}

/* ---- sincronização automática: pagamentos de entregas -> lançamentos ---- */
function sincronizarLancamentoAutoEntrega(entrega) {
  garantirCategoriasVenda();
  const nome = 'Venda ' + entrega.cliente;
  const existente = DB.lancamentos.find(l => !l.deletado && l.origem && l.origem.entregaId === entrega.id && !l.origem.parcelaId);
  if (entrega.dataPagamento) {
    const valor = valorTotalEntrega(entrega);
    if (existente) {
      Object.assign(existente, {
        valor, data: entrega.dataPagamento, categoria: CATEGORIA_VENDA[entrega.tipo], nome,
        observacoes: entrega.obsPagamento || '', updatedAt: Date.now(),
      });
    } else {
      DB.lancamentos.push({
        id: uid('lc'), tipoMov: 'entrada', categoria: CATEGORIA_VENDA[entrega.tipo], nome, valor,
        data: entrega.dataPagamento, observacoes: entrega.obsPagamento || '', automatico: true,
        origem: { entregaId: entrega.id, tipoCliente: entrega.tipo },
        criadoPor: AUTH_USER ? AUTH_USER.email : null, updatedAt: Date.now(), deletado: false,
      });
    }
  } else if (existente) {
    existente.deletado = true;
    existente.updatedAt = Date.now();
  }
}

function sincronizarParcelasComumFinanceiro(entrega) {
  garantirCategoriasVenda();
  const nome = 'Venda ' + entrega.cliente;
  // remove lançamentos de parcelas que não existem mais nessa entrega (ex: após "recalcular parcelas")
  const idsAtuais = new Set((entrega.parcelas || []).map(p => p.id));
  DB.lancamentos.forEach(l => {
    if (!l.deletado && l.origem && l.origem.entregaId === entrega.id && l.origem.parcelaId && !idsAtuais.has(l.origem.parcelaId)) {
      l.deletado = true;
      l.updatedAt = Date.now();
    }
  });
  (entrega.parcelas || []).forEach(p => {
    const existente = DB.lancamentos.find(l => !l.deletado && l.origem && l.origem.entregaId === entrega.id && l.origem.parcelaId === p.id);
    if (p.dataPagamento) {
      const valor = Number(p.valor) || 0;
      if (existente) {
        Object.assign(existente, { valor, data: p.dataPagamento, categoria: CATEGORIA_VENDA[entrega.tipo], nome, updatedAt: Date.now() });
      } else {
        DB.lancamentos.push({
          id: uid('lc'), tipoMov: 'entrada', categoria: CATEGORIA_VENDA[entrega.tipo], nome, valor,
          data: p.dataPagamento, observacoes: '', automatico: true,
          origem: { entregaId: entrega.id, parcelaId: p.id, tipoCliente: entrega.tipo },
          criadoPor: AUTH_USER ? AUTH_USER.email : null, updatedAt: Date.now(), deletado: false,
        });
      }
    } else if (existente) {
      existente.deletado = true;
      existente.updatedAt = Date.now();
    }
  });
}

function removerLancamentosDaEntrega(entregaId) {
  DB.lancamentos.forEach(l => {
    if (!l.deletado && l.origem && l.origem.entregaId === entregaId) {
      l.deletado = true;
      l.updatedAt = Date.now();
    }
  });
}

/* ---- Categorias financeiras (CRUD) ---- */
function renderCategoriasFinanceiras() {
  garantirCategoriasVenda();
  const el = document.getElementById('fin-categorias-list');
  const entradas = DB.categoriasFinanceiras.filter(c => c.tipo === 'entrada').sort((a, b) => a.nome.localeCompare(b.nome));
  const despesas = DB.categoriasFinanceiras.filter(c => c.tipo === 'despesa').sort((a, b) => a.nome.localeCompare(b.nome));

  const linha = (c) => `
    <div class="card" ${c.sistema ? '' : `onclick="openCategoriaFinanceiraForm('${c.id}')" style="cursor:pointer;"`}>
      <div class="card-row">
        <div>
          <div class="card-title">${escapeHtml(c.nome)}</div>
          ${c.sistema ? '<div class="card-sub">Categoria automática (ligada aos pagamentos)</div>' : ''}
        </div>
        <span class="badge ${c.ativo === false ? 'grey' : 'green'}">${c.ativo === false ? 'Inativa' : 'Ativa'}</span>
      </div>
    </div>`;

  el.innerHTML = `
    <div class="section-title">Categorias de entrada</div>
    ${entradas.map(linha).join('') || emptyState('💵', 'Nenhuma categoria de entrada', '')}
    <div class="section-title" style="margin-top:16px;">Categorias de despesa</div>
    ${despesas.map(linha).join('') || emptyState('🧾', 'Nenhuma categoria de despesa ainda', 'Toque no botão + para criar (ex: Gasolina, Pedágio).')}
  `;
}

function openCategoriaFinanceiraForm(id) {
  if (!podeGerenciar()) { toast('Só administradores e gerentes podem editar categorias.'); return; }
  const editando = !!id;
  const c = editando ? DB.categoriasFinanceiras.find(x => x.id === id) : null;
  if (editando && c.sistema) { toast('Essa categoria é automática e não pode ser editada.'); return; }
  const html = `
    <div class="modal-header">
      <h2>${editando ? 'Editar categoria' : 'Nova categoria'}</h2>
      <button class="close-x" onclick="closeModal()">✕</button>
    </div>
    <div class="field">
      <label>Tipo</label>
      <select id="f-cat-tipo" ${editando ? 'disabled' : ''}>
        <option value="entrada" ${!c || c.tipo === 'entrada' ? 'selected' : ''}>Entrada</option>
        <option value="despesa" ${c && c.tipo === 'despesa' ? 'selected' : ''}>Despesa</option>
      </select>
    </div>
    <div class="field">
      <label>Nome da categoria</label>
      <input type="text" id="f-cat-nome" value="${c ? escapeHtml(c.nome) : ''}" placeholder="Ex: Gasolina, Pedágio, Salário de fulano">
    </div>
    <div class="field">
      <label>Status</label>
      <select id="f-cat-ativo">
        <option value="true" ${!c || c.ativo !== false ? 'selected' : ''}>Ativa</option>
        <option value="false" ${c && c.ativo === false ? 'selected' : ''}>Inativa</option>
      </select>
    </div>
    <div class="formbtns">
      <button class="btn ghost block" onclick="closeModal()">Cancelar</button>
      <button class="btn primary block" onclick="saveCategoriaFinanceira(${editando ? `'${id}'` : 'null'})">Salvar</button>
    </div>
    ${editando ? `<button class="danger-link" onclick="deleteCategoriaFinanceira('${id}')">Excluir esta categoria</button>` : ''}
  `;
  openModal(html);
}

function saveCategoriaFinanceira(id) {
  const tipo = document.getElementById('f-cat-tipo').value;
  const nome = document.getElementById('f-cat-nome').value.trim();
  const ativo = document.getElementById('f-cat-ativo').value === 'true';
  if (!nome) { toast('Digite o nome da categoria.'); return; }
  const duplicada = DB.categoriasFinanceiras.some(c => c.nome === nome && c.tipo === tipo && c.id !== id);
  if (duplicada) { toast('Já existe uma categoria com esse nome nesse tipo.'); return; }
  if (id) {
    Object.assign(DB.categoriasFinanceiras.find(x => x.id === id), { nome, ativo });
  } else {
    DB.categoriasFinanceiras.push({ id: uid('cf'), tipo, nome, ativo });
  }
  saveDB();
  closeModal();
  refreshAll();
  toast('Categoria salva.');
}

function deleteCategoriaFinanceira(id) {
  const c = DB.categoriasFinanceiras.find(x => x.id === id);
  if (c && c.sistema) { toast('Essa categoria é automática e não pode ser excluída.'); return; }
  if (!confirm('Excluir esta categoria? Lançamentos já feitos com ela continuam existindo.')) return;
  DB.categoriasFinanceiras = DB.categoriasFinanceiras.filter(x => x.id !== id);
  saveDB();
  closeModal();
  refreshAll();
  toast('Categoria excluída.');
}

/* ---- Lançamentos manuais (CRUD) ---- */
function openLancamentoForm(id) {
  if (!podeGerenciar()) { toast('Só administradores e gerentes podem lançar entradas e saídas.'); return; }
  const editando = !!id;
  const l = editando ? DB.lancamentos.find(x => x.id === id) : null;
  if (editando && l.automatico) { toast('Esse lançamento é automático (gerado por um pagamento) e não pode ser editado aqui.'); return; }
  const tipoMov = l ? l.tipoMov : 'entrada';

  const html = `
    <div class="modal-header">
      <h2>${editando ? 'Editar lançamento' : 'Novo lançamento'}</h2>
      <button class="close-x" onclick="closeModal()">✕</button>
    </div>
    <div class="field">
      <label>Tipo</label>
      <select id="f-lanc-tipo" onchange="atualizarCategoriasLancamentoForm()">
        <option value="entrada" ${tipoMov === 'entrada' ? 'selected' : ''}>Entrada</option>
        <option value="saida" ${tipoMov === 'saida' ? 'selected' : ''}>Saída (gasto)</option>
      </select>
    </div>
    <div class="field">
      <label>Categoria</label>
      <select id="f-lanc-categoria" onchange="onCategoriaLancamentoChange()"></select>
    </div>
    <div class="field">
      <label>Nome</label>
      <input type="text" id="f-lanc-nome" list="dl-lanc-nomes" value="${l ? escapeHtml(l.nome || '') : ''}" placeholder="Ex: Posto Ipiranga, Salário João">
      <datalist id="dl-lanc-nomes"></datalist>
      <div class="hint">Digite um nome novo ou escolha um já usado nessa categoria.</div>
    </div>
    <div class="row2">
      <div class="field">
        <label>Data</label>
        <input type="date" id="f-lanc-data" value="${l ? l.data : todayStr()}">
      </div>
      <div class="field">
        <label>Valor</label>
        <input type="number" id="f-lanc-valor" step="0.01" value="${l ? l.valor : ''}" placeholder="0,00">
      </div>
    </div>
    <div class="field">
      <label>Observações</label>
      <textarea id="f-lanc-obs" placeholder="Opcional">${l ? escapeHtml(l.observacoes || '') : ''}</textarea>
    </div>
    <div class="formbtns">
      <button class="btn ghost block" onclick="closeModal()">Cancelar</button>
      <button class="btn primary block" onclick="saveLancamento(${editando ? `'${id}'` : 'null'})">Salvar</button>
    </div>
    ${editando ? `<button class="danger-link" onclick="deleteLancamento('${id}')">Excluir este lançamento</button>` : ''}
  `;
  openModal(html);
  atualizarCategoriasLancamentoForm(l ? l.categoria : null);
}

const CATEGORIA_NOVA_SENTINEL = '__nova_categoria__';

function atualizarCategoriasLancamentoForm(categoriaSelecionada) {
  const tipoMov = document.getElementById('f-lanc-tipo').value;
  const el = document.getElementById('f-lanc-categoria');
  const categorias = categoriasFinanceirasDoTipo(tipoMov);
  const opcoes = categorias.map(c =>
    `<option value="${escapeHtml(c.nome)}" ${categoriaSelecionada === c.nome ? 'selected' : ''}>${escapeHtml(c.nome)}</option>`
  ).join('');
  const semCategorias = categorias.length === 0
    ? `<option value="" disabled selected>Nenhuma categoria ainda</option>` : '';
  el.innerHTML = semCategorias + opcoes + `<option value="${CATEGORIA_NOVA_SENTINEL}">➕ Criar nova categoria...</option>`;
  atualizarSugestoesNomeLancamento();
}

function onCategoriaLancamentoChange() {
  const el = document.getElementById('f-lanc-categoria');
  if (el.value === CATEGORIA_NOVA_SENTINEL) {
    criarCategoriaRapida();
    return;
  }
  atualizarSugestoesNomeLancamento();
}

function criarCategoriaRapida() {
  const tipoMov = document.getElementById('f-lanc-tipo').value;
  const tipoCategoria = tipoMov === 'entrada' ? 'entrada' : 'despesa';
  const nome = (prompt(`Nome da nova categoria de ${tipoCategoria === 'entrada' ? 'entrada' : 'despesa'}:`) || '').trim();
  if (!nome) { atualizarCategoriasLancamentoForm(); return; }
  const duplicada = DB.categoriasFinanceiras.some(c => c.nome === nome && c.tipo === tipoCategoria);
  if (!duplicada) {
    DB.categoriasFinanceiras.push({ id: uid('cf'), tipo: tipoCategoria, nome, ativo: true });
    saveDB();
  }
  atualizarCategoriasLancamentoForm(nome);
}

function atualizarSugestoesNomeLancamento() {
  const categoria = document.getElementById('f-lanc-categoria').value;
  const dl = document.getElementById('dl-lanc-nomes');
  if (!dl) return;
  const nomes = Array.from(new Set(
    lancamentosAtivos().filter(l => l.categoria === categoria && l.nome).map(l => l.nome)
  )).sort((a, b) => a.localeCompare(b));
  dl.innerHTML = nomes.map(n => `<option value="${escapeHtml(n)}"></option>`).join('');
}

function saveLancamento(id) {
  const tipoMov = document.getElementById('f-lanc-tipo').value;
  const categoria = document.getElementById('f-lanc-categoria').value;
  const nome = document.getElementById('f-lanc-nome').value.trim();
  const data = document.getElementById('f-lanc-data').value;
  const valor = Number(document.getElementById('f-lanc-valor').value) || 0;
  const observacoes = document.getElementById('f-lanc-obs').value.trim();

  if (!data || !categoria || categoria === CATEGORIA_NOVA_SENTINEL || valor <= 0) {
    toast('Preencha categoria, data e um valor maior que zero.');
    return;
  }

  if (id) {
    Object.assign(DB.lancamentos.find(x => x.id === id), { tipoMov, categoria, nome, data, valor, observacoes, updatedAt: Date.now() });
  } else {
    DB.lancamentos.push({
      id: uid('lc'), tipoMov, categoria, nome, data, valor, observacoes, automatico: false, origem: null,
      criadoPor: AUTH_USER ? AUTH_USER.email : null, updatedAt: Date.now(), deletado: false,
    });
  }
  saveDB();
  closeModal();
  refreshAll();
  toast('Lançamento salvo.');
}

function deleteLancamento(id) {
  const l = DB.lancamentos.find(x => x.id === id);
  if (l && l.automatico) { toast('Esse lançamento é automático e não pode ser excluído aqui. Desmarque o pagamento na entrega de origem.'); return; }
  if (!confirm('Excluir este lançamento?')) return;
  l.deletado = true;
  l.updatedAt = Date.now();
  saveDB();
  closeModal();
  refreshAll();
  toast('Lançamento excluído.');
}

/* ---- Tela: Lançamentos (lista recente + adicionar) ---- */
function renderFinLancamentos() {
  const el = document.getElementById('fin-lancamentos-list');
  const lista = lancamentosAtivos().slice().sort((a, b) => (a.data < b.data ? 1 : a.data > b.data ? -1 : (b.updatedAt || 0) - (a.updatedAt || 0))).slice(0, 60);
  const totalEntradas = lancamentosAtivos().filter(l => l.tipoMov === 'entrada').reduce((s, l) => s + l.valor, 0);
  const totalSaidas = lancamentosAtivos().filter(l => l.tipoMov === 'saida').reduce((s, l) => s + l.valor, 0);

  const kpis = `
    <div class="row2" style="margin-bottom:14px;">
      <div class="kpi">
        <div class="label">Total de entradas</div>
        <div class="value" style="color:var(--teal-700);">${fmtMoney(totalEntradas)}</div>
      </div>
      <div class="kpi">
        <div class="label">Total de saídas</div>
        <div class="value" style="color:var(--red-ink);">${fmtMoney(totalSaidas)}</div>
      </div>
    </div>`;

  if (lista.length === 0) {
    el.innerHTML = kpis + emptyState('💵', 'Nenhum lançamento ainda', 'Toque no botão + para lançar uma entrada ou saída manual.');
    return;
  }

  el.innerHTML = kpis + `<div class="hint" style="margin-bottom:8px;">Mostrando os 60 lançamentos mais recentes. Para ver por período, use o Relatório.</div>` + lista.map(l => `
    <div class="card" ${l.automatico ? '' : `onclick="openLancamentoForm('${l.id}')" style="cursor:pointer;"`}>
      <div class="card-row">
        <div>
          <div class="card-title">${l.nome ? escapeHtml(l.nome) : escapeHtml(l.categoria)}</div>
          <div class="card-sub">${l.nome ? escapeHtml(l.categoria) + ' · ' : ''}${fmtDateBR(l.data)}${l.automatico ? ' · automático' : ''}</div>
        </div>
        <div class="card-title" style="color:${l.tipoMov === 'entrada' ? 'var(--teal-700)' : 'var(--red-ink)'};">
          ${l.tipoMov === 'entrada' ? '+' : '−'} ${fmtMoney(l.valor)}
        </div>
      </div>
    </div>
  `).join('');
}

/* ---- Tela: Relatório por período ---- */
function relatorioFinanceiroPeriodo() {
  const ini = document.getElementById('fin-rel-ini') ? document.getElementById('fin-rel-ini').value : '';
  const fim = document.getElementById('fin-rel-fim') ? document.getElementById('fin-rel-fim').value : '';
  return { ini: ini || '0000-01-01', fim: fim || '9999-12-31' };
}

function renderFinRelatorio() {
  const wrap = document.getElementById('fin-relatorio-container');
  const hoje = todayStr();
  const inicioMes = hoje.slice(0, 8) + '01';
  wrap.innerHTML = `
    <div class="row2" style="margin-bottom:10px;">
      <div class="field">
        <label>De</label>
        <input type="date" id="fin-rel-ini" value="${inicioMes}" onchange="renderFinRelatorioResultado()">
      </div>
      <div class="field">
        <label>Até</label>
        <input type="date" id="fin-rel-fim" value="${hoje}" onchange="renderFinRelatorioResultado()">
      </div>
    </div>
    <button type="button" class="btn ghost block" id="fin-rel-btn-filtro" style="margin-bottom:12px;" onclick="abrirFiltroCategoriasRelatorio()">🔎 Filtrar categorias</button>
    <div class="pill-select" style="margin-bottom:12px;">
      <button class="active" id="fin-rel-modo-resumido" onclick="setFinRelatorioModo('resumido')">Resumido</button>
      <button id="fin-rel-modo-detalhado" onclick="setFinRelatorioModo('detalhado')">Detalhado</button>
    </div>
    <div id="fin-relatorio-resultado"></div>
  `;
  FIN_RELATORIO_MODO = 'resumido';
  FIN_RELATORIO_CATEGORIAS = null;
  atualizarBotaoFiltroCategorias();
  renderFinRelatorioResultado();
}

let FIN_RELATORIO_CATEGORIAS = null; // null = todas as categorias (sem filtro)

function atualizarBotaoFiltroCategorias() {
  const btn = document.getElementById('fin-rel-btn-filtro');
  if (!btn) return;
  btn.textContent = FIN_RELATORIO_CATEGORIAS
    ? `🔎 Filtro: ${FIN_RELATORIO_CATEGORIAS.size} categoria${FIN_RELATORIO_CATEGORIAS.size > 1 ? 's' : ''} — toque para ajustar`
    : '🔎 Filtrar categorias';
}

function abrirFiltroCategoriasRelatorio() {
  garantirCategoriasVenda();
  const entradas = DB.categoriasFinanceiras.filter(c => c.tipo === 'entrada').sort((a, b) => a.nome.localeCompare(b.nome));
  const despesas = DB.categoriasFinanceiras.filter(c => c.tipo === 'despesa').sort((a, b) => a.nome.localeCompare(b.nome));
  const marcada = (nome) => !FIN_RELATORIO_CATEGORIAS || FIN_RELATORIO_CATEGORIAS.has(nome);

  const linha = (c) => `
    <label class="card" style="display:flex;align-items:center;gap:10px;cursor:pointer;padding:10px 14px;margin-bottom:8px;">
      <input type="checkbox" class="chk-filtro-categoria" value="${escapeHtml(c.nome)}" ${marcada(c.nome) ? 'checked' : ''} style="width:18px;height:18px;">
      <span>${escapeHtml(c.nome)}</span>
    </label>`;

  const html = `
    <div class="modal-header">
      <h2>Filtrar categorias</h2>
      <button class="close-x" onclick="closeModal()">✕</button>
    </div>
    <div class="row2" style="margin-bottom:12px;">
      <button type="button" class="btn ghost block" onclick="marcarTodasCategoriasFiltro(true)">Marcar todas</button>
      <button type="button" class="btn ghost block" onclick="marcarTodasCategoriasFiltro(false)">Desmarcar todas</button>
    </div>
    <div class="section-title">Entradas</div>
    ${entradas.map(linha).join('') || emptyState('', 'Nenhuma categoria de entrada', '')}
    <div class="section-title" style="margin-top:14px;">Despesas</div>
    ${despesas.map(linha).join('') || emptyState('', 'Nenhuma categoria de despesa', '')}
    <div class="formbtns" style="margin-top:14px;">
      <button class="btn ghost block" onclick="closeModal()">Cancelar</button>
      <button class="btn primary block" onclick="aplicarFiltroCategoriasRelatorio()">Aplicar</button>
    </div>
  `;
  openModal(html);
}

function marcarTodasCategoriasFiltro(valor) {
  document.querySelectorAll('.chk-filtro-categoria').forEach(chk => { chk.checked = valor; });
}

function aplicarFiltroCategoriasRelatorio() {
  const todas = Array.from(document.querySelectorAll('.chk-filtro-categoria'));
  const marcadas = todas.filter(chk => chk.checked).map(chk => chk.value);
  FIN_RELATORIO_CATEGORIAS = (marcadas.length === todas.length) ? null : new Set(marcadas);
  closeModal();
  atualizarBotaoFiltroCategorias();
  renderFinRelatorioResultado();
}

let FIN_RELATORIO_MODO = 'resumido';
function setFinRelatorioModo(modo) {
  FIN_RELATORIO_MODO = modo;
  document.getElementById('fin-rel-modo-resumido').classList.toggle('active', modo === 'resumido');
  document.getElementById('fin-rel-modo-detalhado').classList.toggle('active', modo === 'detalhado');
  renderFinRelatorioResultado();
}

function renderFinRelatorioResultado() {
  const el = document.getElementById('fin-relatorio-resultado');
  if (!el) return;
  const { ini, fim } = relatorioFinanceiroPeriodo();
  const doPeriodo = lancamentosAtivos().filter(l =>
    l.data >= ini && l.data <= fim && (!FIN_RELATORIO_CATEGORIAS || FIN_RELATORIO_CATEGORIAS.has(l.categoria))
  );
  const entradas = doPeriodo.filter(l => l.tipoMov === 'entrada');
  const saidas = doPeriodo.filter(l => l.tipoMov === 'saida');
  const totalEntradas = entradas.reduce((s, l) => s + l.valor, 0);
  const totalSaidas = saidas.reduce((s, l) => s + l.valor, 0);
  const resultado = totalEntradas - totalSaidas;

  const kpis = `
    <div class="row2" style="margin-bottom:14px;">
      <div class="kpi">
        <div class="label">Entradas</div>
        <div class="value" style="color:var(--teal-700);">${fmtMoney(totalEntradas)}</div>
      </div>
      <div class="kpi">
        <div class="label">Saídas</div>
        <div class="value" style="color:var(--red-ink);">${fmtMoney(totalSaidas)}</div>
      </div>
    </div>
    <div class="kpi" style="margin-bottom:14px;background:linear-gradient(135deg,${resultado >= 0 ? '#0E7C7B,#0A4F4E' : '#E5502F,#9A3324'});">
      <div class="label">Resultado do período</div>
      <div class="value">${fmtMoney(resultado)}</div>
    </div>`;

  if (doPeriodo.length === 0) {
    el.innerHTML = kpis + emptyState('📑', 'Nada nesse período', 'Ajuste as datas acima para ver outro intervalo.');
    return;
  }

  if (FIN_RELATORIO_MODO === 'detalhado') {
    const linhas = doPeriodo.slice().sort((a, b) => (a.data < b.data ? 1 : a.data > b.data ? -1 : 0)).map(l => `
      <div class="card">
        <div class="card-row">
          <div>
            <div class="card-title">${l.nome ? escapeHtml(l.nome) : escapeHtml(l.categoria)}</div>
            <div class="card-sub">${l.nome ? escapeHtml(l.categoria) + ' · ' : ''}${fmtDateBR(l.data)}${l.automatico ? ' · automático' : ''}${l.observacoes ? ' · ' + escapeHtml(l.observacoes) : ''}</div>
          </div>
          <div class="card-title" style="color:${l.tipoMov === 'entrada' ? 'var(--teal-700)' : 'var(--red-ink)'};">
            ${l.tipoMov === 'entrada' ? '+' : '−'} ${fmtMoney(l.valor)}
          </div>
        </div>
      </div>`).join('');
    el.innerHTML = kpis + linhas;
  } else {
    const agrupar = (lista) => {
      const mapa = new Map();
      lista.forEach(l => mapa.set(l.categoria, (mapa.get(l.categoria) || 0) + l.valor));
      return Array.from(mapa.entries()).sort((a, b) => b[1] - a[1]);
    };
    const gruposEntrada = agrupar(entradas);
    const gruposSaida = agrupar(saidas);
    const linhaGrupo = (nome, valor, cor) => `
      <div class="card" onclick="abrirDetalheCategoriaPeriodo('${escapeHtml(nome)}', '${ini}', '${fim}')" style="cursor:pointer;">
        <div class="card-row">
          <div class="card-title">${escapeHtml(nome)}</div>
          <div class="card-title" style="color:${cor};">${fmtMoney(valor)}</div>
        </div>
      </div>`;
    el.innerHTML = kpis +
      `<div class="section-title">Entradas por categoria</div>` +
      (gruposEntrada.map(([nome, valor]) => linhaGrupo(nome, valor, 'var(--teal-700)')).join('') || emptyState('', 'Sem entradas nesse período', '')) +
      `<div class="section-title" style="margin-top:14px;">Saídas por categoria</div>` +
      (gruposSaida.map(([nome, valor]) => linhaGrupo(nome, valor, 'var(--red-ink)')).join('') || emptyState('', 'Sem saídas nesse período', ''));
  }
}

/* ---- Detalhe de uma categoria (lista de lançamentos, clicável para editar) ---- */
function abrirDetalheCategoriaLancamentos(categoria, lista) {
  const ordenada = lista.slice().sort((a, b) => (a.data < b.data ? 1 : a.data > b.data ? -1 : 0));
  const total = ordenada.reduce((s, l) => s + l.valor, 0);

  const linha = (l) => {
    let onclick = '';
    if (l.automatico) {
      if (l.origem && l.origem.entregaId) onclick = `onclick="closeModal(); openEntregaForm('${l.origem.entregaId}')" style="cursor:pointer;"`;
    } else {
      onclick = `onclick="closeModal(); openLancamentoForm('${l.id}')" style="cursor:pointer;"`;
    }
    return `
      <div class="card" ${onclick}>
        <div class="card-row">
          <div>
            <div class="card-title">${l.nome ? escapeHtml(l.nome) : escapeHtml(l.categoria)}</div>
            <div class="card-sub">${fmtDateBR(l.data)}${l.automatico ? ' · automático (edite pela entrega)' : ''}</div>
          </div>
          <div class="card-title" style="color:${l.tipoMov === 'entrada' ? 'var(--teal-700)' : 'var(--red-ink)'};">${fmtMoney(l.valor)}</div>
        </div>
      </div>`;
  };

  const html = `
    <div class="modal-header">
      <h2>${escapeHtml(categoria)}</h2>
      <button class="close-x" onclick="closeModal()">✕</button>
    </div>
    <div class="hint" style="margin-bottom:12px;">${ordenada.length} lançamento${ordenada.length !== 1 ? 's' : ''} · ${fmtMoney(total)}</div>
    ${ordenada.map(linha).join('') || emptyState('', 'Nada aqui', '')}
  `;
  openModal(html);
}

function abrirDetalheCategoriaPeriodo(categoria, ini, fim) {
  const lista = lancamentosAtivos().filter(l => l.categoria === categoria && l.data >= ini && l.data <= fim);
  abrirDetalheCategoriaLancamentos(categoria, lista);
}

function abrirDetalheCategoriaMes(categoria, mesKey) {
  const lista = lancamentosAtivos().filter(l => l.categoria === categoria && l.data.slice(0, 7) === mesKey);
  abrirDetalheCategoriaLancamentos(categoria, lista);
}

/* ---- Tela: Resumo mensal financeiro ---- */
function mesesComLancamentos() {
  const set = new Set(lancamentosAtivos().map(l => l.data.slice(0, 7)));
  return Array.from(set).sort().reverse();
}

let FIN_RESUMO_MES = null;
function renderFinResumoMensal() {
  const el = document.getElementById('fin-resumo-container');
  const meses = mesesComLancamentos();
  if (meses.length === 0) {
    el.innerHTML = emptyState('📆', 'Sem dados ainda', 'Assim que houver lançamentos, o resumo por mês aparece aqui.');
    return;
  }
  if (!FIN_RESUMO_MES || !meses.includes(FIN_RESUMO_MES)) FIN_RESUMO_MES = meses[0];

  const seletor = `<div class="field" style="margin-bottom:14px;">
    <label>Mês</label>
    <select id="fin-resumo-mes-select" onchange="mudarFinResumoMes(this.value)">
      ${meses.map(m => `<option value="${m}" ${m === FIN_RESUMO_MES ? 'selected' : ''}>${fmtMesAno(m)}</option>`).join('')}
    </select>
  </div>`;

  const doMes = lancamentosAtivos().filter(l => l.data.slice(0, 7) === FIN_RESUMO_MES);
  const entradas = doMes.filter(l => l.tipoMov === 'entrada');
  const saidas = doMes.filter(l => l.tipoMov === 'saida');
  const totalEntradas = entradas.reduce((s, l) => s + l.valor, 0);
  const totalSaidas = saidas.reduce((s, l) => s + l.valor, 0);
  const resultado = totalEntradas - totalSaidas;

  const agrupar = (lista) => {
    const mapa = new Map();
    lista.forEach(l => mapa.set(l.categoria, (mapa.get(l.categoria) || 0) + l.valor));
    return Array.from(mapa.entries()).sort((a, b) => b[1] - a[1]);
  };
  const linhaGrupo = (nome, valor, cor) => `
    <div class="card" onclick="abrirDetalheCategoriaMes('${escapeHtml(nome)}', '${FIN_RESUMO_MES}')" style="cursor:pointer;">
      <div class="card-row">
        <div class="card-title">${escapeHtml(nome)}</div>
        <div class="card-title" style="color:${cor};">${fmtMoney(valor)}</div>
      </div>
    </div>`;

  el.innerHTML = seletor + `
    <div class="row2" style="margin-bottom:14px;">
      <div class="kpi">
        <div class="label">Entradas</div>
        <div class="value" style="color:var(--teal-700);">${fmtMoney(totalEntradas)}</div>
      </div>
      <div class="kpi">
        <div class="label">Saídas</div>
        <div class="value" style="color:var(--red-ink);">${fmtMoney(totalSaidas)}</div>
      </div>
    </div>
    <div class="kpi" style="margin-bottom:14px;background:linear-gradient(135deg,${resultado >= 0 ? '#0E7C7B,#0A4F4E' : '#E5502F,#9A3324'});">
      <div class="label">Resultado do mês</div>
      <div class="value">${fmtMoney(resultado)}</div>
    </div>
    <div class="section-title">Entradas por categoria</div>
    ${agrupar(entradas).map(([nome, valor]) => linhaGrupo(nome, valor, 'var(--teal-700)')).join('') || emptyState('', 'Sem entradas nesse mês', '')}
    <div class="section-title" style="margin-top:14px;">Saídas por categoria</div>
    ${agrupar(saidas).map(([nome, valor]) => linhaGrupo(nome, valor, 'var(--red-ink)')).join('') || emptyState('', 'Sem saídas nesse mês', '')}
  `;
}

function mudarFinResumoMes(mes) {
  FIN_RESUMO_MES = mes;
  renderFinResumoMensal();
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
  if (AUTH_USER) {
    toast('Desconecte a sincronização (Sair da conta) antes de apagar os dados deste aparelho.');
    return;
  }
  if (!confirm('Isso apaga TODOS os dados salvos neste app (lojas, entregas, feriados). Não pode ser desfeito. Continuar?')) return;
  localStorage.removeItem(STORAGE_KEY);
  loadDB();
  refreshAll();
  toast('Dados apagados.');
}

/* ---------------- localizar/remover entregas duplicadas ---------------- */
function encontrarEntregasDuplicadas() {
  const grupos = new Map();
  entregasAtivas().forEach(e => {
    const chave = [e.cliente, e.data, e.quantidade, e.valorUnitario || 0, e.dataPagamento || ''].join('|');
    if (!grupos.has(chave)) grupos.set(chave, []);
    grupos.get(chave).push(e);
  });
  const duplicadas = [];
  grupos.forEach(lista => {
    if (lista.length > 1) {
      lista.sort((a, b) => (a.updatedAt || 0) - (b.updatedAt || 0));
      duplicadas.push(...lista.slice(1));
    }
  });
  return duplicadas;
}

function verificarDuplicadas() {
  const dups = encontrarEntregasDuplicadas();
  if (dups.length === 0) {
    toast('Nenhuma entrega duplicada encontrada.');
    return;
  }
  const resumo = dups.slice(0, 5).map(e => `${e.cliente} · ${fmtDateBR(e.data)} · ${e.quantidade} un.`).join('\n');
  const msg = `Encontradas ${dups.length} entrega(s) que parecem duplicadas (mesma loja, data, quantidade e valor)` +
    (dups.length > 5 ? `, incluindo:\n${resumo}\n...` : `:\n${resumo}`) +
    `\n\nRemover as cópias extras (mantendo uma de cada)?`;
  if (!confirm(msg)) return;
  dups.forEach(e => { e.deletado = true; e.updatedAt = Date.now(); });
  saveDB();
  refreshAll();
  toast(`${dups.length} entrega(s) duplicada(s) removida(s).`);
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
  if (!podeImportar()) { toast('Só administradores podem importar planilha.'); return; }
  if (!confirm('Importar substitui TODOS os dados atuais deste app pelos dados da planilha. ' +
    'Antes de continuar, será feito um backup automático (local' +
    (AUTH_USER ? ' + nuvem' : '') + ') e uma planilha de segurança será baixada. Continuar?')) return;
  fazerBackupDeSeguranca().then(() => {
    document.getElementById('file-import').click();
  });
}

async function fazerBackupDeSeguranca() {
  const timestamp = Date.now();
  const snapshot = { clientes: DB.clientes, feriados: DB.feriados, entregas: DB.entregas, motoristas: DB.motoristas, categoriasFinanceiras: DB.categoriasFinanceiras, lancamentos: DB.lancamentos, criadoEm: timestamp };
  try {
    localStorage.setItem('picole_backup_pre_import', JSON.stringify(snapshot));
  } catch (e) {
    console.error('backup local falhou', e);
  }
  if (AUTH_USER && MEU_PAPEL) {
    try {
      await ensureFirebase();
      await firebase.firestore()
        .collection('sincronizacoes').doc(WORKSPACE_ID)
        .collection('backups').doc(String(timestamp))
        .set({ json: JSON.stringify(snapshot), criadoEm: timestamp, criadoPor: AUTH_USER.email });
    } catch (err) {
      console.error('backup na nuvem falhou', err);
    }
  }
  try {
    await exportarPlanilha();
  } catch (err) {
    console.error('backup em planilha falhou', err);
  }
  toast('Backup de segurança concluído. Escolha o arquivo para importar.');
}

function restaurarBackupLocal() {
  if (!podeImportar()) { toast('Só administradores podem restaurar backups.'); return; }
  const raw = localStorage.getItem('picole_backup_pre_import');
  if (!raw) { toast('Nenhum backup local encontrado neste aparelho.'); return; }
  let snapshot;
  try { snapshot = JSON.parse(raw); } catch (e) { toast('Backup corrompido.'); return; }

  const quando = new Date(snapshot.criadoEm).toLocaleString('pt-BR');
  const idsNoBackup = new Set((snapshot.entregas || []).map(e => e.id));
  const entregasAfetadas = DB.entregas.filter(e =>
    !idsNoBackup.has(e.id) || (e.updatedAt || 0) > (snapshot.criadoEm || 0)
  ).length;
  const clientesAfetados = Math.max(0, DB.clientes.length - (snapshot.clientes || []).length);

  const html = `
    <div class="modal-header">
      <h2>Restaurar backup local</h2>
      <button class="close-x" onclick="closeModal()">✕</button>
    </div>
    <div class="hint" style="margin-bottom:10px;">Backup feito em ${quando}.</div>
    ${entregasAfetadas > 0 ? `
      <div class="card" style="background:#FDEDEA;border:1px solid #F1B3A6;margin-bottom:12px;">
        <strong style="color:var(--red-ink);">Atenção: isso vai apagar ou reverter ${entregasAfetadas} entrega${entregasAfetadas > 1 ? 's' : ''} lançada${entregasAfetadas > 1 ? 's' : ''}/alterada${entregasAfetadas > 1 ? 's' : ''} depois desse backup.</strong>
      </div>` : ''}
    ${clientesAfetados > 0 ? `<div class="hint" style="margin-bottom:12px;">Também ${clientesAfetados} loja(s) cadastrada(s) depois desse backup deixariam de existir.</div>` : ''}
    <div class="hint" style="margin-bottom:10px;">Antes de restaurar, o app vai salvar automaticamente um backup do que está aqui agora, então dá pra desfazer se for engano.</div>
    <div class="field">
      <label>Para confirmar, digite <strong>RESTAURAR</strong></label>
      <input type="text" id="f-confirma-restaurar" autocomplete="off" autocapitalize="characters" placeholder="RESTAURAR">
    </div>
    <div class="formbtns">
      <button class="btn ghost block" onclick="closeModal()">Cancelar</button>
      <button class="btn primary block" style="background:var(--coral-dark);" onclick="confirmarRestaurarBackupLocal()">Restaurar mesmo assim</button>
    </div>
  `;
  openModal(html);
}

function confirmarRestaurarBackupLocal() {
  if (!podeImportar()) { toast('Só administradores podem restaurar backups.'); closeModal(); return; }
  const digitado = (document.getElementById('f-confirma-restaurar').value || '').trim().toUpperCase();
  if (digitado !== 'RESTAURAR') { toast('Digite RESTAURAR (em maiúsculas) para confirmar.'); return; }

  const raw = localStorage.getItem('picole_backup_pre_import');
  if (!raw) { toast('Nenhum backup local encontrado neste aparelho.'); closeModal(); return; }
  let snapshot;
  try { snapshot = JSON.parse(raw); } catch (e) { toast('Backup corrompido.'); closeModal(); return; }

  // salva o estado atual antes de sobrescrever, pra sempre dar pra voltar atrás
  try {
    const snapshotAtual = { clientes: DB.clientes, feriados: DB.feriados, entregas: DB.entregas, motoristas: DB.motoristas, categoriasFinanceiras: DB.categoriasFinanceiras, lancamentos: DB.lancamentos, criadoEm: Date.now() };
    localStorage.setItem('picole_backup_antes_de_restaurar', JSON.stringify(snapshotAtual));
  } catch (e) {
    console.error('backup de segurança antes de restaurar falhou', e);
  }

  DB.clientes = snapshot.clientes || [];
  DB.feriados = snapshot.feriados || [];
  DB.entregas = snapshot.entregas || [];
  DB.motoristas = snapshot.motoristas || [];
  DB.categoriasFinanceiras = snapshot.categoriasFinanceiras || [];
  DB.lancamentos = snapshot.lancamentos || [];
  migrarTipos();
  saveDB();
  closeModal();
  refreshAll();
  toast('Backup restaurado. O estado anterior também foi salvo, caso precise desfazer.');
}

function desfazerRestauracao() {
  if (!podeImportar()) { toast('Só administradores podem desfazer uma restauração.'); return; }
  const raw = localStorage.getItem('picole_backup_antes_de_restaurar');
  if (!raw) { toast('Não há uma restauração recente para desfazer neste aparelho.'); return; }
  let snapshot;
  try { snapshot = JSON.parse(raw); } catch (e) { toast('Backup corrompido.'); return; }
  const quando = new Date(snapshot.criadoEm).toLocaleString('pt-BR');
  if (!confirm(`Voltar para o estado de antes da última restauração (salvo em ${quando})?`)) return;
  DB.clientes = snapshot.clientes || [];
  DB.feriados = snapshot.feriados || [];
  DB.entregas = snapshot.entregas || [];
  DB.motoristas = snapshot.motoristas || [];
  DB.categoriasFinanceiras = snapshot.categoriasFinanceiras || [];
  DB.lancamentos = snapshot.lancamentos || [];
  migrarTipos();
  saveDB();
  refreshAll();
  toast('Restauração desfeita.');
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
  let semTipoReconhecido = 0;

  // capturados ANTES de qualquer substituição, para servir de fallback quando a célula "Tipo" vier vazia
  const clienteAntigoPorId = new Map(DB.clientes.map(c => [c.id, c]));
  const clienteAntigoPorNome = new Map(DB.clientes.map(c => [c.nome, c]));
  const entregaAntigaPorId = new Map(DB.entregas.map(e => [e.id, e]));

  function resolverTipo(celulaTipo, existente) {
    const tipoDetectado = normalizarTipoImportado(celulaTipo); // null = vazia ou não reconhecida
    if (tipoDetectado) return tipoDetectado;
    if (existente) return existente.tipo;
    semTipoReconhecido++;
    return 'personalizada';
  }

  const shClientes = wb.Sheets['Clientes'];
  if (shClientes) {
    XLSX.utils.sheet_to_json(shClientes, { defval: '' }).forEach(r => {
      const nome = String(r['Nome da Loja'] || '').trim();
      if (!nome) return;
      const ativoRaw = String(r['Ativo'] || '').trim().toLowerCase();
      const ativo = !['não', 'nao', 'false', 'inativo'].includes(ativoRaw);
      const idCelula = String(r['ID'] || '').trim();
      const existente = clienteAntigoPorId.get(idCelula) || clienteAntigoPorNome.get(nome);
      novo.clientes.push({
        id: uid('c'), tipo: resolverTipo(r['Tipo'], existente), nome,
        endereco: String(r['Endereço'] || r['Endereco'] || ''),
        contato: String(r['Contato'] || ''),
        intervaloPadrao: Number(r['Intervalo Padrao Dias']) || 7,
        valorPadrao: Number(r['Valor Padrao'] || r['Valor Padrão']) || 0,
        quantidadePadrao: Number(r['Quantidade Padrao'] || r['Quantidade Padrão']) || 0,
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

  // Parcelas (Entregas Comuns) ficam numa aba à parte, ligadas pela coluna "ID Entrega"
  const parcelasPorEntrega = new Map();
  const shParcelas = wb.Sheets['Parcelas'];
  if (shParcelas) {
    XLSX.utils.sheet_to_json(shParcelas, { defval: '' }).forEach(r => {
      const idEntrega = String(r['ID Entrega'] || '').trim();
      const vencimento = excelDateToKey(r['Vencimento']);
      if (!idEntrega || !vencimento) return;
      if (!parcelasPorEntrega.has(idEntrega)) parcelasPorEntrega.set(idEntrega, []);
      parcelasPorEntrega.get(idEntrega).push({
        id: uid('p'),
        valor: Number(r['Valor Parcela']) || 0,
        vencimento,
        dataPagamento: excelDateToKey(r['Data Pagamento']),
      });
    });
  }

  let importadas = 0;
  const shEntregas = wb.Sheets['Entregas'];
  if (shEntregas) {
    XLSX.utils.sheet_to_json(shEntregas, { defval: '' }).forEach(r => {
      const dataKey = excelDateToKey(r['Data']);
      const cliente = String(r['Cliente'] || '').trim();
      const quantidade = Number(r['Quantidade']) || 0;
      if (!dataKey || !cliente || quantidade <= 0) return;
      const id = String(r['ID'] || '').trim() || uid('e');
      const tipo = resolverTipo(r['Tipo'], entregaAntigaPorId.get(id));
      const base = {
        id, tipo, data: dataKey, cliente, quantidade,
        motorista: String(r['Motorista'] || '').trim() || null,
        observacoes: String(r['Observações'] || r['Observacoes'] || ''),
        updatedAt: Date.now(), deletado: false,
      };
      if (tipo === 'comum') {
        const parcelas = parcelasPorEntrega.get(id) || gerarParcelasPadrao(1, Number(r['Valor Total']) || 0, dataKey, 30);
        novo.entregas.push({ ...base, parcelas });
      } else {
        novo.entregas.push({
          ...base,
          valorUnitario: Number(r['Valor Unitario'] || r['Valor Unitário']) || 0,
          dataPagamento: excelDateToKey(r['Data Pagamento']),
          obsPagamento: String(r['Observações de Pagamento'] || r['Observacoes de Pagamento'] || ''),
        });
      }
      importadas++;
    });
  }

  // Motoristas: mescla com os já cadastrados em vez de substituir (evita perder motoristas se a aba não vier)
  const shMotoristas = wb.Sheets['Motoristas'];
  let motoristasFinal = DB.motoristas || [];
  if (shMotoristas) {
    XLSX.utils.sheet_to_json(shMotoristas, { defval: '' }).forEach(r => {
      const nome = String(r['Nome'] || '').trim();
      if (!nome) return;
      const ativoRaw = String(r['Ativo'] || '').trim().toLowerCase();
      const ativo = !['não', 'nao', 'false', 'inativo'].includes(ativoRaw);
      const existente = motoristasFinal.find(m => m.nome === nome);
      if (existente) existente.ativo = ativo;
      else motoristasFinal.push({ id: uid('m'), nome, ativo });
    });
  }

  // as lojas/entregas importadas substituem só os tipos presentes na planilha; os demais tipos são preservados
  const tiposNaPlanilha = new Set(novo.clientes.map(c => c.tipo).concat(novo.entregas.map(e => e.tipo)));
  const clientesPreservados = DB.clientes.filter(c => !tiposNaPlanilha.has(c.tipo));
  const entregasPreservadas = DB.entregas.filter(e => !tiposNaPlanilha.has(e.tipo));

  DB = {
    clientes: [...novo.clientes, ...clientesPreservados],
    feriados: novo.feriados,
    entregas: [...novo.entregas, ...entregasPreservadas],
    motoristas: motoristasFinal,
  };
  migrarTipos();
  saveDB();
  refreshAll();
  showScreen('painel');
  const avisoTipo = semTipoReconhecido > 0
    ? ` ⚠️ ${semTipoReconhecido} linha(s) sem coluna "Tipo" preenchida foram assumidas como Personalizadas.`
    : '';
  toast(`Importado: ${novo.clientes.length} loja(s), ${importadas} entrega(s).${avisoTipo}`);
}

async function exportarPlanilha() {
  if (!podeExportar()) { toast('Você não tem permissão para exportar a planilha.'); return; }
  toast('Preparando planilha…');
  try {
    await ensureSheetJS();
  } catch (err) {
    toast('Não foi possível exportar. Verifique sua internet.');
    return;
  }

  const wb = XLSX.utils.book_new();

  const clientesRows = DB.clientes.map((c) => ({
    'ID': c.id,
    'Tipo': TIPO_LABEL_CURTO[c.tipo] || 'Personalizadas',
    'Nome da Loja': c.nome,
    'Endereço': c.endereco || '',
    'Contato': c.contato || '',
    'Intervalo Padrao Dias': c.intervaloPadrao,
    'Valor Padrao': c.valorPadrao || '',
    'Quantidade Padrao': c.quantidadePadrao || '',
    'Ativo': c.ativo === false ? 'Não' : 'Sim',
  }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(clientesRows), 'Clientes');

  const feriadosRows = DB.feriados.slice().sort((a, b) => a.data.localeCompare(b.data)).map(f => ({
    'Data': fmtDateBR(f.data),
    'Descrição': f.descricao,
  }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(feriadosRows), 'Feriados');

  const motoristasRows = DB.motoristas.slice().sort((a, b) => a.nome.localeCompare(b.nome)).map(m => ({
    'Nome': m.nome,
    'Ativo': m.ativo === false ? 'Não' : 'Sim',
  }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(motoristasRows), 'Motoristas');

  const entregasAtivasTodas = entregasAtivas().slice().sort((a, b) => a.data.localeCompare(b.data));
  const entregasRows = entregasAtivasTodas.map(e => ({
    'ID': e.id,
    'Tipo': TIPO_LABEL_CURTO[e.tipo] || 'Personalizadas',
    'Data': fmtDateBR(e.data),
    'Cliente': e.cliente,
    'Quantidade': e.quantidade,
    'Motorista': e.motorista || '',
    'Observações': e.observacoes || '',
    'Valor Unitario': e.tipo === 'comum' ? '' : (e.valorUnitario || 0),
    'Valor Total': valorTotalEntrega(e),
    'Data Pagamento': e.tipo === 'comum' ? '' : (e.dataPagamento ? fmtDateBR(e.dataPagamento) : ''),
    'Observações de Pagamento': e.tipo === 'comum' ? '' : (e.obsPagamento || ''),
  }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(entregasRows), 'Entregas');

  const parcelasRows = [];
  entregasAtivasTodas.filter(e => e.tipo === 'comum').forEach(e => {
    (e.parcelas || []).forEach(p => {
      parcelasRows.push({
        'ID Entrega': e.id,
        'Cliente': e.cliente,
        'Data Entrega': fmtDateBR(e.data),
        'Valor Parcela': p.valor || 0,
        'Vencimento': p.vencimento ? fmtDateBR(p.vencimento) : '',
        'Data Pagamento': p.dataPagamento ? fmtDateBR(p.dataPagamento) : '',
      });
    });
  });
  if (parcelasRows.length > 0) {
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(parcelasRows), 'Parcelas');
  }

  garantirCategoriasVenda();
  const categoriasRows = DB.categoriasFinanceiras.slice().sort((a, b) => a.nome.localeCompare(b.nome)).map(c => ({
    'Tipo': c.tipo === 'entrada' ? 'Entrada' : 'Despesa',
    'Nome': c.nome,
    'Automática': c.sistema ? 'Sim' : 'Não',
    'Ativa': c.ativo === false ? 'Não' : 'Sim',
  }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(categoriasRows), 'Categorias Financeiras');

  const lancamentosRows = lancamentosAtivos().slice().sort((a, b) => a.data.localeCompare(b.data)).map(l => ({
    'ID': l.id,
    'Tipo': l.tipoMov === 'entrada' ? 'Entrada' : 'Saída',
    'Categoria': l.categoria,
    'Nome': l.nome || '',
    'Data': fmtDateBR(l.data),
    'Valor': l.valor,
    'Automático': l.automatico ? 'Sim' : 'Não',
    'Observações': l.observacoes || '',
  }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(lancamentosRows), 'Financeiro');

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
showScreen(TELA_INICIAL_POR_AREA[AREA_ATUAL] || 'painel');
tentarRetomarLogin();
