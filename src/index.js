function getApiBase() {
  let raw =
    sessionStorage.getItem('api_base') ||
    sessionStorage.getItem('apibase') ||
    '';

  raw = String(raw || '').trim();

  if (!raw) return '';

  if (!/^https?:\/\//i.test(raw)) {
    raw = `https://${raw}`;
  }

  try {
    const url = new URL(raw);
    return url.href.replace(/\/+$/, '');
  } catch {
    return '';
  }
}

function apiUrl(path) {
  const base = getApiBase();
  if (!path.startsWith('/')) path = `/${path}`;
  return `${base}${path}`;
}

function absUrlFromApi(relOrAbs, apiBase) {
  const s = String(relOrAbs || '').trim();
  if (!s) return '';
  if (/^https?:\/\//i.test(s)) return s;

  const base = String(apiBase || getApiBaseGestaoUsuarios() || '').trim();
  if (!base) return s;

  try {
    if (s.startsWith('/foto-usuario/')) {
      return new URL(`/publicidade${s}`, `${base}/`).href;
    }

    if (s.startsWith('/marketing/')) {
      return new URL(`/publicidade${s}`, `${base}/`).href;
    }

    return new URL(s, `${base}/`).href;
  } catch {
    return s;
  }
}

function DefinirSidebarAberta(abrir){
  const sidebar = document.getElementById('sidebar');
  const overlay = document.getElementById('overlay');
  const floatingMenuBtn = document.getElementById('floatingMenuBtn');

  sidebar?.classList.toggle('is-open', abrir);
  overlay?.classList.toggle('show', abrir);
  floatingMenuBtn?.classList.toggle('is-hidden', abrir);
}

function DefinirPaginaAtiva(pageId, itemClicado){
  document.querySelectorAll('.menu-item').forEach(i => i.classList.remove('bg-secondary'));
  itemClicado?.classList.add('bg-secondary');

  document.querySelectorAll('.page-content').forEach(p => p.classList.remove('active'));
  document.getElementById(pageId)?.classList.add('active');

  if (window.innerWidth <= 900) DefinirSidebarAberta(false);
}

function InicializarHoje(){
  const hoje = new Date();
  const yyyy = hoje.getFullYear();
  const mm = String(hoje.getMonth() + 1).padStart(2,'0');
  const dd = String(hoje.getDate()).padStart(2,'0');
  const iso = `${yyyy}-${mm}-${dd}`;

  const input = document.getElementById('calendarInput');
  if (input) input.value = iso;

  const label = document.getElementById('todayLabel');
  if (label) {
    const dataFormatada = hoje.toLocaleDateString('pt-BR', {
      weekday: 'long',
      year: 'numeric',
      month: 'long',
      day: 'numeric'
    });
    label.textContent = `Hoje é ${dataFormatada}`;
  }

  const usuarioLogin = document.getElementById('helloTitleHome');
  if (usuarioLogin) {
    const nomeCompleto = (sessionStorage.getItem('usuario') || '').trim();

    if (!nomeCompleto) {
      usuarioLogin.textContent = 'Olá, Usuário';
    } else {
      const partes = nomeCompleto.split(/\s+/).filter(Boolean); // separa por 1+ espaços [web:122]
      const primeiro = partes[0];
      const ultimo = partes.length > 1 ? partes[partes.length - 1] : '';
      usuarioLogin.textContent = `Olá, ${ultimo ? `${primeiro} ${ultimo}` : primeiro}`;
    }
  }

}

function ObterPrimeiroEUltimoNome(nomeCompleto) {
  const limpo = (nomeCompleto || '').trim();
  if (!limpo) return { primeiro: '', ultimo: '', exibicao: '' };

  const partes = limpo.split(/\s+/).filter(Boolean); // split por espaços [web:122]
  const primeiro = partes[0] || '';
  const ultimo = (partes.length > 1) ? partes[partes.length - 1] : '';
  const exibicao = ultimo ? `${primeiro} ${ultimo}` : primeiro;

  return { primeiro, ultimo, exibicao };
}

function ObterDominioDoEmail(email) {
  const e = (email || '').trim();
  const at = e.indexOf('@');
  if (at === -1) return '';
  return e.substring(at + 1);
}


function FormatarDataHoraLocal(data) {
  const ano = data.getFullYear();
  const mes = String(data.getMonth() + 1).padStart(2, '0');
  const dia = String(data.getDate()).padStart(2, '0');
  const hora = String(data.getHours()).padStart(2, '0');
  const minuto = String(data.getMinutes()).padStart(2, '0');
  return `${ano}-${mes}-${dia}T${hora}:${minuto}`;
}

function RemoverModalAgendamentoSala() {
  document.getElementById('roomModalOverlay')?.remove();
  document.getElementById('roomModal')?.remove();
}

function AbrirModalAgendamentoSala() {
  RemoverModalAgendamentoSala();

  const overlay = document.createElement('div');
  overlay.id = 'roomModalOverlay';
  overlay.className = 'fixed inset-0 bg-black/30 backdrop-blur-sm z-[70]';
  document.body.appendChild(overlay);

  const modal = document.createElement('div');
  modal.id = 'roomModal';
  modal.className = 'fixed inset-0 z-[80]';
  modal.setAttribute('role', 'dialog');
  modal.setAttribute('aria-modal', 'true');
  modal.setAttribute('aria-labelledby', 'roomModalTitle');

  modal.innerHTML = `
    <div class="w-full h-full overflow-y-auto md:overflow-hidden no-scrollbar">
      <div class="min-h-full flex items-start justify-center p-4 md:p-8">
        <div class="w-full max-w-5xl mx-auto px-4 sm:px-6">
          <!-- Card fixo: sem scroll no card -->
          <div class="glass rounded-2xl shadow-2xl border border-border overflow-hidden h-auto md:h-[92vh] flex flex-col">

            <!-- Header fixo -->
            <div class="px-6 py-5 border-b border-border flex items-start justify-between gap-4 shrink-0">
              <div>
                <h3 id="roomModalTitle" class="text-xl font-semibold text-foreground">Agendar sala de reunião</h3>
                <p class="text-sm text-muted-foreground">Selecione a sala, data, horário e participantes</p>
              </div>

              <button id="closeRoomModal" type="button"
                class="w-10 h-10 rounded-xl bg-white/60 border border-border hover:bg-white transition-all flex items-center justify-center"
                aria-label="Fechar" title="Fechar">
                <i class="fas fa-times"></i>
              </button>
            </div>

            <!-- Conteúdo: 1 col no mobile, 2 cols no md+ -->
            <div class="px-6 py-6 flex-1 overflow-hidden">
              <form id="roomForm" class="h-full">
                <div class="h-full grid grid-cols-1 md:grid-cols-2 gap-6 overflow-hidden">

                  <!-- COLUNA 1: dados do agendamento -->
                  <div class="h-full overflow-hidden flex flex-col">
                    <div class="space-y-4">
                      <div class="space-y-2">
                        <label for="roomName" class="text-sm font-medium">Sala de reunião</label>
                        <select id="roomName"
                          class="w-full rounded-xl border border-border bg-white/70 px-4 py-3 outline-none focus:ring-2 focus:ring-primary/30"
                          required>
                          <option value="" selected disabled hidden>Selecione uma sala...</option>
                          <option value="Sala 01">Sala 01</option>
                          <option value="Sala 02">Sala 02</option>
                          <option value="Sala Diretoria">Sala Diretoria</option>
                        </select>
                      </div>

                      <div class="grid grid-cols-1 sm:grid-cols-2 gap-4">
                        <div class="space-y-2">
                          <label for="roomStartDT" class="text-sm font-medium">Data e hora início</label>
                          <input id="roomStartDT" type="datetime-local"
                            class="w-full rounded-xl border border-border bg-white/70 px-4 py-3 outline-none focus:ring-2 focus:ring-primary/30"
                            required />
                        </div>

                        <div class="space-y-2">
                          <label for="roomEndDT" class="text-sm font-medium">Data e hora fim</label>
                          <input id="roomEndDT" type="datetime-local"
                            class="w-full rounded-xl border border-border bg-white/70 px-4 py-3 outline-none focus:ring-2 focus:ring-primary/30"
                            required />
                        </div>
                      </div>

                      <div class="space-y-2">
                        <label for="roomReason" class="text-sm font-medium">Motivo</label>
                        <input id="roomReason" type="text"
                          class="w-full rounded-xl border border-border bg-white/70 px-4 py-3 outline-none focus:ring-2 focus:ring-primary/30"
                          placeholder="Ex.: Reunião de planejamento"
                          required />
                      </div>

                      <p id="roomFormErro" class="text-sm text-destructive hidden whitespace-pre-line"></p>
                    </div>
                  </div>

                  <!-- COLUNA 2: participantes + botões -->
                  <div class="h-full overflow-hidden flex flex-col">
                    <div class="space-y-4 flex-1 overflow-hidden flex flex-col">

                      <!-- Participantes -->
                      <div class="space-y-2 flex-1 overflow-hidden flex flex-col">
                        <div class="flex items-center justify-between gap-3 flex-wrap">
                          <div class="flex items-center gap-3 flex-wrap">
                            <label class="text-sm font-medium">Participantes</label>
                            <span id="usersCount" class="text-xs text-muted-foreground"></span>
                          </div>

                          <div class="flex items-center gap-2 flex-wrap">
                            <select id="setorFilter"
                              class="rounded-xl border border-border bg-white/70 px-3 py-2 text-sm outline-none focus:ring-2 focus:ring-primary/30"
                              title="Filtrar por setor">
                              <option value="">Todos os setores</option>
                            </select>

                            <button id="btnSelecionarTodosUsers" type="button"
                              class="text-xs px-3 py-2 rounded-xl border border-border bg-white/50 hover:bg-white/70 transition-all">
                              Selecionar todos
                            </button>

                            <button id="btnLimparUsers" type="button"
                              class="text-xs px-3 py-2 rounded-xl border border-border bg-white/50 hover:bg-white/70 transition-all">
                              Limpar
                            </button>
                          </div>
                        </div>

                        <input id="filtroUsers" type="text"
                          class="w-full rounded-xl border border-border bg-white/70 px-4 py-3 outline-none focus:ring-2 focus:ring-primary/30"
                          placeholder="Buscar por nome ou email..." />

                        <!-- Aqui fica o ÚNICO scroll (lista de usuários) -->
                        <div id="usersBox"
                          class="rounded-xl border border-border bg-white/50 p-3 flex-1 overflow-auto no-scrollbar space-y-2 min-h-0">
                          <p id="usersLoading" class="text-sm text-muted-foreground">Carregando usuários...</p>
                        </div>
                      </div>

                      <!-- Botões -->
                      <div class="pt-2 flex flex-col sm:flex-row gap-3 shrink-0">
                        <button id="btnSalvarAgendamentoSala" type="submit"
                          class="sm:flex-1 rounded-xl bg-primary text-white px-4 py-3 font-medium hover:opacity-90 transition-all">
                          Salvar
                        </button>

                        <button id="btnCancelarAgendamentoSala" type="button"
                          class="sm:flex-1 rounded-xl border border-border bg-white/50 px-4 py-3 font-medium hover:bg-white/70 transition-all">
                          Cancelar
                        </button>
                      </div>

                    </div>
                  </div>

                </div>
              </form>
            </div>

          </div>
        </div>
      </div>
    </div>
  `;

  document.body.appendChild(modal);

  const inputInicio = document.getElementById('roomStartDT');
  const inputFim = document.getElementById('roomEndDT');

  const formulario = document.getElementById('roomForm');
  const btnSalvar = document.getElementById('btnSalvarAgendamentoSala');

  const usersBox = document.getElementById('usersBox');
  const usersLoading = document.getElementById('usersLoading');
  const filtroUsers = document.getElementById('filtroUsers');
  const setorFilter = document.getElementById('setorFilter');
  const btnSelTodos = document.getElementById('btnSelecionarTodosUsers');
  const btnLimpar = document.getElementById('btnLimparUsers');
  const usersCount = document.getElementById('usersCount');

  let usuariosCache = []; // {id,nome,email,setor}
  let setoresCache = [];

  const selectedIds = new Set();

  function FecharModalAgendamentoSala() {
    RemoverModalAgendamentoSala();
  }

  overlay.addEventListener('click', FecharModalAgendamentoSala);
  document.getElementById('closeRoomModal')?.addEventListener('click', FecharModalAgendamentoSala);
  document.getElementById('btnCancelarAgendamentoSala')?.addEventListener('click', FecharModalAgendamentoSala);

  btnSalvar?.addEventListener('click', () => {
    formulario?.requestSubmit?.(btnSalvar);
  });

  // -------- Helpers datetime-local --------
  function pad2(n) { return String(n).padStart(2, '0'); }
  function todayDateStr() {
    const d = new Date();
    return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
  }

  function validateNotPastStart() {
    if (!inputInicio?.value) return true;
    const chosen = new Date(inputInicio.value);
    const min = new Date(`${todayDateStr()}T00:00`);
    return chosen >= min;
  }

  function validateEndAfterStart() {
    if (!inputInicio?.value || !inputFim?.value) return true;
    const ini = new Date(inputInicio.value);
    const fim = new Date(inputFim.value);
    return fim > ini;
  }

  (function initDefaults() {
    if (inputInicio) inputInicio.min = `${todayDateStr()}T00:00`;
    if (inputInicio) inputInicio.value = `${todayDateStr()}T07:10`;
    if (inputFim) inputFim.value = `${todayDateStr()}T17:30`;
  })();

  function escapeHtml(s) {
    return String(s ?? '')
      .replaceAll('&', '&amp;')
      .replaceAll('<', '&lt;')
      .replaceAll('>', '&gt;')
      .replaceAll('"', '&quot;')
      .replaceAll("'", '&#039;');
  }

  function setRoomFormError(msg) {
    const el = document.getElementById('roomFormErro');
    if (!el) return;
    if (!msg) { el.textContent = ''; el.classList.add('hidden'); return; }
    el.textContent = msg;
    el.classList.remove('hidden');
  }

  function setSalvarLoading(loading) {
    const btn = document.getElementById('btnSalvarAgendamentoSala');
    if (!btn) return;
    btn.disabled = loading;
    btn.textContent = loading ? 'Salvando...' : 'Salvar';
    btn.classList.toggle('opacity-70', loading);
  }

  function getVisibleUsers() {
    const q = (filtroUsers?.value || '').trim().toLowerCase();
    const setorSel = (setorFilter?.value || '').trim();

    return usuariosCache.filter(u => {
      const nome = (u.nome || '').toLowerCase();
      const email = (u.email || '').toLowerCase();
      const setor = (u.setor || '').toLowerCase();

      const qMatch = !q || nome.includes(q) || email.includes(q);
      const setorMatch = !setorSel || (u.setor || '') === setorSel;

      return qMatch && setorMatch;
    });
  }

  function atualizarContador(visiveis) {
    if (!usersCount) return;
    usersCount.textContent = `${selectedIds.size} selecionado(s) • ${visiveis.length} visível(is)`;
  }

  function renderUsuarios() {
    if (!usersBox) return;

    const visiveis = getVisibleUsers();

    if (!visiveis.length) {
      usersBox.innerHTML = `<p class="text-sm text-muted-foreground">Nenhum usuário encontrado.</p>`;
      atualizarContador([]);
      return;
    }

    usersBox.innerHTML = visiveis.map(u => {
      const idNum = Number(u.id);
      const idStr = String(u.id);

      const nomeAlta = String(u.nome || '').toUpperCase();
      const emailBaixa = String(u.email || '').toLowerCase();
      const setor = u.setor || 'Sem setor';

      const checked = selectedIds.has(idNum) ? 'checked' : '';

      return `
        <label class="flex items-start gap-3 rounded-xl border border-border bg-white/60 hover:bg-white/80 transition-all px-3 py-2 cursor-pointer">
          <input type="checkbox" class="mt-1 checkbox-user" data-user-id="${escapeHtml(idStr)}" ${checked} />
          <div class="min-w-0 flex-1">
            <div class="text-sm font-semibold truncate">${escapeHtml(nomeAlta)}</div>
            <div class="text-xs text-muted-foreground truncate">${escapeHtml(emailBaixa)}</div>
            <div class="text-xs text-muted-foreground truncate">${escapeHtml(setor)}</div>
          </div>
        </label>
      `;
    }).join('');

    usersBox.querySelectorAll('input.checkbox-user').forEach(chk => {
      chk.addEventListener('change', () => {
        const id = Number(chk.getAttribute('data-user-id'));
        if (!Number.isFinite(id)) return;

        if (chk.checked) selectedIds.add(id);
        else selectedIds.delete(id);

        atualizarContador(visiveis);
      });
    });

    atualizarContador(visiveis);
  }

  function selecionarTodosVisiveis() {
    const visiveis = getVisibleUsers();
    visiveis.forEach(u => {
      const id = Number(u.id);
      if (Number.isFinite(id)) selectedIds.add(id);
    });
    renderUsuarios();
  }

  function limparVisiveis() {
    const visiveis = getVisibleUsers();
    visiveis.forEach(u => {
      const id = Number(u.id);
      if (Number.isFinite(id)) selectedIds.delete(id);
    });
    renderUsuarios();
  }

  function getParticipantesSelecionadosIds() {
    return Array.from(selectedIds.values());
  }

  filtroUsers?.addEventListener('input', renderUsuarios);
  setorFilter?.addEventListener('change', renderUsuarios);
  btnSelTodos?.addEventListener('click', selecionarTodosVisiveis);
  btnLimpar?.addEventListener('click', limparVisiveis);

  async function carregarSetores() {
    const API_BASE = sessionStorage.getItem('api_base') || '';
    if (!API_BASE) throw new Error('API_BASE não configurada (sessionStorage api_base).');

    const resp = await fetch(`${API_BASE}/api/setores`, { method: 'GET' });
    const json = await resp.json().catch(() => ({}));
    if (!resp.ok) throw new Error(json?.message || `Erro ao listar setores. Status: ${resp.status}`);

    setoresCache = Array.isArray(json.items)
      ? json.items.map(s => String(s.nome || '').trim()).filter(Boolean)
      : [];

    if (setorFilter) {
      setorFilter.innerHTML =
        `<option value="">Todos os setores</option>` +
        setoresCache.map(nome => `<option value="${escapeHtml(nome)}">${escapeHtml(nome)}</option>`).join('');
    }
  }

  async function carregarUsuarios() {
    try {
      const API_BASE = sessionStorage.getItem('api_base') || '';
      if (!API_BASE) throw new Error('API_BASE não configurada (sessionStorage api_base).');

      const resp = await fetch(`${API_BASE}/api/usuarios`, { method: 'GET' });
      const json = await resp.json().catch(() => ({}));
      if (!resp.ok) throw new Error(json?.message || `Erro ao listar usuários. Status: ${resp.status}`);

      usuariosCache = Array.isArray(json.items) ? json.items : [];
      renderUsuarios();
    } catch (err) {
      if (usersBox) {
        usersBox.innerHTML = `<p class="text-sm text-destructive whitespace-pre-line">${escapeHtml(err?.message || 'Erro ao carregar usuários.')}</p>`;
      }
    } finally {
      usersLoading?.remove();
    }
  }

  (async () => {
    try { await carregarSetores(); } catch (e) { console.error('Setores:', e); }
    await carregarUsuarios();
  })();

  // -------- Modal de conflito (mantido) --------
  function RemoverModalConflito() {
    document.getElementById('conflitoOverlay')?.remove();
    document.getElementById('conflitoModal')?.remove();
  }

  function fmtBR(v) {
    if (!v) return '(não informado)';
    const s = String(v);
    const semZ = s.endsWith('Z') ? s.slice(0, -1) : s;
    const d = new Date(semZ);
    if (Number.isNaN(d.getTime())) return String(v);
    return d.toLocaleString('pt-BR');
  }

  function AbrirModalConflito({ sala, detalhe }) {
    RemoverModalConflito();

    const overlayC = document.createElement('div');
    overlayC.id = 'conflitoOverlay';
    overlayC.className = 'fixed inset-0 bg-black/40 backdrop-blur-sm z-[90]';
    document.body.appendChild(overlayC);

    const modalC = document.createElement('div');
    modalC.id = 'conflitoModal';
    modalC.className = 'fixed inset-0 z-[100]';
    modalC.setAttribute('role', 'dialog');
    modalC.setAttribute('aria-modal', 'true');
    modalC.setAttribute('aria-labelledby', 'conflitoTitle');

    const c = detalhe || {};
    const agendadoPor = c.usuario_agendamento || '(não informado)';

    modalC.innerHTML = `
      <div class="w-full h-full flex items-start justify-center p-4 md:p-8">
        <div class="w-full max-w-xl mx-auto px-4 sm:px-6">
          <div class="glass rounded-2xl shadow-2xl border border-border overflow-hidden">
            <div class="px-6 py-5 border-b border-border flex items-start justify-between gap-4">
              <div>
                <h3 id="conflitoTitle" class="text-xl font-semibold text-foreground">Conflito de agendamento</h3>
                <p class="text-sm text-muted-foreground">Já existe um agendamento nesse intervalo</p>
              </div>

              <button id="btnFecharConflito" type="button"
                class="w-10 h-10 rounded-xl bg-white/60 border border-border hover:bg-white transition-all flex items-center justify-center"
                aria-label="Fechar" title="Fechar">
                <i class="fas fa-times"></i>
              </button>
            </div>

            <div class="px-6 py-6 space-y-3 text-sm">
              <div class="rounded-xl border border-border bg-white/60 p-4 space-y-2">
                <div><span class="text-muted-foreground">Sala:</span> <span class="font-semibold text-foreground">${escapeHtml(sala || '(não informado)')}</span></div>
                <div><span class="text-muted-foreground">Agendado por:</span> <span class="font-semibold text-foreground">${escapeHtml(agendadoPor)}</span></div>
                <div><span class="text-muted-foreground">Motivo:</span> <span class="font-semibold text-foreground">${escapeHtml(c.motivo || '(não informado)')}</span></div>
                <div><span class="text-muted-foreground">Início:</span> <span class="font-semibold text-foreground">${escapeHtml(fmtBR(c.inicio))}</span></div>
                <div><span class="text-muted-foreground">Fim:</span> <span class="font-semibold text-foreground">${escapeHtml(fmtBR(c.fim))}</span></div>
                <div><span class="text-muted-foreground">Registro:</span> <span class="font-semibold text-foreground">${escapeHtml(fmtBR(c.data_agendamento))}</span></div>
              </div>

              <div class="pt-2">
                <button id="btnOkConflito" type="button"
                  class="w-full rounded-xl bg-primary text-white px-4 py-3 font-medium hover:opacity-90 transition-all">
                  OK
                </button>
              </div>
            </div>
          </div>
        </div>
      </div>
    `;

    document.body.appendChild(modalC);

    const fechar = () => RemoverModalConflito();
    overlayC.addEventListener('click', fechar);
    document.getElementById('btnFecharConflito')?.addEventListener('click', fechar);
    document.getElementById('btnOkConflito')?.addEventListener('click', fechar);
  }

  // -------- Submit --------
  formulario?.addEventListener('submit', async (e) => {
    e.preventDefault();
    setRoomFormError('');

    if (!formulario.reportValidity()) return;

    if (!validateNotPastStart()) {
      setRoomFormError('Não é permitido selecionar dias anteriores a hoje para a data/hora início.');
      return;
    }

    if (!validateEndAfterStart()) {
      setRoomFormError('A data/hora fim deve ser maior que a data/hora início.');
      return;
    }

    const sala = document.getElementById('roomName')?.value;
    const inicio = document.getElementById('roomStartDT')?.value;
    const fim = document.getElementById('roomEndDT')?.value;
    const motivo = document.getElementById('roomReason')?.value?.trim();

    if (!sala || !inicio || !fim || !motivo) return;

    const usuarioLogin = sessionStorage.getItem('usuario') || 'desconhecido';
    const criadoEmISO = new Date().toISOString();

    const payload = {
      sala, inicio, fim, motivo,
      usuario: usuarioLogin,
      criadoEm: criadoEmISO,
      participantes: getParticipantesSelecionadosIds()
    };

    try {
      setSalvarLoading(true);

      const API_BASE = sessionStorage.getItem('api_base') || '';
      if (!API_BASE) throw new Error('API_BASE não configurada (sessionStorage api_base).');

      const URL_CHECK = `${API_BASE}/api/agendamentos/sala/verificar`;
      const URL_SAVE  = `${API_BASE}/api/agendamentos/sala`;

      const checkResp = await fetch(URL_CHECK, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ sala: payload.sala, inicio: payload.inicio, fim: payload.fim })
      });

      const checkJson = await checkResp.json().catch(() => ({}));
      if (!checkResp.ok) throw new Error(checkJson?.message || `Erro ao verificar. Status: ${checkResp.status}`);

      if (checkJson?.conflito) {
        const c = checkJson.conflitoDetalhe || checkJson.conflito || {};
        AbrirModalConflito({ sala: payload.sala, detalhe: c });
        return;
      }

      const resp = await fetch(URL_SAVE, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });

      const saveJson = await resp.json().catch(() => ({}));
      if (!resp.ok) throw new Error(saveJson?.message || `Falha ao salvar. Status: ${resp.status}`);

      FecharModalAgendamentoSala();
      alert(saveJson?.message || 'Agendamento salvo com sucesso!');
    } catch (err) {
      setRoomFormError(err?.message || 'Erro ao salvar o agendamento.');
    } finally {
      setSalvarLoading(false);
    }
  });
}


// ===== Sessão expira após 4h sem atividade =====
(function SessionInatividade4h() {
  const LIMITE_MS = 4 * 60 * 60 * 1000; // 4 horas
  const KEY_OPEN = 'sessao_aberta_em_ms';
  const KEY_LAST = 'sessao_ultima_atividade_ms';

  function agoraMs() { return Date.now(); }

  function marcarAberturaSePreciso() {
    const opened = Number(sessionStorage.getItem(KEY_OPEN) || 0);
    if (!opened) sessionStorage.setItem(KEY_OPEN, String(agoraMs()));
  }

  function marcarAtividade() {
    sessionStorage.setItem(KEY_LAST, String(agoraMs()));
  }

  function expirarSessao(motivo) {
    // mantenha se quiser guardar uma mensagem para a tela de login
    // sessionStorage.setItem('logout_reason', motivo);

    sessionStorage.clear();
    alert(motivo || 'Sessão expirada. Faça login novamente.');
    window.location.href = 'index.html'; // ajuste para sua página de login
  }

  function checarInatividade() {
    const last = Number(sessionStorage.getItem(KEY_LAST) || 0);
    const opened = Number(sessionStorage.getItem(KEY_OPEN) || 0);
    const base = last || opened || agoraMs();

    if (agoraMs() - base >= LIMITE_MS) {
      expirarSessao('Sessão expirada por inatividade (4 horas). Faça login novamente.');
    }
  }

  let _resumoIntervalId = null;

  function HomeEstaAtiva() {
    const home = document.getElementById('home');
    return !!home && home.classList.contains('active');
  }

  function IniciarAutoRefreshResumoDia() {
    if (_resumoIntervalId) return; // já iniciou

    // roda imediatamente ao entrar no home
    if (HomeEstaAtiva()) CarregarResumoDoDia();

    _resumoIntervalId = setInterval(() => {
      // opcional: evita rodar em aba oculta
      if (document.visibilityState !== 'visible') return;

      if (HomeEstaAtiva()) {
        CarregarResumoDoDia();
      }
    }, 10_000); // 10s [web:294]
  }

  function PararAutoRefreshResumoDia() {
    if (!_resumoIntervalId) return;
    clearInterval(_resumoIntervalId);
    _resumoIntervalId = null;
  }

  // Inicializa (quando a página abre)
  document.addEventListener('DOMContentLoaded', () => {

    // 1) CHECAGEM DE AUTENTICAÇÃO (PRIMEIRO)
    const apiBase = sessionStorage.getItem('api_base');
    const usuario = sessionStorage.getItem('usuario');
    const email = sessionStorage.getItem('userEmail');

    if (!apiBase || !usuario || !email) {
      // se qualquer um estiver vazio/null/undefined, redireciona para login
      window.location.href = 'index.html';
      return; // para o resto da função
    }

    marcarAberturaSePreciso();
    marcarAtividade(); // considera atividade na abertura

    IniciarAutoRefreshResumoDia();

    // se trocar de aba e voltar, atualiza na hora
    document.addEventListener('visibilitychange', () => {
      if (document.visibilityState === 'visible' && HomeEstaAtiva()) {
        CarregarResumoDoDia();
      }
    });

    // Eventos que contam como atividade
    const eventos = ['click', 'keydown', 'mousemove', 'scroll', 'touchstart'];
    eventos.forEach((ev) => window.addEventListener(ev, marcarAtividade, { passive: true }));

    // Checa a cada 30s
    setInterval(checarInatividade, 30 * 1000); // setInterval [web:294]

    // Checa também quando o usuário volta para a aba
    document.addEventListener('visibilitychange', () => {
      if (document.visibilityState === 'visible') checarInatividade();
    });
    DefinirSidebarAberta(false);

    const elAvatar = document.getElementById('userAvatar');
    const elNome = document.getElementById('userDisplayName');
    const elDominio = document.getElementById('userDomain');

    const nomeCompleto = sessionStorage.getItem('usuario') || '';

    const { primeiro, ultimo, exibicao } = ObterPrimeiroEUltimoNome(nomeCompleto);
    const dominio = ObterDominioDoEmail(email);

    // Nome (primeiro + último)
    if (elNome) elNome.textContent = exibicao || 'Usuário'; // textContent [web:147]

    // Domínio (depois do @)
    if (elDominio) elDominio.textContent = dominio || 'Domínio';

    // Iniciais (primeiro + último)
    const inicial1 = primeiro ? primeiro[0].toUpperCase() : '';
    const inicial2 = ultimo ? ultimo[0].toUpperCase() : (primeiro ? primeiro[0].toUpperCase() : '');
    const iniciais = (inicial1 + inicial2) || 'U';

    if (elAvatar) elAvatar.textContent = iniciais;

    CarregarResumoDoDia();
    document.getElementById('calendarInput')?.addEventListener('change', CarregarResumoDoDia);
    document.getElementById('floatingMenuBtn')?.addEventListener('click', () => DefinirSidebarAberta(true));
    document.getElementById('closeSidebarBtn')?.addEventListener('click', () => DefinirSidebarAberta(false));
    document.getElementById('overlay')?.addEventListener('click', () => DefinirSidebarAberta(false));

    document.querySelectorAll('.menu-item').forEach(item => {
      item.addEventListener('click', function () {
        const page = this.dataset.page;

        DefinirPaginaAtiva(page, this);

        // Se abriu a seção de clientes, carrega automaticamente
        if (page === 'secao-clientes') {
          carregarClientes();
        }
      });
    });


    document.getElementById('logoutBtn')?.addEventListener('click', function(){
      if (confirm('Deseja sair do sistema?')){
        sessionStorage.clear();
        window.location.href = 'index.html';
      }
    });

    document.getElementById('AgendarSalaReuniao')?.addEventListener('click', AbrirModalAgendamentoSala);

    InicializarHoje();
    iniciarLoopMarketingPainel().catch(console.error);

    document.getElementById('btnAtualizarClientes')?.addEventListener('click', carregarClientes);
    document.getElementById('btnNovoCliente')?.addEventListener('click', () => abrirModalCliente({ modo: 'new', cliente: null }));
    document.getElementById('inputBuscaClientes')?.addEventListener('input', () => renderTabelaClientes());

    document.addEventListener('click', async (e) => {
      const btnEdit = e.target.closest('.btnEditarCliente');
      const btnDel = e.target.closest('.btnExcluirCliente');

      if (btnEdit) {
        const id = btnEdit.getAttribute('data-id');
        const cli = clientesCache.find(x => String(x.ID) === String(id));
        if (!cli) return;
        abrirModalCliente({ modo: 'edit', cliente: cli });
      }

      if (btnDel) {
        const id = btnDel.getAttribute('data-id');
        if (!confirm('Desativar este cliente?')) return;
        try {
          await apiSend(`/api/clientes/${encodeURIComponent(id)}`, 'DELETE');
          await carregarClientes();
        } catch (err) {
          alert(err?.message || 'Erro ao desativar.');
        }
      }
    });


  });
})();



function fmtDataHora(v) {
  if (v == null) return '';
  const s = String(v);
  const semZ = s.endsWith('Z') ? s.slice(0, -1) : s;
  const d = new Date(semZ);
  if (Number.isNaN(d.getTime())) return String(v);
  return d.toLocaleString('pt-BR', { hour: '2-digit', minute: '2-digit' });
}

function fmtDataHoraCompleta(v) {
  if (v == null) return '';
  const s = String(v);
  const semZ = s.endsWith('Z') ? s.slice(0, -1) : s;
  const d = new Date(semZ);
  if (Number.isNaN(d.getTime())) return String(v);
  return d.toLocaleString('pt-BR');
}


function RemoverModalDetalheAgendamento() {
  document.getElementById('agDetalheOverlay')?.remove();
  document.getElementById('agDetalheModal')?.remove();
}

function AbrirModalDetalheAgendamento(item) {
  RemoverModalDetalheAgendamento();

  const overlay = document.createElement('div');
  overlay.id = 'agDetalheOverlay';
  overlay.className = 'fixed inset-0 bg-black/40 backdrop-blur-sm z-[90]';
  document.body.appendChild(overlay);

  const modal = document.createElement('div');
  modal.id = 'agDetalheModal';
  modal.className = 'fixed inset-0 z-[100]';
  modal.setAttribute('role', 'dialog');
  modal.setAttribute('aria-modal', 'true');

  modal.innerHTML = `
    <div class="w-full h-full flex items-start justify-center p-4 md:p-8">
      <div class="w-full max-w-xl mx-auto px-4 sm:px-6">
        <div class="glass rounded-2xl shadow-2xl border border-border overflow-hidden">
          <div class="px-6 py-5 border-b border-border flex items-start justify-between gap-4">
            <div>
              <h3 class="text-xl font-semibold text-foreground">Detalhes do agendamento</h3>
            </div>
            <button id="btnFecharAgDetalhe" type="button"
              class="w-10 h-10 rounded-xl bg-white/60 border border-border hover:bg-white transition-all flex items-center justify-center"
              aria-label="Fechar" title="Fechar">
              <i class="fas fa-times"></i>
            </button>
          </div>

          <div class="px-6 py-6 space-y-2 text-sm">
            <div><span class="text-muted-foreground">Sala:</span> <span class="font-semibold text-foreground">${item.sala}</span></div>
            <div><span class="text-muted-foreground">Início:</span> <span class="font-semibold text-foreground">${fmtDataHoraCompleta(item.inicio)}</span></div>
            <div><span class="text-muted-foreground">Fim:</span> <span class="font-semibold text-foreground">${fmtDataHoraCompleta(item.fim)}</span></div>
            <div><span class="text-muted-foreground">Motivo:</span> <span class="font-semibold text-foreground">${item.motivo || '(não informado)'}</span></div>
            <div><span class="text-muted-foreground">Agendado por:</span> <span class="font-semibold text-foreground">${item.usuario_agendamento || '(não informado)'}</span></div>
            <div><span class="text-muted-foreground">Registro:</span> <span class="font-semibold text-foreground">${fmtDataHoraCompleta(item.data_agendamento)}</span></div>

            <div class="pt-4">
              <button id="btnOkAgDetalhe" type="button"
                class="w-full rounded-xl bg-primary text-white px-4 py-3 font-medium hover:opacity-90 transition-all">
                OK
              </button>
            </div>
          </div>
        </div>
      </div>
    </div>
  `;

  document.body.appendChild(modal);

  const fechar = () => RemoverModalDetalheAgendamento();
  overlay.addEventListener('click', fechar);
  document.getElementById('btnFecharAgDetalhe')?.addEventListener('click', fechar);
  document.getElementById('btnOkAgDetalhe')?.addEventListener('click', fechar);
}

async function CarregarResumoDoDia() {
  const lista = document.getElementById('resumoDiaLista');
  const vazio = document.getElementById('resumoDiaVazio');
  const calendar = document.getElementById('calendarInput');

  if (!lista || !vazio) return;

  const API_BASE = sessionStorage.getItem('api_base') || '';
  if (!API_BASE) {
    vazio.textContent = 'API base não configurada.';
    vazio.classList.remove('hidden');
    return;
  }

  const data = calendar?.value || ''; // YYYY-MM-DD
  const url = `${API_BASE}/api/agendamentos/sala/dia?data=${encodeURIComponent(data)}`;

  const resp = await fetch(url, { method: 'GET' });
  const json = await resp.json().catch(() => ({}));
  if (!resp.ok) {
    vazio.textContent = json?.message || 'Erro ao carregar resumo do dia.';
    vazio.classList.remove('hidden');
    return;
  }

  const items = json?.items || [];
  lista.innerHTML = '';

  if (!items.length) {
    vazio.textContent = 'Nenhum evento cadastrado.';
    vazio.classList.remove('hidden');
    return;
  }

  vazio.classList.add('hidden');

  items.forEach((item) => {
    const row = document.createElement('div');
    row.className = 'flex items-center justify-between gap-2 rounded-xl border border-border bg-white/60 hover:bg-white/80 transition-all px-3 py-2';

    // esquerda (clicável)
    const left = document.createElement('button');
    left.type = 'button';
    left.className = 'flex-1 text-left min-w-0';
    left.innerHTML = `
      <div class="flex items-center justify-between gap-2">
        <div class="min-w-0">
          <div class="text-sm font-semibold truncate">${item.sala}</div>
          <div class="text-xs text-muted-foreground truncate">
            ${fmtDataHora(item.inicio)} - ${fmtDataHora(item.fim)}
          </div>
        </div>
      </div>
    `;
    left.addEventListener('click', () => AbrirModalDetalheAgendamento(item));

    // direita (ícone excluir)
    const del = document.createElement('button');
    del.type = 'button';
    del.className = 'w-9 h-9 rounded-xl border border-border bg-white/60 hover:bg-destructive hover:text-white transition-all flex items-center justify-center shrink-0';
    del.title = 'Excluir agendamento';
    del.innerHTML = `<i class="fas fa-trash"></i>`;
    del.addEventListener('click', async (ev) => {
      ev.stopPropagation();

      if (!confirm('Deseja excluir (cancelar) este agendamento?')) return;

      const usuario = sessionStorage.getItem('usuario') || '';
      if (!usuario) {
        alert('Usuário não identificado. Faça login novamente.');
        return;
      }

      const delResp = await fetch(`${API_BASE}/api/cancelar-agendamentos/sala/${item.id}`, {
        method: 'DELETE',
        headers: {
          'x-usuario': usuario
        }
      });

      const delJson = await delResp.json().catch(() => ({}));

      if (!delResp.ok) {
        alert(delJson?.message || 'Erro ao excluir agendamento.');
        return;
      }

      alert(delJson?.message || 'Agendamento cancelado com sucesso.');
      await CarregarResumoDoDia();
    });

    row.appendChild(left);
    row.appendChild(del);
    lista.appendChild(row);
  });
}




// ===============================
// GESTÃO DE USUÁRIOS - COMPLETO ATUALIZADO
// Com locais de trabalho dinâmicos (SF_LOCAL_TRABALHO)
// ===============================

function getApiBaseGestaoUsuarios() {
  let raw =
    sessionStorage.getItem('api_base') ||
    sessionStorage.getItem('apibase') ||
    '';

  raw = String(raw || '').trim();

  if (!raw) return '';

  if (!/^https?:\/\//i.test(raw)) {
    raw = `https://${raw}`;
  }

  try {
    const url = new URL(raw);
    return url.href.replace(/\/+$/, '');
  } catch {
    return '';
  }
}

function montarUrlApiGestao(path) {
  const base = getApiBaseGestaoUsuarios();
  if (!base) throw new Error('API base não configurada.');

  const p = String(path || '').trim();
  return p.startsWith('http://') || p.startsWith('https://')
    ? p
    : `${base}${p.startsWith('/') ? '' : '/'}${p}`;
}

function absUrlFromApiGestaoUsuarios(relOrAbs, apiBase) {
  const s = String(relOrAbs || '').trim();
  if (!s) return '';
  if (/^https?:\/\//i.test(s)) return s;

  const base = String(apiBase || getApiBaseGestaoUsuarios() || '').trim();
  if (!base) return s;

  try {
    if (s.startsWith('/foto-usuario/')) {
      return new URL(`/publicidade${s}`, `${base}/`).href;
    }

    if (s.startsWith('/marketing/')) {
      return new URL(`/publicidade${s}`, `${base}/`).href;
    }

    if (s.startsWith('/publicidade/')) {
      return new URL(s, `${base}/`).href;
    }

    return new URL(s, `${base}/`).href;
  } catch {
    return s;
  }
}

function escapeHtml(s) {
  return String(s ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#039;');
}

function titleCaseNome(nome) {
  return (nome || '')
    .trim()
    .toLowerCase()
    .split(/\s+/)
    .filter(Boolean)
    .map(w => w[0]?.toUpperCase() + w.slice(1))
    .join(' ');
}

function normalizarEmail(email) {
  return (email || '').trim().toLowerCase();
}

function somenteNumeros(v) {
  return (v || '').toString().replace(/\D+/g, '');
}

function formatarCelularBR(raw) {
  const n = somenteNumeros(raw);
  if (n.length === 11) return `(${n.slice(0, 2)}) ${n.slice(2, 3)}${n.slice(3, 7)}-${n.slice(7)}`;
  if (n.length === 10) return `(${n.slice(0, 2)}) ${n.slice(2, 6)}-${n.slice(6)}`;
  return n;
}

function formatarCPF(raw) {
  const n = somenteNumeros(raw).slice(0, 11);
  if (n.length <= 3) return n;
  if (n.length <= 6) return `${n.slice(0, 3)}.${n.slice(3)}`;
  if (n.length <= 9) return `${n.slice(0, 3)}.${n.slice(3, 6)}.${n.slice(6)}`;
  return `${n.slice(0, 3)}.${n.slice(3, 6)}.${n.slice(6, 9)}-${n.slice(9)}`;
}

function setGestaoUsuariosErro(msg) {
  const el = document.getElementById('gestaoUsuariosErro');
  if (!el) return;
  if (!msg) {
    el.textContent = '';
    el.classList.add('hidden');
    return;
  }
  el.textContent = msg;
  el.classList.remove('hidden');
}

function statusBadge(status) {
  const s = (status || '').trim();
  let cls = 'bg-info/15 text-info border-info/20';
  if (s === 'Ativo') cls = 'bg-success/15 text-success border-success/20';
  else if (s === 'Desativado') cls = 'bg-destructive/15 text-destructive border-destructive/20';

  return `<span class="inline-flex items-center px-3 py-1 rounded-full border ${cls} text-xs font-semibold">${escapeHtml(s || '—')}</span>`;
}

function avatarUsuarioHtml(u) {
  const fotoRel = u.FOTO ?? u.foto ?? '';
  const nome = (u.NOME ?? u.nome ?? '').toString().trim();
  const inicial = escapeHtml((nome[0] || 'U').toUpperCase());

  const fotoAbs = fotoRel ? absUrlFromApiGestaoUsuarios(fotoRel) : '';

  if (fotoAbs) {
    return `
      <div class="w-10 h-10 rounded-full overflow-hidden bg-muted shrink-0 border border-border">
        <img
          src="${escapeHtml(fotoAbs)}"
          alt="${escapeHtml(nome || 'Usuário')}"
          class="w-full h-full object-cover"
          onerror="this.style.display='none'; this.parentElement.innerHTML='<div class=&quot;w-10 h-10 rounded-full bg-primary/15 text-primary shrink-0 border border-border flex items-center justify-center font-semibold&quot;>${inicial}</div>';"
        >
      </div>
    `;
  }

  return `
    <div class="w-10 h-10 rounded-full bg-primary/15 text-primary shrink-0 border border-border flex items-center justify-center font-semibold">
      ${inicial}
    </div>
  `;
}

function rowUsuario(u) {
  const nome = u.NOME ?? u.nome ?? '';
  const emailCorporativo = u.EMAIL ?? u.email ?? '';
  const setor = u.SETOR ?? u.setor ?? '';
  const perfil = u.PERFIL ?? u.perfil ?? '';
  const localTrabalho = u.LOCAL_TRABALHO ?? u.local_trabalho ?? u.LOCALTRABALHO ?? u.localtrabalho ?? '';
  const status = u.STATUS ?? u.status ?? '';
  const id = u.ID ?? u.id ?? '';

  return `
    <tr>
      <td class="px-4 py-3">
        <div class="flex items-center gap-3">
          ${avatarUsuarioHtml(u)}
          <div class="min-w-0">
            <div class="font-medium truncate">${escapeHtml(nome)}</div>
            <div class="text-xs text-muted-foreground truncate">${escapeHtml(emailCorporativo)}</div>
          </div>
        </div>
      </td>
      <td class="px-4 py-3">${escapeHtml(setor || '—')}</td>
      <td class="px-4 py-3">${escapeHtml(perfil || '—')}</td>
      <td class="px-4 py-3">${escapeHtml(localTrabalho || '—')}</td>
      <td class="px-4 py-3">${statusBadge(status)}</td>
      <td class="px-4 py-3">
        <div class="flex justify-end gap-2">
          <button
            class="btnViewUsuario w-10 h-10 rounded-xl border border-border bg-white/60 hover:bg-white/90 transition-all"
            data-id="${escapeHtml(id)}"
            aria-label="Visualizar usuário"
            title="Visualizar usuário"
          >
            <i class="fas fa-eye" aria-hidden="true"></i>
          </button>

          <button
            class="btnEditUsuario w-10 h-10 rounded-xl border border-border bg-white/60 hover:bg-white/90 transition-all"
            data-id="${escapeHtml(id)}"
            aria-label="Editar usuário"
            title="Editar usuário"
          >
            <i class="fas fa-pen" aria-hidden="true"></i>
          </button>

          <button
            class="btnDelUsuario w-10 h-10 rounded-xl border border-border bg-white/60 hover:bg-white/90 transition-all"
            data-id="${escapeHtml(id)}"
            aria-label="Excluir usuário"
            title="Excluir usuário"
          >
            <i class="fas fa-trash" aria-hidden="true"></i>
          </button>
        </div>
      </td>
    </tr>
  `;
}


async function apiGet(path) {
  const url = montarUrlApiGestao(path);

  const r = await fetch(url);
  const txt = await r.text();

  let data = null;
  try { data = txt ? JSON.parse(txt) : null; } catch { data = null; }

  if (!r.ok) {
    const msg = data?.message || data?.error || txt || `HTTP ${r.status}`;
    throw new Error(msg);
  }

  return data;
}

async function apiSend(path, method, body) {
  const url = montarUrlApiGestao(path);

  const r = await fetch(url, {
    method,
    headers: { 'Content-Type': 'application/json' },
    body: body ? JSON.stringify(body) : undefined,
  });

  const txt = await r.text();

  let data = null;
  try { data = txt ? JSON.parse(txt) : null; } catch { data = null; }

  if (!r.ok) {
    const msg = data?.message || data?.error || txt || `HTTP ${r.status}`;
    throw new Error(msg);
  }

  return data;
}

async function apiUploadFotoUsuarioBlob(blob, fileName = 'foto-usuario.jpg') {
  const base = getApiBaseGestaoUsuarios();
  if (!base) throw new Error('API base não configurada.');

  const url = `${base}/api/gestao-usuarios/foto`;
  const fd = new FormData();
  fd.append('foto', blob, fileName);

  const r = await fetch(url, {
    method: 'POST',
    body: fd,
  });

  const txt = await r.text();

  let data = null;
  try { data = txt ? JSON.parse(txt) : null; } catch { data = null; }

  if (!r.ok) {
    const msg = data?.message || data?.error || txt || `HTTP ${r.status}`;
    throw new Error(msg);
  }

  return data;
}

function canvasToBlob(canvas, type = 'image/jpeg', quality = 0.9) {
  return new Promise((resolve, reject) => {
    canvas.toBlob((blob) => {
      if (!blob) return reject(new Error('Falha ao gerar imagem.'));
      resolve(blob);
    }, type, quality);
  });
}

function showPage(pageId) {
  document.querySelectorAll('.page-content').forEach(s => s.classList.remove('active'));
  const sec = document.getElementById(pageId);
  if (sec) sec.classList.add('active');
}

let cachePerfisGestao = [];
let cacheSetoresGestao = [];
let cacheLocaisTrabalhoGestao = [];

async function carregarGestaoUsuarios() {
  try {
    setGestaoUsuariosErro('');
    showPage('user-management');

    const tbody = document.getElementById('tbodyGestaoUsuarios');
    if (tbody) {
      tbody.innerHTML = `<tr><td colspan="6" class="px-4 py-6 text-sm text-muted-foreground">Carregando usuários...</td></tr>`;
    }

    const [usuariosResp, perfisResp, setoresResp, locaisResp] = await Promise.all([
      apiGet('/api/gestao-usuarios'),
      apiGet('/api/gestao-usuarios-perfis'),
      apiGet('/api/gestao-usuarios-setores'),
      apiGet('/api/gestao-usuarios-locais-trabalho'),
    ]);

    const usuarios = Array.isArray(usuariosResp?.items) ? usuariosResp.items : [];
    const perfis = Array.isArray(perfisResp?.items) ? perfisResp.items : [];
    const setores = Array.isArray(setoresResp?.items) ? setoresResp.items : [];
    const locais = Array.isArray(locaisResp?.items) ? locaisResp.items : [];

    cachePerfisGestao = perfis;
    cacheSetoresGestao = setores;
    cacheLocaisTrabalhoGestao = locais;

    if (!tbody) return;

    if (!usuarios.length) {
      tbody.innerHTML = `<tr><td colspan="6" class="px-4 py-6 text-sm text-muted-foreground">Nenhum usuário cadastrado.</td></tr>`;
      return;
    }

    tbody.innerHTML = usuarios.map(rowUsuario).join('');
  } catch (err) {
    setGestaoUsuariosErro(err?.message || 'Erro ao carregar usuários.');
    const tbody = document.getElementById('tbodyGestaoUsuarios');
    if (tbody) {
      tbody.innerHTML = `<tr><td colspan="6" class="px-4 py-6 text-sm text-destructive">Falha ao carregar usuários.</td></tr>`;
    }
  }
}

function removerModalGestaoUsuario() {
  document.getElementById('gestaoUsuarioOverlay')?.remove();
  document.getElementById('gestaoUsuarioModal')?.remove();
}

function optionsFromRows(rows, selectedValue, placeholder = 'Selecione...') {
  const sel = (selectedValue || '').toString().trim();

  const opts = (rows || []).map(r => {
    const nome = (r.NOME ?? r.nome ?? '').toString().trim();
    if (!nome) return '';
    const selected = nome === sel ? 'selected' : '';
    return `<option value="${escapeHtml(nome)}" ${selected}>${escapeHtml(nome)}</option>`;
  }).join('');

  return `<option value="" ${sel ? '' : 'selected'}>${escapeHtml(placeholder)}</option>` + opts;
}

function inputValue(id) {
  return document.getElementById(id)?.value ?? '';
}

function setAbaGestaoUsuario(nome) {
  const abaCorp = document.getElementById('guAbaCorporativo');
  const abaPess = document.getElementById('guAbaPessoal');
  const painelCorp = document.getElementById('guPainelCorporativo');
  const painelPess = document.getElementById('guPainelPessoal');

  const pessoalAtiva = nome === 'pessoal';

  if (abaCorp) {
    abaCorp.setAttribute('aria-selected', pessoalAtiva ? 'false' : 'true');
    abaCorp.className = pessoalAtiva
      ? 'px-4 py-2 rounded-xl text-sm font-medium border border-border bg-white/40 text-muted-foreground hover:bg-white/70 transition-all'
      : 'px-4 py-2 rounded-xl text-sm font-medium border border-border bg-white text-foreground shadow-sm transition-all';
  }

  if (abaPess) {
    abaPess.setAttribute('aria-selected', pessoalAtiva ? 'true' : 'false');
    abaPess.className = pessoalAtiva
      ? 'px-4 py-2 rounded-xl text-sm font-medium border border-border bg-white text-foreground shadow-sm transition-all'
      : 'px-4 py-2 rounded-xl text-sm font-medium border border-border bg-white/40 text-muted-foreground hover:bg-white/70 transition-all';
  }

  if (painelCorp) painelCorp.hidden = pessoalAtiva;
  if (painelPess) painelPess.hidden = !pessoalAtiva;
}

function abrirModalGestaoUsuario({ modo, usuario }) {
  removerModalGestaoUsuario();

  const overlay = document.createElement('div');
  overlay.id = 'gestaoUsuarioOverlay';
  overlay.className = 'fixed inset-0 bg-black/40 backdrop-blur-sm z-[90]';
  document.body.appendChild(overlay);

  const modal = document.createElement('div');
  modal.id = 'gestaoUsuarioModal';
  modal.className = 'fixed inset-0 z-[100]';
  modal.setAttribute('role', 'dialog');
  modal.setAttribute('aria-modal', 'true');

  const isEdit = modo === 'edit';
  const isView = modo === 'view';
  const isNew = modo === 'new';
  const u = usuario || {};

  const formDisabledAttr = isView ? 'disabled' : '';
  const inputReadonlyAttr = isView ? 'readonly' : '';
  const fileDisabledAttr = isView ? 'disabled' : '';

  const fotoAtualRel = u.FOTO ?? u.foto ?? '';
  let fotoAtualAbs = '';

  try {
    fotoAtualAbs = fotoAtualRel ? absUrlFromApiGestaoUsuarios(fotoAtualRel) : '';
  } catch (err) {
    console.error('Erro ao resolver fotoAtualAbs:', err);
    fotoAtualAbs = '';
  }

  let perfilOptions = '';
  let setorOptions = '';
  let localTrabalhoOptions = '';

  try {
    perfilOptions = optionsFromRows(cachePerfisGestao, (u.PERFIL ?? u.perfil ?? ''));
  } catch (err) {
    console.error('Erro ao gerar options de perfil:', err);
    perfilOptions = `<option value="" selected>Erro ao carregar perfis</option>`;
  }

  try {
    setorOptions = optionsFromRows(cacheSetoresGestao, (u.SETOR ?? u.setor ?? ''));
  } catch (err) {
    console.error('Erro ao gerar options de setor:', err);
    setorOptions = `<option value="" selected>Erro ao carregar setores</option>`;
  }

  try {
    localTrabalhoOptions = optionsFromRows(
      cacheLocaisTrabalhoGestao,
      (u.LOCAL_TRABALHO ?? u.local_trabalho ?? u.LOCALTRABALHO ?? u.localtrabalho ?? ''),
      'Selecione...'
    );
  } catch (err) {
    console.error('Erro ao gerar options de local de trabalho:', err);
    localTrabalhoOptions = `<option value="" selected>Erro ao carregar locais</option>`;
  }

  const emailCorporativo = u.EMAIL ?? u.email ?? '';
  const telefoneCorporativo = formatarCelularBR(u.TELEFONE ?? u.telefone ?? '');
  const statusAtual = (u.STATUS ?? u.status ?? 'Ativo').toString().trim() || 'Ativo';

  const cpfAtual = formatarCPF(u.CPF ?? u.cpf ?? '');
  const rgAtual = u.RG ?? u.rg ?? '';
  const cnhAtual = u.CNH ?? u.cnh ?? '';
  const cnhCategoriaAtual = u.CNH_CATEGORIA ?? u.cnh_categoria ?? u.CNHCATEGORIA ?? u.cnhcategoria ?? '';
  const dataNascimentoAtual = String(u.DATA_NASCIMENTO ?? u.data_nascimento ?? u.DATANASCIMENTO ?? u.datanascimento ?? '').slice(0, 10);
  const estadoCivilAtual = u.ESTADO_CIVIL ?? u.estado_civil ?? u.ESTADOCIVIL ?? u.estadocivil ?? '';
  const telefonePessoalAtual = formatarCelularBR(u.TELEFONE_PESSOAL ?? u.telefone_pessoal ?? u.TELEFONEPESSOAL ?? u.telefonepessoal ?? '');
  const emailPessoalAtual = u.EMAIL_PESSOAL ?? u.email_pessoal ?? u.EMAILPESSOAL ?? u.emailpessoal ?? '';

  modal.innerHTML = `
    <div class="w-full h-full overflow-auto">
      <div class="min-h-full flex items-start justify-center p-4 md:p-8">
        <div class="w-full max-w-6xl mx-auto">
          <div class="glass rounded-2xl shadow-2xl border border-border overflow-hidden">
            <div class="px-6 py-5 border-b border-border flex items-start justify-between gap-4">
              <div>
                <h3 class="text-xl font-semibold text-foreground">
                  ${isView ? 'Visualizar usuário' : isEdit ? 'Editar usuário' : 'Novo usuário'}
                </h3>
                <p class="text-sm text-muted-foreground">
                  ${isView ? 'Consulta de dados do usuário' : 'Dados corporativos, pessoais e segurança'}
                </p>
              </div>

              <button id="btnFecharGestaoUsuario" type="button"
                class="w-10 h-10 rounded-xl bg-white/60 border border-border hover:bg-white transition-all flex items-center justify-center"
                aria-label="Fechar" title="Fechar">
                <i class="fas fa-times" aria-hidden="true"></i>
              </button>
            </div>

            <form id="formGestaoUsuario" class="px-6 py-6" autocomplete="off">
              <input type="text" name="fakeusernameremembered" autocomplete="username" class="hidden" tabindex="-1">
              <input type="password" name="fakepasswordremembered" autocomplete="new-password" class="hidden" tabindex="-1">

              <div class="grid grid-cols-1 lg:grid-cols-[320px_minmax(0,1fr)] gap-6">
                <div class="space-y-4">
                  <div class="rounded-2xl border border-border bg-white/40 p-4 space-y-3">
                    <label class="text-sm font-medium block">Foto do usuário</label>

                    <div class="w-28 h-28 rounded-full overflow-hidden bg-muted flex items-center justify-center mx-auto border border-border">
                      <img id="guFotoPreview"
                           src="${escapeHtml(fotoAtualAbs)}"
                           alt="Foto do usuário"
                           class="w-full h-full object-cover ${fotoAtualAbs ? '' : 'hidden'}">
                      <span id="guFotoPlaceholder" class="text-xs text-muted-foreground ${fotoAtualAbs ? 'hidden' : ''}">
                        Sem foto
                      </span>
                    </div>

                    <canvas id="guFotoCanvas" class="hidden border border-dashed border-border rounded-xl w-full max-h-64 cursor-grab ${isView ? 'hidden' : ''}"></canvas>

                    ${isView ? '' : `
                    <div class="flex flex-col gap-2">
                      <input id="guFotoInput" type="file" accept="image/*" ${fileDisabledAttr}
                        class="w-full rounded-xl border border-border bg-white/70 px-4 py-2 text-sm"
                        autocomplete="off" />

                      <button id="btnRemoverFotoGU" type="button"
                        class="rounded-xl border border-border bg-white/60 px-4 py-2 text-sm hover:bg-white/90 transition-all">
                        Remover foto
                      </button>
                    </div>

                    <p class="text-xs text-muted-foreground">
                      Selecione uma imagem e use arraste/zoom para ajustar.
                    </p>
                    `}
                  </div>

                  ${isEdit && !isView ? `
                  <div class="rounded-2xl border border-border bg-white/40 p-4 space-y-3">
                    <div>
                      <h4 class="text-sm font-semibold text-foreground">Segurança</h4>
                      <p class="text-xs text-muted-foreground">Troca de senha do usuário</p>
                    </div>

                    <div class="grid grid-cols-1 gap-2">
                      <button id="btnMostrarAlteracaoSenhaGU" type="button"
                        class="rounded-xl border border-border bg-white/70 px-4 py-3 text-sm font-medium hover:bg-white transition-all">
                        Alterar com senha atual
                      </button>
                    </div>

                    <div id="guBlocoAlteracaoSenha" class="hidden space-y-3">
                      <div class="space-y-2">
                        <label class="text-sm font-medium">Senha atual</label>
                        <input id="guSenhaAtual" type="password"
                          class="w-full rounded-xl border border-border bg-white/70 px-4 py-3 outline-none focus:ring-2 focus:ring-primary/30"
                          placeholder="Informe a senha atual"
                          autocomplete="new-password"
                          data-lpignore="true">
                      </div>

                      <div class="space-y-2">
                        <label class="text-sm font-medium">Nova senha</label>
                        <input id="guNovaSenha" type="password" minlength="6"
                          class="w-full rounded-xl border border-border bg-white/70 px-4 py-3 outline-none focus:ring-2 focus:ring-primary/30"
                          placeholder="Mínimo 6 caracteres"
                          autocomplete="new-password"
                          data-lpignore="true">
                      </div>

                      <div class="grid grid-cols-1 gap-2">
                        <button id="btnTrocarSenhaGU" type="button"
                          class="rounded-xl border border-border bg-white/70 px-4 py-3 text-sm font-medium hover:bg-white transition-all">
                          Salvar nova senha
                        </button>
                      </div>
                    </div>

                    <p id="guSenhaMsg" class="text-xs hidden whitespace-pre-line"></p>
                  </div>
                  ` : ''}

                  ${isNew ? `
                  <div class="rounded-2xl border border-border bg-white/40 p-4 space-y-2">
                    <label class="text-sm font-medium">Senha inicial</label>
                    <input id="guSenha" type="password" minlength="6" required
                      class="w-full rounded-xl border border-border bg-white/70 px-4 py-3 outline-none focus:ring-2 focus:ring-primary/30"
                      placeholder="Mínimo 6 caracteres"
                      autocomplete="new-password"
                      data-lpignore="true" />
                    <p class="text-xs text-muted-foreground">Será gravada com hash no banco.</p>
                  </div>
                  ` : ''}
                </div>

                <div class="space-y-4 min-w-0">
                  <div class="flex flex-wrap gap-2">
                    <button id="guAbaCorporativo" type="button" aria-selected="true"
                      class="px-4 py-2 rounded-xl text-sm font-medium border border-border bg-white text-foreground shadow-sm transition-all">
                      Corporativo
                    </button>

                    <button id="guAbaPessoal" type="button" aria-selected="false"
                      class="px-4 py-2 rounded-xl text-sm font-medium border border-border bg-white/40 text-muted-foreground hover:bg-white/70 transition-all">
                      Pessoal
                    </button>
                  </div>

                  <div id="guPainelCorporativo" class="space-y-4">
                    <div class="rounded-2xl border border-border bg-white/40 p-4">
                      <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
                        <div class="space-y-2 md:col-span-2">
                          <label class="text-sm font-medium">Nome</label>
                          <input id="guNome" type="text" required ${inputReadonlyAttr}
                            class="w-full rounded-xl border border-border bg-white/70 px-4 py-3 outline-none focus:ring-2 focus:ring-primary/30"
                            value="${escapeHtml(u.NOME ?? u.nome ?? '')}"
                            autocomplete="off" />
                        </div>

                        <div class="space-y-2">
                          <label class="text-sm font-medium">E-mail corporativo</label>
                          <input id="guEmailCorporativo" type="email" required ${inputReadonlyAttr}
                            class="w-full rounded-xl border border-border bg-white/70 px-4 py-3 outline-none focus:ring-2 focus:ring-primary/30"
                            value="${escapeHtml(emailCorporativo)}"
                            autocomplete="off"
                            autocapitalize="off"
                            autocorrect="off"
                            spellcheck="false"
                            data-lpignore="true" />
                        </div>

                        <div class="space-y-2">
                          <label class="text-sm font-medium">Telefone corporativo</label>
                          <input id="guTelefoneCorporativo" type="text" ${inputReadonlyAttr}
                            class="w-full rounded-xl border border-border bg-white/70 px-4 py-3 outline-none focus:ring-2 focus:ring-primary/30"
                            value="${escapeHtml(telefoneCorporativo)}"
                            placeholder="(77) 9XXXX-XXXX"
                            autocomplete="off" />
                        </div>

                        <div class="space-y-2">
                          <label class="text-sm font-medium">Perfil</label>
                          <select id="guPerfil" required ${formDisabledAttr}
                            class="w-full rounded-xl border border-border bg-white/70 px-4 py-3 outline-none focus:ring-2 focus:ring-primary/30">
                            ${perfilOptions}
                          </select>
                        </div>

                        <div class="space-y-2">
                          <label class="text-sm font-medium">Setor</label>
                          <div class="flex gap-2">
                            <select id="guSetor" required ${formDisabledAttr}
                              class="flex-1 rounded-xl border border-border bg-white/70 px-4 py-3 outline-none focus:ring-2 focus:ring-primary/30">
                              ${setorOptions}
                            </select>

                            ${isView ? '' : `
                            <button id="btnAddSetor" type="button"
                              class="w-12 h-12 rounded-xl border border-border bg-white/60 hover:bg-white/90 transition-all flex items-center justify-center"
                              aria-label="Adicionar setor" title="Adicionar setor">
                              <i class="fas fa-plus" aria-hidden="true"></i>
                            </button>
                            `}
                          </div>
                        </div>

                        <div class="space-y-2">
                          <label class="text-sm font-medium">Local de trabalho</label>
                          <div class="flex gap-2">
                            <select id="guLocalTrabalho" ${formDisabledAttr}
                              class="flex-1 rounded-xl border border-border bg-white/70 px-4 py-3 outline-none focus:ring-2 focus:ring-primary/30">
                              ${localTrabalhoOptions}
                            </select>

                            ${isView ? '' : `
                            <button id="btnAddLocalTrabalho" type="button"
                              class="w-12 h-12 rounded-xl border border-border bg-white/60 hover:bg-white/90 transition-all flex items-center justify-center"
                              aria-label="Adicionar local de trabalho" title="Adicionar local de trabalho">
                              <i class="fas fa-plus" aria-hidden="true"></i>
                            </button>
                            `}
                          </div>
                        </div>

                        <div class="space-y-2">
                          <label class="text-sm font-medium">Status</label>
                          <select id="guStatus" required ${formDisabledAttr}
                            class="w-full rounded-xl border border-border bg-white/70 px-4 py-3 outline-none focus:ring-2 focus:ring-primary/30">
                            ${['Ativo', 'Desativado'].map(s => {
                              const selected = statusAtual === s ? 'selected' : '';
                              return `<option value="${escapeHtml(s)}" ${selected}>${escapeHtml(s)}</option>`;
                            }).join('')}
                          </select>
                        </div>
                      </div>
                    </div>
                  </div>

                  <div id="guPainelPessoal" class="space-y-4" hidden>
                    <div class="rounded-2xl border border-border bg-white/40 p-4">
                      <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
                        <div class="space-y-2">
                          <label class="text-sm font-medium">CPF</label>
                          <input id="guCpf" type="text" ${inputReadonlyAttr}
                            class="w-full rounded-xl border border-border bg-white/70 px-4 py-3 outline-none focus:ring-2 focus:ring-primary/30"
                            value="${escapeHtml(cpfAtual)}"
                            placeholder="000.000.000-00"
                            autocomplete="off" />
                        </div>

                        <div class="space-y-2">
                          <label class="text-sm font-medium">RG</label>
                          <input id="guRg" type="text" ${inputReadonlyAttr}
                            class="w-full rounded-xl border border-border bg-white/70 px-4 py-3 outline-none focus:ring-2 focus:ring-primary/30"
                            value="${escapeHtml(rgAtual)}"
                            autocomplete="off" />
                        </div>

                        <div class="space-y-2">
                          <label class="text-sm font-medium">CNH</label>
                          <input id="guCnh" type="text" ${inputReadonlyAttr}
                            class="w-full rounded-xl border border-border bg-white/70 px-4 py-3 outline-none focus:ring-2 focus:ring-primary/30"
                            value="${escapeHtml(cnhAtual)}"
                            autocomplete="off" />
                        </div>

                        <div class="space-y-2">
                          <label class="text-sm font-medium">Categoria CNH</label>
                          <input id="guCnhCategoria" type="text" maxlength="5" ${inputReadonlyAttr}
                            class="w-full rounded-xl border border-border bg-white/70 px-4 py-3 outline-none focus:ring-2 focus:ring-primary/30 uppercase"
                            value="${escapeHtml(cnhCategoriaAtual)}"
                            placeholder="A, B, AB..."
                            autocomplete="off" />
                        </div>

                        <div class="space-y-2">
                          <label class="text-sm font-medium">Data de nascimento</label>
                          <input id="guDataNascimento" type="date" ${formDisabledAttr}
                            class="w-full rounded-xl border border-border bg-white/70 px-4 py-3 outline-none focus:ring-2 focus:ring-primary/30"
                            value="${escapeHtml(dataNascimentoAtual)}" />
                        </div>

                        <div class="space-y-2">
                          <label class="text-sm font-medium">Estado civil</label>
                          <select id="guEstadoCivil" ${formDisabledAttr}
                            class="w-full rounded-xl border border-border bg-white/70 px-4 py-3 outline-none focus:ring-2 focus:ring-primary/30">
                            ${[
                              '',
                              'Solteiro(a)',
                              'Casado(a)',
                              'Divorciado(a)',
                              'Viúvo(a)',
                              'União estável'
                            ].map(s => {
                              const selected = (estadoCivilAtual || '') === s ? 'selected' : '';
                              const label = s || 'Selecione...';
                              return `<option value="${escapeHtml(s)}" ${selected}>${escapeHtml(label)}</option>`;
                            }).join('')}
                          </select>
                        </div>

                        <div class="space-y-2">
                          <label class="text-sm font-medium">Telefone pessoal</label>
                          <input id="guTelefonePessoal" type="text" ${inputReadonlyAttr}
                            class="w-full rounded-xl border border-border bg-white/70 px-4 py-3 outline-none focus:ring-2 focus:ring-primary/30"
                            value="${escapeHtml(telefonePessoalAtual)}"
                            placeholder="(77) 9XXXX-XXXX"
                            autocomplete="off" />
                        </div>

                        <div class="space-y-2">
                          <label class="text-sm font-medium">E-mail pessoal</label>
                          <input id="guEmailPessoal" type="email" ${inputReadonlyAttr}
                            class="w-full rounded-xl border border-border bg-white/70 px-4 py-3 outline-none focus:ring-2 focus:ring-primary/30"
                            value="${escapeHtml(emailPessoalAtual)}"
                            autocomplete="off"
                            autocapitalize="off"
                            autocorrect="off"
                            spellcheck="false"
                            data-lpignore="true" />
                        </div>
                      </div>
                    </div>
                  </div>

                  <p id="guErro" class="text-sm text-destructive hidden whitespace-pre-line"></p>

                  <div class="pt-2 flex flex-col sm:flex-row gap-3">
                    ${isView ? `
                    <button id="btnCancelarGU" type="button"
                      class="sm:flex-1 rounded-xl border border-border bg-white/50 px-4 py-3 font-medium hover:bg-white/70 transition-all">
                      Fechar
                    </button>
                    ` : `
                    <button id="btnSalvarGU" type="submit"
                      class="sm:flex-1 rounded-xl bg-primary text-white px-4 py-3 font-medium hover:opacity-90 transition-all">
                      Salvar
                    </button>

                    <button id="btnCancelarGU" type="button"
                      class="sm:flex-1 rounded-xl border border-border bg-white/50 px-4 py-3 font-medium hover:bg-white/70 transition-all">
                      Cancelar
                    </button>
                    `}
                  </div>
                </div>
              </div>
            </form>
          </div>
        </div>
      </div>
    </div>
  `;

  document.body.appendChild(modal);

  let guFotoImage = null;
  let guFotoScale = 1;
  let guFotoOffsetX = 0;
  let guFotoOffsetY = 0;
  let guFotoIsDragging = false;
  let guFotoStartX = 0;
  let guFotoStartY = 0;
  let guFotoRemovida = false;

  const fileInput = document.getElementById('guFotoInput');
  const canvas = document.getElementById('guFotoCanvas');
  const previewImg = document.getElementById('guFotoPreview');
  const placeholder = document.getElementById('guFotoPlaceholder');
  const btnRemoverFoto = document.getElementById('btnRemoverFotoGU');
  const form = document.getElementById('formGestaoUsuario');
  const btnFechar = document.getElementById('btnFecharGestaoUsuario');
  const btnCancelar = document.getElementById('btnCancelarGU');
  const btnAddSetor = document.getElementById('btnAddSetor');
  const btnAddLocalTrabalho = document.getElementById('btnAddLocalTrabalho');
  const btnMostrarAlteracaoSenha = document.getElementById('btnMostrarAlteracaoSenhaGU');
  const blocoAlteracaoSenha = document.getElementById('guBlocoAlteracaoSenha');

  const inpTelCorp = document.getElementById('guTelefoneCorporativo');
  const inpTelPess = document.getElementById('guTelefonePessoal');
  const inpCpf = document.getElementById('guCpf');
  const btnAbaCorp = document.getElementById('guAbaCorporativo');
  const btnAbaPess = document.getElementById('guAbaPessoal');
  const btnTrocarSenha = document.getElementById('btnTrocarSenhaGU');
  const ctx = canvas?.getContext('2d');

  function setErr(msg) {
    const el = document.getElementById('guErro');
    if (!el) return;
    if (!msg) {
      el.textContent = '';
      el.classList.add('hidden');
      return;
    }
    el.textContent = msg;
    el.classList.remove('hidden');
  }

  function setSenhaMsg(msg, tipo = 'ok') {
    const el = document.getElementById('guSenhaMsg');
    if (!el) return;
    if (!msg) {
      el.textContent = '';
      el.className = 'text-xs hidden whitespace-pre-line';
      return;
    }
    el.textContent = msg;
    el.className = `text-xs whitespace-pre-line ${tipo === 'erro' ? 'text-destructive' : 'text-success'}`;
  }

  function fechar() {
    removerModalGestaoUsuario();
  }

  overlay.addEventListener('click', fechar);
  btnFechar?.addEventListener('click', fechar);
  btnCancelar?.addEventListener('click', fechar);

  btnAbaCorp?.addEventListener('click', () => setAbaGestaoUsuario('corporativo'));
  btnAbaPess?.addEventListener('click', () => setAbaGestaoUsuario('pessoal'));
  setAbaGestaoUsuario('corporativo');

  if (isView) return;

  if (!canvas || !ctx) throw new Error('Canvas da foto não disponível.');

  const AVATAR_SIZE = 256;
  canvas.width = AVATAR_SIZE;
  canvas.height = AVATAR_SIZE;

  function drawGuFoto() {
    if (!guFotoImage) return;

    ctx.clearRect(0, 0, AVATAR_SIZE, AVATAR_SIZE);
    ctx.save();
    ctx.beginPath();
    ctx.arc(AVATAR_SIZE / 2, AVATAR_SIZE / 2, AVATAR_SIZE / 2, 0, Math.PI * 2);
    ctx.closePath();
    ctx.clip();
    ctx.drawImage(
      guFotoImage,
      guFotoOffsetX,
      guFotoOffsetY,
      guFotoImage.width * guFotoScale,
      guFotoImage.height * guFotoScale
    );
    ctx.restore();

    previewImg.src = canvas.toDataURL('image/jpeg', 0.9);
    previewImg.classList.remove('hidden');
    placeholder.classList.add('hidden');
  }

  function resetFotoUI() {
    guFotoImage = null;
    guFotoScale = 1;
    guFotoOffsetX = 0;
    guFotoOffsetY = 0;
    guFotoIsDragging = false;

    if (fileInput) fileInput.value = '';
    canvas.classList.add('hidden');
    ctx.clearRect(0, 0, AVATAR_SIZE, AVATAR_SIZE);
    canvas.style.cursor = 'default';

    if (fotoAtualAbs && !guFotoRemovida) {
      previewImg.src = fotoAtualAbs;
      previewImg.classList.remove('hidden');
      placeholder.classList.add('hidden');
    } else {
      previewImg.removeAttribute('src');
      previewImg.classList.add('hidden');
      placeholder.classList.remove('hidden');
    }
  }

  inpTelCorp?.addEventListener('input', () => {
    inpTelCorp.value = formatarCelularBR(inpTelCorp.value);
  });

  inpTelPess?.addEventListener('input', () => {
    inpTelPess.value = formatarCelularBR(inpTelPess.value);
  });

  inpCpf?.addEventListener('input', () => {
    inpCpf.value = formatarCPF(inpCpf.value);
  });

  btnMostrarAlteracaoSenha?.addEventListener('click', () => {
    blocoAlteracaoSenha?.classList.toggle('hidden');
  });

  fileInput?.addEventListener('change', e => {
    const file = e.target.files?.[0];
    if (!file) return;

    guFotoRemovida = false;

    const reader = new FileReader();
    reader.onload = () => {
      guFotoImage = new Image();
      guFotoImage.onload = () => {
        guFotoScale = Math.max(
          AVATAR_SIZE / guFotoImage.width,
          AVATAR_SIZE / guFotoImage.height
        );
        guFotoOffsetX = (AVATAR_SIZE - guFotoImage.width * guFotoScale) / 2;
        guFotoOffsetY = (AVATAR_SIZE - guFotoImage.height * guFotoScale) / 2;
        canvas.classList.remove('hidden');
        placeholder.classList.add('hidden');
        previewImg.classList.remove('hidden');
        drawGuFoto();
        canvas.style.cursor = 'grab';
      };
      guFotoImage.onerror = err => console.error('Erro ao carregar imagem:', err);
      guFotoImage.src = reader.result;
    };
    reader.onerror = err => console.error('Erro no FileReader:', err);
    reader.readAsDataURL(file);
  });

  btnRemoverFoto?.addEventListener('click', () => {
    guFotoRemovida = true;
    resetFotoUI();
  });

  canvas.addEventListener('mousedown', e => {
    if (!guFotoImage) return;
    guFotoIsDragging = true;
    guFotoStartX = e.clientX;
    guFotoStartY = e.clientY;
    canvas.style.cursor = 'grabbing';
  });

  window.addEventListener('mousemove', e => {
    if (!guFotoIsDragging || !guFotoImage) return;
    const dx = e.clientX - guFotoStartX;
    const dy = e.clientY - guFotoStartY;
    guFotoStartX = e.clientX;
    guFotoStartY = e.clientY;
    guFotoOffsetX += dx;
    guFotoOffsetY += dy;
    drawGuFoto();
  });

  window.addEventListener('mouseup', () => {
    guFotoIsDragging = false;
    if (guFotoImage) canvas.style.cursor = 'grab';
  });

  canvas.addEventListener('wheel', e => {
    if (!guFotoImage) return;
    e.preventDefault();
    const delta = e.deltaY < 0 ? 0.1 : -0.1;
    guFotoScale = Math.max(0.3, Math.min(4, guFotoScale + delta));
    drawGuFoto();
  }, { passive: false });

  btnAddSetor?.addEventListener('click', async () => {
    const nome = prompt('Nome do setor:');
    if (!nome) return;

    try {
      await apiSend('/api/gestao-usuarios-setores', 'POST', { nome });
      const setoresResp = await apiGet('/api/gestao-usuarios-setores');
      cacheSetoresGestao = Array.isArray(setoresResp?.items) ? setoresResp.items : [];
      const sel = document.getElementById('guSetor');
      if (sel) sel.innerHTML = optionsFromRows(cacheSetoresGestao, titleCaseNome(nome));
    } catch (err) {
      alert('Erro ao adicionar setor: ' + (err?.message || err));
    }
  });

  btnAddLocalTrabalho?.addEventListener('click', async () => {
    const nome = prompt('Nome do local de trabalho:');
    if (!nome) return;

    try {
      await apiSend('/api/gestao-usuarios-locais-trabalho', 'POST', { nome });
      const locaisResp = await apiGet('/api/gestao-usuarios-locais-trabalho');
      cacheLocaisTrabalhoGestao = Array.isArray(locaisResp?.items) ? locaisResp.items : [];
      const sel = document.getElementById('guLocalTrabalho');
      if (sel) sel.innerHTML = optionsFromRows(cacheLocaisTrabalhoGestao, titleCaseNome(nome));
    } catch (err) {
      alert('Erro ao adicionar local de trabalho: ' + (err?.message || err));
    }
  });

  btnTrocarSenha?.addEventListener('click', async () => {
    setSenhaMsg('');
    const userId = (u.ID ?? u.id);
    const senhaAtual = inputValue('guSenhaAtual').toString();
    const novaSenha = inputValue('guNovaSenha').toString();

    if (!senhaAtual) {
      setSenhaMsg('Informe a senha atual.', 'erro');
      return;
    }

    if (!novaSenha || novaSenha.length < 6) {
      setSenhaMsg('A nova senha deve ter no mínimo 6 caracteres.', 'erro');
      return;
    }

    try {
      btnTrocarSenha.disabled = true;
      await apiSend(`/api/gestao-usuarios/${userId}/senha`, 'PATCH', { senhaAtual, novaSenha });
      setSenhaMsg('Senha alterada com sucesso.', 'ok');
      const a = document.getElementById('guSenhaAtual');
      const b = document.getElementById('guNovaSenha');
      if (a) a.value = '';
      if (b) b.value = '';
      if (blocoAlteracaoSenha) blocoAlteracaoSenha.classList.add('hidden');
    } catch (err) {
      setSenhaMsg(err?.message || 'Erro ao alterar senha.', 'erro');
    } finally {
      btnTrocarSenha.disabled = false;
    }
  });

  form?.addEventListener('submit', async (e) => {
    e.preventDefault();
    setErr('');

    const nome = titleCaseNome(inputValue('guNome'));
    const email = normalizarEmail(inputValue('guEmailCorporativo'));
    const telefone = somenteNumeros(inputValue('guTelefoneCorporativo'));
    const perfil = inputValue('guPerfil').trim();
    const setor = inputValue('guSetor').trim();
    const local_trabalho = inputValue('guLocalTrabalho').trim();
    const status = inputValue('guStatus').trim();

    const cpf = somenteNumeros(inputValue('guCpf'));
    const rg = inputValue('guRg').trim();
    const cnh = inputValue('guCnh').trim();
    const cnh_categoria = inputValue('guCnhCategoria').trim().toUpperCase();
    const data_nascimento = inputValue('guDataNascimento').trim();
    const estado_civil = inputValue('guEstadoCivil').trim();
    const telefone_pessoal = somenteNumeros(inputValue('guTelefonePessoal'));
    const email_pessoal = normalizarEmail(inputValue('guEmailPessoal'));

    if (!nome || !email || !perfil || !setor || !status) {
      setErr('Preencha todos os campos obrigatórios da aba corporativa.');
      return;
    }

    const btn = document.getElementById('btnSalvarGU');

    try {
      if (btn) {
        btn.disabled = true;
        btn.classList.add('opacity-70');
      }

      let fotoPayload = undefined;

      if (guFotoRemovida && fotoAtualRel) {
        fotoPayload = null;
      } else if (guFotoImage) {
        const blob = await canvasToBlob(canvas, 'image/jpeg', 0.9);
        const up = await apiUploadFotoUsuarioBlob(blob, 'foto-usuario.jpg');
        fotoPayload = up?.item?.url || null;
      }

      const payload = {
        nome,
        email,
        telefone,
        perfil,
        setor,
        local_trabalho: local_trabalho || '',
        status,
        cpf: cpf || '',
        rg: rg || '',
        cnh: cnh || '',
        cnh_categoria: cnh_categoria || '',
        data_nascimento: data_nascimento || '',
        estado_civil: estado_civil || '',
        telefone_pessoal: telefone_pessoal || '',
        email_pessoal: email_pessoal || '',
      };

      if (fotoPayload !== undefined) payload.foto = fotoPayload;

      if (isNew) {
        const senha = inputValue('guSenha').toString();
        if (!senha || senha.length < 6) {
          setErr('Senha inválida (mínimo 6).');
          return;
        }
        payload.senha = senha;
        await apiSend('/api/gestao-usuarios-adicionar', 'POST', payload);
      } else if (isEdit) {
        const userId = u.ID ?? u.id;
        await apiSend(`/api/gestao-usuarios/${userId}`, 'PUT', payload);
      }

      removerModalGestaoUsuario();
      await carregarGestaoUsuarios();
    } catch (err) {
      console.error('Erro ao salvar usuário:', err);
      setErr(err?.message || 'Erro ao salvar.');
    } finally {
      if (btn) {
        btn.disabled = false;
        btn.classList.remove('opacity-70');
      }
    }
  });
}


// ===== LIGAÇÕES =====

document.addEventListener('click', (e) => {
  const item = e.target.closest('.menu-item[data-page]');
  if (!item) return;

  const page = item.dataset.page;
  if (page !== 'user-management') return;

  carregarGestaoUsuarios();
});

document.addEventListener('click', (e) => {
  const btn = e.target.closest('#btnNovoUsuario');
  if (!btn) return;

  abrirModalGestaoUsuario({ modo: 'new' });
});

document.addEventListener('click', async (e) => {
  const btnView = e.target.closest('.btnViewUsuario');
  const btnEdit = e.target.closest('.btnEditUsuario');
  const btnDel = e.target.closest('.btnDelUsuario');

  if (btnView) {
    const id = btnView.dataset.id;

    try {
      const resp = await apiGet(`/api/gestao-usuarios/${id}`);
      const dados = resp?.item || resp;
      abrirModalGestaoUsuario({ modo: 'view', usuario: dados });
    } catch (err) {
      console.error('Erro ao abrir visualização:', err);
      alert('Erro ao abrir visualização: ' + (err?.message || err));
    }

    return;
  }

  if (btnEdit) {
    const id = btnEdit.dataset.id;

    try {
      const resp = await apiGet(`/api/gestao-usuarios/${id}`);
      const dados = resp?.item || resp;
      abrirModalGestaoUsuario({ modo: 'edit', usuario: dados });
    } catch (err) {
      console.error('Erro ao abrir edição:', err);
      alert('Erro ao abrir edição: ' + (err?.message || err));
    }

    return;
  }

  if (btnDel) {
    const id = btnDel.dataset.id;

    if (!confirm('Confirma excluir este usuário?')) return;

    try {
      await apiSend(`/api/gestao-usuarios/${id}`, 'DELETE');
      await carregarGestaoUsuarios();
    } catch (err) {
      console.error('Erro ao excluir:', err);
      alert('Erro ao excluir: ' + (err?.message || err));
    }
  }
});



// ===== Gestão de Pedido =====//

function ativarAbaPedidos(nomeAba) {
  const abaDash = document.getElementById('abaPedidosDashboard');
  const abaTab = document.getElementById('abaPedidosTabela');
  const painelDash = document.getElementById('painelPedidosDashboard');
  const painelTab = document.getElementById('painelPedidosTabela');
  const btnNovo = document.getElementById('btnNovoPedido');

  const isTabela = nomeAba === 'tabela';

  // aria-selected
  abaDash?.setAttribute('aria-selected', isTabela ? 'false' : 'true');
  abaTab?.setAttribute('aria-selected', isTabela ? 'true' : 'false');

  // classes (estilo ativo/inativo)
  if (abaDash && abaTab) {
    if (!isTabela) {
      abaDash.className = "px-4 py-2 rounded-lg text-sm font-medium transition-all bg-white shadow text-foreground";
      abaTab.className  = "px-4 py-2 rounded-lg text-sm font-medium transition-all text-muted-foreground hover:text-foreground";
    } else {
      abaTab.className  = "px-4 py-2 rounded-lg text-sm font-medium transition-all bg-white shadow text-foreground";
      abaDash.className = "px-4 py-2 rounded-lg text-sm font-medium transition-all text-muted-foreground hover:text-foreground";
    }
  }

  // painéis
  if (painelDash) painelDash.hidden = isTabela;
  if (painelTab)  painelTab.hidden  = !isTabela;

  // botão novo pedido (só na tabela)
  if (btnNovo) btnNovo.classList.toggle('hidden', !isTabela);
}

document.getElementById('abaPedidosDashboard')?.addEventListener('click', () => ativarAbaPedidos('dashboard'));
document.getElementById('abaPedidosTabela')?.addEventListener('click', () => ativarAbaPedidos('tabela'));

// estado inicial
ativarAbaPedidos('dashboard');

function mostrarMsgMarketing(msg) {
  const el = document.getElementById('marketingMsg');
  if (!el) return;
  if (!msg) { el.classList.add('hidden'); el.textContent = ''; return; }
  el.textContent = msg;
  el.classList.remove('hidden');
}

function escapeHtml(s) {
  return String(s ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#039;');
}



// Converte "/publicidade/..." em "https://seuapp.up.railway.app/publicidade/..."
function absUrlFromApi(relOrAbs, apiBase) {
  const s = String(relOrAbs || '').trim();
  if (!s) return '';

  // Se já for absoluta (http/https), mantém
  if (/^https?:\/\//i.test(s)) return s;

  // Resolve relativo usando URL API (padrão correto)
  return new URL(s, apiBase + '/').href;
}

async function listarImagensMarketing() {
  const APIBASE = getApiBase();

  mostrarMsgMarketing('Carregando imagens...');
  const r = await fetch(`${APIBASE}/api/marketing/imagens`, { method: 'GET' });

  const data = await r.json().catch(() => ({}));
  if (!r.ok) throw new Error(data?.message || 'Erro ao listar imagens.');

  const itens = Array.isArray(data.items) ? data.items : [];

  const grid = document.getElementById('gridMarketing');
  const vazio = document.getElementById('marketingVazio');
  if (!grid || !vazio) return;

  grid.innerHTML = '';
  if (!itens.length) {
    vazio.classList.remove('hidden');
    mostrarMsgMarketing('');
    return;
  }
  vazio.classList.add('hidden');

  grid.innerHTML = itens.map((it) => {
    // it: { name, url } (url vem relativa da API)
    const urlAbs = absUrlFromApi(it.url, APIBASE);

    return `
      <div class="rounded-2xl border border-border bg-white/60 overflow-hidden shadow-sm">
        <button type="button"
          class="w-full aspect-[4/3] bg-muted/30 overflow-hidden"
          title="Visualizar"
          onclick="abrirImagemMarketing('${escapeHtml(urlAbs)}')">
          <img src="${escapeHtml(urlAbs)}" alt="${escapeHtml(it.name)}" class="w-full h-full object-cover" />
        </button>

        <div class="p-3 flex items-center justify-between gap-2">
          <div class="min-w-0">
            <div class="text-sm font-semibold truncate">${escapeHtml(it.name)}</div>
          </div>

          <button type="button"
            class="w-10 h-10 rounded-xl border border-border bg-white/60 hover:bg-destructive hover:text-white transition-all
                   flex items-center justify-center shrink-0"
            title="Remover"
            onclick="removerImagemMarketing('${escapeHtml(it.name)}')">
            <i class="fas fa-trash"></i>
          </button>
        </div>
      </div>
    `;
  }).join('');

  mostrarMsgMarketing('');
}

// visualização (nova aba) — agora recebe URL absoluta
function abrirImagemMarketing(urlAbs) {
  if (!urlAbs) return;
  window.open(urlAbs, '_blank', 'noopener,noreferrer');
}

async function enviarImagensMarketing(files) {
  const APIBASE = getApiBase();
  if (!files || !files.length) return;

  const fd = new FormData();
  for (const f of files) fd.append('files', f);

  mostrarMsgMarketing('Enviando...');
  const r = await fetch(`${APIBASE}/api/marketing/imagens`, { method: 'POST', body: fd });
  const data = await r.json().catch(() => ({}));
  if (!r.ok) throw new Error(data?.message || 'Erro ao enviar imagens.');

  mostrarMsgMarketing('Enviado com sucesso.');
  await listarImagensMarketing();
}

async function removerImagemMarketing(nomeArquivo) {
  const APIBASE = getApiBase();
  if (!nomeArquivo) return;

  if (!confirm(`Remover a imagem "${nomeArquivo}"?`)) return;

  mostrarMsgMarketing('Removendo...');
  const r = await fetch(`${APIBASE}/api/marketing/imagens/${encodeURIComponent(nomeArquivo)}`, { method: 'DELETE' });
  const data = await r.json().catch(() => ({}));
  if (!r.ok) throw new Error(data?.message || 'Erro ao remover imagem.');

  await listarImagensMarketing();
  mostrarMsgMarketing('');
}

document.getElementById('btnAtualizarMarketing')?.addEventListener('click', () => {
  listarImagensMarketing().catch(e => alert(e.message || e));
});

document.getElementById('inputImagensMarketing')?.addEventListener('change', async (e) => {
  try {
    const files = e.target.files ? Array.from(e.target.files) : [];
    await enviarImagensMarketing(files);
  } catch (err) {
    alert(err?.message || err);
  } finally {
    e.target.value = '';
  }
});



function absUrlFromApi(relOrAbs, apiBase) {
  const s = String(relOrAbs || '').trim();
  if (!s) return '';
  if (/^https?:\/\//i.test(s)) return s;
  return new URL(s, apiBase + '/').href; // resolve relativo corretamente [web:694]
}

function preloadImage(url) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve(url);
    img.onerror = reject;
    img.src = url;
  });
}

let marketingLoopTimer = null;
let marketingLoopIndex = 0;
let marketingLoopUrls = [];
let marketingLoopLastFetchMs = 0;

async function fetchMarketingUrls() {
  const APIBASE = getApiBase();
  const r = await fetch(`${APIBASE}/api/marketing/imagens`, { method: 'GET' });
  const data = await r.json().catch(() => ({}));
  if (!r.ok) throw new Error(data?.message || 'Erro ao listar imagens.');

  const itens = Array.isArray(data.items) ? data.items : [];
  return itens
    .map(it => absUrlFromApi(it.url, APIBASE))
    .filter(Boolean);
}

async function iniciarLoopMarketingPainel({
  imgId = 'painelMarketingImg',
  fallbackSrc = 'imagens/PaginaPrincipal.jpg',
  intervalMs = 20000,
  refreshListEveryMs = 60000,
} = {}) {
  const el = document.getElementById(imgId);
  if (!el) return;

  // Estado: URL que está aparecendo AGORA (para clique)
  el.dataset.currentUrl = el.src || fallbackSrc;

  // Clique: abre a imagem atual em nova aba
  el.addEventListener('click', () => {
    const url = el.dataset.currentUrl || el.src;
    if (!url) return;
    window.open(url, '_blank', 'noopener,noreferrer');
  });

  async function ensureListFresh() {
    const now = Date.now();
    if (now - marketingLoopLastFetchMs < refreshListEveryMs && marketingLoopUrls.length) return;

    marketingLoopLastFetchMs = now;
    try {
      marketingLoopUrls = await fetchMarketingUrls();
      marketingLoopIndex = 0;
    } catch {
      marketingLoopUrls = [];
    }
  }

  async function tick() {
    await ensureListFresh();

    // Sem imagens => fica na fixa
    if (!marketingLoopUrls.length) {
      el.style.opacity = '1';
      el.src = fallbackSrc;
      el.dataset.currentUrl = fallbackSrc; // clique abre a fixa
      return;
    }

    const url = marketingLoopUrls[marketingLoopIndex % marketingLoopUrls.length];
    marketingLoopIndex++;

    try {
      await preloadImage(url);

      // fade suave sem mexer no tamanho do quadro
      el.style.opacity = '0';
      setTimeout(() => {
        el.src = url;
        el.dataset.currentUrl = url; // clique abre a atual
        el.style.opacity = '1';
      }, 150);
    } catch {
      // se falhar, tenta próxima na próxima rodada
    }
  }

  await tick();

  if (marketingLoopTimer) clearInterval(marketingLoopTimer);
  marketingLoopTimer = setInterval(tick, intervalMs); // troca a cada 10s [web:705]
}


// =====================
// CLIENTES + FILIAIS (MySQL API)
// =====================
let clientesCache = [];
let filiaisForm = [];

function setMsgClientes(msg) {
  const el = document.getElementById('clientesMsg');
  if (!el) return;
  if (!msg) { el.classList.add('hidden'); el.textContent = ''; return; }
  el.textContent = msg;
  el.classList.remove('hidden');
}

function escapeHtml(s) {
  return String(s ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#039;');
}

function somenteNumeros(v) {
  return (v ?? '').toString().replace(/\D+/g, '');
}

function getFiltroClientes() {
  const qRaw = (document.getElementById('inputBuscaClientes')?.value || '').trim();
  const q = qRaw.toLowerCase();
  if (!q) return clientesCache;

  const qNum = somenteNumeros(qRaw);

  return clientesCache.filter(c => {
    const razao = String(c.RAZAOSOCIAL ?? c.RAZAO_SOCIAL ?? '').toLowerCase();
    const doc = String(c.DOCUMENTO ?? '').toLowerCase();

    return razao.includes(q) || (qNum && doc.includes(qNum));
  });
}

async function carregarClientes() {
  try {
    setMsgClientes('Carregando clientes...');
    const q = (document.getElementById('inputBuscaClientes')?.value || '').trim();
    const data = await apiGet(`/api/clientes?q=${encodeURIComponent(q)}`);
    clientesCache = Array.isArray(data?.items) ? data.items : [];
    renderTabelaClientes();
    setMsgClientes('');
  } catch (err) {
    console.error(err);
    setMsgClientes(err?.message || 'Erro ao carregar clientes.');
  }
}

function somenteDigitos(v) {
  return String(v ?? '').replace(/\D/g, '');
}

function formatarTelefoneBR(v) {
  const d = somenteDigitos(v);

  // Se vier com DDI 55 na frente, remove (bem comum em cadastros)
  const num = d.startsWith('55') && d.length > 11 ? d.slice(2) : d;

  // Esperado: 10 (fixo) ou 11 (celular) com DDD
  if (num.length === 11) {
    // (xx) xxxxx-xxxx
    return num.replace(/(\d{2})(\d{5})(\d{4})/, '($1) $2-$3');
  }
  if (num.length === 10) {
    // (xx) xxxx-xxxx
    return num.replace(/(\d{2})(\d{4})(\d{4})/, '($1) $2-$3');
  }

  // Se não bater (ex.: ramal, número curto, etc.), devolve o original
  return String(v ?? '');
}

// CPF (11) => 000.000.000-00 | CNPJ (14) => 00.000.000/0000-00
function formatarCpfCnpj(v) {
  const d = somenteDigitos(v);

  if (d.length === 11) {
    return d.replace(/(\d{3})(\d{3})(\d{3})(\d{2})/, '$1.$2.$3-$4');
  }
  if (d.length === 14) {
    return d.replace(/(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})/, '$1.$2.$3/$4-$5');
  }

  return String(v ?? '');
}

function tipoDoc(c) {
  const d = somenteDigitos(c?.DOCUMENTO);
  if (d.length === 11) return 'CPF';
  if (d.length === 14) return 'CNPJ';
  return '';
}

function renderTabelaClientes() {
  const tbody = document.getElementById('tbodyClientes');
  if (!tbody) return;

  const itens = getFiltroClientes();

  if (!itens.length) {
    tbody.innerHTML = `
      <tr>
        <td class="px-4 py-6 text-sm text-muted-foreground" colspan="5">Nenhum cliente encontrado.</td>
      </tr>
    `;
    return;
  }

  tbody.innerHTML = itens.map(c => {
    const cidadeUf = `${c.CIDADE || '-'} / ${c.UF || '-'}`;
    const contatoNome = c.CONTATO_NOME ? escapeHtml(c.CONTATO_NOME) : '-';
    const telFmt = formatarTelefoneBR(c.CONTATO_TELEFONE);
    const docFmt = c.DOCUMENTO ? formatarCpfCnpj(c.DOCUMENTO) : '';

    const tipo = tipoDoc(c);

    // escolha dos ícones
    const iconClass = tipo === 'CNPJ'
      ? 'fas fa-building'     // empresa
      : tipo === 'CPF'
        ? 'fas fa-id-card'    // pessoa
        : 'fas fa-file-lines';// fallback (se tiver FA6) ou troque por outro

    const badge = tipo
      ? `<span class="text-[10px] px-2 py-0.5 rounded-full border border-border bg-white/60 text-muted-foreground">${tipo}</span>`
      : '';

    return `
      <tr class="hover:bg-white/40 transition-all">
        <td class="px-4 py-3 font-medium">
          <div class="flex items-center gap-2">
            <i class="${iconClass} text-muted-foreground" aria-hidden="true"></i>
            <span>${escapeHtml(c.RAZAO_SOCIAL || '')}</span>
            ${badge}
          </div>
        </td>
        <td class="px-4 py-3 text-muted-foreground">${escapeHtml(docFmt)}</td>
        <td class="px-4 py-3">${escapeHtml(cidadeUf)}</td>
        <td class="px-4 py-3">${contatoNome}</td>
        <td class="px-4 py-3">${escapeHtml(telFmt)}</td>
        <td class="px-4 py-3">
          <div class="flex justify-end gap-2">
            <button class="btnEditarCliente w-10 h-10 rounded-xl border border-border bg-white/60 hover:bg-white/90 transition-all"
              data-id="${escapeHtml(c.ID)}" title="Editar">
              <i class="fas fa-pen"></i>
            </button>
            <button class="btnExcluirCliente w-10 h-10 rounded-xl border border-border bg-white/60 hover:bg-destructive hover:text-white transition-all"
              data-id="${escapeHtml(c.ID)}" title="Desativar">
              <i class="fas fa-trash"></i>
            </button>
          </div>
        </td>
      </tr>
    `;
  }).join('');
}

// ===== Modal Cliente (criado no DOM, igual seu padrão) =====
function removerModalCliente() {
  document.getElementById('clienteOverlay')?.remove();
  document.getElementById('clienteModal')?.remove();
}

async function buscarFiliaisCliente(idCliente) {
  const data = await apiGet(`/api/clientes/${encodeURIComponent(idCliente)}/filiais`);
  return Array.isArray(data?.items) ? data.items : [];
}

function abrirModalCliente({ modo, cliente }) {
  removerModalCliente();

  const overlay = document.createElement('div');
  overlay.id = 'clienteOverlay';
  overlay.className = 'fixed inset-0 bg-black/30 backdrop-blur-sm z-90';
  document.body.appendChild(overlay);

  const modal = document.createElement('div');
  modal.id = 'clienteModal';
  modal.className = 'fixed inset-0 z-100';
  modal.setAttribute('role', 'dialog');
  modal.setAttribute('aria-modal', 'true');

  const isEdit = modo === 'edit';

  modal.innerHTML = `
    <div class="w-full h-full overflow-y-auto no-scrollbar flex items-start justify-center p-4 md:p-8">
      <div class="w-full max-w-4xl mx-auto">
        <div class="glass rounded-2xl shadow-2xl border border-border overflow-hidden">
          <div class="px-6 py-5 border-b border-border flex items-start justify-between gap-4">
            <div>
              <h3 class="text-lg font-semibold text-foreground">${isEdit ? 'Editar cliente' : 'Cadastrar cliente'}</h3>
              <p class="text-xs text-muted-foreground">Dados do cliente e filiais</p>
            </div>
            <button id="btnFecharClienteModal" type="button"
              class="w-10 h-10 rounded-xl bg-white/60 border border-border hover:bg-white transition-all flex items-center justify-center"
              aria-label="Fechar" title="Fechar">
              <i class="fas fa-times"></i>
            </button>
          </div>

          <form id="formClienteModal" class="px-6 py-6 space-y-4 text-sm">
            <input type="hidden" id="cliId" value="${escapeHtml(cliente?.ID || '')}" />

            <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
              <div class="space-y-1">
                <label class="text-sm font-medium">CNPJ/CPF *</label>
                <input id="cliDocumento" required inputmode="numeric"
                  class="w-full rounded-xl border border-border bg-white/70 px-3 py-2 outline-none focus:ring-2 focus:ring-primary/30"
                  value="${escapeHtml(cliente?.DOCUMENTO || '')}" />
              </div>
              <div class="space-y-1">
                <label class="text-sm font-medium">Razão social *</label>
                <input id="cliRazao" required
                  class="w-full rounded-xl border border-border bg-white/70 px-3 py-2 outline-none focus:ring-2 focus:ring-primary/30"
                  value="${escapeHtml(cliente?.RAZAO_SOCIAL || '')}" />
              </div>
            </div>

            <div class="space-y-1">
              <label class="text-sm font-medium">Grupo econômico</label>
              <input id="cliGrupo"
                class="w-full rounded-xl border border-border bg-white/70 px-3 py-2 outline-none focus:ring-2 focus:ring-primary/30"
                value="${escapeHtml(cliente?.GRUPO_ECONOMICO || '')}" />
            </div>

            <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
              <div class="space-y-1">
                <label class="text-sm font-medium">UF *</label>
                <select id="cliUF" required
                  class="w-full rounded-xl border border-border bg-white/70 px-3 py-2 outline-none focus:ring-2 focus:ring-primary/30">
                  <option value="" selected disabled>Selecione...</option>
                </select>
              </div>

              <div class="space-y-1">
                <label class="text-sm font-medium">Cidade *</label>
                <select id="cliCidade" required disabled
                  class="w-full rounded-xl border border-border bg-white/70 px-3 py-2 outline-none focus:ring-2 focus:ring-primary/30">
                  <option value="" selected disabled>Selecione a UF primeiro...</option>
                </select>
              </div>
            </div>

            <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
              <div class="space-y-1">
                <label class="text-sm font-medium">Contato</label>
                <input id="cliContato"
                  class="w-full rounded-xl border border-border bg-white/70 px-3 py-2 outline-none focus:ring-2 focus:ring-primary/30"
                  value="${escapeHtml(cliente?.CONTATO_NOME || '')}" />
              </div>
              <div class="space-y-1">
                <label class="text-sm font-medium">Telefone</label>
                <input id="cliTelefone"
                  class="w-full rounded-xl border border-border bg-white/70 px-3 py-2 outline-none focus:ring-2 focus:ring-primary/30"
                  value="${escapeHtml(cliente?.CONTATO_TELEFONE || '')}" />
              </div>
            </div>

            <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
              <div class="space-y-1">
                <label class="text-sm font-medium">Email</label>
                <input id="cliEmail" type="email"
                  class="w-full rounded-xl border border-border bg-white/70 px-3 py-2 outline-none focus:ring-2 focus:ring-primary/30"
                  value="${escapeHtml(cliente?.CONTATO_EMAIL || '')}" />
              </div>
              <div class="space-y-1">
                <label class="text-sm font-medium">Cultura principal</label>
                <select id="cliCultura"
                  class="w-full rounded-xl border border-border bg-white/70 px-3 py-2 outline-none focus:ring-2 focus:ring-primary/30">
                  <option value="" ${!cliente?.CULTURA_PRINCIPAL ? 'selected' : ''}>Selecione...</option>
                  <option value="soja" ${cliente?.CULTURA_PRINCIPAL === 'soja' ? 'selected' : ''}>Soja</option>
                  <option value="algodao" ${cliente?.CULTURA_PRINCIPAL === 'algodao' ? 'selected' : ''}>Algodão</option>
                  <option value="forrageira" ${cliente?.CULTURA_PRINCIPAL === 'forrageira' ? 'selected' : ''}>Forrageira</option>
                </select>
              </div>
            </div>

            <div class="space-y-1">
              <label class="text-sm font-medium">Hectares estimados</label>
              <input id="cliHectares" type="number" min="0" step="1"
                class="w-full rounded-xl border border-border bg-white/70 px-3 py-2 outline-none focus:ring-2 focus:ring-primary/30"
                value="${escapeHtml((cliente?.HECTARES_ESTIMADOS ?? '') + '')}" />
            </div>

            <div class="space-y-1">
              <label class="text-sm font-medium">Observações</label>
              <textarea id="cliObs" rows="2"
                class="w-full rounded-xl border border-border bg-white/70 px-3 py-2 outline-none focus:ring-2 focus:ring-primary/30">${escapeHtml(cliente?.OBSERVACOES || '')}</textarea>
            </div>

            <hr class="border-border" />

            <div class="space-y-3">
              <div class="flex items-center justify-between gap-3 flex-wrap">
                <div class="flex items-center gap-2">
                  <i class="fas fa-code-branch text-muted-foreground"></i>
                  <h4 class="font-semibold text-sm">Filiais</h4>
                </div>
                <button id="btnAddFilial" type="button"
                  class="rounded-xl border border-border bg-white/60 hover:bg-white/90 px-3 py-2 text-sm font-semibold transition-all">
                  <i class="fas fa-plus mr-2"></i>Adicionar filial
                </button>
              </div>

              <div id="filiaisVazio" class="text-sm text-muted-foreground text-center py-2">Nenhuma filial cadastrada</div>
              <div id="boxFiliais" class="space-y-3"></div>
            </div>

            <p id="cliErro" class="text-sm text-destructive hidden whitespace-pre-line"></p>

            <div class="flex justify-end gap-3 pt-2">
              <button id="btnCancelarClienteModal" type="button"
                class="rounded-xl border border-border bg-white/60 hover:bg-white/90 px-4 py-2 font-semibold transition-all">
                Cancelar
              </button>
              <button id="btnSalvarClienteModal" type="submit"
                class="rounded-xl bg-primary text-white px-4 py-2 font-semibold hover:opacity-90 transition-all">
                Salvar
              </button>
            </div>
          </form>
        </div>
      </div>
    </div>
  `;

  document.body.appendChild(modal);

  const setErr = (msg) => {
    const el = document.getElementById('cliErro');
    if (!el) return;
    if (!msg) { el.textContent = ''; el.classList.add('hidden'); return; }
    el.textContent = msg;
    el.classList.remove('hidden');
  };

  function fechar() { removerModalCliente(); }

  overlay.addEventListener('click', fechar);
  document.getElementById('btnFecharClienteModal')?.addEventListener('click', fechar);
  document.getElementById('btnCancelarClienteModal')?.addEventListener('click', fechar);

  // =========================
  // CPF/CNPJ máscara
  // =========================
  function apenasDigitos(s) { return (s ?? '').toString().replace(/\D+/g, ''); }

  function formatCPF(d) {
    const v = apenasDigitos(d).slice(0, 11);
    return v
      .replace(/^(\d{3})(\d)/, '$1.$2')
      .replace(/^(\d{3})\.(\d{3})(\d)/, '$1.$2.$3')
      .replace(/^(\d{3})\.(\d{3})\.(\d{3})(\d)/, '$1.$2.$3-$4');
  }

  function formatCNPJ(d) {
    const v = apenasDigitos(d).slice(0, 14);
    return v
      .replace(/^(\d{2})(\d)/, '$1.$2')
      .replace(/^(\d{2})\.(\d{3})(\d)/, '$1.$2.$3')
      .replace(/^(\d{2})\.(\d{3})\.(\d{3})(\d)/, '$1.$2.$3/$4')
      .replace(/^(\d{2})\.(\d{3})\.(\d{3})\/(\d{4})(\d)/, '$1.$2.$3/$4-$5');
  }

  function formatDocAuto(raw) {
    const d = apenasDigitos(raw);
    if (d.length <= 11) return formatCPF(d);
    return formatCNPJ(d);
  }

  function maskDocumentoInput(inp) {
    const before = inp.value;
    const after = formatDocAuto(before);
    if (before !== after) inp.value = after;
  }

  const inpDoc = document.getElementById('cliDocumento');
  inpDoc?.addEventListener('input', () => maskDocumentoInput(inpDoc));
  inpDoc?.addEventListener('blur', () => maskDocumentoInput(inpDoc));

  // =========================
  // IBGE UF/Cidade (cache)
  // =========================
  const IBGE_BASE = 'https://servicodados.ibge.gov.br/api/v1/localidades'; // [web:805]
  let cacheUFs = null;
  const cacheMunicipiosPorUF = new Map();

  async function ibgeGetJson(url) {
    const r = await fetch(url);
    if (!r.ok) throw new Error(`IBGE falhou (HTTP ${r.status})`);
    return await r.json();
  }

  async function garantirUFsCache() {
    if (cacheUFs) return cacheUFs;
    const estados = await ibgeGetJson(`${IBGE_BASE}/estados?orderBy=nome`); // [web:805]
    cacheUFs = estados.map(e => ({ sigla: e.sigla, nome: e.nome }));
    return cacheUFs;
  }

  async function getMunicipiosPorUF(uf) {
    const ufUp = (uf || '').toUpperCase().trim();
    if (!ufUp) return [];
    if (cacheMunicipiosPorUF.has(ufUp)) return cacheMunicipiosPorUF.get(ufUp);

    const municipios = await ibgeGetJson(`${IBGE_BASE}/estados/${encodeURIComponent(ufUp)}/municipios?orderBy=nome`); // [web:805]
    const cidades = municipios.map(m => m.nome);
    cacheMunicipiosPorUF.set(ufUp, cidades);
    return cidades;
  }

  function preencherSelectUF(selectEl, ufSelecionada) {
    const atual = (ufSelecionada || '').toUpperCase().trim();
    selectEl.innerHTML =
      `<option value="" disabled ${!atual ? 'selected' : ''}>Selecione...</option>` +
      (cacheUFs || []).map(e => `
        <option value="${escapeHtml(e.sigla)}" ${e.sigla === atual ? 'selected' : ''}>
          ${escapeHtml(`${e.sigla} - ${e.nome}`)}
        </option>
      `).join('');
  }

  function preencherSelectCidade(selectEl, cidades, cidadeSelecionada) {
    const atual = (cidadeSelecionada || '').toString().trim();
    selectEl.innerHTML =
      `<option value="" disabled ${!atual ? 'selected' : ''}>Selecione...</option>` +
      cidades.map(nome => `
        <option value="${escapeHtml(nome)}" ${nome === atual ? 'selected' : ''}>
          ${escapeHtml(nome)}
        </option>
      `).join('');
  }

  async function carregarUFsSelectCliente(ufSelecionada) {
    await garantirUFsCache();
    const selUF = document.getElementById('cliUF');
    if (!selUF) return;
    preencherSelectUF(selUF, ufSelecionada);
  }

  async function carregarCidadesDaUF(uf, cidadeSelecionada) {
    const selCidade = document.getElementById('cliCidade');
    if (!selCidade) return;

    const ufUp = (uf || '').toUpperCase().trim();
    if (!ufUp) {
      selCidade.disabled = true;
      selCidade.innerHTML = `<option value="" selected disabled>Selecione a UF primeiro...</option>`;
      return;
    }

    selCidade.disabled = true;
    selCidade.innerHTML = `<option value="" selected disabled>Carregando cidades...</option>`;

    const cidades = await getMunicipiosPorUF(ufUp);
    preencherSelectCidade(selCidade, cidades, cidadeSelecionada);
    selCidade.disabled = false;
  }

  async function inicializarUfCidade(cli) {
    const ufInicial = (cli?.UF || '').toString().trim().toUpperCase();
    const cidadeInicial = (cli?.CIDADE || '').toString().trim();

    await carregarUFsSelectCliente(ufInicial);

    if (ufInicial) await carregarCidadesDaUF(ufInicial, cidadeInicial);
    else await carregarCidadesDaUF('', '');

    document.getElementById('cliUF')?.addEventListener('change', async (e) => {
      await carregarCidadesDaUF(e.target.value, '');
    });
  }

  // =========================
  // CNPJ -> BrasilAPI -> confirmação
  // =========================
  function removerModalConfirmCNPJ() {
    document.getElementById('cnpjFillOverlay')?.remove();
    document.getElementById('cnpjFillModal')?.remove();
  }

  function abrirModalConfirmCNPJ({ titulo, html, onConfirm }) {
    removerModalConfirmCNPJ();

    const ov = document.createElement('div');
    ov.id = 'cnpjFillOverlay';
    ov.className = 'fixed inset-0 bg-black/35 backdrop-blur-sm z-[110]';
    document.body.appendChild(ov);

    const m = document.createElement('div');
    m.id = 'cnpjFillModal';
    m.className = 'fixed inset-0 z-[120] flex items-start justify-center p-4 md:p-8';
    m.innerHTML = `
      <div class="w-full max-w-xl glass rounded-2xl shadow-2xl border border-border overflow-hidden">
        <div class="px-5 py-4 border-b border-border">
          <h3 class="text-lg font-semibold">${escapeHtml(titulo || 'Encontramos dados para este CNPJ')}</h3>
          <p class="text-sm text-muted-foreground">Deseja preencher automaticamente? Você ainda poderá editar depois.</p>
        </div>
        <div class="px-5 py-4 text-sm space-y-3">
          <div class="rounded-xl border border-border bg-white/50 p-4">
            ${html || ''}
          </div>
          <div class="flex gap-2 justify-end">
            <button id="btnCnpjFillCancelar" type="button"
              class="rounded-xl border border-border bg-white/60 hover:bg-white/90 px-4 py-2 font-semibold transition-all">
              Não
            </button>
            <button id="btnCnpjFillOk" type="button"
              class="rounded-xl bg-primary text-white px-4 py-2 font-semibold hover:opacity-90 transition-all">
              Sim, preencher
            </button>
          </div>
        </div>
      </div>
    `;
    document.body.appendChild(m);

    const fechar = () => removerModalConfirmCNPJ();
    ov.addEventListener('click', fechar);
    document.getElementById('btnCnpjFillCancelar')?.addEventListener('click', fechar);
    document.getElementById('btnCnpjFillOk')?.addEventListener('click', async () => {
      try { await onConfirm?.(); } finally { fechar(); }
    });
  }

  function setIfEmpty(inputId, value) {
    const el = document.getElementById(inputId);
    if (!el) return;
    const cur = (el.value || '').trim();
    if (cur) return;
    el.value = value ?? '';
  }

  let ultimoCnpjConsultado = null;

  async function consultarCNPJ_BrasilAPI(cnpj14) {
    const url = `https://brasilapi.com.br/api/cnpj/v1/${encodeURIComponent(cnpj14)}`; // [web:848]
    const r = await fetch(url);
    if (!r.ok) throw new Error(`Consulta CNPJ falhou (HTTP ${r.status})`);
    return await r.json();
  }

  async function tentarAutopreencherPorCNPJ() {
    const doc = apenasDigitos(document.getElementById('cliDocumento')?.value);
    if (doc.length !== 14) return;

    if (ultimoCnpjConsultado === doc) return;
    ultimoCnpjConsultado = doc;

    let dados;
    try {
      dados = await consultarCNPJ_BrasilAPI(doc); // [web:848]
    } catch (e) {
      console.warn('CNPJ lookup:', e?.message || e);
      return;
    }

    const nome = dados?.razao_social || dados?.nome || '';
    const fantasia = dados?.nome_fantasia || '';
    const municipio = dados?.municipio || '';
    const uf = (dados?.uf || '').toUpperCase();
    const email = dados?.email || '';
    const telefone = dados?.ddd_telefone_1 || dados?.telefone || '';

    const preview = `
      <div><span class="text-muted-foreground">Razão Social:</span> <span class="font-semibold">${escapeHtml(nome || '-')}</span></div>
      <div><span class="text-muted-foreground">Fantasia:</span> <span class="font-semibold">${escapeHtml(fantasia || '-')}</span></div>
      <div><span class="text-muted-foreground">Cidade/UF:</span> <span class="font-semibold">${escapeHtml(municipio || '-')}/${escapeHtml(uf || '-')}</span></div>
      <div><span class="text-muted-foreground">Email:</span> <span class="font-semibold">${escapeHtml(email || '-')}</span></div>
      <div><span class="text-muted-foreground">Telefone:</span> <span class="font-semibold">${escapeHtml(telefone || '-')}</span></div>
    `;

    abrirModalConfirmCNPJ({
      titulo: 'CNPJ encontrado',
      html: preview,
      onConfirm: async () => {
        setIfEmpty('cliRazao', nome);
        setIfEmpty('cliEmail', email);
        setIfEmpty('cliTelefone', telefone);

        const selUF = document.getElementById('cliUF');
        if (selUF && uf) {
          selUF.value = uf;
          await carregarCidadesDaUF(uf, municipio);
        }
      }
    });
  }

  inpDoc?.addEventListener('blur', () => {
    tentarAutopreencherPorCNPJ().catch(console.error);
  });
  inpDoc?.addEventListener('input', () => {
    const d = apenasDigitos(inpDoc.value);
    if (d.length === 14) tentarAutopreencherPorCNPJ().catch(console.error);
  });

  // =========================
  // Filiais render + UF/Cidade por linha
  // =========================
  async function inicializarUfCidadeFiliais() {
    await garantirUFsCache();

    for (let idx = 0; idx < filiaisForm.length; idx++) {
      const selUF = document.getElementById(`filialUF_${idx}`);
      const selCidade = document.getElementById(`filialCidade_${idx}`);
      if (!selUF || !selCidade) continue;

      preencherSelectUF(selUF, filiaisForm[idx]?.uf);

      const ufAtual = (selUF.value || '').toUpperCase().trim();
      if (ufAtual) {
        selCidade.disabled = true;
        selCidade.innerHTML = `<option value="" disabled selected>Carregando cidades...</option>`;
        const cidades = await getMunicipiosPorUF(ufAtual);
        preencherSelectCidade(selCidade, cidades, filiaisForm[idx]?.cidade);
        selCidade.disabled = false;
      } else {
        selCidade.disabled = true;
        selCidade.innerHTML = `<option value="" disabled selected>Selecione a UF primeiro...</option>`;
      }

      selUF.onchange = async () => {
        const ufNova = selUF.value;
        filiaisForm[idx] = { ...filiaisForm[idx], uf: ufNova, cidade: '' };

        selCidade.disabled = true;
        selCidade.innerHTML = `<option value="" disabled selected>Carregando cidades...</option>`;
        const cidades = await getMunicipiosPorUF(ufNova);
        preencherSelectCidade(selCidade, cidades, '');
        selCidade.disabled = false;
      };

      selCidade.onchange = () => {
        filiaisForm[idx] = { ...filiaisForm[idx], cidade: selCidade.value };
      };
    }
  }

  function renderFiliais() {
    const box = document.getElementById('boxFiliais');
    const vazio = document.getElementById('filiaisVazio');
    if (!box || !vazio) return;

    if (!filiaisForm.length) {
      vazio.classList.remove('hidden');
      box.innerHTML = '';
      return;
    }
    vazio.classList.add('hidden');

    box.innerHTML = filiaisForm.map((f, idx) => `
      <div class="border rounded-lg p-4 space-y-3 bg-white/40 text-sm">
        <div class="flex items-center justify-between">
          <span class="font-medium text-sm">Filial ${idx + 1}</span>
          <button type="button"
            class="w-10 h-10 rounded-xl border border-border bg-white/60 hover:bg-destructive hover:text-white transition-all"
            title="Remover filial" data-idx="${idx}">
            <i class="fas fa-trash"></i>
          </button>
        </div>

        <div class="grid grid-cols-1 md:grid-cols-2 gap-3">
          <div class="space-y-1">
            <label class="text-sm font-medium">Nome *</label>
            <input class="w-full rounded-xl border border-border bg-white/70 px-3 py-2 outline-none focus:ring-2 focus:ring-primary/30"
              data-field="nome" data-idx="${idx}" value="${escapeHtml(f.nome || '')}" required />
          </div>

          <div class="space-y-1">
            <label class="text-sm font-medium">Endereço</label>
            <input class="w-full rounded-xl border border-border bg-white/70 px-3 py-2 outline-none focus:ring-2 focus:ring-primary/30"
              data-field="endereco" data-idx="${idx}" value="${escapeHtml(f.endereco || '')}" />
          </div>

          <div class="space-y-1">
            <label class="text-sm font-medium">UF *</label>
            <select id="filialUF_${idx}"
              class="w-full rounded-xl border border-border bg-white/70 px-3 py-2 outline-none focus:ring-2 focus:ring-primary/30"
              data-idx="${idx}" required>
              <option value="" disabled selected>Selecione...</option>
            </select>
          </div>

          <div class="space-y-1">
            <label class="text-sm font-medium">Cidade *</label>
            <select id="filialCidade_${idx}"
              class="w-full rounded-xl border border-border bg-white/70 px-3 py-2 outline-none focus:ring-2 focus:ring-primary/30"
              data-idx="${idx}" required disabled>
              <option value="" disabled selected>Selecione a UF primeiro...</option>
            </select>
          </div>

          <div class="space-y-1">
            <label class="text-sm font-medium">Contato</label>
            <input class="w-full rounded-xl border border-border bg-white/70 px-3 py-2 outline-none focus:ring-2 focus:ring-primary/30"
              data-field="contato_nome" data-idx="${idx}" value="${escapeHtml(f.contato_nome || '')}" />
          </div>

          <div class="space-y-1">
            <label class="text-sm font-medium">Telefone</label>
            <input class="w-full rounded-xl border border-border bg-white/70 px-3 py-2 outline-none focus:ring-2 focus:ring-primary/30"
              data-field="contato_telefone" data-idx="${idx}" value="${escapeHtml(f.contato_telefone || '')}" />
          </div>
        </div>
      </div>
    `).join('');

    box.querySelectorAll('button[data-idx]').forEach(btn => {
      btn.addEventListener('click', () => {
        const idx = Number(btn.getAttribute('data-idx'));
        filiaisForm = filiaisForm.filter((_, i) => i !== idx);
        renderFiliais();
      });
    });

    box.querySelectorAll('input[data-field]').forEach(inp => {
      inp.addEventListener('input', () => {
        const idx = Number(inp.getAttribute('data-idx'));
        const field = inp.getAttribute('data-field');
        if (!Number.isFinite(idx) || !field) return;
        filiaisForm[idx] = { ...filiaisForm[idx], [field]: inp.value };
      });
    });

    inicializarUfCidadeFiliais().catch(err => console.error('UF/Cidade filiais:', err));
  }

  async function carregarFiliaisSeEdicao() {
    const id = document.getElementById('cliId')?.value;
    if (!id) { filiaisForm = []; renderFiliais(); return; }

    try {
      const rows = await buscarFiliaisCliente(id);
      filiaisForm = rows.map(r => ({
        id: r.ID,
        nome: r.NOME,
        endereco: r.ENDERECO,
        cidade: r.CIDADE,
        uf: r.UF,
        contato_nome: r.CONTATO_NOME,
        contato_telefone: r.CONTATO_TELEFONE,
      }));
      renderFiliais();
    } catch (e) {
      console.error(e);
      filiaisForm = [];
      renderFiliais();
    }
  }

  document.getElementById('btnAddFilial')?.addEventListener('click', () => {
    const telCli = document.getElementById('cliTelefone')?.value || '';
    filiaisForm.push({
      id: null,
      nome: '',
      endereco: '',
      cidade: '',
      uf: '',
      contato_nome: '',
      contato_telefone: telCli,
    });
    renderFiliais();
  });

  // =========================
  // Submit
  // =========================
  document.getElementById('formClienteModal')?.addEventListener('submit', async (e) => {
    e.preventDefault();
    setErr('');

    const btn = document.getElementById('btnSalvarClienteModal');
    btn.disabled = true;
    btn.classList.add('opacity-70');
    btn.textContent = 'Salvando...';

    try {
      const id = (document.getElementById('cliId')?.value || '').trim();

      const payload = {
        cliente: {
          id: id || null,
          razao_social: document.getElementById('cliRazao')?.value,
          documento: document.getElementById('cliDocumento')?.value, // vem mascarado
          grupo_economico: document.getElementById('cliGrupo')?.value,
          uf: document.getElementById('cliUF')?.value,
          cidade: document.getElementById('cliCidade')?.value,
          contato_nome: document.getElementById('cliContato')?.value,
          contato_telefone: document.getElementById('cliTelefone')?.value,
          contato_email: document.getElementById('cliEmail')?.value,
          cultura_principal: document.getElementById('cliCultura')?.value,
          hectares_estimados: document.getElementById('cliHectares')?.value,
          observacoes: document.getElementById('cliObs')?.value,
        },
        filiais: filiaisForm.map(f => ({
          id: f.id,
          nome: f.nome,
          endereco: f.endereco,
          uf: f.uf,
          cidade: f.cidade,
          contato_nome: f.contato_nome,
          contato_telefone: f.contato_telefone,
        }))
      };

      await apiSend('/api/clientes/salvar', 'POST', payload);

      removerModalCliente();
      await carregarClientes();
    } catch (err) {
      setErr(err?.message || 'Erro ao salvar.');
    } finally {
      btn.disabled = false;
      btn.classList.remove('opacity-70');
      btn.textContent = 'Salvar';
    }
  });

  // init
  carregarFiliaisSeEdicao();
  inicializarUfCidade(cliente).catch(err => console.error('UF/Cidade:', err));
}

// ====== UF/Cidade (IBGE) ======
const IBGE_BASE = 'https://servicodados.ibge.gov.br/api/v1/localidades'; // [web:805]
let cacheUFs = null;                 // [{ sigla, nome }]
const cacheMunicipiosPorUF = new Map(); // 'BA' -> ['Salvador', '...']

async function ibgeGetJson(url) {
  const r = await fetch(url, { method: 'GET' });
  // padrão fetch: checar ok e só então parsear [web:793]
  if (!r.ok) throw new Error(`IBGE falhou (HTTP ${r.status})`);
  return await r.json(); // [web:794]
}

function option(text, value, selected = false) {
  const sel = selected ? 'selected' : '';
  return `<option value="${escapeHtml(value)}" ${sel}>${escapeHtml(text)}</option>`;
}

async function carregarUFsSelect(valorSelecionado) {
  const selUF = document.getElementById('cliUF');
  if (!selUF) return;

  if (!cacheUFs) {
    const estados = await ibgeGetJson(`${IBGE_BASE}/estados?orderBy=nome`); // [web:805]
    cacheUFs = estados.map(e => ({ sigla: e.sigla, nome: e.nome }));
  }

  const atual = (valorSelecionado || '').toUpperCase();

  selUF.innerHTML =
    `<option value="" disabled ${!atual ? 'selected' : ''}>Selecione...</option>` +
    cacheUFs.map(e => option(`${e.sigla} - ${e.nome}`, e.sigla, e.sigla === atual)).join('');
}

async function carregarCidadesDaUF(uf, cidadeSelecionada) {
  const selCidade = document.getElementById('cliCidade');
  if (!selCidade) return;

  const ufUp = (uf || '').toUpperCase().trim();
  if (!ufUp) {
    selCidade.disabled = true;
    selCidade.innerHTML = `<option value="" selected disabled>Selecione a UF primeiro...</option>`;
    return;
  }

  selCidade.disabled = true;
  selCidade.innerHTML = `<option value="" selected disabled>Carregando cidades...</option>`;

  let cidades = cacheMunicipiosPorUF.get(ufUp);
  if (!cidades) {
    const municipios = await ibgeGetJson(`${IBGE_BASE}/estados/${encodeURIComponent(ufUp)}/municipios?orderBy=nome`); // [web:805]
    cidades = municipios.map(m => m.nome);
    cacheMunicipiosPorUF.set(ufUp, cidades);
  }

  const cidadeCur = (cidadeSelecionada || '').toString().trim();

  selCidade.innerHTML =
    `<option value="" disabled ${!cidadeCur ? 'selected' : ''}>Selecione...</option>` +
    cidades.map(nome => option(nome, nome, nome === cidadeCur)).join('');

  selCidade.disabled = false;
}

async function inicializarUfCidade(cliente) {
  const ufInicial = (cliente?.UF || '').toString().trim().toUpperCase();
  const cidadeInicial = (cliente?.CIDADE || '').toString().trim();

  await carregarUFsSelect(ufInicial);

  // carregue cidades se já tiver UF (edição)
  if (ufInicial) {
    await carregarCidadesDaUF(ufInicial, cidadeInicial);
  } else {
    await carregarCidadesDaUF('', '');
  }

  // quando muda UF, limpa cidade e recarrega
  document.getElementById('cliUF')?.addEventListener('change', async (e) => {
    const uf = e.target.value;
    await carregarCidadesDaUF(uf, '');
  });
}


async function garantirUFsCache() {
  if (cacheUFs) return cacheUFs;
  const estados = await ibgeGetJson(`${IBGE_BASE}/estados?orderBy=nome`); // [web:805]
  cacheUFs = estados.map(e => ({ sigla: e.sigla, nome: e.nome }));
  return cacheUFs;
}

async function getMunicipiosPorUF(uf) {
  const ufUp = (uf || '').toUpperCase().trim();
  if (!ufUp) return [];
  if (cacheMunicipiosPorUF.has(ufUp)) return cacheMunicipiosPorUF.get(ufUp);

  const municipios = await ibgeGetJson(`${IBGE_BASE}/estados/${encodeURIComponent(ufUp)}/municipios?orderBy=nome`); // [web:805]
  const cidades = municipios.map(m => m.nome);
  cacheMunicipiosPorUF.set(ufUp, cidades);
  return cidades;
}

function preencherSelectUF(selectEl, ufSelecionada) {
  const atual = (ufSelecionada || '').toUpperCase().trim();
  selectEl.innerHTML =
    `<option value="" disabled ${!atual ? 'selected' : ''}>Selecione...</option>` +
    cacheUFs.map(e => `
      <option value="${escapeHtml(e.sigla)}" ${e.sigla === atual ? 'selected' : ''}>
        ${escapeHtml(`${e.sigla} - ${e.nome}`)}
      </option>
    `).join('');
}

function preencherSelectCidade(selectEl, cidades, cidadeSelecionada) {
  const atual = (cidadeSelecionada || '').toString().trim();
  selectEl.innerHTML =
    `<option value="" disabled ${!atual ? 'selected' : ''}>Selecione...</option>` +
    cidades.map(nome => `
      <option value="${escapeHtml(nome)}" ${nome === atual ? 'selected' : ''}>
        ${escapeHtml(nome)}
      </option>
    `).join('');
}

// Emails Automáticos
let remetentesCache = [];
let destinatariosCache = [];
let remetentesParaSelect = [];

function statusBadgeEmails(status) {
  const s = (status || '').trim().toLowerCase();
  let cls = 'status-badge bg-info/15 text-info border-info/20';
  if (s === 'ativo') cls = 'status-badge status-ativo';
  else if (s === 'inativo') cls = 'status-badge status-inativo';
  return `<span class="${cls}">${escapeHtml(s)}</span>`;
}

function rowRemetente(rem, idx) {
  return `
    <tr>
      <td class="px-4 py-3 font-mono text-xs">${idx + 1}</td>
      <td class="px-4 py-3 font-medium">${escapeHtml(rem.EMAIL)}</td>
      <td class="px-4 py-3">${escapeHtml(rem.NOME || '')}</td>
      <td class="px-4 py-3">${statusBadgeEmails(rem.ATIVO ? 'Ativo' : 'Inativo')}</td>
      <td class="px-4 py-3">
        <div class="flex justify-end gap-2">
          <button class="btnEditRemetente w-10 h-10 rounded-xl border border-border bg-white/60 hover:bg-white/90 transition-all" data-id="${escapeHtml(rem.ID)}" title="Editar">
            <i class="fas fa-pen"></i>
          </button>
          <button class="btnDelRemetente w-10 h-10 rounded-xl border border-border bg-white/60 hover:bg-destructive hover:text-white transition-all" data-id="${escapeHtml(rem.ID)}" title="Desativar">
            <i class="fas fa-trash"></i>
          </button>
        </div>
      </td>
    </tr>`;
}

function rowDestinatario(dest, idx) {
  return `
    <tr>
      <td class="px-4 py-3">${escapeHtml(dest.remetenteNome || 'N/D')}</td>
      <td class="px-4 py-3 font-medium">${escapeHtml(dest.EMAIL_DESTINATARIO)}</td>
      <td class="px-4 py-3">${escapeHtml(dest.NOME_DESTINATARIO || '')}</td>
      <td class="px-4 py-3">${statusBadgeEmails(dest.ATIVO ? 'Ativo' : 'Inativo')}</td>
      <td class="px-4 py-3">
        <div class="flex justify-end gap-2">
          <button class="btnEditDest w-10 h-10 rounded-xl border border-border bg-white/60 hover:bg-white/90 transition-all" data-id="${escapeHtml(dest.ID)}" title="Editar">
            <i class="fas fa-pen"></i>
          </button>
          <button class="btnDelDest w-10 h-10 rounded-xl border border-border bg-white/60 hover:bg-destructive hover:text-white transition-all" data-id="${escapeHtml(dest.ID)}" title="Desativar">
            <i class="fas fa-trash"></i>
          </button>
        </div>
      </td>
    </tr>`;
}

async function carregarEmailsAutomaticos() {
  try {
    const dataRem = await apiGet('/api/emails/remetentes');
    const dataDest = await apiGet('/api/emails/destinatarios');
    
    remetentesCache = Array.isArray(dataRem?.items) ? dataRem.items : [];
    destinatariosCache = Array.isArray(dataDest?.items) ? dataDest.items : [];
    remetentesParaSelect = remetentesCache.map(r => ({
      id: r.ID,
      nome: `${r.EMAIL}${r.NOME ? ` (${r.NOME})` : ''}`
    }));

    renderEmails();
  } catch (err) {
    mostrarEmailsMsg('Erro ao carregar: ' + err.message);
  }
}

function renderEmails() {
  const tbodyRem = document.getElementById('tbodyRemetentes');
  const tbodyDest = document.getElementById('tbodyDestinatarios');
  const selectRem = document.getElementById('selectRemetenteDest');

  // Remetentes
  if (tbodyRem) {
    if (!remetentesCache.length) {
      tbodyRem.innerHTML = '<tr><td colspan="5" class="px-4 py-6 text-sm text-muted-foreground text-center">Nenhum remetente cadastrado</td></tr>';
    } else {
      tbodyRem.innerHTML = remetentesCache.map((r, i) => rowRemetente(r, i)).join('');
    }
  }

  // Destinatários
  if (tbodyDest) {
    if (!destinatariosCache.length) {
      tbodyDest.innerHTML = '<tr><td colspan="5" class="px-4 py-6 text-sm text-muted-foreground text-center">Nenhum destinatário cadastrado</td></tr>';
    } else {
      tbodyDest.innerHTML = destinatariosCache.map((d, i) => rowDestinatario(d, i)).join('');
    }
  }

  // Select remetentes
  if (selectRem) {
    selectRem.innerHTML = '<option value="">Selecione um remetente...</option>' +
      remetentesParaSelect.map(r => `<option value="${escapeHtml(r.id)}">${escapeHtml(r.nome)}</option>`).join('');
  }
}

function ativarAbaEmails(nomeAba) {
  const abaRem = document.getElementById('abaRemetentes');
  const abaDest = document.getElementById('abaDestinatarios');
  const painelRem = document.getElementById('painelRemetentes');
  const painelDest = document.getElementById('painelDestinatarios');

  // Ativa aba
  if (abaRem) abaRem.setAttribute('aria-selected', nomeAba === 'remetentes');
  if (abaDest) abaDest.setAttribute('aria-selected', nomeAba === 'destinatarios');

  // Classes
  if (abaRem && abaDest) {
    if (nomeAba === 'remetentes') {
      abaRem.className = 'px-4 py-2 rounded-lg text-sm font-medium transition-all bg-white shadow text-foreground';
      abaDest.className = 'px-4 py-2 rounded-lg text-sm font-medium transition-all text-muted-foreground hover:text-foreground';
    } else {
      abaDest.className = 'px-4 py-2 rounded-lg text-sm font-medium transition-all bg-white shadow text-foreground';
      abaRem.className = 'px-4 py-2 rounded-lg text-sm font-medium transition-all text-muted-foreground hover:text-foreground';
    }
  }

  // Painéis
  if (painelRem) painelRem.hidden = nomeAba !== 'remetentes';
  if (painelDest) painelDest.hidden = nomeAba !== 'destinatarios';
}

function mostrarEmailsMsg(msg) {
  const el = document.getElementById('emailsMsg');
  if (!el) return;
  if (!msg) {
    el.classList.add('hidden');
    return;
  }
  el.textContent = msg;
  el.classList.remove('hidden');
}

// Eventos
document.getElementById('abaRemetentes')?.addEventListener('click', () => ativarAbaEmails('remetentes'));
document.getElementById('abaDestinatarios')?.addEventListener('click', () => ativarAbaEmails('destinatarios'));

document.getElementById('btnAtualizarEmailsAuto')?.addEventListener('click', carregarEmailsAutomaticos);
document.getElementById('inputBuscaRemetentes')?.addEventListener('input', renderEmails);
document.getElementById('inputBuscaDestinatarios')?.addEventListener('input', renderEmails);
document.getElementById('selectRemetenteDest')?.addEventListener('change', renderEmails);

// Menu + cliques em botões (adicione no evento geral)
document.addEventListener('click', e => {
  const item = e.target.closest('[data-page="secao-emails-automaticos"]');
  if (item) {
    carregarEmailsAutomaticos().catch(err => mostrarEmailsMsg(err.message));
    ativarAbaEmails('remetentes'); // padrão remetentes
  }
  // Emails - botões novos
  const btnNovoRem = e.target.closest('#btnNovoRemetente');
  if (btnNovoRem) {
    abrirModalRemetente('new');
    return;
  }

  const btnNovoDest = e.target.closest('#btnNovoDestinatario');
  if (btnNovoDest) {
    abrirModalDestinatario('new');
    return;
  }

  // Emails - editar/excluir
  const btnEditRem = e.target.closest('.btnEditRemetente');
  if (btnEditRem) {
    const id = btnEditRem.dataset.id;
    const rem = remetentesCache.find(r => String(r.ID) === id);
    if (rem) abrirModalRemetente('edit', rem);
    return;
  }

  const btnDelRem = e.target.closest('.btnDelRemetente');
  if (btnDelRem) {
    const id = btnDelRem.dataset.id;
    if (!confirm('Deseja desativar este remetente?')) return;
    apiSend(`/api/emails/remetentes/${id}`, 'DELETE')
      .then(() => carregarEmailsAutomaticos())
      .catch(err => mostrarEmailsMsg('Erro: ' + err.message));
    return;
  }

  const btnEditDest = e.target.closest('.btnEditDest');
  if (btnEditDest) {
    const id = btnEditDest.dataset.id;
    const dest = destinatariosCache.find(d => String(d.ID) === id);
    if (dest) abrirModalDestinatario('edit', dest);
    return;
  }

  const btnDelDest = e.target.closest('.btnDelDest');
  if (btnDelDest) {
    const id = btnDelDest.dataset.id;
    if (!confirm('Deseja desativar este destinatário?')) return;
    apiSend(`/api/emails/destinatarios/${id}`, 'DELETE')
      .then(() => carregarEmailsAutomaticos())
      .catch(err => mostrarEmailsMsg('Erro: ' + err.message));
    return;
  }
});

// Modal Remetente
function removerModalRemetente() {
  document.getElementById('modalRemetenteOverlay')?.remove();
  document.getElementById('modalRemetente')?.remove();
}

function abrirModalRemetente(modo, remetente) {
  removerModalRemetente();
  
  const overlay = document.createElement('div');
  overlay.id = 'modalRemetenteOverlay';
  overlay.className = 'fixed inset-0 bg-black/40 backdrop-blur-sm z-[90]';
  document.body.appendChild(overlay);

  const modal = document.createElement('div');
  modal.id = 'modalRemetente';
  modal.className = 'fixed inset-0 z-[100]';
  modal.setAttribute('role', 'dialog');
  modal.setAttribute('aria-modal', 'true');
  
  const isEdit = modo === 'edit';
  const r = remetente || {};
  
  modal.innerHTML = `
    <div class="w-full h-full flex items-start justify-center p-4 md:p-8 overflow-auto">
      <div class="w-full max-w-md mx-auto">
        <div class="glass rounded-2xl shadow-2xl border border-border overflow-hidden">
          <div class="px-6 py-5 border-b border-border flex items-start justify-between gap-4">
            <div>
              <h3 class="text-xl font-semibold text-foreground">${isEdit ? 'Editar remetente' : 'Novo remetente'}</h3>
              <p class="text-sm text-muted-foreground">Preencha os dados abaixo</p>
            </div>
            <button id="btnFecharRemetente" type="button" class="w-10 h-10 rounded-xl bg-white/60 border border-border hover:bg-white transition-all flex items-center justify-center" aria-label="Fechar" title="Fechar">
              <i class="fas fa-times"></i>
            </button>
          </div>
          
          <form id="formRemetente" class="px-6 py-6 space-y-4">
            <div class="space-y-2">
              <label class="text-sm font-medium">Email <span class="text-destructive">*</span></label>
              <input id="remEmail" type="email" required 
                class="w-full rounded-xl border border-border bg-white/70 px-4 py-3 outline-none focus:ring-2 focus:ring-info/30" 
                value="${escapeHtml(r.EMAIL || '')}" placeholder="exemplo@fornecedor.com">
            </div>
            
            <div class="space-y-2">
              <label class="text-sm font-medium">Nome (opcional)</label>
              <input id="remNome" type="text" 
                class="w-full rounded-xl border border-border bg-white/70 px-4 py-3 outline-none focus:ring-2 focus:ring-info/30" 
                value="${escapeHtml(r.NOME || '')}" placeholder="Nome do remetente">
            </div>
            
            <p id="remErro" class="text-sm text-destructive hidden whitespace-pre-line"></p>
            
            <div class="pt-2 flex flex-col sm:flex-row gap-3">
              <button id="btnSalvarRemetente" type="submit" class="sm:flex-1 rounded-xl bg-info text-white px-4 py-3 font-medium hover:opacity-90 transition-all">
                Salvar
              </button>
              <button id="btnCancelarRemetente" type="button" class="sm:flex-1 rounded-xl border border-border bg-white/50 px-4 py-3 font-medium hover:bg-white/70 transition-all">
                Cancelar
              </button>
            </div>
          </form>
        </div>
      </div>
    </div>
  `;
  
  document.body.appendChild(modal);

  // Eventos
  const fechar = removerModalRemetente;
  overlay.addEventListener('click', fechar);
  document.getElementById('btnFecharRemetente')?.addEventListener('click', fechar);
  document.getElementById('btnCancelarRemetente')?.addEventListener('click', fechar);

  // Submit
  document.getElementById('formRemetente')?.addEventListener('submit', async e => {
    e.preventDefault();
    
    const email = document.getElementById('remEmail').value.trim();
    const nome = document.getElementById('remNome').value.trim();
    
    if (!email) {
      setRemErro('Email é obrigatório.');
      return;
    }
    
    try {
      const btn = document.getElementById('btnSalvarRemetente');
      btn.disabled = true;
      btn.classList.add('opacity-70');
      
      if (isEdit) {
        await apiSend(`/api/emails/remetentes/${r.ID}`, 'PUT', { EMAIL: email, NOME: nome || null, ATIVO: 1 });
      } else {
        await apiSend('/api/emails/remetentes', 'POST', { EMAIL: email, NOME: nome || null });
      }
      
      removerModalRemetente();
      await carregarEmailsAutomaticos();
      
    } catch (err) {
      setRemErro(err.message);
    } finally {
      const btn = document.getElementById('btnSalvarRemetente');
      btn.disabled = false;
      btn.classList.remove('opacity-70');
    }
  });
}

function setRemErro(msg) {
  const el = document.getElementById('remErro');
  if (!el) return;
  if (!msg) {
    el.classList.add('hidden');
    return;
  }
  el.textContent = msg;
  el.classList.remove('hidden');
}

// Modal Destinatário
function removerModalDestinatario() {
  document.getElementById('modalDestOverlay')?.remove();
  document.getElementById('modalDest')?.remove();
}

function abrirModalDestinatario(modo, destinatario) {
  removerModalDestinatario();
  
  const overlay = document.createElement('div');
  overlay.id = 'modalDestOverlay';
  overlay.className = 'fixed inset-0 bg-black/40 backdrop-blur-sm z-[90]';
  document.body.appendChild(overlay);

  const modal = document.createElement('div');
  modal.id = 'modalDest';
  modal.className = 'fixed inset-0 z-[100]';
  modal.setAttribute('role', 'dialog');
  modal.setAttribute('aria-modal', 'true');
  
  const isEdit = modo === 'edit';
  const d = destinatario || {};
  
  modal.innerHTML = `
    <div class="w-full h-full flex items-start justify-center p-4 md:p-8 overflow-auto">
      <div class="w-full max-w-md mx-auto">
        <div class="glass rounded-2xl shadow-2xl border border-border overflow-hidden">
          <div class="px-6 py-5 border-b border-border flex items-start justify-between gap-4">
            <div>
              <h3 class="text-xl font-semibold text-foreground">${isEdit ? 'Editar destinatário' : 'Novo destinatário'}</h3>
              <p class="text-sm text-muted-foreground">Preencha os dados abaixo</p>
            </div>
            <button id="btnFecharDest" type="button" class="w-10 h-10 rounded-xl bg-white/60 border border-border hover:bg-white transition-all flex items-center justify-center" aria-label="Fechar" title="Fechar">
              <i class="fas fa-times"></i>
            </button>
          </div>
          
          <form id="formDestinatario" class="px-6 py-6 space-y-4">
            <div class="space-y-2">
              <label class="text-sm font-medium">Remetente <span class="text-destructive">*</span></label>
              <select id="destRemetenteId" required class="w-full rounded-xl border border-border bg-white/70 px-4 py-3 outline-none focus:ring-2 focus:ring-info/30">
                <option value="">Selecione...</option>
                ${remetentesParaSelect.map(r => `<option value="${r.id}" ${r.id == d.ID_REMETENTE ? 'selected' : ''}>${escapeHtml(r.nome)}</option>`).join('')}
              </select>
            </div>
            
            <div class="space-y-2">
              <label class="text-sm font-medium">Email Destinatário <span class="text-destructive">*</span></label>
              <input id="destEmail" type="email" required 
                class="w-full rounded-xl border border-border bg-white/70 px-4 py-3 outline-none focus:ring-2 focus:ring-info/30" 
                value="${escapeHtml(d.EMAIL_DESTINATARIO || '')}">
            </div>
            
            <div class="space-y-2">
              <label class="text-sm font-medium">Nome Destinatário (opcional)</label>
              <input id="destNome" type="text" 
                class="w-full rounded-xl border border-border bg-white/70 px-4 py-3 outline-none focus:ring-2 focus:ring-info/30" 
                value="${escapeHtml(d.NOME_DESTINATARIO || '')}">
            </div>
            
            <p id="destErro" class="text-sm text-destructive hidden whitespace-pre-line"></p>
            
            <div class="pt-2 flex flex-col sm:flex-row gap-3">
              <button id="btnSalvarDest" type="submit" class="sm:flex-1 rounded-xl bg-info text-white px-4 py-3 font-medium hover:opacity-90 transition-all">
                Salvar
              </button>
              <button id="btnCancelarDest" type="button" class="sm:flex-1 rounded-xl border border-border bg-white/50 px-4 py-3 font-medium hover:bg-white/70 transition-all">
                Cancelar
              </button>
            </div>
          </form>
        </div>
      </div>
    </div>
  `;
  
  document.body.appendChild(modal);

  // Eventos
  const fechar = removerModalDestinatario;
  overlay.addEventListener('click', fechar);
  document.getElementById('btnFecharDest')?.addEventListener('click', fechar);
  document.getElementById('btnCancelarDest')?.addEventListener('click', fechar);

  // Submit
  document.getElementById('formDestinatario')?.addEventListener('submit', async e => {
    e.preventDefault();
    
    const idRemetente = Number(document.getElementById('destRemetenteId').value);
    const email = document.getElementById('destEmail').value.trim();
    const nome = document.getElementById('destNome').value.trim();
    
    if (!idRemetente || !email) {
      setDestErro('Remetente e email são obrigatórios.');
      return;
    }
    
    try {
      const btn = document.getElementById('btnSalvarDest');
      btn.disabled = true;
      btn.classList.add('opacity-70');
      
      if (isEdit) {
        await apiSend(`/api/emails/destinatarios/${d.ID}`, 'PUT', { 
          ID_REMETENTE: idRemetente, 
          EMAIL_DESTINATARIO: email, 
          NOME_DESTINATARIO: nome || null 
        });
      } else {
        await apiSend('/api/emails/destinatarios', 'POST', { 
          ID_REMETENTE: idRemetente, 
          EMAIL_DESTINATARIO: email, 
          NOME_DESTINATARIO: nome || null 
        });
      }
      
      removerModalDestinatario();
      await carregarEmailsAutomaticos();
      
    } catch (err) {
      setDestErro(err.message);
    } finally {
      const btn = document.getElementById('btnSalvarDest');
      btn.disabled = false;
      btn.classList.remove('opacity-70');
    }
  });
}

function setDestErro(msg) {
  const el = document.getElementById('destErro');
  if (!el) return;
  if (!msg) {
    el.classList.add('hidden');
    return;
  }
  el.textContent = msg;
  el.classList.remove('hidden');
}

function getApiBaseGestaoUsuarios() {
  let raw =
    sessionStorage.getItem('api_base') ||
    sessionStorage.getItem('apibase') ||
    '';

  raw = String(raw || '').trim();
  if (!raw) return '';

  if (!/^https?:\/\//i.test(raw)) {
    raw = `https://${raw}`;
  }

  try {
    const url = new URL(raw);
    return url.href.replace(/\/+$/, '');
  } catch (err) {
    console.error('getApiBaseGestaoUsuarios base inválida:', raw, err);
    return '';
  }
}

function absUrlFromApiGestaoUsuarios(relOrAbs, apiBase) {
  const s = String(relOrAbs || '').trim();
  if (!s) return '';
  if (/^https?:\/\//i.test(s)) return s;

  const base = String(apiBase || getApiBaseGestaoUsuarios() || '').trim();

  if (!base) {
    return s;
  }

  try {
    if (s.startsWith('/foto-usuario/')) {
      return `${base}/publicidade${s}`.replace(/([^:]\/)\/+/g, '$1');
    }

    if (s.startsWith('/publicidade/')) {
      return `${base}${s}`.replace(/([^:]\/)\/+/g, '$1');
    }

    return new URL(s, `${base}/`).href;
  } catch (err) {
    console.error('Erro em absUrlFromApiGestaoUsuarios:', err);
    return s;
  }
}

function montarUrlApiGestao(path) {
  const base = getApiBaseGestaoUsuarios();
  if (!base) throw new Error('API base inválida ou não configurada.');
  return `${base}${path.startsWith('/') ? '' : '/'}${path}`;
}

// ===== Controle de Estoque ====== //
// ================================ //

let cacheEstoqueEscritorio = [];

function setEstoqueEscritorioErro(msg) {
  const el = document.getElementById('estoqueEscritorioErro');
  if (!el) return;

  if (!msg) {
    el.textContent = '';
    el.classList.add('hidden');
    return;
  }

  el.textContent = msg;
  el.classList.remove('hidden');
}

function setAbaEstoque(nome) {
  const abaEscritorio = document.getElementById('abaEstoqueEscritorio');
  const abaFazenda = document.getElementById('abaEstoqueFazenda');
  const painelEscritorio = document.getElementById('painelEstoqueEscritorio');
  const painelFazenda = document.getElementById('painelEstoqueFazenda');

  const fazendaAtiva = nome === 'fazenda';

  if (abaEscritorio) {
    abaEscritorio.setAttribute('aria-selected', fazendaAtiva ? 'false' : 'true');
    abaEscritorio.className = fazendaAtiva
      ? 'px-4 py-2 rounded-xl text-sm font-medium border border-border bg-white/40 text-muted-foreground hover:bg-white/70 transition-all'
      : 'px-4 py-2 rounded-xl text-sm font-medium border border-border bg-white text-foreground shadow-sm transition-all';
  }

  if (abaFazenda) {
    abaFazenda.setAttribute('aria-selected', fazendaAtiva ? 'true' : 'false');
    abaFazenda.className = fazendaAtiva
      ? 'px-4 py-2 rounded-xl text-sm font-medium border border-border bg-white text-foreground shadow-sm transition-all'
      : 'px-4 py-2 rounded-xl text-sm font-medium border border-border bg-white/40 text-muted-foreground hover:bg-white/70 transition-all';
  }

  if (painelEscritorio) painelEscritorio.classList.toggle('hidden', fazendaAtiva);
  if (painelFazenda) painelFazenda.classList.toggle('hidden', !fazendaAtiva);
}

function rowEstoqueEscritorio(item) {
  const codigo =
    item.CODIGOITEM ??
    item.codigoitem ??
    item.CODIGO_ITEM ??
    item.codigo_item ??
    item.CODIGO ??
    item.codigo ??
    '—';

  const qtdDisponivel =
    item.QTDDISPONIVEL ??
    item.qtddisponivel ??
    item.QTD_DISPONIVEL ??
    item.qtd_disponivel ??
    item.SALDO ??
    item.saldo ??
    0;

  const qtdPedido =
    item.QTDEMPEDIDO ??
    item.qtdempedido ??
    item.QTD_EM_PEDIDO ??
    item.qtd_em_pedido ??
    0;

  const id = item.ID ?? item.id ?? codigo;

  return `
    <tr class="border-b border-border last:border-b-0">
      <td class="px-4 py-3">${escapeHtml(codigo)}</td>
      <td class="px-4 py-3">${escapeHtml(String(qtdDisponivel))}</td>
      <td class="px-4 py-3">${escapeHtml(String(qtdPedido))}</td>
      <td class="px-4 py-3">
        <div class="flex justify-end gap-2">
          <button
            class="btnHistoricoEstoque w-10 h-10 rounded-xl border border-border bg-white/60 hover:bg-white/90 transition-all"
            data-id="${escapeHtml(String(id))}"
            aria-label="Histórico"
            title="Histórico"
          >
            <i class="fas fa-clock-rotate-left" aria-hidden="true"></i>
          </button>

          <button
            class="btnVinculoEstoque w-10 h-10 rounded-xl border border-border bg-white/60 hover:bg-white/90 transition-all"
            data-id="${escapeHtml(String(id))}"
            aria-label="Cadastrar vínculo"
            title="Cadastrar vínculo"
          >
            <i class="fas fa-link" aria-hidden="true"></i>
          </button>
        </div>
      </td>
    </tr>
  `;
}

function renderEstoqueEscritorio(items = []) {
  const tbody = document.getElementById('tbodyEstoqueEscritorio');
  if (!tbody) return;

  if (!Array.isArray(items) || !items.length) {
    tbody.innerHTML = `
      <tr>
        <td colspan="4" class="px-4 py-6 text-sm text-muted-foreground">
          Nenhum material do escritório encontrado.
        </td>
      </tr>
    `;
    return;
  }

  tbody.innerHTML = items.map(rowEstoqueEscritorio).join('');
}

async function carregarControleEstoque() {
  showPage('inventory-control');
  setAbaEstoque('escritorio');
  setEstoqueEscritorioErro('');

  const tbody = document.getElementById('tbodyEstoqueEscritorio');
  if (tbody) {
    tbody.innerHTML = `
      <tr>
        <td colspan="4" class="px-4 py-6 text-sm text-muted-foreground">
          Carregando materiais...
        </td>
      </tr>
    `;
  }

  try {
    // Placeholder temporário até definir API real
    cacheEstoqueEscritorio = [];

    renderEstoqueEscritorio(cacheEstoqueEscritorio);
  } catch (err) {
    setEstoqueEscritorioErro(err?.message || 'Erro ao carregar estoque.');
    renderEstoqueEscritorio([]);
  }
}

document.addEventListener('click', (e) => {
  const item = e.target.closest('.menu-item[data-page]');
  if (!item) return;

  const page = item.dataset.page;
  if (page !== 'inventory-control') return;

  carregarControleEstoque();
});

document.addEventListener('click', (e) => {
  if (e.target.closest('#abaEstoqueEscritorio')) {
    setAbaEstoque('escritorio');
    return;
  }

  if (e.target.closest('#abaEstoqueFazenda')) {
    setAbaEstoque('fazenda');
    return;
  }
});

// ===== Importar PDF =======//

function normalizarTextoNfePdf(texto) {
  return String(texto || '')
    .replace(/\r/g, '')
    .replace(/[ \t]+/g, ' ')
    .replace(/\n{2,}/g, '\n')
    .trim();
}

function validarPdfComoNFe(texto) {
  const t = normalizarTextoNfePdf(texto).toUpperCase();

  const sinais = [
    'DANFE',
    'DOCUMENTO AUXILIAR',
    'NOTA FISCAL',
    'ELETRONICA',
    'CHAVE DE ACESSO',
    'DESTINATÁRIO/REMETENTE',
    'DADOS DO PRODUTO/SERVIÇO'
  ];

  return sinais.filter(s => t.includes(s)).length >= 4;
}

function extrairCampo(regex, texto, fallback = '') {
  const m = texto.match(regex);
  return m?.[1]?.trim?.() || fallback;
}

function limparNumero(str = '') {
  return String(str).replace(/[^\d]/g, '');
}

function extrairCnpjEmitenteDoTexto(textoOriginal) {
  const texto = normalizarTextoNfePdf(textoOriginal);

  const blocoEmitente =
    extrairCampo(/IDENTIFICAÇÃO DO EMITENTE\s+([\s\S]+?)\s+DANFE/i, texto) ||
    extrairCampo(/IDENTIFICAÇÃO DO EMITENTE\s+([\s\S]+?)\s+CHAVE DE ACESSO/i, texto) ||
    extrairCampo(/RECEBEMOS DE\s+([\s\S]+?)\s+OS PRODUTOS/i, texto) ||
    '';

  const cnpjFormatado =
    extrairCampo(/CNPJ[: ]*([0-9]{2}\.[0-9]{3}\.[0-9]{3}\/[0-9]{4}-[0-9]{2})/i, blocoEmitente) ||
    extrairCampo(/CNPJ[: ]*([0-9]{2}\.[0-9]{3}\.[0-9]{3}\/[0-9]{4}-[0-9]{2})/i, texto);

  if (cnpjFormatado) return cnpjFormatado;

  const cnpjNumericoNoBloco = extrairCampo(/CNPJ[: ]*([0-9]{14})/i, blocoEmitente);
  if (cnpjNumericoNoBloco) return cnpjNumericoNoBloco;

  const matchQualquerCnpjNoBloco = blocoEmitente.match(/\d{2}\.\d{3}\.\d{3}\/\d{4}-\d{2}|\d{14}/);
  if (matchQualquerCnpjNoBloco) return matchQualquerCnpjNoBloco[0];

  const matchQualquerCnpjNoTexto = texto.match(/\d{2}\.\d{3}\.\d{3}\/\d{4}-\d{2}|\d{14}/);
  if (matchQualquerCnpjNoTexto) return matchQualquerCnpjNoTexto[0];

  return '';
}

function detectarModeloDocumentoFiscal(textoOriginal) {
  const texto = normalizarTextoNfePdf(textoOriginal);
  const t = texto.toUpperCase();

  if (
    t.includes('SERVIÇOS DE COMUNICAÇÃO ELETRÔNICA') ||
    t.includes('NOTA FISCAL FATURA') ||
    t.includes('/NFCOM') ||
    t.includes('NFCOM')
  ) {
    return 'nfcom';
  }

  if (
    t.includes('DANFE') ||
    t.includes('NOTA FISCAL ELETRÔNICA') ||
    t.includes('CHAVE DE ACESSO')
  ) {
    return 'nfe';
  }

  return 'desconhecido';
}

function validarPdfComoDocumentoFiscal(textoOriginal) {
  return detectarModeloDocumentoFiscal(textoOriginal) !== 'desconhecido';
}

function parseDocumentoFiscalPdf(textoOriginal) {
  const modelo = detectarModeloDocumentoFiscal(textoOriginal);

  if (modelo === 'nfcom') {
    return parseNFComPdfTexto(textoOriginal);
  }

  if (modelo === 'nfe') {
    return parseNfePdfTextoRobusto(textoOriginal);
  }

  return {
    modelo: 'desconhecido',
    textoOriginal: normalizarTextoNfePdf(textoOriginal),
    itens: []
  };
}

function extrairRazaoSocialEmitente(textoOriginal) {
  const texto = normalizarTextoNfePdf(textoOriginal);

  const blocoEmitente =
    extrairCampo(/IDENTIFICAÇÃO DO EMITENTE\s+([\s\S]+?)\s+DANFE/i, texto) ||
    extrairCampo(/IDENTIFICAÇÃO DO EMITENTE\s+([\s\S]+?)\s+CHAVE DE ACESSO/i, texto) ||
    extrairCampo(/RECEBEMOS DE\s+(.+?)\s+OS PRODUTOS/i, texto) ||
    '';

  if (!blocoEmitente) return '';

  const blocoLimpo = blocoEmitente
    .replace(/\s+/g, ' ')
    .trim();

  const marcadoresEndereco = [
    ' R.',
    ' RUA ',
    ' AV.',
    ' AV ',
    ' AVENIDA ',
    ' ALAMEDA ',
    ' TRAVESSA ',
    ' TV. ',
    ' ESTRADA ',
    ' ROD.',
    ' RODOVIA ',
    ' FAZENDA ',
    ' SITIO ',
    ' SÍTIO ',
    ' CHACARA ',
    ' CHÁCARA ',
    ' PRAÇA ',
    ' PRACA ',
    ' QD ',
    ' QUADRA ',
    ' LOTE ',
    ' CEP ',
    ' FONE ',
    ' FONE/FAX ',
    ' CNPJ '
  ];

  let fim = blocoLimpo.length;

  for (const marcador of marcadoresEndereco) {
    const idx = blocoLimpo.toUpperCase().indexOf(marcador.toUpperCase());
    if (idx > 0 && idx < fim) fim = idx;
  }

  return blocoLimpo.slice(0, fim).trim().toUpperCase();
}


function parseNfePdfTextoRobusto(textoOriginal) {
  const texto = normalizarTextoNfePdf(textoOriginal);

  const emitente = extrairRazaoSocialEmitente(texto);

  const numeroNota =
    extrairCampo(/NF-e\s+N[º°.o]*\s*([0-9.\-\/]+)/i, texto) ||
    extrairCampo(/N[º°.o]*\s*([0-9.\-\/]+)\s+S[ÉE]RIE/i, texto) ||
    extrairCampo(/N[º°.o]*\.\s*([0-9.\-\/]+)/i, texto);

  const serie =
    extrairCampo(/S[ÉE]RIE[: ]\s*([0-9]+)/i, texto) ||
    extrairCampo(/S[ée]rie\s*([0-9]+)/i, texto);

  const chaveAcesso = extrairCampo(/CHAVE DE ACESSO\s+([0-9\s]{44,80})/i, texto)
    .replace(/\s+/g, '');

  const dataEmissao =
    extrairCampo(/DATA DA EMISS[ÃA]O\s+([0-9]{2}\/[0-9]{2}\/[0-9]{4})/i, texto) ||
    extrairCampo(/DATA DE EMISS[ÃA]O\s+([0-9]{2}\/[0-9]{2}\/[0-9]{4})/i, texto);

  const naturezaOperacao = extrairCampo(
    /NATUREZA DA OPERAÇÃO\s+([\s\S]+?)\s+PROTOCOLO DE AUTORIZAÇÃO DE USO/i,
    texto
  ).replace(/\n+/g, ' ').trim();

  const emitenteCnpj = extrairCnpjEmitenteDoTexto(texto);

  const destinatario = extrairCampo(
    /DESTINAT[ÁA]RIO\s*\/\s*REMETENTE\s+NOME\s*\/\s*RAZ[ÃA]O SOCIAL\s+([\s\S]+?)\s+CNPJ\s*\/\s*CPF/i,
    texto
  ).replace(/\n+/g, ' ').trim();

  const destinatarioCnpj = extrairCampo(
    /DESTINAT[ÁA]RIO\s*\/\s*REMETENTE[\s\S]*?CNPJ\s*\/\s*CPF\s+([0-9.\-\/]+)/i,
    texto
  );

  const enderecoDestinatario = extrairCampo(
    /DATA DA EMISS[ÃA]O\s+[0-9]{2}\/[0-9]{2}\/[0-9]{4}\s+ENDEREÇO\s+([\s\S]+?)\s+BAIRRO\s*\/\s*DISTRITO/i,
    texto
  ).replace(/\n+/g, ' ').trim();

  const bairroDestinatario = extrairCampo(
    /BAIRRO\s*\/\s*DISTRITO\s+([\s\S]+?)\s+CEP/i,
    texto
  ).replace(/\n+/g, ' ').trim();

  const municipioDestinatario = extrairCampo(
    /MUNIC[ÍI]PIO\s+([\s\S]+?)\s+UF/i,
    texto
  ).replace(/\n+/g, ' ').trim();

  const ufDestinatario = extrairCampo(/UF\s+([A-Z]{2})\s+FONE/i, texto) ||
    extrairCampo(/MUNIC[ÍI]PIO[\s\S]+?UF\s+([A-Z]{2})/i, texto);

  const valorTotalNota =
    extrairCampo(/V\. TOTAL DA NOTA\s+([\d.,]+)/i, texto) ||
    extrairCampo(/VALOR TOTAL DA NOTA\s+([\d.,]+)/i, texto);

  const itens = extrairItensNfePdfRobusto(texto);

  return {
    modelo: 'nfe',
    emitente,
    emitenteCnpj,
    numeroNota,
    serie,
    chaveAcesso,
    dataEmissao,
    naturezaOperacao,
    destinatario,
    destinatarioCnpj,
    enderecoDestinatario,
    bairroDestinatario,
    municipioDestinatario,
    ufDestinatario,
    valorTotalNota,
    itens,
    textoOriginal: texto
  };
}

function extrairItensNfePdfRobusto(textoOriginal) {
  const texto = normalizarTextoNfePdf(textoOriginal);
  const itens = [];

  const bloco = extrairCampo(
    /DADOS DOS PRODUTOS?\s*\/\s*SERVIÇOS\s+([\s\S]+?)\s+DADOS ADICIONAIS/i,
    texto
  ) || extrairCampo(
    /DADOS DO PRODUTO\s*\/\s*SERVIÇO\s+([\s\S]+?)\s+CALCULO DO ISSQN/i,
    texto
  ) || '';


  if (!bloco) {
    console.warn('[NFE PDF] Nenhum bloco de itens encontrado.');
    return itens;
  }

  const blocoLinear = bloco
    .replace(/\n+/g, ' ')
    .replace(/\s{2,}/g, ' ')
    .trim();


  const regexItem = /(\d{3,20})\s+(.+?)\s+(\d{8})\s+(\d{4})\s+(\d{4})\s+([A-Z]{1,6})\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)/g;

  let match;
  let contador = 0;

  while ((match = regexItem.exec(blocoLinear)) !== null) {
    contador++;

    const codigo = match[1] || '';
    let descricao = (match[2] || '').trim();
    const ncm = match[3] || '';
    const cst = match[4] || '';
    const cfop = match[5] || '';
    const unidade = match[6] || '';
    const quantidade = match[7] || '';
    const valorUnitario = match[8] || '';
    const valorTotal = match[9] || '';

    descricao = descricao
      .replace(/^\s*0,00\s+0,00\s+/i, '')
      .replace(/\s+Retido na compra:.*$/i, '')
      .replace(/\s+VALOR ICMS ST=0,00.*$/i, '')
      .replace(/\s+/g, ' ')
      .trim();

    const item = {
      codigo,
      descricao,
      ncm,
      cst,
      cfop,
      unidade,
      quantidade,
      valorUnitario,
      valorTotal
    };


    itens.push(item);
  }

  return itens;
}


function parseNFComPdfTexto(textoOriginal) {
  const texto = normalizarTextoNfePdf(textoOriginal);

  const emitente =
    extrairCampo(/DOCUMENTO AUXILIAR DA NOTA FISCAL FATURA DE SERVIÇOS DE COMUNICAÇÃO ELETRÔNICA\s+([\s\S]+?)\s+CNPJ:/i, texto)
      .replace(/\n+/g, ' ')
      .trim();

  const emitenteCnpj = extrairCampo(/CNPJ:\s*([0-9.\-\/]+)/i, texto);

  const numeroNota =
    extrairCampo(/NOTA FISCAL FATURA No\.\s*([0-9.\-\/]+)/i, texto) ||
    extrairCampo(/NOTA FISCAL FATURA N[ºo]\.?\s*([0-9.\-\/]+)/i, texto);

  const serie = extrairCampo(/S[ÉE]RIE:\s*([0-9]+)/i, texto);

  const chaveAcesso = extrairCampo(/CHAVE DE ACESSO:\s*([0-9\s]{44,80})/i, texto)
    .replace(/\s+/g, '');

  const dataEmissao = extrairCampo(/DATA DE EMISSÃO:\s*([0-9]{2}\/[0-9]{2}\/[0-9]{4})/i, texto);

  const destinatarioCpfCnpj = extrairCampo(/CNPJ\/CPF:\s*([0-9.\-\/]+)/i, texto);

  const destinatario = extrairCampo(
    /BA\s+([A-ZÁÂÃÀÉÊÍÓÔÕÚÇa-záâãàéêíóôõúç ]+)\s+NOTA FISCAL FATURA/i,
    texto
  ).replace(/\n+/g, ' ').trim();

  const valorTotalNota =
    extrairCampo(/TOTAL A PAGAR:\s*([\d.,]+)/i, texto) ||
    extrairCampo(/VALOR TOTAL NFF\s+([\d.,]+)/i, texto);

  const itens = extrairItensNFComPdf(texto);

  return {
    modelo: 'nfcom',
    emitente,
    emitenteCnpj,
    numeroNota,
    serie,
    chaveAcesso,
    dataEmissao,
    destinatario,
    destinatarioCnpj: destinatarioCpfCnpj,
    enderecoDestinatario: '',
    municipioDestinatario: '',
    ufDestinatario: '',
    valorTotalNota,
    itens,
    textoOriginal: texto
  };
}

function extrairItensNFComPdf(textoOriginal) {
  const texto = normalizarTextoNfePdf(textoOriginal);
  const itens = [];

  const bloco = extrairCampo(
    /ITENS DA FATURA\s+([\s\S]+?)\s+INFORMAÇÕES COMPLEMENTARES/i,
    texto,
    ''
  );

  if (!bloco) return itens;

  const linhas = bloco
    .split('\n')
    .map(l => l.trim())
    .filter(Boolean);

  for (let i = 0; i < linhas.length; i++) {
    const descricao = linhas[i];
    const prox = linhas.slice(i + 1, i + 8);

    if (
      descricao &&
      !/^(UN|QUANT|PREÇO UNIT|VALOR TOTAL|PIS\/COFINS|BC ICMS|ALÍQ|VALOR ICMS)$/i.test(descricao) &&
      prox.length >= 7 &&
      /^[A-Z0-9]{1,6}$/.test(prox[0]) &&
      /^[\d.,]+$/.test(prox[1]) &&
      /^[\d.,]+$/.test(prox[2]) &&
      /^[\d.,]+$/.test(prox[3])
    ) {
      itens.push({
        codigo: '',
        descricao,
        unidade: prox[0],
        quantidade: prox[1],
        valorUnitario: prox[2],
        valorTotal: prox[3],
        pisCofins: prox[4] || '',
        bcIcms: prox[5] || '',
        aliquotaIcms: prox[6] || '',
        valorIcms: prox[7] || ''
      });
    }
  }

  return itens;
}

function parseNfePdfTexto(textoOriginal) {
  const texto = normalizarTextoNfePdf(textoOriginal);

  const emitente = extrairCampo(
    /RECEBEMOS DE\s+(.+?)\s+OS PRODUTOS E SERVIÇOS CONSTANTES/i,
    texto
  ) || extrairCampo(
    /IDENTIFICAÇÃO DE ASSINATURA DO RECEBEDOR\s+(.+?)\s+DANFE/i,
    texto
  );

  const numeroNota =
    extrairCampo(/NF-e\s+Nº\s*([0-9.\-\/]+)/i, texto) ||
    extrairCampo(/N[º°o]\s*([0-9.\-\/]+)\s+S[ÉE]RIE/i, texto);

  const serie =
    extrairCampo(/S[ÉE]RIE[: ]\s*([0-9]+)/i, texto) ||
    extrairCampo(/N[º°o]\s*[0-9.\-\/]+\s+S[ÉE]RIE[: ]?\s*([0-9]+)/i, texto);

  const chaveAcesso = extrairCampo(/CHAVE DE ACESSO\s+([0-9\s]{44,60})/i, texto)
    .replace(/\s+/g, '');

  const dataEmissao = extrairCampo(/DATA DE EMISS[ÃA]O\s+([0-9]{2}\/[0-9]{2}\/[0-9]{4})/i, texto);

  const emitenteCnpj = extrairCnpjEmitenteDoTexto(texto);

  const destinatario = extrairCampo(
    /DESTINAT[ÁA]RIO\/REMETENTE\s+NOME\/RAZ[ÃA]O SOCIAL\s+(.+?)\s+CNPJ\/CPF/i,
    texto
  );

  const destinatarioCnpj = extrairCampo(
    /DESTINAT[ÁA]RIO\/REMETENTE[\s\S]*?CNPJ\/CPF\s+([0-9.\-\/]+)/i,
    texto
  );

  const enderecoDestinatario = extrairCampo(
    /DATA DE EMISS[ÃA]O\s+[0-9]{2}\/[0-9]{2}\/[0-9]{4}\s+ENDERE[ÇC]O\s+(.+?)\s+BAIRRO\/DISTRITO/i,
    texto
  );

  const municipioDestinatario = extrairCampo(
    /MUNIC[ÍI]PIO\s+(.+?)\s+FONE\/FAX/i,
    texto
  );

  const ufDestinatario = extrairCampo(/FONE\/FAX\s+UF\s+([A-Z]{2})/i, texto);

  const itens = extrairItensNfePdf(texto);

  return {
    emitente,
    emitenteCnpj,
    numeroNota,
    serie,
    chaveAcesso,
    dataEmissao,
    destinatario,
    destinatarioCnpj,
    enderecoDestinatario,
    municipioDestinatario,
    ufDestinatario,
    itens,
    textoOriginal: texto
  };
}

function extrairItensNfePdf(textoOriginal) {
  const texto = normalizarTextoNfePdf(textoOriginal);
  const itens = [];

  const bloco = extrairCampo(
    /DADOS DO PRODUTO\/SERVIÇO\s+([\s\S]+?)\s+CALCULO DO ISSQN/i,
    texto,
    ''
  );

  if (!bloco) return itens;

  const blocoLinear = bloco
    .replace(/\n+/g, ' ')
    .replace(/\s{2,}/g, ' ')
    .trim();

  const regexItem = /(\d{6,20})\s+(.+?)\s+(\d{8})\s+(\d{4})\s+(\d{4})\s+([A-Z]{1,4})\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+[\d.,]+\s+[\d.,]+\s+[\d.,]+\s+[\d.,]+/g;

  let match;
  while ((match = regexItem.exec(blocoLinear)) !== null) {
    itens.push({
      codigo: match[1] || '',
      descricao: (match[2] || '').trim(),
      ncm: match[3] || '',
      cst: match[4] || '',
      cfop: match[5] || '',
      unidade: match[6] || '',
      quantidade: match[7] || '',
      valorUnitario: match[8] || '',
      valorTotal: match[9] || ''
    });
  }

  return itens;
}

function removerModalImportacaoNfe() {
  document.getElementById('nfePdfOverlay')?.remove();
  document.getElementById('nfePdfModal')?.remove();
}

async function abrirModalImportacaoNfe(dados) {
  removerModalImportacaoNfe();

  const overlay = document.createElement('div');
  overlay.id = 'nfePdfOverlay';
  overlay.className = 'fixed inset-0 bg-black/40 backdrop-blur-sm z-[110]';
  document.body.appendChild(overlay);

  const modal = document.createElement('div');
  modal.id = 'nfePdfModal';
  modal.className = 'fixed inset-0 z-[120]';

  function normalizarTexto(v) {
    return String(v ?? '').trim();
  }

  function escapeHtml(str) {
    return String(str ?? '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function obterUsuarioLogado() {
    return (
      sessionStorage.getItem('usuarioNome') ||
      sessionStorage.getItem('usuario') ||
      sessionStorage.getItem('userName') ||
      sessionStorage.getItem('usuarioEmail') ||
      ''
    );
  }

  async function apiJson(url, options = {}) {
    const resp = await fetch(url, options);
    const texto = await resp.text();
    let data = {};

    try {
      data = texto ? JSON.parse(texto) : {};
    } catch {
      data = {};
    }

    if (!resp.ok || !data.success) {
      throw new Error(
        data.error
          ? `${data.message || 'Erro na requisição.'} Detalhe: ${data.error}`
          : (data.message || `Erro na requisição: ${url}`)
      );
    }

    return data;
  }


  async function salvarAmarracaoItem(payload) {
    return apiJson(apiUrl('/api/estoque/produtos-amarracao'), {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });
  }

  async function confirmarImportacaoPdf(payload) {
    return apiJson(apiUrl('/api/estoque/importacao-pdf/confirmar'), {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });
  }

  async function criarProdutoRapido(item) {
    return new Promise(async (resolve, reject) => {
      try {
        const codigoResp = await apiJson(apiUrl('/api/estoque/produtos/proximo-codigo'));
        const codigoGerado = String(codigoResp?.codigo || '').trim();

        if (!codigoGerado) {
          throw new Error('Não foi possível gerar o código do produto.');
        }

        const overlayProduto = document.createElement('div');
        overlayProduto.id = 'produtoRapidoOverlay';
        overlayProduto.className = 'fixed inset-0 bg-black/40 backdrop-blur-sm z-[130]';

        const modalProduto = document.createElement('div');
        modalProduto.id = 'produtoRapidoModal';
        modalProduto.className = 'fixed inset-0 z-[140]';

        modalProduto.innerHTML = `
          <div class="w-full h-full overflow-auto">
            <div class="min-h-full flex items-center justify-center p-4">
              <div class="w-full max-w-xl rounded-2xl border border-border bg-background shadow-2xl overflow-hidden">
                <div class="px-6 py-5 border-b border-border flex items-start justify-between gap-4">
                  <div>
                    <h3 class="text-lg font-semibold text-foreground">Novo produto</h3>
                    <p class="text-sm text-muted-foreground">Cadastro rápido para vincular o item da nota.</p>
                  </div>

                  <button type="button" id="btnFecharProdutoRapido"
                    class="w-10 h-10 rounded-xl bg-white/60 border border-border hover:bg-white transition-all flex items-center justify-center">
                    <i class="fas fa-times" aria-hidden="true"></i>
                  </button>
                </div>

                <form id="formProdutoRapido" class="px-6 py-6 space-y-4">
                  <div>
                    <label class="block text-sm font-medium text-foreground mb-1">Código</label>
                    <input
                      id="produtoRapidoCodigo"
                      type="text"
                      class="w-full rounded-xl border border-border bg-white/80 px-3 py-2"
                      value="${escapeHtml(codigoGerado)}"
                      readonly
                    />
                  </div>

                  <div>
                    <label class="block text-sm font-medium text-foreground mb-1">Descrição</label>
                    <input
                      id="produtoRapidoDescricao"
                      type="text"
                      class="w-full rounded-xl border border-border bg-white/80 px-3 py-2"
                      value="${escapeHtml(normalizarTexto(item?._descricao).toUpperCase())}"
                      required
                    />
                  </div>

                  <div>
                    <label class="block text-sm font-medium text-foreground mb-1">Unidade</label>
                    <input
                      id="produtoRapidoUnidade"
                      type="text"
                      class="w-full rounded-xl border border-border bg-white/80 px-3 py-2"
                      value="${escapeHtml(normalizarTexto(item?._unidade).toUpperCase() || 'UN')}"
                      maxlength="10"
                      required
                    />
                  </div>

                  <div class="flex justify-end gap-2 pt-2">
                    <button
                      type="button"
                      id="btnCancelarProdutoRapido"
                      class="rounded-xl border border-border bg-white/60 px-4 py-3 text-sm font-medium hover:bg-white/90 transition-all">
                      Cancelar
                    </button>

                    <button
                      type="submit"
                      id="btnSalvarProdutoRapido"
                      class="rounded-xl bg-primary text-white px-4 py-3 text-sm font-medium hover:opacity-90 transition-all">
                      Salvar produto
                    </button>
                  </div>
                </form>
              </div>
            </div>
          </div>
        `;

        function fecharModalProduto() {
          overlayProduto.remove();
          modalProduto.remove();
        }

        function cancelar() {
          fecharModalProduto();
          resolve(null);
        }

        document.body.appendChild(overlayProduto);
        document.body.appendChild(modalProduto);

        overlayProduto.addEventListener('click', cancelar);
        modalProduto.querySelector('#btnFecharProdutoRapido')?.addEventListener('click', cancelar);
        modalProduto.querySelector('#btnCancelarProdutoRapido')?.addEventListener('click', cancelar);

        modalProduto.querySelector('#formProdutoRapido')?.addEventListener('submit', async (e) => {
          e.preventDefault();

          const btnSalvar = modalProduto.querySelector('#btnSalvarProdutoRapido');
          const codigo = normalizarTexto(modalProduto.querySelector('#produtoRapidoCodigo')?.value).toUpperCase();
          const descricao = normalizarTexto(modalProduto.querySelector('#produtoRapidoDescricao')?.value).toUpperCase();
          const unidade = normalizarTexto(modalProduto.querySelector('#produtoRapidoUnidade')?.value).toUpperCase();

          if (!descricao) {
            alert('Descrição é obrigatória.');
            return;
          }

          if (!unidade) {
            alert('Unidade é obrigatória.');
            return;
          }

          try {
            btnSalvar.disabled = true;
            btnSalvar.textContent = 'Salvando...';

            const data = await apiJson(apiUrl('/api/estoque/produtos'), {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({
                codigo,
                descricao,
                unidade
              })
            });

            fecharModalProduto();
            resolve(data.item || null);
          } catch (err) {
            btnSalvar.disabled = false;
            btnSalvar.textContent = 'Salvar produto';

            if ((err.message || '').includes('Já existe produto com esse código')) {
              alert('Esse código foi utilizado por outro usuário neste instante. Abra novamente o cadastro para gerar um novo código.');
              fecharModalProduto();
              resolve(null);
              return;
            }

            reject(err);
          }
        });
      } catch (err) {
        reject(err);
      }
    });
  }


  let produtosSistema = [];
  let fornecedor = dados?._validacao?.fornecedor || null;
  const itensValidacao = Array.isArray(dados?._validacao?.itens) ? dados._validacao.itens : [];

  try {
    produtosSistema = await carregarProdutosSistema();
  } catch (err) {
    console.error('Erro ao carregar produtos do sistema:', err);
    produtosSistema = [];
  }

  const mapaItensValidacao = new Map(
    itensValidacao.map(item => [String(item.codigo ?? '').trim(), item])
  );

  const itensTratados = Array.isArray(dados?.itens)
    ? dados.itens.map((item, index) => {
        const codigoItem = normalizarTexto(item?.codigo);
        const validado = mapaItensValidacao.get(codigoItem);
        const produtoVinculado = validado?.produto || null;

        return {
          ...item,
          _index: index,
          _codigo: codigoItem,
          _descricao: normalizarTexto(item?.descricao),
          _unidade: normalizarTexto(item?.unidade),
          _quantidade: normalizarTexto(item?.quantidade),
          _valorUnitario: normalizarTexto(item?.valorUnitario),
          _valorTotal: normalizarTexto(item?.valorTotal),
          _amarrado: !!validado?.vinculado,
          _produtoSelecionadoId: String(produtoVinculado?.ID ?? ''),
          _produtoSelecionadoCodigo: normalizarTexto(produtoVinculado?.CODIGO).toUpperCase(),
          _produtoSelecionadoDescricao: normalizarTexto(produtoVinculado?.DESCRICAO).toUpperCase()
        };
      })
    : [];

  function gerarOpcoesProdutos(produtoSelecionadoId = '') {
    return produtosSistema.map(prod => {
      const id = String(prod.id ?? prod.ID ?? '');
      const codigo = escapeHtml(String(prod.codigo ?? prod.CODIGO ?? ''));
      const descricao = escapeHtml(String(prod.descricao ?? prod.DESCRICAO ?? ''));
      const selected = String(produtoSelecionadoId) === id ? 'selected' : '';
      return `<option value="${id}" ${selected}>${codigo} - ${descricao}</option>`;
    }).join('');
  }

  function badgeVinculo(item) {
    if (item._amarrado) {
      return `<span class="inline-flex rounded-full border border-green-300 bg-green-50 px-3 py-1 text-xs font-medium text-green-700">Vinculado</span>`;
    }

    if (item._produtoSelecionadoId) {
      return `<span class="inline-flex rounded-full border border-blue-300 bg-blue-50 px-3 py-1 text-xs font-medium text-blue-700">Selecionado</span>`;
    }

    return `<span class="inline-flex rounded-full border border-amber-300 bg-amber-50 px-3 py-1 text-xs font-medium text-amber-700">Sem vínculo</span>`;
  }

  function renderizarLinhasItens() {

    if (!itensTratados.length) {
      return `
        <tr>
          <td colspan="9" class="px-4 py-6 text-sm text-muted-foreground text-center">
            Nenhum item identificado automaticamente no PDF.
          </td>
        </tr>
      `;
    }

    return itensTratados.map((item) => {
      const produtoJaVinculado = item._amarrado && item._produtoSelecionadoId;

      const htmlProdutoSistema = produtoJaVinculado
        ? `
          <div class="rounded-xl border border-green-200 bg-green-50 px-3 py-2 text-sm text-green-800">
            <div class="font-semibold">${escapeHtml(item._produtoSelecionadoCodigo || '')}</div>
            <div class="text-xs opacity-80">${escapeHtml(item._produtoSelecionadoDescricao || '')}</div>
          </div>
        `
        : `
          <select
            class="nfe-item-produto-select w-full rounded-xl border border-border bg-white/80 px-3 py-2"
            data-index="${item._index}"
          >
            <option value="">Selecione...</option>
            ${gerarOpcoesProdutos(item._produtoSelecionadoId)}
          </select>
        `;

      const htmlAcao = produtoJaVinculado
        ? `
          <span class="inline-flex rounded-xl border border-green-200 bg-green-50 px-3 py-2 text-xs font-medium text-green-700">
            Amarração encontrada
          </span>
        `
        : `
          <button
            type="button"
            class="btnNovoProdutoRapido rounded-xl border border-border bg-white/70 px-3 py-2 text-xs font-medium hover:bg-white"
            data-index="${item._index}"
          >
            Novo produto
          </button>
        `;

      return `
        <tr class="border-b border-border last:border-b-0 hover:bg-white/30 transition-colors">
          <td class="px-4 py-3 text-sm">${escapeHtml(item._codigo || '')}</td>
          <td class="px-4 py-3 text-sm">${escapeHtml(item._descricao || '')}</td>
          <td class="px-4 py-3 text-sm text-center">${escapeHtml(item._unidade || '')}</td>
          <td class="px-4 py-3 text-sm text-right">${escapeHtml(item._quantidade || '')}</td>
          <td class="px-4 py-3 text-sm text-right">${escapeHtml(item._valorUnitario || '')}</td>
          <td class="px-4 py-3 text-sm text-right font-semibold">${escapeHtml(item._valorTotal || '')}</td>
          <td class="px-4 py-3 text-sm">${badgeVinculo(item)}</td>
          <td class="px-4 py-3 text-sm min-w-[280px]">
            ${htmlProdutoSistema}
          </td>
          <td class="px-4 py-3 text-sm whitespace-nowrap">
            ${htmlAcao}
          </td>
        </tr>
      `;
    }).join('');
  }


  modal.innerHTML = `
    <div class="w-full h-full overflow-auto">
      <div class="min-h-full flex items-start justify-center p-4 md:p-8">
        <div class="w-full max-w-7xl mx-auto">
          <div class="glass rounded-2xl shadow-2xl border border-border overflow-hidden bg-background">
            <div class="px-6 py-5 border-b border-border flex items-start justify-between gap-4">
              <div>
                <h3 class="text-xl font-semibold text-foreground">Importação de nota fiscal PDF</h3>
                <p class="text-sm text-muted-foreground">Validação e leitura inicial do DANFE/NF-e.</p>
              </div>

              <button id="btnFecharModalNfePdf" type="button"
                class="w-10 h-10 rounded-xl bg-white/60 border border-border hover:bg-white transition-all flex items-center justify-center">
                <i class="fas fa-times" aria-hidden="true"></i>
              </button>
            </div>

            <div class="px-6 py-4 border-b border-border flex gap-2">
              <button id="nfePdfAbaDetalhes" type="button" class="px-4 py-2 rounded-xl text-sm font-medium border border-border bg-white text-foreground shadow-sm transition-all">
                <i class="fas fa-receipt mr-2"></i> Detalhes
              </button>
              <button id="nfePdfAbaItens" type="button" class="px-4 py-2 rounded-xl text-sm font-medium border border-border bg-white/60 text-muted-foreground hover:bg-white/90 transition-all">
                <i class="fas fa-list mr-2"></i> Itens (${itensTratados.length})
              </button>
            </div>

            <div class="px-6 py-6">
              <div id="nfePdfConteudoDetalhes" class="space-y-6">
                <div class="grid grid-cols-1 lg:grid-cols-2 gap-6">
                  <div class="rounded-2xl border border-border bg-white/50 p-5 shadow-sm space-y-4">
                    <div class="grid grid-cols-1 md:grid-cols-2 gap-3">
                      <div class="md:col-span-2 rounded-xl border border-border bg-white/70 px-4 py-3">
                        <div class="text-[11px] uppercase tracking-wide text-muted-foreground">Emitente</div>
                        <div class="mt-1 text-sm font-semibold text-foreground break-words">${escapeHtml(dados.emitente || '—')}</div>
                      </div>

                      <div class="rounded-xl border border-border bg-white/70 px-4 py-3">
                        <div class="text-[11px] uppercase tracking-wide text-muted-foreground">CNPJ do emitente</div>
                        <div class="mt-1 text-sm font-semibold text-foreground">${escapeHtml(dados.emitenteCnpj || '—')}</div>
                      </div>

                      <div class="rounded-xl border border-border bg-white/70 px-4 py-3">
                        <div class="text-[11px] uppercase tracking-wide text-muted-foreground">Número da nota</div>
                        <div class="mt-1 text-sm font-semibold text-foreground">${escapeHtml(dados.numeroNota || '—')}</div>
                      </div>

                      <div class="rounded-xl border border-border bg-white/70 px-4 py-3">
                        <div class="text-[11px] uppercase tracking-wide text-muted-foreground">Série</div>
                        <div class="mt-1 text-sm font-semibold text-foreground">${escapeHtml(dados.serie || '—')}</div>
                      </div>

                      <div class="rounded-xl border border-border bg-white/70 px-4 py-3">
                        <div class="text-[11px] uppercase tracking-wide text-muted-foreground">Data emissão</div>
                        <div class="mt-1 text-sm font-semibold text-foreground">${escapeHtml(dados.dataEmissao || '—')}</div>
                      </div>
                    </div>
                  </div>

                  <div class="rounded-2xl border border-border bg-white/50 p-5 shadow-sm space-y-4">
                    <div class="grid grid-cols-1 md:grid-cols-2 gap-3">
                      <div class="md:col-span-2 rounded-xl border border-border bg-white/70 px-4 py-3">
                        <div class="text-[11px] uppercase tracking-wide text-muted-foreground">Destinatário</div>
                        <div class="mt-1 text-sm font-semibold text-foreground break-words">${escapeHtml(dados.destinatario || '—')}</div>
                      </div>

                      <div class="rounded-xl border border-border bg-white/70 px-4 py-3">
                        <div class="text-[11px] uppercase tracking-wide text-muted-foreground">CNPJ/CPF</div>
                        <div class="mt-1 text-sm font-semibold text-foreground">${escapeHtml(dados.destinatarioCnpj || '—')}</div>
                      </div>

                      <div class="rounded-xl border border-border bg-white/70 px-4 py-3">
                        <div class="text-[11px] uppercase tracking-wide text-muted-foreground">UF</div>
                        <div class="mt-1 text-sm font-semibold text-foreground">${escapeHtml(dados.ufDestinatario || '—')}</div>
                      </div>

                      <div class="md:col-span-2 rounded-xl border border-border bg-white/70 px-4 py-3">
                        <div class="text-[11px] uppercase tracking-wide text-muted-foreground">Fornecedor cadastrado</div>
                        <div class="mt-1 text-sm font-semibold text-foreground break-words">
                          ${fornecedor
                            ? escapeHtml(fornecedor.RAZAO_SOCIAL || fornecedor.razao_social || 'Fornecedor encontrado')
                            : 'Não encontrado automaticamente'}
                        </div>
                      </div>
                    </div>
                  </div>
                </div>
              </div>

              <div id="nfePdfConteudoItens" class="hidden space-y-4">
                <div class="rounded-2xl border border-border bg-white/40 overflow-hidden">
                  <div class="overflow-x-auto">
                    <table class="min-w-full text-sm">
                      <thead class="bg-white/50 border-b border-border sticky top-0">
                        <tr>
                          <th class="px-4 py-3 text-left font-semibold text-foreground">Código</th>
                          <th class="px-4 py-3 text-left font-semibold text-foreground">Descrição</th>
                          <th class="px-4 py-3 text-center font-semibold text-foreground">UN</th>
                          <th class="px-4 py-3 text-right font-semibold text-foreground">Qtd</th>
                          <th class="px-4 py-3 text-right font-semibold text-foreground">Vlr. Unit.</th>
                          <th class="px-4 py-3 text-right font-semibold text-foreground">Vlr. Total</th>
                          <th class="px-4 py-3 text-left font-semibold text-foreground">Status</th>
                          <th class="px-4 py-3 text-left font-semibold text-foreground">Produto sistema</th>
                          <th class="px-4 py-3 text-left font-semibold text-foreground">Ação</th>
                        </tr>
                      </thead>
                      <tbody id="nfePdfTabelaItensBody">
                        ${renderizarLinhasItens()}
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>
            </div>

            <div class="px-6 py-4 border-t border-border flex justify-end gap-2">
              <button id="btnImportarNotaFiscalPdf" type="button"
                class="rounded-xl bg-primary text-white px-4 py-3 text-sm font-medium hover:opacity-90 transition-all">
                Importar
              </button>

              <button id="btnFecharRodapeModalNfePdf" type="button"
                class="rounded-xl border border-border bg-white/60 px-4 py-3 text-sm font-medium hover:bg-white/90 transition-all">
                Fechar
              </button>
            </div>
          </div>
        </div>
      </div>
    </div>
  `;

  document.body.appendChild(modal);

  const btnDetalhes = document.getElementById('nfePdfAbaDetalhes');
  const btnItens = document.getElementById('nfePdfAbaItens');
  const conteudoDetalhes = document.getElementById('nfePdfConteudoDetalhes');
  const conteudoItens = document.getElementById('nfePdfConteudoItens');
  const tbodyItens = document.getElementById('nfePdfTabelaItensBody');
  const btnImportar = document.getElementById('btnImportarNotaFiscalPdf');

  function rerenderTabelaItens() {
    if (!tbodyItens) return;
    tbodyItens.innerHTML = renderizarLinhasItens();
    vincularEventosItens();
  }

  function mostrarAba(aba) {
    if (aba === 'detalhes') {
      conteudoDetalhes.classList.remove('hidden');
      conteudoItens.classList.add('hidden');
    } else {
      conteudoDetalhes.classList.add('hidden');
      conteudoItens.classList.remove('hidden');
    }
  }

  function fechar() {
    removerModalImportacaoNfe();
  }

  function vincularEventosItens() {
    document.querySelectorAll('.nfe-item-produto-select').forEach(select => {
      select.addEventListener('change', (e) => {
        const index = Number(e.currentTarget.dataset.index);
        const item = itensTratados.find(x => x._index === index);
        if (!item) return;

        const value = String(e.currentTarget.value || '');
        item._produtoSelecionadoId = value;

        const produto = produtosSistema.find(
          p => String(p.id ?? p.ID) === value
        );

        item._produtoSelecionadoCodigo = produto
          ? normalizarTexto(produto.codigo ?? produto.CODIGO).toUpperCase()
          : '';

        item._produtoSelecionadoDescricao = produto
          ? normalizarTexto(produto.descricao ?? produto.DESCRICAO).toUpperCase()
          : '';

        rerenderTabelaItens();
      });
    });

    document.querySelectorAll('.btnNovoProdutoRapido').forEach(btn => {
      btn.addEventListener('click', async (e) => {
        const index = Number(e.currentTarget.dataset.index);
        const item = itensTratados.find(x => x._index === index);
        if (!item) return;

        try {
          const produtoCriado = await criarProdutoRapido(item);
          if (!produtoCriado) return;

          produtosSistema.push(produtoCriado);

          item._produtoSelecionadoId = String(produtoCriado.id ?? produtoCriado.ID ?? '');
          item._produtoSelecionadoCodigo = normalizarTexto(produtoCriado.codigo ?? produtoCriado.CODIGO).toUpperCase();
          item._produtoSelecionadoDescricao = normalizarTexto(produtoCriado.descricao ?? produtoCriado.DESCRICAO).toUpperCase();

          rerenderTabelaItens();
        } catch (err) {
          console.error(err);
          alert(err.message || 'Erro ao criar produto.');
        }
      });
    });
  }

  btnDetalhes?.addEventListener('click', () => mostrarAba('detalhes'));
  btnItens?.addEventListener('click', () => mostrarAba('itens'));
  overlay.addEventListener('click', fechar);
  document.getElementById('btnFecharModalNfePdf')?.addEventListener('click', fechar);
  document.getElementById('btnFecharRodapeModalNfePdf')?.addEventListener('click', fechar);

  btnImportar?.addEventListener('click', async () => {
    try {
      btnImportar.disabled = true;
      btnImportar.textContent = 'Importando...';

      if (!itensTratados.length) {
        throw new Error('Nenhum item encontrado para importar.');
      }

      for (const item of itensTratados) {
        if (!item._amarrado && !item._produtoSelecionadoId) {
          mostrarAba('itens');
          throw new Error(`Selecione o produto do sistema para o item ${item._codigo || item._descricao}.`);
        }
      }

      const idFornecedor = Number(fornecedor?.ID ?? fornecedor?.id ?? 0);

      for (const item of itensTratados) {
        if (item._amarrado) continue;

        if (idFornecedor) {
          await salvarAmarracaoItem({
            id_fornecedor: idFornecedor,
            cod_produto_nf: item._codigo,
            descricao_produto_nf: item._descricao,
            id_produto: Number(item._produtoSelecionadoId)
          });
        }

        item._amarrado = true;
      }

      const payloadImportacao = {
        emitente: dados.emitente || '',
        emitenteCnpj: dados.emitenteCnpj || '',
        destinatarioCnpj: dados.destinatarioCnpj || '',
        numeroNota: dados.numeroNota || '',
        serie: dados.serie || '',
        dataEmissao: dados.dataEmissao || '',
        usuarioRegistro: obterUsuarioLogado(),
        itens: itensTratados.map(item => ({
          codigo: item._codigo,
          descricao: item._descricao,
          unidade: item._unidade,
          quantidade: item._quantidade,
          valorUnitario: item._valorUnitario,
          valorTotal: item._valorTotal,
          id_produto: Number(item._produtoSelecionadoId),
          cod_produto_sistema: item._produtoSelecionadoCodigo,
          descricao_produto_sistema: item._produtoSelecionadoDescricao
        }))
      };

      await confirmarImportacaoPdf(payloadImportacao);

      alert('Importação realizada com sucesso.');
      fechar();

      if (typeof carregarEstoque === 'function') {
        try { await carregarEstoque(); } catch {}
      }
    } catch (err) {
      console.error('Erro ao importar nota PDF:', err);
      alert(err.message || 'Erro ao importar nota.');
    } finally {
      btnImportar.disabled = false;
      btnImportar.textContent = 'Importar';
    }
  });

  vincularEventosItens();
}


async function extrairTextoPdfComPdfJs(file) {
  const pdfjs = ensurePdfJsLoaded();

  const buffer = await file.arrayBuffer();
  const pdf = await pdfjs.getDocument({ data: buffer }).promise;

  let texto = '';

  for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
    const page = await pdf.getPage(pageNum);
    const content = await page.getTextContent();
    const pageText = content.items.map(item => item.str).join(' ');
    texto += '\n' + pageText;
  }

  return texto;
}

async function processarImportacaoPdfEstoque(file) {
  if (!file) return;

  const isPdf = file.type === 'application/pdf' || /\.pdf$/i.test(file.name || '');
  if (!isPdf) {
    alert('Selecione um arquivo PDF válido.');
    return;
  }

  try {
    const texto = await extrairTextoPdfComPdfJs(file);

    if (!validarPdfComoDocumentoFiscal(texto)) {
      alert('O arquivo selecionado não parece ser um documento fiscal suportado.');
      return;
    }

    const dados = parseDocumentoFiscalPdf(texto);
    const validacao = await validarImportacaoAntesModal(dados);
    dados._validacao = validacao;

    await abrirModalImportacaoNfe(dados);
  } catch (err) {
    console.error('Erro ao processar PDF da nota fiscal:', err);
    alert('Erro ao ler o PDF: ' + (err?.message || err));
  }
}

document.addEventListener('click', (e) => {
  const btn = e.target.closest('#btnImportarPdfEstoque');
  if (!btn) return;

  document.getElementById('inputImportarPdfEstoque')?.click();
});

document.addEventListener('change', async (e) => {
  const input = e.target.closest('#inputImportarPdfEstoque');
  if (!input) return;

  const file = input.files?.[0];
  await processarImportacaoPdfEstoque(file);
  input.value = '';
});

function getPdfJsLibSafe() {
  return (
    window.pdfjsLib ||
    window['pdfjsLib'] ||
    window['pdfjs-dist/build/pdf'] ||
    window?.exports?.['pdfjs-dist/build/pdf'] ||
    null
  );
}

function ensurePdfJsLoaded() {
  const lib = getPdfJsLibSafe();

  if (!lib) {
    throw new Error('pdf.js não carregado na página.');
  }

  if (lib.GlobalWorkerOptions) {
    lib.GlobalWorkerOptions.workerSrc =
      'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.10.377/pdf.worker.min.js';
  }

  return lib;
}
// ==== Botão para Importar e vlidar dados de nota PDF ======//

async function validarImportacaoAntesModal(dados) {

  const response = await fetch(apiUrl('/api/estoque/importacao-pdf/validar'), {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(dados)
  });

  console.log('[IMPORTACAO PDF][VALIDAR][RESPOSTA STATUS]:', {
    ok: response.ok,
    status: response.status,
    statusText: response.statusText
  });

  const texto = await response.text();

  let result = {};

  try {
    result = texto ? JSON.parse(texto) : {};
  } catch (err) {
    result = {};
  }

  if (!response.ok || !result.success) {

    throw new Error(
      result.error
        ? `${result.message || 'Erro ao validar importação do PDF.'} Detalhe: ${result.error}`
        : (result.message || `Erro ao validar importação do PDF. HTTP ${response.status}`)
    );
  }
  return result;
}


async function carregarProdutosSistema() {
  const response = await fetch(apiUrl('/api/estoque/produtos'));
  const texto = await response.text();
  let result = {};

  try {
    result = texto ? JSON.parse(texto) : {};
  } catch {
    result = {};
  }

  if (!response.ok || !result.success) {
    throw new Error(result.message || 'Erro ao carregar produtos.');
  }

  return result.items || [];
}

function textoLivre(v) {
  return String(v ?? '').trim();
}

function normalizarDocumentoPDF(v) {
  return String(v ?? '').replace(/\D+/g, '').trim();
}

function parseDecimalBr(v) {
  const s = String(v ?? '').trim();
  if (!s) return 0;
  return Number(s.replace(/\./g, '').replace(',', '.')) || 0;
}

function dataBrParaMysql(v) {
  const s = String(v ?? '').trim();
  if (!s) return null;

  const m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (!m) return null;

  const [, dd, mm, yyyy] = m;
  return `${yyyy}-${mm}-${dd}`;
}



