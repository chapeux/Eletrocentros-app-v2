(function () {
  "use strict";

  /* ==========================================================================
     UTILITIES & HELPERS
     ========================================================================== */
  var $ = function (id) { return document.getElementById(id); };
  var root = document.documentElement;

  /* ==========================================================================
     APP STATE & AUTHENTICATION
     ========================================================================== */
  var state = {
    currentView: 'planejamento',
    isMaintenanceUnlocked: false,
    authPasswordDefault: 'admin',
    disciplinas: [],
    disciplina: null,
    campos: {},
    campo: null,
    dirty: null,
    historico: [],
    seletorData: [],
    seletorOriginal: [],
    seletorDirty: false,
    isSeletorActive: false,
    templateBlocksData: null,
    templateBlocksOriginal: null,
    templateBlocksDirty: false,
    isTemplatesActive: false,
    selectedTemplateScenario: 'container_solar_essw_mecanica'
  };

  /* ==========================================================================
     PYTHON / API BRIDGE SETUP
     ========================================================================== */
  var API_BASE_CANDIDATES = (function () {
    var candidates = [];
    if (window.location.protocol !== 'file:') {
      candidates.push(window.location.origin + '/api');
    }
    candidates.push('http://127.0.0.1:8000/api', 'http://localhost:8000/api');
    return candidates;
  })();

  function isPyWebviewAvailable() {
    return window.pywebview && window.pywebview.api;
  }

  function apiCall(path, options) {
    return apiTry(path, options, 0);
  }

  function apiTry(path, options, index) {
    var base = API_BASE_CANDIDATES[index];
    return fetch(base + path, options).then(function (resp) {
      if (!resp.ok) {
        return resp.json().catch(function () { return { detail: resp.statusText }; }).then(function (err) {
          var e = new Error(err.detail || 'Erro na API'); e.status = resp.status; e.base = base; throw e;
        });
      }
      return resp.json();
    }).catch(function (err) {
      if (index + 1 < API_BASE_CANDIDATES.length) {
        return apiTry(path, options, index + 1);
      }
      err.base = base;
      throw err;
    });
  }

  function applyTheme(theme) {
    root.setAttribute('data-theme', theme);
    try { localStorage.setItem('eletrocentros_theme', theme); } catch (e) { }
    var btn = $('themeToggle');
    if (btn) {
      var names = { dark: 'Tema: Escuro (Dark)', dim: 'Tema: Intermediário (Dim)', light: 'Tema: Claro (Light)' };
      btn.setAttribute('title', names[theme] || 'Alternar tema');
    }
  }

  // Restore saved theme on startup (default: 'dim')
  try {
    var savedTheme = localStorage.getItem('eletrocentros_theme');
    var activeTheme = (savedTheme && (savedTheme === 'dark' || savedTheme === 'dim' || savedTheme === 'light')) ? savedTheme : 'dim';
    applyTheme(activeTheme);
  } catch (e) {
    applyTheme('dim');
  }

  $('themeToggle').addEventListener('click', function () {
    var current = root.getAttribute('data-theme') || 'dim';
    var nextTheme = 'light';
    if (current === 'dim') nextTheme = 'light';
    else if (current === 'light') nextTheme = 'dark';
    else nextTheme = 'dim';
    applyTheme(nextTheme);
  });

  /* ==========================================================================
     VIEW SWITCHING & AUTHENTICATION MODAL LOGIC
     ========================================================================== */
  var tabBtnPlanejamento = $('tabBtnPlanejamento');
  var tabBtnResultados = $('tabBtnResultados');
  var tabBtnManutencao = $('tabBtnManutencao');
  var viewPlanejamento = $('view-planejamento');
  var viewResultados = $('view-resultados');
  var viewManutencao = $('view-manutencao');
  var authModalOverlay = $('authModalOverlay');
  var authPasswordInput = $('authPassword');
  var authError = $('authError');
  var modeChip = $('modeChip');
  var ringWrap = $('ringWrap');
  var btnHist = $('btnHist');
  var subTitleText = $('subTitleText');

  function switchView(targetView) {
    if (targetView === 'manutencao' && !state.isMaintenanceUnlocked) {
      openAuthModal();
      return;
    }

    state.currentView = targetView;
    if (tabBtnPlanejamento) tabBtnPlanejamento.classList.toggle('active', targetView === 'planejamento');
    if (tabBtnResultados) tabBtnResultados.classList.toggle('active', targetView === 'resultados');
    if (tabBtnManutencao) tabBtnManutencao.classList.toggle('active', targetView === 'manutencao');

    if (viewPlanejamento) viewPlanejamento.classList.toggle('hidden', targetView !== 'planejamento');
    if (viewResultados) viewResultados.classList.toggle('hidden', targetView !== 'resultados');
    if (viewManutencao) viewManutencao.classList.toggle('hidden', targetView !== 'manutencao');

    if (targetView === 'planejamento') {
      if (modeChip) modeChip.style.display = 'none';
      if (btnHist) btnHist.style.display = 'none';
      if ($('btnCarregar')) $('btnCarregar').style.display = 'flex';
      if (ringWrap) ringWrap.style.display = 'flex';
      if (subTitleText) subTitleText.textContent = 'PCP & PLANEJAMENTO';
    } else if (targetView === 'resultados') {
      if (modeChip) modeChip.style.display = 'none';
      if (btnHist) btnHist.style.display = 'none';
      if ($('btnCarregar')) $('btnCarregar').style.display = 'none';
      if (ringWrap) ringWrap.style.display = 'none';
      if (subTitleText) subTitleText.textContent = 'RESULTADO DO CÁLCULO';
    } else {
      if (modeChip) modeChip.style.display = 'none';
      if (btnHist) btnHist.style.display = 'flex';
      if ($('btnCarregar')) $('btnCarregar').style.display = 'none';
      if (ringWrap) ringWrap.style.display = 'none';
      if (subTitleText) subTitleText.textContent = 'REGRAS & PARÂMETROS';
      if (!state.disciplinas.length) {
        carregarDisciplinas();
      }
    }
  }

  if (tabBtnPlanejamento) tabBtnPlanejamento.addEventListener('click', function () { switchView('planejamento'); });
  if (tabBtnResultados) tabBtnResultados.addEventListener('click', function () { switchView('resultados'); });
  if (tabBtnManutencao) tabBtnManutencao.addEventListener('click', function () { switchView('manutencao'); });

  function openAuthModal() {
    authPasswordInput.value = '';
    authError.classList.remove('show');
    authModalOverlay.classList.add('open');
    setTimeout(function () { authPasswordInput.focus(); }, 100);
  }

  function closeAuthModal() {
    authModalOverlay.classList.remove('open');
  }

  $('btnCloseAuth').addEventListener('click', closeAuthModal);
  $('btnCancelAuth').addEventListener('click', closeAuthModal);

  $('btnTogglePwd').addEventListener('click', function () {
    var type = authPasswordInput.type === 'password' ? 'text' : 'password';
    authPasswordInput.type = type;
  });

  function verifyCredentials() {
    var pwd = authPasswordInput.value.trim();
    if (!pwd) {
      showAuthError('Por favor, informe a senha.');
      return;
    }

    // Check if pywebview API is connected or fallback to client check
    if (isPyWebviewAvailable()) {
      window.pywebview.api.verify_password(pwd).then(function (isValid) {
        if (isValid) {
          onAuthSuccess();
        } else {
          showAuthError('Senha incorreta. Tente novamente.');
        }
      }).catch(function () {
        showAuthError('Erro ao comunicar com backend Python.');
      });
    } else {
      // Local fallback check (accepts "admin" or "1234")
      if (pwd === 'admin' || pwd === '1234' || pwd === state.authPasswordDefault) {
        onAuthSuccess();
      } else {
        showAuthError('Senha incorreta. (Dica padrão: admin ou 1234)');
      }
    }
  }

  function showAuthError(msg) {
    $('authErrorMsg').textContent = msg;
    authError.classList.add('show');
  }

  function onAuthSuccess() {
    state.isMaintenanceUnlocked = true;
    tabBtnManutencao.classList.add('unlocked');
    closeAuthModal();
    showToast('Acesso concedido à aba de Regras!');
    switchView('manutencao');
  }

  $('btnSubmitAuth').addEventListener('click', verifyCredentials);
  authPasswordInput.addEventListener('keydown', function (e) {
    if (e.key === 'Enter') {
      e.preventDefault();
      verifyCredentials();
    }
  });

  /* ==========================================================================
     PLANEJAMENTO VIEW — CONFIG & FORM ENGINE
     ========================================================================== */
  var DEFAULT_CONFIG = {
    listas: {
      tipoestrutura: ['Móvel', 'Semimóvel', 'Modular', 'Fixo', 'Embarcado', 'Container Solar', 'Skid (mecânica)', 'Skid (com elétrica)', 'Pilotis', 'ESSW (mecânica)', 'ESSW (elétrica)', 'Serviço Engenharia'],
      planpin: ['WAU-ELETRO-08', 'WAU-ELETRO-09', 'WAU-ELETRO-04', 'Não aplicável'],
      tipomaq: ['Split', 'Wall Mounted', 'Roof Top', 'Não possui', 'Não aplicável'],
      incendio: ['Com combate', 'Com instalações', 'Somente infra', 'Não aplicável'],
      seguranca: ['CFTV', 'Controle Acesso', 'CFTV + Controle Acesso', 'Não possui', 'Não aplicável'],
      complexidade: ['Simples', 'Médio', 'Complexo', 'Não aplicável'],
      planejadorSel: ['MAGLIONI', 'MAURICIOFA', 'CAMILARM']
    },
    regras: {
      estruturasSemModulo: ['Container Solar', 'Skid (mecânica)', 'Skid (com elétrica)', 'Pilotis', 'ESSW (mecânica)', 'ESSW (elétrica)', 'Serviço Engenharia'],
      estruturasSemValorMec: ['ESSW (elétrica)', 'Serviço Engenharia'],
      estruturasSemValorEletr: ['ESSW (mecânica)', 'Pilotis', 'Skid (mecânica)']
    }
  };

  var CONFIG = JSON.parse(JSON.stringify(DEFAULT_CONFIG));
  var SELECTS = {};

  function loadExternalConfig() {
    if (isPyWebviewAvailable()) {
      window.pywebview.api.get_config().then(function (data) {
        if (data && data.listas) {
          CONFIG = data;
          buildAllSelectsFromConfig();
          recomputeForm();
        }
      }).catch(function (err) {
        console.warn('[Config] Erro ao carregar via pywebview:', err);
      });
    } else {
      fetch('config.json').then(function (resp) {
        if (resp.ok) return resp.json();
      }).then(function (data) {
        if (data && data.listas) {
          CONFIG = data;
          buildAllSelectsFromConfig();
          recomputeForm();
        }
      }).catch(function (err) {
        console.log('[Config] Usando configuração padrão embutida.');
      });
    }
  }

  function buildSelect(containerId, dataId, options, opts) {
    opts = opts || {};
    var container = $(containerId);
    if (!container) return;

    container.innerHTML =
      '<button type="button" class="csel-btn"><span class="lbl ph">' + (opts.placeholder || 'Selecione…') + '</span>' +
      '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4"><path d="M6 9l6 6 6-6"/></svg></button>' +
      '<div class="csel-panel" role="listbox"></div>';

    var panel = container.querySelector('.csel-panel');
    var btn = container.querySelector('.csel-btn');
    var placeholder = opts.placeholder || 'Selecione…';

    var allOpts = [{ value: '', label: placeholder }].concat(options.map(function (o) {
      return typeof o === 'string' ? { value: o, label: o } : o;
    }));

    allOpts.forEach(function (o) {
      var row = document.createElement('div');
      row.className = 'csel-opt';
      row.dataset.value = o.value;
      row.setAttribute('role', 'option');
      row.innerHTML = '<span>' + o.label + '</span><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><path d="M4 12l5 5L20 6"/></svg>';
      row.addEventListener('click', function () {
        setSelectValue(dataId, o.value);
        closeAllSelects();
      });
      panel.appendChild(row);
    });

    btn.addEventListener('click', function (e) {
      e.stopPropagation();
      var wasOpen = container.classList.contains('open');
      closeAllSelects();
      if (!wasOpen && !btn.disabled) container.classList.add('open');
    });

    SELECTS[dataId] = { value: '', options: allOpts, container: container, disabled: false, placeholder: placeholder };
  }

  function closeAllSelects() {
    document.querySelectorAll('.csel.open').forEach(function (c) { c.classList.remove('open'); });
  }
  document.addEventListener('click', closeAllSelects);

  function setSelectValue(id, val) {
    var s = SELECTS[id];
    if (!s) return;
    s.value = val;
    var lbl = s.container.querySelector('.lbl');
    var found = s.options.find(function (o) { return o.value === val; });
    lbl.textContent = found ? found.label : s.placeholder;
    lbl.classList.toggle('ph', val === '');
    s.container.querySelectorAll('.csel-opt').forEach(function (row) {
      row.classList.toggle('sel', row.dataset.value === val);
    });
    recomputeForm();
  }

  function selVal(id) { return SELECTS[id] ? SELECTS[id].value : ''; }

  function selDisable(id, disabled) {
    var s = SELECTS[id];
    if (!s) return;
    s.disabled = disabled;
    s.container.querySelector('.csel-btn').disabled = disabled;
    if (disabled) s.container.classList.remove('open');
  }

  function buildAllSelectsFromConfig() {
    buildSelect('csel-tipoestrutura', 'tipoestrutura', CONFIG.listas.tipoestrutura);
    buildSelect('csel-nrmodulos', 'nrmodulos', [
      { value: '1', label: '1 Módulo' }, { value: '2', label: '2 Módulos' }, { value: '3', label: '3 Módulos' }, { value: '4', label: '4 Módulos' },
      { value: '5', label: '5 Módulos' }, { value: '6', label: '6 Módulos' }, { value: '7', label: '7 Módulos' }, { value: '8', label: '8 Módulos' }
    ]);
    buildSelect('csel-planpin', 'planpin', CONFIG.listas.planpin);
    buildSelect('csel-tipomaq', 'tipomaq', CONFIG.listas.tipomaq);
    buildSelect('csel-incendio', 'incendio', CONFIG.listas.incendio);
    buildSelect('csel-seguranca', 'seguranca', CONFIG.listas.seguranca);
    buildSelect('csel-complexidade', 'complexidade', CONFIG.listas.complexidade);
    buildSelect('csel-planejadorSel', 'planejadorSel', CONFIG.listas.planejadorSel);
  }
  buildAllSelectsFromConfig();

  /* Steppers Handlers */
  document.querySelectorAll('.stepper').forEach(function (st) {
    var input = st.querySelector('input');
    var dec = st.querySelector('.dec');
    var inc = st.querySelector('.inc');
    function step(delta) {
      if (input.disabled) return;
      var v = parseInt(input.value || '0', 10);
      if (isNaN(v)) v = 0;
      v = Math.max(0, v + delta);
      input.value = v;
      input.dispatchEvent(new Event('input', { bubbles: true }));
    }
    if (dec) dec.addEventListener('click', function () { step(-1); });
    if (inc) inc.addEventListener('click', function () { step(1); });
    if (input) {
      input.addEventListener('input', function () {
        input.value = input.value.replace(/[^0-9]/g, '');
      });
    }
  });

  /* Module Rows Setup */
  var moduleRowsEl = $('moduleRows');
  var moduleInputs = [];
  if (moduleRowsEl) {
    for (var i = 1; i <= 8; i++) {
      var row = document.createElement('div');
      row.className = 'mrow locked';
      row.innerHTML =
        '<span class="mlabel"><i class="led na" data-led="mod' + i + '"></i>Módulo ' + i + '</span>' +
        '<span class="unit-suffix"><input type="text" inputmode="decimal" class="ipt mono mod-comp" placeholder="0,00"></span>' +
        '<span class="x">×</span>' +
        '<span class="unit-suffix"><input type="text" inputmode="decimal" class="ipt mono mod-larg" placeholder="0,00"></span>';
      moduleRowsEl.appendChild(row);
      moduleInputs.push(row);
    }
  }

  /* Form Rules & Validation Engine */
  function setLed(name, state) {
    document.querySelectorAll('[data-led="' + name + '"]').forEach(function (el) {
      el.classList.remove('na', 'req', 'ok'); el.classList.add(state);
    });
  }

  function isFilledEl(el) {
    if (!el) return false;
    if (el.type === 'checkbox') return el.checked;
    return String(el.value || '').trim() !== '';
  }

  var reqTotal = 0, reqDone = 0;
  var secCount = {};

  function trackField(section, pending) {
    if (pending === null || pending === undefined) return;
    reqTotal += 1;
    if (!pending) reqDone += 1;
    secCount[section] = (secCount[section] || 0) + (pending ? 1 : 0);
  }

  function applyInput(section, id, required, applicable) {
    var el = $(id);
    if (!el) return;
    el.disabled = !applicable;
    var wrap = el.closest('.stepper');
    if (wrap) wrap.classList.toggle('disabled', !applicable);
    var field = el.closest('.field');
    if (field && section === 'sec-sapinfo') {
      field.style.display = applicable ? '' : 'none';
    }
    if (!applicable) { setLed(id, 'na'); if (wrap) wrap.removeAttribute('data-req'); el.removeAttribute('data-req'); return; }
    var filled = isFilledEl(el);
    el.dataset.req = required ? '1' : '0';
    el.dataset.filled = filled ? '1' : '0';
    if (wrap) { wrap.dataset.req = required ? '1' : '0'; wrap.dataset.filled = filled ? '1' : '0'; }
    setLed(id, required ? (filled ? 'ok' : 'req') : (filled ? 'ok' : 'na'));
    if (required) trackField(section, !filled);
  }

  function applySelect(section, id, required, applicable) {
    selDisable(id, !applicable);
    if (!SELECTS[id]) return;
    var container = SELECTS[id].container;
    var field = container.closest('.field');
    if (field && section === 'sec-sapinfo') {
      field.style.display = applicable ? '' : 'none';
    }
    if (!applicable) { setLed(id, 'na'); container.removeAttribute('data-req'); return; }
    var filled = selVal(id) !== '';
    container.dataset.req = required ? '1' : '0';
    container.dataset.filled = filled ? '1' : '0';
    setLed(id, required ? (filled ? 'ok' : 'req') : (filled ? 'ok' : 'na'));
    if (required) trackField(section, !filled);
  }

  function recomputeForm() {
    reqTotal = 0; reqDone = 0; secCount = {};

    var tipo = selVal('tipoestrutura');
    var nrmod = parseInt(selVal('nrmodulos') || '0', 10);
    var temModulos = tipo !== '' && CONFIG.regras.estruturasSemModulo.indexOf(tipo) === -1;

    applySelect('sec-estrutura', 'tipoestrutura', true, true);
    applySelect('sec-estrutura', 'nrmodulos', true, true);
    applySelect('sec-estrutura', 'planpin', true, true);

    var expandirModulos = temModulos && nrmod > 0;
    if ($('modulesBlock')) $('modulesBlock').style.display = expandirModulos ? '' : 'none';
    moduleInputs.forEach(function (row, i0) {
      var idx = i0 + 1;
      var active = expandirModulos && idx <= nrmod;
      row.style.display = active ? '' : 'none';
      row.classList.toggle('locked', !active);
      row.classList.toggle('last-visible', active && idx === nrmod);
      var comp = row.querySelector('.mod-comp');
      var larg = row.querySelector('.mod-larg');
      comp.disabled = !active; larg.disabled = !active;
      var filled = active && comp.value.trim() !== '' && larg.value.trim() !== '';
      setLed('mod' + idx, !active ? 'na' : (filled ? 'ok' : 'req'));
      if (active) trackField('sec-estrutura', !filled);
    });

    var roofTop1 = selVal('tipomaq') === 'Roof Top' && nrmod === 1 && temModulos;
    if ($('warnRoofTop1Mod')) $('warnRoofTop1Mod').classList.toggle('show', roofTop1);

    var tipomaqVal = selVal('tipomaq');
    var maqAtiva = tipomaqVal !== '' && tipomaqVal !== 'Não possui' && tipomaqVal !== 'Não aplicável';
    applySelect('sec-eletrica', 'tipomaq', true, true);
    applyInput('sec-eletrica', 'qtdmaq', true, maqAtiva);
    applySelect('sec-eletrica', 'incendio', true, true);
    applySelect('sec-eletrica', 'seguranca', true, true);
    applySelect('sec-eletrica', 'complexidade', true, true);
    applyInput('sec-eletrica', 'nrcolunas', true, true);

    var isContainer = tipo === 'Container Solar';
    var blockContainerMaritimo = $('blockContainerMaritimo');
    if (blockContainerMaritimo) {
      blockContainerMaritimo.style.display = isContainer ? '' : 'none';
      if (!isContainer) {
        if ($('progReles')) $('progReles').checked = false;
        if ($('diagBTI')) $('diagBTI').checked = false;
        if ($('diagAgrup')) $('diagAgrup').checked = false;
      }
    }
    if ($('containerHint')) $('containerHint').style.opacity = isContainer ? '1' : '.55';

    var anySap = $('criarDRs').checked ||
      ($('cpc47') && $('cpc47').checked) ||
      ($('solar') && $('solar').checked) ||
      $('planejar').checked ||
      $('camposUsuario').checked ||
      $('amarrarMaterial').checked;

    var sapSection = $('sec-sapinfo');
    if (sapSection) {
      sapSection.style.display = anySap ? '' : 'none';
      sapSection.classList.toggle('disabled', !anySap);
    }
    if ($('sapLockChip')) $('sapLockChip').style.display = anySap ? 'none' : 'flex';
    if ($('warnPep')) $('warnPep').classList.toggle('show', anySap && !isFilledEl($('pep')));

    var hasCamposUser = $('camposUsuario').checked;
    var hasPlanejar = $('planejar').checked;
    var hasAmarrar = $('amarrarMaterial').checked;

    applyInput('sec-sapinfo', 'pep', anySap, anySap);
    applyInput('sec-sapinfo', 'cliente', hasCamposUser, hasCamposUser);
    applySelect('sec-sapinfo', 'planejadorSel', hasCamposUser, hasCamposUser);
    applyInput('sec-sapinfo', 'dataOV', hasCamposUser, hasCamposUser);
    applyInput('sec-sapinfo', 'valorMec', hasCamposUser && CONFIG.regras.estruturasSemValorMec.indexOf(tipo) === -1, hasCamposUser);
    applyInput('sec-sapinfo', 'valorEletr', hasCamposUser && CONFIG.regras.estruturasSemValorEletr.indexOf(tipo) === -1, hasCamposUser);
    applyInput('sec-sapinfo', 'dataInicio', hasPlanejar, hasPlanejar);
    applyInput('sec-sapinfo', 'material', hasAmarrar && tipo !== 'Serviço Engenharia', hasAmarrar);

    ['nrOV', 'itemOV'].forEach(function (id) {
      var el = $(id);
      if (el) {
        el.disabled = !hasCamposUser;
        var f = el.closest('.field');
        if (f) f.style.display = hasCamposUser ? '' : 'none';
      }
    });

    ['materialMec', 'materialEle', 'materialAvo'].forEach(function (id) {
      var el = $(id);
      if (el) {
        el.disabled = !hasAmarrar;
        var f = el.closest('.field');
        if (f) f.style.display = hasAmarrar ? '' : 'none';
      }
    });

    if (sapSection) {
      sapSection.querySelectorAll('.row').forEach(function (r) {
        var visibleChildren = Array.from(r.children).filter(function (child) {
          return child.style.display !== 'none';
        });
        r.style.display = visibleChildren.length > 0 ? '' : 'none';
      });
    }

    var acCount = document.querySelectorAll('.acessorio:checked').length;
    if ($('acessoriosCount')) $('acessoriosCount').textContent = acCount + ' selecionado' + (acCount === 1 ? '' : 's');

    /* Section nav badges */
    document.querySelectorAll('#sectionNavPlanejamento .snav-btn').forEach(function (link) {
      var n = secCount[link.dataset.anchor] || 0;
      var badge = link.querySelector('.n');
      if (badge) {
        badge.textContent = n;
        badge.classList.toggle('zero', n === 0);
      }
      link.classList.toggle('ok-badge', n === 0);
    });

    /* Header Ring & Footer */
    var pct = reqTotal ? reqDone / reqTotal : 0;
    var circumference = 2 * Math.PI * 8.5;
    if ($('ringFill')) $('ringFill').setAttribute('stroke-dasharray', (circumference * pct).toFixed(1) + ' ' + circumference.toFixed(1));
    if ($('ringText')) $('ringText').textContent = reqDone + '/' + reqTotal;
    if ($('ringWrap')) $('ringWrap').classList.toggle('complete', reqTotal > 0 && reqDone === reqTotal);
    if ($('footerMeta')) $('footerMeta').textContent = reqDone + ' / ' + reqTotal + ' obrigatórios';

    /* Atualização da recomendação do Seletor de PEP */
    atualizarSugestaoPep();
  }

  document.querySelectorAll('input, .acessorio').forEach(function (el) {
    el.addEventListener('input', recomputeForm);
    el.addEventListener('change', recomputeForm);
  });

  recomputeForm();

  /* ==========================================================================
     PDF DROPZONE & PARSER AUTOFILL LOGIC
     ========================================================================== */
  var parsedPdfState = {
    lastResult: null,
    filename: ''
  };

  function setupPdfDropzone() {
    var dropzone = $('pdfDropzone');
    var fileInput = $('pdfFileInput');
    var dropLoading = $('pdfDropLoading');
    var dropContent = document.querySelector('.pdf-drop-content');
    var dropIcon = document.querySelector('.pdf-drop-icon');
    var statusBar = $('pdfStatusBar');
    var btnRemover = $('btnRemoverPdf');
    var btnVerCampos = $('btnVerCamposPdf');

    if (!dropzone || !fileInput) return;

    dropzone.addEventListener('click', function () {
      fileInput.click();
    });

    fileInput.addEventListener('change', function (e) {
      if (e.target.files && e.target.files.length > 0) {
        processarPdfArquivo(e.target.files[0]);
      }
    });

    ['dragenter', 'dragover'].forEach(function (eventName) {
      dropzone.addEventListener(eventName, function (e) {
        e.preventDefault();
        e.stopPropagation();
        dropzone.classList.add('dragover');
      });
    });

    ['dragleave', 'drop'].forEach(function (eventName) {
      dropzone.addEventListener(eventName, function (e) {
        e.preventDefault();
        e.stopPropagation();
        dropzone.classList.remove('dragover');
      });
    });

    dropzone.addEventListener('drop', function (e) {
      var dt = e.dataTransfer;
      var files = dt.files;
      if (files && files.length > 0) {
        processarPdfArquivo(files[0]);
      }
    });

    if (btnRemover) {
      btnRemover.addEventListener('click', function (e) {
        e.stopPropagation();
        parsedPdfState.lastResult = null;
        parsedPdfState.filename = '';
        if (statusBar) statusBar.style.display = 'none';
        if (fileInput) fileInput.value = '';
        showToast('Importação de PDF removida.');
      });
    }

    if (btnVerCampos) {
      btnVerCampos.addEventListener('click', function (e) {
        e.stopPropagation();
        if (parsedPdfState.lastResult && parsedPdfState.lastResult.logs) {
          var msg = 'Campos identificados no PDF (' + (parsedPdfState.filename || 'Documento') + '):\n\n' +
            parsedPdfState.lastResult.logs.map(function (l) { return '• ' + l; }).join('\n');
          alert(msg);
        } else {
          alert('Nenhum detalhe disponível para exibição.');
        }
      });
    }
  }

  function processarPdfArquivo(file) {
    if (!file || !file.name.toLowerCase().endsWith('.pdf')) {
      showToast('Por favor selecione um arquivo no formato PDF.', true);
      return;
    }

    var dropLoading = $('pdfDropLoading');
    var dropContent = document.querySelector('.pdf-drop-content');
    var dropIcon = document.querySelector('.pdf-drop-icon');
    var statusBar = $('pdfStatusBar');

    if (dropLoading) dropLoading.style.display = 'flex';
    if (dropContent) dropContent.style.display = 'none';
    if (dropIcon) dropIcon.style.display = 'none';
    if (statusBar) statusBar.style.display = 'none';

    var reader = new FileReader();
    reader.onload = function (e) {
      var base64Data = e.target.result;

      function onResult(res) {
        if (dropLoading) dropLoading.style.display = 'none';
        if (dropContent) dropContent.style.display = '';
        if (dropIcon) dropIcon.style.display = '';

        if (!res || res.status === 'error') {
          showToast(res ? res.message : 'Erro ao processar PDF.', true);
          return;
        }

        parsedPdfState.lastResult = res;
        parsedPdfState.filename = file.name;

        if (statusBar) {
          statusBar.style.display = 'flex';
          if ($('pdfFileName')) $('pdfFileName').textContent = file.name;
          if ($('pdfBadgeFields')) $('pdfBadgeFields').textContent = (res.total_campos || Object.keys(res.fields || {}).length) + ' campos identificados';
        }

        applyParsedPdfData(res.fields || {});
        showToast('PDF importado com sucesso! ' + (res.total_campos || 0) + ' campos preenchidos automaticamente.');
      }

      if (isPyWebviewAvailable()) {
        window.pywebview.api.parse_pdf({
          pdf_base64: base64Data,
          filename: file.name
        }).then(onResult).catch(function (err) {
          onResult({ status: 'error', message: err.message || 'Falha ao executar parser de PDF' });
        });
      } else {
        apiCall('/parse_pdf', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ pdf_base64: base64Data, filename: file.name })
        }).then(onResult).catch(function (err) {
          onResult({ status: 'error', message: err.message || 'Falha na requisição ao backend' });
        });
      }
    };
    reader.onerror = function () {
      if (dropLoading) dropLoading.style.display = 'none';
      if (dropContent) dropContent.style.display = '';
      if (dropIcon) dropIcon.style.display = '';
      showToast('Erro ao ler arquivo PDF do disco.', true);
    };
    reader.readAsDataURL(file);
  }

  function applyParsedPdfData(fields) {
    if (!fields || typeof fields !== 'object') return;

    var modifiedElements = [];

    function findBestMatch(list, targetVal) {
      if (!list || !Array.isArray(list) || !targetVal) return null;
      var cleanTarget = String(targetVal).trim().toLowerCase();
      for (var i = 0; i < list.length; i++) {
        var it = list[i];
        var itemVal = typeof it === 'string' ? it : (it.value || it.label || '');
        if (itemVal.toLowerCase() === cleanTarget || itemVal.toLowerCase().indexOf(cleanTarget) !== -1 || cleanTarget.indexOf(itemVal.toLowerCase()) !== -1) {
          return typeof it === 'string' ? it : it.value;
        }
      }
      return null;
    }

    // 1. Tipo de Estrutura
    if (fields.tipoestrutura && CONFIG && CONFIG.listas && CONFIG.listas.tipoestrutura) {
      var matchTipo = findBestMatch(CONFIG.listas.tipoestrutura, fields.tipoestrutura);
      if (matchTipo) {
        setSelectValue('tipoestrutura', matchTipo);
        var cEl = $('csel-tipoestrutura');
        if (cEl) modifiedElements.push(cEl);
      }
    }

    // 2. Qtd Módulos
    if (fields.nrmodulos) {
      var nmodStr = String(fields.nrmodulos);
      setSelectValue('nrmodulos', nmodStr);
      var cEl = $('csel-nrmodulos');
      if (cEl) modifiedElements.push(cEl);
    }

    // 3. Plano de Pintura
    if (fields.planpin && CONFIG && CONFIG.listas && CONFIG.listas.planpin) {
      var matchPlan = findBestMatch(CONFIG.listas.planpin, fields.planpin);
      if (matchPlan) {
        setSelectValue('planpin', matchPlan);
        var cEl = $('csel-planpin');
        if (cEl) modifiedElements.push(cEl);
      }
    }

    // 4. Dimensões por Módulo
    if (fields.modulos && Array.isArray(fields.modulos)) {
      fields.modulos.forEach(function (m, idx) {
        if (idx < 8 && moduleInputs && moduleInputs[idx]) {
          var row = moduleInputs[idx];
          var compInput = row.querySelector('.mod-comp');
          var largInput = row.querySelector('.mod-larg');
          if (compInput && m.c !== undefined) {
            compInput.value = String(m.c).replace('.', ',');
            modifiedElements.push(compInput);
          }
          if (largInput && m.l !== undefined) {
            largInput.value = String(m.l).replace('.', ',');
            modifiedElements.push(largInput);
          }
        }
      });
    }

    // 5. Checkboxes Estruturais
    if (fields.chapaRemovivel !== undefined && $('chapaRemovivel')) {
      $('chapaRemovivel').checked = Boolean(fields.chapaRemovivel);
      modifiedElements.push($('chapaRemovivel').parentElement);
    }
    if (fields.peDireito !== undefined && $('peDireito')) {
      $('peDireito').checked = Boolean(fields.peDireito);
      modifiedElements.push($('peDireito').parentElement);
    }

    // 6. Ar Condicionado
    if (fields.tipomaq && CONFIG && CONFIG.listas && CONFIG.listas.tipomaq) {
      var matchMaq = findBestMatch(CONFIG.listas.tipomaq, fields.tipomaq);
      if (matchMaq) {
        setSelectValue('tipomaq', matchMaq);
        var cEl = $('csel-tipomaq');
        if (cEl) modifiedElements.push(cEl);
      }
    }
    if (fields.qtdmaq !== undefined && $('qtdmaq')) {
      $('qtdmaq').value = String(fields.qtdmaq);
      modifiedElements.push($('qtdmaq').parentElement);
    }

    // 7. Incêndio e Segurança
    if (fields.incendio && CONFIG && CONFIG.listas && CONFIG.listas.incendio) {
      var matchInc = findBestMatch(CONFIG.listas.incendio, fields.incendio);
      if (matchInc) {
        setSelectValue('incendio', matchInc);
        var cEl = $('csel-incendio');
        if (cEl) modifiedElements.push(cEl);
      }
    }
    if (fields.seguranca && CONFIG && CONFIG.listas && CONFIG.listas.seguranca) {
      var matchSeg = findBestMatch(CONFIG.listas.seguranca, fields.seguranca);
      if (matchSeg) {
        setSelectValue('seguranca', matchSeg);
        var cEl = $('csel-seguranca');
        if (cEl) modifiedElements.push(cEl);
      }
    }

    // 8. Equipamentos e Complexidade
    if (fields.complexidade && CONFIG && CONFIG.listas && CONFIG.listas.complexidade) {
      var matchCpx = findBestMatch(CONFIG.listas.complexidade, fields.complexidade);
      if (matchCpx) {
        setSelectValue('complexidade', matchCpx);
        var cEl = $('csel-complexidade');
        if (cEl) modifiedElements.push(cEl);
      }
    }
    if (fields.nrcolunas !== undefined && $('nrcolunas')) {
      $('nrcolunas').value = String(fields.nrcolunas);
      modifiedElements.push($('nrcolunas').parentElement);
    }
    if (fields.trafoOleo !== undefined && $('trafoOleo')) {
      $('trafoOleo').checked = Boolean(fields.trafoOleo);
      modifiedElements.push($('trafoOleo').parentElement);
    }
    if (fields.testesw !== undefined && $('testesw')) {
      $('testesw').checked = Boolean(fields.testesw);
      modifiedElements.push($('testesw').parentElement);
    }

    // 9. Acessórios
    if (fields.acessorios && Array.isArray(fields.acessorios)) {
      document.querySelectorAll('.acessorio').forEach(function (chk) {
        var flag = chk.dataset.flag;
        if (flag && fields.acessorios.indexOf(flag) !== -1) {
          chk.checked = true;
          modifiedElements.push(chk.parentElement);
        }
      });
    }

    // 10. Container Solar Flags
    if (fields.progReles !== undefined && $('progReles')) $('progReles').checked = Boolean(fields.progReles);
    if (fields.diagBTI !== undefined && $('diagBTI')) $('diagBTI').checked = Boolean(fields.diagBTI);
    if (fields.diagAgrup !== undefined && $('diagAgrup')) $('diagAgrup').checked = Boolean(fields.diagAgrup);

    // 11. PEP, Cliente, Planejador (SAP)
    if (fields.pep && $('pep')) {
      $('pep').value = fields.pep;
      modifiedElements.push($('pep'));
    }
    if (fields.cliente && $('cliente')) {
      $('cliente').value = fields.cliente;
      modifiedElements.push($('cliente'));
    }
    if (fields.planejadorSel && CONFIG && CONFIG.listas && CONFIG.listas.planejadorSel) {
      var matchPlanSel = findBestMatch(CONFIG.listas.planejadorSel, fields.planejadorSel);
      if (matchPlanSel) {
        setSelectValue('planejadorSel', matchPlanSel);
        var cEl = $('csel-planejadorSel');
        if (cEl) modifiedElements.push(cEl);
      }
    }

    // Recalcula todo o formulário
    recomputeForm();

    // Pulse animation nos elementos atualizados
    modifiedElements.forEach(function (el) {
      if (!el) return;
      el.classList.remove('field-highlight-pulse');
      void el.offsetWidth;
      el.classList.add('field-highlight-pulse');
      setTimeout(function () { el.classList.remove('field-highlight-pulse'); }, 2000);
    });
  }

  setupPdfDropzone();

  /* Section Navigation Smooth Scroll & Active Observer */
  var planNavBtns = document.querySelectorAll('#sectionNavPlanejamento .snav-btn');
  planNavBtns.forEach(function (btn) {
    btn.addEventListener('click', function () {
      var targetId = btn.dataset.anchor;
      var targetEl = $(targetId);
      if (targetEl) {
        targetEl.scrollIntoView({ behavior: 'smooth', block: 'start' });
        planNavBtns.forEach(function (b) { b.classList.remove('active'); });
        btn.classList.add('active');

        // Highlight section for 2 seconds
        document.querySelectorAll('section.pane.highlighted').forEach(function (p) {
          p.classList.remove('highlighted');
        });
        void targetEl.offsetWidth; // trigger reflow for clean animation restart
        targetEl.classList.add('highlighted');
        setTimeout(function () {
          targetEl.classList.remove('highlighted');
        }, 2000);
      }
    });
  });

  var planSections = Array.prototype.map.call(planNavBtns, function (btn) {
    return $(btn.dataset.anchor);
  });
  var scrollParent = document.querySelector('.content');
  if ('IntersectionObserver' in window) {
    var snavIo = new IntersectionObserver(function (entries) {
      entries.forEach(function (entry) {
        if (entry.isIntersecting) {
          var id = entry.target.id;
          planNavBtns.forEach(function (btn) {
            btn.classList.toggle('active', btn.dataset.anchor === id);
          });
        }
      });
    }, { root: scrollParent, rootMargin: '-10% 0px -75% 0px', threshold: 0 });
    planSections.forEach(function (sec) {
      if (sec) snavIo.observe(sec);
    });
  }

  /* ==========================================================================
     CÁLCULO DE TEMPOS & RESULTADO MODAL LOGIC
     ========================================================================== */
  function collectFormContext() {
    var tipoestrutura = selVal('tipoestrutura') || 'Móvel';
    var planpin = selVal('planpin') || 'WAU-ELETRO-08';
    var tipomaq = selVal('tipomaq') || 'Não possui';
    var complexidade = selVal('complexidade') || 'Simples';
    var incendio = selVal('incendio') || 'Não aplicável';
    var seguranca = selVal('seguranca') || 'Não possui';

    var firstComp = (moduleInputs && moduleInputs[0] && moduleInputs[0].querySelector('.mod-comp')) ? parseFloat(moduleInputs[0].querySelector('.mod-comp').value.replace(',', '.')) : NaN;
    var firstLarg = (moduleInputs && moduleInputs[0] && moduleInputs[0].querySelector('.mod-larg')) ? parseFloat(moduleInputs[0].querySelector('.mod-larg').value.replace(',', '.')) : NaN;

    var comp = !isNaN(firstComp) && firstComp > 0 ? firstComp : (parseFloat(($('comp') ? $('comp').value : '10').replace(',', '.')) || 10);
    var larg = !isNaN(firstLarg) && firstLarg > 0 ? firstLarg : (parseFloat(($('larg') ? $('larg').value : '3.6').replace(',', '.')) || 3.6);
    var alt = parseFloat(($('alt') ? $('alt').value : '0').replace(',', '.')) || 0;
    var nmod = parseInt(selVal('nrmodulos') || ($('nmod') ? $('nmod').value : '1'), 10) || 1;
    var qtdmaq = parseInt(($('qtdmaq') ? $('qtdmaq').value : '0'), 10) || 0;
    var nrcolunas = parseInt(($('nrcolunas') ? $('nrcolunas').value : '0'), 10) || 0;

    var chapaRemovivel = $('chapaRemovivel') && $('chapaRemovivel').checked ? 'Sim' : 'Não';
    var peDireito = $('peDireito') && $('peDireito').checked ? 'Sim' : 'Não';
    var testesw = $('testesw') && $('testesw').checked ? 'Sim' : 'Não';
    var white_martins = $('whiteMartins') && $('whiteMartins').checked ? 'Sim' : 'Não';
    var trafo_oleo = $('trafoOleo') && $('trafoOleo').checked ? 'Sim' : 'Não';
    var casa_maquinas = (tipomaq === 'Roof Top' || qtdmaq > 0) ? 'Sim' : 'Não';

    var acess_escada_weg = 'Não', acess_escada_esp = 'Não', acess_porao = 'Não', acess_pilotis = 'Não';
    var acess_dutos = 'Não', acess_fundo_falso = 'Não', acess_dutos_bww = 'Não', acess_calhas = 'Não', acess_dutos_gases = 'Não';

    document.querySelectorAll('.acessorio').forEach(function (el) {
      if (!el.checked) return;
      var flag = el.dataset.flag;
      if (!flag) {
        var txt = (el.parentElement ? el.parentElement.textContent : '').toLowerCase();
        if (txt.indexOf('weg') !== -1) flag = 'esc_plat_padao_weg';
        else if (txt.indexOf('especial') !== -1) flag = 'esc_plat_especial';
        else if (txt.indexOf('porão') !== -1 || txt.indexOf('porao') !== -1) flag = 'porao_de_cabos';
        else if (txt.indexOf('pilotis') !== -1) flag = 'pilotis';
        else if (txt.indexOf('rede de dutos') !== -1) flag = 'rede_de_dutos';
        else if (txt.indexOf('fundo falso') !== -1) flag = 'fundo_falso';
        else if (txt.indexOf('bww') !== -1) flag = 'dutos_bww';
        else if (txt.indexOf('calhas') !== -1) flag = 'calhas_pluviais';
        else if (txt.indexOf('gases') !== -1) flag = 'duto_de_gases';
      }
      if (flag === 'esc_plat_padao_weg') acess_escada_weg = 'Sim';
      if (flag === 'esc_plat_especial') acess_escada_esp = 'Sim';
      if (flag === 'porao_de_cabos') acess_porao = 'Sim';
      if (flag === 'pilotis') acess_pilotis = 'Sim';
      if (flag === 'rede_de_dutos') acess_dutos = 'Sim';
      if (flag === 'fundo_falso') acess_fundo_falso = 'Sim';
      if (flag === 'dutos_bww') acess_dutos_bww = 'Sim';
      if (flag === 'calhas_pluviais') acess_calhas = 'Sim';
      if (flag === 'duto_de_gases') acess_dutos_gases = 'Sim';
    });

    var ctx = {
      comp: comp, larg: larg, alt: alt, nmod: nmod,
      tipoestrutura: tipoestrutura, planpin: planpin,
      tipomaq: tipomaq, qtdmaq: qtdmaq, complexidade: complexidade,
      incendio: incendio, seguranca: seguranca, nrcolunas: nrcolunas,
      chapaRemovivel: chapaRemovivel, peDireito: peDireito,
      testesw: testesw, white_martins: white_martins, trafo_oleo: trafo_oleo,
      casa_maquinas: casa_maquinas,
      paineis_interlig: Math.max(1, Math.round(nrcolunas / 5)),
      dur_mcm: 1, dur_tes: 1, dur_ins: 1,
      acess_escada_weg: acess_escada_weg, acess_escada_esp: acess_escada_esp,
      acess_porao: acess_porao, acess_pilotis: acess_pilotis,
      acess_dutos: acess_dutos, acess_fundo_falso: acess_fundo_falso,
      acess_dutos_bww: acess_dutos_bww, acess_calhas: acess_calhas,
      acess_dutos_gases: acess_dutos_gases
    };

    var sumComp = 0;
    for (var i = 1; i <= 8; i++) {
      var mRow = (moduleInputs && moduleInputs[i - 1]) ? moduleInputs[i - 1] : null;
      var mCompVal = (mRow && mRow.querySelector('.mod-comp')) ? parseFloat(mRow.querySelector('.mod-comp').value.replace(',', '.')) : NaN;
      var mLargVal = (mRow && mRow.querySelector('.mod-larg')) ? parseFloat(mRow.querySelector('.mod-larg').value.replace(',', '.')) : NaN;

      var finalCompM = (!isNaN(mCompVal) && mCompVal > 0) ? mCompVal : comp;
      var finalLargM = (!isNaN(mLargVal) && mLargVal > 0) ? mLargVal : larg;

      ctx['comp_m' + i] = finalCompM;
      ctx['larg_m' + i] = finalLargM;

      if (i <= nmod) {
        sumComp += finalCompM;
      }
    }

    ctx.comp_acum = sumComp || (comp * nmod);
    return ctx;
  }

  function collectFormFlags(ctx) {
    return {
      chapa_remov: ctx.chapaRemovivel === 'Sim',
      pe_direito_3_3_m: ctx.peDireito === 'Sim',
      teste_software: ctx.testesw === 'Sim',
      white_martins: ctx.white_martins === 'Sim',
      trafo_a_oleo: ctx.trafo_oleo === 'Sim',
      trafo_oleo: ctx.trafo_oleo === 'Sim',
      casa_maquinas: ctx.casa_maquinas === 'Sim',
      sist_seguranca: ctx.seguranca !== 'Não possui' && ctx.seguranca !== 'Não aplicável' && ctx.seguranca !== '',
      incendio_c_combate: ctx.incendio === 'Com combate',
      climat_c_dutos: ctx.tipomaq === 'Roof Top' || ctx.tipomaq === 'Wall Mounted',
      esc_plat_padao_weg: ctx.acess_escada_weg === 'Sim',
      esc_plat_especial: ctx.acess_escada_esp === 'Sim',
      porao_de_cabos: ctx.acess_porao === 'Sim',
      pilotis: ctx.acess_pilotis === 'Sim',
      rede_de_dutos: ctx.acess_dutos === 'Sim',
      fundo_falso: ctx.acess_fundo_falso === 'Sim',
      dutos_bww: ctx.acess_dutos_bww === 'Sim',
      calhas_pluviais: ctx.acess_calhas === 'Sim',
      duto_de_gases: ctx.acess_dutos_gases === 'Sim'
    };
  }

  /* ==========================================================================
     MOTOR DE CRONOGRAMA & TEMPLATES COM OPERAÇÕES (FASE 3)
     ========================================================================== */
  function classificarDisciplinaTarefa(tarefaCode, desc) {
    var tNum = parseInt(String(tarefaCode || '').trim(), 10) || 0;
    var descUpper = String(desc || '').toUpperCase();

    if (tNum >= 500 && tNum < 600) {
      if (descUpper.indexOf('LOM') !== -1 || descUpper.indexOf('PBS') !== -1 || descUpper.indexOf('PPA') !== -1 ||
          descUpper.indexOf('PCI') !== -1 || descUpper.indexOf('PCE') !== -1 || descUpper.indexOf('PAC') !== -1 ||
          descUpper.indexOf('LCA') !== -1 || descUpper.indexOf('LAA') !== -1 || descUpper.indexOf('LAM') !== -1 ||
          descUpper.indexOf('LMA') !== -1 || descUpper.indexOf('PTR') !== -1 || descUpper.indexOf('LMT') !== -1) {
        return 'Engenharia Mecânica';
      }
      if (descUpper.indexOf('PIL') !== -1 || descUpper.indexOf('PCL') !== -1 || descUpper.indexOf('LMC') !== -1 ||
          descUpper.indexOf('PIN') !== -1 || descUpper.indexOf('LMI') !== -1 || descUpper.indexOf('PSS') !== -1 ||
          descUpper.indexOf('LMS') !== -1 || descUpper.indexOf('DIN') !== -1 || descUpper.indexOf('LMD') !== -1 ||
          descUpper.indexOf('PBA') !== -1 || descUpper.indexOf('LMB') !== -1 || descUpper.indexOf('PRF') !== -1) {
        return 'Engenharia Elétrica';
      }
      if (descUpper.indexOf('ROM') !== -1 || descUpper.indexOf('531') !== -1 || descUpper.indexOf('551') !== -1 ||
          descUpper.indexOf('561') !== -1 || descUpper.indexOf('581') !== -1 || descUpper.indexOf('585') !== -1 ||
          descUpper.indexOf('589') !== -1 || descUpper.indexOf('ROTEIRO') !== -1) {
        return 'Processos';
      }
      return 'Engenharia';
    } else if (tNum >= 600 && tNum < 700) {
      if (descUpper.indexOf('ROTEIRO') !== -1 || descUpper.indexOf('ROM') !== -1) {
        return 'Processos';
      }
      return 'Engenharia Elétrica';
    } else if (tNum >= 700 && tNum < 800) {
      if (descUpper.indexOf('PIN') !== -1 || descUpper.indexOf('PINTURA') !== -1) {
        return 'Pintura';
      }
      if (descUpper.indexOf('COR') !== -1 || descUpper.indexOf('FCH') !== -1 || descUpper.indexOf('PRB') !== -1 ||
          descUpper.indexOf('SBA') !== -1 || descUpper.indexOf('PRE') !== -1 || descUpper.indexOf('SES') !== -1 ||
          descUpper.indexOf('EDF') !== -1 || descUpper.indexOf('CHI') !== -1 || descUpper.indexOf('CHE') !== -1 ||
          descUpper.indexOf('ESTRUTURA') !== -1 || descUpper.indexOf('CALDEIRARIA') !== -1) {
        return 'Mecânica';
      }
      return 'Montagem Mecânica';
    } else if (tNum >= 800 && tNum < 900) {
      if (descUpper.indexOf('LBA') !== -1 || descUpper.indexOf('INT') !== -1 || descUpper.indexOf('TES') !== -1 ||
          descUpper.indexOf('INS') !== -1 || descUpper.indexOf('PEE') !== -1 || descUpper.indexOf('PEM') !== -1 ||
          descUpper.indexOf('FEC') !== -1 || descUpper.indexOf('FEA') !== -1 || descUpper.indexOf('FEQ') !== -1 ||
          descUpper.indexOf('ELETR') !== -1) {
        return 'Elétrica / Testes';
      }
      return 'Acessórios & Elétrica';
    } else if (tNum >= 900) {
      return 'Faturamento & Encerramento';
    }
    return 'Geral';
  }

  function gerarCronogramaCompleto(ctx, calcTimes, templateBlocks, seletorData) {
    if (!templateBlocks || !templateBlocks.cenarios) {
      return { cenario_id: 'padrao', cenario_descricao: 'Template Padrão', qtd_tarefas: 0, total_horas: 0, total_dias: 0, tarefas: [] };
    }

    var isFlagTrue = function (v) { return v === true || v === 'Sim' || v === 'true' || v === 1; };
    var tipo = String(ctx.tipoestrutura || '');
    var betim = isFlagTrue(ctx.Betim1310) || isFlagTrue(ctx.betim);
    var semEng = isFlagTrue(ctx.SemEng) || isFlagTrue(ctx.sem_eng);
    var nmodStr = String(ctx.nrmodulos || (ctx.nmod ? ctx.nmod + ' Módulo' + (ctx.nmod > 1 ? 's' : '') : '1 Módulo'));
    var tipomaq = String(ctx.tipomaq || 'Não possui');
    var seguranca = String(ctx.seguranca || 'Não possui');
    var testesw = isFlagTrue(ctx.testesw) || isFlagTrue(ctx.teste_software);
    var chkFilho = isFlagTrue(ctx.chkFilho) || isFlagTrue(ctx.item_filho);

    var cenarioId = '';
    if (tipo === 'Container Solar' || tipo === 'ESSW (mecânica)') {
      cenarioId = 'container_solar_essw_mecanica';
    } else if (tipo === 'Skid (mecânica)') {
      cenarioId = betim ? 'skid_mecanica_com_betim' : 'skid_mecanica_sem_betim';
    } else if (tipo === 'ESSW (elétrica)') {
      cenarioId = 'essw_eletrica';
    } else if (tipo === 'Pilotis') {
      cenarioId = 'pilotis';
    } else if (tipo === 'Skid (com elétrica)') {
      cenarioId = 'skid_com_eletrica_betim_true';
    } else if (tipo === 'Serviço Engenharia') {
      cenarioId = 'servico_engenharia';
    } else {
      cenarioId = 'eletrocentro_padrao';
    }

    var cenario = templateBlocks.cenarios[cenarioId];
    if (!cenario || !cenario.tarefas) {
      return { cenario_id: cenarioId, cenario_descricao: 'Cenário não encontrado', qtd_tarefas: 0, total_horas: 0, total_dias: 0, tarefas: [] };
    }

    var grid = {};
    cenario.tarefas.forEach(function (t) {
      var r = t.linha_template + 2;
      var durN = parseFloat(String(t.duracao).replace(',', '.')) || 0;
      var trabN = parseFloat(String(t.trabalho).replace(',', '.')) || 0;
      grid[r] = {
        row: r,
        tarefa: t.tarefa,
        descricao_tarefa: t.descricao_tarefa,
        duracao: durN,
        unidade: t.unidade || 'DIA',
        trabalho: trabN,
        calculado: false
      };
    });

    function getVal(k, def) {
      var v = calcTimes ? calcTimes[k] : undefined;
      if (v === undefined || v === null || isNaN(v)) return def !== undefined ? def : 0;
      return parseFloat(v) || 0;
    }

    function setCell(r, hKey, dKey) {
      if (grid[r]) {
        if (calcTimes && hKey && calcTimes[hKey] !== undefined) {
          grid[r].trabalho = Math.round(getVal(hKey) * 10) / 10;
          grid[r].calculado = true;
        }
        if (calcTimes && dKey && calcTimes[dKey] !== undefined) {
          grid[r].duracao = Math.round(getVal(dKey) * 10) / 10;
          grid[r].calculado = true;
        }
      }
    }

    function deleteExcelRows(minR, maxR) {
      var keys = Object.keys(grid).map(Number).sort(function (a, b) { return a - b; });
      var numDeleted = maxR - minR + 1;
      var newGrid = {};

      keys.forEach(function (r) {
        if (r < minR) {
          newGrid[r] = grid[r];
        } else if (r > maxR) {
          var newR = r - numDeleted;
          grid[r].row = newR;
          newGrid[newR] = grid[r];
        }
      });
      grid = newGrid;
    }

    // Aplicação por Cenário
    if (cenarioId === 'container_solar_essw_mecanica') {
      if (tipo === 'Container Solar' || tipo === 'ESSW (mecânica)') {
        if (grid[10]) grid[10].descricao_tarefa = 'PEC - Projeto Estrutura Container';
        if (grid[12]) grid[12].descricao_tarefa = 'EMC - Estagiamento Mat. Estrut. Cont.';
        if (grid[16]) grid[16].descricao_tarefa = 'PEI - Projeto Estrutura Interna';
        if (grid[18]) grid[18].descricao_tarefa = 'EMI - Estagiamento Mat. Estr. Interna';
        if (grid[69]) grid[69].descricao_tarefa = 'FPC - Fabricação Peças Caldeiraria';
        if (grid[74]) grid[74].descricao_tarefa = 'OEE - Ordens Estrutura Interna';
        if (grid[77]) grid[77].descricao_tarefa = 'SEI - Separação Estrutura Interna';
        if (grid[78]) grid[78].descricao_tarefa = 'MEI - Montagem Estrutura Interna';
        if (grid[121]) grid[121].descricao_tarefa = 'OFE - Ordens Fechamento Externo';
        if (grid[124]) grid[124].descricao_tarefa = 'SAF - Separação Almox. Fech. Externo';
        if (grid[126]) grid[126].descricao_tarefa = 'SFE - Separação Fechamento Externo';
        if (grid[127]) grid[127].descricao_tarefa = 'MFE - Montagem Fechamento Externo';
      }

      if (chkFilho && grid[153]) grid[153].tarefa = 899;
      if (grid[131]) grid[131].descricao_tarefa = 'SII - Separação Almox. Instal./ Inc.';
      if (grid[133]) grid[133].descricao_tarefa = 'MII - Montagem Instalações / Incêndio';
      if (seguranca !== 'Não possui' && grid[142]) grid[142].descricao_tarefa = 'LBS - Leito e Bandejamento / Sist. Seg.';

      setCell(4, 'HorLOM', 'DurLOM');
      setCell(8, 'HorLMM', 'DurLMM');
      setCell(10, 'HorPBS', 'DurPBS');
      setCell(13, 'HorPPA', 'DurPPA');
      setCell(16, 'HorPCI', 'DurPCI');
      setCell(19, 'HorPCE', 'DurPCE');
      setCell(22, 'HorPAC', 'DurPAC');
      setCell(24, 'HorLCA', 'DurLCA');
      setCell(27, 'HorLAA', 'DurLAA');
      setCell(30, 'HorLAM', 'DurLAM');
      setCell(33, 'HorLMA', 'DurLMA');
      setCell(37, 'HorPTR', 'DurPTR');
      setCell(41, 'HorPIL', 'DurPIL');
      setCell(45, 'HorPCL', 'DurPCL');
      setCell(47, 'HorLMC', 'DurLMC');
      setCell(49, 'HorPIN', 'DurPIN');
      setCell(51, 'HorLMI', 'DurLMI');

      if (seguranca !== 'Não possui') {
        setCell(53, 'HorPSS', 'DurPSS');
        setCell(55, 'HorLMS', 'DurLMS');
      }

      setCell(43, 'HorLMT', 'DurLMT');
      setCell(57, 'HorDIN', 'DurDIN');
      setCell(59, 'HorLMD', 'DurLMD');
      setCell(61, 'HorPBA', 'DurPBA');
      setCell(63, 'HorLMB', 'DurLMB');
      setCell(65, 'HorPRF', 'DurPRF');

      // Processos
      setCell(11, 'Hor531', 'Dur531');
      setCell(17, 'Hor551', 'Dur551');
      setCell(20, 'Hor561', 'Dur561');
      setCell(28, 'Hor581', 'Dur581');
      setCell(31, 'Hor585', 'Dur585');
      setCell(34, 'Hor589', 'Dur589');

      // Módulo 1
      setCell(69, 'HorCOR1', 'DurCOR1');
      setCell(70, 'HorFCH1', 'DurFCH1');
      setCell(72, 'HorEDF1', 'DurEDF1');
      setCell(75, 'HorPIN1', 'DurPIN1');
      setCell(78, 'HorCHI1', 'DurCHI1');
      setCell(71, 'HorCHE1', 'DurCHE1');

      // Acessórios
      setCell(120, 'HorFAC', 'DurFAC');
      setCell(123, 'HorFCA', 'DurFCA');
      setCell(127, 'HorMAM', 'DurMAM');
      setCell(129, 'HorMAA', 'DurMAA');
      setCell(132, 'HorPRM', 'DurPRM');
      setCell(133, 'HorIST', 'DurIST');
      setCell(135, 'HorMCL', 'DurMCL');

      if (tipomaq === 'Roof Top') setCell(136, 'HorMCM', 'DurMCM');
      if (grid[133]) {
        grid[133].trabalho = Math.round((grid[133].trabalho + getVal('HorMIN', 0)) * 10) / 10;
        grid[133].duracao = Math.round((grid[133].duracao + getVal('DurMIN', 0)) * 10) / 10;
      }

      setCell(140, 'HorFEQ', 'DurFEQ');
      setCell(142, 'HorLBA', 'DurLBA');

      if (seguranca !== 'Não possui' && grid[142]) {
        grid[142].trabalho = Math.round((grid[142].trabalho + getVal('HorMSS', 0)) * 10) / 10;
        grid[142].duracao = Math.round((grid[142].duracao + getVal('DurMSS', 0)) * 10) / 10;
      }

      setCell(145, 'HorINT', 'DurINT');
      setCell(147, 'HorTES', 'DurTES');
      setCell(149, 'HorINS', 'DurINS');
      setCell(150, 'HorPEE', 'DurPEE');
      setCell(151, 'HorPEM', 'DurPEM');
      setCell(152, 'HorFEC', 'DurFEC');
      setCell(130, 'HorFEA', 'DurFEA');

      if (chkFilho) deleteExcelRows(154, 154);
      if (tipo === 'ESSW (mecânica)') deleteExcelRows(131, 152);
      if (!testesw && tipo !== 'ESSW (mecânica)') deleteExcelRows(148, 148);
      if (!betim) deleteExcelRows(143, 143);
      if (tipo !== 'ESSW (mecânica)') deleteExcelRows(139, 139);
      if (tipo !== 'ESSW (mecânica)') deleteExcelRows(137, 138);
      if (tipomaq !== 'Roof Top' && tipo !== 'ESSW (mecânica)') deleteExcelRows(136, 136);

      if (nmodStr === '1 Módulo' || nmodStr === '2 Módulos' || nmodStr === '3 Módulos') deleteExcelRows(107, 118);
      if (nmodStr === '1 Módulo' || nmodStr === '2 Módulos') deleteExcelRows(95, 106);
      if (nmodStr === '1 Módulo') deleteExcelRows(83, 94);

      deleteExcelRows(79, 81);
      if (seguranca === 'Não possui' || tipo === 'ESSW (mecânica)') deleteExcelRows(53, 56);
      deleteExcelRows(25, 25);
      deleteExcelRows(19, 21);
      deleteExcelRows(13, 15);

      if (semEng && tipo === 'Container Solar') deleteExcelRows(4, 57);

    } else if (cenarioId === 'eletrocentro_padrao') {
      if (chkFilho && grid[243]) grid[243].tarefa = 899;
      if (grid[221]) grid[221].descricao_tarefa = 'SII - Separação Almox. Instal./ Inc.';
      if (grid[223]) grid[223].descricao_tarefa = 'MII - Montagem Instalações / Incêndio';
      if (seguranca !== 'Não possui' && grid[232]) grid[232].descricao_tarefa = 'LBS - Leito e Bandejamento / Sist. Seg.';

      setCell(4, 'HorLOM', 'DurLOM');
      setCell(8, 'HorLMM', 'DurLMM');
      if (grid[10]) {
        grid[10].trabalho = Math.round((getVal('HorPBS') + getVal('HorPPA')) * 10) / 10;
        grid[10].duracao = Math.round((getVal('DurPBS') + getVal('DurPPA')) * 10) / 10;
        grid[10].calculado = true;
      }
      setCell(13, 'HorPCI', 'DurPCI');
      setCell(16, 'HorPCE', 'DurPCE');
      setCell(19, 'HorPAC', 'DurPAC');
      setCell(21, 'HorLCA', 'DurLCA');
      setCell(24, 'HorLAA', 'DurLAA');
      setCell(27, 'HorLAM', 'DurLAM');
      setCell(30, 'HorLMA', 'DurLMA');
      setCell(34, 'HorPTR', 'DurPTR');

      setCell(38, 'HorPIL', 'DurPIL');
      setCell(40, 'HorLMT', 'DurLMT');
      setCell(42, 'HorPCL', 'DurPCL');
      setCell(44, 'HorLMC', 'DurLMC');
      setCell(46, 'HorPIN', 'DurPIN');
      setCell(48, 'HorLMI', 'DurLMI');

      if (seguranca !== 'Não possui') {
        setCell(50, 'HorPSS', 'DurPSS');
        setCell(52, 'HorLMS', 'DurLMS');
      }

      setCell(54, 'HorDIN', 'DurDIN');
      setCell(56, 'HorLMD', 'DurLMD');
      setCell(58, 'HorPBA', 'DurPBA');
      setCell(60, 'HorLMB', 'DurLMB');
      setCell(62, 'HorPRF', 'DurPRF');

      // Processos
      setCell(11, 'Hor531', 'Dur531');
      setCell(14, 'Hor551', 'Dur551');
      setCell(17, 'Hor561', 'Dur561');
      setCell(25, 'Hor581', 'Dur581');
      setCell(28, 'Hor585', 'Dur585');
      setCell(31, 'Hor589', 'Dur589');

      // Módulos 1 a 8
      var modMap = [
        [1, 66, 68, 69, 70, 71, 72, 73, 75, 78, 81],
        [2, 84, 86, 87, 88, 89, 90, 91, 93, 96, 99],
        [3, 102, 104, 105, 106, 107, 108, 109, 111, 114, 117],
        [4, 120, 122, 123, 124, 125, 126, 127, 129, 132, 135],
        [5, 138, 140, 141, 142, 143, 144, 145, 147, 150, 153],
        [6, 156, 158, 159, 160, 161, 162, 163, 165, 168, 171],
        [7, 174, 176, 177, 178, 179, 180, 181, 183, 186, 189],
        [8, 192, 194, 195, 196, 197, 198, 199, 201, 204, 207]
      ];
      modMap.forEach(function (m) {
        var mNum = m[0];
        setCell(m[1], 'HorCOR' + mNum, 'DurCOR' + mNum);
        setCell(m[2], 'HorFCH' + mNum, 'DurFCH' + mNum);
        setCell(m[3], 'HorPRB' + mNum, 'DurPRB' + mNum);
        setCell(m[4], 'HorSBA' + mNum, 'DurSBA' + mNum);
        setCell(m[5], 'HorPRE' + mNum, 'DurPRE' + mNum);
        setCell(m[6], 'HorSES' + mNum, 'DurSES' + mNum);
        setCell(m[7], 'HorEDF' + mNum, 'DurEDF' + mNum);
        setCell(m[8], 'HorPIN' + mNum, 'DurPIN' + mNum);
        setCell(m[9], 'HorCHI' + mNum, 'DurCHI' + mNum);
        setCell(m[10], 'HorCHE' + mNum, 'DurCHE' + mNum);
      });

      setCell(210, 'HorFAC', 'DurFAC');
      setCell(213, 'HorFCA', 'DurFCA');
      setCell(217, 'HorMAM', 'DurMAM');
      setCell(219, 'HorMAA', 'DurMAA');
      setCell(222, 'HorPRM', 'DurPRM');
      setCell(223, 'HorIST', 'DurIST');
      setCell(225, 'HorMCL', 'DurMCL');

      if (tipomaq === 'Roof Top') setCell(226, 'HorMCM', 'DurMCM');
      if (grid[223]) {
        grid[223].trabalho = Math.round((grid[223].trabalho + getVal('HorMIN', 0)) * 10) / 10;
        grid[223].duracao = Math.round((grid[223].duracao + getVal('DurMIN', 0)) * 10) / 10;
      }

      setCell(230, 'HorFEQ', 'DurFEQ');
      setCell(232, 'HorLBA', 'DurLBA');

      if (seguranca !== 'Não possui' && grid[232]) {
        grid[232].trabalho = Math.round((grid[232].trabalho + getVal('HorMSS', 0)) * 10) / 10;
        grid[232].duracao = Math.round((grid[232].duracao + getVal('DurMSS', 0)) * 10) / 10;
      }

      setCell(235, 'HorINT', 'DurINT');
      setCell(237, 'HorTES', 'DurTES');
      setCell(239, 'HorINS', 'DurINS');
      setCell(240, 'HorPEE', 'DurPEE');
      setCell(241, 'HorPEM', 'DurPEM');
      setCell(242, 'HorFEC', 'DurFEC');
      setCell(220, 'HorFEA', 'DurFEA');

      // Executa deleções sequenciais ordenadas
      var sortedKeys = Object.keys(grid).map(Number).sort(function (a, b) { return a - b; });
      var taskList = sortedKeys.map(function (k) { return grid[k]; });

      function delRows(minR, maxR) {
        var minI = minR - 4;
        var maxI = maxR - 4;
        taskList = taskList.filter(function (t, i) { return i < minI || i > maxI; });
      }

      if (chkFilho) delRows(244, 244);
      if (!testesw) delRows(238, 238);
      if (!betim) delRows(233, 233);
      delRows(229, 229);
      delRows(227, 228);
      if (tipomaq !== 'Roof Top') delRows(226, 226);

      if (nmodStr === '1 Módulo' || nmodStr === '2 Módulos' || nmodStr === '3 Módulos' || nmodStr === '4 Módulos' || nmodStr === '5 Módulos' || nmodStr === '6 Módulos' || nmodStr === '7 Módulos' || nmodStr === '1') {
        delRows(191, 208);
      }
      if (nmodStr === '1 Módulo' || nmodStr === '2 Módulos' || nmodStr === '3 Módulos' || nmodStr === '4 Módulos' || nmodStr === '5 Módulos' || nmodStr === '6 Módulos' || nmodStr === '1') {
        delRows(173, 190);
      }
      if (nmodStr === '1 Módulo' || nmodStr === '2 Módulos' || nmodStr === '3 Módulos' || nmodStr === '4 Módulos' || nmodStr === '5 Módulos' || nmodStr === '1') {
        delRows(155, 172);
      }
      if (nmodStr === '1 Módulo' || nmodStr === '2 Módulos' || nmodStr === '3 Módulos' || nmodStr === '4 Módulos' || nmodStr === '1') {
        delRows(137, 154);
      }
      if (nmodStr === '1 Módulo' || nmodStr === '2 Módulos' || nmodStr === '3 Módulos' || nmodStr === '1') {
        delRows(119, 136);
      }
      if (nmodStr === '1 Módulo' || nmodStr === '2 Módulos' || nmodStr === '1') {
        delRows(101, 118);
      }
      if (nmodStr === '1 Módulo' || nmodStr === '1') {
        delRows(83, 100);
      }

      delRows(82, 82);
      if (tipo === 'Móvel') delRows(71, 72);
      delRows(67, 67);
      if (seguranca === 'Não possui') delRows(50, 53);
      delRows(22, 22);

      if (semEng && tipo === 'Móvel') delRows(4, 60);

      // Reconstroi grid a partir da lista
      var reconstructedGrid = {};
      taskList.forEach(function (t, i) {
        reconstructedGrid[i + 4] = t;
      });
      grid = reconstructedGrid;

    } else if (cenarioId === 'skid_mecanica_sem_betim' || cenarioId === 'skid_mecanica_com_betim') {
      if (chkFilho) {
        var kList = Object.keys(grid).map(Number);
        var lastR = Math.max.apply(null, kList);
        if (grid[lastR]) grid[lastR].tarefa = 899;
        deleteExcelRows(lastR, lastR);
      }
      if (semEng) deleteExcelRows(4, 15);

    } else if (cenarioId === 'essw_eletrica') {
      if (chkFilho) {
        if (grid[15]) grid[15].tarefa = 899;
        deleteExcelRows(16, 16);
      }

    } else if (cenarioId === 'pilotis') {
      if (chkFilho) {
        if (grid[8]) grid[8].tarefa = 899;
        deleteExcelRows(9, 9);
      }

    } else if (cenarioId === 'skid_com_eletrica_betim_true') {
      if (chkFilho) {
        if (grid[69]) grid[69].tarefa = 899;
        deleteExcelRows(70, 70);
      }
      if (semEng) deleteExcelRows(4, 37);

    } else if (cenarioId === 'servico_engenharia') {
      setCell(4, 'HorLOM', 'DurLOM');
      setCell(8, 'HorLMM', 'DurLMM');
      setCell(9, 'HorPBS', 'DurPBS');
      setCell(10, 'HorPCI', 'DurPCI');
      setCell(11, 'HorPCE', 'DurPCE');
      setCell(12, 'HorPAC', 'DurPAC');
      setCell(14, 'HorLCA', 'DurLCA');
      setCell(15, 'HorLAA', 'DurLAA');
      setCell(16, 'HorLAM', 'DurLAM');
      setCell(17, 'HorLMA', 'DurLMA');
      setCell(19, 'HorPTR', 'DurPTR');
    }

    // Pós-processamento e inserção de ERE
    var finalTasks = [];
    var sortedRows = Object.keys(grid).map(Number).sort(function (a, b) { return a - b; });

    sortedRows.forEach(function (r) {
      var item = grid[r];
      var tCode = ('0000' + item.tarefa).slice(-4);
      var tDesc = String(item.descricao_tarefa || '').trim();

      // Inserção da ERE antes de 0760 CHI (ou após a pintura)
      if (tCode === '0760' && cenarioId === 'eletrocentro_padrao') {
        finalTasks.push({
          row: 75.5,
          tarefa: '0756',
          tarefa_formatada: '0756',
          descricao_tarefa: 'ERE - Emissão de Relatório',
          duracao: 1.0,
          unidade: 'DIA',
          trabalho: 0.1,
          calculado: false,
          disciplina: 'Pintura'
        });
      }

      if ((tCode === '0753' || tCode === '0749' || tCode === '0750') && tDesc.indexOf('PIN') === 0) {
        if (betim) {
          item.tarefa = '0749';
          item.descricao_tarefa = 'ESU - Envio para Subcontratação - PIN';
          item.duracao = 1.0;
          item.trabalho = 0.1;
          item.unidade = 'DIA';
        }
      }

      item.disciplina = classificarDisciplinaTarefa(item.tarefa, item.descricao_tarefa);
      finalTasks.push(item);
    });

    var totalH = 0;
    var totalDUR = 0;
    finalTasks.forEach(function (t, idx) {
      t.ordem = idx + 1;
      var raw = String(t.tarefa || '').trim();
      if (/^\d+$/.test(raw)) {
        t.tarefa_formatada = ('0000' + raw).slice(-4);
      } else {
        t.tarefa_formatada = raw;
      }
      totalH += t.trabalho;
      totalDUR += t.duracao;
    });

    return {
      cenario_id: cenarioId,
      cenario_descricao: cenario.descricao || cenarioId,
      qtd_tarefas: finalTasks.length,
      total_horas: Math.round(totalH * 100) / 100,
      total_dias: Math.round(totalDUR * 10) / 10,
      tarefas: finalTasks
    };
  }

  function executarCalculoTempos() {
    if (!state.regrasData || !state.regrasData.length) {
      showToast('Nenhuma regra disponível para cálculo no momento.', true);
      return null;
    }

    var ctx = collectFormContext();
    var flagsAtivos = collectFormFlags(ctx);
    Object.assign(SIM_CTX, ctx);
    salvarUltimaExecucao(ctx);

    var nmod = ctx.nmod || 1;
    var resultadosAreas = [];
    var totalGeralH = 0;
    var totalGeralDUR = 0;
    var calcTimes = {};

    state.regrasData.forEach(function (areaObj) {
      var areaNome = areaObj.area || 'Geral';
      var camposRes = [];
      var areaTotalH = 0;
      var areaTotalDUR = 0;

      var camposKeys = Object.keys(areaObj.campos || {});
      var isMec = (areaNome.indexOf('MEC') !== -1 && areaNome.indexOf('ENG') === -1);
      var isEngEle = (areaNome.indexOf('EL') !== -1 && areaNome.indexOf('ENG') !== -1);
      var isEletromec = (areaNome.indexOf('ELETROMEC') !== -1);

      camposKeys.forEach(function (cKey) {
        var campoObj = areaObj.campos[cKey];
        var valH = 0;
        var valDUR = 0;

        if (campoObj.H) {
          valH = calcValor(campoObj.H, nmod, flagsAtivos, campoObj, undefined);
          if (isNaN(valH) || !isFinite(valH)) valH = 0;
        }

        if (campoObj.DUR) {
          valDUR = calcValor(campoObj.DUR, nmod, flagsAtivos, campoObj, valH);
          if (isNaN(valDUR) || !isFinite(valDUR)) valDUR = 0;
        }

        valH = Math.round(valH * 100) / 100;
        valDUR = Math.round(valDUR * 10) / 10;

        areaTotalH += valH;
        areaTotalDUR += valDUR;

        if (isEngEle) {
          if (cKey === 'PIN') {
            calcTimes['HorPIN_ELE'] = valH;
            calcTimes['DurPIN_ELE'] = valDUR;
            calcTimes['HorPIN'] = valH;
            calcTimes['DurPIN'] = valDUR;
          } else if (cKey === 'LBA') {
            calcTimes['HorLMB'] = valH;
            calcTimes['DurLMB'] = valDUR;
            calcTimes['HorLBA_ELE'] = valH;
            calcTimes['DurLBA_ELE'] = valDUR;
          } else {
            calcTimes['Hor' + cKey] = valH;
            calcTimes['Dur' + cKey] = valDUR;
          }
        } else if (isMec) {
          if (cKey === 'PIN') {
            calcTimes['HorPIN_MEC'] = valH;
            calcTimes['DurPIN_MEC'] = valDUR;
            for (var mIdx = 1; mIdx <= 8; mIdx++) {
              calcTimes['HorPIN' + mIdx] = valH;
              calcTimes['DurPIN' + mIdx] = valDUR;
            }
          } else {
            calcTimes['Hor' + cKey] = valH;
            calcTimes['Dur' + cKey] = valDUR;
          }
        } else if (isEletromec) {
          if (cKey === 'LBA') {
            calcTimes['HorLBA_ELETR'] = valH;
            calcTimes['DurLBA_ELETR'] = valDUR;
            calcTimes['HorLBA'] = valH;
            calcTimes['DurLBA'] = valDUR;
          } else {
            calcTimes['Hor' + cKey] = valH;
            calcTimes['Dur' + cKey] = valDUR;
          }
        } else {
          calcTimes['Hor' + cKey] = valH;
          calcTimes['Dur' + cKey] = valDUR;
        }

        camposRes.push({
          chave: cKey,
          h: valH,
          dur: valDUR
        });
      });

      totalGeralH += areaTotalH;
      totalGeralDUR += areaTotalDUR;

      resultadosAreas.push({
        area: areaNome,
        campos: camposRes,
        totalH: Math.round(areaTotalH * 100) / 100,
        totalDUR: Math.round(areaTotalDUR * 10) / 10
      });
    });

    // Aliases e expansão de variáveis por módulo para o cronograma
    calcTimes['HorMAM'] = calcTimes['HorMAM'] || calcTimes['HorMAM/MFE'] || calcTimes['HorMFE'] || 0;
    calcTimes['DurMAM'] = calcTimes['DurMAM'] || calcTimes['DurMAM/MFE'] || calcTimes['DurMFE'] || 0;

    for (var m = 1; m <= 8; m++) {
      calcTimes['HorCOR' + m] = calcTimes['HorCOR' + m] || calcTimes['HorCOR/FPC'] || calcTimes['HorCOR'] || 0;
      calcTimes['DurCOR' + m] = calcTimes['DurCOR' + m] || calcTimes['DurCOR/FPC'] || calcTimes['DurCOR'] || 0;

      calcTimes['HorFCH' + m] = calcTimes['HorFCH' + m] || calcTimes['HorFCH'] || 0;
      calcTimes['DurFCH' + m] = calcTimes['DurFCH' + m] || calcTimes['DurFCH'] || 0;

      calcTimes['HorPRB' + m] = calcTimes['HorPRB' + m] || calcTimes['HorPRB'] || 0;
      calcTimes['DurPRB' + m] = calcTimes['DurPRB' + m] || calcTimes['DurPRB'] || 0;

      calcTimes['HorSBA' + m] = calcTimes['HorSBA' + m] || calcTimes['HorSBA'] || 0;
      calcTimes['DurSBA' + m] = calcTimes['DurSBA' + m] || calcTimes['DurSBA'] || 0;

      calcTimes['HorPRE' + m] = calcTimes['HorPRE' + m] || calcTimes['HorPRE'] || 0;
      calcTimes['DurPRE' + m] = calcTimes['DurPRE' + m] || calcTimes['DurPRE'] || 0;

      calcTimes['HorSES' + m] = calcTimes['HorSES' + m] || calcTimes['HorSES'] || 0;
      calcTimes['DurSES' + m] = calcTimes['DurSES' + m] || calcTimes['DurSES'] || 0;

      calcTimes['HorEDF' + m] = calcTimes['HorEDF' + m] || calcTimes['HorEDF'] || 0;
      calcTimes['DurEDF' + m] = calcTimes['DurEDF' + m] || calcTimes['DurEDF'] || 0;

      calcTimes['HorPIN' + m] = calcTimes['HorPIN' + m] || calcTimes['HorPIN'] || 0;
      calcTimes['DurPIN' + m] = calcTimes['DurPIN' + m] || calcTimes['DurPIN'] || 0;

      calcTimes['HorCHI' + m] = calcTimes['HorCHI' + m] || calcTimes['HorCHI/MEI'] || calcTimes['HorCHI'] || 0;
      calcTimes['DurCHI' + m] = calcTimes['DurCHI' + m] || calcTimes['DurCHI/MEI'] || calcTimes['DurCHI'] || 0;

      calcTimes['HorCHE' + m] = calcTimes['HorCHE' + m] || calcTimes['HorCHE/CSC'] || calcTimes['HorCHE'] || 0;
      calcTimes['DurCHE' + m] = calcTimes['DurCHE' + m] || calcTimes['DurCHE/CSC'] || calcTimes['DurCHE'] || 0;
    }

    // Processos (Roteiros) conforme regras do VBA
    var modNum = parseInt(ctx.nmod || 1, 10);
    calcTimes['Hor531'] = Math.min(3.1 * modNum, 15.2);
    calcTimes['Dur531'] = calcTimes['Hor531'] > 7.6 ? 2 : 1;

    calcTimes['Hor551'] = Math.min(0.8 * modNum, 15.2);
    calcTimes['Dur551'] = calcTimes['Hor551'] > 7.6 ? 2 : 1;

    calcTimes['Hor561'] = Math.min(0.8 * modNum, 15.2);
    calcTimes['Dur561'] = calcTimes['Hor561'] > 7.6 ? 2 : 1;

    calcTimes['Hor581'] = Math.min(4.0 * modNum, 15.2);
    calcTimes['Dur581'] = calcTimes['Hor581'] > 7.6 ? 2 : 1;

    calcTimes['Hor585'] = Math.min(4.8 * modNum, 15.2);
    calcTimes['Dur585'] = calcTimes['Hor585'] > 7.6 ? 2 : 1;

    calcTimes['Hor589'] = Math.min(2.0 * modNum, 15.2);
    calcTimes['Dur589'] = calcTimes['Hor589'] > 7.6 ? 2 : 1;

    // Garantir cálculo de PRM se não preenchido ou 0
    if (!calcTimes['HorPRM'] || calcTimes['HorPRM'] === 0) {
      var somaEle = (calcTimes['HorIST'] || 0) + (calcTimes['HorMCL'] || 0) + (calcTimes['HorMIN'] || 0) +
                    (calcTimes['HorMSS'] || 0) + (calcTimes['HorFEQ'] || 0) + (calcTimes['HorLBA'] || 0) +
                    (calcTimes['HorINT'] || 0) + (calcTimes['HorMCM'] || 0);
      calcTimes['HorPRM'] = Math.round((somaEle * 0.1 * 0.8 / 0.9) * 10) / 10;
      calcTimes['DurPRM'] = 3.0;
    }

    // Garantir TES e INS conformes com a complexidade e fórmulas do VBA
    var complex = String(ctx.complexidade || 'Simples');
    var durTes = complex === 'Complexo' ? 15 : (complex === 'Médio' ? 10 : 5);
    var durIns = complex === 'Complexo' ? 15 : 5;

    if (!calcTimes['DurTES'] || calcTimes['DurTES'] < durTes) {
      calcTimes['DurTES'] = durTes;
    }
    if (!calcTimes['HorTES'] || calcTimes['HorTES'] < 20) {
      calcTimes['HorTES'] = Math.round(durTes * 7.04 * 2 * 0.9 * 1.1 * 1.05 * 0.9 * 10) / 10;
    }

    if (!calcTimes['DurINS'] || calcTimes['DurINS'] < durIns) {
      calcTimes['DurINS'] = durIns;
    }
    if (!calcTimes['HorINS'] || calcTimes['HorINS'] < 20) {
      calcTimes['HorINS'] = Math.round(durIns * 7.04 * 2 * 0.9 * 1.1 * 1.05 * 0.85 * 10) / 10;
    }

    var seletorMatch = consultarSeletor(ctx);
    var cronograma = gerarCronogramaCompleto(ctx, calcTimes, state.templateBlocksData, seletorMatch);

    var calcResult = {
      ctx: ctx,
      seletor: seletorMatch,
      resultadosAreas: resultadosAreas,
      totalGeralH: Math.round(totalGeralH * 100) / 100,
      totalGeralDUR: Math.round(totalGeralDUR * 10) / 10,
      calc_times: calcTimes,
      cronograma: cronograma
    };

    // Calcula os totais por disciplina do padrão Excel
    calcResult.totaisDisciplinas = calcularTotaisDisciplinasExcel(cronograma ? cronograma.tarefas : []);

    // Dispara log automático de auditoria das informações preenchidas e calculadas
    enviarLogCalculo(calcResult);

    return calcResult;
  }

  function enviarLogCalculo(res) {
    if (!res || !res.ctx) return;
    try {
      var ctx = res.ctx;
      var crono = res.cronograma || {};
      var totaisDisc = res.totaisDisciplinas || calcularTotaisDisciplinasExcel(crono.tarefas || []);

      var logPayload = {
        ctx: ctx,
        totais_disciplinas: totaisDisc,
        resumo_areas: res.resultadosAreas || [],
        cronograma: {
          cenario_id: crono.cenario_id,
          qtd_tarefas: crono.qtd_tarefas || (crono.tarefas ? crono.tarefas.length : 0),
          total_horas: crono.total_horas || 0,
          total_dias: crono.total_dias || 0
        },
        seletor: res.seletor || {},
        data_hora: new Date().toISOString()
      };

      // Guarda snapshot no localStorage
      try {
        localStorage.setItem('ultimo_log_calculo', JSON.stringify(logPayload));
      } catch (e) {}

      // Log estruturado no console do navegador
      console.group('📊 [LOG CÁLCULO DE TEMPOS — Eletrocentros]');
      console.info('Formulário preenchido:', ctx);
      console.info('Horas por Diagrama (Padrão Excel):', totaisDisc);
      console.info('Resumo 5 Áreas Analíticas:', res.resultadosAreas);
      console.info('Cronograma (tarefas):', crono);
      console.groupEnd();

      // Envia ao backend Python / MySQL
      if (isPyWebviewAvailable()) {
        window.pywebview.api.log_calculo(logPayload).then(function (resp) {
          console.log('[Log Cálculo] Registrado com sucesso via pywebview:', resp);
        }).catch(function (err) {
          console.warn('[Log Cálculo] Falha via pywebview:', err);
        });
      } else {
        apiCall('/log_calculo', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(logPayload)
        }).then(function (resp) {
          console.log('[Log Cálculo] Registrado com sucesso via API:', resp);
        }).catch(function (err) {
          console.warn('[Log Cálculo] Falha via API local:', err);
        });
      }
    } catch (err) {
      console.error('[Log Cálculo] Erro ao enviar log:', err);
    }
  }

  function calcularTotaisDisciplinasExcel(tarefas) {
    var engH = 0;
    var mecH = 0;
    var eleH = 0;

    (tarefas || []).forEach(function (t) {
      var tNum = parseInt(t.tarefa || 0, 10);
      var trab = parseFloat(t.trabalho || 0) || 0;
      var descUpper = String(t.descricao_tarefa || '').toUpperCase();

      // Regra Engenharia: < 703, sem tempo 0.1 e sem tarefas ROM
      if (tNum < 703 && Math.abs(trab - 0.1) > 0.001 && descUpper.indexOf('ROM') === -1) {
        engH += trab;
      }

      // Regra Mecânica: 705 a 798 + tarefa 0894, sem tempo 0.1 e desconsiderar 754, 755, 765, 793, 794
      if (((tNum >= 705 && tNum <= 798) || tNum === 894) && Math.abs(trab - 0.1) > 0.001) {
        if ([754, 755, 765, 793, 794].indexOf(tNum) === -1) {
          mecH += trab;
        }
      }

      // Regra Elétrica: 799 a 893 + tarefa 0895, sem tempo 0.1 e desconsiderar 810, 828, 838, 858, 868
      if (((tNum >= 799 && tNum <= 893) || tNum === 895) && Math.abs(trab - 0.1) > 0.001) {
        if ([810, 828, 838, 858, 868].indexOf(tNum) === -1) {
          eleH += trab;
        }
      }
    });

    engH = Math.round(engH * 10) / 10;
    mecH = Math.round(mecH * 10) / 10;
    eleH = Math.round(eleH * 10) / 10;
    var totalH = Math.round((engH + mecH + eleH) * 10) / 10;

    return {
      eng_h: engH,
      mec_h: mecH,
      ele_h: eleH,
      total_h: totalH
    };
  }

  function initModalResultadoTabs() {
    var tabs = [
      { btn: $('calcTabTotais'), pane: $('calcPaneTotais') },
      { btn: $('calcTabCronograma'), pane: $('calcPaneCronograma') },
      { btn: $('calcTabSeletor'), pane: $('calcPaneSeletor') },
      { btn: $('calcTabResumo'), pane: $('calcPaneRegras') }
    ];

    tabs.forEach(function (t) {
      if (t.btn) {
        t.btn.addEventListener('click', function () {
          tabs.forEach(function (other) {
            if (other.btn) other.btn.classList.remove('active');
            if (other.pane) other.pane.classList.add('hidden');
          });
          t.btn.classList.add('active');
          if (t.pane) t.pane.classList.remove('hidden');
        });
      }
    });
  }

  function renderCronogramaResultado(cronograma) {
    if (!cronograma || !cronograma.tarefas) return;
    var tbody = $('resCronogramaTableBody');
    if (!tbody) return;
    tbody.innerHTML = '';

    if ($('cronogramaCenarioBadge')) {
      $('cronogramaCenarioBadge').textContent = 'Cenário: ' + (cronograma.cenario_descricao || cronograma.cenario_id);
    }
    if ($('resBadgeTarefas')) {
      $('resBadgeTarefas').textContent = String(cronograma.qtd_tarefas || cronograma.tarefas.length);
    }

    var searchTerm = ($('cronogramaSearch') ? $('cronogramaSearch').value.trim().toLowerCase() : '');
    var discFilter = ($('cronogramaDisciplineFilter') ? $('cronogramaDisciplineFilter').value : 'ALL');

    var filtered = cronograma.tarefas.filter(function (t) {
      if (discFilter !== 'ALL' && t.disciplina !== discFilter) return false;
      if (searchTerm) {
        var matchCode = (t.tarefa_formatada || String(t.tarefa)).toLowerCase().indexOf(searchTerm) !== -1;
        var matchDesc = String(t.descricao_tarefa || '').toLowerCase().indexOf(searchTerm) !== -1;
        var matchDisc = String(t.disciplina || '').toLowerCase().indexOf(searchTerm) !== -1;
        return matchCode || matchDesc || matchDisc;
      }
      return true;
    });

    var sumHoras = 0;
    var sumDias = 0;

    filtered.forEach(function (t) {
      sumHoras += (t.trabalho || 0);
      sumDias += (t.duracao || 0);

      var tr = document.createElement('tr');
      tr.className = 'task-row' + (t.calculado ? ' task-calc' : '');

      var discClass = 'disc-mec';
      if (t.disciplina === 'Engenharia Elétrica' || t.disciplina === 'Elétrica / Testes') discClass = 'disc-ele';
      else if (t.disciplina === 'Pintura') discClass = 'disc-pin';
      else if (t.disciplina === 'Processos') discClass = 'disc-pro';

      tr.innerHTML =
        '<td class="mono" style="text-align:center; color:var(--text-faint);">' + t.ordem + '</td>' +
        '<td style="text-align:center;"><span class="task-badge' + (t.calculado ? ' is-calc' : '') + '">' + escapeHtml(t.tarefa_formatada || t.tarefa) + '</span></td>' +
        '<td style="font-weight:' + (t.calculado ? '600' : '400') + '; color:var(--text);">' + escapeHtml(t.descricao_tarefa) + '</td>' +
        '<td class="num" style="color:var(--amber); font-weight:600;">' + t.duracao.toLocaleString('pt-BR', { minimumFractionDigits: 1, maximumFractionDigits: 1 }) + ' d</td>' +
        '<td style="text-align:center; font-size:10.5px; color:var(--text-faint);">' + escapeHtml(t.unidade || 'DIA') + '</td>' +
        '<td class="num" style="font-weight:700; color:' + (t.calculado ? 'var(--accent)' : 'var(--text)') + ';">' + t.trabalho.toLocaleString('pt-BR', { minimumFractionDigits: 1, maximumFractionDigits: 1 }) + ' h</td>' +
        '<td><span class="chip-disc ' + discClass + '">' + escapeHtml(t.disciplina || 'Geral') + '</span></td>';

      tbody.appendChild(tr);
    });

    if ($('cronogramaFootSummary')) {
      $('cronogramaFootSummary').innerHTML = 'Mostrando ' + filtered.length + ' de <b>' + cronograma.tarefas.length + '</b> operações';
    }
    if ($('cronogramaFootDias')) {
      $('cronogramaFootDias').textContent = sumDias.toLocaleString('pt-BR', { minimumFractionDigits: 1, maximumFractionDigits: 1 }) + ' d';
    }
    if ($('cronogramaFootHoras')) {
      $('cronogramaFootHoras').textContent = sumHoras.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' h';
    }
  }

  function renderResultadoCalculoView(res) {
    if (!res) return;

    var ctx = res.ctx;
    var seletorMatch = res.seletor || consultarSeletor(ctx);
    var tarefas = (res.cronograma && res.cronograma.tarefas) ? res.cronograma.tarefas : [];

    // Cálculo das Horas Totais por Diagrama (Padrão Excel Original)
    var totaisDisc = calcularTotaisDisciplinasExcel(tarefas);
    res.totaisDisciplinas = totaisDisc;

    var pepStr = (seletorMatch && seletorMatch['PEP Standard']) ? seletorMatch['PEP Standard'] : (ctx.pep || 'Sob Consulta');

    // Header Legend & Meta
    if ($('resHeaderPep')) $('resHeaderPep').innerHTML = 'PEP <b>' + escapeHtml(pepStr) + '</b>';
    if ($('resHeaderTarefasCount')) $('resHeaderTarefasCount').innerHTML = '<b>' + tarefas.length + '</b> operações processadas';
    if ($('resFooterMeta')) {
      $('resFooterMeta').innerHTML = '<b>' + totaisDisc.total_h.toLocaleString('pt-BR', { minimumFractionDigits: 1, maximumFractionDigits: 1 }) + ' h</b> totais · <b>' + tarefas.length + '</b> operações · PEP <b>' + escapeHtml(pepStr) + '</b>';
    }

    // 1. Atualizar KPIs do Resumo Geral (Aba 1)
    if ($('resKpiEngH')) $('resKpiEngH').textContent = totaisDisc.eng_h.toLocaleString('pt-BR', { minimumFractionDigits: 1, maximumFractionDigits: 1 });
    if ($('resKpiMecH')) $('resKpiMecH').textContent = totaisDisc.mec_h.toLocaleString('pt-BR', { minimumFractionDigits: 1, maximumFractionDigits: 1 });
    if ($('resKpiEleH')) $('resKpiEleH').textContent = totaisDisc.ele_h.toLocaleString('pt-BR', { minimumFractionDigits: 1, maximumFractionDigits: 1 });
    if ($('resKpiTotalH')) $('resKpiTotalH').textContent = totaisDisc.total_h.toLocaleString('pt-BR', { minimumFractionDigits: 1, maximumFractionDigits: 1 });

    // 2. Tabela de Horas Totais por Diagrama (Aba 1)
    var tbodyTotais = $('resTotaisTableBody');
    if (tbodyTotais) {
      tbodyTotais.innerHTML = '';
      var pctEng = totaisDisc.total_h > 0 ? (totaisDisc.eng_h / totaisDisc.total_h * 100).toFixed(1) : '0.0';
      var pctMec = totaisDisc.total_h > 0 ? (totaisDisc.mec_h / totaisDisc.total_h * 100).toFixed(1) : '0.0';
      var pctEle = totaisDisc.total_h > 0 ? (totaisDisc.ele_h / totaisDisc.total_h * 100).toFixed(1) : '0.0';

      var rowsData = [
        { code: 'eng', disc: 'Engenharia (ENG)', regra: 'Tarefas < 0703 — desconsidera tempo 0,1 e tarefas ROM', h: totaisDisc.eng_h, pct: pctEng },
        { code: 'mec', disc: 'Mecânica (MEC)', regra: 'Tarefas 0705 a 0798 + 0894 — desconsidera tempo 0,1 e 754, 755, 765, 793, 794', h: totaisDisc.mec_h, pct: pctMec },
        { code: 'ele', disc: 'Elétrica (ELE)', regra: 'Tarefas 0799 a 0893 + 0895 — desconsidera tempo 0,1 e 810, 828, 838, 858, 868', h: totaisDisc.ele_h, pct: pctEle }
      ];

      rowsData.forEach(function (r) {
        var tr = document.createElement('tr');
        tr.setAttribute('data-d', r.code);
        tr.innerHTML =
          '<td><span class="disc-name-result"><span class="disc-dot-result"></span>' + escapeHtml(r.disc) + '</span></td>' +
          '<td class="rule-text-result">' + escapeHtml(r.regra) + '</td>' +
          '<td class="num">' + r.h.toLocaleString('pt-BR', { minimumFractionDigits: 1, maximumFractionDigits: 1 }) + ' h</td>' +
          '<td class="pct"><div class="pct-bar-wrap-result"><div class="pct-bar-result"><i style="width:' + r.pct + '%;"></i></div>' + r.pct.replace('.', ',') + '%</div></td>';
        tbodyTotais.appendChild(tr);
      });

      var trTot = document.createElement('tr');
      trTot.className = 'total-row';
      trTot.setAttribute('data-d', 'total');
      trTot.innerHTML =
        '<td><span class="disc-name-result"><span class="disc-dot-result"></span>Total geral do projeto</span></td>' +
        '<td class="rule-text-result">Soma consolidada dos 3 diagramas — Engenharia + Mecânica + Elétrica</td>' +
        '<td class="num">' + totaisDisc.total_h.toLocaleString('pt-BR', { minimumFractionDigits: 1, maximumFractionDigits: 1 }) + ' h</td>' +
        '<td class="pct"><div class="pct-bar-wrap-result"><div class="pct-bar-result"><i style="width:100%;"></i></div>100,0%</div></td>';
      tbodyTotais.appendChild(trTot);
    }

    // 3. Info Estrutura e CTs no Resumo Geral
    if ($('resTotaisEstrutura')) {
      $('resTotaisEstrutura').textContent = ctx.nmod + ' Módulo(s) — ' + ctx.comp + 'm × ' + ctx.larg + 'm (' + ctx.tipoestrutura + ')';
    }
    if ($('resTotaisTags')) {
      $('resTotaisTags').innerHTML =
        '<span class="tag-result">Pintura <b>' + escapeHtml(ctx.planpin || '-') + '</b></span>' +
        '<span class="tag-result">Ar cond. <b>' + escapeHtml(ctx.tipomaq || '-') + ' (' + (ctx.qtdmaq || 1) + '×)</b></span>' +
        '<span class="tag-result">Complexidade <b>' + escapeHtml(ctx.complexidade || '-') + '</b></span>' +
        (ctx.incendio ? '<span class="tag-result">Incêndio <b>' + escapeHtml(ctx.incendio) + '</b></span>' : '') +
        (ctx.seguranca ? '<span class="tag-result">Segurança <b>' + escapeHtml(ctx.seguranca) + '</b></span>' : '');
    }
    if ($('resTotaisPepBadge')) {
      $('resTotaisPepBadge').textContent = 'PEP ' + pepStr;
    }
    if ($('resTotaisCtsResumo') && seletorMatch) {
      var ctList = [];
      if (seletorMatch['DR Eng Mec']) ctList.push({ role: 'Eng. Mecânica', dr: seletorMatch['DR Eng Mec'], alt: seletorMatch['Alt Eng Mec'] || '1' });
      if (seletorMatch['DR Eng Ele']) ctList.push({ role: 'Eng. Elétrica', dr: seletorMatch['DR Eng Ele'], alt: seletorMatch['Alt Eng Ele'] || '1' });

      var nmodNum = parseInt(ctx.nmod || '1', 10);
      for (var m = 1; m <= 8; m++) {
        var drM = seletorMatch['DR Mec ' + m] || seletorMatch['DR Mec' + m];
        var altM = seletorMatch['Alt Mec ' + m] || seletorMatch['Alt Mec' + m] || '1';
        if (drM && m <= nmodNum) {
          ctList.push({ role: 'Mecânica Módulo ' + m, dr: drM, alt: altM });
        }
      }
      if (seletorMatch['DR Acess']) ctList.push({ role: 'Acessórios', dr: seletorMatch['DR Acess'], alt: seletorMatch['Alt Acess'] || '1' });
      if (seletorMatch['DR Eletromec']) ctList.push({ role: 'Eletromecânica', dr: seletorMatch['DR Eletromec'], alt: seletorMatch['Alt Eletromec'] || '1' });

      var ctHtml = ctList.map(function(item) {
        return '<div class="ct-row-result"><span class="ct-role-result">' + escapeHtml(item.role) + '</span><span class="ct-code-result">DR ' + escapeHtml(item.dr) + ' <span class="alt">(Alt ' + escapeHtml(item.alt) + ')</span></span></div>';
      }).join('');
      $('resTotaisCtsResumo').innerHTML = ctHtml;
    }

    // 4. KPIs da Aba de Regras & Processos (Aba 4)
    if ($('resTotalH')) $('resTotalH').textContent = res.totalGeralH.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    if ($('resTotalDUR')) $('resTotalDUR').textContent = res.totalGeralDUR.toLocaleString('pt-BR', { minimumFractionDigits: 1, maximumFractionDigits: 1 });
    if ($('resEstruturaInfo')) $('resEstruturaInfo').textContent = ctx.nmod + ' Mód. (' + ctx.comp.toLocaleString('pt-BR') + 'm × ' + ctx.larg.toLocaleString('pt-BR') + 'm)';
    if ($('resDetagensInfo')) $('resDetagensInfo').textContent = ctx.tipoestrutura + ' · ' + ctx.planpin + ' · ' + ctx.tipomaq;

    // Reset para primeira aba (Resumo Geral de Horas)
    if ($('calcTabTotais')) $('calcTabTotais').click();

    // 5. Renderiza Tabela de Regras & Processos (Aba 4)
    var tbodyRegras = $('resTableBody');
    if (tbodyRegras) {
      tbodyRegras.innerHTML = '';
      res.resultadosAreas.forEach(function (area) {
        var trHead = document.createElement('tr');
        trHead.className = 'group-row';
        trHead.innerHTML = '<td colspan="3">' + escapeHtml(area.area) + '</td>';
        tbodyRegras.appendChild(trHead);

        area.campos.forEach(function (c) {
          var trC = document.createElement('tr');
          trC.innerHTML =
            '<td>' + escapeHtml(c.chave) + '</td>' +
            '<td class="num">' + c.h.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' h</td>' +
            '<td class="num" style="color:var(--amber);">' + c.dur.toLocaleString('pt-BR', { minimumFractionDigits: 1, maximumFractionDigits: 1 }) + ' d</td>';
          tbodyRegras.appendChild(trC);
        });

        var trSub = document.createElement('tr');
        trSub.className = 'subtotal-row';
        trSub.innerHTML =
          '<td>Subtotal (' + escapeHtml(area.area) + ')</td>' +
          '<td class="num">' + area.totalH.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' h</td>' +
          '<td class="num">' + area.totalDUR.toLocaleString('pt-BR', { minimumFractionDigits: 1, maximumFractionDigits: 1 }) + ' d</td>';
        tbodyRegras.appendChild(trSub);
      });

      var trGrand = document.createElement('tr');
      trGrand.className = 'total-row';
      trGrand.innerHTML =
        '<td><span class="disc-name-result"><span class="disc-dot-result"></span>TOTAL DAS REGRAS ORÇADAS</span></td>' +
        '<td class="num">' + res.totalGeralH.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' h</td>' +
        '<td class="num" style="color:var(--amber);">' + res.totalGeralDUR.toLocaleString('pt-BR', { minimumFractionDigits: 1, maximumFractionDigits: 1 }) + ' d</td>';
      tbodyRegras.appendChild(trGrand);
    }

    // 6. Renderiza Cronograma de Operações (Aba 2)
    renderCronogramaResultado(res.cronograma);

    // 7. Renderiza Seletor PEP & CTs (Aba 3)
    var ctsGrid = $('resSeletorCtsGrid');
    if (ctsGrid) {
      ctsGrid.innerHTML = '';
      if ($('resSeletorPepBadge')) {
        $('resSeletorPepBadge').textContent = 'PEP Standard: ' + (seletorMatch && seletorMatch['PEP Standard'] ? seletorMatch['PEP Standard'] : 'Sob Consulta');
      }
      if (seletorMatch) {
        var cts = [];
        if (seletorMatch['DR Eng Mec']) cts.push({ label: 'Eng. Mecânica', dr: seletorMatch['DR Eng Mec'], alt: seletorMatch['Alt Eng Mec'] });
        if (seletorMatch['DR Eng Ele']) cts.push({ label: 'Eng. Elétrica', dr: seletorMatch['DR Eng Ele'], alt: seletorMatch['Alt Eng Ele'] });

        var nmodNum = parseInt(ctx.nmod || '1', 10);
        for (var m = 1; m <= 8; m++) {
          var drMec = seletorMatch['DR Mec ' + m] || seletorMatch['DR Mec' + m];
          var altMec = seletorMatch['Alt Mec ' + m] || seletorMatch['Alt Mec' + m];
          if (drMec && m <= nmodNum) {
            cts.push({ label: 'Mecânica Módulo ' + m, dr: drMec, alt: altMec });
          }
        }
        if (seletorMatch['DR Acess']) cts.push({ label: 'Acessórios', dr: seletorMatch['DR Acess'], alt: seletorMatch['Alt Acess'] });
        if (seletorMatch['DR Eletromec']) cts.push({ label: 'Eletromecânica', dr: seletorMatch['DR Eletromec'], alt: seletorMatch['Alt Eletromec'] });

        cts.forEach(function (ct) {
          var card = document.createElement('div');
          card.className = 'ct-card';
          card.innerHTML = '<div class="role">' + escapeHtml(ct.label) + '</div>' +
            '<div class="code">' + escapeHtml(ct.dr) + ' <span class="alt">/ Alt ' + escapeHtml(ct.alt || '1') + '</span></div>';
          ctsGrid.appendChild(card);
        });

        if (ctx.nmod) {
          var modCard = document.createElement('div');
          modCard.className = 'ct-card';
          modCard.innerHTML = '<div class="role">Nº de módulos</div><div class="code">' + escapeHtml(ctx.nmod) + '</div>';
          ctsGrid.appendChild(modCard);
        }
      }
    }

    if (tabBtnResultados) tabBtnResultados.style.display = 'flex';
    switchView('resultados');
    window._lastCalculationResult = res;
  }

  var exibirModalResultadoCalculo = renderResultadoCalculoView;

  if ($('btnVoltarPlanejamento')) {
    $('btnVoltarPlanejamento').addEventListener('click', function () {
      switchView('planejamento');
    });
  }

  initModalResultadoTabs();

  if ($('cronogramaSearch')) {
    $('cronogramaSearch').addEventListener('input', function () {
      if (window._lastCalculationResult && window._lastCalculationResult.cronograma) {
        renderCronogramaResultado(window._lastCalculationResult.cronograma);
      }
    });
  }

  if ($('cronogramaDisciplineFilter')) {
    $('cronogramaDisciplineFilter').addEventListener('change', function () {
      if (window._lastCalculationResult && window._lastCalculationResult.cronograma) {
        renderCronogramaResultado(window._lastCalculationResult.cronograma);
      }
    });
  }

  if ($('btnCopiarResumoCalculo')) {
    $('btnCopiarResumoCalculo').addEventListener('click', function () {
      var res = window._lastCalculationResult;
      if (!res) return;
      var ctx = res.ctx;
      var crono = res.cronograma;
      var totalHVal = (crono && crono.total_horas) ? crono.total_horas : res.totalGeralH;
      var totalDurVal = (crono && crono.total_dias) ? crono.total_dias : res.totalGeralDUR;
      var totalTarefas = crono ? (crono.qtd_tarefas || crono.tarefas.length) : 0;

      var lines = [];
      lines.push('==================================================');
      lines.push('ELETROCENTROS APP — RESUMO DO CÁLCULO DE TEMPOS');
      lines.push('==================================================');
      lines.push('Estrutura: ' + ctx.nmod + ' Módulos — ' + ctx.comp + 'm x ' + ctx.larg + 'm (' + ctx.tipoestrutura + ')');
      lines.push('Plano Pintura: ' + ctx.planpin + ' | Máquina: ' + ctx.tipomaq + ' (' + ctx.qtdmaq + 'x)');
      lines.push('Complexidade: ' + ctx.complexidade + ' | Incêndio: ' + ctx.incendio + ' | Segurança: ' + ctx.seguranca);
      if (res.seletor && res.seletor['PEP Standard']) {
        lines.push('PEP Standard Sugerido: ' + res.seletor['PEP Standard']);
      }
      var totaisDisc = res.totaisDisciplinas || calcularTotaisDisciplinasExcel(crono ? crono.tarefas : []);

      lines.push('--------------------------------------------------');
      lines.push('HORAS TOTAIS POR DIAGRAMA (PADRÃO EXCEL):');
      lines.push(' • Engenharia (ENG): ' + totaisDisc.eng_h.toLocaleString('pt-BR', { minimumFractionDigits: 1 }) + ' h');
      lines.push(' • Mecânica (MEC):   ' + totaisDisc.mec_h.toLocaleString('pt-BR', { minimumFractionDigits: 1 }) + ' h');
      lines.push(' • Elétrica (ELE):   ' + totaisDisc.ele_h.toLocaleString('pt-BR', { minimumFractionDigits: 1 }) + ' h');
      lines.push(' • TOTAL DO PROJETO: ' + totaisDisc.total_h.toLocaleString('pt-BR', { minimumFractionDigits: 1 }) + ' h');
      lines.push('--------------------------------------------------');
      lines.push('DETALHAMENTO DE OPERAÇÕES GERADAS: ' + totalTarefas + ' tarefas');
      lines.push('DURAÇÃO ESTIMADA (DUR): ' + totalDurVal.toLocaleString('pt-BR', { minimumFractionDigits: 1 }) + ' dias');
      lines.push('TOTAL HORAS TODAS OPERAÇÕES: ' + totalHVal.toLocaleString('pt-BR', { minimumFractionDigits: 2 }) + ' h');
      lines.push('--------------------------------------------------');

      res.resultadosAreas.forEach(function (area) {
        lines.push('[' + area.area + ']');
        area.campos.forEach(function (c) {
          lines.push(' - ' + c.chave + ': ' + c.h.toLocaleString('pt-BR', { minimumFractionDigits: 2 }) + ' h | ' + c.dur.toLocaleString('pt-BR', { minimumFractionDigits: 1 }) + ' dias');
        });
        lines.push(' Subtotal: ' + area.totalH.toLocaleString('pt-BR', { minimumFractionDigits: 2 }) + ' h | ' + area.totalDUR.toLocaleString('pt-BR', { minimumFractionDigits: 1 }) + ' dias');
        lines.push('');
      });

      if (res.seletor) {
        lines.push('--------------------------------------------------');
        lines.push('CENTROS DE TRABALHO SUGERIDOS (SELETOR):');
        if (res.seletor['DR Eng Mec']) lines.push(' - Eng. Mecânica: DR ' + res.seletor['DR Eng Mec'] + ' (Alt ' + (res.seletor['Alt Eng Mec'] || '1') + ')');
        if (res.seletor['DR Eng Ele']) lines.push(' - Eng. Elétrica: DR ' + res.seletor['DR Eng Ele'] + ' (Alt ' + (res.seletor['Alt Eng Ele'] || '1') + ')');
        var nmodNum = parseInt(ctx.nmod || '1', 10);
        for (var m = 1; m <= 8; m++) {
          var drM = res.seletor['DR Mec ' + m] || res.seletor['DR Mec' + m];
          var altM = res.seletor['Alt Mec ' + m] || res.seletor['Alt Mec' + m] || '1';
          if (drM && m <= nmodNum) {
            lines.push(' - Mecânica Módulo ' + m + ': DR ' + drM + ' (Alt ' + altM + ')');
          }
        }
        if (res.seletor['DR Acess']) lines.push(' - Acessórios: DR ' + res.seletor['DR Acess'] + ' (Alt ' + (res.seletor['Alt Acess'] || '1') + ')');
        if (res.seletor['DR Eletromec']) lines.push(' - Eletromecânica: DR ' + res.seletor['DR Eletromec'] + ' (Alt ' + (res.seletor['Alt Eletromec'] || '1') + ')');
      }

      var textToCopy = lines.join('\r\n');
      if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(textToCopy).then(function () {
          showToast('Resumo copiado para a área de transferência!');
        }).catch(function () {
          copyFallback(textToCopy);
        });
      } else {
        copyFallback(textToCopy);
      }
    });
  }

  function downloadBlob(blob, filename) {
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(url); }, 2000);
  }

  function getExportFilename(extension) {
    var pepVal = isFilledEl($('pep')) ? $('pep').value.trim() : '';
    var pepClean = pepVal.replace(/[^a-zA-Z0-9_-]/g, '');
    var now = new Date();
    var y = now.getFullYear();
    var m = String(now.getMonth() + 1).padStart(2, '0');
    var d = String(now.getDate()).padStart(2, '0');
    var h = String(now.getHours()).padStart(2, '0');
    var min = String(now.getMinutes()).padStart(2, '0');
    var dateStr = y + m + d + '_' + h + min;
    if (pepClean) {
      return 'Resultado_PEP_' + pepClean + '_' + dateStr + '.' + extension;
    }
    return 'Resultado_Eletrocentro_' + dateStr + '.' + extension;
  }

  if ($('btnExportarCsv')) {
    $('btnExportarCsv').addEventListener('click', async function () {
      var res = window._lastCalculationResult;
      if (!res) {
        showToast('Nenhum resultado de cálculo disponível para exportar.', true);
        return;
      }

      var lines = [];
      lines.push('Tarefa;Descricao da Tarefa;Duracao;Unidade;Trabalho;Diagrama');
      if (res.cronograma && res.cronograma.tarefas) {
        res.cronograma.tarefas.forEach(function (t) {
          lines.push(
            '"' + (t.tarefa_formatada || t.tarefa) + '";' +
            '"' + (t.descricao_tarefa || '').replace(/"/g, '""') + '";' +
            t.duracao.toLocaleString('pt-BR', { minimumFractionDigits: 1 }) + ';' +
            (t.unidade || 'DIA') + ';' +
            t.trabalho.toLocaleString('pt-BR', { minimumFractionDigits: 1 }) + ';' +
            '"' + (t.disciplina || '') + '"'
          );
        });
      }

      var csvContent = '\uFEFF' + lines.join('\r\n');
      var filename = getExportFilename('csv');

      if (isPyWebviewAvailable() && window.pywebview.api.save_file) {
        window.pywebview.api.save_file({
          filename: filename,
          content: csvContent,
          is_base64: false,
          file_type: 'csv'
        }).then(function (resp) {
          if (resp && resp.status === 'success') {
            showToast('Arquivo CSV salvo com sucesso em: ' + (resp.filepath || resp.filename));
          } else if (resp && resp.status === 'cancelled') {
            showToast('Exportação cancelada.');
          } else {
            showToast('Erro ao salvar CSV: ' + ((resp && resp.message) || 'Erro desconhecido'), true);
          }
        }).catch(function (err) {
          showToast('Erro ao salvar CSV via pywebview: ' + (err.message || err), true);
        });
        return;
      }

      if ('showSaveFilePicker' in window) {
        try {
          var handle = await window.showSaveFilePicker({
            suggestedName: filename,
            types: [{
              description: 'Arquivo CSV (*.csv)',
              accept: { 'text/csv': ['.csv'] }
            }]
          });
          var writable = await handle.createWritable();
          await writable.write(new Blob([csvContent], { type: 'text/csv;charset=utf-8;' }));
          await writable.close();
          showToast('Arquivo CSV salvo com sucesso: ' + handle.name);
          return;
        } catch (err) {
          if (err.name === 'AbortError') {
            showToast('Exportação cancelada.');
            return;
          }
        }
      }

      // Fallback via API backend ou download padrão
      apiCall('/save_file', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ filename: filename, content: csvContent, is_base64: false, file_type: 'csv' })
      }).then(function (resp) {
        if (resp && resp.status === 'success') {
          showToast('Arquivo CSV salvo com sucesso em: ' + (resp.filepath || resp.filename));
        } else if (resp && resp.status === 'cancelled') {
          showToast('Exportação cancelada.');
        } else {
          var blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
          downloadBlob(blob, filename);
          showToast('Arquivo CSV exportado com sucesso: ' + filename);
        }
      }).catch(function () {
        var blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        downloadBlob(blob, filename);
        showToast('Arquivo CSV exportado com sucesso: ' + filename);
      });
    });
  }

  if ($('btnExportarExcel')) {
    $('btnExportarExcel').addEventListener('click', async function () {
      var res = window._lastCalculationResult;
      if (!res) {
        showToast('Nenhum resultado de cálculo disponível para exportar.', true);
        return;
      }

      var filename = getExportFilename('xlsx');
      var payload = {
        pep: isFilledEl($('pep')) ? $('pep').value.trim() : '',
        ctx: res.ctx,
        seletor: res.seletor,
        totalGeralH: res.totalGeralH,
        totalGeralDUR: res.totalGeralDUR,
        resultadosAreas: res.resultadosAreas,
        calc_times: res.calc_times,
        cronograma: res.cronograma,
        prompt_save: true
      };

      showToast('Escolha onde salvar o arquivo na janela que foi aberta...');

      if (isPyWebviewAvailable() && window.pywebview.api.export_excel) {
        window.pywebview.api.export_excel(payload).then(function (resp) {
          if (resp && resp.status === 'success') {
            showToast('Planilha Excel salva com sucesso em: ' + (resp.filepath || resp.filename));
          } else if (resp && resp.status === 'cancelled') {
            showToast('Exportação cancelada.');
          } else {
            showToast('Erro ao gerar Excel: ' + ((resp && resp.message) || 'Falha no backend'), true);
          }
        }).catch(function (err) {
          showToast('Erro ao exportar Excel via pywebview: ' + (err.message || err), true);
        });
        return;
      }

      // Se estiver no navegador e suportar showSaveFilePicker
      if ('showSaveFilePicker' in window) {
        try {
          var handle = await window.showSaveFilePicker({
            suggestedName: filename,
            types: [{
              description: 'Planilha Excel (*.xlsx)',
              accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] }
            }]
          });

          payload.prompt_save = false; // solicita o base64 para gravar diretamente no arquivo escolhido
          var resp = await apiCall('/export_excel', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
          });

          if (resp && resp.status === 'success' && resp.base64) {
            var byteCharacters = atob(resp.base64);
            var byteNumbers = new Array(byteCharacters.length);
            for (var i = 0; i < byteCharacters.length; i++) {
              byteNumbers[i] = byteCharacters.charCodeAt(i);
            }
            var byteArray = new Uint8Array(byteNumbers);
            var blob = new Blob([byteArray], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            var writable = await handle.createWritable();
            await writable.write(blob);
            await writable.close();
            showToast('Planilha Excel salva com sucesso: ' + handle.name);
            return;
          } else {
            showToast('Erro ao gerar Excel: ' + ((resp && resp.message) || 'Falha na API'), true);
            return;
          }
        } catch (err) {
          if (err.name === 'AbortError') {
            showToast('Exportação cancelada.');
            return;
          }
        }
      }

      // Fallback via API backend ou download padrão
      apiCall('/export_excel', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      }).then(function (resp) {
        if (resp && resp.status === 'success') {
          if (resp.filepath) {
            showToast('Planilha Excel salva com sucesso em: ' + resp.filepath);
          } else if (resp.base64) {
            var byteCharacters = atob(resp.base64);
            var byteNumbers = new Array(byteCharacters.length);
            for (var i = 0; i < byteCharacters.length; i++) {
              byteNumbers[i] = byteCharacters.charCodeAt(i);
            }
            var byteArray = new Uint8Array(byteNumbers);
            var blob = new Blob([byteArray], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            downloadBlob(blob, resp.filename || filename);
            showToast('Planilha Excel exportada: ' + (resp.filename || filename));
          }
        } else if (resp && resp.status === 'cancelled') {
          showToast('Exportação cancelada.');
        } else {
          showToast('Erro ao gerar Excel: ' + ((resp && resp.message) || 'Falha na API'), true);
        }
      }).catch(function (err) {
        showToast('Erro ao exportar Excel via API: ' + (err.message || err), true);
      });
    });
  }

  /* Footer Actions */
  if ($('btnOk')) {
    $('btnOk').addEventListener('click', function () {
      recomputeForm();
      if (reqDone < reqTotal) {
        var missingNames = [];
        document.querySelectorAll('.csel[data-req="1"], .stepper[data-req="1"], input[data-req="1"]').forEach(function (el) {
          if (el.dataset.filled !== '1') {
            var fieldLabel = el.closest('.field');
            if (fieldLabel) {
              var txt = fieldLabel.querySelector('.field-label .txt');
              if (txt) missingNames.push(txt.textContent.trim());
            }
          }
        });

        var tipo = selVal('tipoestrutura');
        var nrmod = parseInt(selVal('nrmodulos') || '0', 10);
        var temModulos = tipo !== '' && CONFIG.regras.estruturasSemModulo.indexOf(tipo) === -1;
        if (temModulos && nrmod > 0) {
          for (var i = 1; i <= nrmod; i++) {
            var row = $('modRow' + i);
            if (row) {
              var comp = row.querySelector('.mod-comp');
              var larg = row.querySelector('.mod-larg');
              if (!comp || !comp.value.trim() || !larg || !larg.value.trim()) {
                missingNames.push('Dimensões do Módulo ' + i);
              }
            }
          }
        }

        var countMissing = reqTotal - reqDone;
        var msg = 'Preencha todos os campos obrigatórios (' + countMissing + ' pendente' + (countMissing > 1 ? 's' : '') + ')';
        if (missingNames.length > 0) {
          var uniqueMissing = Array.from(new Set(missingNames));
          msg += ': ' + uniqueMissing.slice(0, 3).join(', ') + (uniqueMissing.length > 3 ? '…' : '');
        }
        showToast(msg, true);

        // Highlight and scroll to first missing field
        var firstMissing = document.querySelector('.csel[data-req="1"][data-filled="0"], .stepper[data-req="1"][data-filled="0"], .led.req');
        if (firstMissing) {
          var targetField = firstMissing.closest('.field') || firstMissing;
          targetField.scrollIntoView({ behavior: 'smooth', block: 'center' });
          targetField.classList.add('highlighted');
          setTimeout(function () { targetField.classList.remove('highlighted'); }, 2000);
        }
        return;
      }

      var res = executarCalculoTempos();
      if (res) {
        exibirModalResultadoCalculo(res);
        showToast('Cálculo de tempos concluído com sucesso!');
      } else {
        // Se regras.json ainda não havia sido carregado na memória, tenta carregar imediatamente
        fetch('regras.json').then(function (resp) { return resp.json(); }).then(function (data) {
          if (data && data.length) {
            state.regrasData = data;
            var res2 = executarCalculoTempos();
            if (res2) {
              exibirModalResultadoCalculo(res2);
              showToast('Cálculo de tempos concluído com sucesso!');
            }
          } else {
            showToast('Não foi possível carregar as regras de regras.json.', true);
          }
        }).catch(function (err) {
          showToast('Erro ao carregar regras.json: ' + (err.message || err), true);
        });
      }
    });
  }

  function limparFormularioCompleto() {
    // 1. Limpar todos os selects customizados para vazio / placeholder ("Selecione…")
    if (typeof SELECTS === 'object' && SELECTS !== null) {
      Object.keys(SELECTS).forEach(function (key) {
        setSelectValue(key, '');
      });
    } else {
      setSelectValue('tipoestrutura', '');
      setSelectValue('nrmodulos', '');
      setSelectValue('planpin', '');
      setSelectValue('tipomaq', '');
      setSelectValue('complexidade', '');
      setSelectValue('incendio', '');
      setSelectValue('seguranca', '');
      setSelectValue('planejadorSel', '');
    }

    // 2. Limpar todos os inputs de texto, número e steppers da tela de planejamento
    var planView = $('view-planejamento') || document;
    planView.querySelectorAll('input:not([type="checkbox"]):not([type="radio"]):not([type="button"]):not([type="submit"])').forEach(function (inp) {
      if (inp.id === 'qtdmaq') {
        inp.value = '0';
      } else {
        inp.value = '';
      }
    });

    // 3. Garantir limpeza explícita de todos os campos conhecidos
    if ($('nrcolunas')) $('nrcolunas').value = '';
    if ($('qtdmaq')) $('qtdmaq').value = '0';
    if ($('horLOM')) $('horLOM').value = '';
    if ($('horEDF')) $('horEDF').value = '';
    if ($('horINT')) $('horINT').value = '';
    if ($('pep')) $('pep').value = '';
    if ($('cliente')) $('cliente').value = '';
    if ($('valorMec')) $('valorMec').value = '';
    if ($('valorEletr')) $('valorEletr').value = '';
    if ($('dataOV')) $('dataOV').value = '';
    if ($('nrOV')) $('nrOV').value = '';
    if ($('itemOV')) $('itemOV').value = '';
    if ($('material')) $('material').value = '';
    if ($('dataInicio')) $('dataInicio').value = '';
    if ($('materialMec')) $('materialMec').value = '';
    if ($('materialEle')) $('materialEle').value = '';
    if ($('materialAvo')) $('materialAvo').value = '';

    // 4. Limpar todos os campos da tabela de módulos
    if (moduleInputs && moduleInputs.length > 0) {
      moduleInputs.forEach(function (row) {
        row.querySelectorAll('input').forEach(function (inp) {
          inp.value = '';
        });
      });
    }

    // 5. Desmarcar todos os checkboxes (estrutura, acessórios, automações SAP, fabricação)
    planView.querySelectorAll('input[type="checkbox"]').forEach(function (cb) {
      cb.checked = false;
    });

    // 6. Resetar estado do Seletor de PEP sugerido
    if ($('seletorPepBadge')) $('seletorPepBadge').style.display = 'none';

    // 7. Resetar estado do Parser de PDF Dropzone
    if (typeof parsedPdfState !== 'undefined' && parsedPdfState) {
      parsedPdfState.lastResult = null;
      parsedPdfState.filename = '';
    }
    if ($('pdfStatusBar')) $('pdfStatusBar').style.display = 'none';
    if ($('pdfFileInput')) $('pdfFileInput').value = '';

    // 8. Limpar cache de última sessão no navegador
    try {
      localStorage.removeItem('eletrocentros_last_execution');
    } catch (e) { }

    // 9. Recalcular todo o formulário (LEDs, contadores de obrigatórios, regras de exibição)
    recomputeForm();
    showToast('Formulário limpo com sucesso!');
  }

  var btnLimpar = $('btnLimparDados') || $('btnLimpar') || $('btnSair');
  if (btnLimpar) {
    btnLimpar.addEventListener('click', function () {
      limparFormularioCompleto();
    });
  }

  function salvarUltimaExecucao(ctx) {
    try {
      localStorage.setItem('eletrocentros_last_execution', JSON.stringify(ctx));
    } catch (e) { }
  }

  function carregarUltimaExecucao() {
    var ctx = {
      comp: 10,
      larg: 3.6,
      alt: 0,
      nmod: 1,
      tipoestrutura: 'Fixo',
      planpin: 'WAU-ELETRO-08',
      tipomaq: 'Wall Mounted',
      qtdmaq: 2,
      complexidade: 'Médio',
      incendio: 'Com instalações',
      seguranca: 'CFTV + Controle Acesso',
      nrcolunas: 20,
      chapaRemovivel: 'Sim',
      peDireito: 'Sim',
      testesw: 'Não',
      white_martins: 'Não',
      trafo_oleo: 'Não',
      casa_maquinas: 'Não',
      acess_escada_weg: 'Sim',
      acess_escada_esp: 'Não',
      acess_porao: 'Não',
      acess_pilotis: 'Não',
      acess_dutos: 'Sim',
      acess_fundo_falso: 'Não',
      acess_dutos_bww: 'Não',
      acess_calhas: 'Sim',
      acess_dutos_gases: 'Não'
    };

    // Atualiza campos de entrada
    if ($('comp')) $('comp').value = '10';
    if ($('larg')) $('larg').value = '3,6';
    if ($('alt')) $('alt').value = '2,6';
    if ($('qtdmaq')) $('qtdmaq').value = '2';
    if ($('nrcolunas')) $('nrcolunas').value = '20';

    // Atualiza selects customizados
    setSelectValue('tipoestrutura', 'Fixo');
    setSelectValue('nrmodulos', '1');
    setSelectValue('planpin', 'WAU-ELETRO-08');
    setSelectValue('tipomaq', 'Wall Mounted');
    setSelectValue('complexidade', 'Médio');
    setSelectValue('incendio', 'Com instalações');
    setSelectValue('seguranca', 'CFTV + Controle Acesso');

    // Atualiza checkboxes
    if ($('chapaRemovivel')) $('chapaRemovivel').checked = true;
    if ($('peDireito')) $('peDireito').checked = true;
    if ($('testesw')) $('testesw').checked = false;
    if ($('whiteMartins')) $('whiteMartins').checked = false;
    if ($('trafoOleo')) $('trafoOleo').checked = false;

    // Atualiza inputs específicos da tabela de módulos (Módulo 1)
    if (moduleInputs && moduleInputs[0]) {
      var comp1 = moduleInputs[0].querySelector('.mod-comp');
      var larg1 = moduleInputs[0].querySelector('.mod-larg');
      if (comp1) comp1.value = '10';
      if (larg1) larg1.value = '3,6';
    }

    document.querySelectorAll('.acessorio').forEach(function (el) {
      var flag = el.dataset.flag;
      if (!flag) return;
      if (flag === 'esc_plat_padao_weg') el.checked = (ctx.acess_escada_weg === 'Sim');
      if (flag === 'esc_plat_especial') el.checked = (ctx.acess_escada_esp === 'Sim');
      if (flag === 'porao_de_cabos') el.checked = (ctx.acess_porao === 'Sim');
      if (flag === 'pilotis') el.checked = (ctx.acess_pilotis === 'Sim');
      if (flag === 'rede_de_dutos') el.checked = (ctx.acess_dutos === 'Sim');
      if (flag === 'fundo_falso') el.checked = (ctx.acess_fundo_falso === 'Sim');
      if (flag === 'dutos_bww') el.checked = (ctx.acess_dutos_bww === 'Sim');
      if (flag === 'calhas_pluviais') el.checked = (ctx.acess_calhas === 'Sim');
      if (flag === 'duto_de_gases') el.checked = (ctx.acess_dutos_gases === 'Sim');
    });

    salvarUltimaExecucao(ctx);
    recomputeForm();
    showToast('Última sessão carregada com sucesso (Fixo, 1 Mód. 10×3,6m, Wall Mounted, 20 colunas)!');
  }

  if ($('btnCarregar')) {
    $('btnCarregar').addEventListener('click', function () {
      carregarUltimaExecucao();
    });
  }

  /* ==========================================================================
     MAINTENANCE PANEL LOGIC
     ========================================================================== */
  var FLAGS = [
    { key: 'sist_seguranca', nome: 'Sist. Segurança' }, { key: 'climat_c_dutos', nome: 'Climat. c/ dutos' },
    { key: 'incendio_c_combate', nome: 'Incêndio c/ combate' }, { key: 'casa_maquinas', nome: 'Casa Máquinas' },
    { key: 'teste_software', nome: 'Teste Software' }, { key: 'chapa_remov', nome: 'Chapa Removível' },
    { key: 'trafo_a_oleo', nome: 'Trafo a Óleo' }, { key: 'pe_direito_3_3_m', nome: 'Pé Direito 3,3m' },
    { key: 'esc_plat_padao_weg', nome: 'Esc./Plat. Padrão WEG' }, { key: 'esc_plat_especial', nome: 'Esc./Plat. Especial' },
    { key: 'porao_de_cabos', nome: 'Porão de Cabos' }, { key: 'pilotis', nome: 'Pilotis' },
    { key: 'rede_de_dutos', nome: 'Rede de Dutos' }, { key: 'fundo_falso', nome: 'Fundo Falso' },
    { key: 'dutos_bww', nome: 'Dutos BWW' }, { key: 'calhas_pluviais', nome: 'Calhas Pluviais' },
    { key: 'duto_de_gases', nome: 'Duto de Gases' }
  ];

  function flagNome(key) { var f = FLAGS.find(function (x) { return x.key === key; }); return f ? f.nome : key; }
  var MULT = { 1: 1, 2: 1.5, 3: 2, 4: 2.5, 5: 3, 6: 3.5, 7: 4, 8: 4.5 };

  /* ==========================================================================
     REGRAS VIEW — ESTRUTURA DE AREAS, CAMPOS (LOM, LMM, PBS...) E SUB-ABAS H / DUR
     ========================================================================== */
  state.regrasData = [];
  state.selectedAreaIdx = 0;
  state.selectedCampoKey = null;
  state.selectedSubTab = 'H'; // 'H' or 'DUR'
  state.dirtySubTabRule = null;

  function carregarDisciplinas() {
    var nav = $('sectionNavManutencao');
    if (nav) nav.innerHTML = '<div style="color:var(--text-faint); font-size:12px; padding:8px 12px;">Carregando regras.json…</div>';

    if (isPyWebviewAvailable()) {
      window.pywebview.api.get_regras().then(function (data) {
        if (data && data.length) {
          state.regrasData = data;
          renderRegrasAreasNav();
        } else {
          fetchRegrasFallback();
        }
      }).catch(function (err) {
        console.warn('[Regras] Erro pywebview:', err);
        fetchRegrasFallback();
      });
    } else {
      fetchRegrasFallback();
    }
  }

  function fetchRegrasFallback() {
    fetch('regras.json').then(function (resp) {
      if (resp.ok) return resp.json();
      throw new Error('Status ' + resp.status);
    }).then(function (data) {
      if (data && data.length) {
        state.regrasData = data;
        renderRegrasAreasNav();
      } else {
        var nav = $('sectionNavManutencao');
        if (nav) nav.innerHTML = '<div style="color:var(--text-danger); font-size:12px; padding:8px 12px;">Nenhuma regra encontrada no regras.json.</div>';
      }
    }).catch(function (err) {
      console.warn('[Regras] Erro ao carregar regras.json fallback:', err);
      var nav = $('sectionNavManutencao');
      if (nav) nav.innerHTML = '<div style="color:var(--text-danger); font-size:12px; padding:8px 12px;">Erro ao carregar regras.json.</div>';
    });
  }

  function renderRegrasAreasNav() {
    var nav = $('sectionNavManutencao');
    if (!nav) return;
    nav.innerHTML = '';

    state.regrasData.forEach(function (areaObj, idx) {
      var btn = document.createElement('button');
      btn.className = 'snav-btn snav-btn-area' + (!state.isSeletorActive && !state.isTemplatesActive && idx === state.selectedAreaIdx ? ' active' : '');
      var countCampos = Object.keys(areaObj.campos || {}).length;
      btn.innerHTML = areaObj.area + ' <span class="n">' + countCampos + '</span>';
      btn.addEventListener('click', function () {
        selecionarArea(idx);
      });
      nav.appendChild(btn);
    });

    // Botão dedicado ao Seletor de PEP & Centros de Trabalho
    var btnSeletor = document.createElement('button');
    btnSeletor.className = 'snav-btn snav-btn-seletor' + (state.isSeletorActive ? ' active' : '');
    btnSeletor.style.borderColor = 'rgba(46,196,182,0.4)';
    btnSeletor.innerHTML = '<svg viewBox="0 0 24 24" width="13" height="13" fill="none" stroke="currentColor" stroke-width="2" style="margin-right:4px; vertical-align:middle; color:var(--accent);"><path d="M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z"/></svg> Seletor PEP &amp; CTs <span class="n" style="background:var(--accent); color:#000; font-weight:700;">' + (state.seletorData ? state.seletorData.length : '89') + '</span>';
    btnSeletor.addEventListener('click', function () {
      abrirAbaSeletor();
    });
    nav.appendChild(btnSeletor);

    // Botão dedicado aos Templates de Operações (template_blocks.json)
    var btnTemplates = document.createElement('button');
    btnTemplates.className = 'snav-btn snav-btn-templates' + (state.isTemplatesActive ? ' active' : '');
    btnTemplates.style.borderColor = 'rgba(240,169,62,0.4)';
    btnTemplates.innerHTML = '<svg viewBox="0 0 24 24" width="13" height="13" fill="none" stroke="currentColor" stroke-width="2" style="margin-right:4px; vertical-align:middle; color:var(--amber);"><path d="M4 19.5A2.5 2.5 0 0 1 6.5 17H20"/><path d="M6.5 2H20v20H6.5A2.5 2.5 0 0 1 4 19.5v-15A2.5 2.5 0 0 1 6.5 2z"/></svg> Templates de Operações <span class="n" style="background:var(--amber); color:#000; font-weight:700;">8</span>';
    btnTemplates.addEventListener('click', function () {
      abrirAbaTemplates();
    });
    nav.appendChild(btnTemplates);

    if (state.isTemplatesActive) {
      abrirAbaTemplates();
    } else if (state.isSeletorActive) {
      abrirAbaSeletor();
    } else if (state.regrasData.length) {
      selecionarArea(state.selectedAreaIdx || 0);
    }
  }

  function selecionarArea(areaIdx) {
    state.isSeletorActive = false;
    state.isTemplatesActive = false;
    state.selectedAreaIdx = areaIdx;

    document.querySelectorAll('#sectionNavManutencao .snav-btn').forEach(function (b, i) {
      if (i === areaIdx) b.classList.add('active');
      else b.classList.remove('active');
    });

    if ($('maintContentRegras')) $('maintContentRegras').classList.remove('hidden');
    if ($('maintContentSeletor')) $('maintContentSeletor').classList.add('hidden');
    if ($('maintContentTemplates')) $('maintContentTemplates').classList.add('hidden');
    if ($('legendWrapper')) $('legendWrapper').style.display = '';

    if ($('btnAdicionarLinhaSeletor')) $('btnAdicionarLinhaSeletor').style.display = 'none';

    var areaObj = state.regrasData[areaIdx];
    var camposKeys = Object.keys(areaObj ? areaObj.campos || {} : {});
    state.selectedCampoKey = camposKeys.length ? camposKeys[0] : null;
    state.selectedSubTab = 'H';
    prepararDirtySubTab();
    renderList();
    renderEditor();
  }

  function abrirAbaSeletor() {
    state.isSeletorActive = true;
    state.isTemplatesActive = false;

    document.querySelectorAll('#sectionNavManutencao .snav-btn').forEach(function (b) {
      b.classList.remove('active');
    });
    var btnSel = document.querySelector('#sectionNavManutencao .snav-btn-seletor');
    if (btnSel) btnSel.classList.add('active');

    if ($('maintContentRegras')) $('maintContentRegras').classList.add('hidden');
    if ($('maintContentSeletor')) $('maintContentSeletor').classList.remove('hidden');
    if ($('maintContentTemplates')) $('maintContentTemplates').classList.add('hidden');
    if ($('legendWrapper')) $('legendWrapper').style.display = 'none';

    if ($('btnAdicionarLinhaSeletor')) $('btnAdicionarLinhaSeletor').style.display = 'inline-flex';

    if ($('maintFooterMeta')) {
      $('maintFooterMeta').textContent = 'Seletor de PEP & Centros de Trabalho (' + (state.seletorData ? state.seletorData.length : 0) + ' combinações)';
    }
    if ($('btnSaveFooter')) {
      $('btnSaveFooter').textContent = 'Salvar Seletor';
      $('btnSaveFooter').disabled = !state.seletorDirty;
    }
    renderSeletorTable();
  }

  function abrirAbaTemplates() {
    state.isTemplatesActive = true;
    state.isSeletorActive = false;

    document.querySelectorAll('#sectionNavManutencao .snav-btn').forEach(function (b) {
      b.classList.remove('active');
    });
    var btnTpl = document.querySelector('#sectionNavManutencao .snav-btn-templates');
    if (btnTpl) btnTpl.classList.add('active');

    if ($('maintContentRegras')) $('maintContentRegras').classList.add('hidden');
    if ($('maintContentSeletor')) $('maintContentSeletor').classList.add('hidden');
    if ($('maintContentTemplates')) $('maintContentTemplates').classList.remove('hidden');
    if ($('legendWrapper')) $('legendWrapper').style.display = 'none';

    if ($('btnAdicionarLinhaSeletor')) $('btnAdicionarLinhaSeletor').style.display = 'none';

    if ($('maintFooterMeta')) {
      $('maintFooterMeta').textContent = 'Templates Base de Operações (8 cenários cadastrados)';
    }
    if ($('btnSaveFooter')) {
      $('btnSaveFooter').textContent = 'Salvar Alterações';
      $('btnSaveFooter').disabled = true;
    }

    initTemplateScenarioSelect();
    renderTemplatesTable();
  }

  function initTemplateScenarioSelect() {
    var sel = $('templateScenarioSelect');
    if (!sel || !state.templateBlocksData || !state.templateBlocksData.cenarios) return;
    if (sel.options.length > 0) return;

    var cenarios = state.templateBlocksData.cenarios;
    Object.keys(cenarios).forEach(function (cKey) {
      var opt = document.createElement('option');
      opt.value = cKey;
      opt.textContent = cenarios[cKey].descricao || cKey;
      sel.appendChild(opt);
    });

    if (state.selectedTemplateScenario) {
      sel.value = state.selectedTemplateScenario;
    }

    sel.addEventListener('change', function () {
      state.selectedTemplateScenario = sel.value;
      renderTemplatesTable();
    });
  }

  function renderTemplatesTable() {
    var tbody = $('templateTableBody');
    if (!tbody) return;
    tbody.innerHTML = '';

    if (!state.templateBlocksData || !state.templateBlocksData.cenarios) {
      tbody.innerHTML = '<tr><td colspan="6" style="text-align:center; padding:30px; color:var(--text-dim);">Carregando blocos de templates…</td></tr>';
      return;
    }

    var selKey = state.selectedTemplateScenario || 'container_solar_essw_mecanica';
    var cenario = state.templateBlocksData.cenarios[selKey];
    if (!cenario || !cenario.tarefas) {
      tbody.innerHTML = '<tr><td colspan="6" style="text-align:center; padding:30px; color:var(--text-dim);">Nenhuma tarefa cadastrada neste cenário.</td></tr>';
      return;
    }

    var term = ($('templateSearch') ? $('templateSearch').value.trim().toLowerCase() : '');
    var filtered = cenario.tarefas.filter(function (t) {
      if (!term) return true;
      var matchT = String(t.tarefa || '').toLowerCase().indexOf(term) !== -1;
      var matchD = String(t.descricao_tarefa || '').toLowerCase().indexOf(term) !== -1;
      return matchT || matchD;
    });

    if ($('templateCountBadge')) {
      $('templateCountBadge').textContent = filtered.length + ' de ' + cenario.tarefas.length + ' tarefas';
    }

    filtered.forEach(function (t) {
      var tr = document.createElement('tr');
      var tCode = String(t.tarefa || '');
      if (/^\d+$/.test(tCode)) tCode = ('0000' + tCode).slice(-4);

      tr.innerHTML =
        '<td style="text-align:center; color:var(--text-faint); font-family:\'IBM Plex Mono\';">' + t.linha_template + '</td>' +
        '<td style="text-align:center;"><span class="cronograma-task-badge">' + escapeHtml(tCode) + '</span></td>' +
        '<td style="font-weight:500; color:var(--text);">' + escapeHtml(t.descricao_tarefa) + '</td>' +
        '<td style="text-align:center; font-family:\'IBM Plex Mono\'; font-weight:600; color:var(--amber);">' + t.duracao + '</td>' +
        '<td style="text-align:center; font-size:10.5px; color:var(--text-dim);">' + escapeHtml(t.unidade || 'DIA') + '</td>' +
        '<td style="text-align:right; font-family:\'IBM Plex Mono\'; font-weight:600; color:var(--accent);">' + t.trabalho + ' h</td>';
      tbody.appendChild(tr);
    });
  }

  if ($('templateSearch')) {
    $('templateSearch').addEventListener('input', function () {
      renderTemplatesTable();
    });
  }

  if ($('btnRestaurarTemplates')) {
    $('btnRestaurarTemplates').addEventListener('click', function () {
      if (state.templateBlocksOriginal) {
        state.templateBlocksData = JSON.parse(JSON.stringify(state.templateBlocksOriginal));
        renderTemplatesTable();
        showToast('Templates restaurados com sucesso.');
      }
    });
  }

  function marcarSeletorDirty() {
    state.seletorDirty = true;
    if ($('btnSaveFooter') && state.isSeletorActive) {
      $('btnSaveFooter').disabled = false;
    }
  }

  function renderSeletorTable() {
    var tbody = $('seletorTableBody');
    if (!tbody) return;
    tbody.innerHTML = '';

    if (!state.seletorData || !state.seletorData.length) {
      tbody.innerHTML = '<tr><td colspan="12" style="text-align:center; padding:30px; color:var(--text-dim);">Carregando combinações do Seletor…</td></tr>';
      return;
    }

    var term = ($('seletorSearch') ? $('seletorSearch').value.trim().toLowerCase() : '');
    var countVisible = 0;

    state.seletorData.forEach(function (row, idx) {
      var rowText = (
        (row['Tipo Estrutura'] || '') + ' ' +
        (row['Nº Módulos?'] || '') + ' ' +
        (row['Casa Máq.?'] || '') + ' ' +
        (row['Sist. Seg.?'] || '') + ' ' +
        (row['Teste SW?'] || '') + ' ' +
        (row['PEP Standard'] || '') + ' ' +
        (row['DR Eng Mec'] || '') + ' ' +
        (row['DR Eng Ele'] || '') + ' ' +
        (row['DR Acess'] || '') + ' ' +
        (row['DR Eletromec'] || '')
      ).toLowerCase();

      if (term && rowText.indexOf(term) === -1) return;
      countVisible++;

      var tr = document.createElement('tr');
      tr.style.borderBottom = '1px solid var(--border)';
      tr.style.transition = 'background .15s ease';
      tr.addEventListener('mouseenter', function () { tr.style.background = 'var(--panel-3)'; });
      tr.addEventListener('mouseleave', function () { tr.style.background = ''; });

      var mecMods = [];
      for (var m = 1; m <= 8; m++) {
        var dr = row['DR Mec ' + m] || row['DR Mec' + m];
        var alt = row['Alt Mec ' + m] || row['Alt Mec' + m];
        if (dr) mecMods.push('M' + m + ': ' + dr + (alt ? ' (' + alt + ')' : ''));
      }
      var mecModsStr = mecMods.length ? mecMods.slice(0, 2).join(', ') + (mecMods.length > 2 ? ' +' + (mecMods.length - 2) : '') : '-';

      tr.innerHTML =
        '<td style="padding:8px 12px; font-weight:600; color:var(--text); white-space:nowrap;">' + escapeHtml(row['Tipo Estrutura'] || '-') + '</td>' +
        '<td style="padding:8px 10px; text-align:center; font-family:\'IBM Plex Mono\'; font-weight:600;">' + escapeHtml(row['Nº Módulos?'] || '1') + '</td>' +
        '<td style="padding:8px 10px; text-align:center;">' + (row['Casa Máq.?'] === 'Sim' ? '<span style="color:var(--amber); font-weight:600;">Sim</span>' : '<span style="color:var(--text-dim);">Não</span>') + '</td>' +
        '<td style="padding:8px 10px; text-align:center;">' + (row['Sist. Seg.?'] === 'Sim' ? '<span style="color:var(--accent); font-weight:600;">Sim</span>' : '<span style="color:var(--text-dim);">' + (row['Sist. Seg.?'] || '-') + '</span>') + '</td>' +
        '<td style="padding:8px 10px; text-align:center;">' + (row['Teste SW?'] === 'Sim' ? '<span style="color:var(--accent); font-weight:600;">Sim</span>' : '<span style="color:var(--text-dim);">' + (row['Teste SW?'] || '-') + '</span>') + '</td>' +
        '<td style="padding:8px 12px; font-family:\'IBM Plex Mono\'; font-weight:700; color:var(--accent); white-space:nowrap;">' +
          '<input type="text" class="seletor-pep-ipt" data-idx="' + idx + '" value="' + escapeHtml(row['PEP Standard'] || '') + '" style="background:var(--panel-1); border:1px solid var(--border); color:var(--accent); font-family:inherit; font-weight:bold; padding:3px 6px; border-radius:4px; width:110px; font-size:11px;">' +
        '</td>' +
        '<td style="padding:8px 10px; font-family:\'IBM Plex Mono\'; font-size:11px; white-space:nowrap;">' + (row['DR Eng Mec'] ? row['DR Eng Mec'] + ' <span style="color:var(--text-dim); font-size:10px;">(' + (row['Alt Eng Mec'] || '1') + ')</span>' : '-') + '</td>' +
        '<td style="padding:8px 10px; font-family:\'IBM Plex Mono\'; font-size:11px; white-space:nowrap;">' + (row['DR Eng Ele'] ? row['DR Eng Ele'] + ' <span style="color:var(--text-dim); font-size:10px;">(' + (row['Alt Eng Ele'] || '1') + ')</span>' : '-') + '</td>' +
        '<td style="padding:8px 10px; font-family:\'IBM Plex Mono\'; font-size:10.5px; color:var(--text-dim); white-space:nowrap;" title="' + escapeHtml(mecMods.join('\n')) + '">' + escapeHtml(mecModsStr) + '</td>' +
        '<td style="padding:8px 10px; font-family:\'IBM Plex Mono\'; font-size:11px; white-space:nowrap;">' + (row['DR Acess'] ? row['DR Acess'] + ' <span style="color:var(--text-dim); font-size:10px;">(' + (row['Alt Acess'] || '1') + ')</span>' : '-') + '</td>' +
        '<td style="padding:8px 10px; font-family:\'IBM Plex Mono\'; font-size:11px; white-space:nowrap;">' + (row['DR Eletromec'] ? row['DR Eletromec'] + ' <span style="color:var(--text-dim); font-size:10px;">(' + (row['Alt Eletromec'] || '1') + ')</span>' : '-') + '</td>' +
        '<td style="padding:8px 10px; text-align:center;">' +
          '<button type="button" class="icon-btn btn-edit-seletor-row" data-idx="' + idx + '" title="Editar Centros de Trabalho desta combinação" style="padding:3px 8px; font-size:10.5px; border-radius:4px; border:1px solid var(--border); background:var(--panel-1); color:var(--text); cursor:pointer;">Editar CTs</button>' +
        '</td>';

      tbody.appendChild(tr);
    });

    if ($('seletorCountBadge')) {
      $('seletorCountBadge').textContent = countVisible + ' de ' + state.seletorData.length + ' combinações';
    }

    tbody.querySelectorAll('.seletor-pep-ipt').forEach(function (ipt) {
      ipt.addEventListener('change', function () {
        var rowIdx = parseInt(ipt.dataset.idx, 10);
        state.seletorData[rowIdx]['PEP Standard'] = ipt.value.trim();
        marcarSeletorDirty();
      });
    });

    tbody.querySelectorAll('.btn-edit-seletor-row').forEach(function (btn) {
      btn.addEventListener('click', function () {
        var rowIdx = parseInt(btn.dataset.idx, 10);
        abrirModalEditarLinhaSeletor(rowIdx);
      });
    });
  }

  if ($('seletorSearch')) {
    $('seletorSearch').addEventListener('input', renderSeletorTable);
  }

  if ($('btnRestaurarSeletor')) {
    $('btnRestaurarSeletor').addEventListener('click', function () {
      if (confirm('Deseja recarregar o Seletor do arquivo original?')) {
        carregarSeletor();
        showToast('Seletor recarregado!');
      }
    });
  }

  var currentEditingSeletorIdx = -1;

  function abrirModalNovoSeletorRow() {
    currentEditingSeletorIdx = -1;
    if ($('modalSeletorRowTitle')) $('modalSeletorRowTitle').textContent = 'Nova Combinação do Seletor';
    if ($('btnSalvarModalSeletorRow')) $('btnSalvarModalSeletorRow').textContent = 'Adicionar Combinação';
    if ($('btnExcluirModalSeletorRow')) $('btnExcluirModalSeletorRow').style.display = 'none';

    if ($('iptSelTipoEstrutura')) $('iptSelTipoEstrutura').value = 'Eletrocentro';
    if ($('iptSelNModulos')) $('iptSelNModulos').value = '1';
    if ($('iptSelCasaMaq')) $('iptSelCasaMaq').value = 'Não';
    if ($('iptSelSistSeg')) $('iptSelSistSeg').value = 'Não';
    if ($('iptSelTesteSW')) $('iptSelTesteSW').value = 'Não';
    if ($('iptSelPepStandard')) $('iptSelPepStandard').value = '';

    if ($('iptSelDrEngMec')) $('iptSelDrEngMec').value = '';
    if ($('iptSelAltEngMec')) $('iptSelAltEngMec').value = '1';
    if ($('iptSelDrEngEle')) $('iptSelDrEngEle').value = '';
    if ($('iptSelAltEngEle')) $('iptSelAltEngEle').value = '1';
    if ($('iptSelDrAcess')) $('iptSelDrAcess').value = '';
    if ($('iptSelAltAcess')) $('iptSelAltAcess').value = '1';
    if ($('iptSelDrEletromec')) $('iptSelDrEletromec').value = '';
    if ($('iptSelAltEletromec')) $('iptSelAltEletromec').value = '1';

    for (var m = 1; m <= 8; m++) {
      if ($('iptSelDrMec' + m)) $('iptSelDrMec' + m).value = '';
      if ($('iptSelAltMec' + m)) $('iptSelAltMec' + m).value = '1';
    }

    var modal = $('modalSeletorRow');
    if (modal) modal.classList.add('open');
  }

  function abrirModalEditarLinhaSeletor(rowIdx) {
    var row = state.seletorData[rowIdx];
    if (!row) return;
    currentEditingSeletorIdx = rowIdx;

    if ($('modalSeletorRowTitle')) {
      $('modalSeletorRowTitle').textContent = 'Editar Combinação: ' + (row['Tipo Estrutura'] || '') + ' (' + (row['Nº Módulos?'] || '1') + ' Mód)';
    }
    if ($('btnSalvarModalSeletorRow')) $('btnSalvarModalSeletorRow').textContent = 'Salvar Alterações';
    if ($('btnExcluirModalSeletorRow')) $('btnExcluirModalSeletorRow').style.display = 'inline-flex';

    if ($('iptSelTipoEstrutura')) $('iptSelTipoEstrutura').value = row['Tipo Estrutura'] || '';
    if ($('iptSelNModulos')) $('iptSelNModulos').value = String(row['Nº Módulos?'] || '1');
    if ($('iptSelCasaMaq')) $('iptSelCasaMaq').value = row['Casa Máq.?'] || 'Não';
    if ($('iptSelSistSeg')) $('iptSelSistSeg').value = row['Sist. Seg.?'] || 'Não';
    if ($('iptSelTesteSW')) $('iptSelTesteSW').value = row['Teste SW?'] || 'Não';
    if ($('iptSelPepStandard')) $('iptSelPepStandard').value = row['PEP Standard'] || '';

    if ($('iptSelDrEngMec')) $('iptSelDrEngMec').value = row['DR Eng Mec'] || '';
    if ($('iptSelAltEngMec')) $('iptSelAltEngMec').value = row['Alt Eng Mec'] || '1';
    if ($('iptSelDrEngEle')) $('iptSelDrEngEle').value = row['DR Eng Ele'] || '';
    if ($('iptSelAltEngEle')) $('iptSelAltEngEle').value = row['Alt Eng Ele'] || '1';
    if ($('iptSelDrAcess')) $('iptSelDrAcess').value = row['DR Acess'] || '';
    if ($('iptSelAltAcess')) $('iptSelAltAcess').value = row['Alt Acess'] || '1';
    if ($('iptSelDrEletromec')) $('iptSelDrEletromec').value = row['DR Eletromec'] || '';
    if ($('iptSelAltEletromec')) $('iptSelAltEletromec').value = row['Alt Eletromec'] || '1';

    for (var m = 1; m <= 8; m++) {
      if ($('iptSelDrMec' + m)) $('iptSelDrMec' + m).value = row['DR Mec ' + m] || row['DR Mec' + m] || '';
      if ($('iptSelAltMec' + m)) $('iptSelAltMec' + m).value = row['Alt Mec ' + m] || row['Alt Mec' + m] || '1';
    }

    var modal = $('modalSeletorRow');
    if (modal) modal.classList.add('open');
  }

  function fecharModalSeletorRow() {
    var modal = $('modalSeletorRow');
    if (modal) modal.classList.remove('open');
  }

  if ($('btnAdicionarLinhaSeletor')) {
    $('btnAdicionarLinhaSeletor').addEventListener('click', abrirModalNovoSeletorRow);
  }
  if ($('btnCloseModalSeletorRow')) {
    $('btnCloseModalSeletorRow').addEventListener('click', fecharModalSeletorRow);
  }
  if ($('btnCancelModalSeletorRow')) {
    $('btnCancelModalSeletorRow').addEventListener('click', fecharModalSeletorRow);
  }

  if ($('btnSalvarModalSeletorRow')) {
    $('btnSalvarModalSeletorRow').addEventListener('click', function () {
      var tipo = ($('iptSelTipoEstrutura') ? $('iptSelTipoEstrutura').value : '').trim();
      var pep = ($('iptSelPepStandard') ? $('iptSelPepStandard').value : '').trim();
      if (!tipo) {
        showToast('Informe o Tipo de Estrutura para esta combinação.', true);
        return;
      }
      if (!pep) {
        showToast('Informe o código do PEP Standard.', true);
        return;
      }

      var rowData = {
        "Tipo Estrutura": tipo,
        "Nº Módulos?": $('iptSelNModulos') ? $('iptSelNModulos').value : '1',
        "Casa Máq.?": $('iptSelCasaMaq') ? $('iptSelCasaMaq').value : 'Não',
        "Sist. Seg.?": $('iptSelSistSeg') ? $('iptSelSistSeg').value : 'Não',
        "Teste SW?": $('iptSelTesteSW') ? $('iptSelTesteSW').value : 'Não',
        "PEP Standard": pep,
        "DR Eng Mec": $('iptSelDrEngMec') ? $('iptSelDrEngMec').value.trim() : '',
        "Alt Eng Mec": $('iptSelAltEngMec') ? $('iptSelAltEngMec').value.trim() : '1',
        "DR Eng Ele": $('iptSelDrEngEle') ? $('iptSelDrEngEle').value.trim() : '',
        "Alt Eng Ele": $('iptSelAltEngEle') ? $('iptSelAltEngEle').value.trim() : '1',
        "DR Acess": $('iptSelDrAcess') ? $('iptSelDrAcess').value.trim() : '',
        "Alt Acess": $('iptSelAltAcess') ? $('iptSelAltAcess').value.trim() : '1',
        "DR Eletromec": $('iptSelDrEletromec') ? $('iptSelDrEletromec').value.trim() : '',
        "Alt Eletromec": $('iptSelAltEletromec') ? $('iptSelAltEletromec').value.trim() : '1'
      };

      for (var m = 1; m <= 8; m++) {
        var drMVal = $('iptSelDrMec' + m) ? $('iptSelDrMec' + m).value.trim() : '';
        var altMVal = $('iptSelAltMec' + m) ? $('iptSelAltMec' + m).value.trim() : '1';
        rowData['DR Mec ' + m] = drMVal;
        rowData['Alt Mec ' + m] = altMVal;
      }

      if (currentEditingSeletorIdx >= 0 && currentEditingSeletorIdx < state.seletorData.length) {
        state.seletorData[currentEditingSeletorIdx] = rowData;
        showToast('Combinação atualizada no Seletor.');
      } else {
        state.seletorData.push(rowData);
        showToast('Nova combinação adicionada ao Seletor.');
      }

      marcarSeletorDirty();
      renderSeletorTable();
      fecharModalSeletorRow();
    });
  }

  if ($('btnExcluirModalSeletorRow')) {
    $('btnExcluirModalSeletorRow').addEventListener('click', function () {
      if (currentEditingSeletorIdx >= 0 && currentEditingSeletorIdx < state.seletorData.length) {
        if (confirm('Deseja realmente remover esta combinação do Seletor?')) {
          state.seletorData.splice(currentEditingSeletorIdx, 1);
          marcarSeletorDirty();
          renderSeletorTable();
          fecharModalSeletorRow();
          showToast('Combinação removida do Seletor.');
        }
      }
    });
  }

  function carregarSeletor() {
    if (isPyWebviewAvailable()) {
      window.pywebview.api.get_seletor().then(function (data) {
        if (data && data.length) {
          state.seletorData = data;
          state.seletorOriginal = JSON.parse(JSON.stringify(data));
          atualizarSugestaoPep();
          if (state.isSeletorActive) renderSeletorTable();
        } else {
          fetchSeletorFallback();
        }
      }).catch(function () {
        fetchSeletorFallback();
      });
    } else {
      fetchSeletorFallback();
    }
  }

  function fetchSeletorFallback() {
    fetch('seletor.json').then(function (resp) {
      if (resp.ok) return resp.json();
      throw new Error('Status ' + resp.status);
    }).then(function (data) {
      if (data && data.length) {
        state.seletorData = data;
        state.seletorOriginal = JSON.parse(JSON.stringify(data));
        atualizarSugestaoPep();
        if (state.isSeletorActive) renderSeletorTable();
      }
    }).catch(function (err) {
      console.warn('[Seletor] Erro ao carregar seletor.json:', err);
    });
  }

  function carregarTemplateBlocks() {
    if (isPyWebviewAvailable() && window.pywebview.api.get_template_blocks) {
      window.pywebview.api.get_template_blocks().then(function (data) {
        if (data && data.cenarios && Object.keys(data.cenarios).length) {
          state.templateBlocksData = data;
          state.templateBlocksOriginal = JSON.parse(JSON.stringify(data));
          console.log('[TemplateBlocks] Carregado via pywebview (' + Object.keys(data.cenarios).length + ' cenários)');
          if (state.isTemplatesActive) {
            initTemplateScenarioSelect();
            renderTemplatesTable();
          }
        } else {
          fetchTemplateBlocksFallback();
        }
      }).catch(function (err) {
        console.warn('[TemplateBlocks] Erro pywebview:', err);
        fetchTemplateBlocksFallback();
      });
    } else {
      fetchTemplateBlocksFallback();
    }
  }

  function fetchTemplateBlocksFallback() {
    fetch('template_blocks.json').then(function (resp) {
      if (resp.ok) return resp.json();
      throw new Error('Status ' + resp.status);
    }).then(function (data) {
      if (data && data.cenarios) {
        state.templateBlocksData = data;
        state.templateBlocksOriginal = JSON.parse(JSON.stringify(data));
        console.log('[TemplateBlocks] Carregado via fetch (' + Object.keys(data.cenarios).length + ' cenários)');
        if (state.isTemplatesActive) {
          initTemplateScenarioSelect();
          renderTemplatesTable();
        }
      }
    }).catch(function (err) {
      console.warn('[TemplateBlocks] Erro ao carregar template_blocks.json via fetch:', err);
      apiCall('/get_template_blocks').then(function (data) {
        if (data && data.cenarios) {
          state.templateBlocksData = data;
          state.templateBlocksOriginal = JSON.parse(JSON.stringify(data));
          console.log('[TemplateBlocks] Carregado via /api/get_template_blocks');
          if (state.isTemplatesActive) {
            initTemplateScenarioSelect();
            renderTemplatesTable();
          }
        }
      }).catch(function () {});
    });
  }

  function consultarSeletor(ctx) {
    if (!state.seletorData || !state.seletorData.length || !ctx) return null;

    var tipo = ctx.tipoestrutura || '';
    var nmod = String(ctx.nmod || '1');
    var isRoofTop = ctx.tipomaq === 'Roof Top' || ctx.tipomaq === 'Self + Dutos';
    var casaMaq = isRoofTop ? 'Sim' : 'Não';
    var temSeg = (ctx.seguranca && ctx.seguranca !== 'Não possui' && ctx.seguranca !== 'Não aplicável') ? 'Sim' : 'Não';
    var testeSW = ctx.testesw ? 'Sim' : 'Não';
    var isFab1313 = ctx.fab1313 ? true : false;

    var tipoSeletor = tipo;
    if (tipo === 'Móvel' || tipo === 'Semimóvel' || tipo === 'Modular' || tipo === 'Fixo' || tipo === 'Embarcado' || tipo === 'Eletrocentro') {
      tipoSeletor = isFab1313 ? 'EletrocentroB' : 'Eletrocentro';
    } else if (tipo === 'Container Solar') {
      if (ctx.progReles && ctx.diagBTI) {
        tipoSeletor = 'Container   MarítimoBEI';
      } else if (ctx.progReles) {
        tipoSeletor = 'Container   MarítimoBE';
      } else if (ctx.diagBTI) {
        tipoSeletor = 'Container   MarítimoBI';
      } else if (ctx.diagAgrup) {
        tipoSeletor = 'Container   MarítimoBII';
      } else {
        tipoSeletor = 'Container   Marítimo';
      }
    } else if (tipo === 'Skid (mecânica)') {
      tipoSeletor = isFab1313 ? 'Skid   (mecânica)B' : 'Skid   (mecânica)';
    } else if (tipo === 'Skid (com elétrica)') {
      tipoSeletor = isFab1313 ? 'Skid (com   elétrica)B' : 'Skid (com   elétrica)';
    } else if (tipo === 'Pilotis') {
      tipoSeletor = isFab1313 ? 'PilotisB' : 'Pilotis';
    } else if (tipo === 'ESSW (mecânica)') {
      tipoSeletor = 'ESSW   (mecânica)';
    } else if (tipo === 'ESSW (elétrica)') {
      tipoSeletor = 'ESSW   (elétrica)';
    } else if (tipo === 'Serviço Engenharia') {
      tipoSeletor = 'Serviço   Engenharia';
    }

    var cleanStr = function (s) { return (s || '').toString().replace(/\s+/g, ' ').trim().toLowerCase(); };
    var targetTipo = cleanStr(tipoSeletor);

    var bestMatch = null;
    var bestScore = -1;

    for (var i = 0; i < state.seletorData.length; i++) {
      var row = state.seletorData[i];
      var rowTipo = cleanStr(row['Tipo_Estrutura'] || row['Tipo Estrutura']);
      var rowMod = String(row['N_Modulos'] || row['Nº Módulos?'] || row['N Mdulos?'] || '1').trim();
      var rowCasaMaq = String(row['Casa_Maq'] || row['Casa Máq.?'] || row['Casa Mq.?'] || '').trim();
      var rowSistSeg = String(row['Sist_Seg'] || row['Sist. Seg.?'] || '').trim();
      var rowTesteSW = String(row['Teste_SW'] || row['Teste SW?'] || '').trim();

      var score = 0;

      if (rowTipo === targetTipo) {
        score += 100;
      } else if (targetTipo.indexOf('eletrocentro') !== -1 && rowTipo.indexOf('eletrocentro') !== -1) {
        score += 50;
      } else if (targetTipo.indexOf('container') !== -1 && rowTipo.indexOf('container') !== -1) {
        score += 50;
      } else if (targetTipo.indexOf('skid') !== -1 && rowTipo.indexOf('skid') !== -1) {
        score += 50;
      } else {
        continue;
      }

      if (rowMod === nmod) score += 30;
      if (rowCasaMaq === casaMaq) score += 20;
      if (rowSistSeg === temSeg) score += 15;
      if (rowTesteSW === testeSW) score += 10;

      if (score > bestScore) {
        bestScore = score;
        bestMatch = row;
      }
    }

    return bestMatch;
  }

  function atualizarSugestaoPep() {
    if (typeof coletarContextoFormulario !== 'function') return;
    var currentCtx = coletarContextoFormulario();
    var seletorMatch = consultarSeletor(currentCtx);
    var seletorBadge = $('seletorPepBadge');
    if (seletorBadge) {
      if (seletorMatch && seletorMatch['PEP Standard']) {
        var pepSug = seletorMatch['PEP Standard'];
        if ($('txtSeletorPep')) $('txtSeletorPep').textContent = pepSug;
        seletorBadge.style.display = 'inline-flex';
      } else {
        seletorBadge.style.display = 'none';
      }
    }
  }

  if ($('seletorPepBadge')) {
    $('seletorPepBadge').addEventListener('click', function () {
      var pepInput = $('pep');
      var txt = $('txtSeletorPep') ? $('txtSeletorPep').textContent.trim() : '';
      if (pepInput && txt && txt !== '-') {
        pepInput.value = txt;
        recomputeForm();
        showToast('PEP Standard aplicado com sucesso: ' + txt);
      }
    });
  }

  function salvarSeletor(motivo, anexosPayload) {
    var anexoBase64 = (anexosPayload && anexosPayload.length) ? anexosPayload[0].base64 : null;
    var anexoNome = (anexosPayload && anexosPayload.length) ? anexosPayload[0].nome : null;

    var payload = {
      seletor: state.seletorData,
      motivo: motivo,
      anexo_base64: anexoBase64,
      anexo_nome: anexoNome
    };

    showToast('Salvando Seletor de PEP & Centros de Trabalho...');

    var afterSaveSuccess = function (res) {
      state.seletorOriginal = JSON.parse(JSON.stringify(state.seletorData));
      state.seletorDirty = false;
      if ($('btnSaveFooter')) $('btnSaveFooter').disabled = true;
      showToast('Seletor de PEP & CTs salvo com sucesso no banco e sincronizado!');
      carregarSeletor();
      carregarHistorico();
    };

    if (isPyWebviewAvailable()) {
      window.pywebview.api.save_seletor(payload).then(function (res) {
        if (res && res.status === 'conflito') {
          showToast('⚠️ Conflito no Seletor: ' + (res.message || 'Outro usuário salvou antes.'), true);
          alert('⚠️ Conflito de Salvamento:\n\n' + (res.message || 'Outro usuário salvou o Seletor antes.') + '\n\nOs dados do Seletor serão recarregados do servidor.');
          carregarSeletor();
        } else if (res && res.status === 'success') {
          afterSaveSuccess(res);
        } else {
          showToast('Erro ao salvar Seletor: ' + ((res && res.message) || 'Falha no backend'), true);
        }
      }).catch(function (err) {
        showToast('Erro ao salvar Seletor via pywebview: ' + (err.message || err), true);
      });
    } else {
      apiCall('/save_seletor', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      }).then(function (res) {
        if (res && res.status === 'conflito') {
          showToast('⚠️ Conflito no Seletor: ' + (res.message || 'Outro usuário salvou antes.'), true);
          alert('⚠️ Conflito de Salvamento:\n\n' + (res.message || 'Outro usuário salvou o Seletor antes.') + '\n\nOs dados do Seletor serão recarregados do servidor.');
          carregarSeletor();
        } else if (res && res.status === 'success') {
          afterSaveSuccess(res);
        } else {
          showToast('Erro ao salvar Seletor: ' + ((res && res.message) || 'Falha na API'), true);
        }
      }).catch(function (err) {
        showToast('Erro ao salvar Seletor via API: ' + (err.message || err), true);
      });
    }
  }

  function prepararDirtySubTab() {
    var areaObj = state.regrasData[state.selectedAreaIdx];
    if (!areaObj || !state.selectedCampoKey) {
      state.dirtySubTabRule = null;
      return;
    }
    var campoObj = areaObj.campos[state.selectedCampoKey];
    if (!campoObj) {
      state.dirtySubTabRule = null;
      return;
    }
    if (!campoObj[state.selectedSubTab]) {
      campoObj[state.selectedSubTab] = {
        base: { forma: 'constante', valor: 0 },
        condicoes: []
      };
    }
    state.dirtySubTabRule = JSON.parse(JSON.stringify(campoObj[state.selectedSubTab]));
  }

  function resumoRegraSubTab(ruleObj) {
    if (!ruleObj || !ruleObj.base) return 'não configurado';
    var b = ruleObj.base;
    var str = '';
    if (b.forma === 'constante') str = (b.valor !== undefined ? b.valor : (b.valor_base || 0)) + ' fixo';
    else if (b.forma === 'multiplicativa') {
      var esc = b.escala !== undefined ? b.escala : 0.5;
      str = (b.valor_base !== undefined ? b.valor_base : (b.valor || 0)) + ' base (+' + esc + '×/mód)';
      if (b.adicao_final) str += ' + ' + b.adicao_final + ' (2m+)';
      if (b.subtracao_final) str += ' - ' + b.subtracao_final + ' (2m+)';
    }
    else if (b.forma === 'degrau_fixo') {
      var fat = b.fator !== undefined ? b.fator : (b.escala !== undefined ? b.escala : 1.5);
      str = (b.valor_base !== undefined ? b.valor_base : (b.valor || 0)) + ' base (1m) → ' + fat + '× fixo (2m+)';
      if (b.adicao_final) str += ' + ' + b.adicao_final;
      if (b.subtracao_final) str += ' - ' + b.subtracao_final;
    }
    else if (b.forma === 'tabela') str = 'Valores individuais por módulo (1m-8m)';
    else if (b.forma === 'aditiva') str = (b.valor_base || 0) + ' + ' + (b.passo || 0) + '/mód';
    else if (b.forma === 'derivado_h') {
      var etapas = b.etapas || [];
      if (!etapas.length) {
        if (b.divisao) str = '(H / ' + b.divisao + ') ⌈ceil⌉' + (b.subtracao ? ' - ' + b.subtracao : '');
        else str = 'Fórmula H';
      } else {
        str = 'H → ';
        var parts = [];
        etapas.forEach(function (s) {
          if (s.tipo === 'dividir') parts.push('/ ' + s.valor);
          else if (s.tipo === 'multiplicar') parts.push('× ' + s.valor);
          else if (s.tipo === 'somar') parts.push('+ ' + s.valor);
          else if (s.tipo === 'subtrair') parts.push('- ' + s.valor);
          else if (s.tipo === 'arredondar') parts.push((s.modo === 'baixo' ? '⌊floor⌋' : (s.modo === 'padrao' ? 'round' : '⌈ceil⌉')));
          else if (s.tipo === 'limitar_max') parts.push('≤ ' + s.valor);
          else if (s.tipo === 'limitar_min') parts.push('≥ ' + s.valor);
        });
        str += parts.join(' ');
      }
    }
    else if (b.forma === 'soma_campos') {
      var cLista = b.campos || [];
      str = 'Σ(' + cLista.join(' + ') + ')';
      var etapas = b.etapas || [];
      if (etapas.length) {
        var parts = [];
        etapas.forEach(function (s) {
          if (s.tipo === 'dividir') parts.push('/ ' + s.valor);
          else if (s.tipo === 'multiplicar') parts.push('× ' + s.valor);
          else if (s.tipo === 'somar') parts.push('+ ' + s.valor);
          else if (s.tipo === 'subtrair') parts.push('- ' + s.valor);
          else if (s.tipo === 'arredondar') parts.push((s.modo === 'baixo' ? '⌊floor⌋' : (s.modo === 'padrao' ? 'round' : '⌈ceil⌉')));
          else if (s.tipo === 'limitar_max') parts.push('≤ ' + s.valor);
          else if (s.tipo === 'limitar_min') parts.push('≥ ' + s.valor);
        });
        str += ' ' + parts.join(' ');
      }
    } else str = b.forma;
    var cCount = (ruleObj.condicoes || []).length;
    if (cCount > 0) str += ' (+ ' + cCount + ' cond)';
    if (b.perfis && Array.isArray(b.perfis) && b.perfis.length > 0) str += ' [CS/ESSW]';
    return str;
  }

  function renderList() {
    var wrap = $('colList');
    if (!wrap) return;
    var areaObj = state.regrasData[state.selectedAreaIdx];
    if ($('legendWrapper')) $('legendWrapper').style.display = 'flex';
    if (!areaObj) {
      wrap.innerHTML = '<div class="list-head"><h2>Campos</h2><span>Nenhum campo</span></div>';
      return;
    }
    var camposKeys = Object.keys(areaObj.campos || {});
    var html = '<div class="list-head"><h2>Campos</h2><span>' + camposKeys.length + ' em ' + areaObj.area + '</span></div>';
    camposKeys.forEach(function (key) {
      var campoObj = areaObj.campos[key] || {};
      var hResumo = resumoRegraSubTab(campoObj.H);
      var durResumo = resumoRegraSubTab(campoObj.DUR);
      var isSelected = (state.selectedCampoKey === key);

      html += '<div class="field-row' + (isSelected ? ' active' : '') + '" data-campokey="' + key + '">' +
        '<div class="fr-top"><span class="fr-name">' + key + '</span>' +
        '<span class="badge simples">H / DUR</span></div>' +
        '<span class="fr-meta"><b>H: ' + hResumo + ' | DUR: ' + durResumo + '</b></span>' +
        '</div>';
    });
    wrap.innerHTML = html;
    wrap.querySelectorAll('.field-row').forEach(function (row) {
      row.addEventListener('click', function () {
        state.selectedCampoKey = row.dataset.campokey;
        state.selectedSubTab = 'H';
        prepararDirtySubTab();
        renderList();
        renderEditor();
      });
    });
  }

  /* ==========================================================================
     PERFIS DE ESTRUTURA (Container Solar, ESSW etc.)
     Override universal aplicado ANTES da lógica normal de "forma" (aditiva,
     multiplicativa, tabela, constante...). Serve para estruturas que não
     seguem a escala por nº de módulos (ex.: "Container Solar", "ESSW
     (mecânica)", "ESSW (elétrica)"), que na planilha original ocupam linhas
     próprias e independentes da tabela de 1 a 8 módulos.
     Reaproveita o mesmo mecanismo de "cond" (c/o/val/j) já usado em
     montagens/blocos. Se nenhum perfil casar, cai no comportamento padrão
     (100% retrocompatível — campos sem "perfis" não são afetados).
     ========================================================================== */
  function matchedPerfil(perfis, simCtx) {
    if (!perfis || !Array.isArray(perfis)) return null;
    for (var i = 0; i < perfis.length; i++) {
      if (condOkBloco(perfis[i], simCtx)) return perfis[i];
    }
    return null;
  }

  function getEspeciais(subTabRule) {
    if (!subTabRule) return { solar: { ativo: false, base: { forma: 'constante', valor: 0 }, herdar_condicoes: true, condicoes: [] }, essw: { ativo: false, base: { forma: 'constante', valor: 0 }, herdar_condicoes: true, condicoes: [] } };

    if (!subTabRule.especiais) {
      subTabRule.especiais = {
        solar: { ativo: false, base: { forma: 'constante', valor: 0 }, herdar_condicoes: true, condicoes: [] },
        essw: { ativo: false, base: { forma: 'constante', valor: 0 }, herdar_condicoes: true, condicoes: [] }
      };

      if (subTabRule.base && Array.isArray(subTabRule.base.perfis)) {
        subTabRule.base.perfis.forEach(function (p) {
          if (!p || !p.cond) return;
          var isCs = p.cond.some(function (c) { return c.c === 'tipoestrutura' && c.val === 'Container Solar'; });
          var isEssw = p.cond.some(function (c) { return c.c === 'tipoestrutura' && (c.val === 'ESSW (mecânica)' || c.val === 'ESSW (elétrica)' || c.val === 'ESSW'); });

          var v = 0;
          if (p.it && Array.isArray(p.it) && p.it.length) {
            if (p.it.length === 1 && p.it[0].t === 'num') v = p.it[0].v;
            else v = chainExpr(p.it, false, SIM_CTX);
          } else if (p.valor !== undefined) {
            v = p.valor;
          }

          if (isCs) {
            subTabRule.especiais.solar.ativo = true;
            subTabRule.especiais.solar.base = { forma: 'constante', valor: v, valor_base: v };
          }
          if (isEssw) {
            subTabRule.especiais.essw.ativo = true;
            subTabRule.especiais.essw.base = { forma: 'constante', valor: v, valor_base: v };
          }
        });
      }
    }

    ['solar', 'essw'].forEach(function (k) {
      if (!subTabRule.especiais[k]) {
        subTabRule.especiais[k] = { ativo: false, base: { forma: 'constante', valor: 0 }, herdar_condicoes: true, condicoes: [] };
      }
      if (!subTabRule.especiais[k].base) {
        subTabRule.especiais[k].base = { forma: 'constante', valor: 0 };
      }
      if (subTabRule.especiais[k].herdar_condicoes === undefined) {
        subTabRule.especiais[k].herdar_condicoes = true;
      }
      if (!subTabRule.especiais[k].condicoes) {
        subTabRule.especiais[k].condicoes = [];
      }
    });

    return subTabRule.especiais;
  }

  function syncEspeciaisToPerfis(subTabRule) {
    var esp = getEspeciais(subTabRule);
    var base = subTabRule.base;
    if (!base) return;

    var perfis = [];
    if (esp.solar && esp.solar.ativo) {
      var valS = parseFloat(String(esp.solar.base.valor).replace(',', '.')) || 0;
      perfis.push({
        id: 'perfil_cs',
        nome: 'Container Solar',
        cond: [{ c: 'tipoestrutura', o: '=', val: 'Container Solar', j: 'E' }],
        it: [{ t: 'num', v: valS }],
        valor: valS
      });
    }
    if (esp.essw && esp.essw.ativo) {
      var valE = parseFloat(String(esp.essw.base.valor).replace(',', '.')) || 0;
      perfis.push({
        id: 'perfil_essw',
        nome: 'ESSW',
        cond: [
          { c: 'tipoestrutura', o: '=', val: 'ESSW (mecânica)', j: 'E' },
          { c: 'tipoestrutura', o: '=', val: 'ESSW (elétrica)', j: 'OU' }
        ],
        it: [{ t: 'num', v: valE }],
        valor: valE
      });
    }
    if (perfis.length > 0) {
      base.perfis = perfis;
    } else {
      delete base.perfis;
    }
  }

  function espCondRowHtml(c, i, espKey) {
    var opts = FLAGS.map(function (f) { return '<option value="' + f.key + '"' + (f.key === c.flag ? ' selected' : '') + '>' + f.nome + '</option>'; }).join('');
    var valorAtual = c.forma === 'tabela' ? (c.valores ? c.valores[0] : 0) : (c.valor !== undefined ? c.valor : 0);

    var formaOptions = [
      { tipo: 'fixo', label: '＋ Valor Fixo (+V fixo)' },
      { tipo: 'escala_multiplicativa', label: 'Escala Multiplicativa' },
      { tipo: 'subtrair', label: '－ Subtrair Fixo (-V fixo)' },
      { tipo: 'multiplicar', label: '× Multiplicar Fixo (× V fixo)' },
      { tipo: 'dividir', label: '÷ Dividir Fixo (÷ V fixo)' }
    ];

    var cForma = c.forma || 'fixo';
    var selectFormaHtml = '<select class="cond-select esp-cond-forma" data-esp="' + espKey + '" data-idx="' + i + '" style="max-width:170px; font-weight:500;">' +
      formaOptions.map(function (opt) {
        return '<option value="' + opt.tipo + '"' + (opt.tipo === cForma ? ' selected' : '') + '>' + opt.label + '</option>';
      }).join('') +
      '</select>';

    var valorInputHtml = '<span class="cr-txt">opera</span><input type="text" class="cond-num esp-cond-valor" data-esp="' + espKey + '" data-idx="' + i + '" value="' + valorAtual + '" title="Aceita números ou expressões" style="width:75px;">';

    return '<div class="esp-cond-row" data-esp="' + espKey + '" data-idx="' + i + '" style="flex-wrap:wrap; background:var(--panel-1); padding:6px 8px; border-radius:6px; margin-top:4px; border:1px solid var(--border);">' +
      '<span class="cr-txt">Se</span><select class="cond-select esp-cond-flag" data-esp="' + espKey + '" data-idx="' + i + '" style="max-width:180px;">' + opts + '</select>' +
      valorInputHtml +
      selectFormaHtml +
      '<button type="button" class="cond-remove esp-cond-remove" data-esp="' + espKey + '" data-idx="' + i + '"><svg viewBox="0 0 24 24" width="12" height="12" fill="none" stroke="currentColor" stroke-width="2.4"><path d="M6 6l12 12M18 6L6 18"/></svg></button>' +
      '</div>';
  }

  function perfisBlockHtml(subTabRule, subTab) {
    var esp = getEspeciais(subTabRule);
    var unit = subTab === 'H' ? 'h' : 'dias';
    var hasActive = (esp.solar && esp.solar.ativo) || (esp.essw && esp.essw.ativo);

    var html = '<div class="field-block special-perfis-accordion" style="margin-top:16px; border:1px solid var(--border); border-radius:10px; overflow:hidden; background:var(--panel-1);">' +
      '<div class="accordion-head" id="toggleSpecialRules" style="display:flex; justify-content:space-between; align-items:center; padding:12px 14px; background:var(--panel-2); cursor:pointer; user-select:none; border-bottom:1px solid var(--border);">' +
      '<div style="display:flex; align-items:center; gap:8px;">' +
      '<span style="display:inline-flex; align-items:center; justify-content:center; width:22px; height:22px; border-radius:6px; background:var(--accent-dim); color:var(--accent);"><svg viewBox="0 0 24 24" width="13" height="13" fill="none" stroke="currentColor" stroke-width="2.2"><path d="M4 21v-7M4 10V3M12 21v-9M12 8V3M20 21v-5M20 12V3M1 14h6M9 8h6M17 16h6"/></svg></span>' +
      '<span style="font-size:12.5px; font-weight:700; color:var(--text);">Regras de Estruturas Especiais (Container Solar & ESSW)</span>' +
      (hasActive ? '<span class="badge" style="font-size:10px; padding:2px 6px; border-radius:4px; font-weight:600; background:var(--accent-dim); color:var(--accent);">Personalizado</span>' : '<span class="badge" style="font-size:10px; padding:2px 6px; border-radius:4px; font-weight:600; background:var(--panel-3); color:var(--text-faint);">Padrão</span>') +
      '</div>' +
      '<div style="display:flex; align-items:center; gap:8px;">' +
      '<span style="font-size:11px; color:var(--text-faint); font-family:IBM Plex Mono;">Clique para expandir / ocultar</span>' +
      '<span class="acc-chevron" id="accChevron" style="font-size:12px; color:var(--text-dim); transition:transform 0.2s ease;">' + (state.specialRulesExpanded ? '▲' : '▼') + '</span>' +
      '</div>' +
      '</div>' +

      '<div class="accordion-body" id="specialRulesBody" style="display:' + (state.specialRulesExpanded ? 'block' : 'none') + '; padding:14px; background:var(--panel-1);">' +
      '<div style="display:grid; grid-template-columns: 1fr 1fr; gap:14px;">';

    // 1. Container Solar Card
    var cs = esp.solar;
    var csVal = cs.base ? (cs.base.valor !== undefined ? cs.base.valor : (cs.base.valor_base || 0)) : 0;
    var csForma = cs.base ? (cs.base.forma || 'constante') : 'constante';

    html += '<div class="esp-card" style="border:1px solid ' + (cs.ativo ? 'var(--accent)' : 'var(--border)') + '; border-radius:8px; padding:12px; background:var(--panel-2); display:flex; flex-direction:column; gap:10px;">' +
      '<div style="display:flex; justify-content:space-between; align-items:center;">' +
      '<label style="display:flex; align-items:center; gap:7px; font-weight:700; font-size:12.5px; cursor:pointer;">' +
      '<input type="checkbox" id="chkEspSolar"' + (cs.ativo ? ' checked' : '') + ' style="accent-color:var(--accent); cursor:pointer;">' +
      '<span>Container Solar</span>' +
      '</label>' +
      '<span class="badge" style="font-size:10px; padding:2px 6px; border-radius:4px; font-weight:600; background:' + (cs.ativo ? 'var(--accent-dim)' : 'var(--panel-3)') + '; color:' + (cs.ativo ? 'var(--accent)' : 'var(--text-faint)') + ';">' + (cs.ativo ? 'Ativo' : 'Desativado') + '</span>' +
      '</div>';

    if (cs.ativo) {
      html += '<div style="display:flex; flex-direction:column; gap:10px; border-top:1px solid var(--border); padding-top:10px;">' +
        '<div>' +
        '<div style="font-size:11px; font-weight:600; color:var(--text-dim); margin-bottom:4px;">Cálculo Base (Solar):</div>' +
        '<div style="display:flex; gap:6px;">' +
        '<select id="selFormaEspSolar" class="ipt" style="padding:4px 8px; font-size:11.5px; max-width:130px;">' +
        '<option value="constante"' + (csForma === 'constante' ? ' selected' : '') + '>Valor Fixo</option>' +
        '<option value="derivado_h"' + (csForma === 'derivado_h' ? ' selected' : '') + '>Fórmula (H)</option>' +
        '</select>' +
        '<div class="num-field" style="flex:1;">' +
        '<input type="text" id="iptValEspSolar" value="' + csVal + '" placeholder="ex: 16.2" style="font-weight:600;">' +
        '<span>' + unit + '</span>' +
        '</div>' +
        '</div>' +
        '</div>' +

        '<div style="display:flex; flex-direction:column; gap:6px;">' +
        '<label style="display:flex; align-items:center; gap:6px; font-size:11.5px; color:var(--text); cursor:pointer;">' +
        '<input type="checkbox" id="chkHerdarSolar"' + (cs.herdar_condicoes ? ' checked' : '') + ' style="accent-color:var(--accent); cursor:pointer;">' +
        '<span>Aplicar condições gerais da regra padrão</span>' +
        '</label>' +
        '</div>' +

        '<div style="display:flex; flex-direction:column; gap:6px;">' +
        '<div style="display:flex; justify-content:space-between; align-items:center;">' +
        '<span style="font-size:11px; font-weight:600; color:var(--text-dim);">Condições Específicas para Solar (' + cs.condicoes.length + '):</span>' +
        '<button type="button" class="btn btn-add-esp-cond" data-esp="solar" style="font-size:10.5px; padding:2px 6px;">＋ Condição</button>' +
        '</div>' +
        '<div id="listEspCondSolar" style="display:flex; flex-direction:column; gap:4px;">';

      cs.condicoes.forEach(function (c, i) {
        html += espCondRowHtml(c, i, 'solar');
      });

      html += '</div></div></div>';
    } else {
      html += '<div style="font-size:11px; color:var(--text-faint); font-style:italic;">Utiliza a regra e condições padrão de 1 a 8 módulos.</div>';
    }
    html += '</div>';

    // 2. ESSW Card
    var essw = esp.essw;
    var esswVal = essw.base ? (essw.base.valor !== undefined ? essw.base.valor : (essw.base.valor_base || 0)) : 0;
    var esswForma = essw.base ? (essw.base.forma || 'constante') : 'constante';

    html += '<div class="esp-card" style="border:1px solid ' + (essw.ativo ? 'var(--accent)' : 'var(--border)') + '; border-radius:8px; padding:12px; background:var(--panel-2); display:flex; flex-direction:column; gap:10px;">' +
      '<div style="display:flex; justify-content:space-between; align-items:center;">' +
      '<label style="display:flex; align-items:center; gap:7px; font-weight:700; font-size:12.5px; cursor:pointer;">' +
      '<input type="checkbox" id="chkEspEssw"' + (essw.ativo ? ' checked' : '') + ' style="accent-color:var(--accent); cursor:pointer;">' +
      '<span>ESSW (Mecânica / Elétrica)</span>' +
      '</label>' +
      '<span class="badge" style="font-size:10px; padding:2px 6px; border-radius:4px; font-weight:600; background:' + (essw.ativo ? 'var(--accent-dim)' : 'var(--panel-3)') + '; color:' + (essw.ativo ? 'var(--accent)' : 'var(--text-faint)') + ';">' + (essw.ativo ? 'Ativo' : 'Desativado') + '</span>' +
      '</div>';

    if (essw.ativo) {
      html += '<div style="display:flex; flex-direction:column; gap:10px; border-top:1px solid var(--border); padding-top:10px;">' +
        '<div>' +
        '<div style="font-size:11px; font-weight:600; color:var(--text-dim); margin-bottom:4px;">Cálculo Base (ESSW):</div>' +
        '<div style="display:flex; gap:6px;">' +
        '<select id="selFormaEspEssw" class="ipt" style="padding:4px 8px; font-size:11.5px; max-width:130px;">' +
        '<option value="constante"' + (esswForma === 'constante' ? ' selected' : '') + '>Valor Fixo</option>' +
        '<option value="derivado_h"' + (esswForma === 'derivado_h' ? ' selected' : '') + '>Fórmula (H)</option>' +
        '</select>' +
        '<div class="num-field" style="flex:1;">' +
        '<input type="text" id="iptValEspEssw" value="' + esswVal + '" placeholder="ex: 3.24" style="font-weight:600;">' +
        '<span>' + unit + '</span>' +
        '</div>' +
        '</div>' +
        '</div>' +

        '<div style="display:flex; flex-direction:column; gap:6px;">' +
        '<label style="display:flex; align-items:center; gap:6px; font-size:11.5px; color:var(--text); cursor:pointer;">' +
        '<input type="checkbox" id="chkHerdarEssw"' + (essw.herdar_condicoes ? ' checked' : '') + ' style="accent-color:var(--accent); cursor:pointer;">' +
        '<span>Aplicar condições gerais da regra padrão</span>' +
        '</label>' +
        '</div>' +

        '<div style="display:flex; flex-direction:column; gap:6px;">' +
        '<div style="display:flex; justify-content:space-between; align-items:center;">' +
        '<span style="font-size:11px; font-weight:600; color:var(--text-dim);">Condições Específicas para ESSW (' + essw.condicoes.length + '):</span>' +
        '<button type="button" class="btn btn-add-esp-cond" data-esp="essw" style="font-size:10.5px; padding:2px 6px;">＋ Condição</button>' +
        '</div>' +
        '<div id="listEspCondEssw" style="display:flex; flex-direction:column; gap:4px;">';

      essw.condicoes.forEach(function (c, i) {
        html += espCondRowHtml(c, i, 'essw');
      });

      html += '</div></div></div>';
    } else {
      html += '<div style="font-size:11px; color:var(--text-faint); font-style:italic;">Utiliza a regra e condições padrão de 1 a 8 módulos.</div>';
    }
    html += '</div>';

    html += '</div></div></div>';
    return html;
  }

  function aplicarEtapasSimples(v, etapas) {
    (etapas || []).forEach(function (step) {
      if (!step) return;
      var num = parseFloat(step.valor);
      if (isNaN(num)) num = 0;
      if (step.tipo === 'dividir') { if (num !== 0) v = v / num; }
      else if (step.tipo === 'multiplicar') v = v * num;
      else if (step.tipo === 'somar') v = v + num;
      else if (step.tipo === 'subtrair') v = v - num;
      else if (step.tipo === 'arredondar') {
        var modo = step.modo || step.arredondamento || 'cima';
        if (modo === 'cima') v = Math.ceil(v);
        else if (modo === 'baixo') v = Math.floor(v);
        else if (modo === 'padrao') v = Math.round(v);
      } else if (step.tipo === 'limitar_max') v = Math.min(v, num);
      else if (step.tipo === 'limitar_min') v = Math.max(v, num);
    });
    return v;
  }

  function valorBase(base, mod, campoObj, flagsAtivos, vH, simCtxOverride) {
    if (!base) return 0;

    var activeCtx = Object.assign({}, SIM_CTX, simCtxOverride || {}, { nmod: mod });
    var tipo = (activeCtx && activeCtx.tipoestrutura) || 'Móvel';
    var isSpecialEstrutura = (tipo === 'Container Solar' || tipo === 'ESSW (mecânica)' || tipo === 'ESSW (elétrica)' || tipo === 'ESSW');

    if (base.perfis && Array.isArray(base.perfis) && base.perfis.length) {
      var perfilHit = matchedPerfil(base.perfis, activeCtx);
      if (perfilHit) {
        if (perfilHit.it && Array.isArray(perfilHit.it) && perfilHit.it.length) {
          var vPerfil = evalChain(perfilHit.it, activeCtx, base.blocos);
          return isFinite(vPerfil) ? vPerfil : 0;
        }
        if (perfilHit.etapas) {
          var vBaseDur = (vH !== undefined) ? vH : 0;
          return aplicarEtapasSimples(vBaseDur, perfilHit.etapas);
        }
        if (perfilHit.valor !== undefined) {
          if (typeof perfilHit.valor === 'number') return perfilHit.valor;
          return evalExpr(perfilHit.valor, vH, mod);
        }
      }
    }

    var forma = base.forma;

    // Se for estrutura especial (Solar ou ESSW) e vH estiver definido e positivo (calculando DUR)
    if (isSpecialEstrutura && vH !== undefined && vH !== null && vH > 0) {
      var etapas = base.etapas;
      if (!etapas || !etapas.length) {
        etapas = [];
        if (base.divisao) etapas.push({ tipo: 'dividir', valor: base.divisao });
        if (base.arredondamento) etapas.push({ tipo: 'arredondar', modo: base.arredondamento });
        if (base.subtracao) etapas.push({ tipo: 'subtrair', valor: base.subtracao });
        if (base.soma) etapas.push({ tipo: 'somar', valor: base.soma });
      }
      if (etapas.length > 0) {
        return aplicarEtapasSimples(vH, etapas);
      }
      if (forma === 'derivado_h' || forma === 'tabela' || forma === 'constante' || forma === 'multiplicativa' || forma === 'degrau_fixo' || forma === 'aditiva') {
        return aplicarEtapasSimples(vH, [{ tipo: 'dividir', valor: 7.92 }, { tipo: 'arredondar', modo: 'cima' }]);
      }
    }

    if (forma === 'blocos') {
      var hitM = matchedMontagem(base.montagens, activeCtx);
      var bVal = hitM ? evalChain(hitM.it, activeCtx, base.blocos) : 0;
      return isFinite(bVal) ? bVal : 0;
    }
    if (forma === 'constante') return parseFloat(base.valor !== undefined ? base.valor : base.valor_base) || 0;
    if (forma === 'multiplicativa') {
      var vBase = parseFloat(base.valor_base !== undefined ? base.valor_base : base.valor) || 0;
      var esc = base.escala !== undefined ? parseFloat(base.escala) : 0.5;
      if (isNaN(esc)) esc = 0.5;
      var res = vBase * (1 + (mod - 1) * esc);
      if (mod > 1) {
        if (base.adicao_final) res += parseFloat(base.adicao_final) || 0;
        if (base.subtracao_final) res -= parseFloat(base.subtracao_final) || 0;
      }
      return res;
    }
    if (forma === 'degrau_fixo') {
      var vBase = parseFloat(base.valor_base !== undefined ? base.valor_base : base.valor) || 0;
      if (mod === 1) return vBase;
      var fat = base.fator !== undefined ? parseFloat(base.fator) : (base.escala !== undefined ? parseFloat(base.escala) : 1.5);
      if (isNaN(fat)) fat = 1.5;
      var res = vBase * fat;
      if (base.adicao_final) res += parseFloat(base.adicao_final) || 0;
      if (base.subtracao_final) res -= parseFloat(base.subtracao_final) || 0;
      return res;
    }
    if (forma === 'soma_campos') {
      var total = 0;
      var camposLista = base.campos || [];
      var currentAreaObj = state.regrasData[state.selectedAreaIdx];
      if (currentAreaObj && currentAreaObj.campos) {
        camposLista.forEach(function (cKey) {
          var outroCampoObj = currentAreaObj.campos[cKey];
          if (outroCampoObj && outroCampoObj.H) {
            var valH = calcValor(outroCampoObj.H, mod, flagsAtivos, outroCampoObj, undefined, simCtxOverride);
            if (valH && !isNaN(valH) && isFinite(valH)) {
              total += valH;
            }
          }
        });
      }
      var v = total;
      var etapas = base.etapas || [];
      return aplicarEtapasSimples(v, etapas);
    }
    if (forma === 'aditiva') return (parseFloat(base.valor_base) || 0) + (parseFloat(base.passo) || 0) * (mod - 1);
    if (forma === 'tabela') {
      if (!base.valores || base.valores[mod - 1] === undefined) return 0;
      var hVal = (vH !== undefined) ? vH : 0;
      return evalExpr(base.valores[mod - 1], hVal, mod);
    }
    if (forma === 'derivado_h') {
      var v = (vH !== undefined) ? vH : 0;
      if (vH === undefined && campoObj && campoObj.H && state && state.selectedSubTab !== 'H') {
        v = calcValor(campoObj.H, mod, flagsAtivos, campoObj, undefined, simCtxOverride);
      }
      var etapas = base.etapas;
      if (!etapas || !etapas.length) {
        etapas = [];
        if (base.divisao) etapas.push({ tipo: 'dividir', valor: base.divisao });
        if (base.arredondamento) etapas.push({ tipo: 'arredondar', modo: base.arredondamento });
        if (base.subtracao) etapas.push({ tipo: 'subtrair', valor: base.subtracao });
        if (base.soma) etapas.push({ tipo: 'somar', valor: base.soma });
      }
      return aplicarEtapasSimples(v, etapas);
    }
    return 0;
  }

  function evalExpr(val, vH, mod) {
    if (typeof val === 'number') return isNaN(val) ? 0 : val;
    if (!val || typeof val !== 'string') return 0;
    var str = val.trim();
    if (!str) return 0;

    var hVal = (vH !== undefined && !isNaN(vH)) ? vH : 0;
    var mVal = (mod !== undefined && !isNaN(mod)) ? mod : 1;
    var simCtxM = Object.assign({}, SIM_CTX, { nmod: mVal });
    var compAcumVal = getVarVal('comp_acum', simCtxM);

    var expr = str
      .replace(/^=\s*/, '')
      .replace(/\bH\b/gi, hVal)
      .replace(/\bcomp_acum\b/gi, compAcumVal)
      .replace(/\bJN\b/gi, compAcumVal);

    // Substitui vírgulas decimais entre números (ex: 7,04 -> 7.04)
    expr = expr.replace(/([0-9]+),([0-9]+)/g, function (m, a, b) { return a + '.' + b; });
    // Converte separador de parâmetros ';' em ','
    expr = expr.replace(/;/g, ',');

    expr = expr.replace(/arredondar\.para\.cima\s*\(\s*([^,)]+)(?:\s*,\s*[0-9]+)?\s*\)/gi, 'Math.ceil($1)');
    expr = expr.replace(/ceil\s*\(\s*([^)]+)\s*\)/gi, 'Math.ceil($1)');
    expr = expr.replace(/arredondar\.para\.baixo\s*\(\s*([^,)]+)(?:\s*,\s*[0-9]+)?\s*\)/gi, 'Math.floor($1)');
    expr = expr.replace(/floor\s*\(\s*([^)]+)\s*\)/gi, 'Math.floor($1)');

    if (/\bH\b/i.test(str) && /\/\s*[0-9\.]+/i.test(expr) && !/Math\./i.test(expr) && !/ARRED|SE|IF/i.test(expr)) {
      expr = 'Math.ceil(' + expr + ')';
    }

    try {
      var fn = new Function('SE, IF, ARRED, ROUND, MIN, MAX, Math', '"use strict"; return (' + expr + ');');
      var res = fn(
        function (cond, tVal, fVal) { return cond ? tVal : (fVal !== undefined ? fVal : 0); },
        function (cond, tVal, fVal) { return cond ? tVal : (fVal !== undefined ? fVal : 0); },
        function (x, dec) { var f = Math.pow(10, dec || 0); return Math.round(x * f) / f; },
        function (x, dec) { var f = Math.pow(10, dec || 0); return Math.round(x * f) / f; },
        Math.min,
        Math.max,
        Math
      );
      return typeof res === 'number' && !isNaN(res) ? res : 0;
    } catch (e) {
      var num = parseFloat(expr.replace(/[^0-9\.-]/g, ''));
      return isNaN(num) ? 0 : num;
    }
  }

  function valorBonus(b, mod, vAtual, vH, simCtxOverride) {
    var activeCtx = Object.assign({}, SIM_CTX, simCtxOverride || {}, { nmod: mod });
    if (b.perfis && Array.isArray(b.perfis) && b.perfis.length) {
      var perfilHit = matchedPerfil(b.perfis, activeCtx);
      if (perfilHit) {
        if (perfilHit.it && Array.isArray(perfilHit.it) && perfilHit.it.length) {
          var vPerfil = evalChain(perfilHit.it, activeCtx);
          return isFinite(vPerfil) ? vPerfil : 0;
        }
        if (perfilHit.valor !== undefined) {
          if (typeof perfilHit.valor === 'number') return perfilHit.valor;
          return evalExpr(perfilHit.valor, vH, mod);
        }
      }
    }
    var val = evalExpr(b.valor, vH, mod);
    if (!vAtual) vAtual = 0;
    if (b.forma === 'tabela') {
      if (b.valores && b.valores[mod - 1] !== undefined) {
        return evalExpr(b.valores[mod - 1], vH, mod);
      }
      return val;
    }
    if (b.forma === 'fixo') return val;
    if (b.forma === 'por_modulo') return val * mod;
    if (b.forma === 'escala_multiplicativa' || b.forma === 'escalonado') {
      var esc = b.escala !== undefined ? parseFloat(b.escala) : 0.5;
      if (isNaN(esc)) esc = 0.5;
      return val * (1 + (mod - 1) * esc);
    }
    if (b.forma === 'subtrair' || b.forma === 'subtrair_fixo') return -val;
    if (b.forma === 'subtrair_por_modulo') return -val * mod;

    if (b.forma === 'multiplicar' || b.forma === 'multiplicar_fixo') return vAtual * (val - 1);
    if (b.forma === 'multiplicar_por_modulo') return vAtual * ((val * mod) - 1);

    if (b.forma === 'dividir' || b.forma === 'dividir_fixo') {
      var d = val || 1;
      return vAtual * ((1 / d) - 1);
    }
    if (b.forma === 'dividir_por_modulo') {
      var dm = (val || 1) * mod;
      return dm ? vAtual * ((1 / dm) - 1) : 0;
    }

    return 0;
  }

  function calcValor(analise, mod, flagsAtivos, campoObj, vH, simCtxOverride) {
    if (!analise) return 0;
    var activeCtx = Object.assign({}, SIM_CTX, simCtxOverride || {}, { nmod: mod });
    var tipo = (activeCtx && activeCtx.tipoestrutura) || 'Móvel';
    var isSolar = (tipo === 'Container Solar');
    var isEssw = (tipo === 'ESSW (mecânica)' || tipo === 'ESSW (elétrica)' || tipo === 'ESSW');

    // Se estamos calculando DUR e vH não foi explicitamente fornecido, busca no H do campo (sem recursão em H)
    if (vH === undefined && campoObj && campoObj.H && analise !== campoObj.H && ((campoObj.DUR && analise === campoObj.DUR) || (state && state.selectedSubTab === 'DUR'))) {
      vH = calcValor(campoObj.H, mod, flagsAtivos, campoObj, 0, activeCtx);
    }

    var esp = getEspeciais(analise);

    if (isSolar && esp.solar && esp.solar.ativo) {
      var vBaseSolar = valorBase(esp.solar.base, mod, campoObj, flagsAtivos, vH, activeCtx);
      var vTotalSolar = vBaseSolar;
      if (esp.solar.herdar_condicoes && analise.condicoes) {
        analise.condicoes.forEach(function (c) {
          if (flagsAtivos && flagsAtivos[c.flag]) vTotalSolar += valorBonus(c, mod, vTotalSolar, vH, activeCtx);
        });
      }
      if (esp.solar.condicoes) {
        esp.solar.condicoes.forEach(function (c) {
          if (flagsAtivos && flagsAtivos[c.flag]) vTotalSolar += valorBonus(c, mod, vTotalSolar, vH, activeCtx);
        });
      }
      return vTotalSolar;
    }

    if (isEssw && esp.essw && esp.essw.ativo) {
      var vBaseEssw = valorBase(esp.essw.base, mod, campoObj, flagsAtivos, vH, activeCtx);
      var vTotalEssw = vBaseEssw;
      if (esp.essw.herdar_condicoes && analise.condicoes) {
        analise.condicoes.forEach(function (c) {
          if (flagsAtivos && flagsAtivos[c.flag]) vTotalEssw += valorBonus(c, mod, vTotalEssw, vH, activeCtx);
        });
      }
      if (esp.essw.condicoes) {
        esp.essw.condicoes.forEach(function (c) {
          if (flagsAtivos && flagsAtivos[c.flag]) vTotalEssw += valorBonus(c, mod, vTotalEssw, vH, activeCtx);
        });
      }
      return vTotalEssw;
    }

    if (!analise.base) return 0;
    var v = valorBase(analise.base, mod, campoObj, flagsAtivos, vH, activeCtx);
    (analise.condicoes || []).forEach(function (c) {
      if (flagsAtivos && flagsAtivos[c.flag]) v += valorBonus(c, mod, v, vH, activeCtx);
    });
    return v;
  }

  /* ==========================================================================
     CONSTRUTOR DE BLOCOS CONDICIONAIS (MECÂNICA) ENGINE & UI
     ========================================================================== */
  var VARS_DEFAULT = [
    { grupo: 'Dimensões', nome: 'Comprimento', chave: 'comp', tipo: 'entrada', valor: 12 },
    { grupo: 'Dimensões', nome: 'Largura', chave: 'larg', tipo: 'entrada', valor: 2.4 },
    { grupo: 'Dimensões', nome: 'Altura', chave: 'alt', tipo: 'entrada', valor: 2.6 },
    { grupo: 'Dimensões por Módulo', nome: 'Comprimento Módulo 1', chave: 'comp_m1', tipo: 'entrada', valor: 12 },
    { grupo: 'Dimensões por Módulo', nome: 'Comprimento Módulo 2', chave: 'comp_m2', tipo: 'entrada', valor: 12 },
    { grupo: 'Dimensões por Módulo', nome: 'Comprimento Módulo 3', chave: 'comp_m3', tipo: 'entrada', valor: 12 },
    { grupo: 'Dimensões por Módulo', nome: 'Comprimento Módulo 4', chave: 'comp_m4', tipo: 'entrada', valor: 12 },
    { grupo: 'Dimensões por Módulo', nome: 'Comprimento Módulo 5', chave: 'comp_m5', tipo: 'entrada', valor: 12 },
    { grupo: 'Dimensões por Módulo', nome: 'Comprimento Módulo 6', chave: 'comp_m6', tipo: 'entrada', valor: 12 },
    { grupo: 'Dimensões por Módulo', nome: 'Comprimento Módulo 7', chave: 'comp_m7', tipo: 'entrada', valor: 12 },
    { grupo: 'Dimensões por Módulo', nome: 'Comprimento Módulo 8', chave: 'comp_m8', tipo: 'entrada', valor: 12 },
    { grupo: 'Dimensões por Módulo', nome: 'Largura Módulo 1', chave: 'larg_m1', tipo: 'entrada', valor: 2.4 },
    { grupo: 'Dimensões por Módulo', nome: 'Largura Módulo 2', chave: 'larg_m2', tipo: 'entrada', valor: 2.4 },
    { grupo: 'Dimensões por Módulo', nome: 'Largura Módulo 3', chave: 'larg_m3', tipo: 'entrada', valor: 2.4 },
    { grupo: 'Dimensões por Módulo', nome: 'Largura Módulo 4', chave: 'larg_m4', tipo: 'entrada', valor: 2.4 },
    { grupo: 'Dimensões por Módulo', nome: 'Largura Módulo 5', chave: 'larg_m5', tipo: 'entrada', valor: 2.4 },
    { grupo: 'Dimensões por Módulo', nome: 'Largura Módulo 6', chave: 'larg_m6', tipo: 'entrada', valor: 2.4 },
    { grupo: 'Dimensões por Módulo', nome: 'Largura Módulo 7', chave: 'larg_m7', tipo: 'entrada', valor: 2.4 },
    { grupo: 'Dimensões por Módulo', nome: 'Largura Módulo 8', chave: 'larg_m8', tipo: 'entrada', valor: 2.4 },
    { grupo: 'Sistema', nome: 'Nº de módulos', chave: 'nmod', tipo: 'entrada', valor: 1 },
    { grupo: 'Tempos de Corte', nome: 'Lateral-Fator1', chave: 'lat_f1', tipo: 'constante', valor: 1.14 },
    { grupo: 'Tempos de Corte', nome: 'Lateral-MinB', chave: 'lat_mb', tipo: 'constante', valor: 2.815 },
    { grupo: 'Tempos de Corte', nome: 'FrenFund-Fator1', chave: 'ffd_f1', tipo: 'constante', valor: 1.31 },
    { grupo: 'Tempos de Corte', nome: 'FrenFund-MinB', chave: 'ffd_mb', tipo: 'constante', valor: 1.472 },
    { grupo: 'Tempos de Corte', nome: 'Teto-Fator1', chave: 'tet_f1', tipo: 'constante', valor: 1.29 },
    { grupo: 'Tempos de Corte', nome: 'Teto-Fator2', chave: 'tet_f2', tipo: 'constante', valor: 1.19 },
    { grupo: 'Tempos de Corte', nome: 'Teto-MinB', chave: 'tet_mb', tipo: 'constante', valor: 4.785 },
    { grupo: 'Tempos de Corte', nome: 'Telhado-Fator1', chave: 'tlh_f1', tipo: 'constante', valor: 1.38 },
    { grupo: 'Tempos de Corte', nome: 'Telhado-Fator2', chave: 'tlh_f2', tipo: 'constante', valor: 1.025 },
    { grupo: 'Tempos de Corte', nome: 'Telhado-MinB', chave: 'tlh_mb', tipo: 'constante', valor: 1.36 },
    { grupo: 'Tempos de Corte', nome: 'Base-Fator1', chave: 'bas_f1', tipo: 'constante', valor: 1.4 },
    { grupo: 'Tempos de Corte', nome: 'Base-Fator2', chave: 'bas_f2', tipo: 'constante', valor: 1.405 },
    { grupo: 'Tempos de Corte', nome: 'Base-MinB', chave: 'bas_mb', tipo: 'constante', valor: 14.17 }
  ];

  function getCamposCondBloco() {
    var listas = (CONFIG && CONFIG.listas) ? CONFIG.listas : {};
    var optsEstrutura = listas.tipoestrutura || ['Móvel', 'Semimóvel', 'Modular', 'Fixo', 'Embarcado', 'Container Solar', 'Skid (mecânica)', 'Skid (com elétrica)', 'Pilotis', 'ESSW (mecânica)', 'ESSW (elétrica)', 'Serviço Engenharia'];
    var optsPlanpin = listas.planpin || ['WAU-ELETRO-08', 'WAU-ELETRO-09', 'WAU-ELETRO-04', 'Não aplicável'];
    var optsTipomaq = listas.tipomaq || ['Split', 'Wall Mounted', 'Roof Top', 'Não possui', 'Não aplicável'];
    var optsIncendio = listas.incendio || ['Com combate', 'Com instalações', 'Somente infra', 'Não aplicável'];
    var optsSeguranca = listas.seguranca || ['CFTV', 'Controle Acesso', 'CFTV + Controle Acesso', 'Não possui', 'Não aplicável'];
    var optsComplexidade = listas.complexidade || ['Simples', 'Médio', 'Complexo', 'Não aplicável'];
    var simNao = ['Sim', 'Não'];

    return [
      { k: 'tipoestrutura', n: 'Tipo de estrutura', opts: optsEstrutura },
      { k: 'planpin', n: 'Plano de pintura', opts: optsPlanpin },
      { k: 'complexidade', n: 'Complexidade', opts: optsComplexidade },
      { k: 'nmod', n: 'Nº de módulos', num: true },
      { k: 'comp', n: 'Comprimento (m)', num: true },
      { k: 'larg', n: 'Largura (m)', num: true },
      { k: 'alt', n: 'Altura (m)', num: true },
      { k: 'chapaRemovivel', n: 'Chapa Removível', opts: simNao },
      { k: 'peDireito', n: 'Pé Direito 3,3m', opts: simNao },
      { k: 'trafo_oleo', n: 'Trafo a óleo', opts: simNao },
      { k: 'tipomaq', n: 'Tipo de máquina', opts: optsTipomaq },
      { k: 'qtdmaq', n: 'Qtd. Ar Cond.', num: true },
      { k: 'incendio', n: 'Sistema de incêndio', opts: optsIncendio },
      { k: 'seguranca', n: 'Sistema de segurança', opts: optsSeguranca },
      { k: 'nrcolunas', n: 'Nº de Colunas', num: true },
      { k: 'casa_maquinas', n: 'Casa de Máquinas', opts: simNao },
      { k: 'white_martins', n: 'Cliente White Martins', opts: simNao },
      { k: 'paineis_interlig', n: 'Painéis Interligação', num: true },
      { k: 'testesw', n: 'Teste de Software', opts: simNao },
      { k: 'progReles', n: 'Programação de Relés', opts: simNao },
      { k: 'diagBTI', n: 'Diagrama BTI', opts: simNao },
      { k: 'diagAgrup', n: 'Diagrama Agrupador', opts: simNao },
      { k: 'fab1313', n: 'Fabricação 1313', opts: simNao },
      { k: 'itemFilho', n: 'Item Filho', opts: simNao },
      { k: 'semEngenharia', n: 'Sem Engenharia', opts: simNao },
      { k: 'acess_escada_weg', n: 'Acessório: Escada WEG', opts: simNao },
      { k: 'acess_escada_esp', n: 'Acessório: Escada Especial', opts: simNao },
      { k: 'acess_porao', n: 'Acessório: Porão de Cabos', opts: simNao },
      { k: 'acess_pilotis', n: 'Acessório: Pilotis', opts: simNao },
      { k: 'acess_dutos', n: 'Acessório: Rede de Dutos', opts: simNao },
      { k: 'acess_fundo_falso', n: 'Acessório: Fundo Falso', opts: simNao },
      { k: 'acess_dutos_bww', n: 'Acessório: Dutos BWW', opts: simNao },
      { k: 'acess_calhas', n: 'Acessório: Calhas Pluviais', opts: simNao },
      { k: 'acess_dutos_gases', n: 'Acessório: Dutos de Gases', opts: simNao }
    ];
  }

  var CAMPOS_COND_BLOCO = getCamposCondBloco();

  var SIM_CTX = {
    comp: 15, larg: 3, alt: 2.6, nmod: 1, tipoestrutura: 'Móvel',
    planpin: 'WAU-ELETRO-08', chapaRemovivel: 'Não', peDireito: 'Não',
    trafo_oleo: 'Não', complexidade: 'Simples', tipomaq: 'Split',
    incendio: 'Não aplicável', seguranca: 'CFTV + Controle Acesso',
    casa_maquinas: 'Sim', white_martins: 'Não', qtdmaq: 2, nrcolunas: 10,
    paineis_interlig: 2, testesw: 'Não', progReles: 'Não', diagBTI: 'Não',
    diagAgrup: 'Não', fab1313: 'Não', itemFilho: 'Não', semEngenharia: 'Não',
    acess_escada_weg: 'Não', acess_escada_esp: 'Não', acess_porao: 'Não',
    acess_pilotis: 'Não', acess_dutos: 'Não', acess_fundo_falso: 'Não',
    acess_dutos_bww: 'Não', acess_calhas: 'Não', acess_dutos_gases: 'Não',
    dur_mcm: 1, dur_tes: 1, dur_ins: 1
  };

  var OPS = ['+', '−', '×', '÷', '(', ')', '%', '^'];
  var OPMAP = { '+': '+', '−': '-', '×': '*', '÷': '/', '(': '(', ')': ')', '%': '%', '^': '**' };
  var OPINV = { '+': '+', '-': '−', '*': '×', '/': '÷', '(': '(', ')': ')', '%': '%', '**': '^' };
  var CONDOPS = ['=', '≠', '>', '<', '≥', '≤'];

  function getVARS() {
    return (CONFIG && CONFIG.variaveis && Array.isArray(CONFIG.variaveis)) ? CONFIG.variaveis : VARS_DEFAULT;
  }

  function getVarName(k) {
    var list = getVARS();
    var f = list.filter(function (x) { return x.chave === k || x.k === k; })[0];
    return f ? (f.nome || f.n) : k;
  }

  function getVarVal(k, simCtx) {
    if (k === 'comp_acum') {
      var nmod = parseInt((simCtx && simCtx.nmod) || '1', 10) || 1;
      var sumC = 0;
      for (var mi = 1; mi <= nmod; mi++) {
        var cM = (simCtx && simCtx['comp_m' + mi]) || (simCtx && simCtx.comp) || 12;
        sumC += parseFloat(String(cM).replace(',', '.')) || 0;
      }
      return sumC;
    }
    if (simCtx && simCtx[k] !== undefined && simCtx[k] !== '') return simCtx[k];
    var list = getVARS();
    var f = list.filter(function (x) { return x.chave === k || x.k === k; })[0];
    if (f) {
      if (f.tipo === 'entrada' || f.t === 'entrada') {
        if (simCtx && simCtx[k] !== undefined && simCtx[k] !== '') return simCtx[k];
        if (k.indexOf('comp_m') === 0 && simCtx && simCtx.comp !== undefined) return simCtx.comp;
        if (k.indexOf('larg_m') === 0 && simCtx && simCtx.larg !== undefined) return simCtx.larg;
        return (f.valor !== undefined ? f.valor : f.v);
      }
      return f.valor !== undefined ? f.valor : (f.v !== undefined ? f.v : 0);
    }
    if (k.indexOf('comp_m') === 0 && simCtx && simCtx.comp !== undefined) return simCtx.comp;
    if (k.indexOf('larg_m') === 0 && simCtx && simCtx.larg !== undefined) return simCtx.larg;
    return 0;
  }

  function chainExpr(items, num, simCtx, blocosList) {
    if (!items || !Array.isArray(items)) return '';
    return items.map(function (it) {
      if (it.t === 'op') return OPMAP[it.v] || it.v;
      if (it.t === 'num') return it.v;
      if (it.t === 'var') return num ? getVarVal(it.v, simCtx) : ('[' + getVarName(it.v) + ']');
      if (it.t === 'blk') {
        var bo = (blocosList || []).filter(function (x) { return x.id === it.v; })[0];
        return bo ? (num ? '(' + chainExpr(bo.it, true, simCtx, blocosList) + ')' : '«' + bo.nome + '»') : '0';
      }
      return '';
    }).join(' ');
  }

  function evalChain(items, simCtx, blocosList) {
    try {
      var expr = chainExpr(items, true, simCtx, blocosList);
      if (!expr || !expr.trim()) return NaN;
      var r = Function('"use strict";return (' + expr + ')')();
      return isFinite(r) ? r : NaN;
    } catch (e) {
      return NaN;
    }
  }

  var ev = evalChain;

  function condOkBloco(m, simCtx) {
    if (m.padrao || !m.cond || !m.cond.length) return true;
    var res = null;
    m.cond.forEach(function (c, i) {
      var a = simCtx[c.c];
      var bv = c.val;
      var fn = CAMPOS_COND_BLOCO.filter(function (x) { return x.k === c.c; })[0];
      if (fn && fn.num) {
        a = parseFloat(a) || 0;
        bv = parseFloat(String(bv).replace(',', '.')) || 0;
      }
      var ok = c.o === '=' ? a === bv : c.o === '≠' ? a !== bv : c.o === '>' ? a > bv : c.o === '<' ? a < bv : c.o === '≥' ? a >= bv : a <= bv;
      res = i === 0 ? ok : (c.j === 'OU' ? (res || ok) : (res && ok));
    });
    return res;
  }

  function matchedMontagem(montagens, simCtx) {
    if (!montagens || !Array.isArray(montagens)) return null;
    for (var i = 0; i < montagens.length; i++) {
      if (condOkBloco(montagens[i], simCtx)) return montagens[i];
    }
    return null;
  }

  function varSelectHtml(sel) {
    var list = getVARS();
    var gs = [];
    list.forEach(function (x) { var g = x.grupo || x.g || 'Geral'; if (gs.indexOf(g) < 0) gs.push(g); });
    var h = '<select data-role="var">';
    gs.forEach(function (g) {
      h += '<optgroup label="' + g + '">';
      list.filter(function (x) { return (x.grupo || x.g || 'Geral') === g; }).forEach(function (x) {
        var k = x.chave || x.k;
        var n = x.nome || x.n;
        h += '<option value="' + k + '"' + (k === sel ? ' selected' : '') + '>' + n + '</option>';
      });
      h += '</optgroup>';
    });
    return h + '</select>';
  }

  function opSelectHtml(sel) {
    return '<select data-role="op">' + OPS.map(function (o) {
      return '<option' + (o === (OPINV[sel] || sel) ? ' selected' : '') + '>' + o + '</option>';
    }).join('') + '</select>';
  }

  function blkSelectHtml(sel, skip, blocosList) {
    return '<select data-role="blk">' + (blocosList || []).filter(function (x) { return x.id !== skip; }).map(function (x) {
      return '<option value="' + x.id + '"' + (x.id === sel ? ' selected' : '') + '>' + x.nome + '</option>';
    }).join('') + '</select>';
  }

  function chipHtml(it, i, sc, skip, blocosList) {
    var d = ' data-sc="' + sc + '" data-i="' + i + '"';
    var dragAttr = ' draggable="true"';
    var handle = '<span class="chip-handle" title="Arrastar para reordenar">⋮⋮</span>';

    if (it.t === 'var') return '<span class="chip var"' + d + dragAttr + '>' + handle + varSelectHtml(it.v) + '<button type="button" class="x" data-del="1"' + d + '>✕</button></span>';
    if (it.t === 'op') return '<span class="chip op"' + d + dragAttr + '>' + handle + opSelectHtml(it.v) + '<button type="button" class="x" data-del="1"' + d + '>✕</button></span>';
    if (it.t === 'num') return '<span class="chip num"' + d + dragAttr + '>' + handle + '<input value="' + String(it.v).replace('.', ',') + '"><button type="button" class="x" data-del="1"' + d + '>✕</button></span>';
    if (it.t === 'blk') return '<span class="chip blk"' + d + dragAttr + '>' + handle + blkSelectHtml(it.v, skip, blocosList) + '<button type="button" class="x" data-del="1"' + d + '>✕</button></span>';
    return '';
  }

  function addBtnsHtml(sc, blk) {
    return '<button type="button" class="addbtn" draggable="true" data-add="var" data-sc="' + sc + '" title="Clique para adicionar ou arraste para posicionar">+ variável</button>' +
      '<button type="button" class="addbtn" draggable="true" data-add="op" data-sc="' + sc + '" title="Clique para adicionar ou arraste para posicionar">+ operação</button>' +
      '<button type="button" class="addbtn n" draggable="true" data-add="num" data-sc="' + sc + '" title="Clique para adicionar ou arraste para posicionar">+ valor fixo</button>' +
      (blk ? '<button type="button" class="addbtn b" draggable="true" data-add="blk" data-sc="' + sc + '" title="Clique para adicionar ou arraste para posicionar">+ bloco</button>' : '');
  }

  var dragSession = null;

  function wireChipDragAndDrop() {
    if (window._chipDragWired) return;
    window._chipDragWired = true;

    document.addEventListener('dragstart', function (e) {
      var chip = e.target.closest('.chip[draggable="true"]');
      var addbtn = e.target.closest('.addbtn[draggable="true"]');

      if (chip && chip.dataset && chip.dataset.sc) {
        dragSession = {
          type: 'reorder',
          sc: chip.dataset.sc,
          i: parseInt(chip.dataset.i, 10)
        };
        e.dataTransfer.effectAllowed = 'move';
        try { e.dataTransfer.setData('text/plain', JSON.stringify(dragSession)); } catch (err) { }
        chip.classList.add('dragging');
        return;
      }

      if (addbtn && addbtn.dataset && addbtn.dataset.add) {
        dragSession = {
          type: 'add',
          addType: addbtn.dataset.add,
          sc: addbtn.dataset.sc
        };
        e.dataTransfer.effectAllowed = 'copyMove';
        try { e.dataTransfer.setData('text/plain', JSON.stringify(dragSession)); } catch (err) { }
        addbtn.classList.add('dragging');
        return;
      }
    });

    document.addEventListener('dragend', function () {
      dragSession = null;
      document.querySelectorAll('.chip, .addbtn').forEach(function (el) {
        el.classList.remove('dragging', 'drop-target-left', 'drop-target-right');
      });
      document.querySelectorAll('.chain').forEach(function (el) {
        el.classList.remove('drop-target-empty');
      });
    });

    document.addEventListener('dragover', function (e) {
      if (!dragSession) return;

      var chain = e.target.closest('.chain');
      if (!chain) return;

      e.preventDefault();
      e.dataTransfer.dropEffect = dragSession.type === 'reorder' ? 'move' : 'copy';

      document.querySelectorAll('.chip').forEach(function (el) {
        el.classList.remove('drop-target-left', 'drop-target-right');
      });
      document.querySelectorAll('.chain').forEach(function (el) {
        el.classList.remove('drop-target-empty');
      });

      var chip = e.target.closest('.chip');
      if (chip && chain.contains(chip)) {
        var rect = chip.getBoundingClientRect();
        var midX = rect.left + rect.width / 2;
        if (e.clientX < midX) {
          chip.classList.add('drop-target-left');
        } else {
          chip.classList.add('drop-target-right');
        }
      } else {
        chain.classList.add('drop-target-empty');
      }
    });

    document.addEventListener('drop', function (e) {
      if (!dragSession) return;
      var chain = e.target.closest('.chain');
      if (!chain) return;

      e.preventDefault();

      var base = getCurrentBase();
      if (!base) return;

      var scopeArr = function (sc) {
        var bo = (base.blocos || []).filter(function (x) { return x.id === sc; })[0];
        if (bo) return bo.it;
        var mo = (base.montagens || []).filter(function (x) { return x.id === sc; })[0];
        if (mo) return mo.it;
        return null;
      };

      var chip = e.target.closest('.chip');
      var targetSc = chain.dataset.sc || (chip ? chip.dataset.sc : null);
      var insertIdx = 0;

      if (chip && chip.dataset && chip.dataset.i !== undefined) {
        targetSc = chip.dataset.sc || targetSc;
        var chipIdx = parseInt(chip.dataset.i, 10);
        var rect = chip.getBoundingClientRect();
        var midX = rect.left + rect.width / 2;
        insertIdx = (e.clientX < midX) ? chipIdx : chipIdx + 1;
      } else {
        var chipsInChain = chain.querySelectorAll('.chip');
        if (chipsInChain.length > 0) {
          var lastChip = chipsInChain[chipsInChain.length - 1];
          targetSc = lastChip.dataset.sc || targetSc;
          insertIdx = parseInt(lastChip.dataset.i, 10) + 1;
        } else {
          insertIdx = 0;
        }
      }

      if (!targetSc) return;

      var targetArr = scopeArr(targetSc);
      if (!targetArr) return;

      if (dragSession.type === 'reorder') {
        var sourceArr = scopeArr(dragSession.sc);
        if (!sourceArr || dragSession.i < 0 || dragSession.i >= sourceArr.length) return;

        var movedItem = sourceArr.splice(dragSession.i, 1)[0];

        if (sourceArr === targetArr && dragSession.i < insertIdx) {
          insertIdx--;
        }
        targetArr.splice(insertIdx, 0, movedItem);

      } else if (dragSession.type === 'add') {
        var newItem = null;
        var vars = getVARS();
        var firstVarKey = (vars && vars[0]) ? (vars[0].chave || vars[0].k) : 'comp';
        var firstBlkId = (base.blocos && base.blocos[0]) ? base.blocos[0].id : 'b1';

        if (dragSession.addType === 'var') newItem = { t: 'var', v: firstVarKey };
        else if (dragSession.addType === 'op') newItem = { t: 'op', v: '+' };
        else if (dragSession.addType === 'num') newItem = { t: 'num', v: 1 };
        else if (dragSession.addType === 'blk') newItem = { t: 'blk', v: firstBlkId };

        if (newItem) {
          targetArr.splice(insertIdx, 0, newItem);
        }
      }

      dragSession = null;
      document.querySelectorAll('.chip, .addbtn').forEach(function (el) {
        el.classList.remove('dragging', 'drop-target-left', 'drop-target-right');
      });
      document.querySelectorAll('.chain').forEach(function (el) {
        el.classList.remove('drop-target-empty');
      });

      markDirty();
      renderBlocosModal(state.dirtySubTabRule);
      renderEditor();
    });
  }

  function getCurrentBase() {
    var r = state.dirtySubTabRule;
    return (r && r.base) ? r.base : null;
  }

  function wireBlocosEvents(subTabRule) {
    var base = getCurrentBase();
    if (!base || base.forma !== 'blocos') return;

    wireChipDragAndDrop();

    var scopeArr = function (sc) {
      var b = getCurrentBase();
      if (!b) return null;
      var bo = (b.blocos || []).filter(function (x) { return x.id === sc; })[0];
      if (bo) return bo.it;
      var m = (b.montagens || []).filter(function (x) { return x.id === sc; })[0];
      return m ? m.it : null;
    };

    var byM = function (id) {
      var b = getCurrentBase();
      if (!b) return null;
      return (b.montagens || []).filter(function (x) { return x.id === id; })[0];
    };

    // Modal overlay handlers
    function handleBlocoClick(t, curBase) {
      if (!t || !curBase) return false;

      // Toggle montagem collapse/expand
      var toggleBtn = t.closest('[data-toggle-mcard]') || (t.dataset && t.dataset.toggleMcard ? t : null);
      if (toggleBtn && toggleBtn.dataset && toggleBtn.dataset.toggleMcard) {
        var toggleId = toggleBtn.dataset.toggleMcard;
        var mToggle = (curBase.montagens || []).filter(function (x) { return x.id === toggleId; })[0];
        if (mToggle) {
          mToggle.collapsed = !mToggle.collapsed;
          renderEditor();
        }
        return true;
      }

      var mheadEl = t.closest('.mhead2');
      if (mheadEl && !t.closest('input, select, button')) {
        var headId = mheadEl.dataset.m;
        var mHead = (curBase.montagens || []).filter(function (x) { return x.id === headId; })[0];
        if (mHead) {
          mHead.collapsed = !mHead.collapsed;
          renderEditor();
        }
        return true;
      }

      // Toggle block collapse/expand inside modal
      var toggleBBtn = t.closest('[data-toggle-bcard]') || (t.dataset && t.dataset.toggleBcard ? t : null);
      if (toggleBBtn && toggleBBtn.dataset && toggleBBtn.dataset.toggleBcard) {
        var bIdToggle = toggleBBtn.dataset.toggleBcard;
        var bToggle = (curBase.blocos || []).filter(function (x) { return x.id === bIdToggle; })[0];
        if (bToggle) {
          bToggle.collapsed = !bToggle.collapsed;
          renderBlocosModal(state.dirtySubTabRule);
        }
        return true;
      }

      var bheadEl = t.closest('.bhead');
      if (bheadEl && !t.closest('input, select, button, .bdel')) {
        var headBId = bheadEl.dataset.toggleBcard || (bheadEl.querySelector('[data-bn]') ? bheadEl.querySelector('[data-bn]').dataset.bn : null);
        var bHead = (curBase.blocos || []).filter(function (x) { return x.id === headBId; })[0];
        if (bHead) {
          bHead.collapsed = !bHead.collapsed;
          renderBlocosModal(state.dirtySubTabRule);
        }
        return true;
      }

      // Open Modal "Ver / criar blocos"
      if (t.closest('#openBlk')) {
        renderBlocosModal(state.dirtySubTabRule);
        if ($('ovlBlk')) $('ovlBlk').classList.add('open');
        return true;
      }

      // Close Modal button or overlay backdrop
      if (t === $('ovlBlk') || t.closest('#btnCloseBlk')) {
        if ($('ovlBlk')) $('ovlBlk').classList.remove('open');
        return true;
      }

      // New block button
      if (t.closest('#newBlk')) {
        if (!curBase.blocos) curBase.blocos = [];
        curBase.blocos.push({ id: 'b' + (Date.now() % 900000), nome: 'Bloco ' + (curBase.blocos.length + 1), it: [{ t: 'var', v: 'comp' }] });
        renderBlocosModal(state.dirtySubTabRule);
        markDirty();
        renderEditor();
        return true;
      }

      // New montagem button
      if (t.closest('#newMont')) {
        if (!curBase.montagens) curBase.montagens = [];
        var firstBlkId = (curBase.blocos && curBase.blocos[0]) ? curBase.blocos[0].id : 'b1';
        var hasPadrao = curBase.montagens.some(function (x) { return x.padrao; });
        var insertIdx = hasPadrao ? curBase.montagens.length - 1 : curBase.montagens.length;
        curBase.montagens.splice(insertIdx, 0, {
          id: 'm' + (Date.now() % 900000),
          nome: 'Nova montagem',
          cond: [{ c: 'tipoestrutura', o: '=', val: CAMPOS_COND_BLOCO[0].opts[0], j: 'E' }],
          it: [{ t: 'blk', v: firstBlkId }]
        });
        markDirty();
        renderEditor();
        return true;
      }

      // Delete chip
      var delBtn = t.closest('[data-del]') || (t.dataset && t.dataset.del !== undefined ? t : null);
      if (delBtn && delBtn.dataset && delBtn.dataset.sc) {
        var arr = scopeArr(delBtn.dataset.sc);
        if (arr) {
          arr.splice(+delBtn.dataset.i, 1);
          renderBlocosModal(state.dirtySubTabRule);
          markDirty();
          renderEditor();
        }
        return true;
      }

      // Add item (+ variável, + operação, + valor fixo, + bloco)
      var addBtn = t.closest('[data-add]') || (t.dataset && t.dataset.add ? t : null);
      if (addBtn && addBtn.dataset && addBtn.dataset.sc) {
        var a = scopeArr(addBtn.dataset.sc);
        if (a) {
          var firstVar = getVARS()[0] ? (getVARS()[0].chave || getVARS()[0].k) : 'comp';
          var firstBlk = (curBase.blocos && curBase.blocos[0]) ? curBase.blocos[0].id : 'b1';
          a.push(addBtn.dataset.add === 'var' ? { t: 'var', v: firstVar }
            : addBtn.dataset.add === 'op' ? { t: 'op', v: '*' }
              : addBtn.dataset.add === 'num' ? { t: 'num', v: 1 }
                : { t: 'blk', v: firstBlk });
          renderBlocosModal(state.dirtySubTabRule);
          markDirty();
          renderEditor();
        }
        return true;
      }

      // Move montagem
      var mvBtn = t.closest('[data-mv]') || (t.dataset && t.dataset.mv ? t : null);
      if (mvBtn && mvBtn.dataset && mvBtn.dataset.m) {
        var idx = (curBase.montagens || []).findIndex(function (x) { return x.id === mvBtn.dataset.m; });
        var targetIdx = mvBtn.dataset.mv === 'up' ? idx - 1 : idx + 1;
        if (idx >= 0 && targetIdx >= 0 && targetIdx < curBase.montagens.length) {
          var tmp = curBase.montagens[idx];
          curBase.montagens[idx] = curBase.montagens[targetIdx];
          curBase.montagens[targetIdx] = tmp;
          markDirty();
          renderEditor();
        }
        return true;
      }

      // Delete montagem
      var mdelBtn = t.closest('[data-mdel]') || (t.dataset && t.dataset.mdel ? t : null);
      if (mdelBtn && mdelBtn.dataset) {
        curBase.montagens = curBase.montagens.filter(function (x) { return x.id !== mdelBtn.dataset.mdel; });
        markDirty();
        renderEditor();
        return true;
      }

      // Delete block
      var bdelBtn = t.closest('[data-bdel]') || (t.dataset && t.dataset.bdel ? t : null);
      if (bdelBtn && bdelBtn.dataset) {
        var bId = bdelBtn.dataset.bdel;
        curBase.blocos = (curBase.blocos || []).filter(function (x) { return x.id !== bId; });
        (curBase.montagens || []).forEach(function (m) {
          m.it = (m.it || []).filter(function (it) { return !(it.t === 'blk' && it.v === bId); });
        });
        renderBlocosModal(state.dirtySubTabRule);
        markDirty();
        renderEditor();
        return true;
      }

      // Add condition
      var addCondBtn = t.closest('[data-addcond]') || (t.dataset && t.dataset.addcond ? t : null);
      if (addCondBtn && addCondBtn.dataset) {
        var mObj = byM(addCondBtn.dataset.addcond);
        if (mObj) {
          if (!mObj.cond) mObj.cond = [];
          mObj.cond.push({ c: 'tipoestrutura', o: '=', val: CAMPOS_COND_BLOCO[0].opts[0], j: 'E' });
          markDirty();
          renderEditor();
        }
        return true;
      }

      // Delete condition
      var cdelBtn = t.closest('[data-cdel]') || (t.dataset && t.dataset.cdel ? t : null);
      if (cdelBtn && cdelBtn.dataset && cdelBtn.dataset.m) {
        var mObj2 = byM(cdelBtn.dataset.m);
        if (mObj2 && mObj2.cond) {
          mObj2.cond.splice(+cdelBtn.dataset.ci, 1);
          markDirty();
          renderEditor();
        }
        return true;
      }

      return false;
    }

    function handleBlocoChange(t, curBase) {
      if (!t || !curBase) return false;

      if (t.dataset && t.dataset.role) {
        var ch = t.closest('.chip');
        if (ch && ch.dataset) {
          var a = scopeArr(ch.dataset.sc);
          if (a && a[+ch.dataset.i]) {
            a[+ch.dataset.i].v = t.value;
            renderBlocosModal(state.dirtySubTabRule);
            markDirty();
          }
        }
        return true;
      }

      if (t.dataset && t.dataset.mtype) {
        var mObj = byM(t.dataset.mtype);
        if (mObj) {
          if (t.value === 'padrao') {
            mObj.padrao = true;
          } else {
            delete mObj.padrao;
            if (!mObj.cond || !mObj.cond.length) {
              mObj.cond = [{ c: 'tipoestrutura', o: '=', val: CAMPOS_COND_BLOCO[0].opts[0], j: 'E' }];
            }
          }
          markDirty();
          renderEditor();
        }
        return true;
      }

      if (t.dataset && t.dataset.cf) {
        var row = t.closest('.condrow');
        if (row && row.dataset) {
          var m = byM(row.dataset.m);
          if (m && m.cond && m.cond[+row.dataset.ci]) {
            var cd = m.cond[+row.dataset.ci];
            if (t.dataset.cf === 'campo') {
              cd.c = t.value;
              var fObj = CAMPOS_COND_BLOCO.filter(function (x) { return x.k === t.value; })[0];
              cd.val = fObj && fObj.opts ? fObj.opts[0] : 1;
            } else if (t.dataset.cf === 'op') {
              cd.o = t.value;
            } else if (t.dataset.cf === 'join') {
              cd.j = t.value;
            } else {
              cd.val = t.value;
            }
            markDirty();
            renderEditor();
          }
        }
        return true;
      }

      if (t.dataset && t.dataset.sim) {
        var simK = t.dataset.sim;
        var parsedNum = parseFloat(t.value.replace(',', '.'));
        var isOptField = CAMPOS_COND_BLOCO.filter(function (x) { return x.k === simK && x.opts; })[0];
        SIM_CTX[simK] = (isNaN(parsedNum) || isOptField) ? t.value : parsedNum;
        renderBlocosModal(state.dirtySubTabRule);
        renderEditor();
        return true;
      }

      return false;
    }

    function handleBlocoInput(t, curBase) {
      if (!t || !curBase) return false;

      if (t.dataset && t.dataset.mn) {
        var m = byM(t.dataset.mn);
        if (m) {
          m.nome = t.value;
          markDirty();
        }
        return true;
      }

      if (t.dataset && t.dataset.bn) {
        var bo = (curBase.blocos || []).filter(function (x) { return x.id === t.dataset.bn; })[0];
        if (bo) {
          bo.nome = t.value;
          markDirty();
        }
        return true;
      }

      var chipNum = t.closest('.chip.num');
      if (chipNum && chipNum.dataset) {
        var a = scopeArr(chipNum.dataset.sc);
        if (a && a[+chipNum.dataset.i]) {
          a[+chipNum.dataset.i].v = parseFloat(t.value.replace(',', '.')) || 0;
          renderBlocosModal(state.dirtySubTabRule);
          markDirty();
        }
        return true;
      }

      return false;
    }

    // Modal overlay handlers
    var ovl = $('ovlBlk');
    if (ovl && !ovl._wired) {
      ovl._wired = true;
      ovl.addEventListener('click', function (e) {
        handleBlocoClick(e.target, getCurrentBase());
      });
      ovl.addEventListener('change', function (e) {
        handleBlocoChange(e.target, getCurrentBase());
      });
      ovl.addEventListener('input', function (e) {
        handleBlocoInput(e.target, getCurrentBase());
      });
    }

    var colEd = $('colEditor');
    if (colEd && !colEd._blocosWired) {
      colEd._blocosWired = true;
      colEd.addEventListener('click', function (e) {
        handleBlocoClick(e.target, getCurrentBase());
      });
      colEd.addEventListener('change', function (e) {
        handleBlocoChange(e.target, getCurrentBase());
      });
      colEd.addEventListener('input', function (e) {
        handleBlocoInput(e.target, getCurrentBase());
      });
    }
  }

  function renderBlocosModal(subTabRule) {
    var base = (subTabRule && subTabRule.base) ? subTabRule.base : getCurrentBase();
    var container = $('blocos');
    if (!container || !base || !base.blocos) return;

    container.innerHTML = base.blocos.map(function (bo, bi) {
      var uso = (base.montagens || []).filter(function (m) {
        return (m.it || []).some(function (it) { return it.t === 'blk' && it.v === bo.id; });
      }).map(function (m) { return m.nome; });

      var bVal = ev(bo.it, SIM_CTX, base.blocos);
      if (bo.collapsed === undefined) {
        bo.collapsed = (bi !== 0);
      }
      var isCollapsed = bo.collapsed;
      var bExpr = chainExpr(bo.it, false, SIM_CTX, base.blocos);

      return '<div class="bcard' + (isCollapsed ? ' collapsed' : '') + '">' +
        '<div class="bhead" data-toggle-bcard="' + bo.id + '">' +
        '<button type="button" class="btn-toggle-bcard" data-toggle-bcard="' + bo.id + '" title="' + (isCollapsed ? 'Expandir bloco' : 'Recolher bloco') + '">' +
        '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4"><path d="m6 9 6 6 6-6"/></svg>' +
        '</button>' +
        '<span class="tag">B' + (bi + 1) + '</span>' +
        '<input class="bname" data-bn="' + bo.id + '" value="' + escapeHtml(bo.nome) + '">' +
        '<span class="bhead-expr-summary">' + escapeHtml(bExpr) + '</span>' +
        '<div style="margin-left:auto; display:flex; align-items:center; gap:8px;">' +
        '<span class="bres">= ' + (isFinite(bVal) ? bVal.toFixed(2).replace('.', ',') : '—') + '</span>' +
        '<button type="button" class="bdel" data-bdel="' + bo.id + '">✕</button></div></div>' +
        '<div class="bcard-body">' +
        '<div class="chain" data-sc="' + bo.id + '">' + (bo.it || []).map(function (it, i) { return chipHtml(it, i, bo.id, bo.id, base.blocos); }).join('') + addBtnsHtml(bo.id, true) + '</div>' +
        '<div class="usedby">' + (uso.length ? 'usado em: ' + uso.join(', ') : 'não usado em nenhuma montagem') + '</div></div></div>';
    }).join('');

    if (!container._wired) {
      container._wired = true;
      container.addEventListener('click', function (e) {
        var t = e.target;
        if (!t) return;
        if (t.dataset && t.dataset.bdel) {
          var bId = t.dataset.bdel;
          base.blocos = base.blocos.filter(function (x) { return x.id !== bId; });
          (base.montagens || []).forEach(function (m) {
            m.it = (m.it || []).filter(function (it) { return !(it.t === 'blk' && it.v === bId); });
          });
          renderBlocosModal(subTabRule);
          markDirty();
          renderEditor();
        }
      });
      container.addEventListener('input', function (e) {
        var t = e.target;
        if (!t) return;
        if (t.dataset && t.dataset.bn) {
          var bo = (base.blocos || []).filter(function (x) { return x.id === t.dataset.bn; })[0];
          if (bo) {
            bo.nome = t.value;
            markDirty();
            renderEditor();
          }
        }
      });
    }
  }

  function stepRowHtml(step, idx, totalSteps) {
    var isArred = step.tipo === 'arredondar';
    return '<div class="step-row" data-step-idx="' + idx + '" style="display:flex; align-items:center; gap:8px; background:var(--panel-2); border:1px solid var(--border); border-radius:9px; padding:8px 10px; margin-top:6px;">' +
      '<span class="mono" style="font-size:11px; color:var(--text-faint); min-width:50px;">Passo ' + (idx + 1) + '</span>' +
      '<select class="ipt step-tipo" style="padding:5px 8px; font-size:12px; width:150px; font-weight:500;">' +
      '<option value="dividir"' + (step.tipo === 'dividir' ? ' selected' : '') + '>Dividir (÷)</option>' +
      '<option value="multiplicar"' + (step.tipo === 'multiplicar' ? ' selected' : '') + '>Multiplicar (×)</option>' +
      '<option value="somar"' + (step.tipo === 'somar' ? ' selected' : '') + '>Somar (+)</option>' +
      '<option value="subtrair"' + (step.tipo === 'subtrair' ? ' selected' : '') + '>Subtrair (-)</option>' +
      '<option value="arredondar"' + (step.tipo === 'arredondar' ? ' selected' : '') + '>Arredondar</option>' +
      '<option value="limitar_max"' + (step.tipo === 'limitar_max' ? ' selected' : '') + '>Limitar Máximo / Teto (≤)</option>' +
      '<option value="limitar_min"' + (step.tipo === 'limitar_min' ? ' selected' : '') + '>Limitar Mínimo / Piso (≥)</option>' +
      '</select>' +
      (isArred ?
        '<select class="ipt step-modo" style="padding:5px 8px; font-size:12px; flex:1;">' +
        '<option value="cima"' + (step.modo === 'cima' || !step.modo ? ' selected' : '') + '>Arredondar p/ cima (Ceil)</option>' +
        '<option value="baixo"' + (step.modo === 'baixo' ? ' selected' : '') + '>Arredondar p/ baixo (Floor)</option>' +
        '<option value="padrao"' + (step.modo === 'padrao' ? ' selected' : '') + '>Mais próximo (Round)</option>' +
        '</select>'
        :
        '<input type="number" step="0.01" class="ipt mono step-valor" value="' + (step.valor !== undefined ? step.valor : 1) + '" style="padding:5px 8px; font-size:12.5px; flex:1; text-align:right;">'
      ) +
      '<button type="button" class="icon-btn step-up" title="Mover para cima" style="width:24px; height:24px; font-size:11px;" ' + (idx === 0 ? 'disabled' : '') + '>↑</button>' +
      '<button type="button" class="icon-btn step-down" title="Mover para baixo" style="width:24px; height:24px; font-size:11px;" ' + (idx === totalSteps - 1 ? 'disabled' : '') + '>↓</button>' +
      '<button type="button" class="icon-btn step-del" title="Remover etapa" style="width:24px; height:24px; color:var(--red); border-color:color-mix(in srgb, var(--red) 30%, transparent);">✕</button>' +
      '</div>';
  }

  function renderEditor() {
    var wrap = $('colEditor');
    if (!wrap) return;
    var areaObj = state.regrasData[state.selectedAreaIdx];
    if (!areaObj || !state.selectedCampoKey) {
      wrap.innerHTML = '<div class="pane"><div class="empty-state">' +
        '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M9 3v4M15 3v4M4 8h16M6 6h12a2 2 0 0 1 2 2v11a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2Z"/></svg>' +
        '<div>Selecione um campo à esquerda para editar</div></div></div>';
      if ($('maintFooterMeta')) $('maintFooterMeta').textContent = 'Selecione um campo para editar';
      return;
    }

    var campoObj = areaObj.campos[state.selectedCampoKey];
    var subTabRule = state.dirtySubTabRule;
    if (!subTabRule) return;

    var html = '<div class="pane">' +
      '<div class="editor-head"><div>' +
      '<h1>' + state.selectedCampoKey + '</h1>' +
      '<div class="path">' + areaObj.area + ' &gt; ' + state.selectedCampoKey + '</div>' +
      '</div></div>';

    // Sub-Tabs Header (H vs DUR)
    html += '<div class="modal-tabs" style="padding:0; margin-bottom:18px; border-bottom:1px solid var(--hairline);">' +
      '<button type="button" class="modal-tab btn-subtab ' + (state.selectedSubTab === 'H' ? 'active' : '') + '" data-subtab="H">Horas (H)</button>' +
      '<button type="button" class="modal-tab btn-subtab ' + (state.selectedSubTab === 'DUR' ? 'active' : '') + '" data-subtab="DUR">Duração (DUR)</button>' +
      '</div>';

    var base = subTabRule.base;
    html += '<div class="field-block">' +
      '<div class="field-label" style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:8px;">' +
      '<span>Cálculo Base (' + (state.selectedSubTab === 'H' ? 'Horas' : 'Duração') + ')</span>' +
      '<select id="baseFormaSelect" class="ipt" style="padding:4px 8px; font-size:12px; font-weight:500;">' +
      '<option value="constante"' + (base.forma === 'constante' ? ' selected' : '') + '>Constante Fixo (Todos módulos iguais)</option>' +
      '<option value="aditiva"' + (base.forma === 'aditiva' ? ' selected' : '') + '>Aditiva (Base + Passo × Mód)</option>' +
      '<option value="multiplicativa"' + (base.forma === 'multiplicativa' ? ' selected' : '') + '>Multiplicativa Incremental (+×/mód)</option>' +
      '<option value="degrau_fixo"' + (base.forma === 'degrau_fixo' ? ' selected' : '') + '>Fator Fixo 2m+ (1m Base / 2m+ Base × Fator)</option>' +
      '<option value="tabela"' + (base.forma === 'tabela' ? ' selected' : '') + '>Valores Personalizados por Módulo (Tabela 1m a 8m)</option>' +
      '<option value="derivado_h"' + (base.forma === 'derivado_h' ? ' selected' : '') + '>Derivado de Horas H (Fórmula Operacional)</option>' +
      '<option value="soma_campos"' + (base.forma === 'soma_campos' ? ' selected' : '') + '>Soma de Campos (Agregação de outros campos calculados)</option>' +
      '<option value="blocos"' + (base.forma === 'blocos' ? ' selected' : '') + '>Blocos Condicionais (SE / Variáveis / Fórmulas)</option>' +
      '</select>' +
      '</div>';

    if (base.forma === 'blocos') {
      CAMPOS_COND_BLOCO = getCamposCondBloco();
      if (!base.blocos) {
        base.blocos = [
          { id: 'b1', nome: 'Lateral', it: [{ t: 'var', v: 'lat_f1' }, { t: 'op', v: '*' }, { t: 'var', v: 'comp' }, { t: 'op', v: '+' }, { t: 'num', v: 4 }] },
          { id: 'b2', nome: 'Frente/Fundo', it: [{ t: 'var', v: 'ffd_f1' }, { t: 'op', v: '*' }, { t: 'var', v: 'comp' }, { t: 'op', v: '+' }, { t: 'num', v: 4 }] },
          { id: 'b3', nome: 'Teto', it: [{ t: 'var', v: 'tet_f1' }, { t: 'op', v: '*' }, { t: 'var', v: 'comp' }, { t: 'op', v: '+' }, { t: 'var', v: 'tet_f2' }, { t: 'op', v: '*' }, { t: 'var', v: 'larg' }] },
          { id: 'b4', nome: 'Telhado', it: [{ t: 'var', v: 'tlh_f1' }, { t: 'op', v: '*' }, { t: 'var', v: 'comp' }, { t: 'op', v: '+' }, { t: 'var', v: 'tlh_f2' }, { t: 'op', v: '*' }, { t: 'var', v: 'larg' }] },
          { id: 'b5', nome: 'Base', it: [{ t: 'var', v: 'bas_f1' }, { t: 'op', v: '*' }, { t: 'var', v: 'comp' }, { t: 'op', v: '+' }, { t: 'var', v: 'bas_f2' }, { t: 'op', v: '*' }, { t: 'var', v: 'larg' }] }
        ];
      }
      if (!base.montagens) {
        base.montagens = [
          { id: 'm1', nome: 'Móvel', cond: [{ c: 'tipoestrutura', o: '=', val: 'Móvel', j: 'E' }], it: [{ t: 'blk', v: 'b5' }, { t: 'op', v: '*' }, { t: 'var', v: 'bas_mb' }] },
          { id: 'm2', nome: 'Embarcado', cond: [{ c: 'tipoestrutura', o: '=', val: 'Embarcado', j: 'E' }], it: [{ t: 'blk', v: 'b1' }, { t: 'op', v: '*' }, { t: 'var', v: 'lat_mb' }, { t: 'op', v: '+' }, { t: 'num', v: 2 }, { t: 'op', 'v': '*' }, { t: 'blk', v: 'b2' }, { t: 'op', v: '*' }, { t: 'var', v: 'ffd_mb' }, { t: 'op', v: '+' }, { t: 'blk', v: 'b3' }, { t: 'op', v: '*' }, { t: 'var', v: 'tet_mb' }, { t: 'op', v: '+' }, { t: 'blk', v: 'b5' }, { t: 'op', v: '*' }, { t: 'var', v: 'bas_mb' }] },
          { id: 'm0', nome: 'Demais estruturas', padrao: true, cond: [], it: [{ t: 'blk', v: 'b1' }, { t: 'op', v: '*' }, { t: 'var', v: 'lat_mb' }, { t: 'op', v: '+' }, { t: 'num', v: 2 }, { t: 'op', v: '*' }, { t: 'blk', v: 'b2' }, { t: 'op', v: '*' }, { t: 'var', v: 'ffd_mb' }, { t: 'op', v: '+' }, { t: 'blk', v: 'b3' }, { t: 'op', v: '*' }, { t: 'var', v: 'tet_mb' }, { t: 'op', v: '+' }, { t: 'num', v: 2 }, { t: 'op', v: '*' }, { t: 'blk', v: 'b4' }, { t: 'op', v: '*' }, { t: 'var', v: 'tlh_mb' }, { t: 'op', v: '+' }, { t: 'blk', v: 'b5' }, { t: 'op', v: '*' }, { t: 'var', v: 'bas_mb' }] }
        ];
      }
      if (base.ajuste_final) {
        delete base.ajuste_final;
      }
      (base.montagens || []).forEach(function (m) {
        if (!m.it || !m.it.length) return;
        var hasPlus = m.it.some(function (x) { return x.t === 'op' && (x.v === '+' || x.v === '−'); });
        var divIdx = -1;
        for (var d = 0; d < m.it.length; d++) {
          if (m.it[d].t === 'op' && m.it[d].v === '/') { divIdx = d; break; }
        }
        if (hasPlus && divIdx > 0 && m.it[0].v !== '(') {
          m.it.unshift({ t: 'op', v: '(' });
          m.it.splice(divIdx + 1, 0, { t: 'op', v: ')' });
        }
      });

      // 1. Barra de Blocos
      html += '<div class="blocbar" style="margin-top:12px;">' +
        '<span class="ic"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.3"><rect x="3" y="3" width="8" height="8" rx="1.6"/><rect x="13" y="3" width="8" height="8" rx="1.6"/><rect x="3" y="13" width="8" height="8" rx="1.6"/><rect x="13" y="13" width="8" height="8" rx="1.6"/></svg></span>' +
        '<span class="tt"><b>' + base.blocos.length + ' blocos definidos</b><span>' + base.blocos.map(function (x) { return x.nome; }).join(' · ') + '</span></span>' +
        '<button type="button" class="btn-vio" id="openBlk">Ver / criar blocos</button>' +
        '</div>';

      // 2. Container de Montagens Condicionais
      html += '<div class="block" style="margin-top:13px">' +
        '<div class="bl">Montagens <span class="hint">avaliadas de cima para baixo — vale a primeira que atender a condição</span></div>' +
        '<div id="montagensListContainer">';

      var hitM = matchedMontagem(base.montagens, SIM_CTX);
      base.montagens.forEach(function (m, mi) {
        var isHit = hitM && hitM.id === m.id;
        var rVal = ev(m.it, SIM_CTX, base.blocos);
        if (m.collapsed === undefined) {
          m.collapsed = hitM ? !isHit : (mi !== 0);
        }
        var isCollapsed = m.collapsed;

        var condSummary = '';
        if (m.padrao) {
          condSummary = 'senão';
        } else if (m.cond && m.cond.length) {
          condSummary = m.cond.map(function (c, ci) {
            var fObj = CAMPOS_COND_BLOCO.filter(function (x) { return x.k === c.c; })[0];
            var cName = fObj ? fObj.n.toLowerCase() : c.c;
            var prefix = ci > 0 ? (c.j === 'OU' ? ' ou ' : ' e ') : 'se ';
            return prefix + cName + ' ' + c.o + ' ' + c.val;
          }).join('');
        }

        html += '<div class="mcard' + (isHit ? ' hit' : '') + (m.padrao ? ' pad' : '') + (isCollapsed ? ' collapsed' : '') + '">' +
          '<div class="mhead2" data-m="' + m.id + '">' +
          '<button type="button" class="btn-toggle-mcard" data-toggle-mcard="' + m.id + '" title="' + (isCollapsed ? 'Expandir montagem' : 'Recolher montagem') + '">' +
          '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4"><path d="m6 9 6 6 6-6"/></svg>' +
          '</button>' +
          '<span class="ord">' + (mi + 1) + '</span>' +
          '<input class="mname" data-mn="' + m.id + '" value="' + escapeHtml(m.nome) + '">' +
          '<select class="mtype-sel" data-mtype="' + m.id + '">' +
          '<option value="cond"' + (!m.padrao ? ' selected' : '') + '>Condicional (SE)</option>' +
          '<option value="padrao"' + (m.padrao ? ' selected' : '') + '>Padrão (Senão)</option>' +
          '</select>' +
          '<span class="mhead-cond-summary">' + escapeHtml(condSummary) + '</span>' +
          '<div style="margin-left:auto; display:flex; align-items:center; gap:8px;">' +
          (isHit ? '<span class="hitbadge">APLICADA</span>' : '') +
          '<span class="mres">= ' + (isFinite(rVal) ? rVal.toFixed(1).replace('.', ',') : '—') + ' h</span>' +
          (mi > 0 ? '<button type="button" class="arrbtn" data-mv="up" data-m="' + m.id + '">▲</button>' : '') +
          (mi < base.montagens.length - 1 ? '<button type="button" class="arrbtn" data-mv="dn" data-m="' + m.id + '">▼</button>' : '') +
          '<button type="button" class="bdel" data-mdel="' + m.id + '">✕</button>' + '</div></div>' +
          '<div class="mcard-body">';

        if (m.padrao) {
          html += '<div class="padtxt">Senão — usada quando nenhuma condição acima é atendida</div>';
        } else {
          (m.cond || []).forEach(function (c, ci) {
            html += '<div class="condrow" data-m="' + m.id + '" data-ci="' + ci + '">' +
              (ci === 0 ? '<span class="kw">SE</span>' : '<select data-cf="join" style="width:64px"><option' + (c.j === 'E' ? ' selected' : '') + '>E</option><option' + (c.j === 'OU' ? ' selected' : '') + '>OU</option></select>') +
              '<select data-cf="campo">' + CAMPOS_COND_BLOCO.map(function (fn) { return '<option value="' + fn.k + '"' + (fn.k === c.c ? ' selected' : '') + '>' + fn.n + '</option>'; }).join('') + '</select>' +
              '<select data-cf="op" style="width:56px">' + CONDOPS.map(function (o) { return '<option' + (o === c.o ? ' selected' : '') + '>' + o + '</option>'; }).join('') + '</select>';

            var fObj = CAMPOS_COND_BLOCO.filter(function (x) { return x.k === c.c; })[0];
            if (fObj && fObj.opts) {
              html += '<select data-cf="val">' + fObj.opts.map(function (o) { return '<option' + (o === c.val ? ' selected' : '') + '>' + o + '</option>'; }).join('') + '</select>';
            } else {
              html += '<input data-cf="val" value="' + escapeHtml(c.val) + '">';
            }

            html += '<button type="button" class="bdel" data-cdel="1" data-m="' + m.id + '" data-ci="' + ci + '">✕</button>' +
              (ci === (m.cond.length - 1) ? '<button type="button" class="addbtn c" data-addcond="' + m.id + '">+ condição</button>' : '') +
              '</div>';
          });
          if (!m.cond || !m.cond.length) {
            html += '<div class="condrow"><span class="kw">SE</span><button type="button" class="addbtn c" data-addcond="' + m.id + '">+ condição</button></div>';
          }
        }

        html += '<div class="chain" data-sc="' + m.id + '">' + (m.it || []).map(function (it, i) { return chipHtml(it, i, m.id, null, base.blocos); }).join('') + addBtnsHtml(m.id, true) + '</div>' +
          '<div class="expr-out">' + chainExpr(m.it, false, SIM_CTX, base.blocos) + '</div></div></div>';
      });

      html += '</div>' +
        '<button type="button" class="newblock" id="newMont"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4"><path d="M12 5v14M5 12h14"/></svg>Adicionar montagem condicional</button>' +
        '</div>';

      // 3. Painel de Simulação
      var areaObj = state.regrasData[state.selectedAreaIdx];
      var isAcessorios = (areaObj && areaObj.area === 'ACESSÓRIOS');

      html += '<div class="block" style="margin-bottom:0">' +
        '<div class="bl">Simular <span class="hint">' + (isAcessorios ? 'escolha o número de módulos (1 a 8) e complexidade para conferir o resultado' : 'escolha o cenário e confira qual montagem é aplicada') + '</span></div>' +
        '<div class="simbar">';

      var simFields = isAcessorios
        ? ['nmod', 'complexidade', 'comp', 'comp_m1', 'comp_m2', 'comp_m3', 'comp_m4', 'comp_m5', 'comp_m6', 'comp_m7', 'comp_m8']
        : ['tipoestrutura', 'trafo_oleo', 'comp', 'larg'];

      simFields.forEach(function (k) {
        var cObj = CAMPOS_COND_BLOCO.filter(function (x) { return x.k === k; })[0];
        if (cObj) {
          html += '<span class="siminp"><label>' + cObj.n + '</label>' + (cObj.opts
            ? '<select data-sim="' + cObj.k + '">' + cObj.opts.map(function (o) { return '<option' + (String(o) === String(SIM_CTX[cObj.k]) ? ' selected' : '') + '>' + o + '</option>'; }).join('') + '</select>'
            : '<input data-sim="' + cObj.k + '" value="' + (SIM_CTX[cObj.k] !== undefined ? SIM_CTX[cObj.k] : '') + '">') + '</span>';
        } else {
          var vObj = getVARS().filter(function (y) { return y.chave === k || y.k === k; })[0];
          if (vObj) {
            var valDisp = SIM_CTX[k] !== undefined ? SIM_CTX[k] : (vObj.valor !== undefined ? vObj.valor : (vObj.v !== undefined ? vObj.v : ''));
            html += '<span class="siminp"><label>' + (vObj.nome || vObj.n) + '</label><input data-sim="' + k + '" value="' + String(valDisp).replace('.', ',') + '"></span>';
          }
        }
      });

      html += '</div>';

      var curNmod = parseInt(SIM_CTX.nmod, 10) || 1;
      var totVal = hitM ? ev(hitM.it, SIM_CTX, base.blocos) : NaN;
      if (!hitM) {
        html += '<div class="err">Nenhuma montagem atende ao cenário e não há montagem padrão.</div>';
      } else if (isNaN(totVal)) {
        html += '<div class="err">Expressão inválida — verifique se falta uma operação entre os itens.</div>';
      } else {
        html += '<div class="simtot"><span class="lb">Montagem aplicada para ' + curNmod + ' módulo' + (curNmod > 1 ? 's' : '') + ': <b>' + escapeHtml(hitM.nome) + '</b> · resultado de <b>' + state.selectedCampoKey + ' · ' + state.selectedSubTab + '</b></span>' +
          '<span class="vv">' + totVal.toFixed(1).replace('.', ',') + '</span><span class="un">horas</span></div>';
      }

      if (isAcessorios) {
        html += '<div class="sim-table-wrap" style="margin-top:14px; background:var(--bg-card); border:1px solid var(--border-subtle); border-radius:8px; padding:12px 14px;">' +
          '<div style="font-size:12px; font-weight:700; color:var(--accent); margin-bottom:10px; display:flex; align-items:center; justify-content:space-between;">' +
          '<span>Simulação de Resultado (1 a 8 Módulos)</span>' +
          '<span style="font-weight:400; font-size:11px; color:var(--text-faint);">Complexidade: <b>' + escapeHtml(SIM_CTX.complexidade || 'Simples') + '</b></span>' +
          '</div>' +
          '<div style="display:grid; grid-template-columns: repeat(auto-fit, minmax(130px, 1fr)); gap:8px;">';

        for (var m = 1; m <= 8; m++) {
          var simCtxM = Object.assign({}, SIM_CTX, { nmod: m });
          var hitM_m = matchedMontagem(base.montagens, simCtxM);
          var val_m = hitM_m ? ev(hitM_m.it, simCtxM, base.blocos) : NaN;
          var isCurrent = (curNmod === m);

          html += '<div class="sim-mod-card' + (isCurrent ? ' active' : '') + '" style="padding:8px 10px; background:' + (isCurrent ? 'color-mix(in srgb, var(--accent) 15%, transparent)' : 'var(--bg-body)') + '; border:1px solid ' + (isCurrent ? 'var(--accent)' : 'var(--border-subtle)') + '; border-radius:6px; transition:all 0.15s ease;">' +
            '<div style="display:flex; justify-content:space-between; align-items:center;">' +
            '<span style="font-weight:700; font-size:12px; color:var(--accent);">' + m + ' Módulo' + (m > 1 ? 's' : '') + '</span>' +
            (isCurrent ? '<span style="font-size:9.5px; font-weight:700; background:var(--accent); color:#fff; padding:1px 5px; border-radius:10px;">SELECIONADO</span>' : '') +
            '</div>' +
            '<div style="font-size:10.5px; color:var(--text-faint); margin:3px 0; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;" title="' + escapeHtml(hitM_m ? hitM_m.nome : '—') + '">' + escapeHtml(hitM_m ? hitM_m.nome : 'Nenhuma') + '</div>' +
            '<div style="font-weight:700; font-size:13.5px; color:var(--text);">' + (isFinite(val_m) ? val_m.toFixed(1).replace('.', ',') + ' h' : '—') + '</div>' +
            '</div>';
        }

        html += '</div></div>';
      }

      html += '</div>';
    } else if (base.forma === 'soma_campos') {
      if (!base.campos) base.campos = [];
      if (!base.etapas) base.etapas = [];
      var areaObj = state.regrasData[state.selectedAreaIdx];
      var allCamposKeys = areaObj && areaObj.campos ? Object.keys(areaObj.campos).filter(function (k) { return k !== state.selectedCampoKey; }) : [];

      html += '<div style="margin-top:10px; display:flex; flex-direction:column; gap:10px;">' +
        '<div style="font-size:11.5px; font-weight:600; color:var(--text-dim); display:flex; justify-content:space-between; align-items:center;">' +
        '<span>Campos a Somar na Regra (H dos campos calculados)</span>' +
        '<span style="font-weight:400; font-size:10.5px; color:var(--text-faint);">' + base.campos.length + ' selecionado(s)</span>' +
        '</div>' +
        '<div style="display:flex; flex-wrap:wrap; gap:6px; background:var(--panel-2); padding:10px; border-radius:8px; border:1px solid var(--border);">';

      allCamposKeys.forEach(function (cKey) {
        var isChecked = base.campos.indexOf(cKey) >= 0;
        html += '<button type="button" class="btn-toggle-campo-soma" data-ckey="' + cKey + '" style="cursor:pointer; padding:4px 10px; font-size:11.5px; font-weight:600; border-radius:6px; border:1px solid ' + (isChecked ? 'var(--accent)' : 'var(--border)') + '; background:' + (isChecked ? 'var(--accent)' : 'var(--bg-body)') + '; color:' + (isChecked ? '#fff' : 'var(--text)') + '; transition:all 0.15s ease;">' +
          (isChecked ? '✓ ' : '+ ') + cKey +
          '</button>';
      });

      html += '</div>' +
        '<div class="scale-chip" style="margin-top:4px;"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M4 12h16"/></svg>Fórmula da Soma: <b>' + (base.campos.length ? 'Σ (' + base.campos.join(' + ') + ')' : 'Nenhum campo selecionado') + '</b></div>' +
        '<div style="margin-top:8px; font-size:11.5px; font-weight:600; color:var(--text-dim); display:flex; justify-content:space-between; align-items:center;">' +
        '<span>Etapas de Ajuste Matemático sobre o Total da Soma</span>' +
        '<span style="font-weight:400; font-size:10.5px; color:var(--text-faint);">' + base.etapas.length + ' etapa(s)</span>' +
        '</div>' +
        '<div id="stepsList" style="display:flex; flex-direction:column; gap:6px;">';

      base.etapas.forEach(function (s, idx) {
        html += stepRowHtml(s, idx, base.etapas.length);
      });

      html += '</div>' +
        '<button type="button" class="btn" id="btnAddStep" style="margin-top:4px; font-size:11.5px; border-style:dashed;">＋ Adicionar Etapa de Ajuste</button>' +
        '</div>';
    } else if (base.forma === 'derivado_h') {
      if (!base.etapas || !base.etapas.length) {
        base.etapas = [];
        if (base.divisao) base.etapas.push({ tipo: 'dividir', valor: base.divisao });
        if (base.arredondamento) base.etapas.push({ tipo: 'arredondar', modo: base.arredondamento });
        if (base.subtracao) base.etapas.push({ tipo: 'subtrair', valor: base.subtracao });
      }

      html += '<div style="margin-top:10px; display:flex; flex-direction:column; gap:6px;">' +
        '<div style="font-size:11.5px; font-weight:600; color:var(--text-dim); display:flex; justify-content:space-between; align-items:center;">' +
        '<span>Etapas Sequenciais da Fórmula (Início: Valor H de Horas)</span>' +
        '<span style="font-weight:400; font-size:10.5px; color:var(--text-faint);">' + base.etapas.length + ' etapa' + (base.etapas.length === 1 ? '' : 's') + '</span>' +
        '</div>' +
        '<div id="stepsList" style="display:flex; flex-direction:column; gap:6px;">';

      base.etapas.forEach(function (s, idx) {
        html += stepRowHtml(s, idx, base.etapas.length);
      });

      html += '</div>' +
        '<button type="button" class="btn" id="btnAddStep" style="margin-top:8px; font-size:11.5px; border-style:dashed;">＋ Adicionar Etapa de Cálculo</button>' +
        '</div>';
    } else if (base.forma === 'tabela') {
      if (!base.valores || base.valores.length < 8) {
        var oldVals = base.valores || [];
        base.valores = [0, 0, 0, 0, 0, 0, 0, 0];
        for (var i = 0; i < Math.min(8, oldVals.length); i++) base.valores[i] = oldVals[i];
      }
      var unit = state.selectedSubTab === 'H' ? 'h' : 'dias';
      html += '<div style="width:100%; margin-top:10px; display:flex; flex-direction:column; gap:6px;">' +
        '<div style="font-size:11.5px; font-weight:600; color:var(--text-dim);">Informe o valor fixo ou fórmula para cada quantidade de módulos (1m a 8m):</div>' +
        '<div style="display:grid; grid-template-columns:repeat(4, 1fr); gap:8px;">';
      for (var m = 1; m <= 8; m++) {
        var valM = base.valores[m - 1] !== undefined ? base.valores[m - 1] : '';
        html += '<div style="display:flex; flex-direction:column; gap:3px; background:var(--panel-2); border:1px solid var(--border); border-radius:8px; padding:6px 8px;">' +
          '<span style="font-size:11px; font-weight:600; color:var(--text-faint); font-family:IBM Plex Mono;">' + m + ' Módulo' + (m > 1 ? 's' : '') + ' (' + m + 'm)</span>' +
          '<div class="num-field" style="width:100%;">' +
          '<input type="text" class="tabela-mod-ipt" data-tabela-mod="' + m + '" value="' + valM + '" placeholder="ex: H/7.92 ou 6" title="Aceita números (ex: 6), expressões (ex: 16*0.9) ou fórmulas com H (ex: H/7.92)">' +
          '</div></div>';
      }
      html += '</div>' +
        '<div class="scale-chip" style="margin-top:8px;"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M4 12h16"/></svg>Valores ou Fórmulas por módulo (aceita números, expressões ou H / 7.92)</div></div>';
    } else {
      var currentBaseVal = base.forma === 'constante' ? (base.valor !== undefined ? base.valor : (base.valor_base || 0)) : (base.valor_base !== undefined ? base.valor_base : (base.valor || 0));
      html += '<div class="base-row" style="gap:10px; align-items:center; margin-top:8px;">' +
        '<div class="num-field"><input type="number" step="0.01" id="baseInput" value="' + currentBaseVal + '"><span>' + (state.selectedSubTab === 'H' ? 'h' : 'dias') + '</span></div>';

      if (base.forma === 'multiplicativa') {
        var valEscala = base.escala !== undefined ? base.escala : 0.5;
        var valAdicao = base.adicao_final || 0;
        var valSubtracao = base.subtracao_final || 0;
        var unit = state.selectedSubTab === 'H' ? 'h' : 'dias';

        var chipExtra = '';
        if (valEscala || valAdicao || valSubtracao) {
          var formulaPart = '(Base × (1 + (m-1)×' + valEscala + '))';
          if (valAdicao) formulaPart += ' + ' + valAdicao + unit;
          if (valSubtracao) formulaPart += ' - ' + valSubtracao + unit;
          chipExtra = ' | 2m+: ' + formulaPart;
        }

        html += '<div class="num-field" title="Quantidade da escala por módulo adicional (ex: 0.50 = +0,5x a cada módulo)"><input type="number" step="0.01" id="escalaInput" value="' + valEscala + '"><span>+×/mód</span></div>' +
          '<div class="num-field" title="Adição de valor fixo no final (a partir do 2º módulo)"><input type="number" step="0.01" id="adicaoFinalInput" value="' + valAdicao + '"><span>+ ' + unit + ' final (2m+)</span></div>' +
          '<div class="num-field" title="Subtração de valor fixo no final (a partir do 2º módulo)"><input type="number" step="0.01" id="subtracaoFinalInput" value="' + valSubtracao + '"><span>- ' + unit + ' final (2m+)</span></div>' +
          '<div class="scale-chip"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 17l6-6 4 4 8-8"/></svg>1m = ' + currentBaseVal + unit + chipExtra + '</div>';
      } else if (base.forma === 'degrau_fixo') {
        var valFator = base.fator !== undefined ? base.fator : (base.escala !== undefined ? base.escala : 1.5);
        var valAdicao = base.adicao_final || 0;
        var valSubtracao = base.subtracao_final || 0;
        var unit = state.selectedSubTab === 'H' ? 'h' : 'dias';
        var val2m = (currentBaseVal * valFator) + valAdicao - valSubtracao;

        html += '<div class="num-field" title="Fator multiplicador fixo para 2 ou mais módulos (ex: 1.50)"><input type="number" step="0.01" id="fatorInput" value="' + valFator + '"><span>× fator (2m+)</span></div>' +
          '<div class="num-field" title="Adição de valor fixo no final (a partir do 2º módulo)"><input type="number" step="0.01" id="adicaoFinalInput" value="' + valAdicao + '"><span>+ ' + unit + ' final (2m+)</span></div>' +
          '<div class="num-field" title="Subtração de valor fixo no final (a partir do 2º módulo)"><input type="number" step="0.01" id="subtracaoFinalInput" value="' + valSubtracao + '"><span>- ' + unit + ' final (2m+)</span></div>' +
          '<div class="scale-chip"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 17l6-6 4 4 8-8"/></svg>1m = ' + currentBaseVal + unit + ' | 2m a 8m = ' + val2m.toFixed(2) + unit + ' (Fixo)</div>';
      } else if (base.forma === 'aditiva') {
        html += '<div class="num-field"><input type="number" step="0.01" id="passoInput" value="' + (base.passo || 0) + '"><span>' + (state.selectedSubTab === 'H' ? 'h/mód' : 'dias/mód') + '</span></div>';
      } else {
        html += '<div class="scale-chip"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M4 12h16"/></svg>Valor constante</div>';
      }
      html += '</div>';
    }
    html += '</div>';

    // Seção de Perfis Especiais (Container Solar & ESSW)
    html += perfisBlockHtml(subTabRule, state.selectedSubTab);

    html += '<div class="field-block"><div class="field-label">Condições adicionais (' + state.selectedSubTab + ') <span class="hint">acréscimo por opção</span></div><div class="cond-list" id="condList">';
    (subTabRule.condicoes || []).forEach(function (c, i) { html += condRowHtml(c, i); });
    html += '</div><button type="button" class="add-cond" id="btnAddCond"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><path d="M12 5v14M5 12h14"/></svg>Adicionar condição</button></div>';

    html += '<div class="field-block"><div class="field-label">Prévia de cálculo por estrutura (' + state.selectedSubTab + ') <span class="hint">1m a 8m, Container Solar e ESSW</span></div>' +
      '<table class="preview-table"><thead><tr><th>Condição</th>' + [1, 2, 3, 4, 5, 6, 7, 8].map(function (m) { return '<th>' + m + 'm</th>'; }).join('') + '<th class="th-special" title="Estrutura Container Solar (Linha fixa)">Solar</th><th class="th-special" title="Estrutura ESSW (Linha fixa)">ESSW</th></tr></thead><tbody id="previewBody"></tbody></table></div>';

    html += '</div>';

    wrap.innerHTML = html;

    // Wire subtab buttons
    wrap.querySelectorAll('.btn-subtab').forEach(function (b) {
      b.addEventListener('click', function () {
        state.selectedSubTab = b.dataset.subtab;
        prepararDirtySubTab();
        renderEditor();
      });
    });

    wireEditorEvents(campoObj[state.selectedSubTab], subTabRule);
    renderPreview(campoObj[state.selectedSubTab], subTabRule);
    markDirty();
    if ($('maintFooterMeta')) $('maintFooterMeta').textContent = areaObj.area + ' · ' + state.selectedCampoKey + ' (' + state.selectedSubTab + ')';
    if ($('btnSaveFooter')) $('btnSaveFooter').textContent = 'Salvar regras de ' + state.selectedCampoKey;
  }

  function condRowHtml(c, i) {
    var opts = FLAGS.map(function (f) { return '<option value="' + f.key + '"' + (f.key === c.flag ? ' selected' : '') + '>' + f.nome + '</option>'; }).join('');
    var valorAtual = c.forma === 'tabela' ? (c.valores ? c.valores[0] : 0) : (c.valor !== undefined ? c.valor : 0);
    var escalaAtual = c.escala !== undefined ? c.escala : 0.5;

    var formaOptions = [
      { tipo: 'fixo', label: '＋ Valor Fixo (+V fixo)' },
      { tipo: 'por_modulo', label: '＋ Adicionar por Módulo (+V × mód)' },
      { tipo: 'escala_multiplicativa', label: 'Escala Multiplicativa' },
      { tipo: 'tabela', label: 'Tabela por Módulo (1m a 8m)' },
      { tipo: 'subtrair', label: '－ Subtrair Fixo (-V fixo)' },
      { tipo: 'subtrair_por_modulo', label: '－ Subtrair por Módulo (-V × mód)' },
      { tipo: 'multiplicar', label: '× Multiplicar Fixo (× V fixo)' },
      { tipo: 'multiplicar_por_modulo', label: '× Multiplicar por Módulo (× V × mód)' },
      { tipo: 'dividir', label: '÷ Dividir Fixo (÷ V fixo)' },
      { tipo: 'dividir_por_modulo', label: '÷ Dividir por Módulo (÷ (V × mód))' }
    ];

    var cForma = c.forma || 'fixo';
    if (cForma === 'subtrair_fixo') cForma = 'subtrair';
    if (cForma === 'multiplicar_fixo') cForma = 'multiplicar';
    if (cForma === 'dividir_fixo') cForma = 'dividir';

    var selectFormaHtml = '<select class="cond-select cond-forma" style="max-width:210px; font-weight:500;">' +
      formaOptions.map(function (opt) {
        return '<option value="' + opt.tipo + '"' + (opt.tipo === cForma ? ' selected' : '') + '>' + opt.label + '</option>';
      }).join('') +
      '</select>';

    var escalaInputHtml = '';
    if (cForma === 'escala_multiplicativa') {
      escalaInputHtml = '<span class="cr-txt" style="margin-left:2px; font-size:11px;">(+</span>' +
        '<input type="number" step="0.01" class="cond-num cond-escala" title="Fator de escala por módulo adicional (ex: 0.50 = +0,5x por módulo)" value="' + escalaAtual + '" style="width:62px;">' +
        '<span class="cr-txt" style="font-size:11px;">×/mód)</span>';
    }

    var tabelaCondHtml = '';
    if (cForma === 'tabela') {
      if (!c.valores || c.valores.length < 8) {
        var oldC = c.valores || [];
        c.valores = [0, 0, 0, 0, 0, 0, 0, 0];
        for (var k = 0; k < Math.min(8, oldC.length); k++) c.valores[k] = oldC[k];
      }
      tabelaCondHtml = '<div style="width:100%; margin-top:8px; display:grid; grid-template-columns:repeat(8, 1fr); gap:6px;">';
      for (var m = 1; m <= 8; m++) {
        var vM = c.valores[m - 1] !== undefined ? c.valores[m - 1] : '';
        tabelaCondHtml += '<div style="display:flex; flex-direction:column; gap:2px; background:var(--panel-2); border:1px solid var(--border); border-radius:6px; padding:4px 6px; text-align:center;">' +
          '<span style="font-size:10px; font-weight:600; color:var(--text-faint); font-family:IBM Plex Mono;">' + m + 'm</span>' +
          '<input type="text" class="cond-tabela-mod" data-cond-idx="' + i + '" data-tabela-mod="' + m + '" value="' + vM + '" placeholder="0" title="Digite um número ou expressão (ex: 16*0.9*0.9)" style="padding:3px; font-size:11px; text-align:center; width:100%; font-weight:600; background:var(--panel-1); border:1px solid var(--border); border-radius:4px; color:var(--text);">' +
          '</div>';
      }
      tabelaCondHtml += '</div>';
    }

    var valorInputHtml = cForma === 'tabela' ? '' : '<span class="cr-txt">marcado, opera</span><input type="text" class="cond-num cond-valor" value="' + valorAtual + '" title="Aceita números ou expressões (ex: 16*0.9*0.9)" style="width:90px;">';

    return '<div class="cond-row" data-idx="' + i + '" style="flex-wrap:wrap;">' +
      '<span class="cr-txt">Se</span><select class="cond-select cond-flag">' + opts + '</select>' +
      valorInputHtml +
      selectFormaHtml +
      escalaInputHtml +
      '<button type="button" class="cond-remove" data-idx="' + i + '"><svg viewBox="0 0 24 24" width="12" height="12" fill="none" stroke="currentColor" stroke-width="2.4"><path d="M6 6l12 12M18 6L6 18"/></svg></button>' +
      tabelaCondHtml +
      '</div>';
  }

  function markDirty() {
    var areaObj = state.regrasData[state.selectedAreaIdx];
    if (!areaObj || !state.selectedCampoKey) return;
    var origSubTab = areaObj.campos[state.selectedCampoKey][state.selectedSubTab];
    var changed = JSON.stringify(state.dirtySubTabRule) !== JSON.stringify(origSubTab);
    ['btnSave', 'btnSaveFooter'].forEach(function (id) {
      var btn = $(id);
      if (btn) btn.disabled = !changed;
    });
  }

  function wireEditorEvents(originalSubTab, subTabRule) {
    var base = subTabRule.base;

    if ($('baseFormaSelect')) {
      $('baseFormaSelect').addEventListener('change', function () {
        base.forma = this.value;
        if (base.forma === 'constante' && base.valor === undefined) base.valor = base.valor_base || 0;
        if (base.forma === 'aditiva' && base.valor_base === undefined) base.valor_base = base.valor || 0;
        if (base.forma === 'multiplicativa') {
          if (base.valor_base === undefined) base.valor_base = base.valor || 0;
          if (base.escala === undefined) base.escala = 0.5;
        }
        if (base.forma === 'degrau_fixo') {
          if (base.valor_base === undefined) base.valor_base = base.valor || 0;
          if (base.fator === undefined) base.fator = 1.5;
        }
        if (base.forma === 'tabela' && (!base.valores || base.valores.length < 8)) {
          var oldVals = base.valores || [];
          base.valores = [0, 0, 0, 0, 0, 0, 0, 0];
          for (var i = 0; i < Math.min(8, oldVals.length); i++) base.valores[i] = oldVals[i];
        }
        if (base.forma === 'soma_campos') {
          if (!base.campos || !base.campos.length) {
            var areaObj = state.regrasData[state.selectedAreaIdx];
            var allKeys = areaObj && areaObj.campos ? Object.keys(areaObj.campos).filter(function (k) { return k !== state.selectedCampoKey; }) : [];
            base.campos = allKeys.slice(0, 3);
          }
          if (!base.etapas) {
            base.etapas = [
              { tipo: 'dividir', valor: 0.9 },
              { tipo: 'multiplicar', valor: 0.1 },
              { tipo: 'multiplicar', valor: 0.8 }
            ];
          }
        }
        if (base.forma === 'derivado_h' && (!base.etapas || !base.etapas.length)) {
          base.etapas = [
            { tipo: 'dividir', valor: 7.92 },
            { tipo: 'arredondar', modo: 'cima' },
            { tipo: 'subtrair', valor: 1.0 }
          ];
        }
        if (base.forma === 'blocos') {
          if (!base.blocos) {
            base.blocos = [
              { id: 'b1', nome: 'Lateral', it: [{ t: 'var', v: 'lat_f1' }, { t: 'op', v: '*' }, { t: 'var', v: 'comp' }, { t: 'op', v: '+' }, { t: 'num', v: 4 }] },
              { id: 'b2', nome: 'Frente/Fundo', it: [{ t: 'var', v: 'ffd_f1' }, { t: 'op', v: '*' }, { t: 'var', v: 'comp' }, { t: 'op', v: '+' }, { t: 'num', v: 4 }] },
              { id: 'b3', nome: 'Teto', it: [{ t: 'var', v: 'tet_f1' }, { t: 'op', v: '*' }, { t: 'var', v: 'comp' }, { t: 'op', v: '+' }, { t: 'var', v: 'tet_f2' }, { t: 'op', v: '*' }, { t: 'var', v: 'larg' }] },
              { id: 'b4', nome: 'Telhado', it: [{ t: 'var', v: 'tlh_f1' }, { t: 'op', v: '*' }, { t: 'var', v: 'comp' }, { t: 'op', v: '+' }, { t: 'var', v: 'tlh_f2' }, { t: 'op', v: '*' }, { t: 'var', v: 'larg' }] },
              { id: 'b5', nome: 'Base', it: [{ t: 'var', v: 'bas_f1' }, { t: 'op', v: '*' }, { t: 'var', v: 'comp' }, { t: 'op', v: '+' }, { t: 'var', v: 'bas_f2' }, { t: 'op', v: '*' }, { t: 'var', v: 'larg' }] }
            ];
          }
          if (!base.montagens) {
            base.montagens = [
              { id: 'm1', nome: 'Móvel', cond: [{ c: 'tipoestrutura', o: '=', val: 'Móvel', j: 'E' }], it: [{ t: 'blk', v: 'b5' }, { t: 'op', v: '*' }, { t: 'var', v: 'bas_mb' }] },
              { id: 'm2', nome: 'Embarcado', cond: [{ c: 'tipoestrutura', o: '=', val: 'Embarcado', j: 'E' }], it: [{ t: 'blk', v: 'b1' }, { t: 'op', v: '*' }, { t: 'var', v: 'lat_mb' }, { t: 'op', v: '+' }, { t: 'num', v: 2 }, { t: 'op', 'v': '*' }, { t: 'blk', v: 'b2' }, { t: 'op', v: '*' }, { t: 'var', v: 'ffd_mb' }, { t: 'op', v: '+' }, { t: 'blk', v: 'b3' }, { t: 'op', v: '*' }, { t: 'var', v: 'tet_mb' }, { t: 'op', v: '+' }, { t: 'blk', v: 'b5' }, { t: 'op', v: '*' }, { t: 'var', v: 'bas_mb' }] },
              { id: 'm0', nome: 'Demais estruturas', padrao: true, cond: [], it: [{ t: 'blk', v: 'b1' }, { t: 'op', v: '*' }, { t: 'var', v: 'lat_mb' }, { t: 'op', v: '+' }, { t: 'num', v: 2 }, { t: 'op', 'v': '*' }, { t: 'blk', v: 'b2' }, { t: 'op', v: '*' }, { t: 'var', v: 'ffd_mb' }, { t: 'op', v: '+' }, { t: 'blk', v: 'b3' }, { t: 'op', v: '*' }, { t: 'var', v: 'tet_mb' }, { t: 'op', v: '+' }, { t: 'num', v: 2 }, { t: 'op', 'v': '*' }, { t: 'blk', v: 'b4' }, { t: 'op', v: '*' }, { t: 'var', v: 'tlh_mb' }, { t: 'op', v: '+' }, { t: 'blk', v: 'b5' }, { t: 'op', v: '*' }, { t: 'var', v: 'bas_mb' }] }
            ];
          }
        }
        markDirty();
        renderEditor();
      });
    }

    document.querySelectorAll('#colEditor .tabela-mod-ipt').forEach(function (inp) {
      inp.addEventListener('input', function () {
        var mIdx = (+inp.dataset.tabelaMod) - 1;
        if (!base.valores) base.valores = [0, 0, 0, 0, 0, 0, 0, 0];
        base.valores[mIdx] = this.value;
        markDirty();
        renderPreview(originalSubTab, subTabRule);
      });
    });

    document.querySelectorAll('.btn-toggle-campo-soma').forEach(function (btn) {
      btn.addEventListener('click', function () {
        var cKey = this.dataset.ckey;
        if (!base.campos) base.campos = [];
        var idx = base.campos.indexOf(cKey);
        if (idx >= 0) {
          base.campos.splice(idx, 1);
        } else {
          base.campos.push(cKey);
        }
        markDirty();
        renderEditor();
      });
    });

    if ($('btnAddStep')) {
      $('btnAddStep').addEventListener('click', function () {
        if (!base.etapas) base.etapas = [];
        base.etapas.push({ tipo: 'somar', valor: 1 });
        renderEditor();
      });
    }
    wireStepRows(originalSubTab, subTabRule);

    if ($('baseInput')) {
      $('baseInput').addEventListener('input', function () {
        var val = parseFloat(this.value) || 0;
        base.valor = val;
        base.valor_base = val;
        markDirty(); renderPreview(originalSubTab, subTabRule);
      });
    }
    if ($('passoInput')) {
      $('passoInput').addEventListener('input', function () {
        base.passo = parseFloat(this.value) || 0;
        markDirty(); renderPreview(originalSubTab, subTabRule);
      });
    }
    if ($('escalaInput')) {
      $('escalaInput').addEventListener('input', function () {
        base.escala = parseFloat(this.value) || 0;
        markDirty(); renderPreview(originalSubTab, subTabRule);
      });
    }
    if ($('fatorInput')) {
      $('fatorInput').addEventListener('input', function () {
        base.fator = parseFloat(this.value) || 0;
        markDirty(); renderPreview(originalSubTab, subTabRule);
      });
    }
    if ($('adicaoFinalInput')) {
      $('adicaoFinalInput').addEventListener('input', function () {
        base.adicao_final = parseFloat(this.value) || 0;
        markDirty(); renderPreview(originalSubTab, subTabRule);
      });
    }
    if ($('subtracaoFinalInput')) {
      $('subtracaoFinalInput').addEventListener('input', function () {
        base.subtracao_final = parseFloat(this.value) || 0;
        markDirty(); renderPreview(originalSubTab, subTabRule);
      });
    }

    // Wiring Collapsible Accordion & Special Rules (Solar / ESSW)
    var btnToggleEsp = $('toggleSpecialRules');
    if (btnToggleEsp) {
      btnToggleEsp.addEventListener('click', function () {
        state.specialRulesExpanded = !state.specialRulesExpanded;
        var bodyEl = $('specialRulesBody');
        var chevEl = $('accChevron');
        if (bodyEl) bodyEl.style.display = state.specialRulesExpanded ? 'block' : 'none';
        if (chevEl) chevEl.textContent = state.specialRulesExpanded ? '▲' : '▼';
      });
    }

    var esp = getEspeciais(subTabRule);

    // Solar Wiring
    var chkSolar = $('chkEspSolar');
    if (chkSolar) {
      chkSolar.addEventListener('change', function () {
        esp.solar.ativo = this.checked;
        if (this.checked && (!esp.solar.base || esp.solar.base.valor === undefined || esp.solar.base.valor === 0)) {
          var curBase = base.valor !== undefined ? base.valor : (base.valor_base || 0);
          esp.solar.base = { forma: 'constante', valor: curBase, valor_base: curBase };
        }
        syncEspeciaisToPerfis(subTabRule);
        markDirty();
        renderEditor();
      });
    }

    var selFormaSolar = $('selFormaEspSolar');
    if (selFormaSolar) {
      selFormaSolar.addEventListener('change', function () {
        esp.solar.base.forma = this.value;
        syncEspeciaisToPerfis(subTabRule);
        markDirty();
        renderPreview(originalSubTab, subTabRule);
      });
    }

    var iptValSolar = $('iptValEspSolar');
    if (iptValSolar) {
      iptValSolar.addEventListener('input', function () {
        var num = parseFloat(this.value.replace(',', '.'));
        esp.solar.base.valor = !isNaN(num) ? num : this.value;
        esp.solar.base.valor_base = !isNaN(num) ? num : this.value;
        syncEspeciaisToPerfis(subTabRule);
        markDirty();
        renderPreview(originalSubTab, subTabRule);
      });
    }

    var chkHerdarSolar = $('chkHerdarSolar');
    if (chkHerdarSolar) {
      chkHerdarSolar.addEventListener('change', function () {
        esp.solar.herdar_condicoes = this.checked;
        markDirty();
        renderPreview(originalSubTab, subTabRule);
      });
    }

    // ESSW Wiring
    var chkEssw = $('chkEspEssw');
    if (chkEssw) {
      chkEssw.addEventListener('change', function () {
        esp.essw.ativo = this.checked;
        if (this.checked && (!esp.essw.base || esp.essw.base.valor === undefined || esp.essw.base.valor === 0)) {
          var curBase = base.valor !== undefined ? base.valor : (base.valor_base || 0);
          esp.essw.base = { forma: 'constante', valor: curBase, valor_base: curBase };
        }
        syncEspeciaisToPerfis(subTabRule);
        markDirty();
        renderEditor();
      });
    }

    var selFormaEssw = $('selFormaEspEssw');
    if (selFormaEssw) {
      selFormaEssw.addEventListener('change', function () {
        esp.essw.base.forma = this.value;
        syncEspeciaisToPerfis(subTabRule);
        markDirty();
        renderPreview(originalSubTab, subTabRule);
      });
    }

    var iptValEssw = $('iptValEspEssw');
    if (iptValEssw) {
      iptValEssw.addEventListener('input', function () {
        var num = parseFloat(this.value.replace(',', '.'));
        esp.essw.base.valor = !isNaN(num) ? num : this.value;
        esp.essw.base.valor_base = !isNaN(num) ? num : this.value;
        syncEspeciaisToPerfis(subTabRule);
        markDirty();
        renderPreview(originalSubTab, subTabRule);
      });
    }

    var chkHerdarEssw = $('chkHerdarEssw');
    if (chkHerdarEssw) {
      chkHerdarEssw.addEventListener('change', function () {
        esp.essw.herdar_condicoes = this.checked;
        markDirty();
        renderPreview(originalSubTab, subTabRule);
      });
    }

    // Wiring Add Condition Buttons for Solar / ESSW
    document.querySelectorAll('.btn-add-esp-cond').forEach(function (btn) {
      btn.addEventListener('click', function () {
        var espKey = this.dataset.esp;
        if (esp[espKey]) {
          if (!esp[espKey].condicoes) esp[espKey].condicoes = [];
          esp[espKey].condicoes.push({ flag: FLAGS[0].key, rotulo: FLAGS[0].nome, forma: 'fixo', valor: 1 });
          markDirty();
          renderEditor();
        }
      });
    });

    // Wiring Condition Rows for Solar / ESSW
    document.querySelectorAll('.esp-cond-row').forEach(function (row) {
      var espKey = row.dataset.esp;
      var idx = +row.dataset.idx;
      var targetList = esp[espKey] ? esp[espKey].condicoes : null;
      if (!targetList || !targetList[idx]) return;
      var c = targetList[idx];

      var flagSel = row.querySelector('.esp-cond-flag');
      if (flagSel) {
        flagSel.addEventListener('change', function () {
          c.flag = this.value;
          c.rotulo = flagNome(this.value);
          markDirty();
          renderPreview(originalSubTab, subTabRule);
        });
      }

      var valIpt = row.querySelector('.esp-cond-valor');
      if (valIpt) {
        valIpt.addEventListener('input', function () {
          c.valor = this.value;
          markDirty();
          renderPreview(originalSubTab, subTabRule);
        });
      }

      var formaSel = row.querySelector('.esp-cond-forma');
      if (formaSel) {
        formaSel.addEventListener('change', function () {
          c.forma = this.value;
          markDirty();
          renderPreview(originalSubTab, subTabRule);
        });
      }

      var delBtn = row.querySelector('.esp-cond-remove');
      if (delBtn) {
        delBtn.addEventListener('click', function () {
          targetList.splice(idx, 1);
          markDirty();
          renderEditor();
        });
      }
    });

    if ($('btnAddCond')) {
      $('btnAddCond').addEventListener('click', function () {
        subTabRule.condicoes.push({ flag: FLAGS[0].key, rotulo: FLAGS[0].nome, forma: 'fixo', valor: 1 });
        renderEditor();
      });
    }
    wireCondRows(originalSubTab, subTabRule);
    wireBlocosEvents(subTabRule);

    if ($('btnCancel')) $('btnCancel').addEventListener('click', function () { prepararDirtySubTab(); renderEditor(); });
    if ($('btnSave')) $('btnSave').addEventListener('click', function () { saveRegrasCampo(); });
  }

  function wireStepRows(originalSubTab, subTabRule) {
    var base = subTabRule.base;
    if (!base.etapas) return;
    document.querySelectorAll('.step-row').forEach(function (row) {
      var idx = +row.dataset.stepIdx;
      var step = base.etapas[idx];
      if (!step) return;

      var tipoSel = row.querySelector('.step-tipo');
      if (tipoSel) {
        tipoSel.addEventListener('change', function () {
          step.tipo = this.value;
          if (step.tipo === 'arredondar' && !step.modo) step.modo = 'cima';
          if (step.tipo !== 'arredondar' && step.valor === undefined) step.valor = 1;
          renderEditor();
        });
      }

      var valorIpt = row.querySelector('.step-valor');
      if (valorIpt) {
        valorIpt.addEventListener('input', function () {
          step.valor = parseFloat(this.value) || 0;
          markDirty();
          renderPreview(originalSubTab, subTabRule);
        });
      }

      var modoSel = row.querySelector('.step-modo');
      if (modoSel) {
        modoSel.addEventListener('change', function () {
          step.modo = this.value;
          markDirty();
          renderPreview(originalSubTab, subTabRule);
        });
      }

      var btnUp = row.querySelector('.step-up');
      if (btnUp) {
        btnUp.addEventListener('click', function () {
          if (idx > 0) {
            var tmp = base.etapas[idx - 1];
            base.etapas[idx - 1] = base.etapas[idx];
            base.etapas[idx] = tmp;
            renderEditor();
          }
        });
      }

      var btnDown = row.querySelector('.step-down');
      if (btnDown) {
        btnDown.addEventListener('click', function () {
          if (idx < base.etapas.length - 1) {
            var tmp = base.etapas[idx + 1];
            base.etapas[idx + 1] = base.etapas[idx];
            base.etapas[idx] = tmp;
            renderEditor();
          }
        });
      }

      var btnDel = row.querySelector('.step-del');
      if (btnDel) {
        btnDel.addEventListener('click', function () {
          base.etapas.splice(idx, 1);
          renderEditor();
        });
      }
    });
  }

  function wireCondRows(originalSubTab, subTabRule) {
    document.querySelectorAll('#condList > .cond-row').forEach(function (row) {
      var idx = +row.dataset.idx;
      var c = subTabRule.condicoes ? subTabRule.condicoes[idx] : null;
      if (!c) return;

      var flagSel = row.querySelector('.cond-flag');
      if (flagSel) {
        flagSel.addEventListener('change', function () {
          c.flag = this.value;
          c.rotulo = flagNome(this.value);
          markDirty();
          renderPreview(originalSubTab, subTabRule);
        });
      }
      var valorIpt = row.querySelector('.cond-valor');
      if (valorIpt) {
        valorIpt.addEventListener('input', function () {
          c.valor = this.value;
          markDirty();
          renderPreview(originalSubTab, subTabRule);
        });
      }
      var formaSel = row.querySelector('.cond-forma');
      if (formaSel) {
        formaSel.addEventListener('change', function () {
          c.forma = this.value;
          if (c.forma === 'escala_multiplicativa' && c.escala === undefined) c.escala = 0.5;
          if (c.forma === 'tabela' && (!c.valores || c.valores.length < 8)) {
            c.valores = [0, 0, 0, 0, 0, 0, 0, 0];
          }
          markDirty();
          renderEditor();
        });
      }
      var escalaIpt = row.querySelector('.cond-escala');
      if (escalaIpt) {
        escalaIpt.addEventListener('input', function () {
          c.escala = parseFloat(this.value) || 0;
          markDirty();
          renderPreview(originalSubTab, subTabRule);
        });
      }
      row.querySelectorAll('.cond-tabela-mod').forEach(function (inp) {
        inp.addEventListener('input', function () {
          var mIdx = (+inp.dataset.tabelaMod) - 1;
          if (!c.valores) c.valores = [0, 0, 0, 0, 0, 0, 0, 0];
          c.valores[mIdx] = this.value;
          markDirty();
          renderPreview(originalSubTab, subTabRule);
        });
      });
      var remBtn = row.querySelector('.cond-remove');
      if (remBtn) {
        remBtn.addEventListener('click', function () {
          subTabRule.condicoes.splice(idx, 1);
          renderEditor();
        });
      }
    });
  }

  function renderPreview(originalSubTab, subTabRule) {
    var body = $('previewBody');
    if (!body) return;
    var areaObj = state.regrasData[state.selectedAreaIdx];
    var campoObj = (areaObj && areaObj.campos) ? areaObj.campos[state.selectedCampoKey] : null;

    // Collect all condition flags from both H and DUR (and subTabRule and especiais)
    var conds = [];
    if (campoObj && campoObj.H && campoObj.H.condicoes) conds = conds.concat(campoObj.H.condicoes);
    if (campoObj && campoObj.DUR && campoObj.DUR.condicoes) conds = conds.concat(campoObj.DUR.condicoes);
    if (subTabRule && subTabRule.condicoes) conds = conds.concat(subTabRule.condicoes);
    var espRule = getEspeciais(subTabRule);
    if (espRule.solar && espRule.solar.condicoes) conds = conds.concat(espRule.solar.condicoes);
    if (espRule.essw && espRule.essw.condicoes) conds = conds.concat(espRule.essw.condicoes);

    var uniqueFlags = [];
    var seenFlags = {};
    conds.forEach(function (c) {
      if (c && c.flag && !seenFlags[c.flag]) {
        seenFlags[c.flag] = true;
        uniqueFlags.push(c);
      }
    });

    var cols = [
      { k: '1m', m: 1, ctx: { tipoestrutura: 'Móvel', nmod: 1 } },
      { k: '2m', m: 2, ctx: { tipoestrutura: 'Móvel', nmod: 2 } },
      { k: '3m', m: 3, ctx: { tipoestrutura: 'Móvel', nmod: 3 } },
      { k: '4m', m: 4, ctx: { tipoestrutura: 'Móvel', nmod: 4 } },
      { k: '5m', m: 5, ctx: { tipoestrutura: 'Móvel', nmod: 5 } },
      { k: '6m', m: 6, ctx: { tipoestrutura: 'Móvel', nmod: 6 } },
      { k: '7m', m: 7, ctx: { tipoestrutura: 'Móvel', nmod: 7 } },
      { k: '8m', m: 8, ctx: { tipoestrutura: 'Móvel', nmod: 8 } },
      { k: 'Solar', m: 1, ctx: { tipoestrutura: 'Container Solar', nmod: 1 }, special: true },
      { k: 'ESSW', m: 1, ctx: { tipoestrutura: 'ESSW (mecânica)', nmod: 1 }, special: true }
    ];

    function rowFor(label, flagsAtivos) {
      var isDur = state.selectedSubTab === 'DUR';
      var subTabH = campoObj ? campoObj['H'] : null;

      var hOrigVals = cols.map(function (col) {
        return (isDur && subTabH) ? calcValor(subTabH, col.m, flagsAtivos, campoObj, 0, col.ctx) : undefined;
      });
      var hNewVals = cols.map(function (col) {
        return (isDur && subTabH) ? calcValor(subTabH, col.m, flagsAtivos, campoObj, 0, col.ctx) : undefined;
      });

      var origVals = cols.map(function (col, idx) {
        return calcValor(originalSubTab, col.m, flagsAtivos, campoObj, isDur ? hOrigVals[idx] : undefined, col.ctx);
      });
      var newVals = cols.map(function (col, idx) {
        return calcValor(subTabRule, col.m, flagsAtivos, campoObj, isDur ? hNewVals[idx] : undefined, col.ctx);
      });

      return '<tr><td>' + label + '</td>' + newVals.map(function (v, i) {
        var changed = v !== null && origVals[i] !== null && Math.abs(v - origVals[i]) > 0.001;
        var cls = [];
        if (cols[i].special) cls.push('td-special');
        if (changed) cls.push('changed');
        return '<td class="' + cls.join(' ') + '">' + (v !== null ? (typeof v === 'number' ? v.toFixed(2) : v) : '—') + '</td>';
      }).join('') + '</tr>';
    }

    var html = rowFor('nenhuma marcada', {});
    uniqueFlags.forEach(function (c) {
      var m = {}; m[c.flag] = true;
      html += rowFor((c.rotulo || flagNome(c.flag)) + ' marcado', m);
    });
    body.innerHTML = html;
  }

  function saveRegrasCampo(motivo, anexosList) {
    ['btnSave', 'btnSaveFooter'].forEach(function (id) {
      var btn = $(id);
      if (btn) { btn.disabled = true; btn.textContent = 'Salvando…'; }
    });

    var areaObj = state.regrasData[state.selectedAreaIdx];
    areaObj.campos[state.selectedCampoKey][state.selectedSubTab] = JSON.parse(JSON.stringify(state.dirtySubTabRule));

    function resetSaveButtons() {
      ['btnSave', 'btnSaveFooter'].forEach(function (id) {
        var btn = $(id);
        if (btn) {
          btn.disabled = false;
          btn.textContent = 'Salvar regras de ' + state.selectedCampoKey;
        }
      });
    }

    var payload = {
      regras: state.regrasData,
      motivo: motivo || '',
      anexos: anexosList || []
    };

    if (isPyWebviewAvailable()) {
      window.pywebview.api.save_regras(payload).then(function (res) {
        if (res && res.status === 'conflito') {
          showToast('⚠️ Conflito: ' + (res.message || 'Outro usuário salvou antes.'), true);
          alert('⚠️ Conflito de Salvamento:\n\n' + (res.message || 'Outro usuário salvou alterações antes.') + '\n\nOs dados mais recentes serão recarregados do servidor MySQL.');
          carregarDisciplinas();
          resetSaveButtons();
        } else if (res && res.status === 'error') {
          showToast('Erro ao salvar: ' + (res.message || 'Erro desconhecido'), true);
          resetSaveButtons();
        } else {
          onRegrasSaveSuccess();
        }
      }).catch(function (err) {
        showToast('Erro ao salvar regras: ' + (err ? (err.message || err) : 'Erro'), true);
        resetSaveButtons();
      });
    } else {
      fetch('/api/save_regras', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      }).then(function (r) { return r.json(); }).then(function (res) {
        if (res && res.status === 'conflito') {
          showToast('⚠️ Conflito: ' + (res.message || 'Outro usuário salvou antes.'), true);
          alert('⚠️ Conflito de Salvamento:\n\n' + (res.message || 'Outro usuário salvou antes.') + '\n\nOs dados mais recentes serão recarregados do servidor MySQL.');
          carregarDisciplinas();
          resetSaveButtons();
        } else if (res && res.status === 'error') {
          showToast('Erro ao salvar: ' + (res.message || 'Erro desconhecido'), true);
          resetSaveButtons();
        } else {
          showToast('Regras salvas na sessão!');
          onRegrasSaveSuccess();
        }
      }).catch(function () {
        showToast('Regras salvas localmente!');
        onRegrasSaveSuccess();
      });
    }
  }

  var currentAnexosList = [];

  function resetAnexoInput() {
    currentAnexosList = [];
    var ipt = $('inputAnexoSave');
    if (ipt) ipt.value = '';
    renderAnexosModalList();
  }

  function renderAnexosModalList() {
    var listEl = $('anexosFileList');
    var lbl = $('anexoLabelText');
    var area = $('anexoDropArea');
    if (!listEl) return;

    if (!currentAnexosList.length) {
      listEl.innerHTML = '';
      if (lbl) lbl.textContent = 'Clique para selecionar arquivo(s) do seu computador...';
      if (area) area.classList.remove('has-file');
      return;
    }

    if (lbl) lbl.textContent = '＋ Adicionar mais arquivo(s) (' + currentAnexosList.length + ' selecionado' + (currentAnexosList.length === 1 ? '' : 's') + ')...';
    if (area) area.classList.add('has-file');

    var html = '';
    currentAnexosList.forEach(function (fileObj, idx) {
      html += '<div class="anexo-item-modal">' +
        '<div class="anexo-item-name">' +
        '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="m21.44 11.05-9.19 9.19a6 6 0 0 1-8.49-8.49l8.57-8.57A4 4 0 1 1 18 8.84l-8.59 8.57a2 2 0 0 1-2.83-2.83l8.49-8.48"/></svg>' +
        '<span>' + escapeHtml(fileObj.nome) + ' (' + formatBytes(fileObj.tamanho || 0) + ')</span>' +
        '</div>' +
        '<button type="button" class="anexo-item-del" data-idx="' + idx + '" title="Remover este anexo">' +
        '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" style="width:13px;height:13px;"><path d="M18 6 6 18M6 6l12 12"/></svg>' +
        '</button>' +
        '</div>';
    });
    listEl.innerHTML = html;

    listEl.querySelectorAll('.anexo-item-del').forEach(function (btn) {
      btn.addEventListener('click', function (e) {
        e.stopPropagation();
        var idx = +this.dataset.idx;
        currentAnexosList.splice(idx, 1);
        renderAnexosModalList();
      });
    });
  }

  function formatBytes(bytes) {
    if (!bytes || bytes === 0) return '0 B';
    var k = 1024;
    var sizes = ['B', 'KB', 'MB', 'GB'];
    var i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
  }

  function updateMotivoCharCounter() {
    var input = $('inputMotivoSave');
    var counter = $('motivoCharCount');
    var btn = $('btnConfirmSaveMotivo');
    if (!input) return;
    var len = input.value.trim().length;
    if (counter) {
      if (len >= 20) {
        counter.textContent = len + ' / 20 caracteres (OK)';
        counter.style.color = 'var(--green)';
      } else {
        counter.textContent = len + ' / 20 mín.';
        counter.style.color = 'var(--red)';
      }
    }
    if (btn) {
      btn.disabled = len < 20;
    }
  }

  function promptMotivoESalvar() {
    var modal = $('modalMotivo');
    var input = $('inputMotivoSave');
    resetAnexoInput();
    if (modal && input) {
      input.value = '';
      updateMotivoCharCounter();
      modal.classList.add('open');
      setTimeout(function () { input.focus(); }, 100);
    } else {
      saveRegrasCampo('', []);
    }
  }

  function handleFilesInput(filesList) {
    var files = Array.from(filesList || []);
    if (!files.length) return;

    files.forEach(function (file) {
      if (file.size > 50 * 1024 * 1024) {
        showToast('O arquivo "' + file.name + '" é muito grande (máximo 50MB).', true);
        return;
      }

      var reader = new FileReader();
      reader.onload = function (evt) {
        currentAnexosList.push({
          nome: file.name,
          base64: evt.target.result,
          tamanho: file.size
        });
        renderAnexosModalList();
      };
      reader.onerror = function () {
        showToast('Erro ao ler o arquivo "' + file.name + '".', true);
      };
      reader.readAsDataURL(file);
    });
  }

  var dropAreaEl = $('anexoDropArea');
  if (dropAreaEl) {
    dropAreaEl.addEventListener('click', function (e) {
      if (e.target && (e.target.classList.contains('anexo-item-del') || e.target.closest('.anexo-item-del'))) {
        return;
      }
      if ($('inputAnexoSave')) $('inputAnexoSave').click();
    });

    ['dragenter', 'dragover'].forEach(function (evtName) {
      dropAreaEl.addEventListener(evtName, function (e) {
        e.preventDefault();
        e.stopPropagation();
        dropAreaEl.classList.add('drag-over');
      });
    });

    ['dragleave', 'dragend'].forEach(function (evtName) {
      dropAreaEl.addEventListener(evtName, function (e) {
        e.preventDefault();
        e.stopPropagation();
        dropAreaEl.classList.remove('drag-over');
      });
    });

    dropAreaEl.addEventListener('drop', function (e) {
      e.preventDefault();
      e.stopPropagation();
      dropAreaEl.classList.remove('drag-over');
      if (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files.length) {
        handleFilesInput(e.dataTransfer.files);
      }
    });
  }

  if ($('inputAnexoSave')) {
    $('inputAnexoSave').addEventListener('change', function (e) {
      handleFilesInput(this.files);
      this.value = '';
    });
  }

  if ($('inputMotivoSave')) {
    $('inputMotivoSave').addEventListener('input', updateMotivoCharCounter);
    $('inputMotivoSave').addEventListener('keydown', function (e) {
      if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) {
        e.preventDefault();
        var len = this.value.trim().length;
        if (len >= 20 && $('btnConfirmSaveMotivo')) {
          $('btnConfirmSaveMotivo').click();
        } else if (len < 20) {
          showToast('O motivo da alteração deve conter pelo menos 20 caracteres.', true);
        }
      }
    });
  }

  if ($('btnSaveFooter')) {
    $('btnSaveFooter').addEventListener('click', promptMotivoESalvar);
  }

  if ($('btnConfirmSaveMotivo')) {
    $('btnConfirmSaveMotivo').addEventListener('click', function () {
      var motivo = ($('inputMotivoSave') ? $('inputMotivoSave').value : '').trim();
      if (motivo.length < 20) {
        showToast('O motivo da alteração precisa ter no mínimo 20 caracteres.', true);
        return;
      }
      var anexosPayload = currentAnexosList.map(function (x) {
        return { nome: x.nome, base64: x.base64 };
      });
      if ($('modalMotivo')) $('modalMotivo').classList.remove('open');
      if (state.isSeletorActive) {
        salvarSeletor(motivo, anexosPayload);
      } else {
        saveRegrasCampo(motivo, anexosPayload);
      }
    });
  }

  if ($('btnCancelMotivo')) {
    $('btnCancelMotivo').addEventListener('click', function () {
      if ($('modalMotivo')) $('modalMotivo').classList.remove('open');
      resetAnexoInput();
    });
  }

  if ($('btnCloseMotivoModal')) {
    $('btnCloseMotivoModal').addEventListener('click', function () {
      if ($('modalMotivo')) $('modalMotivo').classList.remove('open');
      resetAnexoInput();
    });
  }

  if ($('inputMotivoSave')) {
    $('inputMotivoSave').addEventListener('keydown', function (e) {
      if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) {
        e.preventDefault();
        if ($('btnConfirmSaveMotivo')) $('btnConfirmSaveMotivo').click();
      }
    });
  }

  if ($('btnToggleLegend')) {
    $('btnToggleLegend').addEventListener('click', function () {
      var bar = $('chipLegendBar');
      if (!bar) return;
      var isHidden = bar.classList.contains('hidden');
      if (isHidden) {
        bar.classList.remove('hidden');
        this.querySelector('span').textContent = 'Ocultar legenda';
      } else {
        bar.classList.add('hidden');
        this.querySelector('span').textContent = 'Mostrar legenda';
      }
    });
  }

  function onRegrasSaveSuccess() {
    state.historico.unshift({
      disciplina: state.regrasData[state.selectedAreaIdx].area,
      campo: state.selectedCampoKey + ' (' + state.selectedSubTab + ')',
      quando: new Date().toLocaleString('pt-BR'),
      resumo: 'Regra de ' + state.selectedSubTab + ' atualizada'
    });
    carregarHistoricoBD();
    showToast('Regras do campo ' + state.selectedCampoKey + ' salvas com sucesso!');
    prepararDirtySubTab();
    renderList();
    renderEditor();
  }

  function renderHistorico() {
    carregarHistoricoBD();
  }

  /* ==========================================================================
     HISTORY POPUP MODAL LOGIC (BANCO DE DADOS MYSQL)
     ========================================================================== */
  var histModalOverlay = $('histModalOverlay');
  var histModalBody = $('histModalBody');
  var histSearchInput = $('histSearchInput');
  var histFooterCount = $('histFooterCount');
  var currentLogs = [];

  function openHist() {
    if (histModalOverlay) {
      histModalOverlay.classList.add('open');
      carregarHistoricoBD();
    }
  }

  function closeHist() {
    if (histModalOverlay) {
      histModalOverlay.classList.remove('open');
    }
  }

  if (btnHist) btnHist.addEventListener('click', openHist);
  if ($('btnHistFooter')) $('btnHistFooter').addEventListener('click', openHist);
  if ($('btnCloseHist')) $('btnCloseHist').addEventListener('click', closeHist);
  if ($('btnCloseHistFoot')) $('btnCloseHistFoot').addEventListener('click', closeHist);
  if ($('btnRefreshHist')) $('btnRefreshHist').addEventListener('click', carregarHistoricoBD);
  if (histModalOverlay) {
    histModalOverlay.addEventListener('click', function (e) {
      if (e.target === histModalOverlay) closeHist();
    });
  }

  if (histSearchInput) {
    histSearchInput.addEventListener('input', function () {
      filtrarERenderizarLogs();
    });
  }

  function carregarHistoricoBD() {
    if (!histModalBody) return;
    histModalBody.innerHTML = '<div class="hist-empty">Carregando histórico do banco de dados…</div>';

    if (isPyWebviewAvailable()) {
      window.pywebview.api.get_logs(100).then(function (logs) {
        currentLogs = logs || [];
        filtrarERenderizarLogs();
      }).catch(function (err) {
        console.warn('[Histórico BD] Erro via PyWebView:', err);
        renderHistoricoFallback();
      });
    } else {
      fetch('/api/logs').then(function (resp) {
        if (resp.ok) return resp.json();
      }).then(function (logs) {
        if (logs && Array.isArray(logs)) {
          currentLogs = logs;
          filtrarERenderizarLogs();
        } else {
          renderHistoricoFallback();
        }
      }).catch(function () {
        renderHistoricoFallback();
      });
    }
  }

  function filtrarERenderizarLogs() {
    var query = (histSearchInput ? histSearchInput.value : '').trim().toLowerCase();
    var filtrados = currentLogs;

    if (query) {
      filtrados = currentLogs.filter(function (log) {
        var txt = (log.grupo_area + ' ' + log.regra_campo + ' ' + log.subtab + ' ' + log.usuario + ' ' + log.data_hora + ' ' + (log.detalhes || '')).toLowerCase();
        return txt.indexOf(query) !== -1;
      });
    }

    renderLogsCards(filtrados);
  }

  function renderLogsCards(logs) {
    if (!histModalBody) return;

    if (histFooterCount) {
      histFooterCount.textContent = logs.length + ' alteraç' + (logs.length === 1 ? 'ão encontrada' : 'ões encontradas');
    }

    if (!logs || !logs.length) {
      histModalBody.innerHTML = '<div class="hist-empty">Nenhuma alteração de regra encontrada no histórico.</div>';
      return;
    }

    histModalBody.innerHTML = '';

    logs.forEach(function (log) {
      var subtab = (log.subtab || 'H').toUpperCase();
      var subtabClass = subtab.toLowerCase() === 'dur' ? 'dur' : '';
      var campoNome = log.regra_campo || 'Regra';
      var areaNome = log.grupo_area || 'Área';
      var usuario = log.usuario || 'usuario';

      // Formata data abreviada: "10/08 16:31"
      var dataStr = log.data_hora || '';
      if (dataStr.length >= 16) {
        var partes = dataStr.split(' ');
        if (partes.length >= 2) {
          var dataPart = partes[0].substring(0, 5); // "10/08"
          var horaPart = partes[1].substring(0, 5); // "16:31"
          dataStr = dataPart + ' ' + horaPart;
        }
      }

      var diffObj = gerarDiffListasHtml(log.antes, log.depois);

      var item = document.createElement('div');
      item.className = 'hitem';

      item.innerHTML =
        '<div class="hitem-row">' +
        '<span class="htag ' + subtabClass + '">' + escapeHtml(subtab) + '</span>' +
        '<span class="hname">' + escapeHtml(campoNome) + '</span>' +
        '<span class="hsep">·</span>' +
        '<span class="harea">' + escapeHtml(areaNome) + '</span>' +
        '<span class="hmeta">' +
        '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="4" width="18" height="18" rx="2"/><path d="M16 2v4M8 2v4M3 10h18"/></svg>' +
        ' ' + escapeHtml(dataStr) +
        '<span class="u">' + escapeHtml(usuario) + '</span>' +
        '</span>' +
        '<span class="hchevron"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"><path d="m18 15-6-6-6 6"/></svg></span>' +
        '</div>' +

        '<div class="hdetail">' +
        (log.motivo ? '<div class="hmotivo-box"><strong>Motivo:</strong> ' + escapeHtml(log.motivo) + '</div>' : '') +
        (function () {
          var anexosList = log.anexos || [];
          if (!anexosList.length && (log.anexo_caminho || log.anexo_nome)) {
            anexosList = [{
              nome: log.anexo_nome || (log.anexo_caminho ? log.anexo_caminho.split(/[\\/]/).pop() : 'Anexo'),
              caminho: log.anexo_caminho
            }];
          }
          if (!anexosList.length) return '';
          var h = '<div style="display:flex; flex-direction:column; gap:6px; margin-bottom:10px;">';
          anexosList.forEach(function (anx) {
            var aNome = anx.nome || (anx.caminho ? anx.caminho.split(/[\\/]/).pop() : 'Anexo');
            h += '<div class="hanexo-box">' +
              '<div class="hanexo-info">' +
              '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="m21.44 11.05-9.19 9.19a6 6 0 0 1-8.49-8.49l8.57-8.57A4 4 0 1 1 18 8.84l-8.59 8.57a2 2 0 0 1-2.83-2.83l8.49-8.48"/></svg>' +
              '<span class="hanexo-name" title="' + escapeHtml(aNome) + '">' + escapeHtml(aNome) + '</span>' +
              '</div>' +
              (anx.caminho ? '<button type="button" class="btn-abrir-anexo" data-caminho="' + escapeHtml(anx.caminho) + '"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="width:12px;height:12px;"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/></svg> Abrir Anexo</button>' : '') +
              '</div>';
          });
          h += '</div>';
          return h;
        })() +
        '<div class="hdetail-label">O que mudou</div>' +
        '<div class="diff">' +
        '<div class="diff-col before">' +
        '<div class="diff-head"><span class="dot"></span>Antes</div>' +
        '<ul class="diff-list">' + diffObj.antesListHtml + '</ul>' +
        '</div>' +
        '<div class="diff-col after">' +
        '<div class="diff-head"><span class="dot"></span>Depois</div>' +
        '<ul class="diff-list">' + diffObj.depoisListHtml + '</ul>' +
        '</div>' +
        '</div>' +
        '</div>';

      // Event listener para expandir/recolher
      var row = item.querySelector('.hitem-row');
      row.addEventListener('click', function () {
        item.classList.toggle('open');
      });

      // Event listeners para abrir anexos na rede
      item.querySelectorAll('.btn-abrir-anexo').forEach(function (btnAbrir) {
        btnAbrir.addEventListener('click', function (e) {
          e.stopPropagation();
          var path = this.dataset.caminho;
          if (!path) return;
          if (isPyWebviewAvailable()) {
            window.pywebview.api.open_attachment(path).then(function (res) {
              if (res && res.status === 'error') {
                showToast(res.message || 'Erro ao abrir anexo.', true);
              }
            }).catch(function (err) {
              showToast('Erro ao abrir anexo: ' + (err ? (err.message || err) : 'Erro'), true);
            });
          } else {
            showToast('Abertura de anexo disponível no aplicativo desktop.', false);
          }
        });
      });

      histModalBody.appendChild(item);
    });
  }

  function formatCondSummary(m) {
    if (!m) return '';
    if (m.padrao) return 'Senão';
    if (!m.cond || !m.cond.length) return 'Sem condição';
    return m.cond.map(function (c, ci) {
      var fObj = CAMPOS_COND_BLOCO.filter(function (x) { return x.k === c.c; })[0];
      var cName = fObj ? fObj.n : c.c;
      var prefix = ci > 0 ? (c.j === 'OU' ? ' OU ' : ' E ') : 'SE ';
      return prefix + cName + ' ' + (c.o || '=') + ' ' + c.val;
    }).join('');
  }

  function gerarDiffListasHtml(antes, depois) {
    var aBase = (antes && antes.base) ? antes.base : {};
    var dBase = (depois && depois.base) ? depois.base : {};

    var aForma = (aBase.forma || 'Desconhecida').replace('_', ' ');
    var dForma = (dBase.forma || 'Desconhecida').replace('_', ' ');
    aForma = aForma.charAt(0).toUpperCase() + aForma.slice(1);
    dForma = dForma.charAt(0).toUpperCase() + dForma.slice(1);

    var aVal = aBase.valor_base !== undefined ? aBase.valor_base : (aBase.valor !== undefined ? aBase.valor : 0);
    var dVal = dBase.valor_base !== undefined ? dBase.valor_base : (dBase.valor !== undefined ? dBase.valor : 0);

    var chForma = aForma !== dForma;
    var chVal = String(aVal) !== String(dVal);

    var antesLis = [];
    var depoisLis = [];

    // 1. Forma base
    if (chForma) {
      antesLis.push('<li class="changed">Forma base: <b class="num-diff num-diff-before">' + escapeHtml(aForma) + '</b></li>');
      depoisLis.push('<li class="changed">Forma base: <b class="num-diff num-diff-after">' + escapeHtml(dForma) + '</b></li>');
    } else {
      antesLis.push('<li>Forma base: <b>' + escapeHtml(aForma) + '</b></li>');
      depoisLis.push('<li>Forma base: <b>' + escapeHtml(dForma) + '</b></li>');
    }

    // 2. Valor base (se não for blocos)
    if (aBase.forma !== 'blocos' && dBase.forma !== 'blocos') {
      if (chVal) {
        antesLis.push('<li class="changed">Valor base: <b class="num-diff num-diff-before">' + escapeHtml(aVal) + '</b></li>');
        depoisLis.push('<li class="changed">Valor base: <b class="num-diff num-diff-after">' + escapeHtml(dVal) + '</b></li>');
      } else {
        antesLis.push('<li>Valor base: <b>' + escapeHtml(aVal) + '</b></li>');
        depoisLis.push('<li>Valor base: <b>' + escapeHtml(dVal) + '</b></li>');
      }
    }

    // 3. Passo (se aditiva)
    if (aBase.passo !== undefined || dBase.passo !== undefined) {
      var aPasso = aBase.passo !== undefined ? aBase.passo : '—';
      var dPasso = dBase.passo !== undefined ? dBase.passo : '—';
      var chPasso = String(aPasso) !== String(dPasso);
      if (chPasso) {
        antesLis.push('<li class="changed">Passo por módulo: <b class="num-diff num-diff-before">+' + escapeHtml(aPasso) + '</b></li>');
        depoisLis.push('<li class="changed">Passo por módulo: <b class="num-diff num-diff-after">+' + escapeHtml(dPasso) + '</b></li>');
      } else {
        antesLis.push('<li>Passo por módulo: <b>+' + escapeHtml(aPasso) + '</b></li>');
        depoisLis.push('<li>Passo por módulo: <b>+' + escapeHtml(dPasso) + '</b></li>');
      }
    }

    // 4. Tabela (1m–8m) se aplicável
    var aArr = (aBase.valores && Array.isArray(aBase.valores)) ? aBase.valores : null;
    var dArr = (dBase.valores && Array.isArray(dBase.valores)) ? dBase.valores : null;

    if (aArr || dArr) {
      var maxLen = Math.max(aArr ? aArr.length : 0, dArr ? dArr.length : 0, 8);
      var antesTabParts = [];
      var depoisTabParts = [];
      var chTab = false;

      for (var i = 0; i < maxLen; i++) {
        var vA = aArr && aArr[i] !== undefined ? aArr[i] : '0';
        var vD = dArr && dArr[i] !== undefined ? dArr[i] : '0';

        if (String(vA) !== String(vD)) {
          chTab = true;
          antesTabParts.push('<b class="num-diff num-diff-before">' + escapeHtml(vA) + '</b>');
          depoisTabParts.push('<b class="num-diff num-diff-after">' + escapeHtml(vD) + '</b>');
        } else {
          antesTabParts.push(escapeHtml(vA));
          depoisTabParts.push(escapeHtml(vD));
        }
      }

      antesLis.push('<li class="' + (chTab ? 'changed' : '') + '">Tabela (1m–8m): ' + antesTabParts.join(', ') + '</li>');
      depoisLis.push('<li class="' + (chTab ? 'changed' : '') + '">Tabela (1m–8m): ' + depoisTabParts.join(', ') + '</li>');
    }

    // 5. Etapas Derivadas (se existirem em DUR)
    if ((aBase.etapas && aBase.etapas.length) || (dBase.etapas && dBase.etapas.length)) {
      var aEt = (aBase.etapas || []).map(function (e) { return e.tipo + '(' + (e.valor !== undefined ? e.valor : (e.modo || '')) + ')'; }).join(', ') || 'Nenhuma';
      var dEt = (dBase.etapas || []).map(function (e) { return e.tipo + '(' + (e.valor !== undefined ? e.valor : (e.modo || '')) + ')'; }).join(', ') || 'Nenhuma';
      var chEt = aEt !== dEt;
      if (chEt) {
        antesLis.push('<li class="changed">Etapas derivadas: <b class="num-diff num-diff-before">' + escapeHtml(aEt) + '</b></li>');
        depoisLis.push('<li class="changed">Etapas derivadas: <b class="num-diff num-diff-after">' + escapeHtml(dEt) + '</b></li>');
      } else {
        antesLis.push('<li>Etapas derivadas: <b>' + escapeHtml(aEt) + '</b></li>');
        depoisLis.push('<li>Etapas derivadas: <b>' + escapeHtml(dEt) + '</b></li>');
      }
    }

    // 6. Blocos & Montagens
    var aBlocos = (aBase.blocos && Array.isArray(aBase.blocos)) ? aBase.blocos : [];
    var dBlocos = (dBase.blocos && Array.isArray(dBase.blocos)) ? dBase.blocos : [];
    var aMont = (aBase.montagens && Array.isArray(aBase.montagens)) ? aBase.montagens : [];
    var dMont = (dBase.montagens && Array.isArray(dBase.montagens)) ? dBase.montagens : [];

    if (aBase.forma === 'blocos' || dBase.forma === 'blocos' || aBlocos.length || dBlocos.length || aMont.length || dMont.length) {
      if (aBlocos.length > 0 || dBlocos.length > 0) {
        var maxB = Math.max(aBlocos.length, dBlocos.length);
        for (var bi = 0; bi < maxB; bi++) {
          var bA = aBlocos[bi];
          var bD = dBlocos[bi];
          var nameA = bA ? bA.nome : (bD ? bD.nome : 'Bloco ' + (bi + 1));
          var exprA = bA ? chainExpr(bA.it, false, null, aBlocos) : '—';
          var exprD = bD ? chainExpr(bD.it, false, null, dBlocos) : '—';

          var chB = exprA !== exprD || (bA && bD && bA.nome !== bD.nome);
          if (chB) {
            antesLis.push('<li class="changed">Bloco "' + escapeHtml(nameA) + '": <b class="num-diff num-diff-before">' + escapeHtml(exprA) + '</b></li>');
            depoisLis.push('<li class="changed">Bloco "' + escapeHtml(nameA) + '": <b class="num-diff num-diff-after">' + escapeHtml(exprD) + '</b></li>');
          } else {
            antesLis.push('<li>Bloco "' + escapeHtml(nameA) + '": <b>' + escapeHtml(exprA) + '</b></li>');
            depoisLis.push('<li>Bloco "' + escapeHtml(nameA) + '": <b>' + escapeHtml(exprD) + '</b></li>');
          }
        }
      }

      if (aMont.length > 0 || dMont.length > 0) {
        var maxM = Math.max(aMont.length, dMont.length);
        for (var mi = 0; mi < maxM; mi++) {
          var mA = aMont[mi];
          var mD = dMont[mi];

          var mNameA = mA ? mA.nome : (mD ? mD.nome : 'Montagem ' + (mi + 1));
          var cA = mA ? formatCondSummary(mA) : '';
          var cD = mD ? formatCondSummary(mD) : '';
          var eA = mA ? chainExpr(mA.it, false, null, aBlocos) : '';
          var eD = mD ? chainExpr(mD.it, false, null, dBlocos) : '';

          var fullA = mA ? (mNameA + ' [' + cA + '] = ' + eA) : '—';
          var fullD = mD ? (mNameA + ' [' + cD + '] = ' + eD) : '—';

          var chM = fullA !== fullD;
          if (chM) {
            antesLis.push('<li class="changed">Montagem "' + escapeHtml(mNameA) + '": <b class="num-diff num-diff-before">' + escapeHtml(fullA) + '</b></li>');
            depoisLis.push('<li class="changed">Montagem "' + escapeHtml(mNameA) + '": <b class="num-diff num-diff-after">' + escapeHtml(fullD) + '</b></li>');
          } else {
            antesLis.push('<li>Montagem "' + escapeHtml(mNameA) + '": <b>' + escapeHtml(fullA) + '</b></li>');
            depoisLis.push('<li>Montagem "' + escapeHtml(mNameA) + '": <b>' + escapeHtml(fullD) + '</b></li>');
          }
        }
      }
    }

    // 7. Condições Adicionais
    var aCondsArr = (antes && antes.condicoes) ? antes.condicoes : [];
    var dCondsArr = (depois && depois.condicoes) ? depois.condicoes : [];

    if (aCondsArr.length > 0 || dCondsArr.length > 0) {
      var aCondMap = {};
      aCondsArr.forEach(function (c) { aCondMap[c.flag] = c; });
      var dCondMap = {};
      dCondsArr.forEach(function (c) { dCondMap[c.flag] = c; });

      var allFlags = {};
      aCondsArr.forEach(function (c) { allFlags[c.flag] = true; });
      dCondsArr.forEach(function (c) { allFlags[c.flag] = true; });

      var aCondParts = [];
      var dCondParts = [];
      var chConds = false;

      Object.keys(allFlags).forEach(function (flag) {
        var cA = aCondMap[flag];
        var cD = dCondMap[flag];

        var rotA = cA ? (cA.rotulo || cA.flag) : null;
        var rotD = cD ? (cD.rotulo || cD.flag) : null;
        var valA = cA ? (cA.valor !== undefined ? cA.valor : '') : null;
        var valD = cD ? (cD.valor !== undefined ? cD.valor : '') : null;

        if (String(valA) !== String(valD)) {
          chConds = true;
          if (cA) aCondParts.push(escapeHtml(rotA) + ': <b class="num-diff num-diff-before">' + escapeHtml(valA) + '</b>');
          if (cD) dCondParts.push(escapeHtml(rotD) + ': <b class="num-diff num-diff-after">' + escapeHtml(valD) + '</b>');
        } else {
          if (cA) aCondParts.push(escapeHtml(rotA) + ': <b>' + escapeHtml(valA) + '</b>');
          if (cD) dCondParts.push(escapeHtml(rotD) + ': <b>' + escapeHtml(valD) + '</b>');
        }
      });

      var aCondStr = aCondParts.length ? aCondParts.join('; ') : 'Nenhuma';
      var dCondStr = dCondParts.length ? dCondParts.join('; ') : 'Nenhuma';

      antesLis.push('<li class="' + (chConds ? 'changed' : '') + '">Condições adicionais: ' + aCondStr + '</li>');
      depoisLis.push('<li class="' + (chConds ? 'changed' : '') + '">Condições adicionais: ' + dCondStr + '</li>');
    } else if (aBase.forma !== 'blocos' && dBase.forma !== 'blocos') {
      antesLis.push('<li>Condições: <b>Nenhuma</b></li>');
      depoisLis.push('<li>Condições: <b>Nenhuma</b></li>');
    }

    // 8. Regras Especiais: Container Solar & ESSW
    var aEsp = getEspeciais(antes);
    var dEsp = getEspeciais(depois);

    // 8.1 Container Solar
    var aCs = aEsp ? aEsp.solar : { ativo: false };
    var dCs = dEsp ? dEsp.solar : { ativo: false };
    if (aCs.ativo || dCs.ativo) {
      var aCsVal = (aCs.base && (aCs.base.valor !== undefined ? aCs.base.valor : aCs.base.valor_base)) !== undefined ? (aCs.base.valor !== undefined ? aCs.base.valor : aCs.base.valor_base) : 0;
      var dCsVal = (dCs.base && (dCs.base.valor !== undefined ? dCs.base.valor : dCs.base.valor_base)) !== undefined ? (dCs.base.valor !== undefined ? dCs.base.valor : dCs.base.valor_base) : 0;
      var aCsConds = (aCs.condicoes || []).map(function (c) { return (c.rotulo || flagNome(c.flag)) + ' (' + (c.forma || 'fixo') + ': ' + (c.valor !== undefined ? c.valor : '') + ')'; });
      var dCsConds = (dCs.condicoes || []).map(function (c) { return (c.rotulo || flagNome(c.flag)) + ' (' + (c.forma || 'fixo') + ': ' + (c.valor !== undefined ? c.valor : '') + ')'; });

      var aCsHerdar = aCs.herdar_condicoes !== false;
      var dCsHerdar = dCs.herdar_condicoes !== false;
      var aCsHerdarTxt = aCsHerdar ? 'Aplicar cond. padrão: Sim' : 'Aplicar cond. padrão: Não';
      var dCsHerdarTxt = dCsHerdar ? 'Aplicar cond. padrão: Sim' : 'Aplicar cond. padrão: Não';

      var aCsParts = [];
      if (aCs.ativo) {
        aCsParts.push('Base: ' + aCsVal);
        aCsParts.push(aCsHerdarTxt);
        if (aCsConds.length) aCsParts.push('Condições específicas: ' + aCsConds.join(', '));
      }
      var dCsParts = [];
      if (dCs.ativo) {
        dCsParts.push('Base: ' + dCsVal);
        dCsParts.push(dCsHerdarTxt);
        if (dCsConds.length) dCsParts.push('Condições específicas: ' + dCsConds.join(', '));
      }

      var aCsSummary = aCs.ativo ? aCsParts.join(' | ') : 'Padrão (Desativado)';
      var dCsSummary = dCs.ativo ? dCsParts.join(' | ') : 'Padrão (Desativado)';

      var chCs = (aCs.ativo !== dCs.ativo) || (String(aCsVal) !== String(dCsVal)) || (aCsConds.join('|') !== dCsConds.join('|')) || (aCsHerdar !== dCsHerdar);

      if (chCs) {
        antesLis.push('<li class="changed">Container Solar: <b class="num-diff num-diff-before">' + escapeHtml(aCsSummary) + '</b></li>');
        depoisLis.push('<li class="changed">Container Solar: <b class="num-diff num-diff-after">' + escapeHtml(dCsSummary) + '</b></li>');
      } else {
        antesLis.push('<li>Container Solar: <b>' + escapeHtml(aCsSummary) + '</b></li>');
        depoisLis.push('<li>Container Solar: <b>' + escapeHtml(dCsSummary) + '</b></li>');
      }
    }

    // 8.2 ESSW
    var aEssw = aEsp ? aEsp.essw : { ativo: false };
    var dEssw = dEsp ? dEsp.essw : { ativo: false };
    if (aEssw.ativo || dEssw.ativo) {
      var aEsswVal = (aEssw.base && (aEssw.base.valor !== undefined ? aEssw.base.valor : aEssw.base.valor_base)) !== undefined ? (aEssw.base.valor !== undefined ? aEssw.base.valor : aEssw.base.valor_base) : 0;
      var dEsswVal = (dEssw.base && (dEssw.base.valor !== undefined ? dEssw.base.valor : dEssw.base.valor_base)) !== undefined ? (dEssw.base.valor !== undefined ? dEssw.base.valor : dEssw.base.valor_base) : 0;
      var aEsswConds = (aEssw.condicoes || []).map(function (c) { return (c.rotulo || flagNome(c.flag)) + ' (' + (c.forma || 'fixo') + ': ' + (c.valor !== undefined ? c.valor : '') + ')'; });
      var dEsswConds = (dEssw.condicoes || []).map(function (c) { return (c.rotulo || flagNome(c.flag)) + ' (' + (c.forma || 'fixo') + ': ' + (c.valor !== undefined ? c.valor : '') + ')'; });

      var aEsswHerdar = aEssw.herdar_condicoes !== false;
      var dEsswHerdar = dEssw.herdar_condicoes !== false;
      var aEsswHerdarTxt = aEsswHerdar ? 'Aplicar cond. padrão: Sim' : 'Aplicar cond. padrão: Não';
      var dEsswHerdarTxt = dEsswHerdar ? 'Aplicar cond. padrão: Sim' : 'Aplicar cond. padrão: Não';

      var aEsswParts = [];
      if (aEssw.ativo) {
        aEsswParts.push('Base: ' + aEsswVal);
        aEsswParts.push(aEsswHerdarTxt);
        if (aEsswConds.length) aEsswParts.push('Condições específicas: ' + aEsswConds.join(', '));
      }
      var dEsswParts = [];
      if (dEssw.ativo) {
        dEsswParts.push('Base: ' + dEsswVal);
        dEsswParts.push(dEsswHerdarTxt);
        if (dEsswConds.length) dEsswParts.push('Condições específicas: ' + dEsswConds.join(', '));
      }

      var aEsswSummary = aEssw.ativo ? aEsswParts.join(' | ') : 'Padrão (Desativado)';
      var dEsswSummary = dEssw.ativo ? dEsswParts.join(' | ') : 'Padrão (Desativado)';

      var chEssw = (aEssw.ativo !== dEssw.ativo) || (String(aEsswVal) !== String(dEsswVal)) || (aEsswConds.join('|') !== dEsswConds.join('|')) || (aEsswHerdar !== dEsswHerdar);

      if (chEssw) {
        antesLis.push('<li class="changed">ESSW: <b class="num-diff num-diff-before">' + escapeHtml(aEsswSummary) + '</b></li>');
        depoisLis.push('<li class="changed">ESSW: <b class="num-diff num-diff-after">' + escapeHtml(dEsswSummary) + '</b></li>');
      } else {
        antesLis.push('<li>ESSW: <b>' + escapeHtml(aEsswSummary) + '</b></li>');
        depoisLis.push('<li>ESSW: <b>' + escapeHtml(dEsswSummary) + '</b></li>');
      }
    }

    return {
      antesListHtml: antesLis.join(''),
      depoisListHtml: depoisLis.join('')
    };
  }

  function formatRuleSummary(rule) {
    if (!rule || typeof rule !== 'object') return '<em>Sem dados anteriores (Regra nova)</em>';

    var items = [];
    var base = rule.base || {};

    if (base.forma) {
      var formaNome = base.forma.toUpperCase().replace('_', ' ');
      items.push('<b>Forma Base:</b> ' + formaNome);
    }
    if (base.forma !== 'blocos' && (base.valor_base !== undefined || base.valor !== undefined)) {
      var v = base.valor_base !== undefined ? base.valor_base : base.valor;
      items.push('<b>Valor Base:</b> <code style="font-family:IBM Plex Mono; font-weight:700;">' + v + '</code>');
    }
    if (base.passo !== undefined) {
      items.push('<b>Passo por Módulo:</b> +' + base.passo);
    }
    if (base.escala !== undefined) {
      items.push('<b>Fator Escala:</b> ' + base.escala + '×');
    }
    if (base.fator !== undefined) {
      items.push('<b>Fator Multiplicativo:</b> ' + base.fator + '×');
    }

    if (base.etapas && Array.isArray(base.etapas) && base.etapas.length > 0) {
      var etapasStr = base.etapas.map(function (et, i) {
        var val = et.valor !== undefined ? et.valor : (et.modo || '');
        return (i + 1) + '. ' + et.tipo.toUpperCase() + ' (' + val + ')';
      }).join(', ');
      items.push('<b>Etapas Derivadas (DUR):</b> ' + etapasStr);
    }

    if (base.valores && Array.isArray(base.valores)) {
      items.push('<b>Valores Tabela (1m a 8m):</b> [' + base.valores.join(', ') + ']');
    }

    if (base.forma === 'blocos') {
      if (base.blocos && Array.isArray(base.blocos) && base.blocos.length > 0) {
        var blocosStr = base.blocos.map(function (b) {
          return '<li><b>' + escapeHtml(b.nome) + ':</b> ' + escapeHtml(chainExpr(b.it, false, null, base.blocos)) + '</li>';
        }).join('');
        items.push('<b>Blocos (' + base.blocos.length + '):</b><ul style="margin:4px 0 0 0;">' + blocosStr + '</ul>');
      }
      if (base.montagens && Array.isArray(base.montagens) && base.montagens.length > 0) {
        var montStr = base.montagens.map(function (m) {
          return '<li><b>' + escapeHtml(m.nome) + '</b> <i>(' + escapeHtml(formatCondSummary(m)) + ')</i>: <code>' + escapeHtml(chainExpr(m.it, false, null, base.blocos)) + '</code></li>';
        }).join('');
        items.push('<b>Montagens (' + base.montagens.length + '):</b><ul style="margin:4px 0 0 0;">' + montStr + '</ul>');
      }
    }

    var conds = rule.condicoes || [];
    if (conds.length > 0) {
      var condsStr = conds.map(function (c) {
        var rot = c.rotulo || c.flag;
        var val = c.valor !== undefined ? c.valor : '';
        return '<li>' + escapeHtml(rot) + ' (' + c.forma + '): <b>' + val + '</b></li>';
      }).join('');
      items.push('<b>Condições Adicionais (' + conds.length + '):</b><ul style="margin:4px 0 0 0;">' + condsStr + '</ul>');
    } else if (base.forma !== 'blocos') {
      items.push('<b>Condições:</b> <em>Nenhuma condição adicional</em>');
    }

    var esp = getEspeciais(rule);
    if (esp.solar && esp.solar.ativo) {
      var sVal = (esp.solar.base && (esp.solar.base.valor !== undefined ? esp.solar.base.valor : esp.solar.base.valor_base)) || 0;
      var sConds = (esp.solar.condicoes || []).map(function (c) { return (c.rotulo || flagNome(c.flag)) + ' (' + (c.forma || 'fixo') + ': ' + (c.valor !== undefined ? c.valor : '') + ')'; });
      var sHerdar = esp.solar.herdar_condicoes !== false ? 'Sim' : 'Não';
      items.push('<b>Container Solar:</b> Base: ' + sVal + ' | Aplicar cond. padrão: ' + sHerdar + (sConds.length ? ' | Cond. específicas: ' + sConds.join(', ') : ''));
    }
    if (esp.essw && esp.essw.ativo) {
      var eVal = (esp.essw.base && (esp.essw.base.valor !== undefined ? esp.essw.base.valor : esp.essw.base.valor_base)) || 0;
      var eConds = (esp.essw.condicoes || []).map(function (c) { return (c.rotulo || flagNome(c.flag)) + ' (' + (c.forma || 'fixo') + ': ' + (c.valor !== undefined ? c.valor : '') + ')'; });
      var eHerdar = esp.essw.herdar_condicoes !== false ? 'Sim' : 'Não';
      items.push('<b>ESSW:</b> Base: ' + eVal + ' | Aplicar cond. padrão: ' + eHerdar + (eConds.length ? ' | Cond. específicas: ' + eConds.join(', ') : ''));
    }

    return items.map(function (it) { return '<div style="margin-bottom:3px;">• ' + it + '</div>'; }).join('');
  }

  function escapeHtml(str) {
    if (str === null || str === undefined) return '';
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  function renderHistoricoFallback() {
    if (!histModalBody) return;
    if (!state.historico || !state.historico.length) {
      histModalBody.innerHTML = '<div class="hist-empty">Nenhuma alteração registrada ainda.</div>';
      return;
    }

    histModalBody.innerHTML = state.historico.map(function (h) {
      return '<div class="hist-card" style="padding:12px;">' +
        '<div style="font-weight:700; font-size:13px;">' + escapeHtml(h.campo) + '</div>' +
        '<div style="font-size:12px;">' + escapeHtml(h.resumo) + '</div>' +
        '<div style="font-size:11px; color:var(--text-faint); margin-top:4px;">Data: ' + escapeHtml(h.quando) + '</div>' +
        '</div>';
    }).join('');
  }

  /* Settings Modal Setup */
  var settingsOverlay = $('settingsOverlay');
  var jsonEditor = $('jsonEditor');
  var jsonStatus = $('jsonStatus');
  var jsonStatusText = $('jsonStatusText');
  var cfgKeySelect = $('cfgKeySelect');
  var cfgItemsList = $('cfgItemsList');
  var cfgNewItemInput = $('cfgNewItemInput');
  var btnAddCfgItem = $('btnAddCfgItem');

  var LIST_NAMES = {
    tipoestrutura: 'Tipo de Estrutura',
    planpin: 'Plano de Pintura',
    tipomaq: 'Tipo Máq. Ar Condicionado',
    incendio: 'Tipo Sist. Incêndio',
    seguranca: 'Tipo Sist. Segurança',
    complexidade: 'Complexidade',
    planejadorSel: 'Planejadores'
  };

  function renderCfgItems() {
    if (!cfgKeySelect || !cfgItemsList) return;
    var selectedKey = cfgKeySelect.value || 'tipoestrutura';
    var items = CONFIG.listas[selectedKey] || [];

    if ($('cfgListTitle')) $('cfgListTitle').textContent = 'Itens de "' + (LIST_NAMES[selectedKey] || selectedKey) + '"';
    if ($('cfgItemCount')) $('cfgItemCount').textContent = items.length + ' item' + (items.length === 1 ? '' : 's');

    cfgItemsList.innerHTML = '';
    if (!items.length) {
      cfgItemsList.innerHTML = '<div style="color:var(--text-faint); font-size:12px; padding:12px; text-align:center;">Nenhum item nesta lista. Adicione um abaixo.</div>';
      return;
    }

    items.forEach(function (itemText, idx) {
      var row = document.createElement('div');
      row.className = 'cfg-item-row';
      row.style.cssText = 'display:flex; align-items:center; gap:6px; background:var(--panel-2); border:1px solid var(--border); border-radius:8px; padding:4px 6px;';
      row.innerHTML =
        '<span style="font-family:\'IBM Plex Mono\'; font-size:11px; color:var(--text-faint); min-width:20px; text-align:center;">' + (idx + 1) + '</span>' +
        '<input type="text" class="ipt mono" value="' + itemText.replace(/"/g, '&quot;') + '" style="flex:1; border:none; background:transparent; padding:4px 6px;">' +
        '<button type="button" class="icon-btn btn-up" title="Mover para cima" style="width:26px; height:26px; font-size:11px;" ' + (idx === 0 ? 'disabled' : '') + '>↑</button>' +
        '<button type="button" class="icon-btn btn-down" title="Mover para baixo" style="width:26px; height:26px; font-size:11px;" ' + (idx === items.length - 1 ? 'disabled' : '') + '>↓</button>' +
        '<button type="button" class="icon-btn btn-del" title="Remover item" style="width:26px; height:26px; color:var(--red); border-color:color-mix(in srgb, var(--red) 30%, transparent);">✕</button>';

      var inputEl = row.querySelector('input');
      inputEl.addEventListener('input', function () {
        CONFIG.listas[selectedKey][idx] = inputEl.value;
        syncJsonEditorFromConfig();
      });

      row.querySelector('.btn-up').addEventListener('click', function () {
        if (idx > 0) {
          var temp = CONFIG.listas[selectedKey][idx - 1];
          CONFIG.listas[selectedKey][idx - 1] = CONFIG.listas[selectedKey][idx];
          CONFIG.listas[selectedKey][idx] = temp;
          syncJsonEditorFromConfig();
          renderCfgItems();
        }
      });

      row.querySelector('.btn-down').addEventListener('click', function () {
        if (idx < items.length - 1) {
          var temp = CONFIG.listas[selectedKey][idx + 1];
          CONFIG.listas[selectedKey][idx + 1] = CONFIG.listas[selectedKey][idx];
          CONFIG.listas[selectedKey][idx] = temp;
          syncJsonEditorFromConfig();
          renderCfgItems();
        }
      });

      row.querySelector('.btn-del').addEventListener('click', function () {
        CONFIG.listas[selectedKey].splice(idx, 1);
        syncJsonEditorFromConfig();
        renderCfgItems();
      });

      cfgItemsList.appendChild(row);
    });
  }

  function addItemToSelectedList() {
    if (!cfgNewItemInput || !cfgKeySelect) return;
    var newTxt = cfgNewItemInput.value.trim();
    if (!newTxt) return;

    var selectedKey = cfgKeySelect.value || 'tipoestrutura';
    if (!CONFIG.listas[selectedKey]) CONFIG.listas[selectedKey] = [];

    CONFIG.listas[selectedKey].push(newTxt);
    cfgNewItemInput.value = '';
    syncJsonEditorFromConfig();
    renderCfgItems();
  }

  if (btnAddCfgItem) btnAddCfgItem.addEventListener('click', addItemToSelectedList);
  if (cfgNewItemInput) {
    cfgNewItemInput.addEventListener('keydown', function (e) {
      if (e.key === 'Enter') {
        e.preventDefault();
        addItemToSelectedList();
      }
    });
  }
  if (cfgKeySelect) cfgKeySelect.addEventListener('change', renderCfgItems);

  function validateJson() {
    if (!jsonEditor) return null;
    try {
      var parsed = JSON.parse(jsonEditor.value);
      if (typeof parsed !== 'object' || parsed === null) throw new Error('O JSON deve ser um objeto válido');
      if (jsonStatus) {
        jsonStatus.classList.remove('err');
        jsonStatusText.textContent = 'JSON de Configuração válido';
      }
      return parsed;
    } catch (e) {
      if (jsonStatus) {
        jsonStatus.classList.add('err');
        jsonStatusText.textContent = 'Erro: ' + e.message;
      }
      return null;
    }
  }

  function syncJsonEditorFromConfig() {
    if (jsonEditor) {
      jsonEditor.value = JSON.stringify(CONFIG, null, 2);
      validateJson();
    }
  }

  var regrasJsonEditor = $('regrasJsonEditor');
  var regrasJsonStatus = $('regrasJsonStatus');
  var regrasJsonStatusText = $('regrasJsonStatusText');

  function validateRegrasJson() {
    if (!regrasJsonEditor) return null;
    try {
      var parsed = JSON.parse(regrasJsonEditor.value);
      if (!Array.isArray(parsed)) throw new Error('regras.json deve ser uma lista (array) de áreas');
      if (regrasJsonStatus) {
        regrasJsonStatus.classList.remove('err');
        regrasJsonStatusText.textContent = 'JSON de Regras válido (' + parsed.length + ' área' + (parsed.length === 1 ? '' : 's') + ')';
      }
      return parsed;
    } catch (e) {
      if (regrasJsonStatus) {
        regrasJsonStatus.classList.add('err');
        regrasJsonStatusText.textContent = 'Erro: ' + e.message;
      }
      return null;
    }
  }

  function openSettings() {
    if (!state.isMaintenanceUnlocked) {
      showToast('Apenas usuários autenticados no modo mantenedor podem alterar as configurações.', true);
      openAuthModal();
      return;
    }
    renderCfgItems();
    if (jsonEditor) {
      jsonEditor.value = JSON.stringify(CONFIG, null, 2);
      validateJson();
    }

    if (regrasJsonEditor) {
      regrasJsonEditor.value = JSON.stringify(state.regrasData, null, 2);
      validateRegrasJson();
    }
    settingsOverlay.classList.add('open');
  }

  function closeSettings() { settingsOverlay.classList.remove('open'); }

  if ($('btnSettings')) $('btnSettings').addEventListener('click', openSettings);
  if ($('btnCloseSettings')) $('btnCloseSettings').addEventListener('click', closeSettings);
  if (settingsOverlay) settingsOverlay.addEventListener('click', function (e) { if (e.target === settingsOverlay) closeSettings(); });

  if ($('btnRestoreDefault')) {
    $('btnRestoreDefault').addEventListener('click', function () {
      if (confirm('Deseja restaurar as configurações padrão?')) {
        CONFIG = JSON.parse(JSON.stringify(DEFAULT_CONFIG));
        syncJsonEditorFromConfig();
        renderCfgItems();
        buildAllSelectsFromConfig();
        recomputeForm();
        showToast('Configurações restauradas para o padrão!');
      }
    });
  }

  if ($('btnDownloadJson')) {
    $('btnDownloadJson').addEventListener('click', function () {
      var dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(CONFIG, null, 2));
      var downloadAnchor = document.createElement('a');
      downloadAnchor.setAttribute("href", dataStr);
      downloadAnchor.setAttribute("download", "config_eletrocentros.json");
      document.body.appendChild(downloadAnchor);
      downloadAnchor.click();
      downloadAnchor.remove();
    });
  }

  document.querySelectorAll('.modal-tab').forEach(function (tab) {
    tab.addEventListener('click', function () {
      document.querySelectorAll('.modal-tab').forEach(function (t) { t.classList.remove('active'); });
      document.querySelectorAll('.modal-pane').forEach(function (p) { p.classList.remove('active'); });
      tab.classList.add('active');
      var targetPane = $(tab.dataset.pane);
      if (targetPane) targetPane.classList.add('active');
    });
  });

  if (jsonEditor) jsonEditor.addEventListener('input', validateJson);
  if (regrasJsonEditor) regrasJsonEditor.addEventListener('input', validateRegrasJson);

  if ($('btnApplyConfig')) {
    $('btnApplyConfig').addEventListener('click', function () {
      var parsedConfig = validateJson();
      if (!parsedConfig) return;

      var parsedRegras = validateRegrasJson();

      CONFIG = parsedConfig;
      buildAllSelectsFromConfig();
      recomputeForm();

      if (parsedRegras) {
        state.regrasData = parsedRegras;
        renderRegrasAreasNav();
      }

      closeSettings();

      if (isPyWebviewAvailable()) {
        window.pywebview.api.save_config(parsedConfig).then(function (res) {
          if (res && res.status === 'conflito') {
            showToast('⚠️ Conflito ao salvar configurações: ' + (res.message || 'Outro usuário salvou antes.'), true);
            alert('⚠️ Conflito de Configuração:\n\n' + (res.message || 'Outro usuário salvou alterações antes.') + '\n\nAs configurações mais recentes serão recarregadas do servidor.');
            loadExternalConfig();
          } else if (res && res.status === 'error') {
            showToast('Erro ao salvar configurações: ' + (res.message || 'Erro'), true);
          } else if (!parsedRegras) {
            showToast('Configurações salvas com sucesso no servidor!');
          }
        }).catch(function () { });

        if (parsedRegras) {
          window.pywebview.api.save_regras(parsedRegras).then(function (res) {
            if (res && res.status === 'conflito') {
              showToast('⚠️ Conflito ao salvar regras: ' + (res.message || 'Outro usuário salvou antes.'), true);
              alert('⚠️ Conflito de Regras:\n\n' + (res.message || 'Outro usuário salvou alterações antes.') + '\n\nAs regras mais recentes serão recarregadas do servidor.');
              carregarDisciplinas();
            } else if (res && res.status === 'error') {
              showToast('Erro ao salvar regras: ' + (res.message || 'Erro'), true);
            } else {
              showToast('Configurações e regras salvas com sucesso no servidor!');
            }
          }).catch(function () {
            showToast('Configurações aplicadas.');
          });
        }
      } else {
        showToast('Configurações e Regras aplicadas!');
      }
    });
  }

  /* Toast Helper */
  function showToast(msg, isError) {
    var t = $('toast');
    if (!t) return;
    $('toastText').textContent = msg;
    t.style.borderColor = isError ? 'var(--red)' : '';
    t.classList.add('show');
    clearTimeout(t._timer);
    t._timer = setTimeout(function () { t.classList.remove('show'); }, 3800);
  }

  // Load external configuration, rules, and selector on startup
  loadExternalConfig();
  carregarDisciplinas();
  carregarSeletor();
  carregarTemplateBlocks();

})();