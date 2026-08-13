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
    historico: []
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
    try { localStorage.setItem('eletrocentros_theme', theme); } catch (e) {}
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
  var tabBtnManutencao = $('tabBtnManutencao');
  var viewPlanejamento = $('view-planejamento');
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
    tabBtnPlanejamento.classList.toggle('active', targetView === 'planejamento');
    tabBtnManutencao.classList.toggle('active', targetView === 'manutencao');

    if (targetView === 'planejamento') {
      viewPlanejamento.classList.remove('hidden');
      viewManutencao.classList.add('hidden');
      if (modeChip) modeChip.style.display = 'none';
      btnHist.style.display = 'none';
      ringWrap.style.display = 'flex';
      subTitleText.textContent = 'PCP & PLANEJAMENTO';
    } else {
      viewPlanejamento.classList.add('hidden');
      viewManutencao.classList.remove('hidden');
      if (modeChip) modeChip.style.display = 'none';
      btnHist.style.display = 'flex';
      ringWrap.style.display = 'none';
      subTitleText.textContent = 'REGRAS & PARÂMETROS';
      if (!state.disciplinas.length) {
        carregarDisciplinas();
      }
    }
  }

  tabBtnPlanejamento.addEventListener('click', function () { switchView('planejamento'); });
  tabBtnManutencao.addEventListener('click', function () { switchView('manutencao'); });

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
  }

  document.querySelectorAll('input, .acessorio').forEach(function (el) {
    el.addEventListener('input', recomputeForm);
    el.addEventListener('change', recomputeForm);
  });

  recomputeForm();

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

  /* Footer Actions */
  if ($('btnOk')) {
    $('btnOk').addEventListener('click', function () {
      recomputeForm();
      if (reqDone < reqTotal) {
        alert('Existem ' + (reqTotal - reqDone) + ' campo(s) obrigatório(s) pendente(s). Verifique os indicadores âmbar.');
      } else {
        showToast('Dados validados com sucesso! Pronto para integração Python.');
      }
    });
  }

  if ($('btnSair')) {
    $('btnSair').addEventListener('click', function () {
      if (confirm('Deseja realmente encerrar a sessão atual?')) {
        location.reload();
      }
    });
  }

  if ($('btnCarregar')) {
    $('btnCarregar').addEventListener('click', function () {
      showToast('Carregando última sessão salva...');
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

  function carregarDisciplinas() {
    var nav = $('sectionNavManutencao');
    if (!nav) return;
    nav.innerHTML = '<span style="color:var(--text-faint); font-size:12px; padding:8px 4px;">Carregando disciplinas...</span>';

    if (isPyWebviewAvailable()) {
      window.pywebview.api.get_disciplinas().then(function (lista) {
        renderDisciplinasNav(lista);
      }).catch(function (err) {
        showNavError(err.message);
      });
    } else {
      apiCall('/manutencao/disciplinas').then(function (lista) {
        renderDisciplinasNav(lista);
      }).catch(function (err) {
        // Fallback mock disciplinas if local API backend is not active yet
        var mockLista = [
          { disciplina: 'Mecânica — Estrutura', total_campos: 14 },
          { disciplina: 'Elétrica & Equipamentos', total_campos: 18 },
          { disciplina: 'Acessórios & Adicionais', total_campos: 9 }
        ];
        renderDisciplinasNav(mockLista);
      });
    }
  }

  function showNavError(msg) {
    var nav = $('sectionNavManutencao');
    nav.innerHTML = '<span style="color:var(--red); font-size:12px; padding:8px 4px;">Não foi possível carregar as disciplinas. ' + msg + '</span>';
  }

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
    }).then(function (data) {
      if (data && data.length) {
        state.regrasData = data;
        renderRegrasAreasNav();
      }
    }).catch(function (err) {
      console.log('[Regras] Erro ao carregar regras.json fallback.');
    });
  }

  function renderRegrasAreasNav() {
    var nav = $('sectionNavManutencao');
    if (!nav) return;
    nav.innerHTML = '';
    state.regrasData.forEach(function (areaObj, idx) {
      var btn = document.createElement('button');
      btn.className = 'snav-btn' + (idx === state.selectedAreaIdx ? ' active' : '');
      var countCampos = Object.keys(areaObj.campos || {}).length;
      btn.innerHTML = areaObj.area + ' <span class="n">' + countCampos + '</span>';
      btn.addEventListener('click', function () {
        document.querySelectorAll('#sectionNavManutencao .snav-btn').forEach(function (b) { b.classList.remove('active'); });
        btn.classList.add('active');
        selecionarArea(idx);
      });
      nav.appendChild(btn);
    });
    if (state.regrasData.length) selecionarArea(state.selectedAreaIdx || 0);
  }

  function selecionarArea(areaIdx) {
    state.selectedAreaIdx = areaIdx;
    var areaObj = state.regrasData[areaIdx];
    var camposKeys = Object.keys(areaObj.campos || {});
    state.selectedCampoKey = camposKeys.length ? camposKeys[0] : null;
    state.selectedSubTab = 'H';
    prepararDirtySubTab();
    renderList();
    renderEditor();
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
        });
        str += parts.join(' ');
      }
    } else str = b.forma;
    var cCount = (ruleObj.condicoes || []).length;
    if (cCount > 0) str += ' (+ ' + cCount + ' cond)';
    return str;
  }

  function renderList() {
    var wrap = $('colList');
    if (!wrap) return;
    var areaObj = state.regrasData[state.selectedAreaIdx];
    var isMecanica = areaObj && (areaObj.area === 'MECÂNICA' || areaObj.area === 'Mecanica' || (areaObj.area && areaObj.area.toUpperCase().indexOf('MEC') >= 0));
    if ($('legendWrapper')) $('legendWrapper').style.display = isMecanica ? 'flex' : 'none';
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

  function valorBase(base, mod, campoObj, flagsAtivos, vH) {
    if (!base) return 0;
    var forma = base.forma;
    if (forma === 'blocos') {
      var simCtx = Object.assign({}, SIM_CTX, { nmod: mod });
      var hitM = matchedMontagem(base.montagens, simCtx);
      var bVal = hitM ? evalChain(hitM.it, simCtx, base.blocos) : 0;
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
    if (forma === 'aditiva') return (parseFloat(base.valor_base) || 0) + (parseFloat(base.passo) || 0) * (mod - 1);
    if (forma === 'tabela') {
      if (!base.valores || base.valores[mod - 1] === undefined) return 0;
      var hVal = (vH !== undefined) ? vH : 0;
      return evalExpr(base.valores[mod - 1], hVal);
    }
    if (forma === 'derivado_h') {
      var v = (vH !== undefined) ? vH : 0;
      if (vH === undefined && campoObj && campoObj.H && state.selectedSubTab !== 'H') {
        v = calcValor(campoObj.H, mod, flagsAtivos, campoObj);
      }

      var etapas = base.etapas;
      if (!etapas || !etapas.length) {
        etapas = [];
        if (base.divisao) etapas.push({ tipo: 'dividir', valor: base.divisao });
        if (base.arredondamento) etapas.push({ tipo: 'arredondar', modo: base.arredondamento });
        if (base.subtracao) etapas.push({ tipo: 'subtrair', valor: base.subtracao });
        if (base.soma) etapas.push({ tipo: 'somar', valor: base.soma });
      }

      etapas.forEach(function (step) {
        if (!step) return;
        var num = parseFloat(step.valor);
        if (isNaN(num)) num = 0;

        if (step.tipo === 'dividir') {
          if (num !== 0) v = v / num;
        } else if (step.tipo === 'multiplicar') {
          v = v * num;
        } else if (step.tipo === 'somar') {
          v = v + num;
        } else if (step.tipo === 'subtrair') {
          v = v - num;
        } else if (step.tipo === 'arredondar') {
          var modo = step.modo || step.arredondamento || 'cima';
          if (modo === 'cima') v = Math.ceil(v);
          else if (modo === 'baixo') v = Math.floor(v);
          else if (modo === 'padrao') v = Math.round(v);
        }
      });

      return v;
    }
    return 0;
  }

  function evalExpr(val, vH) {
    if (typeof val === 'number') return isNaN(val) ? 0 : val;
    if (!val || typeof val !== 'string') return 0;
    var str = val.trim();
    if (!str) return 0;

    var hVal = (vH !== undefined && !isNaN(vH)) ? vH : 0;

    var expr = str.replace(/\bH\b/gi, hVal).replace(/,/g, '.');

    expr = expr.replace(/arredondar\.para\.cima\s*\(\s*([^;)]+)(?:\s*;\s*[0-9]+)?\s*\)/gi, 'Math.ceil($1)');
    expr = expr.replace(/ceil\s*\(\s*([^)]+)\s*\)/gi, 'Math.ceil($1)');

    if (/\bH\b/i.test(str) && /\/\s*[0-9\.]+/i.test(expr) && !/Math\.ceil/i.test(expr)) {
      expr = 'Math.ceil(' + expr + ')';
    }

    try {
      var fn = new Function('Math', '"use strict"; return (' + expr + ');');
      var res = fn(Math);
      return typeof res === 'number' && !isNaN(res) ? res : 0;
    } catch (e) {
      var num = parseFloat(expr.replace(/[^0-9\.-]/g, ''));
      return isNaN(num) ? 0 : num;
    }
  }

  function valorBonus(b, mod, vAtual, vH) {
    var val = evalExpr(b.valor, vH);
    if (!vAtual) vAtual = 0;
    if (b.forma === 'tabela') {
      if (b.valores && b.valores[mod - 1] !== undefined) {
        return evalExpr(b.valores[mod - 1], vH);
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

  function calcValor(analise, mod, flagsAtivos, campoObj, vH) {
    if (!analise || !analise.base) return 0;
    var v = valorBase(analise.base, mod, campoObj, flagsAtivos, vH);
    (analise.condicoes || []).forEach(function (c) {
      if (flagsAtivos && flagsAtivos[c.flag]) v += valorBonus(c, mod, v, vH);
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

  var CAMPOS_COND_BLOCO = [
    { k: 'tipoestrutura', n: 'Tipo de estrutura', opts: ['Móvel', 'Semimóvel', 'Modular', 'Fixo', 'Embarcado', 'Container Solar'] },
    { k: 'trafo_oleo', n: 'Trafo a óleo', opts: ['Sim', 'Não'] },
    { k: 'nmod', n: 'Nº de módulos', num: true },
    { k: 'comp', n: 'Comprimento (m)', num: true },
    { k: 'larg', n: 'Largura (m)', num: true },
    { k: 'alt', n: 'Altura (m)', num: true },
    { k: 'tipomaq', n: 'Tipo de máquina', opts: ['Split', 'Wall Mounted', 'Roof Top', 'Não possui', 'Não aplicável'] },
    { k: 'incendio', n: 'Sistema de incêndio', opts: ['Com combate', 'Com instalações', 'Somente infra', 'Não aplicável'] }
  ];

  var SIM_CTX = { comp: 15, larg: 3, alt: 2.6, nmod: 1, tipoestrutura: 'Móvel', trafo_oleo: 'Não', tipomaq: 'Split', incendio: 'Não aplicável' };

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
    var list = getVARS();
    var f = list.filter(function (x) { return x.chave === k || x.k === k; })[0];
    if (!f) return (simCtx && simCtx[k] !== undefined) ? simCtx[k] : 0;
    if (f.tipo === 'entrada' || f.t === 'entrada') return (simCtx && simCtx[k] !== undefined) ? simCtx[k] : (f.valor !== undefined ? f.valor : f.v);
    return f.valor !== undefined ? f.valor : (f.v !== undefined ? f.v : 0);
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
    if (it.t === 'var') return '<span class="chip var"' + d + '>' + varSelectHtml(it.v) + '<button type="button" class="x" data-del="1"' + d + '>✕</button></span>';
    if (it.t === 'op') return '<span class="chip op"' + d + '>' + opSelectHtml(it.v) + '<button type="button" class="x" data-del="1"' + d + '>✕</button></span>';
    if (it.t === 'num') return '<span class="chip num"' + d + '><input value="' + String(it.v).replace('.', ',') + '"><button type="button" class="x" data-del="1"' + d + '>✕</button></span>';
    if (it.t === 'blk') return '<span class="chip blk"' + d + '>' + blkSelectHtml(it.v, skip, blocosList) + '<button type="button" class="x" data-del="1"' + d + '>✕</button></span>';
    return '';
  }

  function addBtnsHtml(sc, blk) {
    return '<button type="button" class="addbtn" data-add="var" data-sc="' + sc + '">+ variável</button>' +
      '<button type="button" class="addbtn" data-add="op" data-sc="' + sc + '">+ operação</button>' +
      '<button type="button" class="addbtn n" data-add="num" data-sc="' + sc + '">+ valor fixo</button>' +
      (blk ? '<button type="button" class="addbtn b" data-add="blk" data-sc="' + sc + '">+ bloco</button>' : '');
  }

  function getCurrentBase() {
    var r = state.dirtySubTabRule;
    return (r && r.base) ? r.base : null;
  }

  function wireBlocosEvents(subTabRule) {
    var base = getCurrentBase();
    if (!base || base.forma !== 'blocos') return;

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
        if (idx >= 0 && targetIdx >= 0 && targetIdx < (curBase.montagens.length - 1)) {
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
            renderEditor();
          }
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
        '<div class="chain">' + (bo.it || []).map(function (it, i) { return chipHtml(it, i, bo.id, bo.id, base.blocos); }).join('') + addBtnsHtml(bo.id, true) + '</div>' +
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
      '<select class="ipt step-tipo" style="padding:5px 8px; font-size:12px; width:130px; font-weight:500;">' +
      '<option value="dividir"' + (step.tipo === 'dividir' ? ' selected' : '') + '>Dividir (÷)</option>' +
      '<option value="multiplicar"' + (step.tipo === 'multiplicar' ? ' selected' : '') + '>Multiplicar (×)</option>' +
      '<option value="somar"' + (step.tipo === 'somar' ? ' selected' : '') + '>Somar (+)</option>' +
      '<option value="subtrair"' + (step.tipo === 'subtrair' ? ' selected' : '') + '>Subtrair (-)</option>' +
      '<option value="arredondar"' + (step.tipo === 'arredondar' ? ' selected' : '') + '>Arredondar</option>' +
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
      '<option value="blocos"' + (base.forma === 'blocos' ? ' selected' : '') + '>Blocos Condicionais (SE / Variáveis / Fórmulas)</option>' +
      '</select>' +
      '</div>';

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
          '<span class="mhead-cond-summary">' + escapeHtml(condSummary) + '</span>' +
          '<div style="margin-left:auto; display:flex; align-items:center; gap:8px;">' +
          (isHit ? '<span class="hitbadge">APLICADA</span>' : '') +
          '<span class="mres">= ' + (isFinite(rVal) ? rVal.toFixed(1).replace('.', ',') : '—') + ' h</span>' +
          (mi > 0 && !m.padrao ? '<button type="button" class="arrbtn" data-mv="up" data-m="' + m.id + '">▲</button>' : '') +
          (mi < base.montagens.length - 2 ? '<button type="button" class="arrbtn" data-mv="dn" data-m="' + m.id + '">▼</button>' : '') +
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

        html += '<div class="chain">' + (m.it || []).map(function (it, i) { return chipHtml(it, i, m.id, null, base.blocos); }).join('') + addBtnsHtml(m.id, true) + '</div>' +
          '<div class="expr-out">' + chainExpr(m.it, false, SIM_CTX, base.blocos) + '</div></div></div>';
      });

      html += '</div>' +
        '<button type="button" class="newblock" id="newMont"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4"><path d="M12 5v14M5 12h14"/></svg>Adicionar montagem condicional</button>' +
        '</div>';

      // 3. Painel de Simulação
      html += '<div class="block" style="margin-bottom:0">' +
        '<div class="bl">Simular <span class="hint">escolha o cenário e confira qual montagem é aplicada</span></div>' +
        '<div class="simbar">';

      ['tipoestrutura', 'trafo_oleo', 'comp', 'larg'].forEach(function (k) {
        var cObj = CAMPOS_COND_BLOCO.filter(function (x) { return x.k === k; })[0];
        if (cObj) {
          html += '<span class="siminp"><label>' + cObj.n + '</label>' + (cObj.opts
            ? '<select data-sim="' + cObj.k + '">' + cObj.opts.map(function (o) { return '<option' + (o === SIM_CTX[cObj.k] ? ' selected' : '') + '>' + o + '</option>'; }).join('') + '</select>'
            : '<input data-sim="' + cObj.k + '" value="' + SIM_CTX[cObj.k] + '">') + '</span>';
        } else {
          var vObj = getVARS().filter(function (y) { return y.chave === k || y.k === k; })[0];
          if (vObj) {
            html += '<span class="siminp"><label>' + (vObj.nome || vObj.n) + '</label><input data-sim="' + k + '" value="' + String(SIM_CTX[k]).replace('.', ',') + '"></span>';
          }
        }
      });

      html += '</div>';

      var totVal = hitM ? ev(hitM.it, SIM_CTX, base.blocos) : NaN;
      if (!hitM) {
        html += '<div class="err">Nenhuma montagem atende ao cenário e não há montagem padrão.</div>';
      } else if (isNaN(totVal)) {
        html += '<div class="err">Expressão inválida — verifique se falta uma operação entre os itens.</div>';
      } else {
        html += '<div class="simtot"><span class="lb">Montagem aplicada: <b>' + escapeHtml(hitM.nome) + '</b> · resultado de <b>' + state.selectedCampoKey + ' · ' + state.selectedSubTab + '</b></span>' +
          '<span class="vv">' + totVal.toFixed(1).replace('.', ',') + '</span><span class="un">horas</span></div>';
      }

      html += '</div>';
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

    html += '<div class="field-block"><div class="field-label">Condições adicionais (' + state.selectedSubTab + ') <span class="hint">acréscimo por opção</span></div><div class="cond-list" id="condList">';
    (subTabRule.condicoes || []).forEach(function (c, i) { html += condRowHtml(c, i); });
    html += '</div><button type="button" class="add-cond" id="btnAddCond"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><path d="M12 5v14M5 12h14"/></svg>Adicionar condição</button></div>';

    html += '<div class="field-block"><div class="field-label">Prévia de cálculo por nº de módulos (' + state.selectedSubTab + ')</div>' +
      '<table class="preview-table"><thead><tr><th>Condição</th>' + [1, 2, 3, 4, 5, 6, 7, 8].map(function (m) { return '<th>' + m + 'm</th>'; }).join('') + '</tr></thead><tbody id="previewBody"></tbody></table></div>';

    html += '<div class="editor-foot"><button type="button" class="btn" id="btnCancel">Cancelar</button>' +
      '<button type="button" class="btn primary" id="btnSave" disabled>Salvar Regras de ' + state.selectedCampoKey + '</button></div></div>';

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
      { tipo: 'escala_multiplicativa', label: '📈 Escala Multiplicativa' },
      { tipo: 'tabela', label: '📊 Tabela por Módulo (1m a 8m)' },
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
              { id: 'm0', nome: 'Demais estruturas', padrao: true, cond: [], it: [{ t: 'blk', v: 'b1' }, { t: 'op', v: '*' }, { t: 'var', v: 'lat_mb' }, { t: 'op', v: '+' }, { t: 'num', v: 2 }, { t: 'op', v: '*' }, { t: 'blk', v: 'b2' }, { t: 'op', v: '*' }, { t: 'var', v: 'ffd_mb' }, { t: 'op', v: '+' }, { t: 'blk', v: 'b3' }, { t: 'op', v: '*' }, { t: 'var', v: 'tet_mb' }, { t: 'op', v: '+' }, { t: 'num', v: 2 }, { t: 'op', v: '*' }, { t: 'blk', v: 'b4' }, { t: 'op', v: '*' }, { t: 'var', v: 'tlh_mb' }, { t: 'op', v: '+' }, { t: 'blk', v: 'b5' }, { t: 'op', v: '*' }, { t: 'var', v: 'bas_mb' }] }
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
    document.querySelectorAll('.cond-row').forEach(function (row) {
      var idx = +row.dataset.idx;
      var c = subTabRule.condicoes[idx];
      if (!c) return;
      row.querySelector('.cond-flag').addEventListener('change', function () {
        c.flag = this.value;
        c.rotulo = flagNome(this.value);
        markDirty();
        renderPreview(originalSubTab, subTabRule);
      });
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
      row.querySelector('.cond-remove').addEventListener('click', function () {
        subTabRule.condicoes.splice(idx, 1);
        renderEditor();
      });
    });
  }

  function renderPreview(originalSubTab, subTabRule) {
    var body = $('previewBody');
    if (!body) return;
    var areaObj = state.regrasData[state.selectedAreaIdx];
    var campoObj = (areaObj && areaObj.campos) ? areaObj.campos[state.selectedCampoKey] : null;

    // Collect all condition flags from both H and DUR (and subTabRule)
    var conds = [];
    if (campoObj && campoObj.H && campoObj.H.condicoes) conds = conds.concat(campoObj.H.condicoes);
    if (campoObj && campoObj.DUR && campoObj.DUR.condicoes) conds = conds.concat(campoObj.DUR.condicoes);
    if (subTabRule && subTabRule.condicoes) conds = conds.concat(subTabRule.condicoes);

    var uniqueFlags = [];
    var seenFlags = {};
    conds.forEach(function (c) {
      if (c && c.flag && !seenFlags[c.flag]) {
        seenFlags[c.flag] = true;
        uniqueFlags.push(c);
      }
    });

    function rowFor(label, flagsAtivos) {
      var isDur = state.selectedSubTab === 'DUR';
      var subTabH = campoObj ? campoObj['H'] : null;

      var hOrigVals = [];
      var hNewVals = [];
      if (isDur && subTabH) {
        hOrigVals = [1, 2, 3, 4, 5, 6, 7, 8].map(function (m) { return calcValor(subTabH, m, flagsAtivos, campoObj); });
        hNewVals = [1, 2, 3, 4, 5, 6, 7, 8].map(function (m) { return calcValor(subTabH, m, flagsAtivos, campoObj); });
      }

      var origVals = [1, 2, 3, 4, 5, 6, 7, 8].map(function (m, idx) {
        return calcValor(originalSubTab, m, flagsAtivos, campoObj, isDur ? hOrigVals[idx] : undefined);
      });
      var newVals = [1, 2, 3, 4, 5, 6, 7, 8].map(function (m, idx) {
        return calcValor(subTabRule, m, flagsAtivos, campoObj, isDur ? hNewVals[idx] : undefined);
      });
      return '<tr><td>' + label + '</td>' + newVals.map(function (v, i) {
        var changed = v !== null && origVals[i] !== null && Math.abs(v - origVals[i]) > 0.001;
        return '<td class="' + (changed ? 'changed' : '') + '">' + (v !== null ? (typeof v === 'number' ? v.toFixed(2) : v) : '—') + '</td>';
      }).join('') + '</tr>';
    }

    var html = rowFor('nenhuma marcada', {});
    uniqueFlags.forEach(function (c) {
      var m = {}; m[c.flag] = true;
      html += rowFor((c.rotulo || flagNome(c.flag)) + ' marcado', m);
    });
    body.innerHTML = html;
  }

  function saveRegrasCampo() {
    ['btnSave', 'btnSaveFooter'].forEach(function (id) {
      var btn = $(id);
      if (btn) { btn.disabled = true; btn.textContent = 'Salvando…'; }
    });

    var areaObj = state.regrasData[state.selectedAreaIdx];
    areaObj.campos[state.selectedCampoKey][state.selectedSubTab] = JSON.parse(JSON.stringify(state.dirtySubTabRule));

    if (isPyWebviewAvailable()) {
      window.pywebview.api.save_regras(state.regrasData).then(function (res) {
        if (res && res.status === 'success') {
          onRegrasSaveSuccess();
        } else {
          showToast('Regras salvas na sessão.');
          onRegrasSaveSuccess();
        }
      }).catch(function (err) {
        showToast('Erro ao salvar em regras.json: ' + err.message, true);
        ['btnSave', 'btnSaveFooter'].forEach(function (id) {
          var btn = $(id);
          if (btn) { btn.disabled = false; btn.textContent = 'Salvar regras de ' + state.selectedCampoKey; }
        });
      });
    } else {
      showToast('Regras salvas na sessão!');
      onRegrasSaveSuccess();
    }
  }

  if ($('btnSaveFooter')) {
    $('btnSaveFooter').addEventListener('click', saveRegrasCampo);
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

      histModalBody.appendChild(item);
    });
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

    // 2. Valor base
    if (chVal) {
      antesLis.push('<li class="changed">Valor base: <b class="num-diff num-diff-before">' + escapeHtml(aVal) + '</b></li>');
      depoisLis.push('<li class="changed">Valor base: <b class="num-diff num-diff-after">' + escapeHtml(dVal) + '</b></li>');
    } else {
      antesLis.push('<li>Valor base: <b>' + escapeHtml(aVal) + '</b></li>');
      depoisLis.push('<li>Valor base: <b>' + escapeHtml(dVal) + '</b></li>');
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

    // 6. Condições
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

      antesLis.push('<li class="' + (chConds ? 'changed' : '') + '">Condições: ' + aCondStr + '</li>');
      depoisLis.push('<li class="' + (chConds ? 'changed' : '') + '">Condições: ' + dCondStr + '</li>');
    } else {
      antesLis.push('<li>Condições: <b>Nenhuma</b></li>');
      depoisLis.push('<li>Condições: <b>Nenhuma</b></li>');
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
    if (base.valor_base !== undefined || base.valor !== undefined) {
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

    var conds = rule.condicoes || [];
    if (conds.length > 0) {
      var condsStr = conds.map(function (c) {
        var rot = c.rotulo || c.flag;
        var val = c.valor !== undefined ? c.valor : '';
        return '<li>' + escapeHtml(rot) + ' (' + c.forma + '): <b>' + val + '</b></li>';
      }).join('');
      items.push('<b>Condições Adicionais (' + conds.length + '):</b><ul style="margin:4px 0 0 0;">' + condsStr + '</ul>');
    } else {
      items.push('<b>Condições:</b> <em>Nenhuma condição adicional</em>');
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
        '<div style="font-size:11px; color:var(--text-faint); margin-top:4px;">📅 ' + escapeHtml(h.quando) + '</div>' +
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
        window.pywebview.api.save_config(parsedConfig).catch(function () {});
        if (parsedRegras) {
          window.pywebview.api.save_regras(parsedRegras).then(function (res) {
            showToast('Configurações e regras.json salvas com sucesso!');
          }).catch(function () {
            showToast('Configurações aplicadas.');
          });
        } else {
          showToast('Configurações salvas em config.json!');
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

  // Load external configuration on startup
  loadExternalConfig();

})();
