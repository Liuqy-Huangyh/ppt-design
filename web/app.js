/* PPT Studio 前端：零构建、零依赖、原生 DOM。
 *
 * 约定：
 * - 所有用户可控文本一律走 textContent / value，绝不用 HTML 字符串拼进页面；
 * - 请求失败时读取服务端统一错误体 {"error":{"message":...}}，把 message 显示给用户；
 * - 三块能力（生成 / 案例库 / 反馈）各自有 loading、error、empty 三种状态。
 */
'use strict';

(function () {
  const MAX_UPLOAD_BYTES = 25 * 1024 * 1024;
  const ALLOWED_UPLOAD_EXT = ['.pptx', '.ppt'];

  const state = {
    themes: [],
    profiles: [],
    decks: [],
    lastDeck: null
  };

  /* ------------------------------------------------------------------ *
   * 小工具
   * ------------------------------------------------------------------ */

  function $(selector, scope) {
    return (scope || document).querySelector(selector);
  }

  function $$(selector, scope) {
    return Array.prototype.slice.call((scope || document).querySelectorAll(selector));
  }

  function show(element, visible) {
    if (element) element.hidden = !visible;
  }

  function setStatus(element, text, busy) {
    if (!element) return;
    element.textContent = text || '';
    element.hidden = !text;
    element.dataset.busy = busy && text ? 'true' : 'false';
  }

  function setError(element, message) {
    if (!element) return;
    element.textContent = message || '';
    element.hidden = !message;
  }

  function textOf(value, fallback) {
    if (value === null || value === undefined || value === '') {
      return fallback === undefined ? '—' : fallback;
    }
    return String(value);
  }

  function numberText(value, digits, fallback) {
    const number = Number(value);
    if (!Number.isFinite(number)) return fallback === undefined ? '—' : fallback;
    return digits > 0 ? number.toFixed(digits) : String(Math.round(number));
  }

  function formatDate(value) {
    const text = String(value || '');
    if (!text) return '—';
    return text.replace('T', ' ').slice(0, 16);
  }

  function isHexColor(value) {
    return /^#[0-9A-Fa-f]{6}$/.test(String(value || ''));
  }

  function safeFilename(title) {
    const base = String(title || 'ppt-studio').replace(/[\\/:*?"<>|\r\n\t]+/g, '-').trim();
    return (base || 'ppt-studio').slice(0, 60);
  }

  function todayLabel() {
    const now = new Date();
    return now.getFullYear() + '.' + String(now.getMonth() + 1).padStart(2, '0');
  }

  function el(tag, className, text) {
    const node = document.createElement(tag);
    if (className) node.className = className;
    if (text !== undefined && text !== null) node.textContent = String(text);
    return node;
  }

  async function request(path, options) {
    const response = await fetch(path, options || {});
    const raw = await response.text();
    let payload = null;
    if (raw) {
      try {
        payload = JSON.parse(raw);
      } catch (error) {
        payload = null;
      }
    }
    if (!response.ok) {
      const detail = payload && payload.error ? payload.error : null;
      const message = detail && detail.message
        ? String(detail.message)
        : '请求失败（服务返回 ' + response.status + '）';
      const failure = new Error(message);
      failure.code = detail && detail.code ? String(detail.code) : '';
      failure.status = response.status;
      throw failure;
    }
    return payload;
  }

  function setBusy(button, busy, busyLabel, idleLabel) {
    if (!button) return;
    button.disabled = busy;
    button.textContent = busy ? busyLabel : idleLabel;
  }

  /* ------------------------------------------------------------------ *
   * 启动
   * ------------------------------------------------------------------ */

  document.addEventListener('DOMContentLoaded', function () {
    const dateInput = $('#f-date');
    if (dateInput && !dateInput.value) dateInput.value = todayLabel();

    bindModeToggle();
    bindGenerate();
    bindUpload();
    bindFeedback();

    window.addEventListener('resize', fitPreview);
    const preview = $('#gen-preview');
    if (preview) preview.addEventListener('load', fitPreview);

    loadHealth();
    loadThemes();
    loadProfiles();
    loadDecks();
  });

  /* ------------------------------------------------------------------ *
   * 服务状态 / 主题
   * ------------------------------------------------------------------ */

  async function loadHealth() {
    const badge = $('#service-status');
    const label = $('#service-label');
    try {
      const info = await request('/api/health');
      const llm = Boolean(info && info.llm);
      label.textContent = llm ? 'LLM 已启用' : '离线模式';
      badge.classList.add('is-ok');
      badge.classList.remove('is-error');
      $('#service-version').textContent = 'v' + textOf(info && info.version, '1.0.0');
      $('#service-mode').textContent = llm ? '在线扩写大纲' : '离线启发式大纲';
    } catch (error) {
      label.textContent = '服务未连接 · 请用 .\\ppt.cmd web 启动后打开';
      badge.classList.remove('is-ok');
      badge.classList.add('is-error');
      $('#service-mode').textContent = '未连接';
    }
  }

  async function loadThemes() {
    const select = $('#f-theme');
    try {
      const themes = await request('/api/themes');
      state.themes = Array.isArray(themes) ? themes : [];
      if (!state.themes.length) throw new Error('服务没有返回任何主题');
      select.replaceChildren();
      state.themes.forEach(function (theme) {
        const option = document.createElement('option');
        option.value = String(theme.key || '');
        option.textContent = textOf(theme.label, theme.key);
        select.append(option);
      });
    } catch (error) {
      setError($('#gen-error'), '主题加载失败：' + error.message);
    }
  }

  /* ------------------------------------------------------------------ *
   * 生成
   * ------------------------------------------------------------------ */

  function bindModeToggle() {
    const promptPanel = $('#prompt-panel');
    const outlinePanel = $('#outline-panel');
    const radios = $$('input[name="content-mode"]');

    function sync(shouldFocus) {
      const checked = $('input[name="content-mode"]:checked');
      const mode = checked ? checked.value : 'prompt';
      show(promptPanel, mode === 'prompt');
      show(outlinePanel, mode === 'outline');
      if (shouldFocus) {
        const target = mode === 'prompt' ? $('#f-prompt') : $('#f-outline');
        if (target) target.focus();
      }
    }

    radios.forEach(function (radio) {
      radio.addEventListener('change', function () { sync(true); });
    });
    sync(false);
  }

  function collectGeneratePayload() {
    const payload = {};
    ['title', 'subtitle', 'author', 'date'].forEach(function (key) {
      const field = $('#f-' + key);
      const value = field ? field.value.trim() : '';
      if (value) payload[key] = value;
    });

    const theme = $('#f-theme').value;
    if (theme) payload.theme = theme;

    const pagesField = $('#f-pages');
    const pages = Number(pagesField ? pagesField.value : '');
    payload.pages = Number.isFinite(pages) ? Math.min(40, Math.max(2, Math.round(pages))) : 10;

    const profileId = $('#f-profile').value;
    if (profileId) payload.profile_id = profileId;

    const checked = $('input[name="content-mode"]:checked');
    const mode = checked && checked.value === 'outline' ? 'outline' : 'prompt';
    const field = mode === 'outline' ? $('#f-outline') : $('#f-prompt');
    const text = field ? field.value.trim() : '';
    if (text) payload[mode] = text;

    if (!payload.prompt && !payload.outline) {
      throw new Error(mode === 'outline'
        ? '请在大纲里贴入 Markdown（每页以 # 标题 开头）'
        : '请填写提示词，或切换到大纲并贴入 Markdown');
    }
    return payload;
  }

  function bindGenerate() {
    const form = $('#deck-form');
    form.addEventListener('submit', async function (event) {
      event.preventDefault();
      const submit = $('#gen-submit');
      const errorBox = $('#gen-error');
      const status = $('#gen-status');

      setError(errorBox, '');
      let payload;
      try {
        payload = collectGeneratePayload();
      } catch (error) {
        setError(errorBox, error.message);
        return;
      }

      setStatus(status, '正在生成：解析大纲、绘制页面…', true);
      setBusy(submit, true, '生成中…', '生成演示');
      form.setAttribute('aria-busy', 'true');
      show($('#gen-empty'), false);
      try {
        const result = await request('/api/generate', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload)
        });
        renderDeck(result, result.title || payload.title);
        state.lastDeck = result;
        await loadDecks(result.id);
        setStatus(status, '完成：共 ' + numberText(result.pages, 0) + ' 页');
      } catch (error) {
        setStatus(status, '');
        setError(errorBox, error.message);
        show($('#gen-result'), false);
        show($('#gen-empty'), true);
      } finally {
        setBusy(submit, false, '生成中…', '生成演示');
        form.setAttribute('aria-busy', 'false');
      }
    });
  }

  // deck.html 以 1280×720 为舞台尺寸设计；iframe 保持这个原生尺寸，
  // 再由 16:9 容器统一缩放，预览里才能看到完整的整页版面。
  function fitPreview() {
    const frame = $('#preview-frame');
    if (!frame) return;
    const width = frame.clientWidth;
    if (!width) return;
    frame.style.setProperty('--preview-scale', String(width / 1280));
  }

  function renderDeck(result, title) {
    if (!result || !result.files) throw new Error('服务端返回缺少文件信息');
    // 服务端会用 meta.title 兜底（标题留空时取提示词/首标题），优先采用它，
    // 避免结果卡显示“未命名演示”、下载名与 meta 不一致。
    const displayTitle = textOf(result.title || title, '');
    const files = result.files;
    const deckId = result.id ? encodeURIComponent(result.id) : '';
    const pptxUrl = files.pptx || (deckId ? '/files/' + deckId + '/deck.pptx' : '');
    const htmlUrl = files.html || (deckId ? '/files/' + deckId + '/deck.html' : '');

    const pptxLink = $('#gen-pptx');
    pptxLink.href = pptxUrl || '#';
    pptxLink.textContent = '下载 .pptx（' + numberText(result.pages, 0) + ' 页）';
    pptxLink.setAttribute('download', safeFilename(displayTitle) + '.pptx');

    const htmlLink = $('#gen-html');
    htmlLink.href = htmlUrl || '#';

    // 先让结果区可见，再给 iframe 设 src：deck.html 在加载时按窗口尺寸缩放舞台，
    // 若容器还带着 hidden，内窗宽高为 0，舞台会被 scale(0) 吞掉（只剩黑底）。
    show($('#gen-result'), true);
    show($('#gen-empty'), false);

    const preview = $('#gen-preview');
    if (htmlUrl) {
      preview.src = htmlUrl;
      fitPreview();
    } else {
      preview.removeAttribute('src');
    }

    $('#gen-title').textContent = textOf(displayTitle, '未命名演示');
    $('#gen-pages').textContent = numberText(result.pages, 0);
    $('#gen-theme').textContent = textOf(result.theme);
    $('#gen-llm').textContent = result.llm_used ? 'LLM 扩写' : '离线大纲';

    const notes = $('#gen-notes');
    notes.replaceChildren();
    (Array.isArray(result.notes) ? result.notes : []).forEach(function (note) {
      notes.append(el('li', '', String(note)));
    });
  }

  /* ------------------------------------------------------------------ *
   * 案例库（上传 / 档案卡片）
   * ------------------------------------------------------------------ */

  function bindUpload() {
    const zone = $('#dropzone');
    const input = $('#file-input');

    function openPicker() {
      if (input) input.click();
    }

    zone.addEventListener('click', function (event) {
      if (event.target !== input) openPicker();
    });
    zone.addEventListener('keydown', function (event) {
      if (event.key === 'Enter' || event.key === ' ') {
        event.preventDefault();
        openPicker();
      }
    });
    input.addEventListener('change', function () {
      const file = input.files && input.files[0];
      input.value = '';
      if (file) uploadFile(file);
    });

    ['dragenter', 'dragover'].forEach(function (name) {
      zone.addEventListener(name, function (event) {
        event.preventDefault();
        zone.classList.add('is-dragging');
      });
    });
    ['dragleave', 'dragend'].forEach(function (name) {
      zone.addEventListener(name, function () { zone.classList.remove('is-dragging'); });
    });
    zone.addEventListener('drop', function (event) {
      event.preventDefault();
      zone.classList.remove('is-dragging');
      const file = event.dataTransfer && event.dataTransfer.files && event.dataTransfer.files[0];
      if (file) uploadFile(file);
    });

    const refresh = $('#profiles-refresh');
    if (refresh) {
      refresh.addEventListener('click', function () {
        refresh.disabled = true;
        loadProfiles().finally(function () { refresh.disabled = false; });
      });
    }
  }

  async function uploadFile(file) {
    const status = $('#upload-status');
    const errorBox = $('#upload-error');
    const zone = $('#dropzone');
    setError(errorBox, '');

    const name = String(file.name || '');
    const lower = name.toLowerCase();
    const ext = lower.slice(lower.lastIndexOf('.'));
    if (ALLOWED_UPLOAD_EXT.indexOf(ext) === -1) {
      setError(errorBox, '只支持 .pptx / .ppt 文件：' + name);
      return;
    }
    if (file.size > MAX_UPLOAD_BYTES) {
      setError(errorBox, '「' + name + '」超过 25 MB 上限，请先精简后再上传。');
      return;
    }

    setStatus(status, '正在分析「' + name + '」的配色、字体与信息密度…', true);
    zone.classList.add('is-busy');
    try {
      const result = await request('/api/profiles', {
        method: 'POST',
        headers: {
          'X-Filename': encodeURIComponent(name),
          'Content-Type': 'application/octet-stream'
        },
        body: file
      });
      const profile = (result && result.profile) || {};
      setStatus(status, '已提取档案：' + textOf(profile.name, name)
        + '（密度 ' + numberText(profile.style && profile.style.density, 2) + '）');
      await loadProfiles();
      const select = $('#f-profile');
      const id = result && result.id;
      if (id && Array.prototype.some.call(select.options, function (option) {
        return option.value === id;
      })) {
        select.value = id;
      }
    } catch (error) {
      setStatus(status, '');
      setError(errorBox, error.message);
    } finally {
      zone.classList.remove('is-busy');
    }
  }

  async function loadProfiles() {
    const list = $('#profile-list');
    const empty = $('#profiles-empty');
    try {
      const summaries = await request('/api/profiles');
      const items = Array.isArray(summaries) ? summaries : [];
      const details = await Promise.allSettled(items.map(function (item) {
        return request('/api/profiles/' + encodeURIComponent(item.id));
      }));
      const profiles = items.map(function (item, index) {
        const full = details[index] && details[index].status === 'fulfilled'
          ? details[index].value
          : null;
        const merged = Object.assign({}, item, full && typeof full === 'object' ? full : {});
        if (!merged.metrics && full && full.metrics) merged.metrics = full.metrics;
        return merged;
      });
      state.profiles = profiles;
      list.replaceChildren();
      profiles.forEach(function (profile) { list.append(profileCard(profile)); });
      show(empty, profiles.length === 0);
      renderProfileOptions();
      updateFeedbackHint();
      setError($('#upload-error'), '');
    } catch (error) {
      setError($('#upload-error'), '档案列表加载失败：' + error.message);
    }
  }

  function renderProfileOptions() {
    const select = $('#f-profile');
    if (!select) return;
    const current = select.value;
    select.replaceChildren();
    const none = document.createElement('option');
    none.value = '';
    none.textContent = '不使用档案（基础主题）';
    select.append(none);
    state.profiles.forEach(function (profile) {
      const option = document.createElement('option');
      option.value = String(profile.id || '');
      option.textContent = textOf(profile.name, profile.id)
        + ' · 密度 ' + numberText(profile.style && profile.style.density, 2);
      select.append(option);
    });
    const keep = Array.prototype.some.call(select.options, function (option) {
      return option.value === current;
    });
    select.value = keep ? current : '';
  }

  function fontLine(fonts) {
    if (!fonts || typeof fonts !== 'object') return '—';
    const title = textOf(fonts.title, '');
    const body = textOf(fonts.body, '');
    if (title && body && title !== body) return title + ' / ' + body;
    return title || body || '—';
  }

  function profileCard(profile) {
    const card = el('article', 'profile-card');
    card.dataset.profileId = String(profile.id || '');

    const head = el('div', 'profile-head');
    head.append(el('h4', 'profile-name', textOf(profile.name, '未命名案例')));
    head.append(el('p', 'profile-date', formatDate(profile.created_at)));
    card.append(head);

    const palette = profile.palette && typeof profile.palette === 'object' ? profile.palette : {};
    const labels = {
      bg: '底色', ink: '文字', muted: '次要', accent: '主色',
      accent_soft: '浅主色', rule: '分隔线'
    };
    const swatches = el('div', 'swatches');
    Object.keys(palette).forEach(function (key) {
      const color = palette[key];
      if (!isHexColor(color)) return;
      const cell = el('span', 'swatch');
      const chip = el('span', 'swatch-chip');
      chip.style.backgroundColor = color;
      chip.title = (labels[key] || key) + ' ' + color;
      cell.append(chip, el('span', 'swatch-label', labels[key] || key));
      swatches.append(cell);
    });
    if (swatches.childNodes.length) card.append(swatches);

    const metrics = profile.metrics && typeof profile.metrics === 'object' ? profile.metrics : {};
    const style = profile.style && typeof profile.style === 'object' ? profile.style : {};
    const learn = profile.learn && typeof profile.learn === 'object' ? profile.learn : {};

    const facts = el('dl', 'facts');
    [
      ['字体', fontLine(profile.fonts)],
      ['每页字数', numberText(metrics.chars_per_slide, 1)],
      ['每页要点', numberText(metrics.bullets_per_slide, 1)],
      ['页数', numberText(metrics.slides !== undefined ? metrics.slides : profile.slides, 0)],
      ['信息密度', numberText(style.density, 2)],
      ['反馈轮次', numberText(learn.rounds, 0)]
    ].forEach(function (pair) {
      facts.append(el('dt', '', pair[0]), el('dd', '', pair[1]));
    });
    card.append(facts);

    const use = el('button', 'button button--ghost', '用这个档案生成');
    use.type = 'button';
    use.addEventListener('click', function () {
      const select = $('#f-profile');
      select.value = String(profile.id || '');
      setStatus($('#gen-status'), '已选择档案「' + textOf(profile.name, profile.id) + '」');
      $('#panel-generate').scrollIntoView({ behavior: 'smooth', block: 'start' });
    });
    card.append(use);
    return card;
  }

  /* ------------------------------------------------------------------ *
   * 反馈
   * ------------------------------------------------------------------ */

  function bindFeedback() {
    $('#fb-deck').addEventListener('change', updateFeedbackHint);
    $('#feedback-form').addEventListener('submit', async function (event) {
      event.preventDefault();
      const submit = $('#fb-submit');
      const status = $('#fb-status');
      const errorBox = $('#fb-error');
      setError(errorBox, '');
      show($('#fb-result'), false);

      const deck = selectedDeck();
      if (!deck) {
        setError(errorBox, '请先选择一份 deck；还没有生成记录的话，到 01 生成一份。');
        return;
      }
      const checked = $('input[name="rating"]:checked');
      const rating = checked ? Number(checked.value) : NaN;
      if (!(rating >= 1 && rating <= 5)) {
        setError(errorBox, '请选择 1–5 分。');
        return;
      }
      const notesField = $('#fb-notes');

      setStatus(status, '正在提交评分，并写回风格档案…', true);
      setBusy(submit, true, '提交中…', '提交反馈');
      try {
        const result = await request('/api/decks/' + encodeURIComponent(deck.id) + '/feedback', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ rating: rating, notes: notesField ? notesField.value.trim() : '' })
        });
        renderFeedback(result, deck);
        await loadProfiles();
        setStatus(status, '已记录。档案已更新，下次生成立刻生效。');
      } catch (error) {
        setStatus(status, '');
        setError(errorBox, error.message);
        show($('#fb-empty'), true);
      } finally {
        setBusy(submit, false, '提交中…', '提交反馈');
      }
    });
  }

  function selectedDeck() {
    const select = $('#fb-deck');
    if (!select) return null;
    const id = select.value;
    if (!id) return null;
    return state.decks.find(function (deck) { return String(deck.id) === id; }) || null;
  }

  function updateFeedbackHint() {
    const hint = $('#fb-hint');
    if (!hint) return;
    const deck = selectedDeck();
    if (!deck) {
      hint.textContent = '生成一份 deck 后，就能在这里给它打分。';
      return;
    }
    if (!deck.profile_id) {
      hint.textContent = '这份 deck 生成时没有选择风格档案：评分会被记录，但不会调整档案。';
      return;
    }
    const profile = state.profiles.filter(function (item) {
      return String(item.id) === String(deck.profile_id);
    })[0];
    hint.textContent = '将调整档案：' + textOf(profile && profile.name, deck.profile_id)
      + '（当前密度 ' + numberText(profile && profile.style && profile.style.density, 2) + '）';
  }

  function renderFeedback(result, deck) {
    const message = result && result.message ? String(result.message) : '已记录评分';
    $('#fb-message').textContent = message;

    const style = result && result.style && typeof result.style === 'object' ? result.style : {};
    const parts = [];
    if (Number.isFinite(Number(style.density))) parts.push('密度 ' + Number(style.density).toFixed(2));
    if (Number.isFinite(Number(style.title_scale))) parts.push('标题比例 ' + Number(style.title_scale).toFixed(2));
    $('#fb-style').textContent = parts.length ? '当前档案：' + parts.join(' · ') : '';
    $('#fb-profile').textContent = '档案 ' + textOf(result && result.profile_id)
      + ' · deck ' + textOf(deck && deck.id);

    show($('#fb-result'), true);
    show($('#fb-empty'), false);
  }

  async function loadDecks(selectId) {
    const select = $('#fb-deck');
    if (!select) return;
    try {
      const decks = await request('/api/decks');
      state.decks = Array.isArray(decks) ? decks : [];
      const previous = selectId || select.value;
      select.replaceChildren();
      if (!state.decks.length) {
        const option = document.createElement('option');
        option.value = '';
        option.textContent = '（还没有生成记录）';
        select.append(option);
        updateFeedbackHint();
        return;
      }
      state.decks.forEach(function (deck) {
        const option = document.createElement('option');
        option.value = String(deck.id || '');
        option.textContent = textOf(deck.title, '未命名演示')
          + ' · ' + numberText(deck.pages, 0) + ' 页 · ' + formatDate(deck.created_at);
        select.append(option);
      });
      const keep = Array.prototype.some.call(select.options, function (option) {
        return option.value === previous;
      });
      if (keep) select.value = previous;
      updateFeedbackHint();
    } catch (error) {
      setError($('#fb-error'), '生成记录加载失败：' + error.message);
    }
  }
})();
