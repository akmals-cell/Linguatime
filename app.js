  const SUPABASE_URL = 'https://srpiqxfeosusayqqfvrq.supabase.co';
  const SUPABASE_ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InNycGlxeGZlb3N1c2F5cXFmdnJxIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzg4MjAzMTQsImV4cCI6MjA5NDM5NjMxNH0.PyOLqLn93h7I0_VYL-vE9O8VTEgdyDmS7UWxSAKA3Ss';
  const CREATE_TRANSLATOR_URL = `${SUPABASE_URL}/functions/v1/create-translator`;

  const sb = window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

  let currentUser = null;
  let languagePairs = [];
  let pairRowCounter = 0;

  // ═══════════════════════════════════════════════════════════════════════
  // Инфраструктура: TTL-кеш, ленивая загрузка скриптов, debounce, часовые пояса.
  // Добавлено для оптимизации — не меняет бизнес-логику.
  // ═══════════════════════════════════════════════════════════════════════

  // TTL-кеш для серверных запросов. Ключ — строка вида "rateHistory:uid1,uid2".
  // Инвалидация по префиксу: invalidateCache('rateHistory:') чистит всю группу.
  const CACHE_TTL_MS = 60 * 1000; // 60 секунд — компромисс свежесть/скорость
  const callCache = new Map();    // key → { data, expires }

  function cachedCall(key, ttlMs, loader) {
    const hit = callCache.get(key);
    if (hit && hit.expires > Date.now()) return Promise.resolve(hit.data);
    return loader().then(data => {
      callCache.set(key, { data, expires: Date.now() + ttlMs });
      return data;
    });
  }
  function invalidateCache(prefix) {
    for (const key of Array.from(callCache.keys())) {
      if (key.startsWith(prefix)) callCache.delete(key);
    }
  }

  // Ленивая загрузка стороннего скрипта (XLSX, jsPDF). Кеширует промис, поэтому
  // повторные вызовы дёшевы. Если один и тот же src уже на странице — резолвим сразу.
  const scriptLoadPromises = new Map();
  function loadScript(src) {
    if (scriptLoadPromises.has(src)) return scriptLoadPromises.get(src);
    const p = new Promise((resolve, reject) => {
      const existing = document.querySelector(`script[src="${src}"]`);
      if (existing) { resolve(); return; }
      const s = document.createElement('script');
      s.src = src;
      s.async = true;
      s.onload = () => resolve();
      s.onerror = () => reject(new Error('Не удалось загрузить ' + src));
      document.head.appendChild(s);
    });
    scriptLoadPromises.set(src, p);
    return p;
  }
  async function ensureXlsx() {
    if (window.XLSX) return;
    await loadScript('https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js');
  }
  async function ensureJspdf() {
    if (window.jspdf && window.jspdf.jsPDF && window.jspdf.jsPDF.API.autoTable) return;
    await loadScript('https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js');
    await loadScript('https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.8.0/jspdf.plugin.autotable.min.js');
  }

  // Простой debounce — задерживает вызов до тех пор, пока не пройдёт ms тишины.
  function debounce(fn, ms) {
    let t;
    return function debounced(...args) {
      clearTimeout(t);
      t = setTimeout(() => fn.apply(this, args), ms);
    };
  }

  // Список часовых поясов — единый источник правды, заполняет все селекты с классом
  // js-timezone-select при старте приложения.
  const TIMEZONES = [
    ['Asia/Tashkent',       'Asia/Tashkent (UTC+5)'],
    ['Asia/Almaty',         'Asia/Almaty (UTC+5)'],
    ['Europe/Moscow',       'Europe/Moscow (UTC+3)'],
    ['Europe/Kyiv',         'Europe/Kyiv (UTC+2)'],
    ['Europe/London',       'Europe/London (UTC+0)'],
    ['Europe/Berlin',       'Europe/Berlin (UTC+1)'],
    ['Asia/Seoul',          'Asia/Seoul (UTC+9)'],
    ['Asia/Tokyo',          'Asia/Tokyo (UTC+9)'],
    ['Asia/Shanghai',       'Asia/Shanghai (UTC+8)'],
    ['Asia/Dubai',          'Asia/Dubai (UTC+4)'],
    ['America/New_York',    'America/New_York (UTC-5)'],
    ['America/Los_Angeles', 'America/Los_Angeles (UTC-8)'],
  ];
  function fillTimezoneSelects() {
    const html = TIMEZONES.map(([v, l]) => `<option value="${v}">${l}</option>`).join('');
    document.querySelectorAll('.js-timezone-select').forEach(sel => {
      // Сохраняем выбранное значение, если оно есть
      const prev = sel.value;
      sel.innerHTML = html;
      if (prev) sel.value = prev;
      else sel.value = 'Asia/Tashkent'; // дефолт
    });
  }

  // Текущая страница — для пропуска повторных загрузок при клике на тот же пункт меню.
  let currentPage = null;

  function showScreen(id) {
    ['loading', 'login', 'forgot', 'reset', 'app'].forEach(s =>
      document.getElementById('screen-' + s).classList.add('hidden'));
    document.getElementById('screen-' + id).classList.remove('hidden');
  }
  function showPage(pageId) {
    currentPage = pageId; // синхронизируем — listener использует это для skip-логики
    document.querySelectorAll('.page').forEach(p => p.classList.add('hidden'));
    document.getElementById('page-' + pageId).classList.remove('hidden');
    document.querySelectorAll('.sb-item').forEach(b => b.classList.remove('active'));
    const btn = document.querySelector(`.sb-item[data-page="${pageId}"]`);
    if (btn) btn.classList.add('active');
  }

  async function checkSession() {
    // Сначала проверяем, не вернулся ли пользователь по ссылке recovery
    // Supabase кладёт хеш-параметры вида #access_token=...&type=recovery
    const hash = window.location.hash;
    if (hash && hash.includes('type=recovery')) {
      // Не делаем ничего здесь — onAuthStateChange ниже среагирует на PASSWORD_RECOVERY
      // и переключит экран
      return;
    }

    const { data: { session } } = await sb.auth.getSession();
    if (session) await loadProfile(session.user.id);
    else showScreen('login');
  }

  // Supabase шлёт событие PASSWORD_RECOVERY когда пользователь приходит по ссылке из email
  sb.auth.onAuthStateChange((event, session) => {
    if (event === 'PASSWORD_RECOVERY') {
      showScreen('reset');
      // Очищаем хеш чтобы при следующем входе не срабатывало
      // (но только после того как покажем экран — Supabase ещё использует параметры)
    }
  });

  async function loadProfile(userId) {
    const { data, error } = await sb.from('users')
      .select('id, email, name, role, timezone, is_active, is_super_admin, can_access_clients')
      .eq('id', userId).single();

    if (error || !data) {
      await sb.auth.signOut();
      showScreen('login');
      showError('login-error', 'Профиль не найден. Обратитесь к менеджеру.');
      return;
    }
    if (!data.is_active) {
      await sb.auth.signOut();
      showScreen('login');
      showError('login-error', 'Учётная запись деактивирована.');
      return;
    }

    currentUser = data;
    document.getElementById('sb-user-name').textContent = data.name;
    document.getElementById('sb-user-email').textContent = data.email;

    if (data.role === 'manager') {
      document.getElementById('nav-manager').classList.remove('hidden');
      document.getElementById('nav-translator').classList.add('hidden');
      // Секция «Клиенты» — только для менеджеров с доступом
      const clientsNav = document.querySelector('#nav-manager [data-page="clients"]');
      if (clientsNav) {
        clientsNav.style.display = data.can_access_clients ? '' : 'none';
      }
      showScreen('app');
      showPage('dashboard');
      await loadLanguagePairs();
      await loadDashboard();
      refreshRequestsBadge();
      refreshSwapBadge();
    } else {
      document.getElementById('nav-translator').classList.remove('hidden');
      document.getElementById('nav-manager').classList.add('hidden');
      showScreen('app');
      showPage('calendar');
      await loadLanguagePairs();
      await loadCalendar();
    }
  }

  async function handleLogin() {
    hideError('login-error');
    const email = document.getElementById('email').value.trim();
    const password = document.getElementById('password').value;
    if (!email || !password) { showError('login-error', 'Введите email и пароль.'); return; }
    const btn = document.getElementById('btn-login');
    btn.disabled = true; btn.textContent = 'Входим…';
    const { data, error } = await sb.auth.signInWithPassword({ email, password });
    btn.disabled = false; btn.textContent = 'Войти';
    if (error) {
      showError('login-error', error.message.includes('Invalid login credentials') ? 'Неверный email или пароль.' : ('Ошибка: ' + error.message));
      return;
    }
    await loadProfile(data.user.id);
  }

  async function handleLogout() {
    await sb.auth.signOut();
    currentUser = null;
    document.getElementById('email').value = '';
    document.getElementById('password').value = '';
    hideError('login-error');
    showScreen('login');
  }

  // ────────────────────────────────────────────────────────────────────
  // ВОССТАНОВЛЕНИЕ ПАРОЛЯ
  // ────────────────────────────────────────────────────────────────────

  async function handleForgotPassword() {
    hideError('forgot-error');
    const email = document.getElementById('forgot-email').value.trim().toLowerCase();
    if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      showError('forgot-error', 'Введите корректный email.');
      return;
    }
    const btn = document.getElementById('btn-forgot-send');
    btn.disabled = true; btn.textContent = 'Отправляем…';

    // Указываем куда вернуть пользователя после клика по ссылке в письме
    const redirectTo = window.location.origin + window.location.pathname;

    const { error } = await sb.auth.resetPasswordForEmail(email, {
      redirectTo,
    });

    btn.disabled = false; btn.textContent = 'Отправить ссылку';

    if (error) {
      showError('forgot-error', 'Ошибка: ' + error.message);
      return;
    }

    // Из соображений безопасности всегда показываем «отправлено», даже если email не найден
    // (чтобы не палить регистрации). Supabase делает то же самое.
    document.getElementById('forgot-form-wrap').classList.add('hidden');
    document.getElementById('forgot-success').classList.remove('hidden');
  }

  async function handleResetPassword() {
    hideError('reset-error');
    const pw1 = document.getElementById('reset-password').value;
    const pw2 = document.getElementById('reset-password2').value;
    if (!pw1 || pw1.length < 6) {
      showError('reset-error', 'Пароль должен быть не короче 6 символов.');
      return;
    }
    if (pw1 !== pw2) {
      showError('reset-error', 'Пароли не совпадают.');
      return;
    }

    const btn = document.getElementById('btn-reset-save');
    btn.disabled = true; btn.textContent = 'Сохраняем…';

    const { data, error } = await sb.auth.updateUser({ password: pw1 });

    if (error) {
      btn.disabled = false; btn.textContent = 'Сохранить пароль';
      showError('reset-error', 'Ошибка: ' + error.message);
      return;
    }

    // Успех! Показываем сообщение и автоматически входим
    document.getElementById('reset-form-wrap').classList.add('hidden');
    document.getElementById('reset-success').classList.remove('hidden');

    // Очищаем хеш (recovery параметры) и через 1.5 сек загружаем профиль
    setTimeout(async () => {
      window.history.replaceState(null, '', window.location.pathname);
      if (data && data.user) {
        await loadProfile(data.user.id);
      } else {
        showScreen('login');
      }
    }, 1500);
  }

  async function loadLanguagePairs() {
    languagePairs = await cachedCall('languagePairs', 5 * 60 * 1000, async () => {
      const { data } = await sb.from('language_pairs')
        .select('id, code, display_name').eq('active', true).order('code');
      return data || [];
    });
  }

  // ====================================================================
  // СТРАНИЦА: МЕНЕДЖЕРЫ
  // ====================================================================
  async function loadManagers() {
    hideError('managers-error');
    const content = document.getElementById('managers-content');
    const subtitle = document.getElementById('managers-subtitle');
    const actionsWrap = document.getElementById('managers-actions-wrap');
    content.innerHTML = '<div class="loading-state">Загрузка…</div>';

    // Загружаем менеджеров
    const { data: managers, error } = await sb.from('users')
      .select('id, email, name, is_active, is_super_admin, can_access_clients, created_at')
      .eq('role', 'manager')
      .order('created_at', { ascending: true });

    if (error) {
      showError('managers-error', 'Ошибка загрузки: ' + error.message);
      content.innerHTML = '';
      return;
    }

    const activeManagers = managers.filter(m => m.is_active);
    subtitle.textContent =
      `${activeManagers.length} ${pluralize(activeManagers.length, 'активный', 'активных', 'активных')}` +
      (managers.length > activeManagers.length
        ? ` · ${managers.length - activeManagers.length} неактивных`
        : '');

    // Кнопка «Добавить менеджера» — только для super-admin
    if (currentUser.is_super_admin) {
      actionsWrap.innerHTML = `
        <button class="btn" onclick="openAddManagerModal()">+ Добавить менеджера</button>
      `;
    } else {
      actionsWrap.innerHTML = '';
    }

    if (managers.length === 0) {
      content.innerHTML = `
        <div class="empty-state">
          <div class="empty-state-text">В системе пока нет менеджеров.</div>
        </div>
      `;
      return;
    }

    let html = `<table><thead><tr>
      <th>Менеджер</th>
      <th>Email</th>
      <th>Статус</th>
      <th>Добавлен</th>
      <th></th>
    </tr></thead><tbody>`;

    for (const m of managers) {
      const initials = (m.name || '?').split(' ').map(s => s[0]).join('').slice(0, 2).toUpperCase();
      const isCurrent = m.id === currentUser.id;

      // Бейджи
      let badges = '';
      if (m.is_super_admin) {
        badges += `<span class="badge badge-super" style="margin-left: 8px;">★ super-admin</span>`;
      }
      if (isCurrent) {
        badges += `<span class="badge badge-info" style="margin-left: 6px;">это вы</span>`;
      }

      // Статус
      const statusBadge = m.is_active
        ? '<span class="badge badge-good">● Активен</span>'
        : '<span class="badge badge-warn">● Неактивен</span>';

      // Бейдж доступа к клиентам
      const clientsBadge = m.can_access_clients
        ? '<span class="badge badge-info" style="margin-left:6px;">Клиенты ✓</span>'
        : '<span class="badge badge-neutral" style="margin-left:6px;">Клиенты ✕</span>';

      // Кнопки действий (только super-admin, не себя, не других super-admin)
      let actionBtn = '';
      if (currentUser.is_super_admin && !isCurrent && !m.is_super_admin) {
        // Тумблер доступа к клиентам
        if (m.can_access_clients) {
          actionBtn += `<button class="btn btn-ghost btn-sm" onclick="toggleClientsAccess('${m.id}', false, '${escapeHtml(m.name).replace(/'/g, "\\'")}')">Убрать клиентов</button> `;
        } else {
          actionBtn += `<button class="btn btn-ghost btn-sm" onclick="toggleClientsAccess('${m.id}', true, '${escapeHtml(m.name).replace(/'/g, "\\'")}')">Дать клиентов</button> `;
        }
        // Деактивация/активация
        if (m.is_active) {
          actionBtn += `<button class="btn btn-ghost btn-sm" onclick="toggleManagerActive('${m.id}', false, '${escapeHtml(m.name).replace(/'/g, "\\'")}')">Деактивировать</button>`;
        } else {
          actionBtn += `<button class="btn btn-ghost btn-sm" onclick="toggleManagerActive('${m.id}', true, '${escapeHtml(m.name).replace(/'/g, "\\'")}')">Активировать</button>`;
        }
      }

      const created = new Date(m.created_at).toLocaleDateString('ru-RU');

      html += `<tr>
        <td>
          <div class="emp-cell">
            <div class="emp-avatar">${initials}</div>
            <div>
              <div class="emp-name">${escapeHtml(m.name)}${badges}${clientsBadge}</div>
            </div>
          </div>
        </td>
        <td style="color: #475569;">${escapeHtml(m.email)}</td>
        <td>${statusBadge}</td>
        <td style="font-family: 'JetBrains Mono', monospace; color: #475569; font-size: 12px;">${created}</td>
        <td style="text-align: right;">${actionBtn}</td>
      </tr>`;
    }

    html += '</tbody></table>';
    content.innerHTML = html;
  }

  // ──── Модалка создания менеджера ────────────────────────────────────
  function openAddManagerModal() {
    if (!currentUser.is_super_admin) {
      alert('Только super-admin может добавлять менеджеров.');
      return;
    }
    hideError('add-manager-error');
    document.getElementById('mgr-name').value = '';
    document.getElementById('mgr-email').value = '';
    document.getElementById('mgr-password').value = '';
    document.getElementById('add-manager-modal').classList.add('open');
    setTimeout(() => document.getElementById('mgr-name').focus(), 50);
  }

  function closeAddManagerModal() {
    document.getElementById('add-manager-modal').classList.remove('open');
  }

  async function createManager() {
    hideError('add-manager-error');
    const btn = document.getElementById('btn-save-manager');

    const name = document.getElementById('mgr-name').value.trim();
    const email = document.getElementById('mgr-email').value.trim().toLowerCase();
    const password = document.getElementById('mgr-password').value;

    if (!name) { showError('add-manager-error', 'Введите имя.'); return; }
    if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      showError('add-manager-error', 'Введите корректный email.');
      return;
    }
    if (!password || password.length < 6) {
      showError('add-manager-error', 'Пароль должен быть не короче 6 символов.');
      return;
    }

    btn.disabled = true; btn.textContent = 'Создание…';

    try {
      const { data: { session } } = await sb.auth.getSession();
      if (!session) throw new Error('Сессия истекла. Войдите заново.');

      const response = await fetch(CREATE_TRANSLATOR_URL, {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${session.access_token}`,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          email, password, name,
          role: 'manager',
          timezone: 'Asia/Tashkent',
        }),
      });

      const result = await response.json();
      if (!response.ok || !result.ok) {
        throw new Error(result.error || 'Не удалось создать менеджера');
      }

      closeAddManagerModal();
      await loadManagers();
    } catch (e) {
      showError('add-manager-error', e.message);
    } finally {
      btn.disabled = false; btn.textContent = 'Создать менеджера';
    }
  }

  // ──── Активация/деактивация менеджера ───────────────────────────────
  async function toggleManagerActive(userId, makeActive, managerName) {
    const action = makeActive ? 'активировать' : 'деактивировать';
    if (!confirm(`Точно ${action} менеджера ${managerName}?`)) return;

    try {
      const { error } = await sb.from('users')
        .update({ is_active: makeActive })
        .eq('id', userId);
      if (error) throw new Error(error.message);
      await loadManagers();
    } catch (e) {
      alert('Ошибка: ' + e.message);
    }
  }

  // Выдать/убрать доступ к секции «Клиенты» (только super-admin)
  async function toggleClientsAccess(userId, grant, managerName) {
    const action = grant ? 'дать доступ к клиентам' : 'убрать доступ к клиентам';
    if (!confirm(`Точно ${action} для ${managerName}?`)) return;

    try {
      const { error } = await sb.from('users')
        .update({ can_access_clients: grant })
        .eq('id', userId);
      if (error) throw new Error(error.message);
      await loadManagers();
    } catch (e) {
      alert('Ошибка: ' + e.message);
    }
  }


  async function loadTranslators() {
    const content = document.getElementById('translators-content');
    content.innerHTML = '<div class="loading-state">Загрузка переводчиков…</div>';

    const { data, error } = await sb.from('users')
      .select(`id, email, name, timezone, is_active, created_at,
               default_shift_start, default_shift_end, default_shift_minutes,
               translator_pairs ( id, rate_per_hour, is_primary, proficiency,
                                  language_pairs (code, display_name) )`)
      .eq('role', 'translator')
      .order('created_at', { ascending: false });

    if (error) {
      content.innerHTML = '';
      showError('translators-error', 'Ошибка загрузки: ' + error.message);
      return;
    }

    document.getElementById('translators-count').textContent =
      `${data.length} ${pluralize(data.length, 'переводчик', 'переводчика', 'переводчиков')} в системе`;

    if (data.length === 0) {
      content.innerHTML = `
        <div class="empty-state">
          <div class="empty-state-title">Переводчиков пока нет</div>
          <div class="empty-state-text">Нажмите «+ Добавить переводчика», чтобы создать первую учётную запись.</div>
        </div>`;
      return;
    }

    let html = `<table><thead><tr>
      <th>Переводчик</th><th>Языковые пары</th>
      <th class="numeric">Часов в этом мес.</th>
      <th class="numeric">К выплате</th>
      <th>Статус</th>
      <th></th>
    </tr></thead><tbody>`;

    // Получим часы каждого переводчика за текущий месяц одним запросом
    const monthStart = formatDate(new Date(new Date().getFullYear(), new Date().getMonth(), 1));
    const monthEnd = formatDate(new Date(new Date().getFullYear(), new Date().getMonth() + 1, 0));
    const userIds = data.map(u => u.id);
    let daysData = [];
    if (userIds.length > 0) {
      const { data: monthDays } = await sb
        .from('work_days')
        .select(`
          user_id, work_date, day_type,
          work_intervals ( language_pair_id, duration_minutes ),
          breaks ( duration_minutes )
        `)
        .in('user_id', userIds)
        .gte('work_date', monthStart)
        .lte('work_date', monthEnd);
      daysData = monthDays || [];
    }

    // Считаем для каждого переводчика
    const statsByUser = {}; // user_id -> { minutes, amount }
    for (const u of data) statsByUser[u.id] = { minutes: 0, amount: 0 };

    // Подгружаем историю ставок для всех переводчиков
    const allHistoryByUser = await loadRateHistoryForUsers(userIds);

    // O(1) поиск пользователя по id вместо .find() — ускоряет dashboard при многих переводчиках
    const usersById = new Map(data.map(u => [u.id, u]));
    // Преcчитываем rateByPair один раз на пользователя, а не на каждый день
    const rateByPairByUser = new Map();
    for (const u of data) {
      const rateByPair = {};
      for (const p of (u.translator_pairs || [])) {
        rateByPair[p.language_pair_id] = Number(p.rate_per_hour);
      }
      rateByPairByUser.set(u.id, rateByPair);
    }

    for (const d of daysData) {
      if (d.day_type !== 'working') continue;
      const user = usersById.get(d.user_id);
      if (!user) continue;
      const rateByPair = rateByPairByUser.get(d.user_id);
      const userHistory = allHistoryByUser[d.user_id] || {};
      const calc = calcDayWithHistory(d, userHistory, rateByPair);
      statsByUser[d.user_id].minutes += calc.netMinutes;
      statsByUser[d.user_id].amount += calc.amount;
    }

    for (const u of data) {
      const initials = (u.name || '?').split(' ').map(s => s[0]).join('').slice(0, 2).toUpperCase();
      const pairs = (u.translator_pairs || []).sort((a, b) => (b.is_primary ? 1 : 0) - (a.is_primary ? 1 : 0));
      const pairsHtml = pairs.length === 0
        ? '<span style="color:#94A3B8;font-size:12px;">нет пар</span>'
        : pairs.map(p => `
            <div class="pair-row">
              <span class="badge badge-neutral">${p.language_pairs.code}</span>
              <span class="pair-rate">$${Number(p.rate_per_hour).toFixed(2)}/ч</span>
              ${p.is_primary ? '<span style="color:#94A3B8;font-size:11px;">основная</span>' : ''}
            </div>`).join('');
      const statusBadge = u.is_active
        ? '<span class="badge badge-good">● Активен</span>'
        : '<span class="badge badge-warn">● Неактивен</span>';

      const stats = statsByUser[u.id];
      const hoursDisplay = stats.minutes > 0
        ? formatHoursMinutes(stats.minutes)
        : '<span style="color:#94A3B8;">—</span>';
      const amountDisplay = stats.amount > 0
        ? '$' + stats.amount.toFixed(2)
        : '<span style="color:#94A3B8;">—</span>';

      html += `<tr style="cursor:pointer;" data-user-id="${u.id}" onclick="openTranslatorDetail('${u.id}')">
        <td><div class="emp-cell">
          <div class="emp-avatar">${initials}</div>
          <div><div class="emp-name">${escapeHtml(u.name)}</div>
               <div class="emp-email">${escapeHtml(u.email)}</div></div>
        </div></td>
        <td><div class="pair-list">${pairsHtml}</div></td>
        <td style="text-align:right; font-family: 'JetBrains Mono', monospace; font-weight: 500;">${hoursDisplay}</td>
        <td style="text-align:right; font-family: 'JetBrains Mono', monospace; font-weight: 600; color: #16A34A;">${amountDisplay}</td>
        <td>${statusBadge}</td>
        <td style="text-align: right;"><span style="color:#94A3B8; font-size: 11px;">Открыть →</span></td>
      </tr>`;
    }
    html += '</tbody></table>';
    content.innerHTML = html;
  }

  // ====================================================================
  // ДАШБОРД (для менеджера)
  // ====================================================================
  async function loadDashboard() {
    hideError('dash-error');

    // Текущий период
    const today = new Date();
    const year = today.getFullYear();
    const month = today.getMonth();
    const monthLabel = MONTH_NAMES_RU[month] + ' ' + year;

    // Дата справа
    const dateStr = today.toLocaleDateString('ru-RU', {
      weekday: 'long', day: 'numeric', month: 'long'
    });
    document.getElementById('dash-date').textContent = dateStr;

    // Заголовок — месяц
    document.getElementById('dash-title').textContent = monthLabel;

    // Сводку наполним позже, после загрузки данных
    document.getElementById('dash-summary').innerHTML = 'Загрузка данных…';

    document.getElementById('dash-chart-month').textContent = monthLabel;

    const periodStart = formatDate(new Date(year, month, 1));
    const periodEnd = formatDate(new Date(year, month + 1, 0));

    // ─── Загружаем активных переводчиков и их пары ─────────────────────────
    const { data: users, error: usersError } = await sb
      .from('users')
      .select(`
        id, name, is_active, default_shift_minutes,
        translator_pairs ( language_pair_id, rate_per_hour,
                            language_pairs (code) )
      `)
      .eq('role', 'translator');

    if (usersError) {
      showError('dash-error', 'Ошибка загрузки: ' + usersError.message);
      return;
    }

    const activeUsers = (users || []).filter(u => u.is_active);
    const inactiveCount = (users || []).length - activeUsers.length;

    // ─── Загружаем все дни месяца по всем переводчикам ─────────────────────
    const userIds = activeUsers.map(u => u.id);
    let allDays = [];
    if (userIds.length > 0) {
      const { data: days } = await sb
        .from('work_days')
        .select(`
          user_id, work_date, day_type,
          work_intervals ( duration_minutes, language_pair_id ),
          breaks ( duration_minutes )
        `)
        .in('user_id', userIds)
        .gte('work_date', periodStart)
        .lte('work_date', periodEnd);
      allDays = days || [];
    }

    // ─── История ставок ────────────────────────────────────────────────────
    const allHistoryByUser = await loadRateHistoryForUsers(userIds);

    // ─── Считаем глобальные KPI и распределение ────────────────────────────
    let totalMinutes = 0;
    let totalAmount = 0;
    const minutesByUser = {}; // user_id -> minutes
    const minutesByPair = {}; // code -> minutes
    const minutesByDay = {};  // date -> minutes
    const codeByPair = {};

    // O(1) поиск пользователя и преcчитанные rateByPair — на dashboard с 30+ переводчиками
    // экономия времени измеряется секундами.
    const usersById = new Map(activeUsers.map(u => [u.id, u]));
    const rateByPairByUser = new Map();
    for (const u of activeUsers) {
      minutesByUser[u.id] = 0;
      const rateByPair = {};
      for (const p of (u.translator_pairs || [])) {
        rateByPair[p.language_pair_id] = Number(p.rate_per_hour);
        codeByPair[p.language_pair_id] = p.language_pairs.code;
      }
      rateByPairByUser.set(u.id, rateByPair);
    }

    for (const d of allDays) {
      if (d.day_type !== 'working') continue;
      const user = usersById.get(d.user_id);
      if (!user) continue;

      const rateByPair = rateByPairByUser.get(d.user_id);
      const userHistory = allHistoryByUser[d.user_id] || {};

      const intervalsSum = (d.work_intervals || []).reduce((s, i) => s + (i.duration_minutes || 0), 0);
      const breaksSum = (d.breaks || []).reduce((s, b) => s + (b.duration_minutes || 0), 0);
      const netMin = Math.max(0, intervalsSum - breaksSum);
      if (netMin <= 0) continue;

      const breakRatio = intervalsSum > 0 ? (netMin / intervalsSum) : 0;
      let dayAmount = 0;
      for (const interv of (d.work_intervals || [])) {
        const rate = getRateForDate(d.work_date, interv.language_pair_id, userHistory, rateByPair);
        const intervalNetMin = (interv.duration_minutes || 0) * breakRatio;
        dayAmount += (intervalNetMin / 60) * rate;
        const code = codeByPair[interv.language_pair_id] || '—';
        minutesByPair[code] = (minutesByPair[code] || 0) + intervalNetMin;
      }

      totalMinutes += netMin;
      totalAmount += dayAmount;
      minutesByUser[d.user_id] += netMin;
      minutesByDay[d.work_date] = (minutesByDay[d.work_date] || 0) + netMin;
    }

    // ─── Считаем овертайм команды ────────────────────────────────────────
    const allShiftsByUser = await loadShiftsForUsers(
      activeUsers.map(u => u.id), periodStart, periodEnd
    );
    let totalOvertime = 0;
    const overtimeByUser = {};
    for (const u of activeUsers) overtimeByUser[u.id] = 0;

    for (const d of allDays) {
      if (d.day_type !== 'working') continue;
      const user = usersById.get(d.user_id);
      if (!user) continue;

      const intervalsSum = (d.work_intervals || []).reduce((s, i) => s + (i.duration_minutes || 0), 0);
      const breaksSum = (d.breaks || []).reduce((s, b) => s + (b.duration_minutes || 0), 0);
      const netMin = Math.max(0, intervalsSum - breaksSum);
      if (netMin <= 0) continue;

      const userShifts = allShiftsByUser[d.user_id] || {};
      const defaultMin = user.default_shift_minutes || 480;
      const planned = getPlannedMinutes(d.work_date, userShifts, defaultMin);
      // Овертайм = присутствие (gross) − план. Брейк нейтрален (он и так не оплачивается).
      const ot = intervalsSum - planned;
      totalOvertime += ot;
      overtimeByUser[d.user_id] += ot;
    }

    // ─── Заполняем KPI ─────────────────────────────────────────────────────
    const overtimeEl = document.getElementById('dash-kpi-overtime');
    const overtimeMetaEl = document.getElementById('dash-kpi-overtime-meta');
    if (Math.abs(totalOvertime) < 1) {
      overtimeEl.innerHTML = '0<span class="kpi-unit">мин</span>';
      overtimeEl.style.color = '';
      overtimeMetaEl.textContent = 'команда точно по плану';
    } else if (totalOvertime > 0) {
      overtimeEl.textContent = '+' + formatHoursMinutes(totalOvertime);
      overtimeEl.style.color = '#B45309';
      // Топ овертаймщик
      const topUserId = Object.keys(overtimeByUser).sort((a, b) => overtimeByUser[b] - overtimeByUser[a])[0];
      const topUser = usersById.get(topUserId);
      if (topUser && overtimeByUser[topUserId] > 0) {
        overtimeMetaEl.textContent = `больше всех: ${topUser.name}`;
      } else {
        overtimeMetaEl.textContent = 'переработано командой';
      }
    } else {
      overtimeEl.textContent = '−' + formatHoursMinutes(Math.abs(totalOvertime));
      overtimeEl.style.color = '#1E40AF';
      overtimeMetaEl.textContent = 'недоработано командой';
    }

    document.getElementById('dash-kpi-hours').textContent = formatHoursMinutes(totalMinutes);
    const workingTranslators = activeUsers.filter(u => minutesByUser[u.id] > 0).length;
    document.getElementById('dash-kpi-hours-meta').textContent =
      workingTranslators > 0
        ? `${workingTranslators} ${pluralize(workingTranslators, 'работал', 'работали', 'работали')} в мес.`
        : 'никто не работал';

    // Дней до закрытия периода
    const lastDay = new Date(year, month + 1, 0);
    const daysLeft = Math.max(0, Math.ceil((lastDay - today) / (1000 * 60 * 60 * 24)));
    document.getElementById('dash-kpi-days').textContent = daysLeft;
    const daysMeta = document.getElementById('dash-kpi-days-meta');
    if (daysLeft === 0) {
      daysMeta.textContent = 'период завершён';
      daysMeta.className = 'dash-kpi-meta good';
    } else if (daysLeft <= 3) {
      daysMeta.textContent = 'скоро закрытие';
      daysMeta.className = 'dash-kpi-meta warning';
    } else {
      daysMeta.textContent = pluralize(daysLeft, 'день в текущем мес.', 'дня в текущем мес.', 'дней в текущем мес.');
      daysMeta.className = 'dash-kpi-meta';
    }

    document.getElementById('dash-kpi-amount').textContent = '$' + totalAmount.toFixed(2);

    // ─── Сводка-предложение в заголовке ────────────────────────────────────
    const summaryParts = [];
    if (workingTranslators > 0) {
      summaryParts.push(`<strong>$${totalAmount.toFixed(2)}</strong> заработано командой`);
      summaryParts.push(`<span class="accent-blue">${formatHoursMinutes(totalMinutes)}</span> отработано`);
    } else {
      summaryParts.push('Пока никто не работал в этом месяце');
    }
    if (daysLeft > 0 && daysLeft <= 5) {
      summaryParts.push(`до закрытия периода ${daysLeft} ${pluralize(daysLeft, 'день', 'дня', 'дней')}`);
    } else if (daysLeft === 0) {
      summaryParts.push('период завершён, можно закрывать');
    }
    document.getElementById('dash-summary').innerHTML = summaryParts.join(' · ');

    // ─── График по дням ────────────────────────────────────────────────────
    renderDashChart(year, month, minutesByDay);

    // ─── Топ переводчиков ──────────────────────────────────────────────────
    renderDashLeaderboard(activeUsers, minutesByUser);

    // ─── Распределение по парам ────────────────────────────────────────────
    renderDashPairs(minutesByPair);
  }

  function renderDashChart(year, month, minutesByDay) {
    const grid = document.getElementById('dash-chart');
    const axis = document.getElementById('dash-chart-axis');

    const daysInMonth = new Date(year, month + 1, 0).getDate();
    const todayStr = formatDate(new Date());
    const maxMin = Math.max(...Object.values(minutesByDay), 1);

    const gridFrag = document.createDocumentFragment();
    for (let d = 1; d <= daysInMonth; d++) {
      const dateStr = `${year}-${String(month + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
      const min = minutesByDay[dateStr] || 0;
      const dayOfWeek = new Date(year, month, d).getDay();
      const isWeekend = (dayOfWeek === 0 || dayOfWeek === 6);

      const bar = document.createElement('div');
      bar.className = 'dash-chart-bar';
      if (min === 0) bar.classList.add('empty');
      else if (isWeekend) bar.classList.add('weekend');
      if (dateStr === todayStr) bar.classList.add('today');

      const heightPct = (min / maxMin) * 100;
      bar.style.height = Math.max(2, heightPct) + '%';

      if (min > 0) {
        bar.dataset.tooltip = `${d} ${MONTH_NAMES_RU[month].toLowerCase()}: ${formatHoursMinutes(min)}`;
      } else {
        bar.dataset.tooltip = `${d}: нет часов`;
      }
      gridFrag.appendChild(bar);
    }
    grid.replaceChildren(gridFrag);

    // Ось снизу — каждые 5 дней
    const axisFrag = document.createDocumentFragment();
    for (let d = 1; d <= daysInMonth; d += 5) {
      const lbl = document.createElement('span');
      lbl.textContent = d;
      lbl.style.flex = d === 1 ? '0 0 auto' : '1 1 auto';
      axisFrag.appendChild(lbl);
    }
    const lastLbl = document.createElement('span');
    lastLbl.textContent = daysInMonth;
    axisFrag.appendChild(lastLbl);
    axis.replaceChildren(axisFrag);
  }

  function renderDashLeaderboard(users, minutesByUser) {
    const block = document.getElementById('dash-leaderboard');
    block.innerHTML = '';
    const sorted = users
      .map(u => ({ user: u, minutes: minutesByUser[u.id] || 0 }))
      .filter(x => x.minutes > 0)
      .sort((a, b) => b.minutes - a.minutes)
      .slice(0, 5);

    if (sorted.length === 0) {
      block.innerHTML = '<div class="dash-empty">Никто не работал в этом месяце.</div>';
      return;
    }

    const maxMin = sorted[0].minutes;
    sorted.forEach((x, idx) => {
      const initials = (x.user.name || '?').split(' ').map(s => s[0]).join('').slice(0, 2).toUpperCase();
      const widthPct = (x.minutes / maxMin) * 100;
      const row = document.createElement('div');
      row.className = 'dash-leader-row';
      row.style.cursor = 'pointer';
      row.onclick = () => {
        showPage('translator-detail');
        openTranslatorDetail(x.user.id);
      };
      row.innerHTML = `
        <div class="dash-leader-rank ${idx === 0 ? 'gold' : ''}">${idx + 1}</div>
        <div class="dash-leader-info">
          <div class="dash-leader-name">${escapeHtml(x.user.name)}</div>
          <div class="dash-leader-bar-track">
            <div class="dash-leader-bar-fill" style="width: ${widthPct}%;"></div>
          </div>
        </div>
        <div class="dash-leader-value">${formatHoursMinutes(x.minutes)}</div>
      `;
      block.appendChild(row);
    });
  }

  function renderDashPairs(minutesByPair) {
    const block = document.getElementById('dash-pairs');
    block.innerHTML = '';
    const sorted = Object.entries(minutesByPair)
      .filter(([_, m]) => m > 0)
      .sort((a, b) => b[1] - a[1]);

    if (sorted.length === 0) {
      block.innerHTML = '<div class="dash-empty">Нет данных за этот месяц.</div>';
      return;
    }

    const maxMin = sorted[0][1];
    sorted.forEach(([code, min]) => {
      const widthPct = (min / maxMin) * 100;
      const row = document.createElement('div');
      row.className = 'dash-pair-row';
      row.innerHTML = `
        <div class="dash-pair-code">${escapeHtml(code)}</div>
        <div class="dash-pair-bar-track">
          <div class="dash-pair-bar-fill" style="width: ${widthPct}%;"></div>
        </div>
        <div class="dash-pair-value">${formatHoursMinutes(Math.round(min))}</div>
      `;
      block.appendChild(row);
    });
  }

  // ====================================================================
  // СТРАНИЦА: ЗАРПЛАТА (для менеджера)
  // ====================================================================
  // Состояние страницы зарплаты
  let payrollYear = new Date().getFullYear();
  let payrollMonth = new Date().getMonth();
  let currentPeriod = null; // запись из payroll_periods
  let currentPeriodEntries = []; // payroll_entries для закрытого периода
  let currentPayrollRows = []; // {user, totalMinutes, totalAmount, breakdown} — для экспорта

  function payrollMonthLabel(y, m) {
    return MONTH_NAMES_RU[m] + ' ' + y;
  }
  function payrollLabel(y, m) {
    // YYYY-MM для period_label
    return `${y}-${String(m + 1).padStart(2, '0')}`;
  }

  async function loadPayroll() {
    hideError('payroll-error');

    // Заполнить селектор (один раз достаточно, но безопасно)
    fillPayrollMonthSelector();

    document.getElementById('payroll-subtitle').textContent =
      'Ведомость за ' + payrollMonthLabel(payrollYear, payrollMonth);

    // Загружаем период (если существует) и его статус
    const label = payrollLabel(payrollYear, payrollMonth);
    const { data: period } = await sb
      .from('payroll_periods')
      .select('id, period_label, period_start, period_end, status, closed_at, closed_by')
      .eq('period_label', label)
      .maybeSingle();

    currentPeriod = period || null;
    currentPeriodEntries = [];

    // Если период закрыт — берём данные из payroll_entries (снимок)
    if (period && period.status === 'closed') {
      const { data: entries } = await sb
        .from('payroll_entries')
        .select('user_id, total_hours, total_amount, breakdown_by_pair')
        .eq('period_id', period.id);
      currentPeriodEntries = entries || [];
      renderPayrollClosed(period);
    } else {
      renderPayrollOpen(period);
    }

    await renderPayrollContent();
  }

  function fillPayrollMonthSelector() {
    const sel = document.getElementById('payroll-month-select');
    if (sel.options.length > 0) {
      sel.value = `${payrollYear}-${payrollMonth}`;
      return;
    }
    const now = new Date();
    for (let i = 0; i < 12; i++) {
      const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
      const opt = document.createElement('option');
      opt.value = `${d.getFullYear()}-${d.getMonth()}`;
      opt.textContent = payrollMonthLabel(d.getFullYear(), d.getMonth());
      sel.appendChild(opt);
    }
    sel.value = `${payrollYear}-${payrollMonth}`;
    sel.onchange = () => {
      const [y, m] = sel.value.split('-').map(Number);
      payrollYear = y; payrollMonth = m;
      loadPayroll();
    };
  }

  function renderPayrollOpen(period) {
    const block = document.getElementById('payroll-status-block');
    block.className = 'payroll-status status-open';
    block.classList.remove('hidden');
    block.innerHTML = `
      <span class="payroll-status-icon"></span>
      Период <strong>открыт</strong>. Данные рассчитываются на лету по текущим часам и ставкам.
    `;
    document.getElementById('payroll-snapshot-note').classList.add('hidden');

    // По умолчанию скрываем кнопку — покажем после расчёта, если есть данные
    document.getElementById('btn-close-period').classList.add('hidden');
    document.getElementById('btn-reopen-period').classList.add('hidden');
  }

  // Решение, показывать ли кнопку «Закрыть период» — после того как данные посчитаны.
  // Условия:
  //   1) месяц должен быть полностью завершён (today > последний день месяца)
  //   2) в периоде есть хотя бы один переводчик с часами
  function updateCloseButtonVisibility(hasData) {
    if (currentPeriod && currentPeriod.status === 'closed') return;
    const today = new Date();
    const periodEnd = new Date(payrollYear, payrollMonth + 1, 0);
    const monthFinished = today > periodEnd;
    const btn = document.getElementById('btn-close-period');
    btn.classList.toggle('hidden', !(monthFinished && hasData));
  }

  function renderPayrollClosed(period) {
    const block = document.getElementById('payroll-status-block');
    block.className = 'payroll-status status-closed';
    block.classList.remove('hidden');
    const closedAt = period.closed_at
      ? new Date(period.closed_at).toLocaleDateString('ru-RU')
      : '—';
    block.innerHTML = `
      <span class="payroll-status-icon"></span>
      Период <strong>закрыт</strong> ${closedAt}. Данные ниже — зафиксированный снимок.
    `;
    document.getElementById('payroll-snapshot-note').classList.remove('hidden');

    // Кнопки
    document.getElementById('btn-close-period').classList.add('hidden');
    // Открыть обратно — только если прошло ≤ 30 дней с закрытия
    const closedDate = period.closed_at ? new Date(period.closed_at) : null;
    const daysSinceClose = closedDate
      ? Math.floor((Date.now() - closedDate.getTime()) / (1000 * 60 * 60 * 24))
      : 999;
    const canReopen = daysSinceClose <= 30;
    document.getElementById('btn-reopen-period').classList.toggle('hidden', !canReopen);
  }

  async function renderPayrollContent() {
    const content = document.getElementById('payroll-content');
    content.innerHTML = '<div class="loading-state">Расчёт…</div>';

    // Если период закрыт — рендерим из payroll_entries
    if (currentPeriod && currentPeriod.status === 'closed') {
      await renderPayrollFromSnapshot();
      return;
    }

    // Иначе считаем на лету
    await renderPayrollLive();
  }

  // Расчёт «на лету» — серверная RPC payroll_for_period.
  // SQL-функция считает gross/net/breaks/ставки-на-дату/овертайм одним запросом,
  // что экономит ~5 запросов на страницу и десятки KB трафика по сравнению со
  // старой JS-реализацией.
  async function renderPayrollLive() {
    const periodStart = formatDate(new Date(payrollYear, payrollMonth, 1));
    const periodEnd   = formatDate(new Date(payrollYear, payrollMonth + 1, 0));

    const { data, error } = await sb.rpc('payroll_for_period', {
      p_period_start: periodStart,
      p_period_end:   periodEnd,
    });

    if (error) {
      showError('payroll-error', 'Ошибка расчёта: ' + error.message);
      return;
    }

    // RPC возвращает массив объектов в форме, ожидаемой renderPayrollTable.
    // Числа на всякий случай приводим к Number — Supabase может вернуть строки
    // для NUMERIC-полей при большой точности.
    const result = (data || []).map(row => ({
      user: row.user,
      totalMinutes:         row.totalMinutes,
      totalAmount:          Number(row.totalAmount),
      totalPlannedMinutes:  row.totalPlannedMinutes,
      totalOvertimeMinutes: row.totalOvertimeMinutes,
      breakdown: Object.fromEntries(
        Object.entries(row.breakdown || {}).map(([code, b]) => [code, {
          minutes: b.minutes,
          amount:  Number(b.amount),
          rate:    Number(b.rate),
        }])
      ),
    }));

    renderPayrollTable(result);
  }

  // Рендер из снимка (закрытый период)
  async function renderPayrollFromSnapshot() {
    // Нужны имена переводчиков → запрос
    const userIds = currentPeriodEntries.map(e => e.user_id);
    if (userIds.length === 0) {
      renderPayrollEmpty('Период закрыт, но в нём нет записей.');
      return;
    }
    const { data: users } = await sb
      .from('users')
      .select('id, name, email')
      .in('id', userIds);
    const usersById = {};
    for (const u of (users || [])) usersById[u.id] = u;

    const result = currentPeriodEntries.map(e => {
      const breakdown = {};
      const breakdownArr = e.breakdown_by_pair || [];
      for (const b of breakdownArr) {
        breakdown[b.pair_code] = {
          minutes: Math.round((Number(b.hours) || 0) * 60),
          rate: Number(b.rate) || 0,
          amount: Number(b.amount) || 0,
        };
      }
      return {
        user: usersById[e.user_id] || { name: '—', email: '' },
        totalMinutes: Math.round((Number(e.total_hours) || 0) * 60),
        totalAmount: Number(e.total_amount) || 0,
        breakdown,
      };
    });
    renderPayrollTable(result);
  }

  function renderPayrollEmpty(msg) {
    const content = document.getElementById('payroll-content');
    content.innerHTML = `<div class="empty-state">
      <div class="empty-state-text">${escapeHtml(msg)}</div>
    </div>`;
    document.getElementById('payroll-kpi-count').textContent = '0';
    document.getElementById('payroll-kpi-hours').textContent = '0ч 00м';
    document.getElementById('payroll-kpi-amount').textContent = '$0.00';
    updateCloseButtonVisibility(false);
    document.getElementById('btn-export-excel').classList.add('hidden');
    document.getElementById('btn-export-pdf').classList.add('hidden');
    currentPayrollRows = [];
  }

  function renderPayrollTable(rows) {
    const content = document.getElementById('payroll-content');
    currentPayrollRows = rows; // сохраняем для экспорта

    if (rows.length === 0) {
      renderPayrollEmpty('Никто не работал в этом месяце.');
      return;
    }

    // Кнопка «Закрыть период» — показать только если есть данные и период не закрыт
    updateCloseButtonVisibility(true);
    // Кнопки экспорта — показать когда есть данные
    document.getElementById('btn-export-excel').classList.remove('hidden');
    document.getElementById('btn-export-pdf').classList.remove('hidden');

    // KPI
    const totalAmount = rows.reduce((s, r) => s + r.totalAmount, 0);
    const totalMinutes = rows.reduce((s, r) => s + r.totalMinutes, 0);
    const totalOvertimeMinutes = rows.reduce((s, r) => s + (r.totalOvertimeMinutes || 0), 0);
    document.getElementById('payroll-kpi-count').textContent = String(rows.length);
    document.getElementById('payroll-kpi-hours').textContent = formatHoursMinutes(totalMinutes);
    document.getElementById('payroll-kpi-amount').textContent = '$' + totalAmount.toFixed(2);

    // Таблица
    let html = `<table><thead><tr>
      <th>Переводчик</th>
      <th>Разбивка по парам</th>
      <th style="text-align:right;">Часы</th>
      <th style="text-align:right;">Овертайм</th>
      <th style="text-align:right;">К выплате</th>
    </tr></thead><tbody>`;

    for (const r of rows) {
      const initials = (r.user.name || '?').split(' ').map(s => s[0]).join('').slice(0, 2).toUpperCase();
      const breakdownHtml = Object.entries(r.breakdown)
        .sort()
        .map(([code, b]) => `
          <div class="payroll-pair-row">
            <span class="badge badge-neutral">${code}</span>
            <span class="payroll-pair-hours">${formatHoursMinutes(Math.round(b.minutes))}</span>
            <span class="payroll-pair-rate">× $${b.rate.toFixed(2)}/ч</span>
            <span class="payroll-pair-amount">= $${b.amount.toFixed(2)}</span>
          </div>
        `).join('');

      // Виджет овертайма
      const ot = r.totalOvertimeMinutes || 0;
      let otHtml;
      if (Math.abs(ot) < 1) {
        otHtml = `<span style="color: #94A3B8; font-family: 'JetBrains Mono', monospace; font-size: 12px;">точно</span>`;
      } else if (ot > 0) {
        otHtml = `<span class="overtime-widget over">+${formatHoursMinutes(ot)}</span>`;
      } else {
        otHtml = `<span class="overtime-widget under">−${formatHoursMinutes(Math.abs(ot))}</span>`;
      }

      html += `<tr>
        <td><div class="emp-cell">
          <div class="emp-avatar">${initials}</div>
          <div><div class="emp-name">${escapeHtml(r.user.name)}</div>
               <div class="emp-email">${escapeHtml(r.user.email)}</div></div>
        </div></td>
        <td>${breakdownHtml}</td>
        <td style="text-align:right; font-family: 'JetBrains Mono', monospace; font-weight: 500;">${formatHoursMinutes(r.totalMinutes)}</td>
        <td style="text-align:right;">${otHtml}</td>
        <td style="text-align:right; font-family: 'JetBrains Mono', monospace; font-weight: 600; color: #16A34A;">$${r.totalAmount.toFixed(2)}</td>
      </tr>`;
    }

    // Итог
    let totalOtHtml;
    if (Math.abs(totalOvertimeMinutes) < 1) {
      totalOtHtml = `<span style="color: rgba(255,255,255,0.5); font-family: 'JetBrains Mono', monospace; font-size: 12px;">точно</span>`;
    } else if (totalOvertimeMinutes > 0) {
      totalOtHtml = `<span style="color: #FBBF24; font-family: 'JetBrains Mono', monospace; font-weight: 600;">+${formatHoursMinutes(totalOvertimeMinutes)}</span>`;
    } else {
      totalOtHtml = `<span style="color: #93C5FD; font-family: 'JetBrains Mono', monospace; font-weight: 600;">−${formatHoursMinutes(Math.abs(totalOvertimeMinutes))}</span>`;
    }

    html += `<tr class="payroll-total-row">
      <td colspan="2" style="text-align: right; color: #94A3B8; text-transform: uppercase; font-family: 'JetBrains Mono', monospace; font-size: 11px; letter-spacing: 0.1em;">Итого по команде</td>
      <td style="text-align:right; font-family: 'JetBrains Mono', monospace;">${formatHoursMinutes(totalMinutes)}</td>
      <td style="text-align:right;">${totalOtHtml}</td>
      <td style="text-align:right;" class="payroll-grand-total">$${totalAmount.toFixed(2)}</td>
    </tr>`;

    html += '</tbody></table>';
    content.innerHTML = html;
  }

  // ──── Закрытие/открытие периода ────────────────────────────────────────
  async function closePeriod() {
    if (!confirm('Закрыть период ' + payrollMonthLabel(payrollYear, payrollMonth) +
                 '?\n\nПосле закрытия данные за этот месяц станут зафиксированными. ' +
                 'Изменения часов или ставок не повлияют на расчёт.')) return;

    const btn = document.getElementById('btn-close-period');
    btn.disabled = true; btn.textContent = 'Закрываем…';

    try {
      const label = payrollLabel(payrollYear, payrollMonth);
      const periodStart = formatDate(new Date(payrollYear, payrollMonth, 1));
      const periodEnd = formatDate(new Date(payrollYear, payrollMonth + 1, 0));

      // 1) Создаём или обновляем запись периода
      let periodId = currentPeriod?.id;
      if (!periodId) {
        const { data: newPeriod, error: pErr } = await sb
          .from('payroll_periods')
          .insert({
            period_label: label,
            period_start: periodStart,
            period_end: periodEnd,
            status: 'open',
          })
          .select('id').single();
        if (pErr) throw new Error('Создание периода: ' + pErr.message);
        periodId = newPeriod.id;
      }

      // 2) Рассчитываем актуальную ведомость через ту же RPC, что использует
      //    страница «Зарплата». Гарантия: цифры в payroll_entries совпадут с
      //    тем, что менеджер видел в момент закрытия.
      const { data: payrollRows, error: rpcErr } = await sb.rpc('payroll_for_period', {
        p_period_start: periodStart,
        p_period_end:   periodEnd,
      });
      if (rpcErr) throw new Error('Расчёт ведомости: ' + rpcErr.message);

      const entriesToInsert = (payrollRows || []).map(row => ({
        period_id: periodId,
        user_id: row.user.id,
        total_hours:  Number((row.totalMinutes / 60).toFixed(2)),
        total_amount: Number(Number(row.totalAmount).toFixed(2)),
        breakdown_by_pair: Object.entries(row.breakdown || {}).map(([code, b]) => ({
          pair_code: code,
          hours:  Number((b.minutes / 60).toFixed(2)),
          rate:   Number(b.rate),
          amount: Number(Number(b.amount).toFixed(2)),
        })),
        analytics_breakdown: null, // в этап 8 наполним
      }));

      // 3) Удалим старые записи (на случай повторного закрытия) и вставим новые
      await sb.from('payroll_entries').delete().eq('period_id', periodId);
      if (entriesToInsert.length > 0) {
        const { error: eErr } = await sb.from('payroll_entries').insert(entriesToInsert);
        if (eErr) throw new Error('Сохранение записей: ' + eErr.message);
      }

      // 4) Меняем статус периода на closed
      const { error: updErr } = await sb
        .from('payroll_periods')
        .update({
          status: 'closed',
          closed_at: new Date().toISOString(),
          closed_by: currentUser.id,
        })
        .eq('id', periodId);
      if (updErr) throw new Error('Обновление статуса: ' + updErr.message);

      // Перезагружаем
      await loadPayroll();
    } catch (e) {
      showError('payroll-error', e.message);
    } finally {
      btn.disabled = false; btn.textContent = 'Закрыть период';
    }
  }

  // ──── Экспорт ведомости ────────────────────────────────────────────────

  // Excel-файл с двумя листами:
  //   - "Сводка" — компактная ведомость
  //   - "Разбивка" — детально по парам
  async function exportPayrollExcel() {
    if (currentPayrollRows.length === 0) {
      alert('Нет данных для экспорта.');
      return;
    }

    // Ленивая загрузка XLSX — библиотека ~900 КБ, грузится только при первом клике
    const btn = document.getElementById('btn-export-excel');
    const originalText = btn.textContent;
    btn.disabled = true;
    btn.textContent = 'Загрузка библиотеки…';
    try {
      await ensureXlsx();
    } catch (e) {
      btn.disabled = false;
      btn.textContent = originalText;
      alert('Не удалось загрузить библиотеку XLSX: ' + e.message);
      return;
    }
    btn.disabled = false;
    btn.textContent = originalText;

    const monthLabel = payrollMonthLabel(payrollYear, payrollMonth);
    const isClosed = currentPeriod && currentPeriod.status === 'closed';
    const closedNote = isClosed
      ? `Период закрыт ${new Date(currentPeriod.closed_at).toLocaleDateString('ru-RU')}`
      : 'Период открыт (предварительный расчёт)';

    // ─── Лист 1: Сводка ────────────────────────────────────────────────
    const summaryData = [
      ['Ведомость зарплаты переводчиков'],
      [monthLabel],
      [closedNote],
      [`Сформирован: ${new Date().toLocaleString('ru-RU')}`],
      [],
      ['№', 'Переводчик', 'Email', 'Часы', 'Овертайм (ч)', 'К выплате ($)'],
    ];
    let totalMinutes = 0;
    let totalAmount = 0;
    let totalOvertime = 0;
    currentPayrollRows.forEach((r, i) => {
      const otHours = Number(((r.totalOvertimeMinutes || 0) / 60).toFixed(2));
      summaryData.push([
        i + 1,
        r.user.name || '—',
        r.user.email || '',
        Number((r.totalMinutes / 60).toFixed(2)),
        otHours,
        Number(r.totalAmount.toFixed(2)),
      ]);
      totalMinutes += r.totalMinutes;
      totalAmount += r.totalAmount;
      totalOvertime += (r.totalOvertimeMinutes || 0);
    });
    summaryData.push([]);
    summaryData.push(['', 'ИТОГО:', '',
      Number((totalMinutes / 60).toFixed(2)),
      Number((totalOvertime / 60).toFixed(2)),
      Number(totalAmount.toFixed(2))
    ]);

    const wsSummary = XLSX.utils.aoa_to_sheet(summaryData);
    wsSummary['!cols'] = [
      { wch: 4 },   // №
      { wch: 28 },  // имя
      { wch: 26 },  // email
      { wch: 10 },  // часы
      { wch: 12 },  // овертайм
      { wch: 14 },  // сумма
    ];
    // Объединим заголовок
    wsSummary['!merges'] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: 5 } },
      { s: { r: 1, c: 0 }, e: { r: 1, c: 5 } },
      { s: { r: 2, c: 0 }, e: { r: 2, c: 5 } },
      { s: { r: 3, c: 0 }, e: { r: 3, c: 5 } },
    ];

    // ─── Лист 2: Разбивка по парам ─────────────────────────────────────
    const breakdownData = [
      ['Разбивка по языковым парам'],
      [monthLabel],
      [],
      ['Переводчик', 'Email', 'Языковая пара', 'Часы', 'Ставка ($/ч)', 'Сумма ($)'],
    ];
    let detailTotalMinutes = 0;
    let detailTotalAmount = 0;
    for (const r of currentPayrollRows) {
      const pairs = Object.entries(r.breakdown).sort();
      pairs.forEach(([code, b], idx) => {
        breakdownData.push([
          idx === 0 ? r.user.name : '',
          idx === 0 ? r.user.email : '',
          code,
          Number((b.minutes / 60).toFixed(2)),
          Number(b.rate.toFixed(2)),
          Number(b.amount.toFixed(2)),
        ]);
        detailTotalMinutes += b.minutes;
        detailTotalAmount += b.amount;
      });
    }
    breakdownData.push([]);
    breakdownData.push(['', '', 'ИТОГО:',
      Number((detailTotalMinutes / 60).toFixed(2)),
      '',
      Number(detailTotalAmount.toFixed(2))
    ]);

    const wsBreakdown = XLSX.utils.aoa_to_sheet(breakdownData);
    wsBreakdown['!cols'] = [
      { wch: 28 }, { wch: 26 }, { wch: 14 }, { wch: 10 }, { wch: 14 }, { wch: 14 }
    ];
    wsBreakdown['!merges'] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: 5 } },
      { s: { r: 1, c: 0 }, e: { r: 1, c: 5 } },
    ];

    // ─── Собираем книгу и сохраняем ─────────────────────────────────────
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, wsSummary, 'Сводка');
    XLSX.utils.book_append_sheet(wb, wsBreakdown, 'Разбивка');

    const fname = `LinguaTime_${payrollLabel(payrollYear, payrollMonth)}.xlsx`;
    XLSX.writeFile(wb, fname);
  }

  // PDF-ведомость: одна страница с таблицей и итогом
  async function exportPayrollPDF() {
    if (currentPayrollRows.length === 0) {
      alert('Нет данных для экспорта.');
      return;
    }

    // Ленивая загрузка jsPDF — ~700 КБ, грузится только при первом клике
    const btn = document.getElementById('btn-export-pdf');
    const originalText = btn.textContent;
    btn.disabled = true;
    btn.textContent = 'Загрузка библиотеки…';
    try {
      await ensureJspdf();
    } catch (e) {
      btn.disabled = false;
      btn.textContent = originalText;
      alert('Не удалось загрузить библиотеку jsPDF: ' + e.message);
      return;
    }
    btn.disabled = false;
    btn.textContent = originalText;

    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });

    const monthLabel = payrollMonthLabel(payrollYear, payrollMonth);
    const isClosed = currentPeriod && currentPeriod.status === 'closed';
    const closedNote = isClosed
      ? `Period closed: ${new Date(currentPeriod.closed_at).toLocaleDateString('ru-RU')}`
      : 'Preliminary calculation (open period)';
    const generated = `Generated: ${new Date().toLocaleString('ru-RU')}`;

    // ВАЖНО: стандартный шрифт jsPDF не поддерживает кириллицу.
    // Используем транслитерацию для надёжности на любых системах.
    const titleEN = 'Payroll Report — Translators';
    const monthEN = transliterate(monthLabel);

    // Заголовок
    doc.setFontSize(16);
    doc.setFont('helvetica', 'bold');
    doc.text(titleEN, 14, 18);

    doc.setFontSize(11);
    doc.setFont('helvetica', 'normal');
    doc.text(monthEN, 14, 26);

    doc.setFontSize(9);
    doc.setTextColor(120);
    doc.text(closedNote, 14, 32);
    doc.text(generated, 14, 37);
    doc.setTextColor(0);

    // Таблица сводки
    let totalMinutes = 0;
    let totalAmount = 0;
    let totalOvertime = 0;
    const body = currentPayrollRows.map((r, i) => {
      totalMinutes += r.totalMinutes;
      totalAmount += r.totalAmount;
      const ot = r.totalOvertimeMinutes || 0;
      totalOvertime += ot;
      const otStr = Math.abs(ot) < 1 ? '—'
        : (ot > 0 ? '+' : '−') + formatHoursMinutes(Math.abs(ot));
      return [
        i + 1,
        transliterate(r.user.name || '—'),
        r.user.email || '',
        formatHoursMinutes(r.totalMinutes),
        otStr,
        '$' + r.totalAmount.toFixed(2),
      ];
    });

    const totalOtStr = Math.abs(totalOvertime) < 1 ? '—'
      : (totalOvertime > 0 ? '+' : '−') + formatHoursMinutes(Math.abs(totalOvertime));

    doc.autoTable({
      startY: 44,
      head: [['#', 'Translator', 'Email', 'Hours', 'Overtime', 'Amount']],
      body,
      foot: [['', transliterate('ИТОГО:'), '',
        formatHoursMinutes(totalMinutes),
        totalOtStr,
        '$' + totalAmount.toFixed(2)
      ]],
      theme: 'grid',
      headStyles: { fillColor: [26, 26, 23], textColor: 255, fontStyle: 'bold' },
      footStyles: { fillColor: [26, 26, 23], textColor: 255, fontStyle: 'bold' },
      styles: { fontSize: 9, cellPadding: 3 },
      columnStyles: {
        0: { halign: 'center', cellWidth: 10 },
        3: { halign: 'right' },
        4: { halign: 'right' },
        5: { halign: 'right' },
      },
    });

    // Подпись/футер
    const finalY = doc.lastAutoTable.finalY + 16;
    doc.setFontSize(9);
    doc.setTextColor(80);
    doc.text(`Translators count: ${currentPayrollRows.length}`, 14, finalY);
    doc.text(`Total hours: ${formatHoursMinutes(totalMinutes)}`, 14, finalY + 5);
    doc.text(`Total payout: $${totalAmount.toFixed(2)}`, 14, finalY + 10);

    doc.setFontSize(8);
    doc.setTextColor(160);
    doc.text('LinguaTime CRM', 14, 286);
    doc.text(`Page 1 of 1`, 195, 286, { align: 'right' });

    const fname = `LinguaTime_${payrollLabel(payrollYear, payrollMonth)}.pdf`;
    doc.save(fname);
  }

  // Транслитерация русского текста в латиницу (для совместимости со стандартным шрифтом jsPDF)
  function transliterate(text) {
    const map = {
      'а':'a','б':'b','в':'v','г':'g','д':'d','е':'e','ё':'yo','ж':'zh','з':'z',
      'и':'i','й':'y','к':'k','л':'l','м':'m','н':'n','о':'o','п':'p','р':'r',
      'с':'s','т':'t','у':'u','ф':'f','х':'kh','ц':'ts','ч':'ch','ш':'sh','щ':'sch',
      'ъ':'','ы':'y','ь':'','э':'e','ю':'yu','я':'ya',
      'А':'A','Б':'B','В':'V','Г':'G','Д':'D','Е':'E','Ё':'Yo','Ж':'Zh','З':'Z',
      'И':'I','Й':'Y','К':'K','Л':'L','М':'M','Н':'N','О':'O','П':'P','Р':'R',
      'С':'S','Т':'T','У':'U','Ф':'F','Х':'Kh','Ц':'Ts','Ч':'Ch','Ш':'Sh','Щ':'Sch',
      'Ъ':'','Ы':'Y','Ь':'','Э':'E','Ю':'Yu','Я':'Ya'
    };
    return String(text || '').split('').map(c => map[c] !== undefined ? map[c] : c).join('');
  }

  async function reopenPeriod() {
    if (!confirm('Открыть период обратно?\n\nДанные снова станут редактируемыми, ' +
                 'но старый снимок ведомости будет удалён. После повторного закрытия ' +
                 'снимок пересоздастся с актуальными значениями.')) return;

    const btn = document.getElementById('btn-reopen-period');
    btn.disabled = true; btn.textContent = 'Открываем…';

    try {
      if (!currentPeriod) return;
      // Удаляем записи и сбрасываем статус
      await sb.from('payroll_entries').delete().eq('period_id', currentPeriod.id);
      const { error } = await sb
        .from('payroll_periods')
        .update({
          status: 'open',
          closed_at: null,
          closed_by: null,
        })
        .eq('id', currentPeriod.id);
      if (error) throw new Error(error.message);
      await loadPayroll();
    } catch (e) {
      showError('payroll-error', e.message);
    } finally {
      btn.disabled = false; btn.textContent = 'Открыть период';
    }
  }

  // ====================================================================
  // СТРАНИЦА: ДЕТАЛИ ПЕРЕВОДЧИКА (для менеджера)
  // ====================================================================
  let tdUser = null;          // объект переводчика
  let tdCalYear = new Date().getFullYear();
  let tdCalMonth = new Date().getMonth();
  let tdRateByPair = {};      // language_pair_id -> rate
  let tdHistoryByPair = {};   // language_pair_id -> [history items DESC]
  let tdMyPairs = [];         // массив для отображения

  function backToTranslators() {
    showPage('translators');
    loadTranslators();
  }

  // ────────────────────────────────────────────────────────────────
  // РЕДАКТИРОВАНИЕ ПЕРЕВОДЧИКА
  // ────────────────────────────────────────────────────────────────
  let editPairRowCounter = 0;
  let editExistingPairs = [];   // оригинальные пары переводчика для сравнения

  // Считает длительность смены HH:MM → HH:MM в минутах. Поддерживает переход через полночь.
  function calcShiftMinutes(startStr, endStr) {
    if (!startStr || !endStr) return 0;
    const [sh, sm] = startStr.split(':').map(Number);
    const [eh, em] = endStr.split(':').map(Number);
    let startMin = sh * 60 + sm;
    let endMin = eh * 60 + em;
    if (endMin <= startMin) endMin += 24 * 60; // переход через полночь
    return endMin - startMin;
  }

  function updateShiftDuration() {
    const start = document.getElementById('edit-shift-start').value;
    const end = document.getElementById('edit-shift-end').value;
    const minutes = calcShiftMinutes(start, end);
    const display = document.getElementById('edit-shift-duration');
    if (minutes > 0) {
      display.textContent = formatHoursMinutes(minutes);
      display.style.color = '#16A34A';
    } else {
      display.textContent = '—';
      display.style.color = '#94A3B8';
    }
  }

  async function openEditModal() {
    if (!tdUser) return;

    editPairRowCounter = 0;
    document.getElementById('edit-name').value = tdUser.name || '';
    document.getElementById('edit-timezone').value = tdUser.timezone || 'Asia/Tashkent';

    // Дефолтная смена
    const shiftStart = (tdUser.default_shift_start || '09:00:00').substring(0, 5);
    const shiftEnd = (tdUser.default_shift_end || '18:00:00').substring(0, 5);
    document.getElementById('edit-shift-start').value = shiftStart;
    document.getElementById('edit-shift-end').value = shiftEnd;
    updateShiftDuration();

    // Слушатели для пересчёта длительности при изменении
    document.getElementById('edit-shift-start').oninput = updateShiftDuration;
    document.getElementById('edit-shift-end').oninput = updateShiftDuration;
    document.getElementById('edit-pairs-list').innerHTML = '';
    hideError('edit-error');

    // Запомним текущие пары (для сравнения при сохранении)
    editExistingPairs = (tdUser.translator_pairs || []).map(p => ({
      language_pair_id: p.language_pair_id,
      rate_per_hour: Number(p.rate_per_hour),
      is_primary: p.is_primary,
      code: p.language_pairs.code,
      display_name: p.language_pairs.display_name,
    }));

    // Сортируем: основная сверху
    editExistingPairs.sort((a, b) => (b.is_primary ? 1 : 0) - (a.is_primary ? 1 : 0));

    // Рендерим существующие пары
    for (const p of editExistingPairs) {
      addEditPairRow(p);
    }
    if (editExistingPairs.length === 0) {
      addEditPairRow();
    }

    // История ставок
    await loadRateHistory();

    // Список клиентов для привязки
    await populateEditTranslatorClients();

    document.getElementById('edit-modal').classList.add('open');
  }

  // Заполняет dropdown клиентов в модалке редактирования переводчика.
  // Показывает всех клиентов; текущий выбирается автоматически.
  async function populateEditTranslatorClients() {
    const sel = document.getElementById('edit-translator-client');
    const badge = document.getElementById('edit-client-current-badge');
    sel.innerHTML = '<option value="">— без клиента —</option>';

    const { data: clients } = await sb
      .from('clients')
      .select('id, name')
      .eq('is_active', true)
      .order('name');

    for (const c of (clients || [])) {
      const opt = document.createElement('option');
      opt.value = c.id;
      opt.textContent = c.name;
      sel.appendChild(opt);
    }

    sel.value = tdUser.client_id || '';

    // Бейдж текущего клиента
    if (tdUser.client_id) {
      const current = (clients || []).find(c => c.id === tdUser.client_id);
      badge.textContent = current ? current.name : 'привязан';
      badge.classList.remove('hidden');
    } else {
      badge.classList.add('hidden');
    }
  }

  function closeEditModal() {
    document.getElementById('edit-modal').classList.remove('open');
  }

  function addEditPairRow(existing = null) {
    editPairRowCounter++;
    const id = 'ep-' + editPairRowCounter;
    const today = new Date().toISOString().split('T')[0];

    // Опции пар: все активные пары
    const options = languagePairs.map(p => {
      const selected = existing && existing.language_pair_id === p.id ? 'selected' : '';
      return `<option value="${p.id}" ${selected}>${p.code} — ${p.display_name}</option>`;
    }).join('');

    const div = document.createElement('div');
    div.className = 'pair-editor' + (existing ? ' existing' : '');
    div.id = id;
    div.dataset.originalRate = existing ? existing.rate_per_hour : '';
    div.dataset.originalLangId = existing ? existing.language_pair_id : '';
    div.dataset.wasPrimary = existing && existing.is_primary ? '1' : '0';

    if (existing) {
      // Существующая пара
      div.innerHTML = `
        <select class="pair-lang" disabled>${options}</select>
        <div class="pair-rate-input">
          <span style="color:#475569;">$</span>
          <input type="number" class="pair-rate-input-field" step="0.50" min="0.50" max="500" value="${existing.rate_per_hour.toFixed(2)}">
          <span style="color:#475569;font-size:12px;">/ч</span>
        </div>
        <input type="date" class="pair-effective-date" value="${today}" title="Действует с (для новой ставки)">
        <label class="primary-check">
          <input type="radio" name="edit-primary-pair" value="${id}" ${existing.is_primary ? 'checked' : ''}>
          Основная
        </label>
        <button class="pair-remove" onclick="markPairRemoved('${id}')" title="Удалить пару">×</button>
      `;
    } else {
      // Новая пара
      div.innerHTML = `
        <select class="pair-lang">${options}</select>
        <div class="pair-rate-input">
          <span style="color:#475569;">$</span>
          <input type="number" class="pair-rate-input-field" step="0.50" min="0.50" max="500" placeholder="25.00">
          <span style="color:#475569;font-size:12px;">/ч</span>
        </div>
        <input type="date" class="pair-effective-date" value="${today}" title="Действует с">
        <label class="primary-check">
          <input type="radio" name="edit-primary-pair" value="${id}">
          Основная
        </label>
        <button class="pair-remove" onclick="document.getElementById('${id}').remove()" title="Убрать строку">×</button>
      `;
    }

    document.getElementById('edit-pairs-list').appendChild(div);
  }

  function markPairRemoved(id) {
    const el = document.getElementById(id);
    if (!el) return;
    if (el.classList.contains('removed')) {
      // Отмена удаления
      el.classList.remove('removed');
      el.querySelector('.pair-remove').textContent = '×';
      el.querySelector('.pair-remove').title = 'Удалить пару';
    } else {
      if (!confirm('Удалить эту языковую пару? Она исчезнет из доступных переводчику. Историческая ставка останется в rate_history.')) return;
      el.classList.add('removed');
      el.querySelector('.pair-remove').textContent = '↺';
      el.querySelector('.pair-remove').title = 'Отменить удаление';
    }
  }

  async function loadRateHistory() {
    const { data, error } = await sb
      .from('rate_history')
      .select(`
        id, old_rate, new_rate, changed_at, reason,
        language_pairs (code)
      `)
      .eq('user_id', tdUser.id)
      .order('changed_at', { ascending: false });

    const block = document.getElementById('rate-history-list');
    const count = document.getElementById('rate-history-count');
    count.textContent = (data && data.length) || 0;

    if (!data || data.length === 0) {
      block.innerHTML = '<div style="color:#94A3B8; font-size: 12px; padding: 6px 0;">История пуста.</div>';
      return;
    }

    block.innerHTML = data.map(h => `
      <div class="rate-history-item">
        <span class="rate-history-date">${new Date(h.changed_at).toLocaleDateString('ru-RU')}</span>
        <span class="rate-history-change">
          <span class="badge badge-neutral">${h.language_pairs.code}</span>
          ${h.old_rate !== null
            ? `<span class="rate-history-old">$${Number(h.old_rate).toFixed(2)}</span><span class="rate-history-arrow">→</span>`
            : ''}
          <span class="rate-history-new">$${Number(h.new_rate).toFixed(2)}</span>
        </span>
        <span style="font-size: 11px; color: #94A3B8;">${escapeHtml(h.reason || '')}</span>
      </div>
    `).join('');
  }

  function toggleRateHistory() {
    const list = document.getElementById('rate-history-list');
    const arrow = document.getElementById('rate-history-arrow');
    if (list.classList.contains('hidden')) {
      list.classList.remove('hidden');
      arrow.textContent = '▼';
    } else {
      list.classList.add('hidden');
      arrow.textContent = '▶';
    }
  }

  async function saveEditTranslator() {
    hideError('edit-error');
    const btn = document.getElementById('btn-save-edit');

    const name = document.getElementById('edit-name').value.trim();
    const timezone = document.getElementById('edit-timezone').value;
    const shiftStart = document.getElementById('edit-shift-start').value;
    const shiftEnd = document.getElementById('edit-shift-end').value;
    const shiftMinutes = calcShiftMinutes(shiftStart, shiftEnd);

    if (!name) {
      showError('edit-error', 'Имя не может быть пустым.');
      return;
    }
    if (!shiftStart || !shiftEnd || shiftMinutes <= 0) {
      showError('edit-error', 'Укажите корректную стандартную смену.');
      return;
    }
    if (shiftMinutes > 16 * 60) {
      showError('edit-error', 'Стандартная смена не может быть больше 16 часов.');
      return;
    }

    // Собираем пары из формы
    const rows = document.querySelectorAll('#edit-pairs-list .pair-editor');
    if (rows.length === 0) {
      showError('edit-error', 'Должна быть хотя бы одна языковая пара.');
      return;
    }

    const formPairs = []; // { rowId, langId, rate, effectiveDate, isPrimary, isExisting, originalRate, removed }
    let primaryFound = false;
    let validationError = null;

    rows.forEach(row => {
      const removed = row.classList.contains('removed');
      const langId = row.querySelector('.pair-lang').value;
      const rate = parseFloat(row.querySelector('.pair-rate-input-field').value);
      const effectiveDate = row.querySelector('.pair-effective-date').value;
      const isPrimary = row.querySelector('input[type="radio"]').checked;
      const isExisting = row.classList.contains('existing');
      const originalRate = row.dataset.originalRate ? Number(row.dataset.originalRate) : null;
      const originalLangId = row.dataset.originalLangId || null;

      if (removed) {
        formPairs.push({ rowId: row.id, langId: originalLangId || langId, removed: true, isExisting });
        return;
      }

      if (!langId) { validationError = 'Выберите языковую пару во всех строках.'; return; }
      if (!rate || rate < 0.5 || rate > 500) { validationError = 'Все ставки должны быть от $0.50 до $500.00.'; return; }
      if (!effectiveDate) { validationError = 'Укажите дату начала действия ставки.'; return; }
      if (isPrimary) primaryFound = true;

      formPairs.push({
        rowId: row.id, langId, rate, effectiveDate, isPrimary, isExisting,
        originalRate, originalLangId
      });
    });

    if (validationError) { showError('edit-error', validationError); return; }

    // Проверки: должна остаться хотя бы одна не-удалённая пара
    const activePairs = formPairs.filter(p => !p.removed);
    if (activePairs.length === 0) {
      showError('edit-error', 'Должна остаться хотя бы одна языковая пара.');
      return;
    }
    if (!primaryFound) {
      showError('edit-error', 'Отметьте одну пару как основную.');
      return;
    }

    // Уникальность пар
    const langIds = activePairs.map(p => p.langId);
    if (new Set(langIds).size !== langIds.length) {
      showError('edit-error', 'Одна и та же языковая пара указана несколько раз.');
      return;
    }

    btn.disabled = true; btn.textContent = 'Сохранение…';

    try {
      // 1) Обновляем имя и часовой пояс
      const { error: profileErr } = await sb
        .from('users')
        .update({ name, timezone })
        .eq('id', tdUser.id);
      if (profileErr) throw new Error('Профиль: ' + profileErr.message);

      // 1a) График — через set_schedule (атомарно пишет shift_history), только если изменился
      const curStart = (tdUser.default_shift_start || '').substring(0, 5);
      const curEnd = (tdUser.default_shift_end || '').substring(0, 5);
      const curMin = tdUser.default_shift_minutes || 0;
      if (curStart !== shiftStart || curEnd !== shiftEnd || curMin !== shiftMinutes) {
        const { error: schedErr } = await sb.rpc('set_schedule', {
          p_user: tdUser.id,
          p_start: shiftStart + ':00',
          p_end: shiftEnd + ':00',
          p_minutes: shiftMinutes,
          p_reason: 'Ручная правка графика в карточке',
        });
        if (schedErr) throw new Error('График: ' + schedErr.message);
      }

      // 1b) Привязка к клиенту (если изменилась)
      const newClientId = document.getElementById('edit-translator-client').value || null;
      if (newClientId !== (tdUser.client_id || null)) {
        const { error: clientLinkErr } = await sb
          .from('users')
          .update({ client_id: newClientId })
          .eq('id', tdUser.id);
        if (clientLinkErr) throw new Error('Привязка к клиенту: ' + clientLinkErr.message);
        // При привязке к клиенту — автозаведение тренинга (если ещё нет)
        if (newClientId) {
          await createTrainingForTranslator(tdUser.id, newClientId);
        }
        // Клиент изменился — чистим кеш прибыльности
        invalidateCache('clientProfit:');
      }

      // 2) Обрабатываем пары
      const rateHistoryInserts = [];

      // Сначала: сбрасываем is_primary у всех (потом установим у одной)
      await sb.from('translator_pairs')
        .update({ is_primary: false })
        .eq('user_id', tdUser.id);

      for (const p of formPairs) {
        if (p.removed && p.isExisting) {
          // Удаляем пару
          await sb.from('translator_pairs')
            .delete()
            .eq('user_id', tdUser.id)
            .eq('language_pair_id', p.langId);
          continue;
        }
        if (p.removed) continue;

        if (p.isExisting) {
          // Существующая пара. Если ставка изменилась — пишем в rate_history.
          const rateChanged = Math.abs(p.rate - p.originalRate) > 0.001;
          if (rateChanged) {
            rateHistoryInserts.push({
              user_id: tdUser.id,
              language_pair_id: p.langId,
              old_rate: p.originalRate,
              new_rate: p.rate,
              changed_by: currentUser.id,
              effective_from: p.effectiveDate,
              reason: `Изменение ставки с ${p.effectiveDate}`,
            });
          }
          // Обновляем translator_pairs
          await sb.from('translator_pairs')
            .update({
              rate_per_hour: p.rate,
              is_primary: p.isPrimary,
              effective_from: p.effectiveDate,
            })
            .eq('user_id', tdUser.id)
            .eq('language_pair_id', p.langId);
        } else {
          // Новая пара
          await sb.from('translator_pairs').insert({
            user_id: tdUser.id,
            language_pair_id: p.langId,
            rate_per_hour: p.rate,
            is_primary: p.isPrimary,
            effective_from: p.effectiveDate,
          });
          rateHistoryInserts.push({
            user_id: tdUser.id,
            language_pair_id: p.langId,
            old_rate: null,
            new_rate: p.rate,
            changed_by: currentUser.id,
            effective_from: p.effectiveDate,
            reason: `Добавлена пара с ${p.effectiveDate}`,
          });
        }
      }

      // 3) Запись истории
      if (rateHistoryInserts.length > 0) {
        await sb.from('rate_history').insert(rateHistoryInserts);
      }

      // Изменения ставок/пар влияют на расчёты dashboard/payroll/calendar — чистим кеши
      invalidateCache('rateHistory:');

      // 4) Закрываем модалку и перезагружаем профиль
      closeEditModal();
      await openTranslatorDetail(tdUser.id);
    } catch (e) {
      showError('edit-error', e.message);
    } finally {
      btn.disabled = false; btn.textContent = 'Сохранить';
    }
  }

  // ────────────────────────────────────────────────────────────────
  // ДЕАКТИВАЦИЯ / АКТИВАЦИЯ
  // ────────────────────────────────────────────────────────────────
  async function toggleActive() {
    if (!tdUser) return;
    const newState = !tdUser.is_active;
    const action = newState ? 'активировать' : 'деактивировать';

    const message = newState
      ? `Активировать ${tdUser.name}? Переводчик снова сможет входить в систему.`
      : `Деактивировать ${tdUser.name}?\n\n` +
        `Переводчик не сможет войти в систему.\n` +
        `Все его данные и история сохранятся.\n` +
        `Активировать обратно можно в любое время.`;

    if (!confirm(message)) return;

    try {
      const updateData = newState
        ? { is_active: true, deactivated_at: null }
        : { is_active: false, deactivated_at: new Date().toISOString() };

      const { error } = await sb
        .from('users')
        .update(updateData)
        .eq('id', tdUser.id);

      if (error) throw new Error(error.message);

      // Перезагружаем профиль
      await openTranslatorDetail(tdUser.id);
    } catch (e) {
      alert('Ошибка ' + action + ': ' + e.message);
    }
  }

  async function openTranslatorDetail(userId) {
    // Загружаем профиль и его пары
    const { data: user, error } = await sb
      .from('users')
      .select(`
        id, email, name, timezone, is_active, created_at, client_id,
        default_shift_start, default_shift_end, default_shift_minutes,
        translator_pairs ( language_pair_id, rate_per_hour, is_primary, language_pairs (code, display_name) )
      `)
      .eq('id', userId)
      .single();

    if (error || !user) {
      alert('Не удалось загрузить переводчика: ' + (error?.message || ''));
      return;
    }

    tdUser = user;
    tdRateByPair = {};
    tdMyPairs = (user.translator_pairs || []).sort((a, b) => (b.is_primary ? 1 : 0) - (a.is_primary ? 1 : 0));
    for (const p of tdMyPairs) {
      tdRateByPair[p.language_pair_id] = Number(p.rate_per_hour);
    }
    // Загружаем историю ставок этого переводчика
    const historyByUser = await loadRateHistoryForUsers([user.id]);
    tdHistoryByPair = historyByUser[user.id] || {};

    // Текущий месяц по умолчанию
    tdCalYear = new Date().getFullYear();
    tdCalMonth = new Date().getMonth();

    // Заполняем шапку
    const initials = (user.name || '?').split(' ').map(s => s[0]).join('').slice(0, 2).toUpperCase();
    document.getElementById('td-avatar').textContent = initials;
    document.getElementById('td-name').textContent = user.name;
    document.getElementById('td-email').textContent = user.email;

    // Бейджи: смена + языковые пары
    const shiftStart = (user.default_shift_start || '09:00:00').substring(0, 5);
    const shiftEnd = (user.default_shift_end || '18:00:00').substring(0, 5);
    const shiftMin = user.default_shift_minutes || 480;
    const shiftBadge = `<span class="badge badge-info" title="Стандартная смена">⏱ ${shiftStart}–${shiftEnd} · ${formatHoursMinutes(shiftMin)}</span>`;
    const pairsBadges = tdMyPairs.map(p =>
      `<span class="badge badge-neutral">${p.language_pairs.code} · $${Number(p.rate_per_hour).toFixed(2)}/ч</span>`
    ).join('');
    document.getElementById('td-pairs').innerHTML = shiftBadge + pairsBadges;

    // Кнопка деактивации
    const btnToggle = document.getElementById('btn-toggle-active');
    if (user.is_active) {
      btnToggle.textContent = 'Деактивировать';
      btnToggle.style.color = '#DC2626';
    } else {
      btnToggle.textContent = 'Активировать';
      btnToggle.style.color = '#16A34A';
    }

    // Заполняем селектор месяцев (последние 12 месяцев + текущий)
    fillMonthSelector();

    showPage('translator-detail');
    await loadTranslatorDetailMonth();
    await loadTranslatorShiftHistory();
  }

  // Загружает и отображает историю запросов на изменение графика для tdUser
  async function loadTranslatorShiftHistory() {
    if (!tdUser) return;
    const section = document.getElementById('td-shift-history-section');
    const list = document.getElementById('td-shift-history-list');

    const { data, error } = await sb.from('shift_change_requests')
      .select(`id, status, created_at, reviewed_at, reason, review_note,
               requested_start, requested_end, requested_minutes, shift_date,
               final_start, final_end, final_minutes, final_effective_from`)
      .eq('user_id', tdUser.id)
      .order('created_at', { ascending: false })
      .limit(20);

    if (error || !data || data.length === 0) {
      section.style.display = 'none';
      return;
    }

    section.style.display = '';
    list.innerHTML = data.map(r => renderShiftRequestHistoryItem(r, true)).join('');
  }

  function fillMonthSelector() {
    const sel = document.getElementById('td-month-select');
    sel.innerHTML = '';
    const now = new Date();
    for (let i = 0; i < 12; i++) {
      const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
      const value = `${d.getFullYear()}-${d.getMonth()}`;
      const label = MONTH_NAMES_RU[d.getMonth()] + ' ' + d.getFullYear();
      const opt = document.createElement('option');
      opt.value = value;
      opt.textContent = label;
      sel.appendChild(opt);
    }
    sel.value = `${tdCalYear}-${tdCalMonth}`;
    sel.onchange = () => {
      const [y, m] = sel.value.split('-').map(Number);
      tdCalYear = y; tdCalMonth = m;
      loadTranslatorDetailMonth();
    };
  }

  // Состояние календаря менеджера (детализация переводчика) — для single-cell update
  let tdCalState = null;

  async function loadTranslatorDetailMonth() {
    hideError('td-error');

    document.getElementById('td-cal-month-label').textContent =
      MONTH_NAMES_RU[tdCalMonth] + ' ' + tdCalYear;

    const periodStart = formatDate(new Date(tdCalYear, tdCalMonth, 1));
    const periodEnd = formatDate(new Date(tdCalYear, tdCalMonth + 1, 0));

    const [daysRes, shiftsRes] = await Promise.all([
      sb.from('work_days')
        .select(`
          id, work_date, day_type,
          work_intervals ( id, duration_minutes, language_pair_id ),
          breaks ( id, duration_minutes )
        `)
        .eq('user_id', tdUser.id)
        .gte('work_date', periodStart)
        .lte('work_date', periodEnd),
      sb.from('shifts')
        .select('shift_date, planned_minutes')
        .eq('user_id', tdUser.id)
        .gte('shift_date', periodStart)
        .lte('shift_date', periodEnd)
    ]);

    if (daysRes.error) {
      showError('td-error', 'Ошибка загрузки: ' + daysRes.error.message);
      return;
    }

    // Маппинг смен по дате
    const shiftsByDate = {};
    for (const s of (shiftsRes.data || [])) {
      shiftsByDate[s.shift_date] = s.planned_minutes;
    }
    const defaultMin = tdUser.default_shift_minutes || 480;

    const dayMap = {};
    let totalMinutes = 0, totalAmount = 0, absentDays = 0, workingDays = 0;
    let totalOvertime = 0;
    for (const d of daysRes.data || []) {
      if (d.day_type === 'absent') {
        dayMap[d.work_date] = { type: 'absent' };
        absentDays++;
        continue;
      }
      const calc = calcDayWithHistory(d, tdHistoryByPair, tdRateByPair);
      if (calc.netMinutes > 0) {
        dayMap[d.work_date] = { type: 'working', minutes: calc.netMinutes, amount: calc.amount };
        totalMinutes += calc.netMinutes;
        totalAmount += calc.amount;
        workingDays++;

        const planned = getPlannedMinutes(d.work_date, shiftsByDate, defaultMin);
        totalOvertime += (calc.netMinutes - planned);
      }
    }

    tdCalState = {
      year: tdCalYear, month: tdCalMonth,
      dayMap, shiftsByDate, defaultMin,
      totals: { minutes: totalMinutes, amount: totalAmount, workingDays, absentDays, overtime: totalOvertime },
    };

    updateTdKpis();
    renderTdCalendar(dayMap);
  }

  function updateTdKpis() {
    if (!tdCalState) return;
    const { totals } = tdCalState;
    document.getElementById('td-hours').textContent = formatHoursMinutes(totals.minutes);
    document.getElementById('td-hours-meta').textContent =
      `${totals.workingDays} ${pluralize(totals.workingDays, 'рабочий день', 'рабочих дня', 'рабочих дней')}`;
    document.getElementById('td-absent').innerHTML =
      totals.absentDays + '<span class="kpi-unit">' +
      (totals.absentDays === 0 ? 'дней' : pluralize(totals.absentDays, 'день', 'дня', 'дней')) +
      '</span>';
    document.getElementById('td-amount').textContent = '$' + totals.amount.toFixed(2);

    // Овертайм в новой KPI карточке
    const otEl = document.getElementById('td-overtime');
    const otMetaEl = document.getElementById('td-overtime-meta');
    if (otEl) {
      if (Math.abs(totals.overtime) < 1) {
        otEl.innerHTML = '0<span class="kpi-unit">мин</span>';
        otEl.style.color = '';
        otMetaEl.textContent = 'точно по плану';
      } else if (totals.overtime > 0) {
        otEl.innerHTML = '+' + formatHoursMinutes(totals.overtime);
        otEl.style.color = '#B45309';
        otMetaEl.textContent = 'переработано';
      } else {
        otEl.innerHTML = '−' + formatHoursMinutes(Math.abs(totals.overtime));
        otEl.style.color = '#1E40AF';
        otMetaEl.textContent = 'недоработано';
      }
    }
  }

  // Single-cell обновление календаря в режиме менеджера. Используется и для дня
  // (изменение часов/отгула), и для смены — обновление shiftsByDate тоже корректно
  // пересчитывает овертайм.
  async function refreshTdCalendarDay(dateStr) {
    if (!tdCalState || tdCalState.year !== tdCalYear || tdCalState.month !== tdCalMonth) {
      return loadTranslatorDetailMonth();
    }
    const periodStart = formatDate(new Date(tdCalYear, tdCalMonth, 1));
    const periodEnd = formatDate(new Date(tdCalYear, tdCalMonth + 1, 0));
    if (dateStr < periodStart || dateStr > periodEnd) return loadTranslatorDetailMonth();

    // Параллельно перечитываем день и смену на этот день
    const [dayRes, shiftRes] = await Promise.all([
      sb.from('work_days')
        .select(`
          id, work_date, day_type,
          work_intervals ( id, duration_minutes, language_pair_id ),
          breaks ( id, duration_minutes )
        `)
        .eq('user_id', tdUser.id)
        .eq('work_date', dateStr)
        .maybeSingle(),
      sb.from('shifts')
        .select('shift_date, planned_minutes')
        .eq('user_id', tdUser.id)
        .eq('shift_date', dateStr)
        .maybeSingle()
    ]);
    if (dayRes.error) return loadTranslatorDetailMonth();

    // Откатываем вклад старого дня и его овертайма
    const old = tdCalState.dayMap[dateStr];
    if (old) {
      if (old.type === 'working') {
        tdCalState.totals.minutes -= old.minutes;
        tdCalState.totals.amount -= old.amount;
        tdCalState.totals.workingDays -= 1;
        const oldPlanned = getPlannedMinutes(dateStr, tdCalState.shiftsByDate, tdCalState.defaultMin);
        tdCalState.totals.overtime -= (old.minutes - oldPlanned);
      } else if (old.type === 'absent') {
        tdCalState.totals.absentDays -= 1;
      }
      delete tdCalState.dayMap[dateStr];
    }

    // Обновляем shiftsByDate (смена могла измениться)
    if (shiftRes.data) {
      tdCalState.shiftsByDate[dateStr] = shiftRes.data.planned_minutes;
    } else {
      delete tdCalState.shiftsByDate[dateStr];
    }

    // Применяем новый день
    const day = dayRes.data;
    if (day) {
      if (day.day_type === 'absent') {
        tdCalState.dayMap[dateStr] = { type: 'absent' };
        tdCalState.totals.absentDays += 1;
      } else {
        const calc = calcDayWithHistory(day, tdHistoryByPair, tdRateByPair);
        if (calc.netMinutes > 0) {
          tdCalState.dayMap[dateStr] = { type: 'working', minutes: calc.netMinutes, amount: calc.amount };
          tdCalState.totals.minutes += calc.netMinutes;
          tdCalState.totals.amount += calc.amount;
          tdCalState.totals.workingDays += 1;
          const planned = getPlannedMinutes(dateStr, tdCalState.shiftsByDate, tdCalState.defaultMin);
          tdCalState.totals.overtime += (calc.netMinutes - planned);
        }
      }
    }

    // Перерисовываем только эту ячейку
    const oldCell = document.querySelector(`#td-calendar-grid .cal-day[data-date="${dateStr}"]`);
    if (oldCell) {
      const dayNum = Number(dateStr.split('-')[2]);
      const todayStr = formatDate(new Date());
      const newCell = buildCalendarCell(dayNum, dateStr, tdCalYear, tdCalMonth, tdCalState.dayMap, todayStr);
      newCell.addEventListener('click', () => openDayViewModal(dateStr));
      oldCell.replaceWith(newCell);
    }
    updateTdKpis();
  }

  function renderTdCalendar(dayMap) {
    const grid = document.getElementById('td-calendar-grid');
    const frag = document.createDocumentFragment();

    DAY_NAMES.forEach(n => {
      const el = document.createElement('div');
      el.className = 'cal-day-name';
      el.textContent = n;
      frag.appendChild(el);
    });
    const firstDay = new Date(tdCalYear, tdCalMonth, 1);
    let dow = firstDay.getDay() - 1;
    if (dow < 0) dow = 6;
    for (let i = 0; i < dow; i++) {
      const empty = document.createElement('div');
      empty.className = 'cal-day empty';
      frag.appendChild(empty);
    }
    const daysInMonth = new Date(tdCalYear, tdCalMonth + 1, 0).getDate();
    const todayStr = formatDate(new Date());

    for (let d = 1; d <= daysInMonth; d++) {
      const dateStr = `${tdCalYear}-${String(tdCalMonth + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
      const cell = buildCalendarCell(d, dateStr, tdCalYear, tdCalMonth, dayMap, todayStr);
      // Клик на день — открыть просмотр (любой день, чтобы можно было назначить смену)
      cell.addEventListener('click', () => openDayViewModal(dateStr));
      frag.appendChild(cell);
    }

    grid.replaceChildren(frag);
  }

  function changeTdMonth(delta) {
    tdCalMonth += delta;
    if (tdCalMonth < 0) { tdCalMonth = 11; tdCalYear--; }
    if (tdCalMonth > 11) { tdCalMonth = 0; tdCalYear++; }
    const sel = document.getElementById('td-month-select');
    sel.value = `${tdCalYear}-${tdCalMonth}`;
    loadTranslatorDetailMonth();
  }

  // ────────────────────────────────────────────────────────────────
  // Read-only просмотр дня (для менеджера)
  // Менеджер: открыть модалку ввода/редактирования дня за переводчика
  async function openDayModalAsManager() {
    const dateStr = tdViewDate || document.getElementById('day-view-modal').dataset.date;
    if (!dateStr || !tdUser) {
      alert('Ошибка: дата или переводчик не определены.');
      return;
    }
    // Закроем модалку просмотра
    closeDayViewModal();
    // Откроем модалку ввода с targetUser = tdUser (она сама добавит .open в конце)
    await openDayModal(dateStr, tdUser);
  }

  // ────────────────────────────────────────────────────────────────
  // Глобальное состояние для модалки просмотра дня (для менеджера)
  let tdViewDate = null;
  let tdViewShift = null; // shifts row or null

  async function openDayViewModal(dateStr) {
    tdViewDate = dateStr;
    document.getElementById('day-view-modal').dataset.date = dateStr;
    const date = new Date(dateStr + 'T00:00:00');
    const dayName = ['Воскресенье','Понедельник','Вторник','Среда','Четверг','Пятница','Суббота'][date.getDay()];
    const monthName = MONTH_NAMES_RU[date.getMonth()];
    document.getElementById('day-view-title').textContent =
      `${date.getDate()} ${monthName.toLowerCase()} · ${tdUser.name}`;
    document.getElementById('day-view-subtitle').textContent = dayName;

    const body = document.getElementById('day-view-body');
    body.innerHTML = '<div class="loading-state">Загрузка…</div>';

    // Сначала прячем обе кнопки — покажем нужную ниже
    document.getElementById('btn-manager-fill-day').classList.add('hidden');
    document.getElementById('btn-manager-edit-day').classList.add('hidden');

    document.getElementById('day-view-modal').classList.add('open');

    // Параллельно загружаем день и смену
    const [dayRes, shiftRes] = await Promise.all([
      sb.from('work_days')
        .select(`
          id, day_type, note,
          work_intervals ( id, language_pair_id, start_at, end_at, duration_minutes,
                           language_pairs (code) ),
          breaks ( id, start_at, end_at, duration_minutes, note )
        `)
        .eq('user_id', tdUser.id)
        .eq('work_date', dateStr)
        .maybeSingle(),
      sb.from('shifts')
        .select('id, planned_start, planned_end, planned_minutes, note')
        .eq('user_id', tdUser.id)
        .eq('shift_date', dateStr)
        .maybeSingle()
    ]);

    const day = dayRes.data;
    tdViewShift = shiftRes.data;

    // Показываем нужную кнопку в футере
    if (day) {
      document.getElementById('btn-manager-edit-day').classList.remove('hidden');
    } else {
      document.getElementById('btn-manager-fill-day').classList.remove('hidden');
    }

    // Рендерим блок СМЕНЫ всегда (даже если работы не было)
    const shiftHtml = renderShiftSection(tdViewShift);

    if (!day) {
      // Нет записи о работе — показываем только смену и подсказку
      body.innerHTML = shiftHtml + `
        <div class="day-readonly-empty" style="margin-top: 12px;">
          На этот день нет записей о работе.
        </div>
      `;
      return;
    }

    if (day.day_type === 'absent') {
      body.innerHTML = shiftHtml + `
        <div style="padding: 18px; text-align: center; color: #1E40AF; background: #DBEAFE; border-radius: 8px; margin-top: 12px;">
          <div style="font-size: 14px; font-weight: 500; margin-bottom: 4px;">Отгул</div>
          <div style="font-size: 12px;">Этот день отмечен как отгул. Оплата не начисляется.</div>
        </div>
      `;
      return;
    }

    // Рабочий день
    const intervals = (day.work_intervals || []).sort((a, b) => a.start_at.localeCompare(b.start_at));
    const breaks = (day.breaks || []).sort((a, b) => a.start_at.localeCompare(b.start_at));

    const grossMin = intervals.reduce((s, i) => s + (i.duration_minutes || 0), 0);
    const breaksMin = breaks.reduce((s, b) => s + (b.duration_minutes || 0), 0);
    const netMin = Math.max(0, grossMin - breaksMin);

    const intervalRates = intervals.map(i =>
      getRateForDate(dateStr, i.language_pair_id, tdHistoryByPair, tdRateByPair)
    );

    let dayAmount = 0;
    intervals.forEach((interv, idx) => {
      dayAmount += (interv.duration_minutes / 60) * intervalRates[idx];
    });
    if (grossMin > 0) dayAmount = dayAmount * (netMin / grossMin);

    // Овертайм: фактические минуты vs плановые
    const plannedMin = tdViewShift
      ? tdViewShift.planned_minutes
      : (tdUser.default_shift_minutes || 480);
    const overtimeMin = netMin - plannedMin;
    let overtimeHtml = '';
    if (Math.abs(overtimeMin) >= 1) {
      const cls = overtimeMin > 0 ? 'over' : 'under';
      const sign = overtimeMin > 0 ? '+' : '−';
      overtimeHtml = `
        <span class="overtime-widget ${cls}" title="Разница с плановой сменой">
          ${sign}${formatHoursMinutes(Math.abs(overtimeMin))}
          ${overtimeMin > 0 ? '· овертайм' : '· недоработка'}
        </span>
      `;
    } else {
      overtimeHtml = `<span class="overtime-widget exact">Точно по плану</span>`;
    }

    let html = shiftHtml + `
      <div class="day-readonly-section">
        <div class="day-readonly-title">Рабочие интервалы</div>
        ${intervals.length === 0
          ? '<div class="day-readonly-empty">Нет интервалов.</div>'
          : intervals.map((i, idx) => `
            <div class="day-readonly-item">
              <span class="day-readonly-time">${formatTimeHM(i.start_at)} → ${formatTimeHM(i.end_at)}</span>
              <span class="day-readonly-meta">${i.language_pairs.code} · $${intervalRates[idx].toFixed(2)}/ч</span>
              <span class="day-readonly-dur">${formatHoursMinutes(i.duration_minutes)}</span>
            </div>
          `).join('')}
      </div>
    `;

    if (breaks.length > 0) {
      html += `
        <div class="day-readonly-section">
          <div class="day-readonly-title">Брейки</div>
          ${breaks.map(b => `
            <div class="day-readonly-item break-item">
              <span class="day-readonly-time">${formatTimeHM(b.start_at)} → ${formatTimeHM(b.end_at)}</span>
              <span class="day-readonly-meta">${escapeHtml(b.note || '')}</span>
              <span class="day-readonly-dur">${formatHoursMinutes(b.duration_minutes)}</span>
            </div>
          `).join('')}
        </div>
      `;
    }

    html += `
      <div class="day-summary">
        <div class="day-summary-cell">
          <div class="day-summary-label">Чистое время</div>
          <div class="day-summary-value green">${formatHoursMinutes(netMin)}</div>
          <div style="margin-top: 4px;">${overtimeHtml}</div>
        </div>
        <div class="day-summary-cell">
          <div class="day-summary-label">Брейки</div>
          <div class="day-summary-value">${formatHoursMinutes(breaksMin)}</div>
        </div>
        <div class="day-summary-cell">
          <div class="day-summary-label">К оплате</div>
          <div class="day-summary-value green">$${dayAmount.toFixed(2)}</div>
        </div>
      </div>
    `;

    body.innerHTML = html;
  }

  // Рендерит блок смены (либо назначенной, либо дефолтной)
  function renderShiftSection(shift) {
    if (shift) {
      const start = shift.planned_start.substring(0, 5);
      const end = shift.planned_end.substring(0, 5);
      const dur = formatHoursMinutes(shift.planned_minutes);
      return `
        <div class="shift-section assigned">
          <div class="shift-section-header">
            <span class="shift-section-label">Особая смена назначена</span>
            <div class="shift-section-actions">
              <button class="shift-section-btn" onclick="openShiftModal()">✎ Изменить</button>
              <button class="shift-section-btn danger" onclick="deleteShift()">× Удалить</button>
            </div>
          </div>
          <div class="shift-section-content">
            <span class="shift-section-time">${start} – ${end}</span>
            <span class="shift-section-duration">· ${dur} плана</span>
          </div>
          ${shift.note ? `<div class="shift-section-note">${escapeHtml(shift.note)}</div>` : ''}
        </div>
      `;
    } else {
      const defStart = (tdUser.default_shift_start || '09:00:00').substring(0, 5);
      const defEnd = (tdUser.default_shift_end || '18:00:00').substring(0, 5);
      const defMin = tdUser.default_shift_minutes || 480;
      return `
        <div class="shift-section">
          <div class="shift-section-header">
            <span class="shift-section-label">Стандартная смена (дефолт)</span>
            <div class="shift-section-actions">
              <button class="shift-section-btn" onclick="openShiftModal()">+ Назначить особую</button>
            </div>
          </div>
          <div class="shift-section-content">
            <span class="shift-section-time">${defStart} – ${defEnd}</span>
            <span class="shift-section-duration">· ${formatHoursMinutes(defMin)} плана</span>
          </div>
        </div>
      `;
    }
  }

  // ────────────────────────────────────────────────────────────────
  // НАЗНАЧЕНИЕ / ИЗМЕНЕНИЕ СМЕНЫ
  // ────────────────────────────────────────────────────────────────
  function openShiftModal() {
    hideError('shift-modal-error');
    // Защита: читаем дату из data-attribute если переменная потерялась
    const dateStr = tdViewDate || document.getElementById('day-view-modal').dataset.date;
    if (!dateStr) {
      alert('Ошибка: дата не определена. Закройте модалку и откройте день заново.');
      return;
    }
    tdViewDate = dateStr; // восстанавливаем переменную
    const date = new Date(dateStr + 'T00:00:00');
    const dayName = ['Воскресенье','Понедельник','Вторник','Среда','Четверг','Пятница','Суббота'][date.getDay()];
    const monthName = MONTH_NAMES_RU[date.getMonth()];
    document.getElementById('shift-modal-title').textContent =
      tdViewShift ? 'Изменить смену' : 'Назначить смену';
    document.getElementById('shift-modal-subtitle').textContent =
      `${date.getDate()} ${monthName.toLowerCase()}, ${dayName}`;

    // Префилл — либо текущая смена, либо дефолт
    const start = tdViewShift
      ? tdViewShift.planned_start.substring(0, 5)
      : (tdUser.default_shift_start || '09:00:00').substring(0, 5);
    const end = tdViewShift
      ? tdViewShift.planned_end.substring(0, 5)
      : (tdUser.default_shift_end || '18:00:00').substring(0, 5);

    document.getElementById('shift-start').value = start;
    document.getElementById('shift-end').value = end;
    document.getElementById('shift-note').value = tdViewShift?.note || '';
    updateShiftModalDuration();

    document.getElementById('shift-start').oninput = updateShiftModalDuration;
    document.getElementById('shift-end').oninput = updateShiftModalDuration;

    document.getElementById('shift-modal').classList.add('open');
  }

  function closeShiftModal() {
    document.getElementById('shift-modal').classList.remove('open');
  }

  function updateShiftModalDuration() {
    const start = document.getElementById('shift-start').value;
    const end = document.getElementById('shift-end').value;
    const minutes = calcShiftMinutes(start, end);
    const display = document.getElementById('shift-duration');
    if (minutes > 0) {
      display.textContent = formatHoursMinutes(minutes);
      display.style.color = '#16A34A';
    } else {
      display.textContent = '—';
      display.style.color = '#94A3B8';
    }
  }

  async function saveShift() {
    hideError('shift-modal-error');
    const btn = document.getElementById('btn-save-shift');

    const start = document.getElementById('shift-start').value;
    const end = document.getElementById('shift-end').value;
    const note = document.getElementById('shift-note').value.trim();
    const minutes = calcShiftMinutes(start, end);

    // Защита: читаем дату из переменной или data-attribute
    const dateStr = tdViewDate || document.getElementById('day-view-modal').dataset.date;
    if (!dateStr) {
      showError('shift-modal-error', 'Ошибка: дата смены не определена. Закройте модалку и откройте день заново.');
      return;
    }

    if (!start || !end || minutes <= 0) {
      showError('shift-modal-error', 'Укажите корректное время смены.');
      return;
    }
    if (minutes > 16 * 60) {
      showError('shift-modal-error', 'Смена не может быть больше 16 часов.');
      return;
    }

    btn.disabled = true; btn.textContent = 'Сохранение…';

    try {
      const crossesMidnight = (end <= start);
      const payload = {
        user_id: tdUser.id,
        shift_date: dateStr,
        planned_start: start + ':00',
        planned_end: end + ':00',
        planned_minutes: minutes,
        crosses_midnight: crossesMidnight,
        assigned_by: currentUser.id,
        note: note || null,
      };

      if (tdViewShift) {
        // UPDATE
        const { error } = await sb.from('shifts')
          .update(payload)
          .eq('id', tdViewShift.id);
        if (error) throw new Error(error.message);
      } else {
        // INSERT
        const { error } = await sb.from('shifts').insert(payload);
        if (error) throw new Error(error.message);
      }

      closeShiftModal();
      // Кеш смен мог содержать этот период — выкидываем, чтобы dashboard/payroll увидели изменение
      invalidateCache('shifts:');
      // Перезагружаем модалку дня и обновляем только эту ячейку календаря
      await openDayViewModal(dateStr);
      await refreshTdCalendarDay(dateStr);
    } catch (e) {
      showError('shift-modal-error', e.message);
    } finally {
      btn.disabled = false; btn.textContent = 'Сохранить';
    }
  }

  async function deleteShift() {
    if (!tdViewShift) return;
    if (!confirm('Удалить назначенную смену? Будет использоваться стандартная смена переводчика.')) return;

    const dateStr = tdViewDate || document.getElementById('day-view-modal').dataset.date;

    try {
      const { error } = await sb.from('shifts')
        .delete()
        .eq('id', tdViewShift.id);
      if (error) throw new Error(error.message);

      invalidateCache('shifts:');
      if (dateStr) await openDayViewModal(dateStr);
      if (dateStr) await refreshTdCalendarDay(dateStr);
    } catch (e) {
      alert('Ошибка удаления: ' + e.message);
    }
  }

  function closeDayViewModal() {
    document.getElementById('day-view-modal').classList.remove('open');
    document.getElementById('day-view-modal').dataset.date = '';
    tdViewDate = null;
    tdViewShift = null;
  }

  // ====================================================================
  // ИСТОРИЧЕСКИЕ СТАВКИ
  // ====================================================================
  // Для расчётов нужна актуальная ставка на конкретную дату.
  //
  // Логика: смотрим rate_history для (user_id, language_pair_id),
  // берём запись с changed_at::date <= work_date с самой свежей датой.
  // Если такой нет — используем текущую ставку из translator_pairs.
  //
  // Эта функция работает с УЖЕ ЗАГРУЖЕННЫМИ данными — её вызывают после
  // того, как вытянули и пары, и историю одним запросом, чтобы не делать
  // запрос на каждую минуту.
  //
  // Параметры:
  //   workDate      — строка 'YYYY-MM-DD'
  //   pairId        — language_pair_id
  //   historyByPair — { pair_id: [{ new_rate, changed_at }, ...] (отсорт. по убыв.) }
  //   currentRates  — { pair_id: rate } (текущие, fallback)
  function getRateForDate(workDate, pairId, historyByPair, currentRates) {
    const history = historyByPair[pairId] || [];
    // История отсортирована по changed_at DESC (самые свежие изменения первыми).
    // Ищем первую запись с effective_from <= workDate.
    // Это даёт правильный ответ даже если более новое изменение откатывает
    // ставку задним числом ("изменение с 1 мая" применится ко всему маю).
    for (const h of history) {
      const effective = h.effective_from || h.changed_at.split('T')[0];
      if (effective <= workDate) {
        return Number(h.new_rate);
      }
    }
    // Не нашли в истории — fallback на текущую ставку
    return Number(currentRates[pairId] || 0);
  }

  // Загружает историю ставок для набора пользователей и группирует по парам.
  // Возвращает { user_id: { pair_id: [history items DESC by changed_at] } }
  // Сортировка по changed_at: последнее изменение применяется первым.
  // Если оно содержит effective_from в прошлом — оно корректно применится
  // ко всем дням >= effective_from, даже к тем, что попадали под более старые ставки.
  async function loadRateHistoryForUsers(userIds) {
    if (userIds.length === 0) return {};
    const key = 'rateHistory:' + [...userIds].sort().join(',');
    return cachedCall(key, CACHE_TTL_MS, async () => {
      const { data } = await sb
        .from('rate_history')
        .select('user_id, language_pair_id, new_rate, changed_at, effective_from')
        .in('user_id', userIds)
        .order('changed_at', { ascending: false });

      const byUser = {};
      for (const h of (data || [])) {
        if (!byUser[h.user_id]) byUser[h.user_id] = {};
        if (!byUser[h.user_id][h.language_pair_id]) byUser[h.user_id][h.language_pair_id] = [];
        byUser[h.user_id][h.language_pair_id].push(h);
      }
      return byUser;
    });
  }

  // Универсальная функция расчёта дня с учётом исторических ставок.
  // Возвращает { netMinutes, amount }.
  function calcDayWithHistory(day, historyByPair, currentRates) {
    if (day.day_type !== 'working') return { netMinutes: 0, amount: 0 };
    const intervals = day.work_intervals || [];
    const breaks = day.breaks || [];
    const intervalsSum = intervals.reduce((s, i) => s + (i.duration_minutes || 0), 0);
    const breaksSum = breaks.reduce((s, b) => s + (b.duration_minutes || 0), 0);
    const netMin = Math.max(0, intervalsSum - breaksSum);
    if (netMin <= 0) return { netMinutes: 0, amount: 0 };

    // Брейки распределяются пропорционально
    const breakRatio = intervalsSum > 0 ? (netMin / intervalsSum) : 0;
    let amount = 0;
    for (const interv of intervals) {
      const rate = getRateForDate(day.work_date, interv.language_pair_id, historyByPair, currentRates);
      const intervalNetMin = (interv.duration_minutes || 0) * breakRatio;
      amount += (intervalNetMin / 60) * rate;
    }
    return { netMinutes: netMin, amount };
  }

  // Загружает смены (shifts) для группы переводчиков за период.
  // Возвращает { user_id: { date: planned_minutes } }
  async function loadShiftsForUsers(userIds, periodStart, periodEnd) {
    if (userIds.length === 0) return {};
    const key = 'shifts:' + periodStart + ':' + periodEnd + ':' + [...userIds].sort().join(',');
    return cachedCall(key, CACHE_TTL_MS, async () => {
      const { data } = await sb.from('shifts')
        .select('user_id, shift_date, planned_minutes')
        .in('user_id', userIds)
        .gte('shift_date', periodStart)
        .lte('shift_date', periodEnd);

      const byUser = {};
      for (const s of (data || [])) {
        if (!byUser[s.user_id]) byUser[s.user_id] = {};
        byUser[s.user_id][s.shift_date] = s.planned_minutes;
      }
      return byUser;
    });
  }

  // Считает план для конкретной даты переводчика
  // shifts: { date: planned_minutes }
  // defaultMin: дефолтные минуты переводчика (fallback)
  function getPlannedMinutes(workDate, shiftsByDate, defaultMin) {
    if (shiftsByDate && shiftsByDate[workDate] !== undefined) {
      return shiftsByDate[workDate];
    }
    return defaultMin;
  }

  // Расчёт с разбивкой по парам ранее делался функцией calcDayWithBreakdown.
  // Теперь эту работу выполняет SQL-функция payroll_for_period (см. миграцию
  // payroll_for_period.sql) — серверный расчёт быстрее и гарантирует
  // консистентность ведомости на лету и в payroll_entries.

  // ====================================================================
  // КАЛЕНДАРЬ (для переводчика)
  // ====================================================================
  let calYear = new Date().getFullYear();
  let calMonth = new Date().getMonth(); // 0-11

  const MONTH_NAMES_RU = [
    'Январь','Февраль','Март','Апрель','Май','Июнь',
    'Июль','Август','Сентябрь','Октябрь','Ноябрь','Декабрь'
  ];
  const DAY_NAMES = ['Пн','Вт','Ср','Чт','Пт','Сб','Вс'];

  // Состояние календаря переводчика — для обновления одной ячейки после сохранения дня
  // без перерисовки всего месяца. Восстанавливается при каждом loadCalendar().
  let calState = null;

  async function loadCalendar() {
    hideError('cal-error');

    document.getElementById('cal-page-title').textContent =
      MONTH_NAMES_RU[calMonth] + ' ' + calYear;
    document.getElementById('cal-month-label').textContent =
      MONTH_NAMES_RU[calMonth] + ' ' + calYear;

    const periodStart = formatDate(new Date(calYear, calMonth, 1));
    const periodEnd = formatDate(new Date(calYear, calMonth + 1, 0));

    // Параллелизуем три независимых запроса — раньше шли последовательно
    const [daysRes, pairsRes, historyByUser] = await Promise.all([
      sb.from('work_days')
        .select(`
          id, work_date, day_type, note,
          work_intervals ( id, duration_minutes, language_pair_id ),
          breaks ( id, duration_minutes )
        `)
        .eq('user_id', currentUser.id)
        .gte('work_date', periodStart)
        .lte('work_date', periodEnd),
      sb.from('translator_pairs')
        .select('language_pair_id, rate_per_hour')
        .eq('user_id', currentUser.id),
      loadRateHistoryForUsers([currentUser.id]),
    ]);

    if (daysRes.error) {
      showError('cal-error', 'Ошибка загрузки: ' + daysRes.error.message);
      return;
    }
    const days = daysRes.data;
    const myPairs = pairsRes.data;

    const rateByPair = {};
    for (const p of (myPairs || [])) {
      rateByPair[p.language_pair_id] = Number(p.rate_per_hour);
    }
    const historyByPair = historyByUser[currentUser.id] || {};

    const dayMap = {};
    let totalMinutes = 0;
    let totalAmount = 0;
    let absentDays = 0;
    let workingDays = 0;

    for (const d of days || []) {
      if (d.day_type === 'absent') {
        dayMap[d.work_date] = { type: 'absent' };
        absentDays++;
        continue;
      }
      const calc = calcDayWithHistory(d, historyByPair, rateByPair);
      if (calc.netMinutes > 0) {
        dayMap[d.work_date] = { type: 'working', minutes: calc.netMinutes, amount: calc.amount };
        totalMinutes += calc.netMinutes;
        totalAmount += calc.amount;
        workingDays++;
      }
    }

    // Сохраняем состояние — пригодится для refreshCalendarDay
    calState = {
      year: calYear, month: calMonth,
      dayMap, rateByPair, historyByPair,
      totals: { minutes: totalMinutes, amount: totalAmount, workingDays, absentDays },
    };

    updateCalendarKpis();
    renderCalendar(dayMap);
  }

  // Обновляет блок KPI календаря из calState — отдельной функцией, чтобы вызывать
  // и после полной загрузки, и после single-cell обновления.
  function updateCalendarKpis() {
    if (!calState) return;
    const { totals } = calState;
    document.getElementById('kpi-hours').innerHTML =
      formatHoursMinutes(totals.minutes) + '<span class="kpi-unit"></span>';
    document.getElementById('kpi-hours-meta').textContent =
      `${totals.workingDays} ${pluralize(totals.workingDays, 'рабочий день', 'рабочих дня', 'рабочих дней')}`;
    document.getElementById('kpi-absent').innerHTML =
      totals.absentDays + '<span class="kpi-unit">' +
      (totals.absentDays === 0 ? 'дней' : pluralize(totals.absentDays, 'день', 'дня', 'дней')) +
      '</span>';
    document.getElementById('kpi-amount').textContent = '$' + totals.amount.toFixed(2);
  }

  // Обновление одного дня в календаре без полной перезагрузки месяца.
  // Вызывается после saveDay/deleteDay — даёт мгновенную обратную связь.
  async function refreshCalendarDay(dateStr) {
    // Если состояния нет или мы вне текущего месяца — fallback на полную загрузку
    if (!calState || calState.year !== calYear || calState.month !== calMonth) {
      return loadCalendar();
    }
    const periodStart = formatDate(new Date(calYear, calMonth, 1));
    const periodEnd = formatDate(new Date(calYear, calMonth + 1, 0));
    if (dateStr < periodStart || dateStr > periodEnd) return loadCalendar();

    // Запрашиваем только этот день
    const { data: day, error } = await sb
      .from('work_days')
      .select(`
        id, work_date, day_type, note,
        work_intervals ( id, duration_minutes, language_pair_id ),
        breaks ( id, duration_minutes )
      `)
      .eq('user_id', currentUser.id)
      .eq('work_date', dateStr)
      .maybeSingle();
    if (error) return loadCalendar();

    // Откатываем вклад старого состояния
    const old = calState.dayMap[dateStr];
    if (old) {
      if (old.type === 'working') {
        calState.totals.minutes -= old.minutes;
        calState.totals.amount -= old.amount;
        calState.totals.workingDays -= 1;
      } else if (old.type === 'absent') {
        calState.totals.absentDays -= 1;
      }
      delete calState.dayMap[dateStr];
    }

    // Применяем новое
    if (day) {
      if (day.day_type === 'absent') {
        calState.dayMap[dateStr] = { type: 'absent' };
        calState.totals.absentDays += 1;
      } else {
        const calc = calcDayWithHistory(day, calState.historyByPair, calState.rateByPair);
        if (calc.netMinutes > 0) {
          calState.dayMap[dateStr] = { type: 'working', minutes: calc.netMinutes, amount: calc.amount };
          calState.totals.minutes += calc.netMinutes;
          calState.totals.amount += calc.amount;
          calState.totals.workingDays += 1;
        }
      }
    }

    // Перерисовываем только эту ячейку
    const oldCell = document.querySelector(`#calendar-grid .cal-day[data-date="${dateStr}"]`);
    if (oldCell) {
      const dayNum = Number(dateStr.split('-')[2]);
      const todayStr = formatDate(new Date());
      const newCell = buildCalendarCell(dayNum, dateStr, calYear, calMonth, calState.dayMap, todayStr);
      newCell.addEventListener('click', () => onDayClick(dateStr));
      oldCell.replaceWith(newCell);
    }
    updateCalendarKpis();
  }

  // Рендер одной ячейки календаря (используется и при полной перерисовке, и при single-cell update)
  // Возвращает DOM-узел. dayMap — общий словарь, calYear/calMonth — текущие глобальные.
  function buildCalendarCell(d, dateStr, calYear_, calMonth_, dayMap, todayStr) {
    const cell = document.createElement('div');
    cell.className = 'cal-day';
    cell.dataset.date = dateStr;

    const dayOfWeek = new Date(calYear_, calMonth_, d).getDay();
    const isWeekend = (dayOfWeek === 0 || dayOfWeek === 6);
    const info = dayMap[dateStr];

    if (info && info.type === 'working') cell.classList.add('has-hours');
    else if (info && info.type === 'absent') cell.classList.add('absent');
    else if (isWeekend) cell.classList.add('weekend');

    if (dateStr === todayStr) cell.classList.add('today');

    const num = document.createElement('div');
    num.className = 'cal-day-num';
    num.textContent = d;
    cell.appendChild(num);

    if (info && info.type === 'working') {
      const h = document.createElement('div');
      h.className = 'cal-day-hours';
      h.textContent = formatHoursMinutes(info.minutes);
      cell.appendChild(h);
    } else if (info && info.type === 'absent') {
      const t = document.createElement('div');
      t.className = 'cal-day-tag';
      t.textContent = 'Отгул';
      cell.appendChild(t);
    }

    return cell;
  }

  function renderCalendar(dayMap) {
    const grid = document.getElementById('calendar-grid');
    // DocumentFragment — все вставки в один reflow вместо ~37
    const frag = document.createDocumentFragment();

    DAY_NAMES.forEach(n => {
      const el = document.createElement('div');
      el.className = 'cal-day-name';
      el.textContent = n;
      frag.appendChild(el);
    });

    const firstDay = new Date(calYear, calMonth, 1);
    // Сдвигаем понедельник = 0
    let dow = firstDay.getDay() - 1;
    if (dow < 0) dow = 6;

    for (let i = 0; i < dow; i++) {
      const empty = document.createElement('div');
      empty.className = 'cal-day empty';
      frag.appendChild(empty);
    }

    const daysInMonth = new Date(calYear, calMonth + 1, 0).getDate();
    const todayStr = formatDate(new Date());

    for (let d = 1; d <= daysInMonth; d++) {
      const dateStr = `${calYear}-${String(calMonth + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
      const cell = buildCalendarCell(d, dateStr, calYear, calMonth, dayMap, todayStr);
      cell.addEventListener('click', () => onDayClick(dateStr));
      frag.appendChild(cell);
    }

    grid.replaceChildren(frag);
  }

  function onDayClick(dateStr) {
    openDayModal(dateStr);
  }

  // ====================================================================
  // ЗАПРОСЫ НА ИЗМЕНЕНИЕ СМЕНЫ
  // ====================================================================

  // ──── Переводчик: открыть модалку запроса изменения графика ──────────
  async function openShiftRequestModal() {
    hideError('shift-req-error');

    // Загружаем текущий дефолтный график
    const { data: profile } = await sb.from('users')
      .select('default_shift_start, default_shift_end, default_shift_minutes')
      .eq('id', currentUser.id)
      .single();

    const currentStart = (profile?.default_shift_start || '09:00:00').substring(0, 5);
    const currentEnd = (profile?.default_shift_end || '18:00:00').substring(0, 5);
    const currentMin = profile?.default_shift_minutes || 480;

    document.getElementById('shift-req-current').textContent =
      `${currentStart} – ${currentEnd} · ${formatHoursMinutes(currentMin)}`;

    // Предзаполняем форму текущим графиком
    document.getElementById('shift-req-start').value = currentStart;
    document.getElementById('shift-req-end').value = currentEnd;

    // Дата начала действия — по умолчанию +7 дней от сегодня
    const inWeek = new Date();
    inWeek.setDate(inWeek.getDate() + 7);
    document.getElementById('shift-req-effective').value = inWeek.toISOString().split('T')[0];

    document.getElementById('shift-req-reason').value = '';
    updateShiftReqDuration();

    document.getElementById('shift-req-start').oninput = updateShiftReqDuration;
    document.getElementById('shift-req-end').oninput = updateShiftReqDuration;

    document.getElementById('shift-request-modal').classList.add('open');
  }

  function closeShiftRequestModal() {
    document.getElementById('shift-request-modal').classList.remove('open');
  }

  function updateShiftReqDuration() {
    const start = document.getElementById('shift-req-start').value;
    const end = document.getElementById('shift-req-end').value;
    const minutes = calcShiftMinutes(start, end);
    const display = document.getElementById('shift-req-duration');
    if (minutes > 0) {
      display.textContent = formatHoursMinutes(minutes);
      display.style.color = '#16A34A';
    } else {
      display.textContent = '—';
      display.style.color = '#94A3B8';
    }
  }

  async function sendShiftRequest() {
    hideError('shift-req-error');
    const btn = document.getElementById('btn-send-shift-req');

    const start = document.getElementById('shift-req-start').value;
    const end = document.getElementById('shift-req-end').value;
    const effectiveFrom = document.getElementById('shift-req-effective').value;
    const reason = document.getElementById('shift-req-reason').value.trim();
    const minutes = calcShiftMinutes(start, end);

    if (!start || !end || minutes <= 0) {
      showError('shift-req-error', 'Укажите корректное время.');
      return;
    }
    if (minutes > 16 * 60) {
      showError('shift-req-error', 'Смена не может быть больше 16 часов.');
      return;
    }
    if (!effectiveFrom) {
      showError('shift-req-error', 'Укажите дату начала действия нового графика.');
      return;
    }
    if (!reason) {
      showError('shift-req-error', 'Укажите причину изменения графика.');
      return;
    }

    btn.disabled = true; btn.textContent = 'Отправка…';

    try {
      // Проверим, нет ли уже pending запроса
      const { data: existing } = await sb
        .from('shift_change_requests')
        .select('id')
        .eq('user_id', currentUser.id)
        .eq('status', 'pending')
        .maybeSingle();

      if (existing) {
        showError('shift-req-error', 'У вас уже есть незавершённый запрос. Дождитесь решения менеджера.');
        return;
      }

      const { error } = await sb.from('shift_change_requests').insert({
        user_id: currentUser.id,
        shift_date: effectiveFrom,         // используем поле shift_date как effective_from
        requested_start: start + ':00',
        requested_end: end + ':00',
        requested_minutes: minutes,
        reason,
      });
      if (error) throw new Error(error.message);

      closeShiftRequestModal();
      alert('Запрос отправлен. Менеджер рассмотрит его в ближайшее время.');
      await loadMyProfile(); // обновим страницу профиля чтобы показать pending
    } catch (e) {
      showError('shift-req-error', e.message);
    } finally {
      btn.disabled = false; btn.textContent = 'Отправить запрос';
    }
  }

  // ──── Менеджер: страница запросов ──────────────────────────────────
  async function loadRequests() {
    hideError('req-error');
    const list = document.getElementById('req-list');
    list.innerHTML = '<div class="loading-state">Загрузка…</div>';

    const filter = document.getElementById('req-filter').value;
    let q = sb.from('shift_change_requests')
      .select(`
        id, shift_date, requested_start, requested_end, requested_minutes,
        final_start, final_end, final_minutes, final_effective_from,
        reason, status, reviewed_at, review_note, created_at,
        user_id,
        users:user_id ( name, email, default_shift_start, default_shift_end, default_shift_minutes )
      `)
      .order('created_at', { ascending: false });

    if (filter !== 'all') {
      q = q.eq('status', filter);
    }

    const { data: requests, error } = await q;
    if (error) {
      showError('req-error', 'Ошибка загрузки: ' + error.message);
      return;
    }

    // Подсчёт для подзаголовка
    const allReqs = await sb.from('shift_change_requests').select('status');
    const counts = { pending: 0, approved: 0, rejected: 0 };
    for (const r of (allReqs.data || [])) counts[r.status]++;
    document.getElementById('req-subtitle').textContent =
      `${counts.pending} ожидают · ${counts.approved} одобрено · ${counts.rejected} отклонено`;

    // Обновляем бейдж в сайдбаре
    updateRequestsBadge(counts.pending);

    if (!requests || requests.length === 0) {
      list.innerHTML = `
        <div class="empty-state">
          <div class="empty-state-text">Нет запросов по выбранному фильтру.</div>
        </div>
      `;
      return;
    }

    list.innerHTML = requests.map(r => renderRequestCard(r)).join('');
  }

  function renderRequestCard(r) {
    const user = r.users;
    const initials = (user.name || '?').split(' ').map(s => s[0]).join('').slice(0, 2).toUpperCase();

    // Текущий график переводчика
    const ds = (user.default_shift_start || '09:00:00').substring(0, 5);
    const de = (user.default_shift_end || '18:00:00').substring(0, 5);
    const dm = user.default_shift_minutes || 480;
    const currentLabel = `${ds}–${de} (${formatHoursMinutes(dm)})`;

    // Запрошенный график
    const rs = r.requested_start.substring(0,5);
    const re = r.requested_end.substring(0,5);
    const reqLabel = `${rs}–${re} (${formatHoursMinutes(r.requested_minutes)})`;

    // Дата начала
    const effDate = formatDateRu(r.shift_date);

    if (r.status === 'pending') {
      // Pending: показываем форму редактирования с возможностью одобрить
      return `
        <div class="req-card pending">
          <div class="req-avatar">${initials}</div>
          <div class="req-info" style="grid-column: 2 / -1;">
            <div style="display: flex; justify-content: space-between; align-items: baseline;">
              <div>
                <div class="req-translator-name">${escapeHtml(user.name)}</div>
                <div class="req-date">Подал ${new Date(r.created_at).toLocaleDateString('ru-RU')}</div>
              </div>
              <span class="req-status-badge pending">● Ожидает</span>
            </div>

            <div class="req-change" style="margin-top: 10px;">
              <span class="req-change-from">Сейчас: ${currentLabel}</span>
              <span class="req-change-arrow">→</span>
              <span class="req-change-to">Хочет: ${reqLabel} с ${effDate}</span>
            </div>

            <div class="req-reason">${escapeHtml(r.reason)}</div>

            <div style="margin-top: 14px; padding: 12px; background: #FAFAFA; border-radius: 8px;">
              <div style="font-size: 11px; font-weight: 600; color: #475569; margin-bottom: 8px; text-transform: uppercase; letter-spacing: 0.05em;">
                Решение менеджера · можете скорректировать перед одобрением
              </div>
              <div class="field-row">
                <div class="field" style="margin: 0;">
                  <label class="field-label">Начало</label>
                  <input type="time" id="rev-start-${r.id}" class="input" value="${rs}">
                </div>
                <div class="field" style="margin: 0;">
                  <label class="field-label">Окончание</label>
                  <input type="time" id="rev-end-${r.id}" class="input" value="${re}">
                </div>
                <div class="field" style="margin: 0;">
                  <label class="field-label">Действует с</label>
                  <input type="date" id="rev-date-${r.id}" class="input" value="${r.shift_date}">
                </div>
              </div>
              <div class="field" style="margin-top: 10px; margin-bottom: 0;">
                <label class="field-label">Заметка (опционально)</label>
                <input type="text" id="rev-note-${r.id}" class="input" placeholder="Например: согласовано до конца квартала">
              </div>
              <div style="display: flex; gap: 8px; margin-top: 12px; justify-content: flex-end;">
                <button class="btn btn-ghost btn-sm" onclick="reviewRequest('${r.id}', 'rejected')">Отклонить</button>
                <button class="btn btn-sm" onclick="reviewRequest('${r.id}', 'approved')">Одобрить с этими параметрами</button>
              </div>
            </div>
          </div>
        </div>
      `;
    }

    // approved / rejected: только просмотр
    let finalLabel = '';
    if (r.status === 'approved' && r.final_start) {
      const fs = r.final_start.substring(0,5);
      const fe = r.final_end.substring(0,5);
      const fd = r.final_effective_from ? formatDateRu(r.final_effective_from) : effDate;
      finalLabel = `<div class="req-change" style="margin-top: 8px; background: #DCFCE7;">
        <span class="req-change-from" style="color: #475569;">Утверждено:</span>
        <span class="req-change-to">${fs}–${fe} (${formatHoursMinutes(r.final_minutes)}) с ${fd}</span>
      </div>`;
    }
    if (r.review_note) {
      finalLabel += `<div class="req-reason" style="margin-top: 6px;">${escapeHtml(r.review_note)}</div>`;
    }

    const statusClass = r.status === 'approved' ? 'approved' : 'rejected';
    const statusLabel = r.status === 'approved' ? '● Одобрен' : '● Отклонён';
    return `
      <div class="req-card ${statusClass}">
        <div class="req-avatar">${initials}</div>
        <div class="req-info">
          <div class="req-translator-name">${escapeHtml(user.name)}</div>
          <div class="req-date">Подал ${new Date(r.created_at).toLocaleDateString('ru-RU')} · решено ${new Date(r.reviewed_at).toLocaleDateString('ru-RU')}</div>
          <div class="req-change">
            <span class="req-change-from">Запросил: ${reqLabel} с ${effDate}</span>
          </div>
          <div class="req-reason">${escapeHtml(r.reason)}</div>
          ${finalLabel}
        </div>
        <div class="req-actions">
          <span class="req-status-badge ${statusClass}">${statusLabel}</span>
        </div>
      </div>
    `;
  }

  function formatDateRu(dateStr) {
    const d = new Date(dateStr + 'T00:00:00');
    return `${d.getDate()} ${MONTH_NAMES_RU[d.getMonth()].toLowerCase()} ${d.getFullYear()}`;
  }

  async function reviewRequest(reqId, decision) {
    let finalStart, finalEnd, finalDate, finalMin, reviewNote;

    if (decision === 'approved') {
      // Читаем значения из инпутов
      finalStart = document.getElementById(`rev-start-${reqId}`).value;
      finalEnd = document.getElementById(`rev-end-${reqId}`).value;
      finalDate = document.getElementById(`rev-date-${reqId}`).value;
      reviewNote = document.getElementById(`rev-note-${reqId}`).value.trim();
      finalMin = calcShiftMinutes(finalStart, finalEnd);

      if (!finalStart || !finalEnd || finalMin <= 0) {
        alert('Укажите корректное время начала и окончания.');
        return;
      }
      if (!finalDate) {
        alert('Укажите дату начала действия графика.');
        return;
      }
    } else {
      reviewNote = prompt('Причина отклонения (опционально):') || '';
    }

    const action = decision === 'approved' ? 'одобрить' : 'отклонить';
    if (!confirm(`Точно ${action} запрос?`)) return;

    try {
      // Загружаем запрос
      const { data: req, error: reqErr } = await sb
        .from('shift_change_requests')
        .select('*')
        .eq('id', reqId)
        .single();
      if (reqErr || !req) throw new Error('Запрос не найден');

      // При одобрении: обновляем дефолт + записываем историю
      if (decision === 'approved') {
        // Текущий дефолт переводчика (для истории)
        const { data: u } = await sb.from('users')
          .select('default_shift_start, default_shift_end, default_shift_minutes')
          .eq('id', req.user_id)
          .single();

        // 1) Запись в shift_history (старый → новый)
        await sb.from('shift_history').insert({
          user_id: req.user_id,
          old_start: u?.default_shift_start,
          old_end: u?.default_shift_end,
          old_minutes: u?.default_shift_minutes,
          new_start: finalStart + ':00',
          new_end: finalEnd + ':00',
          new_minutes: finalMin,
          effective_from: finalDate,
          changed_by: currentUser.id,
          reason: `По запросу: ${req.reason}` + (reviewNote ? ` · ${reviewNote}` : ''),
          request_id: reqId,
        });

        // 2) Обновляем дефолтный график переводчика
        await sb.from('users').update({
          default_shift_start: finalStart + ':00',
          default_shift_end: finalEnd + ':00',
          default_shift_minutes: finalMin,
        }).eq('id', req.user_id);
      }

      // 3) Обновляем сам запрос
      const updatePayload = {
        status: decision,
        reviewed_by: currentUser.id,
        reviewed_at: new Date().toISOString(),
        review_note: reviewNote || null,
      };
      if (decision === 'approved') {
        updatePayload.final_start = finalStart + ':00';
        updatePayload.final_end = finalEnd + ':00';
        updatePayload.final_minutes = finalMin;
        updatePayload.final_effective_from = finalDate;
      }
      const { error: updErr } = await sb
        .from('shift_change_requests')
        .update(updatePayload)
        .eq('id', reqId);
      if (updErr) throw new Error(updErr.message);

      // Перезагружаем список
      await loadRequests();
    } catch (e) {
      alert('Ошибка: ' + e.message);
    }
  }

  // ──── Бейдж количества pending в сайдбаре ──────────────────────────
  async function refreshRequestsBadge() {
    if (!currentUser || currentUser.role !== 'manager') return;
    const { count } = await sb
      .from('shift_change_requests')
      .select('*', { count: 'exact', head: true })
      .eq('status', 'pending');
    updateRequestsBadge(count || 0);
  }

  function updateRequestsBadge(count) {
    const badge = document.getElementById('sb-req-badge');
    if (!badge) return;
    if (count > 0) {
      badge.textContent = String(count);
      badge.classList.remove('hidden');
    } else {
      badge.classList.add('hidden');
    }
  }

  // Подгружает данные о смене и обновляет блок-подсказку в модалке ввода дня (только просмотр)
  async function loadDayShiftInfo(dateStr) {
    const userId = (targetUser && targetUser.id) ? targetUser.id : currentUser.id;
    const [shiftRes, profileRes] = await Promise.all([
      sb.from('shifts')
        .select('planned_start, planned_end, planned_minutes')
        .eq('user_id', userId)
        .eq('shift_date', dateStr)
        .maybeSingle(),
      sb.from('users')
        .select('default_shift_start, default_shift_end, default_shift_minutes')
        .eq('id', userId)
        .single()
    ]);

    const block = document.getElementById('day-shift-info');
    const label = document.getElementById('day-shift-label');
    const timeEl = document.getElementById('day-shift-time');
    const durEl = document.getElementById('day-shift-dur');

    let start, end, minutes;
    if (shiftRes.data) {
      start = shiftRes.data.planned_start.substring(0, 5);
      end = shiftRes.data.planned_end.substring(0, 5);
      minutes = shiftRes.data.planned_minutes;
      label.textContent = 'Особая смена на этот день';
      block.className = 'shift-section assigned';
    } else {
      start = (profileRes.data?.default_shift_start || '09:00:00').substring(0, 5);
      end = (profileRes.data?.default_shift_end || '18:00:00').substring(0, 5);
      minutes = profileRes.data?.default_shift_minutes || 480;
      label.textContent = 'Стандартная смена';
      block.className = 'shift-section';
    }
    timeEl.textContent = `${start} – ${end}`;
    durEl.textContent = `· ${formatHoursMinutes(minutes)} плана`;
  }

  // ====================================================================
  // МОДАЛКА: ВВОД РАБОЧЕГО ДНЯ
  // ====================================================================
  let currentDayDate = null;
  let currentDayId = null;
  let currentDayType = 'working';
  let intervalRowCounter = 0;
  let breakRowCounter = 0;
  let myLanguagePairs = []; // [{ language_pair_id, code, rate_per_hour }]

  // Кого редактируем (себя или переводчика-за-кого-менеджер вводит)
  let targetUser = null; // объект пользователя, чей день мы редактируем
  let targetUserPairs = []; // языковые пары этого пользователя

  async function loadUserLanguagePairs(userId) {
    const { data } = await sb
      .from('translator_pairs')
      .select(`
        language_pair_id, rate_per_hour, is_primary,
        language_pairs ( code, display_name )
      `)
      .eq('user_id', userId);
    return (data || [])
      .sort((a, b) => (b.is_primary ? 1 : 0) - (a.is_primary ? 1 : 0))
      .map(p => ({
        language_pair_id: p.language_pair_id,
        code: p.language_pairs.code,
        rate_per_hour: p.rate_per_hour
      }));
  }

  async function ensureMyLanguagePairs() {
    if (myLanguagePairs.length > 0) return;
    myLanguagePairs = await loadUserLanguagePairs(currentUser.id);
  }

  // Открыть модалку ввода дня
  // forUser: если задан — менеджер вводит за переводчика. Иначе переводчик за себя.
  async function openDayModal(dateStr, forUser = null) {
    // Определяем target user
    if (forUser) {
      targetUser = forUser;
      targetUserPairs = await loadUserLanguagePairs(forUser.id);
    } else {
      // Переводчик за себя
      await ensureMyLanguagePairs();
      targetUser = currentUser;
      targetUserPairs = myLanguagePairs;
    }

    if (targetUserPairs.length === 0) {
      const msg = forUser
        ? `У переводчика ${targetUser.name} нет назначенных языковых пар. Сначала добавьте пары.`
        : 'У вас нет назначенных языковых пар. Обратитесь к менеджеру.';
      alert(msg);
      return;
    }

    currentDayDate = dateStr;
    currentDayId = null;
    currentDayType = 'working';
    intervalRowCounter = 0;
    breakRowCounter = 0;

    const date = new Date(dateStr + 'T00:00:00');
    const dayName = ['Воскресенье','Понедельник','Вторник','Среда','Четверг','Пятница','Суббота'][date.getDay()];
    const monthName = MONTH_NAMES_RU[date.getMonth()];
    let title = `${date.getDate()} ${monthName.toLowerCase()}`;
    if (forUser) title += ` · ${targetUser.name}`;
    document.getElementById('day-modal-title').textContent = title;
    document.getElementById('day-modal-subtitle').textContent =
      forUser ? `${dayName} · ввод менеджером` : dayName;

    document.getElementById('intervals-list').innerHTML = '';
    document.getElementById('breaks-list').innerHTML = '';
    document.getElementById('btn-delete-day').classList.add('hidden');
    document.getElementById('day-warnings').classList.add('hidden');
    hideError('day-error');
    setDayType('working');

    // Загружаем смену на этот день для подсказки
    await loadDayShiftInfo(dateStr);

    // Существующая запись?
    const { data: existingDay } = await sb
      .from('work_days')
      .select(`
        id, day_type, note,
        work_intervals ( id, language_pair_id, start_at, end_at, duration_minutes ),
        breaks ( id, start_at, end_at, duration_minutes, note )
      `)
      .eq('user_id', targetUser.id)
      .eq('work_date', dateStr)
      .maybeSingle();

    if (existingDay) {
      currentDayId = existingDay.id;
      setDayType(existingDay.day_type);
      document.getElementById('btn-delete-day').classList.remove('hidden');

      for (const interv of (existingDay.work_intervals || [])) {
        addIntervalRow({
          languagePairId: interv.language_pair_id,
          startTime: formatTimeHM(interv.start_at),
          endTime: formatTimeHM(interv.end_at),
        });
      }
      for (const br of (existingDay.breaks || [])) {
        addBreakRow({
          startTime: formatTimeHM(br.start_at),
          endTime: formatTimeHM(br.end_at),
          note: br.note || '',
        });
      }
      if ((existingDay.work_intervals || []).length === 0 && existingDay.day_type === 'working') {
        addIntervalRow();
      }
    } else {
      addIntervalRow();
    }

    updateDaySummary();
    document.getElementById('day-modal').classList.add('open');
  }

  function closeDayModal() {
    document.getElementById('day-modal').classList.remove('open');
    // Сбрасываем targetUser чтобы при следующем открытии переводчиком вернулось к currentUser
    targetUser = null;
    targetUserPairs = [];
  }

  function setDayType(type) {
    currentDayType = type;
    document.getElementById('chip-working').classList.toggle('active', type === 'working');
    document.getElementById('chip-absent').classList.toggle('active', type === 'absent');
    document.getElementById('working-content').classList.toggle('hidden', type !== 'working');
    document.getElementById('absent-content').classList.toggle('hidden', type !== 'absent');
    // При выборе "Отгул" — скрываем предупреждения и ошибки
    if (type === 'absent') {
      document.getElementById('day-warnings').classList.add('hidden');
      hideError('day-error');
    } else {
      updateDaySummary();
    }
  }

  function addIntervalRow(prefill = null) {
    intervalRowCounter++;
    const id = 'iv-' + intervalRowCounter;
    const pairs = (targetUserPairs && targetUserPairs.length > 0) ? targetUserPairs : myLanguagePairs;
    const options = pairs.map(p =>
      `<option value="${p.language_pair_id}" ${prefill && prefill.languagePairId === p.language_pair_id ? 'selected' : ''}>${p.code} ($${Number(p.rate_per_hour).toFixed(2)}/ч)</option>`
    ).join('');
    const div = document.createElement('div');
    div.className = 'interval-row';
    div.id = id;
    div.innerHTML = `
      <input type="time" class="iv-start" value="${prefill ? prefill.startTime : '09:00'}">
      <span class="arrow">→</span>
      <input type="time" class="iv-end" value="${prefill ? prefill.endTime : '18:00'}">
      <select class="iv-pair">${options}</select>
      <button class="remove-btn" onclick="removeIntervalRow('${id}')">×</button>
    `;
    document.getElementById('intervals-list').appendChild(div);
    div.querySelectorAll('input, select').forEach(el => {
      el.addEventListener('input', debouncedUpdateDaySummary);
      el.addEventListener('change', updateDaySummary);
    });
    updateDaySummary();
  }

  function removeIntervalRow(id) {
    const el = document.getElementById(id);
    if (el) el.remove();
    updateDaySummary();
  }

  function addBreakRow(prefill = null) {
    breakRowCounter++;
    const id = 'br-' + breakRowCounter;
    const div = document.createElement('div');
    div.className = 'interval-row break-row';
    div.id = id;
    div.innerHTML = `
      <input type="time" class="br-start" value="${prefill ? prefill.startTime : '13:00'}">
      <span class="arrow">→</span>
      <input type="time" class="br-end" value="${prefill ? prefill.endTime : '13:30'}">
      <input type="text" class="break-note" maxlength="200" placeholder="Заметка (необязательно)" value="${prefill ? escapeAttr(prefill.note) : ''}">
      <button class="remove-btn" onclick="removeBreakRow('${id}')">×</button>
    `;
    document.getElementById('breaks-list').appendChild(div);
    div.querySelectorAll('input').forEach(el => {
      el.addEventListener('input', debouncedUpdateDaySummary);
    });
    updateDaySummary();
  }

  function removeBreakRow(id) {
    const el = document.getElementById(id);
    if (el) el.remove();
    updateDaySummary();
  }

  // ─── ВАЛИДАЦИЯ ─────────────────────────────────────────────────────────
  // Возвращает все нарушения, найденные в текущей форме.
  // Также подсвечивает проблемные строки CSS-классом 'invalid'.
  function validateDayForm() {
    const warnings = [];
    const invalidRows = new Set();

    // Сбрасываем подсветку
    document.querySelectorAll('#intervals-list .interval-row, #breaks-list .interval-row')
      .forEach(r => r.classList.remove('invalid'));

    if (currentDayType !== 'working') return { warnings, invalidRows };

    // Собираем интервалы и брейки в виде [start, end] минут от начала currentDayDate
    // (с учётом перехода через полночь — конец может быть > 24h)
    const intervals = []; // { rowId, start, end, idx, langId }
    const breaks = [];    // { rowId, start, end, idx }

    const intervalRows = document.querySelectorAll('#intervals-list .interval-row');
    intervalRows.forEach((row, idx) => {
      const s = row.querySelector('.iv-start').value;
      const e = row.querySelector('.iv-end').value;
      const langId = row.querySelector('.iv-pair').value;
      if (!s || !e) {
        warnings.push(`Интервал #${idx + 1}: укажите время начала и окончания.`);
        invalidRows.add(row.id);
        return;
      }
      if (!langId) {
        warnings.push(`Интервал #${idx + 1}: выберите языковую пару.`);
        invalidRows.add(row.id);
        return;
      }
      const startMin = hmToMinutes(s);
      let endMin = hmToMinutes(e);
      if (endMin <= startMin) endMin += 24 * 60;
      if (endMin - startMin <= 0) {
        warnings.push(`Интервал #${idx + 1}: время окончания должно быть позже начала.`);
        invalidRows.add(row.id);
        return;
      }
      intervals.push({ rowId: row.id, start: startMin, end: endMin, idx: idx + 1, langId });
    });

    const breakRows = document.querySelectorAll('#breaks-list .interval-row');
    breakRows.forEach((row, idx) => {
      const s = row.querySelector('.br-start').value;
      const e = row.querySelector('.br-end').value;
      if (!s || !e) {
        warnings.push(`Брейк #${idx + 1}: укажите время начала и окончания.`);
        invalidRows.add(row.id);
        return;
      }
      const startMin = hmToMinutes(s);
      let endMin = hmToMinutes(e);
      if (endMin <= startMin) endMin += 24 * 60;
      if (endMin - startMin <= 0) {
        warnings.push(`Брейк #${idx + 1}: время окончания должно быть позже начала.`);
        invalidRows.add(row.id);
        return;
      }
      breaks.push({ rowId: row.id, start: startMin, end: endMin, idx: idx + 1 });
    });

    // ── Пересечения интервалов между собой ──────────────────────────────
    for (let i = 0; i < intervals.length; i++) {
      for (let j = i + 1; j < intervals.length; j++) {
        const a = intervals[i], b = intervals[j];
        if (overlaps(a.start, a.end, b.start, b.end)) {
          warnings.push(`Интервалы #${a.idx} и #${b.idx} пересекаются по времени.`);
          invalidRows.add(a.rowId);
          invalidRows.add(b.rowId);
        }
      }
    }

    // ── Пересечения брейков между собой ─────────────────────────────────
    for (let i = 0; i < breaks.length; i++) {
      for (let j = i + 1; j < breaks.length; j++) {
        const a = breaks[i], b = breaks[j];
        if (overlaps(a.start, a.end, b.start, b.end)) {
          warnings.push(`Брейки #${a.idx} и #${b.idx} пересекаются по времени.`);
          invalidRows.add(a.rowId);
          invalidRows.add(b.rowId);
        }
      }
    }

    // ── Каждый брейк должен быть ВНУТРИ хотя бы одного интервала ────────
    // Учитываем ночные смены: интервал может быть сдвинут на +24ч
    // (если start > end → end += 1440). Брейк, нарисованный во второй
    // (послеполночной) части смены, имеет числа меньше start интервала,
    // и тогда нужно сравнивать брейк со сдвигом +24ч.
    for (const br of breaks) {
      const isInside = intervals.some(iv => {
        // Прямая проверка: брейк лежит в [iv.start, iv.end]
        if (iv.start <= br.start && br.end <= iv.end) return true;
        // Ночная смена: пробуем сдвинуть брейк на +24ч
        // (актуально, если интервал пересекает полночь, т.е. iv.end > 1440)
        if (iv.end > 24 * 60) {
          const brStartShifted = br.start + 24 * 60;
          const brEndShifted = br.end + 24 * 60;
          if (iv.start <= brStartShifted && brEndShifted <= iv.end) return true;
        }
        return false;
      });
      if (!isInside && intervals.length > 0) {
        warnings.push(`Брейк #${br.idx} находится вне рабочих интервалов.`);
        invalidRows.add(br.rowId);
      }
    }

    // Применяем подсветку
    invalidRows.forEach(rowId => {
      const el = document.getElementById(rowId);
      if (el) el.classList.add('invalid');
    });

    return { warnings, invalidRows };
  }

  // Пересечение интервалов [a1, a2] и [b1, b2] (с открытыми границами)
  function overlaps(a1, a2, b1, b2) {
    return a1 < b2 && b1 < a2;
  }

  function hmToMinutes(hm) {
    const [h, m] = hm.split(':').map(Number);
    return h * 60 + m;
  }

  function renderWarnings(warnings) {
    const block = document.getElementById('day-warnings');
    const list = document.getElementById('day-warnings-list');
    if (warnings.length === 0) {
      block.classList.add('hidden');
      list.innerHTML = '';
      return;
    }
    list.innerHTML = warnings.map(w => `<li>${escapeHtml(w)}</li>`).join('');
    block.classList.remove('hidden');
  }

  function updateDaySummary() {
    let grossMin = 0, breaksMin = 0;
    document.querySelectorAll('#intervals-list .interval-row').forEach(row => {
      const s = row.querySelector('.iv-start').value;
      const e = row.querySelector('.iv-end').value;
      const dur = minutesBetween(s, e);
      if (dur > 0) grossMin += dur;
    });
    document.querySelectorAll('#breaks-list .interval-row').forEach(row => {
      const s = row.querySelector('.br-start').value;
      const e = row.querySelector('.br-end').value;
      const dur = minutesBetween(s, e);
      if (dur > 0) breaksMin += dur;
    });
    const netMin = Math.max(0, grossMin - breaksMin);
    document.getElementById('sum-gross').textContent = formatHoursMinutes(grossMin);
    document.getElementById('sum-breaks').textContent = '−' + formatHoursMinutes(breaksMin);
    document.getElementById('sum-net').textContent = formatHoursMinutes(netMin);

    // Запускаем валидацию и показываем предупреждения
    const { warnings } = validateDayForm();
    renderWarnings(warnings);
  }
  // Debounced версия — навешивается на input-события (печать в полях времени),
  // чтобы не перевалидировать форму на каждое нажатие клавиши.
  const debouncedUpdateDaySummary = debounce(updateDaySummary, 120);

  function minutesBetween(startHM, endHM) {
    if (!startHM || !endHM) return 0;
    const [sh, sm] = startHM.split(':').map(Number);
    const [eh, em] = endHM.split(':').map(Number);
    let startMin = sh * 60 + sm;
    let endMin = eh * 60 + em;
    if (endMin <= startMin) endMin += 24 * 60; // переход через полночь
    return endMin - startMin;
  }

  async function saveDay() {
    hideError('day-error');
    const btn = document.getElementById('btn-save-day');

    // Запускаем валидацию для рабочего дня
    if (currentDayType === 'working') {
      const { warnings } = validateDayForm();
      if (warnings.length > 0) {
        showError('day-error', 'Исправьте ошибки перед сохранением: ' + warnings.length + ' замечание(й) ниже.');
        renderWarnings(warnings);
        return;
      }
    }

    btn.disabled = true; btn.textContent = 'Сохранение…';
    try {
      const isManagerEdit = targetUser && targetUser.id !== currentUser.id;

      if (currentDayType === 'absent') {
        await upsertDay({ day_type: 'absent', intervals: [], breaks: [] });
        const savedDate = currentDayDate;
        closeDayModal();
        if (isManagerEdit) {
          await refreshTdCalendarDay(savedDate);
        } else {
          await refreshCalendarDay(savedDate);
        }
        return;
      }
      const intervals = [];
      let basicValid = true;
      document.querySelectorAll('#intervals-list .interval-row').forEach(row => {
        const s = row.querySelector('.iv-start').value;
        const e = row.querySelector('.iv-end').value;
        const langId = row.querySelector('.iv-pair').value;
        const dur = minutesBetween(s, e);
        if (!s || !e || !langId) basicValid = false;
        if (dur <= 0) basicValid = false;
        intervals.push({ start: s, end: e, language_pair_id: langId, duration: dur });
      });
      const breaks = [];
      document.querySelectorAll('#breaks-list .interval-row').forEach(row => {
        const s = row.querySelector('.br-start').value;
        const e = row.querySelector('.br-end').value;
        const note = row.querySelector('.break-note').value.trim() || null;
        const dur = minutesBetween(s, e);
        if (!s || !e) basicValid = false;
        if (dur <= 0) basicValid = false;
        breaks.push({ start: s, end: e, duration: dur, note });
      });
      if (intervals.length === 0) {
        showError('day-error', 'Добавьте хотя бы один рабочий интервал, либо отметьте день как отгул.');
        return;
      }
      if (!basicValid) {
        showError('day-error', 'Проверьте поля: время и языковая пара заполнены, длительность > 0.');
        return;
      }
      await upsertDay({ day_type: 'working', intervals, breaks });
      const savedDate = currentDayDate;
      closeDayModal();
      if (isManagerEdit) {
        await refreshTdCalendarDay(savedDate);
      } else {
        await refreshCalendarDay(savedDate);
      }
    } catch (e) {
      showError('day-error', 'Ошибка: ' + e.message);
    } finally {
      btn.disabled = false; btn.textContent = 'Сохранить';
    }
  }

  async function upsertDay({ day_type, intervals, breaks }) {
    const userId = (targetUser && targetUser.id) ? targetUser.id : currentUser.id;
    let workDayId = currentDayId;
    if (!workDayId) {
      const { data: newDay, error: insertErr } = await sb
        .from('work_days')
        .insert({ user_id: userId, work_date: currentDayDate, day_type })
        .select('id').single();
      if (insertErr) throw new Error(insertErr.message);
      workDayId = newDay.id;
    } else {
      const { error: updErr } = await sb
        .from('work_days').update({ day_type }).eq('id', workDayId);
      if (updErr) throw new Error(updErr.message);
      await sb.from('work_intervals').delete().eq('work_day_id', workDayId);
      await sb.from('breaks').delete().eq('work_day_id', workDayId);
    }

    if (day_type === 'absent') return;

    if (intervals.length > 0) {
      const intervalRows = intervals.map(i => ({
        work_day_id: workDayId,
        language_pair_id: i.language_pair_id,
        start_at: toUtcISO(currentDayDate, i.start),
        end_at:   toUtcEndISO(currentDayDate, i.start, i.end),
        duration_minutes: i.duration,
      }));
      const { error } = await sb.from('work_intervals').insert(intervalRows);
      if (error) throw new Error(error.message);
    }
    if (breaks.length > 0) {
      const breakRows = breaks.map(b => ({
        work_day_id: workDayId,
        start_at: toUtcISO(currentDayDate, b.start),
        end_at:   toUtcEndISO(currentDayDate, b.start, b.end),
        duration_minutes: b.duration,
        note: b.note,
      }));
      const { error } = await sb.from('breaks').insert(breakRows);
      if (error) throw new Error(error.message);
    }
  }

  async function deleteDay() {
    if (!currentDayId) return;
    if (!confirm('Удалить запись за этот день? Все интервалы и брейки будут стёрты.')) return;
    const { error } = await sb.from('work_days').delete().eq('id', currentDayId);
    if (error) {
      showError('day-error', 'Ошибка удаления: ' + error.message);
      return;
    }
    const isManagerEdit = targetUser && targetUser.id !== currentUser.id;
    const deletedDate = currentDayDate;
    closeDayModal();
    if (isManagerEdit) {
      await refreshTdCalendarDay(deletedDate);
    } else {
      await refreshCalendarDay(deletedDate);
    }
  }

  // Локальное "HH:MM" на дату currentDayDate → UTC ISO для записи в БД
  function toUtcISO(dateStr, hm) {
    const [h, m] = hm.split(':').map(Number);
    const [y, mo, d] = dateStr.split('-').map(Number);
    return new Date(y, mo - 1, d, h, m, 0).toISOString();
  }
  function toUtcEndISO(dateStr, startHM, endHM) {
    const [sh, sm] = startHM.split(':').map(Number);
    const [eh, em] = endHM.split(':').map(Number);
    const [y, mo, d] = dateStr.split('-').map(Number);
    const sDate = new Date(y, mo - 1, d, sh, sm, 0);
    let eDate = new Date(y, mo - 1, d, eh, em, 0);
    if (eDate <= sDate) eDate = new Date(y, mo - 1, d + 1, eh, em, 0);
    return eDate.toISOString();
  }

  function formatTimeHM(isoString) {
    const d = new Date(isoString);
    return String(d.getHours()).padStart(2, '0') + ':' + String(d.getMinutes()).padStart(2, '0');
  }

  function escapeAttr(s) {
    return String(s || '').replace(/"/g, '&quot;');
  }

  function changeMonth(delta) {
    calMonth += delta;
    if (calMonth < 0) { calMonth = 11; calYear--; }
    if (calMonth > 11) { calMonth = 0; calYear++; }
    loadCalendar();
  }

  function formatDate(date) {
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, '0');
    const d = String(date.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }

  function formatHoursMinutes(totalMinutes) {
    const h = Math.floor(totalMinutes / 60);
    const m = totalMinutes % 60;
    return `${h}ч ${String(m).padStart(2, '0')}м`;
  }

  async function loadMyProfile() {
    await loadMyShiftBlock();
    await loadMyPairsTable();
  }

  async function loadMyShiftBlock() {
    const content = document.getElementById('my-shift-content');
    const btn = document.getElementById('btn-request-shift');

    // Параллельно: профиль (текущий график), pending запрос и ВСЯ история запросов
    const [profileRes, pendingReqRes, allRequestsRes] = await Promise.all([
      sb.from('users')
        .select('default_shift_start, default_shift_end, default_shift_minutes')
        .eq('id', currentUser.id)
        .single(),
      sb.from('shift_change_requests')
        .select('id, requested_start, requested_end, requested_minutes, shift_date, reason, created_at')
        .eq('user_id', currentUser.id)
        .eq('status', 'pending')
        .maybeSingle(),
      sb.from('shift_change_requests')
        .select(`id, status, created_at, reviewed_at, reason, review_note,
                 requested_start, requested_end, requested_minutes, shift_date,
                 final_start, final_end, final_minutes, final_effective_from`)
        .eq('user_id', currentUser.id)
        .order('created_at', { ascending: false })
        .limit(20)
    ]);

    const profile = profileRes.data;
    const start = (profile?.default_shift_start || '09:00:00').substring(0, 5);
    const end = (profile?.default_shift_end || '18:00:00').substring(0, 5);
    const minutes = profile?.default_shift_minutes || 480;

    let html = `
      <div class="shift-section">
        <div class="shift-section-header">
          <span class="shift-section-label">Действует сейчас</span>
        </div>
        <div class="shift-section-content">
          <span class="shift-section-time">${start} – ${end}</span>
          <span class="shift-section-duration">· ${formatHoursMinutes(minutes)} плана</span>
        </div>
      </div>
    `;

    // Pending запрос — выделенный блок
    if (pendingReqRes.data) {
      const r = pendingReqRes.data;
      const rs = r.requested_start.substring(0,5);
      const re = r.requested_end.substring(0,5);
      html += `
        <div style="background: #FEF3C7; border-left: 3px solid #B45309; padding: 10px 14px; border-radius: 8px; margin-top: 12px;">
          <div style="font-size: 11px; font-weight: 600; color: #B45309; text-transform: uppercase; letter-spacing: 0.05em; margin-bottom: 4px;">
            ⏱ Запрос на рассмотрении у менеджера
          </div>
          <div style="font-family: 'JetBrains Mono', monospace; font-size: 13px; color: #0F1B3D;">
            ${rs}–${re} (${formatHoursMinutes(r.requested_minutes)}) с ${formatDateRu(r.shift_date)}
          </div>
          <div style="font-style: italic; color: #475569; font-size: 12px; margin-top: 4px;">«${escapeHtml(r.reason)}»</div>
        </div>
      `;
      btn.textContent = '⏱ Запрос в ожидании';
      btn.disabled = true;
      btn.style.opacity = '0.5';
    } else {
      btn.textContent = '✎ Запросить изменение';
      btn.disabled = false;
      btn.style.opacity = '';
    }

    // ─── ПОЛНАЯ ИСТОРИЯ ЗАПРОСОВ ───────────────────────────────────────
    const allReqs = allRequestsRes.data || [];
    if (allReqs.length > 0) {
      html += `
        <div style="margin-top: 20px; padding-top: 14px; border-top: 1px solid #E5E7EB;">
          <div style="font-size: 11px; color: #94A3B8; text-transform: uppercase; letter-spacing: 0.05em; margin-bottom: 12px; font-family: 'JetBrains Mono', monospace; font-weight: 500;">
            История запросов
          </div>
          <div class="shift-history-list">
      `;

      for (const r of allReqs) {
        html += renderShiftRequestHistoryItem(r, false);
      }

      html += `</div></div>`;
    }

    content.innerHTML = html;
  }

  // Универсальный рендер строки истории запроса (для переводчика и менеджера)
  // isManagerView = true → шире карточка, добавляется имя переводчика (передаётся отдельно)
  function renderShiftRequestHistoryItem(r, isManagerView, translatorName = null) {
    const rs = r.requested_start.substring(0,5);
    const re = r.requested_end.substring(0,5);
    const reqDur = formatHoursMinutes(r.requested_minutes);
    const submitDate = new Date(r.created_at).toLocaleDateString('ru-RU');

    let statusLabel, statusClass, leftBorder;
    if (r.status === 'pending') {
      statusLabel = '● Ожидает';
      statusClass = 'pending';
      leftBorder = '#B45309';
    } else if (r.status === 'approved') {
      statusLabel = '● Одобрен';
      statusClass = 'approved';
      leftBorder = '#16A34A';
    } else {
      statusLabel = '● Отклонён';
      statusClass = 'rejected';
      leftBorder = '#DC2626';
    }

    let finalBlock = '';
    if (r.status === 'approved' && r.final_start) {
      const fs = r.final_start.substring(0,5);
      const fe = r.final_end.substring(0,5);
      const fDur = formatHoursMinutes(r.final_minutes);
      const fDate = r.final_effective_from ? formatDateRu(r.final_effective_from) : formatDateRu(r.shift_date);
      const sameAsRequested = (fs === rs && fe === re && r.final_minutes === r.requested_minutes
                               && r.final_effective_from === r.shift_date);
      if (sameAsRequested) {
        finalBlock = `<div style="font-size: 12px; color: #16A34A; margin-top: 4px;">✓ Одобрено как запрошено</div>`;
      } else {
        finalBlock = `<div style="font-size: 12px; color: #16A34A; margin-top: 4px; font-family: 'JetBrains Mono', monospace;">
          ✓ Утверждено: <strong>${fs}–${fe} (${fDur}) с ${fDate}</strong>
        </div>`;
      }
    } else if (r.status === 'rejected') {
      finalBlock = `<div style="font-size: 12px; color: #DC2626; margin-top: 4px;">✕ Отклонено</div>`;
    }

    let noteBlock = '';
    if (r.review_note) {
      noteBlock = `<div style="font-size: 12px; color: #475569; margin-top: 6px; font-style: italic;">
        Менеджер: «${escapeHtml(r.review_note)}»
      </div>`;
    }

    const decisionDate = r.reviewed_at
      ? ` · решено ${new Date(r.reviewed_at).toLocaleDateString('ru-RU')}`
      : '';

    return `
      <div style="padding: 10px 12px; border-left: 3px solid ${leftBorder}; background: #FAFAFA; border-radius: 6px; margin-bottom: 8px;">
        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 4px;">
          <div style="font-size: 12px; color: #94A3B8; font-family: 'JetBrains Mono', monospace;">
            ${isManagerView && translatorName ? `<strong style="color: #0F1B3D;">${escapeHtml(translatorName)}</strong> · ` : ''}
            подан ${submitDate}${decisionDate}
          </div>
          <span class="req-status-badge ${statusClass}">${statusLabel}</span>
        </div>
        <div style="font-family: 'JetBrains Mono', monospace; font-size: 13px; color: #0F1B3D;">
          Запросил: ${rs}–${re} (${reqDur}) с ${formatDateRu(r.shift_date)}
        </div>
        <div style="font-style: italic; color: #475569; font-size: 12px; margin-top: 4px;">«${escapeHtml(r.reason)}»</div>
        ${finalBlock}
        ${noteBlock}
      </div>
    `;
  }

  async function loadMyPairsTable() {
    const content = document.getElementById('my-profile-content');
    const { data, error } = await sb.from('translator_pairs')
      .select(`id, rate_per_hour, is_primary, proficiency, effective_from,
               language_pairs (code, display_name)`)
      .eq('user_id', currentUser.id);

    if (error) {
      content.innerHTML = `<div style="padding:24px;"><div class="error">Ошибка: ${error.message}</div></div>`;
      return;
    }
    if (!data || data.length === 0) {
      content.innerHTML = `
        <div class="empty-state">
          <div class="empty-state-title">У вас пока нет назначенных языковых пар</div>
          <div class="empty-state-text">Обратитесь к менеджеру для назначения языков и ставок.</div>
        </div>`;
      return;
    }

    const sorted = data.sort((a, b) => (b.is_primary ? 1 : 0) - (a.is_primary ? 1 : 0));
    let html = `<table><thead><tr>
      <th>Языковая пара</th><th>Уровень</th><th>Ставка / час</th><th>Действует с</th>
    </tr></thead><tbody>`;
    for (const p of sorted) {
      html += `<tr>
        <td><strong>${p.language_pairs.code}</strong>
            ${p.is_primary ? '<span class="badge badge-neutral" style="margin-left: 6px;">основная</span>' : ''}</td>
        <td>${p.proficiency || '—'}</td>
        <td style="font-family: 'JetBrains Mono', monospace; font-weight: 500;">$${Number(p.rate_per_hour).toFixed(2)}</td>
        <td style="font-family: 'JetBrains Mono', monospace; color: #475569;">${new Date(p.effective_from).toLocaleDateString('ru-RU')}</td>
      </tr>`;
    }
    html += '</tbody></table>';
    content.innerHTML = html;
  }

  function openAddModal() {
    document.getElementById('new-name').value = '';
    document.getElementById('new-email').value = '';
    document.getElementById('new-password').value = '';
    document.getElementById('new-timezone').value = 'Asia/Tashkent';
    // Дефолт effective_from — первое число текущего месяца
    const now = new Date();
    const firstOfMonth = new Date(now.getFullYear(), now.getMonth(), 1);
    const yyyy = firstOfMonth.getFullYear();
    const mm = String(firstOfMonth.getMonth() + 1).padStart(2, '0');
    const dd = '01';
    document.getElementById('new-effective-from').value = `${yyyy}-${mm}-${dd}`;
    document.getElementById('pairs-list').innerHTML = '';
    pairRowCounter = 0;
    addPairRow(true);
    hideError('add-error');
    document.getElementById('add-modal').classList.add('open');
  }
  function closeAddModal() {
    document.getElementById('add-modal').classList.remove('open');
  }

  function addPairRow(isPrimary = false) {
    pairRowCounter++;
    const id = 'pair-' + pairRowCounter;
    const options = languagePairs.map(p =>
      `<option value="${p.id}">${p.code} — ${p.display_name}</option>`
    ).join('');

    const div = document.createElement('div');
    div.className = 'pair-editor';
    div.id = id;
    div.innerHTML = `
      <select class="pair-lang">${options}</select>
      <div class="pair-rate-input">
        <span style="color:#475569;">$</span>
        <input type="number" class="pair-rate-input-field" step="0.50" min="0.50" max="500" placeholder="25.00">
        <span style="color:#475569;font-size:12px;">/ч</span>
      </div>
      <label class="primary-check">
        <input type="radio" name="primary-pair" value="${id}" ${isPrimary ? 'checked' : ''}>
        Основная
      </label>
      <button class="pair-remove" onclick="document.getElementById('${id}').remove()">×</button>
    `;
    document.getElementById('pairs-list').appendChild(div);
  }

  async function saveTranslator() {
    hideError('add-error');
    const name = document.getElementById('new-name').value.trim();
    const email = document.getElementById('new-email').value.trim();
    const password = document.getElementById('new-password').value;
    const timezone = document.getElementById('new-timezone').value;
    const effectiveFrom = document.getElementById('new-effective-from').value;

    if (!name || !email || !password) { showError('add-error', 'Заполните имя, email и пароль.'); return; }
    if (password.length < 6) { showError('add-error', 'Пароль должен быть не менее 6 символов.'); return; }
    if (!effectiveFrom) { showError('add-error', 'Укажите дату вступления ставки в силу.'); return; }

    const rows = document.querySelectorAll('#pairs-list .pair-editor');
    if (rows.length === 0) { showError('add-error', 'Добавьте хотя бы одну языковую пару.'); return; }

    const pairs = [];
    let primaryFound = false;
    for (const row of rows) {
      const langId = row.querySelector('.pair-lang').value;
      const rate = parseFloat(row.querySelector('.pair-rate-input-field').value);
      const isPrimary = row.querySelector('input[type="radio"]').checked;
      if (!rate || rate < 0.5 || rate > 500) {
        showError('add-error', 'Все ставки должны быть от $0.50 до $500.00.');
        return;
      }
      if (isPrimary) primaryFound = true;
      pairs.push({ language_pair_id: langId, rate_per_hour: rate, is_primary: isPrimary, proficiency: null });
    }
    if (!primaryFound) { showError('add-error', 'Отметьте одну пару как основную.'); return; }

    const codes = pairs.map(p => p.language_pair_id);
    if (new Set(codes).size !== codes.length) {
      showError('add-error', 'Одна и та же языковая пара указана несколько раз.');
      return;
    }

    const btn = document.getElementById('btn-save-translator');
    btn.disabled = true; btn.textContent = 'Сохраняем…';

    try {
      const { data: { session } } = await sb.auth.getSession();
      const response = await fetch(CREATE_TRANSLATOR_URL, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${session.access_token}`,
          'apikey': SUPABASE_ANON_KEY,
        },
        body: JSON.stringify({ email, password, name, timezone, pairs, effective_from: effectiveFrom })
      });

      const result = await response.json();
      if (!response.ok) {
        showError('add-error', result.error || 'Ошибка сервера.');
        return;
      }
      // Создан новый переводчик с парами и стартовой ставкой — чистим кеши
      invalidateCache('rateHistory:');
      closeAddModal();
      await loadTranslators();
    } catch (e) {
      showError('add-error', 'Ошибка сети: ' + e.message);
    } finally {
      btn.disabled = false; btn.textContent = 'Сохранить';
    }
  }

  // ════════════════════════════════════════════════════════════════════
  // МОДУЛЬ: КЛИЕНТЫ И ПРИБЫЛЬНОСТЬ
  // ════════════════════════════════════════════════════════════════════
  // Модель: тариф единый на (клиент, языковая пара). К одному клиенту можно
  // привязать несколько переводчиков (users.client_id). Резерв предоплаты —
  // индивидуальный на переводчика. RPC client_profitability возвращает массив
  // клиентов; у каждого агрегаты + вложенный массив translators[] со своей
  // разбивкой и своим резервом.
  // ════════════════════════════════════════════════════════════════════

  let clientsYear = new Date().getFullYear();
  let clientsMonth = new Date().getMonth();

  let clientDetailId = null;   // id клиента, открытого в детальной карточке
  let clientEditId = null;     // id клиента, редактируемого в edit-client-modal
  let clientRatePairRows = 0;  // счётчик строк тарифов в форме создания

  // ── Загрузка данных прибыльности (RPC) с кешем ──────────────────────
  async function fetchClientProfitability(year, month) {
    const periodStart = formatDate(new Date(year, month, 1));
    const periodEnd   = formatDate(new Date(year, month + 1, 0));
    const cacheKey = `clientProfit:${periodStart}:${periodEnd}`;
    return cachedCall(cacheKey, CACHE_TTL_MS, async () => {
      const { data, error } = await sb.rpc('client_profitability', {
        p_period_start: periodStart,
        p_period_end:   periodEnd,
      });
      if (error) throw new Error(error.message);
      return (data || []).map(normalizeClientRow);
    });
  }

  // Приведение числовых полей клиента и его переводчиков к Number
  function normalizeClientRow(row) {
    return {
      client_id:        row.client_id,
      client_name:      row.client_name,
      contact_email:    row.contact_email || '',
      revenue:          Number(row.revenue) || 0,
      cost:             Number(row.cost) || 0,
      profit:           Number(row.profit) || 0,
      margin_pct:       Number(row.margin_pct) || 0,
      translator_count: Number(row.translator_count) || 0,
      rates_by_pair:    row.rates_by_pair || [],
      training_summary: row.training_summary || null,
      translators:      (row.translators || []).map(normalizeTranslatorObj),
    };
  }

  function normalizeTranslatorObj(t) {
    return {
      user_id:            t.user_id,
      translator_name:    t.translator_name,
      is_active:          t.is_active !== false, // по умолчанию активен
      revenue:            Number(t.revenue) || 0,
      cost:               Number(t.cost) || 0,
      profit:             Number(t.profit) || 0,
      margin_pct:         Number(t.margin_pct) || 0,
      hours_period:       Number(t.hours_period) || 0,
      hours_purchased:    Number(t.hours_purchased) || 0,
      hours_worked_total: Number(t.hours_worked_total) || 0,
      hours_remaining:    Number(t.hours_remaining) || 0,
      total_paid:         Number(t.total_paid) || 0,
      unearned_loss:      Number(t.unearned_loss) || 0,
      training:           t.training || null,
      prepay_history:     t.prepay_history || [],
      breakdown:          t.breakdown || {},
    };
  }

  function clientsMonthLabel(y, m) {
    return MONTH_NAMES_RU[m] + ' ' + y;
  }

  function clientInitials(name) {
    return (name || '?').split(' ').map(s => s[0]).join('').slice(0, 2).toUpperCase();
  }

  // Режим резерва по остатку часов → { mode, badgeCls, barColor, textColor }
  function reserveStatus(hoursRemaining) {
    if (hoursRemaining <= 0) {
      return { mode: 'Постоплата', badgeCls: 'badge-warn', barColor: '#DC2626', textColor: '#DC2626' };
    }
    if (hoursRemaining <= 20) {
      return { mode: 'Мало резерва', badgeCls: 'badge-warn', barColor: '#B45309', textColor: '#B45309' };
    }
    return { mode: 'Предоплата', badgeCls: 'badge-good', barColor: '#16A34A', textColor: '#16A34A' };
  }

  // Сводный статус клиента для бейджа в списке. Приоритет:
  //   • убыток (неактивный с положительным остатком) → красный «Убыток»
  //   • резерв на исходе (активный, 0 < остаток ≤ 20) → жёлтый «Мало резерва»
  //   • иначе → нейтральный «—» (постоплата/норма ничем не выделяется)
  function clientReserveStatus(translators) {
    if (!translators || translators.length === 0) {
      return { mode: '—', badgeCls: 'badge-neutral', barColor: '#94A3B8', textColor: '#94A3B8' };
    }
    const hasLoss = translators.some(t => !t.is_active && t.hours_remaining > 0);
    if (hasLoss) {
      return { mode: 'Убыток', badgeCls: 'badge-warn', barColor: '#DC2626', textColor: '#DC2626' };
    }
    const hasLow = translators.some(t => t.is_active && t.hours_remaining > 0 && t.hours_remaining <= 20);
    if (hasLow) {
      return { mode: 'Мало резерва', badgeCls: 'badge-warn', barColor: '#B45309', textColor: '#B45309' };
    }
    return { mode: '—', badgeCls: 'badge-neutral', barColor: '#16A34A', textColor: '#16A34A' };
  }

  // ── СТРАНИЦА: СПИСОК КЛИЕНТОВ ───────────────────────────────────────
  async function loadClients() {
    hideError('clients-error');
    fillClientsMonthSelector();

    const content = document.getElementById('clients-content');
    content.innerHTML = '<div class="loading-state">Загрузка…</div>';
    document.getElementById('clients-subtitle').textContent =
      clientsMonthLabel(clientsYear, clientsMonth);

    let rows;
    try {
      rows = await fetchClientProfitability(clientsYear, clientsMonth);
    } catch (e) {
      content.innerHTML = '';
      showError('clients-error', 'Ошибка загрузки: ' + e.message);
      return;
    }

    // KPI по всем клиентам
    const totalRevenue = rows.reduce((s, r) => s + r.revenue, 0);
    const totalCost    = rows.reduce((s, r) => s + r.cost, 0);
    const totalProfit  = totalRevenue - totalCost;
    const totalMargin  = totalRevenue > 0 ? (totalProfit / totalRevenue * 100) : 0;
    // Резерв суммируем по всем переводчикам всех клиентов
    let totalReserve = 0, totalBought = 0;
    for (const r of rows) {
      for (const t of r.translators) {
        totalReserve += Math.max(0, t.hours_remaining);
        totalBought  += t.hours_purchased;
      }
    }

    document.getElementById('clients-kpi-revenue').textContent = '$' + totalRevenue.toFixed(2);
    document.getElementById('clients-kpi-cost').textContent = '$' + totalCost.toFixed(2);
    document.getElementById('clients-kpi-profit').textContent = '$' + totalProfit.toFixed(2);
    document.getElementById('clients-kpi-margin').textContent = 'маржа ' + totalMargin.toFixed(1) + '%';
    document.getElementById('clients-kpi-reserve').textContent = totalReserve.toFixed(0) + ' ч';
    document.getElementById('clients-kpi-reserve-meta').textContent =
      'из ' + totalBought.toFixed(0) + ' ч куплено';

    if (rows.length === 0) {
      content.innerHTML = `
        <div class="empty-state">
          <div class="empty-state-title">Клиентов пока нет</div>
          <div class="empty-state-text">Нажмите «+ Добавить клиента», чтобы создать первого клиента и задать тарифы по парам.</div>
        </div>`;
      return;
    }

    let html = `<table><thead><tr>
      <th>Клиент</th>
      <th class="numeric">Переводчиков</th>
      <th class="numeric">Доход</th>
      <th class="numeric">Расход</th>
      <th class="numeric">Прибыль</th>
      <th class="numeric">Маржа</th>
      <th>Резерв</th>
      <th></th>
    </tr></thead><tbody>`;

    rows.sort((a, b) => b.profit - a.profit);

    for (const r of rows) {
      const st = clientReserveStatus(r.translators);
      const initials = clientInitials(r.client_name);
      const profitColor = r.profit >= 0 ? '#16A34A' : '#DC2626';
      const profitSign = r.profit >= 0 ? '+' : '−';
      const marginBadge = r.margin_pct >= 0
        ? `<span class="badge badge-good">${r.margin_pct.toFixed(1)}%</span>`
        : `<span class="badge badge-warn">${r.margin_pct.toFixed(1)}%</span>`;

      html += `<tr style="cursor:pointer;" onclick="openClientDetail('${r.client_id}')">
        <td><div class="emp-cell">
          <div class="emp-avatar">${initials}</div>
          <div><div class="emp-name">${escapeHtml(r.client_name)}</div>
               ${r.contact_email ? `<div class="emp-email">${escapeHtml(r.contact_email)}</div>` : ''}</div>
        </div></td>
        <td style="text-align:right; font-family:'JetBrains Mono',monospace;">${r.translator_count}</td>
        <td style="text-align:right; font-family:'JetBrains Mono',monospace;">$${r.revenue.toFixed(2)}</td>
        <td style="text-align:right; font-family:'JetBrains Mono',monospace; color:#94A3B8;">$${r.cost.toFixed(2)}</td>
        <td style="text-align:right; font-family:'JetBrains Mono',monospace; font-weight:600; color:${profitColor};">${profitSign}$${Math.abs(r.profit).toFixed(2)}</td>
        <td style="text-align:right;">${marginBadge}</td>
        <td><span class="badge ${st.badgeCls}">${st.mode}</span></td>
        <td style="text-align:right;"><span style="color:#94A3B8; font-size:11px;">Открыть →</span></td>
      </tr>`;
    }
    html += '</tbody></table>';
    content.innerHTML = html;
  }

  function fillClientsMonthSelector() {
    const sel = document.getElementById('clients-month-select');
    if (sel.options.length > 0) {
      sel.value = `${clientsYear}-${clientsMonth}`;
      return;
    }
    const now = new Date();
    for (let i = 0; i < 12; i++) {
      const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
      const opt = document.createElement('option');
      opt.value = `${d.getFullYear()}-${d.getMonth()}`;
      opt.textContent = clientsMonthLabel(d.getFullYear(), d.getMonth());
      sel.appendChild(opt);
    }
    sel.value = `${clientsYear}-${clientsMonth}`;
    sel.onchange = () => {
      const [y, m] = sel.value.split('-').map(Number);
      clientsYear = y; clientsMonth = m;
      loadClients();
    };
  }

  // ── СТРАНИЦА: ДЕТАЛЬНАЯ КАРТОЧКА КЛИЕНТА ────────────────────────────
  async function openClientDetail(clientId) {
    clientDetailId = clientId;
    showPage('client-detail');
    fillClientDetailMonthSelector();
    await renderClientDetail();
  }

  function backToClients() {
    showPage('clients');
    loadClients();
  }

  function fillClientDetailMonthSelector() {
    const sel = document.getElementById('cd-month-select');
    sel.innerHTML = '';
    const now = new Date();
    for (let i = 0; i < 12; i++) {
      const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
      const opt = document.createElement('option');
      opt.value = `${d.getFullYear()}-${d.getMonth()}`;
      opt.textContent = clientsMonthLabel(d.getFullYear(), d.getMonth());
      sel.appendChild(opt);
    }
    sel.value = `${clientsYear}-${clientsMonth}`;
    sel.onchange = () => {
      const [y, m] = sel.value.split('-').map(Number);
      clientsYear = y; clientsMonth = m;
      renderClientDetail();
    };
  }

  async function renderClientDetail() {
    hideError('cd-error');
    let rows;
    try {
      rows = await fetchClientProfitability(clientsYear, clientsMonth);
    } catch (e) {
      showError('cd-error', 'Ошибка загрузки: ' + e.message);
      return;
    }

    const r = rows.find(x => x.client_id === clientDetailId);
    if (!r) {
      showError('cd-error', 'Клиент не найден в данных за этот месяц.');
      return;
    }

    // Шапка
    document.getElementById('cd-avatar').textContent = clientInitials(r.client_name);
    document.getElementById('cd-name').textContent = r.client_name;
    document.getElementById('cd-email').textContent = r.contact_email || '—';
    document.getElementById('cd-meta').innerHTML =
      `<span style="font-size:12px; color:#475569;">${r.translator_count} ${pluralize(r.translator_count,'переводчик','переводчика','переводчиков')}</span>`;

    // KPI клиента (агрегаты)
    document.getElementById('cd-kpi-revenue').textContent = '$' + r.revenue.toFixed(2);
    document.getElementById('cd-kpi-revenue-meta').textContent = 'по всем переводчикам';
    document.getElementById('cd-kpi-cost').textContent = '$' + r.cost.toFixed(2);
    document.getElementById('cd-kpi-cost-meta').textContent = 'зарплата';
    document.getElementById('cd-kpi-profit').textContent = '$' + r.profit.toFixed(2);
    document.getElementById('cd-kpi-margin').textContent = 'маржа ' + r.margin_pct.toFixed(1) + '%';

    // Сводный резерв клиента = сумма остатков переводчиков
    const totalRemaining = r.translators.reduce((s, t) => s + Math.max(0, t.hours_remaining), 0);
    const totalBought = r.translators.reduce((s, t) => s + t.hours_purchased, 0);
    document.getElementById('cd-kpi-reserve').textContent = totalRemaining.toFixed(0) + ' ч';
    document.getElementById('cd-kpi-reserve-meta').textContent = 'из ' + totalBought.toFixed(0) + ' ч';

    // Алерт — приоритет: убыток (красный) > резерв на исходе (жёлтый) > тишина.
    //
    //  • Убыток: неактивный переводчик с положительным остатком. Клиент оплатил
    //    часы, которые переводчик не отработал (уволился). Резерв вырабатывается
    //    другими, но для агентства это минус. Показываем сумму.
    //  • На исходе: активный переводчик, у которого резерв заканчивается
    //    (0 < остаток ≤ 20). Сигнал «первый месяц кончается, скоро постоплата».
    //  • Всё остальное (в т.ч. любой минус у активных) — норма, молчим.
    const alert = document.getElementById('cd-reserve-alert');
    const alertText = document.getElementById('cd-reserve-alert-text');

    const lossOnes = r.translators.filter(t => !t.is_active && t.hours_remaining > 0);
    const lowOnes  = r.translators.filter(t => t.is_active && t.hours_remaining > 0 && t.hours_remaining <= 20);

    if (lossOnes.length > 0) {
      const totalLoss = lossOnes.reduce((s, t) => s + t.unearned_loss, 0);
      const totalUnearnedH = lossOnes.reduce((s, t) => s + t.hours_remaining, 0);
      const names = lossOnes.map(t => t.translator_name).join(', ');
      // Красный стиль (убыток)
      alert.style.background = '#FCEBEB';
      alert.style.borderLeftColor = '#A32D2D';
      alert.style.color = '#A32D2D';
      alertText.textContent =
        `Убыток по неотработанному резерву: ${names} деактивирован(ы), ` +
        `осталось ${totalUnearnedH.toFixed(0)} ч оплаченного клиентом времени` +
        (totalLoss > 0 ? ` (~$${totalLoss.toFixed(2)}).` : '.');
      alert.classList.remove('hidden');
    } else if (lowOnes.length > 0) {
      // Жёлтый стиль (резерв на исходе)
      alert.style.background = '#FEF3C7';
      alert.style.borderLeftColor = '#B45309';
      alert.style.color = '#92400E';
      alertText.textContent =
        `У ${lowOnes.length} ${pluralize(lowOnes.length,'переводчика','переводчиков','переводчиков')} ` +
        `резерв заканчивается (≤ 20 ч) — скоро переход на постоплату.`;
      alert.classList.remove('hidden');
    } else {
      alert.classList.add('hidden');
    }

    // Сводный бейдж резерва клиента — по новой логике риска
    const st = clientReserveStatus(r.translators);
    const modeBadge = document.getElementById('cd-reserve-mode-badge');
    if (st.mode === '—') {
      modeBadge.classList.add('hidden');
    } else {
      modeBadge.classList.remove('hidden');
      modeBadge.textContent = st.mode;
      modeBadge.className = 'badge ' + st.badgeCls;
    }

    const usedTotal = Math.max(0, totalBought - totalRemaining);
    const usedPct = totalBought > 0 ? Math.max(0, Math.min(100, usedTotal / totalBought * 100)) : 0;
    document.getElementById('cd-reserve-used-label').textContent =
      `${usedTotal.toFixed(0)} ч из ${totalBought.toFixed(0)} ч`;
    const bar = document.getElementById('cd-reserve-bar');
    bar.style.width = usedPct + '%';
    bar.style.background = st.barColor;
    document.getElementById('cd-reserve-remaining-label').textContent = `${totalRemaining.toFixed(0)} ч осталось`;
    document.getElementById('cd-reserve-remaining-label').style.color = st.textColor;
    document.getElementById('cd-reserve-total-label').textContent = totalBought.toFixed(0) + ' ч';

    // Список переводчиков клиента с индивидуальными резервами.
    // Цвет и подпись отражают реальный смысл, а не просто знак остатка:
    //   • неактивный с остатком > 0 → красный (убыток по неотработанному резерву)
    //   • активный, 0 < остаток ≤ 20 → жёлтый (скоро постоплата)
    //   • активный в минусе → норма (постоплата), нейтрально
    const histEl = document.getElementById('cd-prepay-history');
    histEl.innerHTML = r.translators.map(t => {
      let color, note;
      if (!t.is_active && t.hours_remaining > 0) {
        color = '#DC2626';
        note = `убыток ~$${t.unearned_loss.toFixed(2)} · ${t.hours_remaining.toFixed(0)} ч не отработано`;
      } else if (t.is_active && t.hours_remaining > 0 && t.hours_remaining <= 20) {
        color = '#B45309';
        note = `осталось ${t.hours_remaining.toFixed(0)} ч — скоро постоплата`;
      } else if (t.hours_remaining > 0) {
        color = '#16A34A';
        note = `резерв ${t.hours_remaining.toFixed(0)} ч`;
      } else {
        color = '#94A3B8';
        note = 'постоплата';
      }
      const pct = t.hours_purchased > 0
        ? Math.max(0, Math.min(100, t.hours_remaining / t.hours_purchased * 100)) : 0;
      const inactiveBadge = !t.is_active
        ? ' <span class="badge badge-warn" style="font-size:10px; padding:1px 6px;">деактивирован</span>' : '';
      return `<div style="padding:10px 0; border-bottom:1px solid #F1F5F9;">
        <div style="display:flex; align-items:center; justify-content:space-between; margin-bottom:6px;">
          <span style="font-size:13px; color:#0F1B3D; font-weight:500;">${escapeHtml(t.translator_name)}${inactiveBadge}</span>
          <span style="font-family:'JetBrains Mono',monospace; font-size:12px; color:${color};">${t.hours_remaining.toFixed(0)} ч</span>
        </div>
        <div style="width:100%; height:6px; background:#F1F5F9; border:1px solid #E5E7EB; border-radius:100px; overflow:hidden;">
          <div style="height:100%; width:${pct}%; background:${color}; border-radius:100px;"></div>
        </div>
        <div style="display:flex; justify-content:space-between; align-items:center; margin-top:3px;">
          <span style="font-size:11px; color:${color === '#94A3B8' ? '#94A3B8' : color};">${note}</span>
          <button class="btn btn-ghost btn-sm" style="padding:2px 8px; font-size:11px;" onclick="openAddPrepaymentModal('${t.user_id}')">+ Предоплата</button>
        </div>
      </div>`;
    }).join('');

    // Тарифы по парам (клиентские)
    const ratesEl = document.getElementById('cd-rates-by-pair');
    if (!r.rates_by_pair || r.rates_by_pair.length === 0) {
      ratesEl.innerHTML = '<div style="color:#94A3B8; font-size:13px; padding:8px 0;">Тарифы клиента не заданы. Задайте их через «Редактировать».</div>';
    } else {
      ratesEl.innerHTML = r.rates_by_pair.map(p => {
        const cRate = Number(p.client_rate) || 0;
        return `<div style="display:flex; align-items:center; justify-content:space-between; padding:8px 0; border-bottom:1px solid #F1F5F9;">
          <span style="font-family:'JetBrains Mono',monospace; font-size:12px; font-weight:600;">${escapeHtml(p.pair_code)}</span>
          <span style="color:#1E40AF; font-family:'JetBrains Mono',monospace; font-size:13px; font-weight:600;">$${cRate.toFixed(2)}/ч</span>
        </div>`;
      }).join('');
    }

    // Разбивка по переводчикам за месяц
    document.getElementById('cd-breakdown-month').textContent = clientsMonthLabel(clientsYear, clientsMonth);
    const bdEl = document.getElementById('cd-breakdown');
    if (r.translators.length === 0) {
      bdEl.innerHTML = '<div class="empty-state"><div class="empty-state-text">К клиенту не привязан ни один переводчик.</div></div>';
    } else {
      let bdHtml = `<table><thead><tr>
        <th>Переводчик</th>
        <th class="numeric">Часы (мес.)</th>
        <th class="numeric">Доход</th>
        <th class="numeric">Расход</th>
        <th class="numeric">Прибыль</th>
        <th class="numeric">Маржа</th>
      </tr></thead><tbody>`;
      const sorted = [...r.translators].sort((a, b) => b.profit - a.profit);
      for (const t of sorted) {
        const profitColor = t.profit >= 0 ? '#16A34A' : '#DC2626';
        bdHtml += `<tr>
          <td>${escapeHtml(t.translator_name)}</td>
          <td style="text-align:right; font-family:'JetBrains Mono',monospace;">${t.hours_period.toFixed(1)} ч</td>
          <td style="text-align:right; font-family:'JetBrains Mono',monospace;">$${t.revenue.toFixed(2)}</td>
          <td style="text-align:right; font-family:'JetBrains Mono',monospace; color:#94A3B8;">$${t.cost.toFixed(2)}</td>
          <td style="text-align:right; font-family:'JetBrains Mono',monospace; font-weight:600; color:${profitColor};">$${t.profit.toFixed(2)}</td>
          <td style="text-align:right;"><span class="badge ${t.margin_pct>=0?'badge-good':'badge-warn'}">${t.margin_pct.toFixed(1)}%</span></td>
        </tr>`;
      }
      bdHtml += '</tbody></table>';
      bdEl.innerHTML = bdHtml;
    }

    // ── Тренинги ────────────────────────────────────────────────────────
    renderTrainingBlock(r);
  }

  // Статусы тренинга → человекочитаемо + цвет
  function trainingStatusMeta(status) {
    switch (status) {
      case 'payable': return { label: 'Готов к выплате', cls: 'badge-warn', color: '#B45309' };
      case 'paid':    return { label: 'Выплачено',       cls: 'badge-good', color: '#16A34A' };
      case 'waived':  return { label: 'Не выплачивается', cls: 'badge-neutral', color: '#94A3B8' };
      default:        return { label: 'Ждёт 3 мес',      cls: 'badge-neutral', color: '#475569' };
    }
  }

  function renderTrainingBlock(r) {
    const summaryEl = document.getElementById('cd-training-summary');
    const bodyEl = document.getElementById('cd-training');
    const trainAlert = document.getElementById('cd-training-alert');

    // Сводка по клиенту
    const ts = r.training_summary;
    if (ts) {
      const parts = [];
      parts.push(`доход $${Number(ts.revenue).toFixed(2)}`);
      if (Number(ts.cost_paid) > 0) parts.push(`выплачено $${Number(ts.cost_paid).toFixed(2)}`);
      if (Number(ts.payable_count) > 0) parts.push(`<span style="color:#B45309; font-weight:600;">${ts.payable_count} к выплате</span>`);
      summaryEl.innerHTML = parts.join(' · ');

      // Алерт «пора выплатить» — если есть созревшие тренинги
      if (Number(ts.payable_count) > 0) {
        document.getElementById('cd-training-alert-text').textContent =
          `${ts.payable_count} ${pluralize(ts.payable_count,'переводчик отработал','переводчика отработали','переводчиков отработали')} 3 месяца — пора выплатить тренинг.`;
        trainAlert.classList.remove('hidden');
      } else {
        trainAlert.classList.add('hidden');
      }
    } else {
      summaryEl.textContent = 'нет данных';
      trainAlert.classList.add('hidden');
    }

    // Переводчики, у которых есть тренинг
    const withTraining = r.translators.filter(t => t.training);
    if (withTraining.length === 0) {
      bodyEl.innerHTML = '<div class="empty-state"><div class="empty-state-text">Тренинги не заведены. Они создаются автоматически при привязке переводчика к клиенту.</div></div>';
      return;
    }

    let html = `<table><thead><tr>
      <th>Переводчик</th>
      <th class="numeric">Часы</th>
      <th class="numeric">Доход</th>
      <th class="numeric">Расход</th>
      <th class="numeric">Маржа</th>
      <th>Статус</th>
      <th></th>
    </tr></thead><tbody>`;

    for (const t of withTraining) {
      const tr = t.training;
      const meta = trainingStatusMeta(tr.status);
      const revenue = Number(tr.revenue) || 0;
      const cost = Number(tr.cost) || 0;
      const margin = Number(tr.margin) || 0;

      // Кнопки действий по статусу
      let actions = '';
      if (tr.status === 'payable') {
        actions = `<button class="btn btn-sm" style="padding:3px 10px; font-size:12px;" onclick="markTrainingPaid('${t.user_id}')">Выплатить</button>
                   <button class="btn btn-ghost btn-sm" style="padding:3px 10px; font-size:12px;" onclick="markTrainingWaived('${t.user_id}')">Не выплачивать</button>`;
      } else if (tr.status === 'paid') {
        actions = `<span style="font-size:11px; color:#94A3B8;">${tr.paid_at ? formatDateRu(tr.paid_at) : ''}</span>`;
      } else if (tr.status === 'waived') {
        actions = `<button class="btn btn-ghost btn-sm" style="padding:3px 10px; font-size:12px;" onclick="markTrainingPaid('${t.user_id}')">Всё же выплатить</button>`;
      } else {
        // pending — показываем сколько ждать
        actions = `<span style="font-size:11px; color:#94A3B8;">${tr.start_date ? 'с ' + formatDateRu(tr.start_date) : 'ждёт первого дня'}</span>`;
      }

      html += `<tr>
        <td>${escapeHtml(t.translator_name)}${!t.is_active ? ' <span class="badge badge-warn" style="font-size:10px; padding:1px 6px;">деактивирован</span>' : ''}</td>
        <td style="text-align:right; font-family:'JetBrains Mono',monospace;">${Number(tr.hours).toFixed(0)} ч</td>
        <td style="text-align:right; font-family:'JetBrains Mono',monospace;">$${revenue.toFixed(2)}</td>
        <td style="text-align:right; font-family:'JetBrains Mono',monospace; color:#94A3B8;">$${cost.toFixed(2)}</td>
        <td style="text-align:right; font-family:'JetBrains Mono',monospace; font-weight:600; color:#16A34A;">$${margin.toFixed(2)}</td>
        <td><span class="badge ${meta.cls}">${meta.label}</span></td>
        <td style="text-align:right; white-space:nowrap;">${actions}</td>
      </tr>`;
    }
    html += '</tbody></table>';
    bodyEl.innerHTML = html;
  }

  // Отметить тренинг выплаченным (расход фиксируется, дата = сегодня)
  async function markTrainingPaid(userId) {
    if (!confirm('Отметить тренинг как выплаченный? Это зафиксирует расход по тренингу.')) return;
    try {
      const { error } = await sb
        .from('client_trainings')
        .update({ payout_status: 'paid', paid_at: new Date().toISOString().split('T')[0] })
        .eq('user_id', userId);
      if (error) throw new Error(error.message);
      invalidateCache('clientProfit:');
      await renderClientDetail();
    } catch (e) {
      showError('cd-error', 'Не удалось отметить выплату: ' + e.message);
    }
  }

  // Отметить тренинг как не подлежащий выплате (ушёл раньше срока и т.п.)
  async function markTrainingWaived(userId) {
    if (!confirm('Отметить тренинг как «не выплачивается»? Обычно — если переводчик ушёл раньше 3 месяцев.')) return;
    try {
      const { error } = await sb
        .from('client_trainings')
        .update({ payout_status: 'waived', paid_at: null })
        .eq('user_id', userId);
      if (error) throw new Error(error.message);
      invalidateCache('clientProfit:');
      await renderClientDetail();
    } catch (e) {
      showError('cd-error', 'Не удалось изменить статус: ' + e.message);
    }
  }

  // ── МОДАЛКА: СОЗДАНИЕ КЛИЕНТА (тарифы на клиента по парам) ───────────
  function openAddClientModal() {
    hideError('add-client-error');
    document.getElementById('client-name').value = '';
    document.getElementById('client-email').value = '';
    document.getElementById('client-note').value = '';
    document.getElementById('client-rates-wrap').classList.remove('hidden');
    document.getElementById('client-rates-list').innerHTML = '';
    clientRatePairRows = 0;

    // Одна пустая строка тарифа по умолчанию
    addClientRateRow();

    document.getElementById('add-client-modal').classList.add('open');
  }

  // Строка «языковая пара → тариф клиента» в форме создания
  function addClientRateRow(existing = null) {
    clientRatePairRows++;
    const rowId = 'crr-' + clientRatePairRows;
    const options = languagePairs.map(p => {
      const sel = existing && existing.language_pair_id === p.id ? 'selected' : '';
      return `<option value="${p.id}" ${sel}>${p.code} — ${p.display_name}</option>`;
    }).join('');

    const div = document.createElement('div');
    div.className = 'pair-editor';
    div.id = rowId;
    div.style.cssText = 'display:flex; align-items:center; gap:10px; margin-bottom:8px;';
    div.innerHTML = `
      <select class="select client-rate-pair" style="flex:1;">${options}</select>
      <div style="display:flex; align-items:center; gap:4px;">
        <span style="color:#475569;">$</span>
        <input type="number" class="input client-rate-value" style="width:90px;"
               step="0.50" min="0.50" max="1000"
               value="${existing ? Number(existing.rate_per_hour).toFixed(2) : ''}"
               placeholder="8.00">
        <span style="color:#475569; font-size:12px;">/ч</span>
      </div>
      <button class="pair-remove" onclick="removeClientRateRow('${rowId}')" title="Удалить">×</button>
    `;
    document.getElementById('client-rates-list').appendChild(div);
  }

  function removeClientRateRow(rowId) {
    const el = document.getElementById(rowId);
    if (el) el.remove();
  }

  function closeAddClientModal() {
    document.getElementById('add-client-modal').classList.remove('open');
  }

  async function saveClient() {
    hideError('add-client-error');
    const btn = document.getElementById('btn-save-client');

    const name = document.getElementById('client-name').value.trim();
    const email = document.getElementById('client-email').value.trim();
    const note = document.getElementById('client-note').value.trim();

    if (!name) {
      showError('add-client-error', 'Укажите название клиента.');
      return;
    }

    // Собираем тарифы, проверяем на дубли пар
    const rowEls = document.querySelectorAll('#client-rates-list .pair-editor');
    const rates = [];
    const seenPairs = new Set();
    for (const row of rowEls) {
      const pairId = row.querySelector('.client-rate-pair').value;
      const val = parseFloat(row.querySelector('.client-rate-value').value);
      if (!val || val <= 0) {
        showError('add-client-error', 'Укажите тариф для всех пар (больше 0).');
        return;
      }
      if (seenPairs.has(pairId)) {
        showError('add-client-error', 'Языковая пара указана дважды. Оставьте одну строку на пару.');
        return;
      }
      seenPairs.add(pairId);
      rates.push({ language_pair_id: pairId, rate_per_hour: val });
    }
    if (rates.length === 0) {
      showError('add-client-error', 'Добавьте хотя бы одну пару с тарифом.');
      return;
    }

    btn.disabled = true; btn.textContent = 'Создаём…';
    try {
      const { data: client, error: clientErr } = await sb
        .from('clients')
        .insert({ name, contact_email: email || null, note: note || null, created_by: currentUser.id })
        .select('id')
        .single();
      if (clientErr) throw new Error('Клиент: ' + clientErr.message);

      const rateRows = rates.map(r => ({
        client_id: client.id,
        language_pair_id: r.language_pair_id,
        rate_per_hour: r.rate_per_hour,
      }));
      const { error: ratesErr } = await sb.from('client_rates').insert(rateRows);
      if (ratesErr) throw new Error('Тарифы: ' + ratesErr.message);

      invalidateCache('clientProfit:');
      closeAddClientModal();
      await loadClients();
    } catch (e) {
      showError('add-client-error', e.message);
    } finally {
      btn.disabled = false; btn.textContent = 'Создать клиента';
    }
  }

  // ── МОДАЛКА: РЕДАКТИРОВАНИЕ КЛИЕНТА ─────────────────────────────────
  async function openEditClientModal() {
    if (!clientDetailId) return;
    hideError('edit-client-error');

    const { data: client, error } = await sb
      .from('clients')
      .select('id, name, contact_email, note')
      .eq('id', clientDetailId)
      .single();
    if (error || !client) {
      showError('cd-error', 'Не удалось загрузить клиента.');
      return;
    }

    clientEditId = client.id;
    document.getElementById('edit-client-name').value = client.name || '';
    document.getElementById('edit-client-email').value = client.contact_email || '';
    document.getElementById('edit-client-note').value = client.note || '';

    // Текущие тарифы клиента
    const { data: currentRates } = await sb
      .from('client_rates')
      .select('language_pair_id, rate_per_hour')
      .eq('client_id', clientEditId);

    const rateByPair = {};
    for (const cr of (currentRates || [])) rateByPair[cr.language_pair_id] = Number(cr.rate_per_hour);

    // Рисуем строку на каждую активную языковую пару, подставляя существующий тариф
    const list = document.getElementById('edit-client-rates-list');
    list.innerHTML = languagePairs.map(p => {
      const existing = rateByPair[p.id];
      return `<div style="display:flex; align-items:center; gap:10px; padding:8px 0; border-bottom:1px solid #F1F5F9;">
        <span class="badge badge-neutral" style="min-width:96px;">${p.code}</span>
        <span style="flex:1; color:#94A3B8; font-size:12px;">${escapeHtml(p.display_name)}</span>
        <div style="display:flex; align-items:center; gap:4px;">
          <span style="color:#475569;">$</span>
          <input type="number" class="input edit-client-rate-field" data-pair-id="${p.id}"
                 style="width:90px;" step="0.50" min="0" max="1000"
                 value="${existing != null ? existing.toFixed(2) : ''}"
                 placeholder="—">
          <span style="color:#475569; font-size:12px;">/ч</span>
        </div>
      </div>`;
    }).join('');
    document.getElementById('edit-client-rates-hint').textContent =
      'Пустое поле = тариф для этой пары не задан. Заполненное — создаётся или обновляется.';

    document.getElementById('edit-client-modal').classList.add('open');
  }

  function closeEditClientModal() {
    document.getElementById('edit-client-modal').classList.remove('open');
  }

  async function saveEditClient() {
    hideError('edit-client-error');
    const btn = document.getElementById('btn-save-edit-client');

    const name = document.getElementById('edit-client-name').value.trim();
    const email = document.getElementById('edit-client-email').value.trim();
    const note = document.getElementById('edit-client-note').value.trim();

    if (!name) {
      showError('edit-client-error', 'Укажите название клиента.');
      return;
    }
    if (!clientEditId) {
      showError('edit-client-error', 'Клиент не определён.');
      return;
    }

    // Тарифы: заполненные → upsert; пустые → удалить (если были)
    const fields = document.querySelectorAll('#edit-client-rates-list .edit-client-rate-field');
    const toUpsert = [];
    const toDelete = [];
    for (const f of fields) {
      const val = parseFloat(f.value);
      if (val && val > 0) {
        toUpsert.push({ client_id: clientEditId, language_pair_id: f.dataset.pairId, rate_per_hour: val });
      } else {
        toDelete.push(f.dataset.pairId);
      }
    }

    btn.disabled = true; btn.textContent = 'Сохраняем…';
    try {
      const { error: clientErr } = await sb
        .from('clients')
        .update({ name, contact_email: email || null, note: note || null })
        .eq('id', clientEditId);
      if (clientErr) throw new Error('Клиент: ' + clientErr.message);

      if (toUpsert.length > 0) {
        const { error: upErr } = await sb
          .from('client_rates')
          .upsert(toUpsert, { onConflict: 'client_id,language_pair_id' });
        if (upErr) throw new Error('Тарифы: ' + upErr.message);
      }
      if (toDelete.length > 0) {
        const { error: delErr } = await sb
          .from('client_rates')
          .delete()
          .eq('client_id', clientEditId)
          .in('language_pair_id', toDelete);
        if (delErr) throw new Error('Удаление тарифов: ' + delErr.message);
      }

      invalidateCache('clientProfit:');
      closeEditClientModal();
      await renderClientDetail();
    } catch (e) {
      showError('edit-client-error', e.message);
    } finally {
      btn.disabled = false; btn.textContent = 'Сохранить';
    }
  }

  // ── МОДАЛКА: ПРИВЯЗКА ПЕРЕВОДЧИКА К КЛИЕНТУ ─────────────────────────
  async function openAttachTranslatorModal() {
    if (!clientDetailId) return;
    hideError('attach-translator-error');
    document.getElementById('attach-translator-rates-warn').classList.add('hidden');

    // Название клиента в подзаголовок
    const nameEl = document.getElementById('cd-name');
    document.getElementById('attach-translator-subtitle').textContent =
      'Клиент: ' + (nameEl ? nameEl.textContent : '');

    // Загружаем свободных переводчиков (client_id IS NULL) + их пары
    const sel = document.getElementById('attach-translator-select');
    sel.innerHTML = '<option value="">— выберите переводчика —</option>';

    const { data: free, error } = await sb
      .from('users')
      .select('id, name, translator_pairs ( language_pair_id, language_pairs (code) )')
      .eq('role', 'translator')
      .eq('is_active', true)
      .is('client_id', null)
      .order('name');

    if (error) {
      showError('attach-translator-error', 'Ошибка загрузки: ' + error.message);
      return;
    }

    // Кеш пар переводчика — для проверки недостающих тарифов клиента
    window._attachFreeCache = {};
    for (const t of (free || [])) {
      window._attachFreeCache[t.id] = t;
      const opt = document.createElement('option');
      opt.value = t.id;
      opt.textContent = t.name;
      sel.appendChild(opt);
    }

    if ((free || []).length === 0) {
      sel.innerHTML = '<option value="">— нет свободных переводчиков —</option>';
    }

    // При выборе — проверяем, есть ли у клиента тарифы под пары переводчика
    sel.onchange = () => checkAttachTranslatorRates(sel.value);

    document.getElementById('attach-translator-modal').classList.add('open');
  }

  // Проверяет, для всех ли языковых пар переводчика у клиента задан тариф.
  // Если чего-то не хватает — показывает предупреждение (не блокирует привязку).
  async function checkAttachTranslatorRates(userId) {
    const warn = document.getElementById('attach-translator-rates-warn');
    if (!userId || !window._attachFreeCache || !window._attachFreeCache[userId]) {
      warn.classList.add('hidden');
      return;
    }
    const t = window._attachFreeCache[userId];
    const pairs = (t.translator_pairs || []);
    if (pairs.length === 0) {
      warn.classList.add('hidden');
      return;
    }

    // Тарифы клиента
    const { data: rates } = await sb
      .from('client_rates')
      .select('language_pair_id')
      .eq('client_id', clientDetailId);
    const haveRates = new Set((rates || []).map(r => r.language_pair_id));

    const missing = pairs
      .filter(p => !haveRates.has(p.language_pair_id))
      .map(p => p.language_pairs.code);

    if (missing.length > 0) {
      document.getElementById('attach-translator-rates-warn-text').textContent =
        `У клиента не задан тариф для пар: ${missing.join(', ')}. Доход по этим парам будет $0, пока не зададите тариф через «Редактировать».`;
      warn.classList.remove('hidden');
    } else {
      warn.classList.add('hidden');
    }
  }

  function closeAttachTranslatorModal() {
    document.getElementById('attach-translator-modal').classList.remove('open');
  }

  async function attachTranslator() {
    hideError('attach-translator-error');
    const btn = document.getElementById('btn-attach-translator');
    const userId = document.getElementById('attach-translator-select').value;

    if (!userId) {
      showError('attach-translator-error', 'Выберите переводчика.');
      return;
    }

    btn.disabled = true; btn.textContent = 'Привязываем…';
    try {
      const { error } = await sb
        .from('users')
        .update({ client_id: clientDetailId })
        .eq('id', userId);
      if (error) throw new Error(error.message);

      // Автозаведение тренинга (70 ч) со снимком ставок на момент выхода.
      // Тариф клиента — по основной паре переводчика; ставка переводчика — оттуда же.
      // Если записи ещё нет (UNIQUE user_id) — создаём; если есть — не трогаем.
      await createTrainingForTranslator(userId, clientDetailId);

      invalidateCache('clientProfit:');
      closeAttachTranslatorModal();
      await renderClientDetail();
    } catch (e) {
      showError('attach-translator-error', e.message);
    } finally {
      btn.disabled = false; btn.textContent = 'Привязать';
    }
  }

  // Создаёт запись тренинга для переводчика при выходе на клиента.
  // Снимок ставок: тариф клиента и ставка переводчика по основной паре.
  // start_date не заполняем — RPC вычислит по первому рабочему дню.
  async function createTrainingForTranslator(userId, clientId) {
    // Уже есть тренинг? (UNIQUE user_id) — не дублируем
    const { data: existing } = await sb
      .from('client_trainings')
      .select('id')
      .eq('user_id', userId)
      .maybeSingle();
    if (existing) return;

    // Основная пара переводчика + его ставка
    const { data: pairs } = await sb
      .from('translator_pairs')
      .select('language_pair_id, rate_per_hour, is_primary')
      .eq('user_id', userId);

    let clientRate = 0, translatorRate = 0;
    if (pairs && pairs.length > 0) {
      const primary = pairs.find(p => p.is_primary) || pairs[0];
      translatorRate = Number(primary.rate_per_hour) || 0;
      // Тариф клиента по этой паре
      const { data: cr } = await sb
        .from('client_rates')
        .select('rate_per_hour')
        .eq('client_id', clientId)
        .eq('language_pair_id', primary.language_pair_id)
        .maybeSingle();
      if (cr) clientRate = Number(cr.rate_per_hour) || 0;
    }

    await sb.from('client_trainings').insert({
      client_id: clientId,
      user_id: userId,
      training_hours: 70,
      client_rate: clientRate,
      translator_rate: translatorRate,
      payout_status: 'pending',
      created_by: currentUser.id,
    });
  }

  // ── МОДАЛКА: ВНЕСЕНИЕ ПРЕДОПЛАТЫ (под конкретного переводчика) ───────
  let prepayUserId = null; // переводчик, которому вносим предоплату

  async function openAddPrepaymentModal(userId) {
    prepayUserId = userId;
    hideError('add-prepayment-error');

    let rows;
    try {
      rows = await fetchClientProfitability(clientsYear, clientsMonth);
    } catch (e) {
      showError('cd-error', 'Ошибка: ' + e.message);
      return;
    }
    const client = rows.find(x => x.client_id === clientDetailId);
    const t = client ? client.translators.find(x => x.user_id === userId) : null;
    if (!t) {
      showError('cd-error', 'Переводчик не найден.');
      return;
    }

    document.getElementById('prepay-modal-subtitle').textContent =
      client.client_name + ' · ' + t.translator_name;
    document.getElementById('prepay-current-reserve').textContent = t.hours_remaining.toFixed(0) + ' ч';
    document.getElementById('prepay-reserve-meta').textContent =
      `Куплено ${t.hours_purchased.toFixed(0)} ч · отработано ${t.hours_worked_total.toFixed(0)} ч`;

    document.getElementById('prepay-hours').value = '';
    document.getElementById('prepay-rate').value = '';
    document.getElementById('prepay-note').value = '';
    document.getElementById('prepay-paid-at').value = new Date().toISOString().split('T')[0];
    document.getElementById('prepay-total-amount').textContent = '$0.00';

    const recalc = () => {
      const h = parseFloat(document.getElementById('prepay-hours').value) || 0;
      const rate = parseFloat(document.getElementById('prepay-rate').value) || 0;
      document.getElementById('prepay-total-amount').textContent = '$' + (h * rate).toFixed(2);
    };
    document.getElementById('prepay-hours').oninput = recalc;
    document.getElementById('prepay-rate').oninput = recalc;

    document.getElementById('add-prepayment-modal').classList.add('open');
  }

  function closeAddPrepaymentModal() {
    document.getElementById('add-prepayment-modal').classList.remove('open');
  }

  async function savePrepayment() {
    hideError('add-prepayment-error');
    const btn = document.getElementById('btn-save-prepayment');

    const hours = parseFloat(document.getElementById('prepay-hours').value);
    const rate = parseFloat(document.getElementById('prepay-rate').value);
    const paidAt = document.getElementById('prepay-paid-at').value;
    const note = document.getElementById('prepay-note').value.trim();

    if (!hours || hours <= 0) { showError('add-prepayment-error', 'Укажите количество часов.'); return; }
    if (!rate || rate <= 0) { showError('add-prepayment-error', 'Укажите тариф.'); return; }
    if (!paidAt) { showError('add-prepayment-error', 'Укажите дату оплаты.'); return; }
    if (!prepayUserId) { showError('add-prepayment-error', 'Переводчик не определён.'); return; }

    btn.disabled = true; btn.textContent = 'Вносим…';
    try {
      const { error } = await sb.from('client_prepayments').insert({
        client_id: clientDetailId,
        user_id: prepayUserId,
        hours_purchased: hours,
        rate_per_hour: rate,
        paid_at: paidAt,
        note: note || null,
        created_by: currentUser.id,
      });
      if (error) throw new Error(error.message);

      invalidateCache('clientProfit:');
      closeAddPrepaymentModal();
      await renderClientDetail();
    } catch (e) {
      showError('add-prepayment-error', e.message);
    } finally {
      btn.disabled = false; btn.textContent = 'Внести предоплату';
    }
  }

  // ═══════════════════════════════════════════════════════════════════════
  // Биржа обмена постоянными графиками (schedule swap).
  // Бэкенд: RPC swap_* (миграции schedule_swap_01..04). Постоянный график
  // живёт в users.default_shift_*; здесь только выставление, заявки и своп.
  // ═══════════════════════════════════════════════════════════════════════

  const SWAP_STATUS_META = {
    open:             { label: '● Открыта',             cls: 'pending'  },
    manager_approved: { label: '● Одобрена менеджером',  cls: 'pending'  },
    pending_icall:    { label: '● Ждёт I-Call',          cls: 'pending'  },
    approved:         { label: '● Обмен выполнен',       cls: 'approved' },
    rejected:         { label: '● Отклонена',            cls: 'rejected' },
    cancelled:        { label: '● Отменена',             cls: 'rejected' },
  };
  const SWAP_ACTIVE = ['open', 'manager_approved', 'pending_icall'];

  function fmtWin(s, e, m) {
    return `${String(s).substring(0,5)}–${String(e).substring(0,5)} (${formatHoursMinutes(m)})`;
  }

  // Понятные сообщения для кодов ошибок из RPC
  function swapErr(error) {
    const msg = (error && error.message) || '';
    const map = {
      pair_mismatch: 'Языковые пары не совпадают — обмен возможен только между одинаковыми парами.',
      requester_frozen: 'Ваш обмен сейчас заморожен.',
      target_frozen: 'У этого графика активна заморозка.',
      target_in_progress: 'По этому графику уже идёт обмен.',
      requester_has_active: 'У вас уже есть активная заявка на обмен.',
      listing_not_found: 'Этот график уже недоступен.',
      cannot_swap_self: 'Нельзя обменяться с самим собой.',
      schedule_frozen: 'Действие недоступно во время заморозки.',
      target_not_agreed: 'Сначала зафиксируйте согласие второй стороны.',
      cannot_cancel: 'Заявку уже нельзя отменить.',
      manager_only: 'Действие доступно только менеджеру.',
      not_found_or_not_open: 'Заявку нельзя изменить в текущем статусе.',
      not_found_or_wrong_state: 'Заявку нельзя изменить в текущем статусе.',
      wrong_state: 'Неверный статус заявки для этого действия.',
      effective_from_required: 'Укажите дату вступления в силу.',
      frozen_use_override: 'Один из участников заморожен — включите «обойти заморозку».',
      requester_invalid: 'Доступно только для переводчиков.',
      target_inactive: 'Переводчик неактивен.',
    };
    for (const k in map) if (msg.includes(k)) return map[k];
    return 'Ошибка: ' + msg;
  }

  // ──── Переводчик: биржа ────────────────────────────────────────────
  async function loadSwapMarket() {
    hideError('swap-error');
    await renderMyListingBlock();
    await renderIncomingOffers();
    await renderMarketAndOutgoing();
  }

  async function renderMyListingBlock() {
    const box = document.getElementById('swap-my-listing');
    box.innerHTML = '<div class="loading-state">Загрузка…</div>';

    const [profRes, listRes] = await Promise.all([
      sb.from('users').select('default_shift_start, default_shift_end, default_shift_minutes')
        .eq('id', currentUser.id).single(),
      sb.from('schedule_listings').select('is_listed, swap_frozen_until, note')
        .eq('user_id', currentUser.id).maybeSingle(),
    ]);
    const p = profRes.data || {};
    const listing = listRes.data || { is_listed: false, swap_frozen_until: null, note: '' };
    const se = `${(p.default_shift_start || '').substring(0,5)}–${(p.default_shift_end || '').substring(0,5)}`;
    const mins = p.default_shift_minutes || 0;

    const todayISO = new Date().toISOString().split('T')[0];
    const frozen = !!listing.swap_frozen_until && listing.swap_frozen_until > todayISO;

    let control;
    if (frozen) {
      control = `<div class="swap-freeze-note">🔒 Обмен заморожен до ${formatDateRu(listing.swap_frozen_until)} — после недавнего обмена нужно отработать новый график.</div>`;
    } else if (listing.is_listed) {
      control = `
        <div class="swap-listed-note">✓ Ваш график на бирже — коллеги видят это окно без вашего имени и могут предложить обмен.</div>
        <button class="btn btn-ghost btn-sm" onclick="toggleMyListing(false)">Снять с биржи</button>`;
    } else {
      control = `
        <div class="field" style="margin-bottom:10px;">
          <label class="field-label">Комментарий для менеджера (не виден другим переводчикам)</label>
          <input type="text" id="swap-note" class="input" maxlength="200" placeholder="Например: хочу более раннее окончание" value="${escapeHtml(listing.note || '')}">
        </div>
        <button class="btn btn-sm" onclick="toggleMyListing(true)">Выставить мой график на биржу</button>`;
    }

    box.innerHTML = `
      <div class="shift-section ${listing.is_listed && !frozen ? 'assigned' : ''}">
        <div class="shift-section-header"><span class="shift-section-label">Действует сейчас</span></div>
        <div class="shift-section-content">
          <span class="shift-section-time">${se}</span>
          <span class="shift-section-duration">· ${formatHoursMinutes(mins)} плана</span>
        </div>
      </div>
      <div style="margin-top:12px;">${control}</div>`;
  }

  async function toggleMyListing(list) {
    hideError('swap-error');
    const note = list ? (document.getElementById('swap-note')?.value.trim() || null) : null;
    const { error } = await sb.rpc('swap_my_listing_set', { p_is_listed: list, p_note: note });
    if (error) { showError('swap-error', swapErr(error)); return; }
    await loadSwapMarket();
  }

  async function renderIncomingOffers() {
    const box = document.getElementById('swap-incoming');
    const { data, error } = await sb.rpc('swap_my_incoming');
    if (error || !data || data.length === 0) { box.innerHTML = ''; return; }
    box.innerHTML = `
      <div class="section">
        <div style="font-size: 14px; color: #0F1B3D; font-weight: 600; margin-bottom: 4px;">Мне предлагают обмен</div>
        <div style="font-size:12px;color:#94A3B8;margin-bottom:12px;">Заявки на ваш выставленный график. Решение принимает менеджер — он свяжется с вами. Личность отправителя скрыта.</div>
        ${data.map(o => {
          const meta = SWAP_STATUS_META[o.status] || { label:o.status, cls:'' };
          const mine = fmtWin(o.my_start, o.my_end, o.my_minutes);
          const prop = fmtWin(o.proposed_start, o.proposed_end, o.proposed_minutes);
          return `
            <div class="req-card ${meta.cls}">
              <div class="req-info" style="grid-column:1 / -1;">
                <div style="display:flex;justify-content:space-between;align-items:center;gap:12px;flex-wrap:wrap;">
                  <div class="req-change" style="margin:0;">
                    <span class="req-change-from">Ваше окно: ${mine}</span>
                    <span class="req-change-arrow">→</span>
                    <span class="req-change-to">Предлагают: ${prop}</span>
                  </div>
                  <span class="req-status-badge ${meta.cls}">${meta.label}</span>
                </div>
              </div>
            </div>`;
        }).join('')}
      </div>`;
  }

  async function renderMarketAndOutgoing() {
    const marketBox = document.getElementById('swap-market-list');
    const outBox = document.getElementById('swap-outgoing');

    const [mktRes, outRes] = await Promise.all([
      sb.rpc('swap_market_list'),
      sb.rpc('swap_my_outgoing'),
    ]);
    if (mktRes.error) showError('swap-error', swapErr(mktRes.error));
    const market = mktRes.data || [];
    const outgoing = outRes.data || [];
    const activeOut = outgoing.find(o => SWAP_ACTIVE.includes(o.status));

    if (market.length === 0) {
      marketBox.innerHTML = `<div class="empty-state"><div class="empty-state-text">Сейчас нет выставленных графиков вашей языковой пары.</div></div>`;
    } else {
      marketBox.innerHTML = market.map(c => {
        const se = `${c.window_start.substring(0,5)}–${c.window_end.substring(0,5)}`;
        const btn = activeOut
          ? `<button class="btn btn-sm" disabled title="У вас уже есть активная заявка">Недоступно</button>`
          : `<button class="btn btn-sm" onclick="openApplyModal('${c.listing_id}','${se} (${formatHoursMinutes(c.minutes)})')">Предложить обмен</button>`;
        return `
          <div class="req-card">
            <div class="req-info" style="grid-column:1 / -1;">
              <div style="display:flex;justify-content:space-between;align-items:center;gap:12px;flex-wrap:wrap;">
                <div>
                  <div class="req-translator-name" style="font-family:'JetBrains Mono',monospace;">${se} · ${formatHoursMinutes(c.minutes)}</div>
                  <div class="req-date">${escapeHtml(c.pair_code || '')}</div>
                </div>
                ${btn}
              </div>
            </div>
          </div>`;
      }).join('');
    }

    if (outgoing.length === 0) { outBox.innerHTML = ''; return; }
    outBox.innerHTML = `
      <div class="swap-subhead">Мои заявки на обмен</div>
      ${outgoing.map(o => {
        const meta = SWAP_STATUS_META[o.status] || { label:o.status, cls:'' };
        const give = fmtWin(o.give_start, o.give_end, o.give_minutes);
        const get = fmtWin(o.get_start, o.get_end, o.get_minutes);
        const cancel = SWAP_ACTIVE.includes(o.status)
          ? `<button class="btn btn-ghost btn-sm" onclick="cancelMyOutgoing('${o.request_id}')">Отозвать</button>` : '';
        const eff = o.effective_from ? `Вступает в силу с ${formatDateRu(o.effective_from)}` : '';
        return `
          <div class="req-card ${meta.cls}">
            <div class="req-info" style="grid-column:1 / -1;">
              <div style="display:flex;justify-content:space-between;align-items:center;gap:12px;flex-wrap:wrap;">
                <div class="req-change" style="margin:0;">
                  <span class="req-change-from">Отдаю: ${give}</span>
                  <span class="req-change-arrow">→</span>
                  <span class="req-change-to">Получаю: ${get}</span>
                </div>
                <span class="req-status-badge ${meta.cls}">${meta.label}</span>
              </div>
              ${eff ? `<div class="req-date" style="margin-top:6px;">${eff}</div>` : ''}
              ${cancel ? `<div style="margin-top:8px;text-align:right;">${cancel}</div>` : ''}
            </div>
          </div>`;
      }).join('')}`;
  }

  let applyListingId = null;
  function openApplyModal(listingId, label) {
    applyListingId = listingId;
    hideError('apply-error');
    document.getElementById('apply-target-window').textContent = label;
    document.getElementById('apply-reason').value = '';
    document.getElementById('apply-modal').classList.add('open');
  }
  function closeApplyModal() {
    document.getElementById('apply-modal').classList.remove('open');
    applyListingId = null;
  }
  async function confirmApply() {
    if (!applyListingId) return;
    hideError('apply-error');
    const btn = document.getElementById('btn-confirm-apply');
    const reason = document.getElementById('apply-reason').value.trim() || null;
    btn.disabled = true; btn.textContent = 'Отправка…';
    const { error } = await sb.rpc('swap_request_create', { p_listing_id: applyListingId, p_reason: reason });
    btn.disabled = false; btn.textContent = 'Отправить заявку';
    if (error) { showError('apply-error', swapErr(error)); return; }
    closeApplyModal();
    alert('Заявка отправлена. Менеджер согласует обмен с обеими сторонами.');
    await loadSwapMarket();
  }
  async function cancelMyOutgoing(reqId) {
    if (!confirm('Отозвать заявку на обмен?')) return;
    const { error } = await sb.rpc('swap_request_cancel', { p_request_id: reqId });
    if (error) { showError('swap-error', swapErr(error)); return; }
    await loadSwapMarket();
  }

  // ──── Менеджер: обмены графиков ────────────────────────────────────
  async function loadSwaps() {
    hideError('swaps-error');
    const list = document.getElementById('swaps-list');
    list.innerHTML = '<div class="loading-state">Загрузка…</div>';
    const filter = document.getElementById('swaps-filter').value;

    const { data, error } = await sb.rpc('swap_board');
    if (error) { showError('swaps-error', swapErr(error)); list.innerHTML = ''; return; }
    let rows = data || [];

    const counts = { active: 0, approved: 0, rejected: 0 };
    for (const r of rows) {
      if (SWAP_ACTIVE.includes(r.status)) counts.active++;
      else if (r.status === 'approved') counts.approved++;
      else counts.rejected++;
    }
    document.getElementById('swaps-subtitle').textContent =
      `${counts.active} в работе · ${counts.approved} выполнено · ${counts.rejected} отклонено/отменено`;
    updateSwapBadge(counts.active);

    if (filter === 'active') rows = rows.filter(r => SWAP_ACTIVE.includes(r.status));
    else if (filter === 'approved') rows = rows.filter(r => r.status === 'approved');
    else if (filter === 'rejected') rows = rows.filter(r => r.status === 'rejected' || r.status === 'cancelled');

    if (rows.length === 0) {
      list.innerHTML = `<div class="empty-state"><div class="empty-state-text">Нет обменов по выбранному фильтру.</div></div>`;
      return;
    }
    list.innerHTML = rows.map(renderSwapCard).join('');
  }

  function renderSwapCard(r) {
    const meta = SWAP_STATUS_META[r.status] || { label: r.status, cls: '' };
    const give = fmtWin(r.give_start, r.give_end, r.give_minutes);   // окно А → получит Б
    const get  = fmtWin(r.get_start, r.get_end, r.get_minutes);      // окно Б → получит А
    const rid = r.request_id;
    const rn = escapeHtml(r.requester_name), tn = escapeHtml(r.target_name);

    let actions = '';
    if (r.status === 'open') {
      actions = `
        <div class="swap-actions">
          <label class="swap-check">
            <input type="checkbox" id="agree-${rid}" ${r.target_agreed ? 'checked' : ''} onchange="swapSetAgreed('${rid}', this.checked)">
            Согласие ${tn} получено
          </label>
          <div style="display:flex;gap:8px;">
            <button class="btn btn-ghost btn-sm" onclick="swapRejectReq('${rid}')">Отклонить</button>
            <button class="btn btn-sm" onclick="swapApprove('${rid}')" ${r.target_agreed ? '' : 'disabled'}>Одобрить</button>
          </div>
        </div>`;
    } else if (r.status === 'manager_approved') {
      actions = `
        <div class="swap-actions" style="justify-content:flex-end;">
          <button class="btn btn-ghost btn-sm" onclick="swapRejectReq('${rid}')">Отклонить</button>
          <button class="btn btn-sm" onclick="swapSendIcall('${rid}')">Отправить в I-Call</button>
        </div>`;
    } else if (r.status === 'pending_icall') {
      const _nm = new Date(new Date().getFullYear(), new Date().getMonth() + 1, 1);
      const defEff = `${_nm.getFullYear()}-${String(_nm.getMonth() + 1).padStart(2, '0')}-01`;
      actions = `
        <div class="swap-icall">
          <div class="swap-subhead" style="margin-top:0;">Подтверждение I-Call</div>
          <div class="field-row">
            <div class="field" style="margin:0;">
              <label class="field-label">Действует с</label>
              <input type="date" id="eff-${rid}" class="input" value="${defEff}">
            </div>
            <div class="field" style="margin:0;">
              <label class="field-label">Референс I-Call</label>
              <input type="text" id="ref-${rid}" class="input" placeholder="№ / дата письма">
            </div>
          </div>
          <div style="font-size:12px;color:#94A3B8;margin-top:8px;">💡 Рекомендуется 1-е число месяца — тогда овертайм за месяц обмена будет точным.</div>
          <label class="swap-check" style="margin-top:10px;">
            <input type="checkbox" id="ovr-${rid}"> Обойти заморозку (форс-мажор)
          </label>
          <div style="display:flex;gap:8px;justify-content:flex-end;margin-top:10px;">
            <button class="btn btn-ghost btn-sm" onclick="swapRejectReq('${rid}')">Отклонить</button>
            <button class="btn btn-sm" onclick="swapIcallApprove('${rid}')">I-Call одобрил — выполнить обмен</button>
          </div>
        </div>`;
    } else {
      const extra = [];
      if (r.effective_from) extra.push(`с ${formatDateRu(r.effective_from)}`);
      if (r.override_freeze) extra.push('обход заморозки');
      if (r.decision_note) extra.push(escapeHtml(r.decision_note));
      actions = extra.length ? `<div class="req-date" style="margin-top:8px;">${extra.join(' · ')}</div>` : '';
    }

    return `
      <div class="req-card ${meta.cls}">
        <div class="req-info" style="grid-column:1 / -1;">
          <div style="display:flex;justify-content:space-between;align-items:baseline;gap:12px;flex-wrap:wrap;">
            <div class="req-translator-name">${rn} ⇄ ${tn}
              <span style="font-weight:400;color:#94A3B8;font-size:12px;">· ${escapeHtml(r.pair_code || '')}</span></div>
            <span class="req-status-badge ${meta.cls}">${meta.label}</span>
          </div>
          <div class="req-change" style="margin-top:10px;">
            <span class="req-change-from">${rn}: ${give}</span>
            <span class="req-change-arrow">⇄</span>
            <span class="req-change-to">${tn}: ${get}</span>
          </div>
          ${r.reason ? `<div class="req-reason">${escapeHtml(r.reason)}</div>` : ''}
          ${actions}
        </div>
      </div>`;
  }

  async function swapSetAgreed(rid, agreed) {
    const { error } = await sb.rpc('swap_mark_target_agreed', { p_request_id: rid, p_agreed: agreed, p_note: null });
    if (error) showError('swaps-error', swapErr(error));
    await loadSwaps();
  }
  async function swapApprove(rid) {
    const { error } = await sb.rpc('swap_manager_approve', { p_request_id: rid });
    if (error) { showError('swaps-error', swapErr(error)); return; }
    await loadSwaps();
  }
  async function swapSendIcall(rid) {
    const { error } = await sb.rpc('swap_send_to_icall', { p_request_id: rid });
    if (error) { showError('swaps-error', swapErr(error)); return; }
    await loadSwaps();
  }
  async function swapIcallApprove(rid) {
    hideError('swaps-error');
    const eff = document.getElementById('eff-' + rid).value;
    const ref = document.getElementById('ref-' + rid).value.trim() || null;
    const ovr = document.getElementById('ovr-' + rid).checked;
    if (!eff) { showError('swaps-error', 'Укажите дату вступления в силу.'); return; }
    if (!confirm(`Подтвердить обмен? Графики поменяются местами с ${formatDateRu(eff)}. Оба участника получат заморозку на 2 месяца.`)) return;
    const { error } = await sb.rpc('swap_icall_approve', { p_request_id: rid, p_effective_from: eff, p_icall_ref: ref, p_override: ovr });
    if (error) { showError('swaps-error', swapErr(error)); return; }
    alert('Обмен выполнен. Графики обновлены, заморозка установлена на 2 месяца.');
    await loadSwaps();
  }
  async function swapRejectReq(rid) {
    const note = prompt('Причина отклонения (опционально):');
    if (note === null) return; // отмена диалога
    const { error } = await sb.rpc('swap_reject', { p_request_id: rid, p_note: note.trim() || null });
    if (error) { showError('swaps-error', swapErr(error)); return; }
    await loadSwaps();
  }

  async function refreshSwapBadge() {
    if (!currentUser || currentUser.role !== 'manager') return;
    const { count } = await sb.from('schedule_swap_requests')
      .select('*', { count: 'exact', head: true })
      .in('status', SWAP_ACTIVE);
    updateSwapBadge(count || 0);
  }
  function updateSwapBadge(count) {
    const badge = document.getElementById('sb-swap-badge');
    if (!badge) return;
    if (count > 0) { badge.textContent = String(count); badge.classList.remove('hidden'); }
    else badge.classList.add('hidden');
  }

  function showError(elId, msg) {
    const el = document.getElementById(elId);
    el.textContent = msg;
    el.classList.remove('hidden');
  }
  function hideError(elId) {
    document.getElementById(elId).classList.add('hidden');
  }
  function escapeHtml(s) {
    return String(s || '').replace(/[&<>"']/g, c => ({
      '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
    }[c]));
  }
  function pluralize(n, one, few, many) {
    const mod10 = n % 10, mod100 = n % 100;
    if (mod10 === 1 && mod100 !== 11) return one;
    if (mod10 >= 2 && mod10 <= 4 && (mod100 < 10 || mod100 >= 20)) return few;
    return many;
  }

  document.getElementById('btn-login').addEventListener('click', handleLogin);
  document.getElementById('btn-logout').addEventListener('click', handleLogout);
  document.getElementById('password').addEventListener('keydown', e => {
    if (e.key === 'Enter') handleLogin();
  });
  document.getElementById('email').addEventListener('keydown', e => {
    if (e.key === 'Enter') document.getElementById('password').focus();
  });

  // Восстановление пароля
  document.getElementById('btn-forgot-link').addEventListener('click', (e) => {
    e.preventDefault();
    hideError('forgot-error');
    document.getElementById('forgot-form-wrap').classList.remove('hidden');
    document.getElementById('forgot-success').classList.add('hidden');
    document.getElementById('forgot-email').value =
      document.getElementById('email').value;
    showScreen('forgot');
    setTimeout(() => document.getElementById('forgot-email').focus(), 50);
  });
  document.getElementById('btn-back-to-login').addEventListener('click', (e) => {
    e.preventDefault();
    showScreen('login');
  });
  document.getElementById('btn-forgot-send').addEventListener('click', handleForgotPassword);
  document.getElementById('forgot-email').addEventListener('keydown', e => {
    if (e.key === 'Enter') handleForgotPassword();
  });
  document.getElementById('btn-reset-save').addEventListener('click', handleResetPassword);
  document.getElementById('reset-password2').addEventListener('keydown', e => {
    if (e.key === 'Enter') handleResetPassword();
  });
  document.getElementById('reset-password').addEventListener('keydown', e => {
    if (e.key === 'Enter') document.getElementById('reset-password2').focus();
  });
  document.querySelectorAll('.sb-item').forEach(b => {
    b.addEventListener('click', async () => {
      const page = b.dataset.page;
      // Если уже на этой странице — не перезагружаем (экономит цепочку запросов)
      if (page === currentPage) return;
      currentPage = page;
      showPage(page);
      if (page === 'dashboard') await loadDashboard();
      else if (page === 'translators') await loadTranslators();
      else if (page === 'managers') await loadManagers();
      else if (page === 'requests') await loadRequests();
      else if (page === 'swaps') await loadSwaps();
      else if (page === 'my-profile') await loadMyProfile();
      else if (page === 'swap-market') await loadSwapMarket();
      else if (page === 'calendar') await loadCalendar();
      else if (page === 'payroll') await loadPayroll();
      else if (page === 'clients') {
        // Защита: секция клиентов недоступна без флага
        if (!currentUser || !currentUser.can_access_clients) {
          showError('clients-error', 'Нет доступа к секции «Клиенты».');
          return;
        }
        await loadClients();
      }
    });
  });
  document.getElementById('add-modal').addEventListener('click', e => {
    if (e.target.id === 'add-modal') closeAddModal();
  });
  document.getElementById('day-modal').addEventListener('click', e => {
    if (e.target.id === 'day-modal') closeDayModal();
  });

  document.getElementById('cal-prev').addEventListener('click', () => changeMonth(-1));
  document.getElementById('cal-next').addEventListener('click', () => changeMonth(1));
  document.getElementById('td-cal-prev').addEventListener('click', () => changeTdMonth(-1));
  document.getElementById('td-cal-next').addEventListener('click', () => changeTdMonth(1));
  document.getElementById('day-view-modal').addEventListener('click', e => {
    if (e.target.id === 'day-view-modal') closeDayViewModal();
  });
  document.getElementById('edit-modal').addEventListener('click', e => {
    if (e.target.id === 'edit-modal') closeEditModal();
  });
  document.getElementById('shift-modal').addEventListener('click', e => {
    if (e.target.id === 'shift-modal') closeShiftModal();
  });
  document.getElementById('shift-request-modal').addEventListener('click', e => {
    if (e.target.id === 'shift-request-modal') closeShiftRequestModal();
  });
  document.getElementById('add-manager-modal').addEventListener('click', e => {
    if (e.target.id === 'add-manager-modal') closeAddManagerModal();
  });
  document.getElementById('add-client-modal').addEventListener('click', e => {
    if (e.target.id === 'add-client-modal') closeAddClientModal();
  });
  document.getElementById('edit-client-modal').addEventListener('click', e => {
    if (e.target.id === 'edit-client-modal') closeEditClientModal();
  });
  document.getElementById('add-prepayment-modal').addEventListener('click', e => {
    if (e.target.id === 'add-prepayment-modal') closeAddPrepaymentModal();
  });
  document.getElementById('attach-translator-modal').addEventListener('click', e => {
    if (e.target.id === 'attach-translator-modal') closeAttachTranslatorModal();
  });
  document.getElementById('req-filter').addEventListener('change', () => loadRequests());

  // Заполняем select'ы часовых поясов из единого массива TIMEZONES
  fillTimezoneSelects();

  checkSession();

  // ═══════════════════════════════════════════════════════════════════
  // Mobile navigation: burger menu + bottom nav + topbar title
  // ═══════════════════════════════════════════════════════════════════
  // Использует ДЕЛЕГИРОВАНИЕ событий через document.addEventListener —
  // это работает независимо от того, когда появились элементы в DOM
  // и какие расширения браузера установлены.
  // ═══════════════════════════════════════════════════════════════════
  function setupMobileNav() {
    try {
      console.log('[mobile-nav] setup started');

      function openDrawer() {
        const sidebar = document.querySelector('.sidebar');
        const backdrop = document.getElementById('mobile-backdrop');
        if (sidebar) sidebar.classList.add('mobile-open');
        if (backdrop) backdrop.classList.add('open');
      }

      function closeDrawer() {
        const sidebar = document.querySelector('.sidebar');
        const backdrop = document.getElementById('mobile-backdrop');
        if (sidebar) sidebar.classList.remove('mobile-open');
        if (backdrop) backdrop.classList.remove('open');
      }

      function updateMobileTopbarTitle(activeItem) {
        const topbarTitle = document.getElementById('mobile-topbar-title');
        if (!topbarTitle || !activeItem) return;
        const text = activeItem.textContent.trim().split('\n')[0].trim();
        topbarTitle.textContent = text || 'LinguaTime';
      }

      function syncBottomNav(page) {
        const bottomNav = document.getElementById('mobile-bottom-nav-translator');
        if (!bottomNav) return;
        bottomNav.querySelectorAll('.mobile-bottom-nav-item').forEach(b => {
          b.classList.toggle('active', b.dataset.page === page);
        });
      }

      function refreshBottomNavVisibility() {
        const bottomNav = document.getElementById('mobile-bottom-nav-translator');
        if (!bottomNav) return;
        const transNav = document.getElementById('nav-translator');
        const isTranslator = transNav && !transNav.classList.contains('hidden');
        bottomNav.style.display = isTranslator ? 'flex' : 'none';
      }

      // ── ГЛОБАЛЬНЫЙ обработчик кликов через делегирование ────────
      // Работает на любом элементе с нужным id/классом, в т.ч. если
      // он появился позже в DOM. Не зависит от момента загрузки.
      document.addEventListener('click', function(e) {
        // Тап по бургеру → открыть drawer
        if (e.target.closest('#mobile-burger-btn')) {
          openDrawer();
          return;
        }
        // Тап по затемнению → закрыть drawer
        if (e.target.closest('#mobile-backdrop')) {
          closeDrawer();
          return;
        }
        // Тап по пункту sidebar (в открытом drawer) → закрыть + обновить заголовок
        const sbItem = e.target.closest('.sb-item');
        if (sbItem) {
          closeDrawer();
          updateMobileTopbarTitle(sbItem);
          syncBottomNav(sbItem.dataset.page);
          return;
        }
        // Тап по нижней навигации переводчика → кликаем соответствующий sb-item
        const navItem = e.target.closest('.mobile-bottom-nav-item');
        if (navItem) {
          const page = navItem.dataset.page;
          const target = document.querySelector('.sb-item[data-page="' + page + '"]');
          if (target) target.click();
          return;
        }
      });

      // Наблюдаем за изменением видимости nav-translator
      const transNav = document.getElementById('nav-translator');
      if (transNav) {
        const observer = new MutationObserver(refreshBottomNavVisibility);
        observer.observe(transNav, { attributes: true, attributeFilter: ['class'] });
      }
      refreshBottomNavVisibility();

      // На старте — обновляем заголовок по активному пункту
      const initialActive = document.querySelector('.sb-item.active');
      if (initialActive) updateMobileTopbarTitle(initialActive);

      console.log('[mobile-nav] setup completed');
    } catch (err) {
      console.error('[mobile-nav] setup failed:', err);
    }
  }

  // Запускаем сразу + страховка через DOMContentLoaded
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', setupMobileNav);
  } else {
    setupMobileNav();
  }
