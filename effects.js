// ═══════════════════════════════════════════════════════════════════════
// LinguaTime — анимационные эффекты интерфейса.
// Based on React Bits by David Haz (MIT + Commons Clause) — https://github.com/DavidHDev/react-bits
// Портированы в чистый JS/CSS для этого приложения: CountUp, SpotlightCard,
// AnimatedList, BlurText, FadeContent (CSS в styles.css) и Aurora (WebGL2).
// Не отдельная библиотека — используется только как часть LinguaTime.
//
// app.js не трогаем: эффекты подключаются через data-атрибуты в разметке
// и MutationObserver — app.js пишет значения как раньше, а здесь они анимируются.
//   data-fx-count  — CountUp: числа внутри элемента «досчитываются»
//   data-fx-list   — AnimatedList: новые строки появляются каскадом
//   data-fx-bars   — столбики графика вырастают снизу
//   data-fx-blur   — BlurText: заголовок проявляется по словам из размытия
// При prefers-reduced-motion всё выключено, значения выводятся сразу.
// ═══════════════════════════════════════════════════════════════════════
(function () {
  'use strict';

  const reduceMq = window.matchMedia('(prefers-reduced-motion: reduce)');
  const reduced = () => reduceMq.matches;
  const isShown = el => el.getClientRects().length > 0;
  const easeOut = t => 1 - Math.pow(1 - t, 4);

  // ── CountUp ─────────────────────────────────────────────────────────
  // Анимируются только текстовые узлы, разметка внутри (kpi-unit) не меняется.
  // Все числа элемента идут с общим прогрессом: «12ч 30м» считается как 0ч 00м → 12ч 30м.
  // Повторное обновление (например, после правки дня) считает от прошлого значения.
  const COUNT_MS = 900;
  const NUM_RE = /\d+(?:\.\d+)?/g;

  function parseNums(el) {
    const parts = [];
    const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
    let node, idx = 0;
    while ((node = walker.nextNode())) {
      const text = node.nodeValue;
      const chunks = [];
      let last = 0, m;
      NUM_RE.lastIndex = 0;
      while ((m = NUM_RE.exec(text))) {
        chunks.push(text.slice(last, m.index));
        const tok = m[0];
        const dot = tok.indexOf('.');
        chunks.push({
          to: parseFloat(tok),
          dec: dot === -1 ? 0 : tok.length - dot - 1,
          // минуты в «Xч 05м» — второе и далее число сохраняет ширину с ведущим нулём
          pad: idx > 0 ? (dot === -1 ? tok.length : dot) : 0,
        });
        idx++;
        last = m.index + tok.length;
      }
      if (chunks.length) {
        chunks.push(text.slice(last));
        parts.push({ node, chunks });
      }
    }
    return parts;
  }

  function fmtNum(v, n) {
    let s = v.toFixed(n.dec);
    if (n.pad) {
      const [i, d] = s.split('.');
      s = i.padStart(n.pad, '0') + (d !== undefined ? '.' + d : '');
    }
    return s;
  }

  function renderNums(parts, values) {
    let k = 0;
    for (const p of parts) {
      p.node.nodeValue = p.chunks.map(c => (typeof c === 'string' ? c : fmtNum(values[k++], c))).join('');
    }
  }

  function countUp(el) {
    cancelAnimationFrame(el._fxRaf || 0);
    const parts = parseNums(el);
    const nums = parts.flatMap(p => p.chunks.filter(c => typeof c !== 'string'));
    const to = nums.map(n => n.to);
    const shape = parts.map(p => p.chunks.map(c => (typeof c === 'string' ? c : '#')).join('')).join('|');
    const prev = el._fxLast && el._fxLast.shape === shape ? el._fxLast.values : null;
    el._fxLast = { shape, values: to };

    if (!nums.length || reduced() || !isShown(el)) return;
    const from = prev || to.map(() => 0);
    if (from.every((v, i) => v === to[i])) return;

    const start = performance.now();
    const tick = now => {
      const t = Math.min(1, (now - start) / COUNT_MS);
      const e = easeOut(t);
      renderNums(parts, from.map((f, i) => f + (to[i] - f) * e));
      if (t < 1) el._fxRaf = requestAnimationFrame(tick);
    };
    // MutationObserver срабатывает до отрисовки — первый кадр уже стартовое значение
    renderNums(parts, from);
    el._fxRaf = requestAnimationFrame(tick);
  }

  // ── AnimatedList ────────────────────────────────────────────────────
  // Каскадное появление строк (в оригинале scale .7 → 1; здесь мягче — .98 и сдвиг 6px,
  // под стиль apple.com). Анимируются только новые строки: после действия в списке
  // (одобрили запрос) прыгает только изменившаяся карточка, а не весь список.
  // При каждом открытии раздела память сбрасывается — каскад проигрывается заново.
  const LIST_MAX = 12;
  const LIST_STEP_MS = 40;
  const seen = new WeakMap();
  const sig = el => el.textContent.replace(/\s+/g, ' ').trim().slice(0, 240);

  function animateList(box) {
    const known = seen.get(box) || new Set();
    const next = new Set();
    let n = 0;
    for (const child of box.children) {
      if (child.matches('.loading-state, .empty-state, .dash-empty')) continue;
      const key = sig(child);
      next.add(key);
      if (known.has(key) || reduced() || n >= LIST_MAX) continue;
      child.style.setProperty('--fx-d', (n++ * LIST_STEP_MS) + 'ms');
      child.classList.add('fx-item');
      child.addEventListener('animationend', function done(e) {
        if (e.target !== child) return;
        child.classList.remove('fx-item');
        child.removeEventListener('animationend', done);
      });
    }
    seen.set(box, next);
  }

  // ── Столбики графика ────────────────────────────────────────────────
  function growBars(box) {
    if (reduced() || !isShown(box)) return;
    let i = 0;
    for (const bar of box.children) bar.style.setProperty('--fx-i', i++);
    box.classList.remove('fx-grow');
    void box.offsetWidth; // перезапуск анимации
    box.classList.add('fx-grow');
    clearTimeout(box._fxT);
    box._fxT = setTimeout(() => box.classList.remove('fx-grow'), 1200);
  }

  // ── BlurText ────────────────────────────────────────────────────────
  // Слова проявляются по очереди: blur(10px) → 0, снизу вверх. Текст для скринридеров
  // остаётся целым (aria-label), анимированные span скрыты от них.
  function blurText(el) {
    const text = el.textContent;
    if (reduced() || !text.trim() || !isShown(el)) return;
    const words = text.trim().split(/\s+/);
    const frag = document.createDocumentFragment();
    words.forEach((w, i) => {
      const s = document.createElement('span');
      s.className = 'fx-word';
      s.setAttribute('aria-hidden', 'true');
      s.style.setProperty('--fx-d', (i * 120) + 'ms');
      s.textContent = w;
      frag.appendChild(s);
      if (i < words.length - 1) frag.appendChild(document.createTextNode(' '));
    });
    el._fxSelf = true;
    el.setAttribute('aria-label', text.trim());
    el.replaceChildren(frag);
  }

  // ── Наблюдение за разметкой ─────────────────────────────────────────
  const handlers = [
    ['[data-fx-count]', countUp],
    ['[data-fx-list]', animateList],
    ['[data-fx-bars]', growBars],
    ['[data-fx-blur]', el => {
      if (el._fxSelf) { el._fxSelf = false; return; } // наша же замена на span
      el.removeAttribute('aria-label');
      blurText(el);
    }],
  ];

  function watch() {
    const mo = new MutationObserver(records => {
      const done = new Set();
      for (const r of records) {
        if (done.has(r.target)) continue;
        done.add(r.target);
        for (const [sel, fn] of handlers) if (r.target.matches(sel)) fn(r.target);
      }
    });
    for (const [sel, fn] of handlers) {
      document.querySelectorAll(sel).forEach(el => {
        mo.observe(el, { childList: true });
        if (fn === countUp) countUp(el); // запоминаем исходную форму значения
      });
    }

    // Открыли раздел — сбрасываем память списков, чтобы каскад проиграл заново
    const pageMo = new MutationObserver(records => {
      for (const r of records) {
        if (r.target.classList.contains('hidden')) continue;
        r.target.querySelectorAll('[data-fx-list]').forEach(box => seen.delete(box));
      }
    });
    document.querySelectorAll('.page').forEach(p => pageMo.observe(p, { attributes: true, attributeFilter: ['class'] }));
  }

  // ── SpotlightCard ───────────────────────────────────────────────────
  // Мягкий свет внутри плиток под курсором. Один делегированный слушатель на документ;
  // на тач-экранах не включается. Сам градиент — в styles.css.
  const SPOT_SEL = '.kpi, .dash-kpi, .dash-card';
  function spotlight() {
    if (!window.matchMedia('(hover: hover) and (pointer: fine)').matches) return;
    let pending = null, raf = 0;
    document.addEventListener('pointermove', e => {
      const card = e.target.closest && e.target.closest(SPOT_SEL);
      if (!card) return;
      pending = { card, x: e.clientX, y: e.clientY };
      if (raf) return;
      raf = requestAnimationFrame(() => {
        raf = 0;
        const { card: c, x, y } = pending;
        const r = c.getBoundingClientRect();
        c.style.setProperty('--fx-mx', (x - r.left) + 'px');
        c.style.setProperty('--fx-my', (y - r.top) + 'px');
      });
    }, { passive: true });
  }

  // ── Aurora (экран входа) ────────────────────────────────────────────
  // Шейдер — из React Bits Aurora; вместо ogl — чистый WebGL2 (один треугольник на весь
  // экран), чтобы не тянуть зависимость. Рисуется только пока открыт вход /
  // восстановление пароля, после входа цикл останавливается и контекст освобождается.
  const AURORA_VERT = `#version 300 es
in vec2 position;
void main() { gl_Position = vec4(position, 0.0, 1.0); }`;

  const AURORA_FRAG = `#version 300 es
precision highp float;
uniform float uTime;
uniform float uAmplitude;
uniform vec3 uColorStops[3];
uniform vec2 uResolution;
uniform float uBlend;
uniform float uLightMode;
out vec4 fragColor;

vec3 permute(vec3 x) { return mod(((x * 34.0) + 1.0) * x, 289.0); }

float snoise(vec2 v) {
  const vec4 C = vec4(0.211324865405187, 0.366025403784439, -0.577350269189626, 0.024390243902439);
  vec2 i  = floor(v + dot(v, C.yy));
  vec2 x0 = v - i + dot(i, C.xx);
  vec2 i1 = (x0.x > x0.y) ? vec2(1.0, 0.0) : vec2(0.0, 1.0);
  vec4 x12 = x0.xyxy + C.xxzz;
  x12.xy -= i1;
  i = mod(i, 289.0);
  vec3 p = permute(permute(i.y + vec3(0.0, i1.y, 1.0)) + i.x + vec3(0.0, i1.x, 1.0));
  vec3 m = max(0.5 - vec3(dot(x0, x0), dot(x12.xy, x12.xy), dot(x12.zw, x12.zw)), 0.0);
  m = m * m;
  m = m * m;
  vec3 x = 2.0 * fract(p * C.www) - 1.0;
  vec3 h = abs(x) - 0.5;
  vec3 ox = floor(x + 0.5);
  vec3 a0 = x - ox;
  m *= 1.79284291400159 - 0.85373472095314 * (a0 * a0 + h * h);
  vec3 g;
  g.x  = a0.x  * x0.x  + h.x  * x0.y;
  g.yz = a0.yz * x12.xz + h.yz * x12.yw;
  return 130.0 * dot(m, g);
}

void main() {
  vec2 uv = gl_FragCoord.xy / uResolution;
  vec3 rampColor = uv.x < 0.5
    ? mix(uColorStops[0], uColorStops[1], uv.x * 2.0)
    : mix(uColorStops[1], uColorStops[2], (uv.x - 0.5) * 2.0);

  float height = snoise(vec2(uv.x * 2.0 + uTime * 0.1, uTime * 0.25)) * 0.5 * uAmplitude;
  height = exp(height);
  height = (uv.y * 2.0 - height + 0.2);
  float intensity = 0.6 * height;

  float midPoint = 0.20;
  float auroraAlpha = smoothstep(midPoint - uBlend * 0.5, midPoint + uBlend * 0.5, intensity);
  vec3 auroraColor = intensity * rampColor;

  if (uLightMode > 0.5) {
    float energy = clamp(max(intensity, 0.0), 0.0, 1.0);
    float coverage = clamp(auroraAlpha * (0.55 + 0.45 * energy), 0.0, 0.86);
    vec3 chroma = pow(clamp(rampColor, 0.0, 1.0), vec3(1.2));
    float chromaPeak = max(chroma.r, max(chroma.g, chroma.b));
    chroma /= max(chromaPeak, 0.0001);
    float a = min(coverage * 1.08, 0.94);
    fragColor = vec4(chroma * a, a);
  } else {
    fragColor = vec4(auroraColor * auroraAlpha, auroraAlpha);
  }
}`;

  // Цвета — из палитры приложения: синий акцент → бирюзовый → сиреневый
  const AURORA_THEMES = {
    light: { stops: ['#0071e3', '#5ac8fa', '#af52de'], light: 1 },
    dark: { stops: ['#0a84ff', '#40c8e0', '#bf5af2'], light: 0 },
  };
  const AURORA_SPEED = 0.5;

  const hexRgb = h => [1, 3, 5].map(i => parseInt(h.slice(i, i + 2), 16) / 255);

  function createAurora() {
    const wrap = document.createElement('div');
    wrap.className = 'fx-aurora';
    wrap.setAttribute('aria-hidden', 'true');
    const canvas = document.createElement('canvas');
    wrap.appendChild(canvas);

    const gl = canvas.getContext('webgl2', { alpha: true, premultipliedAlpha: true, antialias: false });
    if (!gl) return null;

    const sh = (type, src) => {
      const s = gl.createShader(type);
      gl.shaderSource(s, src);
      gl.compileShader(s);
      return gl.getShaderParameter(s, gl.COMPILE_STATUS) ? s : null;
    };
    const vs = sh(gl.VERTEX_SHADER, AURORA_VERT);
    const fs = sh(gl.FRAGMENT_SHADER, AURORA_FRAG);
    if (!vs || !fs) return null;
    const prog = gl.createProgram();
    gl.attachShader(prog, vs);
    gl.attachShader(prog, fs);
    gl.linkProgram(prog);
    if (!gl.getProgramParameter(prog, gl.LINK_STATUS)) return null;
    gl.useProgram(prog);

    const buf = gl.createBuffer();
    gl.bindBuffer(gl.ARRAY_BUFFER, buf);
    gl.bufferData(gl.ARRAY_BUFFER, new Float32Array([-1, -1, 3, -1, -1, 3]), gl.STATIC_DRAW);
    const loc = gl.getAttribLocation(prog, 'position');
    gl.enableVertexAttribArray(loc);
    gl.vertexAttribPointer(loc, 2, gl.FLOAT, false, 0, 0);

    const u = name => gl.getUniformLocation(prog, name);
    const uTime = u('uTime'), uRes = u('uResolution'), uStops = u('uColorStops[0]'), uLight = u('uLightMode');
    gl.uniform1f(u('uAmplitude'), 1.0);
    gl.uniform1f(u('uBlend'), 0.5);
    gl.clearColor(0, 0, 0, 0);
    gl.enable(gl.BLEND);
    gl.blendFunc(gl.ONE, gl.ONE_MINUS_SRC_ALPHA);

    const darkMq = window.matchMedia('(prefers-color-scheme: dark)');
    const applyTheme = () => {
      const th = darkMq.matches ? AURORA_THEMES.dark : AURORA_THEMES.light;
      gl.uniform3fv(uStops, th.stops.flatMap(hexRgb));
      gl.uniform1f(uLight, th.light);
    };
    applyTheme();

    const resize = () => {
      const dpr = Math.min(window.devicePixelRatio || 1, 1.5);
      const w = Math.max(1, Math.round(wrap.clientWidth * dpr));
      const h = Math.max(1, Math.round(wrap.clientHeight * dpr));
      if (canvas.width !== w || canvas.height !== h) {
        canvas.width = w;
        canvas.height = h;
        gl.viewport(0, 0, w, h);
        gl.uniform2f(uRes, w, h);
      }
    };
    const draw = t => {
      resize();
      gl.uniform1f(uTime, t * 0.01 * AURORA_SPEED * 0.1);
      gl.clear(gl.COLOR_BUFFER_BIT);
      gl.drawArrays(gl.TRIANGLES, 0, 3);
    };

    let raf = 0, running = false;
    const loop = t => { draw(t); raf = requestAnimationFrame(loop); };
    const start = () => {
      if (running) return;
      running = true;
      // reduced motion — один неподвижный кадр вместо анимации
      if (reduced()) { requestAnimationFrame(() => draw(4000)); return; }
      raf = requestAnimationFrame(loop);
    };
    const stop = () => { running = false; cancelAnimationFrame(raf); };
    const redraw = () => { if (running && reduced()) draw(4000); };

    const off = new AbortController();
    const opts = { signal: off.signal };
    darkMq.addEventListener('change', () => { applyTheme(); redraw(); }, opts);
    window.addEventListener('resize', redraw, opts);
    document.addEventListener('visibilitychange', () => {
      if (document.hidden) stop(); else start();
    }, opts);

    return {
      el: wrap,
      start,
      destroy() {
        stop();
        off.abort();
        wrap.remove();
        gl.getExtension('WEBGL_lose_context')?.loseContext();
      },
    };
  }

  const AUTH_SCREENS = ['screen-login', 'screen-forgot', 'screen-reset'];
  function aurora() {
    let inst = null;
    const sync = () => {
      const screen = AUTH_SCREENS.map(id => document.getElementById(id))
        .find(el => el && !el.classList.contains('hidden'));
      if (!screen) {
        if (inst) { inst.destroy(); inst = null; }
        return;
      }
      if (!inst) inst = createAurora();
      if (!inst) return; // нет WebGL2 — остаётся обычный фон
      if (inst.el.parentNode !== screen) screen.prepend(inst.el);
      inst.start();
    };
    const mo = new MutationObserver(sync);
    AUTH_SCREENS.concat('screen-app').forEach(id => {
      const el = document.getElementById(id);
      if (el) mo.observe(el, { attributes: true, attributeFilter: ['class'] });
    });
    sync();
  }

  function init() {
    watch();
    spotlight();
    aurora();
  }
  // Скрипт подключён в конце body перед app.js — разметка уже есть, ждать DOMContentLoaded
  // нельзя: app.js к тому времени может успеть заполнить экраны
  init();
})();
