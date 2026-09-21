/* ============================================================
   APCSA 题库 · 公共脚本
   ============================================================ */
(function () {
  'use strict';

  var DATA = window.QB_DATA || { meta: {}, questions: [] };
  var TOPICS = window.QB_TOPICS || { units: [] };

  /* ---------- 索引 ---------- */
  var TOPIC_INDEX = {};   // "4.3" -> {unit, name, nameCn, kw}
  var UNIT_INDEX = {};    // "U4" -> unit object
  TOPICS.units.forEach(function (u) {
    UNIT_INDEX[u.id] = u;
    u.topics.forEach(function (t) {
      TOPIC_INDEX[t.id] = { unit: u.id, name: t.name, nameCn: t.nameCn, kw: t.kw || [] };
    });
  });

  /* ---------- 工具 ---------- */
  function esc(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }
  function qs(name) {
    var m = new RegExp('[?&]' + name + '=([^&#]*)').exec(location.search);
    return m ? decodeURIComponent(m[1].replace(/\+/g, ' ')) : null;
  }
  function unitColor(uid) { return (UNIT_INDEX[uid] && UNIT_INDEX[uid].color) || '#64748b'; }
  function unitShort(uid) { return (UNIT_INDEX[uid] && UNIT_INDEX[uid].code) || uid || '未归类'; }
  function topicName(tid) {
    if (!tid) return '未标注';
    var t = TOPIC_INDEX[tid];
    return t ? (tid + ' ' + t.nameCn) : tid;
  }

  var CODE_LINE = /(;|\{|\}|^\s*\/\/|^\s*\*|^\s*(public|private|protected|static|void|int|double|boolean|char|String|if|for|while|do|return|else|class|new|final)\b)/;

  /* 把「散文 + 代码」混合文本渲染为 HTML：连续代码行包进 <pre> */
  function formatBody(text) {
    if (!text) return '';
    var lines = String(text).split('\n');
    var html = '', buf = [], inCode = false;
    function flushProse() {
      if (!buf.length) return;
      var t = buf.join('\n').trim();
      if (t) html += '<p class="stem">' + esc(t) + '</p>';
      buf = [];
    }
    function flushCode() {
      if (!buf.length) return;
      html += '<pre class="code">' + esc(buf.join('\n')) + '</pre>';
      buf = [];
    }
    lines.forEach(function (ln) {
      var isCode = CODE_LINE.test(ln) || /^\s{4,}/.test(ln);
      if (isCode !== inCode) {
        inCode ? flushCode() : flushProse();
        inCode = isCode;
      }
      buf.push(ln);
    });
    inCode ? flushCode() : flushProse();
    return html;
  }

  /* ---------- 题车（localStorage） ---------- */
  var CART_KEY = 'apcsa_qbank_cart_v1';
  function getCart() {
    try { return JSON.parse(localStorage.getItem(CART_KEY) || '[]'); } catch (e) { return []; }
  }
  function setCart(arr) {
    try { localStorage.setItem(CART_KEY, JSON.stringify(arr)); } catch (e) { }
    document.dispatchEvent(new CustomEvent('cartchange', { detail: arr }));
  }
  function inCart(id) { return getCart().indexOf(id) >= 0; }
  function toggleCart(id, on) {
    var c = getCart(), i = c.indexOf(id);
    if (on === undefined) on = i < 0;
    if (on && i < 0) c.push(id);
    if (!on && i >= 0) c.splice(i, 1);
    setCart(c);
    return on;
  }
  function clearCart() { setCart([]); }
  function cartQuestions() {
    var byId = {};
    DATA.questions.forEach(function (q) { byId[q.id] = q; });
    return getCart().map(function (id) { return byId[id]; }).filter(Boolean);
  }

  /* ---------- 统计 ---------- */
  function computeStats() {
    var qs = DATA.questions;
    var byUnit = {}, byYear = {}, byTopic = {}, bySource = {}, byType = { MCQ: 0, FRQ: 0 };
    qs.forEach(function (q) {
      byType[q.type] = (byType[q.type] || 0) + 1;
      if (q.unit) byUnit[q.unit] = (byUnit[q.unit] || 0) + 1;
      if (q.year) byYear[q.year] = (byYear[q.year] || 0) + 1;
      if (q.topic) byTopic[q.topic] = (byTopic[q.topic] || 0) + 1;
      if (q.sourceLabel) bySource[q.sourceLabel] = (bySource[q.sourceLabel] || 0) + 1;
    });
    return { byUnit: byUnit, byYear: byYear, byTopic: byTopic, bySource: bySource, byType: byType, total: qs.length };
  }

  /* ---------- 单题渲染 ---------- */
  function renderCard(q, opts) {
    opts = opts || {};
    var showAns = !!opts.showAnswer;
    var ucolor = unitColor(q.unit);
    var tagBits = [];
    if (q.unit) tagBits.push('<span class="chip" style="background:' + ucolor + '18;color:' + ucolor + '">' + esc(unitShort(q.unit)) + '</span>');
    if (q.topic) {
      var tname = TOPIC_INDEX[q.topic] ? TOPIC_INDEX[q.topic].nameCn : '';
      tagBits.push('<span class="chip" title="' + esc(q.topic + ' ' + tname) + '">Topic ' + esc(q.topic) + '</span>');
    }
    if (q.type === 'FRQ') tagBits.push('<span class="chip chip-frq">FRQ</span>');
    if (q.year) tagBits.push('<span class="chip chip-year">' + q.year + '</span>');
    if (q.topicTagSource === 'official') tagBits.push('<span class="chip chip-official" title="考点标注来自官方 Skill / Learning Objective 表">官方标注' + (q.officialLO ? ' · ' + esc(q.officialLO) : '') + '</span>');

    var bodyHtml = '';
    if (q.type === 'FRQ') {
      bodyHtml += formatBody(q.prompt || q.stem || '');
    } else {
      bodyHtml += formatBody(q.stem || '');
      if (q.options && q.options.length) {
        bodyHtml += '<div class="opts">';
        q.options.forEach(function (o) {
          var cls = 'opt' + (showAns && String(q.answer) === String(o.label) ? ' correct' : '');
          bodyHtml += '<div class="' + cls + '"><span class="lb">(' + esc(o.label) + ')</span><span class="tx">' + esc(o.text) + '</span></div>';
        });
        bodyHtml += '</div>';
      }
    }

    var hasAnsContent = q.type === 'FRQ' ? !!q.solution : !!(q.answer || q.explanation);
    var answerHtml = '';
    if (!showAns) {
      answerHtml = hasAnsContent
        ? '<div class="answerbox" style="border-top-style:dashed"><span class="src" style="color:var(--muted)">答案与解析已隐藏 —— 勾选右上角「显示答案」后查看</span></div>'
        : '';
    } else if (q.type === 'FRQ') {
      if (q.solution) {
        answerHtml = '<div class="answerbox wb-ans">' +
          '<div class="expl-label">评分标准 / 参考答案（Scoring Guidelines）</div>' +
          '<div class="expl">' + esc(q.solution) + '</div></div>';
      }
    } else if (hasAnsContent) {
      answerHtml = '<div class="answerbox wb-ans">' +
        (q.answer ? '<div class="ansline">答案：<b>' + esc(q.answer) + '</b></div>' : '<div class="ansline" style="color:var(--warn)">本题答案暂缺（可参考原卷）</div>') +
        (q.explanation ? '<div class="expl-label">解析</div><div class="expl">' + esc(q.explanation) + '</div>' : '') +
        '</div>';
    }

    var footHtml = '<div class="qfoot">' +
      (answerHtml ? '<button class="toggle wb-toggle">查看答案与解析</button>' : '<span class="src">暂无答案 / 解析</span>') +
      '<span class="src">· 来源：' + esc(q.label || q.sourceLabel || '') + (q.id ? ' #' + esc(String(q.id).split('-').pop()) : '') + '</span>' +
      '</div>';

    return '<article class="qcard" data-id="' + esc(q.id) + '"' + (inCart(q.id) ? ' data-selected="1"' : '') + '>' +
      '<div class="qhead">' +
      '<input type="checkbox" class="pick" ' + (inCart(q.id) ? 'checked' : '') + ' title="加入题车">' +
      '<span class="qno">' + esc(String(q.id).split('-').pop()) + '.</span>' +
      '<div class="tags">' + tagBits.join('') + '</div>' +
      '</div>' +
      '<div class="qbody">' + bodyHtml + '</div>' +
      footHtml +
      '</article>';
  }

  /* 绑定题卡交互（事件委托） */
  function bindCards(root, hooks) {
    hooks = hooks || {};
    root.addEventListener('click', function (e) {
      var t = e.target;
      if (t.classList.contains('wb-toggle')) {
        var card = t.closest('.qcard');
        var box = card.querySelector('.wb-ans');
        if (box) {
          var hide = !box.hidden;
          box.hidden = hide;
          t.textContent = hide ? '查看答案与解析' : '收起答案与解析';
        } else if (typeof window.wbAskShowAnswer === 'function') {
          window.wbAskShowAnswer();
        }
        return;
      }
    });
    root.addEventListener('change', function (e) {
      if (e.target.classList.contains('pick')) {
        var card = e.target.closest('.qcard');
        var id = card.getAttribute('data-id');
        var on = toggleCart(id, e.target.checked);
        card.classList.toggle('selected', on);
        if (hooks.onPick) hooks.onPick(id, on);
      }
    });
  }

  /* ---------- 题车浮条 ---------- */
  function renderCartBar() {
    var el = document.getElementById('cartBar');
    if (!el) return;
    var n = getCart().length;
    if (!n) { el.hidden = true; return; }
    el.hidden = false;
    el.innerHTML = '<span>已选 <span class="cnum">' + n + '</span> 题</span>' +
      '<a class="btn btn-light" href="paper.html">进入组卷</a>' +
      '<button class="btn" id="cartClear">清空</button>';
    var b = document.getElementById('cartClear');
    if (b) b.onclick = function () {
      if (confirm('确定清空已选题目的题车吗？')) {
        clearCart();
        document.querySelectorAll('.qcard .pick').forEach(function (c) { c.checked = false; c.closest('.qcard').classList.remove('selected'); });
        if (typeof window.wbAfterCartClear === 'function') window.wbAfterCartClear();
      }
    };
  }

  /* ---------- Toast ---------- */
  function toast(msg) {
    var el = document.getElementById('toast');
    if (!el) { el = document.createElement('div'); el.id = 'toast'; el.className = 'toast'; document.body.appendChild(el); }
    el.textContent = msg;
    el.classList.add('show');
    clearTimeout(el._t);
    el._t = setTimeout(function () { el.classList.remove('show'); }, 2200);
  }

  /* ---------- 导航高亮 ---------- */
  if (typeof document !== 'undefined') {
    document.addEventListener('DOMContentLoaded', function () {
      document.querySelectorAll('.nav a').forEach(function (a) {
        if (a.getAttribute('href') && location.pathname.indexOf(a.getAttribute('href')) >= 0) a.classList.add('on');
      });
      renderCartBar();
    });
    document.addEventListener('cartchange', renderCartBar);
  }

  /* ---------- 导出 ---------- */
  window.QB = {
    DATA: DATA, TOPICS: TOPICS, TOPIC_INDEX: TOPIC_INDEX, UNIT_INDEX: UNIT_INDEX,
    esc: esc, qs: qs, formatBody: formatBody, renderCard: renderCard, bindCards: bindCards,
    getCart: getCart, setCart: setCart, cartQuestions: cartQuestions, clearCart: clearCart,
    toggleCart: toggleCart, inCart: inCart, computeStats: computeStats, toast: toast,
    unitColor: unitColor, unitShort: unitShort, topicName: topicName
  };
})();
