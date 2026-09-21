/* ============================================================
   APCSA 题库 · 组卷与导出
   依赖：lib/docx.min.js（UMD，暴露 window.docx）
   ============================================================ */
(function () {
  'use strict';

  var CODE_LINE = /(;|\{|\}|^\s*\/\/|^\s*\*|^\s*(public|private|protected|static|void|int|double|boolean|char|String|if|for|while|do|return|else|class|new|final)\b)/;
  var MONO_FONT = 'Consolas';
  var CN_FONT = 'DengXian';
  var CODEFONT_FALLBACK = 'Courier New';

  function splitStem(text) {
    /* 返回 [{type:'prose'|'code', text}] */
    var out = [], buf = [], inCode = false;
    String(text || '').split('\n').forEach(function (ln) {
      var isCode = CODE_LINE.test(ln) || /^\s{4,}/.test(ln);
      if (isCode !== inCode) {
        if (buf.length) out.push({ type: inCode ? 'code' : 'prose', text: buf.join('\n') });
        buf = []; inCode = isCode;
      }
      buf.push(ln);
    });
    if (buf.length) out.push({ type: inCode ? 'code' : 'prose', text: buf.join('\n') });
    return out;
  }

  function cfgFromUI() {
    return {
      title: document.getElementById('pTitle').value.trim() || 'AP Computer Science A 练习题',
      subtitle: document.getElementById('pSub').value.trim(),
      headInfo: document.getElementById('pHead').value.trim(),
      includeSource: document.getElementById('optSource').checked,
      includeAnswer: document.getElementById('optAnswer').checked,
      includeExpl: document.getElementById('optExpl').checked,
      blankLines: Math.max(0, parseInt(document.getElementById('optBlank').value, 10) || 0),
      pageBreak: document.getElementById('optBreak').checked,
      showTopic: document.getElementById('optTopic').checked
    };
  }

  /* ---------- 构建 docx ---------- */
  function buildDocx(questions, cfg) {
    var D = window.docx;
    if (!D) throw new Error('Word 导出库未加载，请检查 lib/docx.min.js');
    var P = D.Paragraph, T = D.TextRun, A = D.AlignmentType;
    var kids = [];
    var TOTAL = questions.length;

    function para(text, o) {
      return new P(Object.assign({ children: [new T({ text: text, font: CN_FONT, size: 21 })] }, o || {}));
    }
    function blank(n) { for (var i = 0; i < n; i++) kids.push(new P({ children: [new T({ text: '', size: 21 })] })); }

    /* 卷头 */
    kids.push(new P({ alignment: A.CENTER, spacing: { after: 60 }, children: [new T({ text: cfg.title, bold: true, size: 32, font: CN_FONT })] }));
    if (cfg.subtitle) kids.push(new P({ alignment: A.CENTER, spacing: { after: 60 }, children: [new T({ text: cfg.subtitle, size: 21, color: '555555', font: CN_FONT })] }));
    kids.push(new P({
      alignment: A.CENTER, spacing: { after: 120 },
      children: [new T({ text: '共 ' + TOTAL + ' 题' + (cfg.headInfo ? '　' + cfg.headInfo : ''), size: 20, color: '666666', font: CN_FONT })]
    }));
    kids.push(new P({
      spacing: { after: 240 },
      children: [new T({
        text: '班级：____________　　姓名：____________　　学号：____________　　成绩：__________',
        size: 20, font: CN_FONT
      })],
      border: { bottom: { style: D.BorderStyle.SINGLE, size: 6, color: 'BFBFBF', space: 6 } }
    }));

    if (cfg.includeAnswer === false && shouldShowMark(cfg)) { /* noop */ }

    questions.forEach(function (q, idx) {
      if (idx > 0 && cfg.pageBreak) kids.push(new P({ children: [new D.PageBreak()] }));

      /* 题号行 + 标签 */
      var headRuns = [new T({ text: (idx + 1) + '. ', bold: true, size: 22, font: CN_FONT })];
      var metaBits = [];
      if (cfg.showTopic) {
        if (q.topic) metaBits.push('Topic ' + q.topic);
        if (q.type === 'FRQ') metaBits.push('FRQ');
      }
      if (cfg.includeSource) metaBits.push(q.label || q.sourceLabel || '');
      if (metaBits.length) headRuns.push(new T({ text: '［' + metaBits.join(' · ') + '］', size: 17, color: '888888', font: CN_FONT }));
      kids.push(new P({ spacing: { before: 120, after: 60 }, children: headRuns, keepNext: true }));

      /* 题干 / 提示 */
      var bodyText = q.type === 'FRQ' ? (q.prompt || q.stem || '') : (q.stem || '');
      splitStem(bodyText).forEach(function (seg) {
        if (seg.type === 'code') {
          seg.text.split('\n').forEach(function (ln) {
            kids.push(new P({
              spacing: { after: 0 }, indent: { left: 280 },
              children: [new T({ text: ln || ' ', font: MONO_FONT, size: 19 })]
            }));
          });
          kids.push(new P({ spacing: { after: 60 }, children: [new T({ text: '', size: 10 })] }));
        } else {
          seg.text.split('\n').forEach(function (ln) {
            kids.push(new P({ spacing: { after: 60, line: 300 }, children: [new T({ text: ln, size: 21, font: CN_FONT })] }));
          });
        }
      });

      /* 选项 */
      if (q.options && q.options.length) {
        q.options.forEach(function (o) {
          var isAns = String(q.answer) === String(o.label);
          kids.push(new P({
            spacing: { after: 40 }, indent: { left: 280 },
            children: [new T({
              text: '(' + o.label + ') ' + o.text.replace(/\n/g, ' '),
              size: 21, font: CN_FONT,
              bold: cfg.includeAnswer && isAns
            })]
          }));
        });
      }

      /* 答案 / 解析 */
      if (cfg.includeAnswer || cfg.includeExpl) {
        if (q.type === 'FRQ') {
          if (cfg.includeAnswer && q.solution) {
            kids.push(new P({ spacing: { before: 100, after: 40 }, children: [new T({ text: '【评分标准 / 参考答案】', bold: true, size: 20, color: '1D4ED8', font: CN_FONT })] }));
            q.solution.split('\n').forEach(function (ln) {
              kids.push(new P({ spacing: { after: 20, line: 280 }, children: [new T({ text: ln, size: 19, font: CN_FONT })] }));
            });
          }
        } else {
          if (cfg.includeAnswer) {
            kids.push(new P({
              spacing: { before: 100, after: 40 },
              children: [new T({ text: '答案：' + (q.answer || '（略）'), bold: true, size: 21, color: '059669', font: CN_FONT })]
            }));
          }
          if (cfg.includeExpl && q.explanation) {
            kids.push(new P({ spacing: { after: 20 }, children: [new T({ text: '解析：', bold: true, size: 20, color: '1D4ED8', font: CN_FONT })] }));
            q.explanation.split('\n').forEach(function (ln) {
              kids.push(new P({ spacing: { after: 20, line: 280 }, children: [new T({ text: ln, size: 19, font: CN_FONT })] }));
            });
          }
        }
      } else {
        blank(cfg.blankLines);
      }
    });

    return new D.Document({
      creator: 'APCSA 题库',
      title: cfg.title,
      styles: { default: { document: { run: { font: CN_FONT, size: 21 } } } },
      sections: [{
        properties: { page: { margin: { top: 1000, right: 1000, bottom: 1000, left: 1100 } } },
        children: kids
      }]
    });
  }

  function shouldShowMark() { return false; }

  function downloadBlob(blob, filename) {
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url; a.download = filename;
    document.body.appendChild(a); a.click();
    setTimeout(function () { document.body.removeChild(a); URL.revokeObjectURL(url); }, 400);
  }

  function safeName(s) { return String(s).replace(/[\\\/:*?"<>|\s]+/g, '_').slice(0, 60); }

  /* 暴露给测试 / 高级用法 */
  if (typeof window !== 'undefined') {
    window.QB_DOCX = { buildDocx: buildDocx, splitStem: splitStem, downloadBlob: downloadBlob, safeName: safeName };
  }

  /* ---------- 页面逻辑（仅在组卷页执行） ---------- */
  if (typeof document === 'undefined' || !document.getElementById('paperList')) return;

  /* ---------- 页面逻辑 ---------- */
  var listEl = document.getElementById('paperList');
  var emptyEl = document.getElementById('paperEmpty');

  function render() {
    var qs = QB.cartQuestions();
    var n = qs.length;
    document.getElementById('pCount').textContent = n;
    emptyEl.hidden = n > 0;
    listEl.hidden = n === 0;
    document.getElementById('actBar').hidden = n === 0;

    var byUnit = {};
    qs.forEach(function (q) { byUnit[q.unit || '—'] = (byUnit[q.unit || '—'] || 0) + 1; });
    document.getElementById('pSummary').textContent = n
      ? Object.keys(byUnit).sort().map(function (u) { return QB.unitShort(u) + ' ' + byUnit[u] + '题'; }).join('　·　')
      : '';

    listEl.innerHTML = qs.map(function (q, i) {
      var stem = (q.type === 'FRQ' ? (q.prompt || q.stem) : q.stem) || '';
      stem = stem.replace(/\n/g, ' ').slice(0, 150);
      var uc = QB.unitColor(q.unit);
      return '<div class="paper-item" data-id="' + QB.esc(q.id) + '">' +
        '<div class="ord">' + (i + 1) + '</div>' +
        '<div class="main"><div class="t">' + QB.esc(stem) + (stem.length >= 150 ? '…' : '') + '</div>' +
        '<div class="m">' +
        (q.unit ? '<span class="chip" style="background:' + uc + '18;color:' + uc + '">' + QB.unitShort(q.unit) + '</span>' : '') +
        (q.topic ? '<span class="chip">Topic ' + QB.esc(q.topic) + '</span>' : '') +
        '<span class="chip">' + (q.type === 'FRQ' ? 'FRQ' : 'MCQ') + '</span>' +
        (q.year ? '<span class="chip chip-year">' + q.year + '</span>' : '') +
        '<span>' + QB.esc(q.label || '') + '</span>' +
        '</div></div>' +
        '<div class="acts">' +
        '<button class="iconbtn" data-act="up" title="上移">↑</button>' +
        '<button class="iconbtn" data-act="down" title="下移">↓</button>' +
        '<button class="iconbtn danger" data-act="del" title="移除">✕</button>' +
        '</div></div>';
    }).join('');
  }

  listEl.addEventListener('click', function (e) {
    var b = e.target.closest('[data-act]');
    if (!b) return;
    var id = e.target.closest('.paper-item').getAttribute('data-id');
    var cart = QB.getCart();
    var i = cart.indexOf(id);
    if (i < 0) return;
    var act = b.getAttribute('data-act');
    if (act === 'del') cart.splice(i, 1);
    if (act === 'up' && i > 0) { cart[i] = cart[i - 1]; cart[i - 1] = id; }
    if (act === 'down' && i < cart.length - 1) { cart[i] = cart[i + 1]; cart[i + 1] = id; }
    QB.setCart(cart);
    render();
  });

  document.addEventListener('cartchange', render);

  document.getElementById('btnDocx').onclick = function () {
    var qs = QB.cartQuestions();
    if (!qs.length) return QB.toast('题车为空，请先到题库勾选题目');
    var cfg = cfgFromUI();
    var btn = this; btn.disabled = true; btn.textContent = '正在生成…';
    try {
      var doc = buildDocx(qs, cfg);
      window.docx.Packer.toBlob(doc).then(function (blob) {
        downloadBlob(blob, safeName(cfg.title) + '_' + qs.length + '题.docx');
        QB.toast('已导出 Word 试卷（' + qs.length + ' 题）');
      }).catch(function (err) {
        console.error(err); QB.toast('生成失败：' + err.message);
      }).then(function () { btn.disabled = false; btn.textContent = '导出 Word (.docx)'; });
    } catch (e) {
      console.error(e); QB.toast('生成失败：' + e.message);
      btn.disabled = false; btn.textContent = '导出 Word (.docx)';
    }
  };

  document.getElementById('btnPrint').onclick = function () {
    var qs = QB.cartQuestions();
    if (!qs.length) return QB.toast('题车为空');
    var cfg = cfgFromUI();
    var shown = cfg.includeAnswer;
    var html = '<div id="printRoot"><h1 style="text-align:center;font-size:22pt;margin:0 0 6px">' + QB.esc(cfg.title) + '</h1>';
    if (cfg.subtitle) html += '<p style="text-align:center;color:#555;margin:0 0 4px">' + QB.esc(cfg.subtitle) + '</p>';
    html += '<p style="text-align:center;color:#666;font-size:10pt;margin:0 0 4px">共 ' + qs.length + ' 题' +
      (cfg.headInfo ? '　' + QB.esc(cfg.headInfo) : '') + '</p>';
    html += '<p style="border-bottom:1px solid #bbb;padding-bottom:8px;font-size:10pt">班级：__________　姓名：__________　学号：__________　成绩：________</p>';
    qs.forEach(function (q, i) {
      html += '<div class="qcard" style="padding:0 0 10pt;margin-bottom:10pt;border-bottom:1px solid #eee">';
      html += '<div class="qhead"><span class="qno">' + (i + 1) + '.</span><div class="tags">' +
        (cfg.showTopic && q.topic ? '<span class="chip">Topic ' + q.topic + '</span>' : '') +
        (cfg.includeSource ? '<span class="chip">' + QB.esc(q.label || '') + '</span>' : '') + '</div></div>';
      html += '<div class="qbody">' + QB.formatBody(q.type === 'FRQ' ? (q.prompt || q.stem) : q.stem);
      if (q.options && q.options.length) {
        html += '<div class="opts">' + q.options.map(function (o) { return '<div class="opt"><span class="lb">(' + o.label + ')</span><span class="tx">' + QB.esc(o.text) + '</span></div>'; }).join('') + '</div>';
      }
      if (shown) {
        if (q.type === 'FRQ' && q.solution) html += '<div class="answerbox"><div class="expl-label">评分标准 / 参考答案</div><div class="expl">' + QB.esc(q.solution) + '</div></div>';
        else if (q.type === 'MCQ') html += '<div class="answerbox">' + (q.answer ? '<div class="ansline">答案：<b>' + q.answer + '</b></div>' : '') + (cfg.includeExpl && q.explanation ? '<div class="expl-label">解析</div><div class="expl">' + QB.esc(q.explanation) + '</div>' : '') + '</div>';
      } else {
        for (var k = 0; k < cfg.blankLines; k++) html += '<div style="height:22pt"></div>';
      }
      html += '</div></div>';
    });
    html += '</div>';
    var root = document.getElementById('printRoot');
    if (root) root.remove();
    var div = document.createElement('div');
    div.id = 'printRoot';
    div.innerHTML = html;
    document.body.appendChild(div);
    document.body.classList.add('printing');
    setTimeout(function () { window.print(); }, 120);
  };

  document.getElementById('btnCopyMd').onclick = function () {
    var qs = QB.cartQuestions();
    if (!qs.length) return QB.toast('题车为空');
    var cfg = cfgFromUI();
    var md = '# ' + cfg.title + '\n\n';
    qs.forEach(function (q, i) {
      md += '**' + (i + 1) + '.** ' + ((q.type === 'FRQ' ? (q.prompt || q.stem) : q.stem) || '').replace(/\n/g, '  \n') + '\n\n';
      (q.options || []).forEach(function (o) { md += '- (' + o.label + ') ' + o.text.replace(/\n/g, ' ') + '\n'; });
      if (cfg.includeAnswer && q.answer) md += '\n> 答案：' + q.answer + '\n';
      if (cfg.includeExpl && q.explanation) md += '> 解析：' + q.explanation.replace(/\n/g, ' ') + '\n';
      md += '\n';
    });
    navigator.clipboard.writeText(md).then(function () { QB.toast('已复制 Markdown 到剪贴板'); },
      function () { QB.toast('复制失败，请检查浏览器权限'); });
  };

  document.getElementById('btnClearCart').onclick = function () {
    if (!QB.getCart().length) return;
    if (confirm('确定清空题车吗？')) { QB.clearCart(); render(); }
  };

  render();
})();
