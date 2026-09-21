/* ============================================================
   APCSA 题库 · 组卷与导出
   依赖：lib/docx.min.js（UMD，暴露 window.docx）
   ============================================================ */
(function () {
  'use strict';

  var CODE_LINE = /(;|\{|\}|^\s*\/\/|^\s*\*|^\s*(public|private|protected|static|void|int|double|boolean|char|String|if|for|while|do|return|else|class|new|final)\b)/;
  var IMG_LINE = /^!\[\]\((data:image\/[a-z]+;base64,([A-Za-z0-9+/=]+))\)$/;
  var MONO_FONT = 'Consolas';
  var CN_FONT = 'DengXian';
  var CODEFONT_FALLBACK = 'Courier New';

  function splitStem(text) {
    /* 返回 [{type:'prose'|'code'|'code', text}]（向后兼容）*/
    return splitSegments(text).filter(function (s) { return s.type !== 'table' && s.type !== 'image'; });
  }

  /* 文本 → [{type:'prose'|'code'|'table'|'image', text|b64}] */
  function splitSegments(text) {
    var out = [], prose = [], code = [], table = [];
    function flushProse() {
      if (prose.length) { out.push({ type: 'prose', text: prose.join('\n') }); prose = []; }
    }
    function flushCode() {
      if (code.length) { out.push({ type: 'code', text: code.join('\n') }); code = []; }
    }
    function flushTable() {
      if (table.length) { out.push({ type: 'table', text: table.join('\n') }); table = []; }
    }
    function flushAll() { flushProse(); flushCode(); flushTable(); }
    String(text || '').split('\n').forEach(function (ln) {
      var m = IMG_LINE.exec(ln.trim());
      if (m) { flushAll(); out.push({ type: 'image', b64: m[2] }); return; }
      var s = ln.trim();
      if (s.charAt(0) === '|' && s.charAt(s.length - 1) === '|' && s.length > 2) {
        flushProse(); flushCode(); table.push(s); return;
      }
      if (table.length) flushTable();
      var isCode = CODE_LINE.test(ln) || /^\s{4,}/.test(ln);
      if (isCode) { flushProse(); code.push(ln); }
      else { flushCode(); prose.push(ln); }
    });
    flushAll();
    return out;
  }

  /* Markdown 表格文本 → 行数组（剔除分隔行） */
  function mdTableRows(mdText) {
    return mdText.split('\n').map(function (ln) {
      return ln.replace(/^\s*\|/, '').replace(/\|\s*$/, '').split('|').map(function (c) { return c.trim(); });
    }).filter(function (r) {
      return !r.every(function (c) { return c === '' || /^:?-{2,}:?$/.test(c); });
    });
  }

  /* PNG base64 → {w, h, bin}（读 IHDR） */
  function pngSize(b64) {
    try {
      var bin = atob(b64);
      if (bin.length < 24 || bin.charCodeAt(0) !== 0x89) return null;
      var w = ((bin.charCodeAt(16) << 24) | (bin.charCodeAt(17) << 16) | (bin.charCodeAt(18) << 8) | bin.charCodeAt(19)) >>> 0;
      var h = ((bin.charCodeAt(20) << 24) | (bin.charCodeAt(21) << 16) | (bin.charCodeAt(22) << 8) | bin.charCodeAt(23)) >>> 0;
      return (w > 0 && h > 0) ? { w: w, h: h } : null;
    } catch (e) { return null; }
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

    /* 混合片段（散文/代码/表格/图片）→ docx 元素 */
    function pushSegments(text, codeSize) {
      codeSize = codeSize || 19;
      splitSegments(text).forEach(function (seg) {
        if (seg.type === 'prose') {
          seg.text.split('\n').forEach(function (ln) {
            if (ln.trim()) kids.push(new P({ spacing: { after: 60, line: 300 }, children: [new T({ text: ln, size: 21, font: CN_FONT })] }));
          });
        } else if (seg.type === 'code') {
          seg.text.split('\n').forEach(function (ln) {
            kids.push(new P({
              spacing: { after: 0 }, indent: { left: 280 },
              children: [new T({ text: ln || ' ', font: MONO_FONT, size: codeSize })]
            }));
          });
          kids.push(new P({ spacing: { after: 60 }, children: [new T({ text: '', size: 10 })] }));
        } else if (seg.type === 'table') {
          var rows = mdTableRows(seg.text);
          if (rows.length) {
            kids.push(new D.Table({
              width: { size: 100, type: D.WidthType.PERCENTAGE },
              rows: rows.map(function (r) {
                return new D.TableRow({
                  children: r.map(function (c) {
                    return new D.TableCell({
                      children: [new P({ children: [new T({ text: c, size: 18, font: CN_FONT })] })]
                    });
                  })
                });
              })
            }));
            kids.push(new P({ spacing: { after: 60 }, children: [new T({ text: '', size: 10 })] }));
          }
        } else if (seg.type === 'image') {
          var dim = pngSize(seg.b64);
          if (dim) {
            var scale = Math.min(1, 520 / dim.w);
            kids.push(new P({
              alignment: A.CENTER, spacing: { before: 60, after: 80 },
              children: [new D.ImageRun({
                data: seg.b64,
                transformation: { width: Math.round(dim.w * scale), height: Math.round(dim.h * scale) }
              })]
            }));
          }
        }
      });
    }

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
      pushSegments(bodyText);

      /* 选项 */
      if (q.options && q.options.length) {
        q.options.forEach(function (o) {
          var isAns = String(q.answer) === String(o.label);
          if (o.image) {
            var dim = pngSize(o.image);
            if (dim) {
              var scale = Math.min(1, 460 / dim.w);
              kids.push(new P({
                spacing: { after: 40 }, indent: { left: 280 },
                children: [
                  new T({ text: '(' + o.label + ') ', size: 21, font: CN_FONT, bold: cfg.includeAnswer && isAns }),
                  new D.ImageRun({
                    data: o.image,
                    transformation: { width: Math.round(dim.w * scale), height: Math.round(dim.h * scale) }
                  })
                ]
              }));
            }
          } else {
            kids.push(new P({
              spacing: { after: 40 }, indent: { left: 280 },
              children: [new T({
                text: '(' + o.label + ') ' + o.text.replace(/\n/g, ' '),
                size: 21, font: CN_FONT,
                bold: cfg.includeAnswer && isAns
              })]
            }));
          }
        });
      }

      /* 答案 / 解析 */
      if (cfg.includeAnswer || cfg.includeExpl) {
        if (q.type === 'FRQ') {
          if (cfg.includeAnswer && q.solution) {
            kids.push(new P({ spacing: { before: 100, after: 40 }, children: [new T({ text: '【参考答案程序（Canonical Solution）】', bold: true, size: 20, color: '1D4ED8', font: CN_FONT })] }));
            q.solution.split('\n').forEach(function (ln) {
              kids.push(new P({ spacing: { after: 20, line: 280 }, indent: { left: 280 }, children: [new T({ text: ln, size: 19, font: MONO_FONT })] }));
            });
          }
          if (cfg.includeAnswer && q.rubricTable && q.rubricTable.length) {
            /* 官方评分表格（与 CB Scoring Guidelines 版式一致） */
            q.rubricTable.forEach(function (p) {
              kids.push(new P({ spacing: { before: 120, after: 40 }, children: [new T({
                text: '【官方评分标准】' + (p.part ? ' Part (' + p.part + ') ' : ' ') + (p.name || '') + (p.total ? '　' + p.total : ''),
                bold: true, size: 20, color: '1D4ED8', font: CN_FONT
              })] }));
              var headRow = new D.TableRow({
                tableHeader: true,
                children: ['Scoring Criteria', 'Decision Rules', 'Points'].map(function (h) {
                  return new D.TableCell({
                    shading: { fill: 'CFE6FB' },
                    children: [new P({ children: [new T({ text: h, bold: true, size: 18, font: CN_FONT })] })]
                  });
                })
              });
              var bodyRows = (p.rows || []).map(function (r) {
                var ruleLines = [];
                if (r.can && r.can.length) {
                  ruleLines.push('Responses can still earn the point even if they');
                  r.can.forEach(function (b) { ruleLines.push('• ' + b); });
                }
                if (r.wont && r.wont.length) {
                  ruleLines.push('Responses will not earn the point if they');
                  r.wont.forEach(function (b) { ruleLines.push('• ' + b); });
                }
                (r.notes || []).forEach(function (n) { ruleLines.push(n); });
                if (!ruleLines.length) ruleLines.push('—');
                var strip = function (s) { return String(s || '').replace(/`/g, ''); };
                function cellP(text, bold) {
                  return new P({ children: [new T({ text: text, size: 17, bold: !!bold, font: CN_FONT })] });
                }
                return new D.TableRow({
                  children: [
                    new D.TableCell({ children: [cellP((r.n ? r.n + '  ' : '') + strip(r.criteria))] }),
                    new D.TableCell({ children: ruleLines.map(function (l) { return cellP(l); }) }),
                    new D.TableCell({ children: [cellP(r.points || '', true)] })
                  ]
                });
              });
              kids.push(new D.Table({
                width: { size: 100, type: D.WidthType.PERCENTAGE },
                rows: [headRow].concat(bodyRows)
              }));
              kids.push(new P({ spacing: { after: 60 }, children: [new T({ text: '', size: 10 })] }));
            });
          } else if (cfg.includeAnswer && q.rubric) {
            kids.push(new P({ spacing: { before: 120, after: 40 }, children: [new T({ text: '【官方评分标准（Scoring Guidelines）】', bold: true, size: 20, color: '1D4ED8', font: CN_FONT })] }));
            pushSegments(q.rubric, 18);
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
      if (q.stemImg) html += '<div class="figbox"><img src="' + q.stemImg + '" alt="题干图"></div>';
      if (q.options && q.options.length) {
        html += '<div class="opts">' + q.options.map(function (o) {
          if (o.image) return '<div class="opt"><span class="lb">(' + o.label + ')</span><span class="tx"><img class="opt-img" src="' + o.image + '" alt="选项"></span></div>';
          return '<div class="opt"><span class="lb">(' + o.label + ')</span><span class="tx">' + QB.esc(o.text) + '</span></div>';
        }).join('') + '</div>';
      }
      if (shown) {
        if (q.type === 'FRQ') {
          if (q.solution) html += '<div class="answerbox"><div class="expl-label">参考答案程序（Canonical Solution）</div><div class="expl">' + QB.formatBody(q.solution) + '</div></div>';
          if (q.rubricTable && q.rubricTable.length) html += '<div class="answerbox"><div class="expl-label">官方评分标准（Scoring Guidelines）</div><div class="expl expl-rubric">' + QB.rubricTableHtml(q.rubricTable) + '</div></div>';
          else if (q.rubric) html += '<div class="answerbox"><div class="expl-label">官方评分标准（Scoring Guidelines）</div><div class="expl expl-rubric">' + QB.formatBody(q.rubric) + '</div></div>';
        }
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
      (q.options || []).forEach(function (o) {
        md += '- (' + o.label + ') ' + (o.image ? '（图片选项，见网站）' : o.text.replace(/\n/g, ' ')) + '\n';
      });
      if (cfg.includeAnswer && q.answer) md += '\n> 答案：' + q.answer + '\n';
      if (cfg.includeExpl && q.explanation) md += '> 解析：' + q.explanation.replace(/\n/g, ' ') + '\n';
      if (cfg.includeAnswer && q.type === 'FRQ' && q.solution) md += '\n> 参考答案程序：\n> ```java\n' + q.solution.replace(/^/gm, '> ') + '\n> ```\n';
      if (cfg.includeAnswer && q.type === 'FRQ' && q.rubricTable && q.rubricTable.length) {
        md += '\n> 官方评分标准：\n>\n';
        q.rubricTable.forEach(function (p) {
          md += '> **' + (p.part ? '(' + p.part + ') ' : '') + (p.name || '') + (p.total ? '　' + p.total : '') + '**\n>\n';
          md += '> | # | Scoring Criteria | Decision Rules | Points |\n> | --- | --- | --- | --- |\n';
          (p.rows || []).forEach(function (r) {
            var rules = [];
            if (r.can && r.can.length) { rules.push('Responses can still earn the point even if they'); r.can.forEach(function (b) { rules.push('• ' + b); }); }
            if (r.wont && r.wont.length) { rules.push('Responses will not earn the point if they'); r.wont.forEach(function (b) { rules.push('• ' + b); }); }
            (r.notes || []).forEach(function (n) { rules.push(n); });
            md += '> | ' + (r.n || '') + ' | ' + String(r.criteria || '').replace(/\|/g, '\\|').replace(/`/g, '') +
              ' | ' + rules.join('<br>').replace(/\|/g, '\\|').replace(/`/g, '') + ' | ' + (r.points || '') + ' |\n';
          });
          md += '>\n';
        });
      } else if (cfg.includeAnswer && q.type === 'FRQ' && q.rubric) {
        md += '\n> 官方评分标准：\n' + q.rubric.replace(/^/gm, '> ') + '\n';
      }
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
