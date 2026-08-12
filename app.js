/* 重量波动折线图分析 —— 纯前端离线版 */
(function () {
  'use strict';

  // ---------- 全局状态 ----------
  var state = {
    workbook: null,
    rows: [],            // 当前工作表原始二维数组（含/不含表头行）
    hasHeader: false,
    timeCol: 0,
    weightCol: 1,
    windowSec: 10,
    deltaKg: 0.01,
    highlightStable: true,
    followZoom: true,
    parsed: [],          // [{t:Date, w:Number}]
    displayData: [],     // [[Number(ms), Number]] 用于 ECharts 绘制
    isEvent: [],         // Boolean[]
    eventType: [],       // 'load'|'unload'|null
    stableFlags: [],     // Boolean[] true=平稳(噪声)
    stableRanges: [],    // [[startMs,endMs], ...]
    events: [],          // [{type,t,from,to,mag}] 聚合后的事件段
    stats: null,
    colCount: 0
  };

  var chart = null;

  // ---------- 工具函数 ----------
  function $(id) { return document.getElementById(id); }

  function colLetter(c) {
    var s = '';
    c = c + 1;
    while (c > 0) {
      var m = (c - 1) % 26;
      s = String.fromCharCode(65 + m) + s;
      c = Math.floor((c - 1) / 26);
    }
    return s;
  }

  function pad(n) { return n < 10 ? '0' + n : '' + n; }

  function formatTime(d) {
    return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()) +
      ' ' + pad(d.getHours()) + ':' + pad(d.getMinutes()) + ':' + pad(d.getSeconds());
  }

  // Excel 时间序列化 / 字符串 / Date 统一解析为 Date
  function parseTime(v) {
    if (v == null || v === '') return null;
    if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
    if (typeof v === 'number') {
      // Excel 序列号（含日期+时间的小数），1900 日期系统
      if (v > 0 && v < 100000) return new Date((v - 25569) * 86400000);
      var d = new Date(v);
      return isNaN(d.getTime()) ? null : d;
    }
    if (typeof v === 'string') {
      var s = v.trim();
      var d2 = new Date(s);
      if (!isNaN(d2.getTime())) return d2;
      var m = /^(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?$/.exec(s);
      if (m) {
        var base = new Date();
        base.setHours(0, 0, 0, 0);
        base.setHours(+m[1], +m[2], m[3] ? +m[3] : 0, 0);
        return base;
      }
      return null;
    }
    return null;
  }

  function parseWeight(v) {
    if (v == null) return NaN;
    if (typeof v === 'number') return v;
    var s = String(v).trim().replace(/,/g, '');
    if (s === '') return NaN;
    return parseFloat(s);
  }

  function colLabel(c) {
    if (state.hasHeader && state.rows[0] && state.rows[0][c] != null && String(state.rows[0][c]).trim() !== '') {
      return colLetter(c) + ' · ' + String(state.rows[0][c]).trim();
    }
    return colLetter(c);
  }

  function downloadFile(content, filename, mime) {
    var blob = new Blob([content], { type: mime + ';charset=utf-8' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url; a.download = filename;
    document.body.appendChild(a); a.click();
    setTimeout(function () { URL.revokeObjectURL(url); document.body.removeChild(a); }, 1000);
  }

  // ---------- 数据解析 ----------
  function handleFile(file) {
    var reader = new FileReader();
    reader.onload = function (e) {
      try {
        var wb = XLSX.read(e.target.result, { type: 'array', cellDates: true });
        state.workbook = wb;
        var names = wb.SheetNames;
        if (names.length > 1) {
          var sel = $('sheet');
          sel.innerHTML = '';
          names.forEach(function (n) {
            var o = document.createElement('option');
            o.value = n; o.textContent = n;
            sel.appendChild(o);
          });
          $('sheetField').style.display = '';
        } else {
          $('sheetField').style.display = 'none';
        }
        loadSheet(names[0]);
      } catch (err) {
        alert('读取文件失败：' + err.message);
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function loadSheet(name) {
    if (!state.workbook) return;
    var ws = state.workbook.Sheets[name];
    // raw:true + cellDates:true -> 日期是 Date，数字保留完整精度
    var rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true, blankrows: false });
    while (rows.length && rows[rows.length - 1].every(function (c) { return c == null || c === ''; })) rows.pop();
    state.rows = rows;

    var cc = 0;
    rows.forEach(function (r) { if (r && r.length > cc) cc = r.length; });
    state.colCount = cc;

    if (rows.length) {
      var t0 = parseTime(rows[0][0]);
      var w0 = parseWeight(rows[0][1]);
      var guessHeader = (t0 == null && isNaN(w0)) && cc >= 2;
      $('hasHeader').checked = guessHeader;
      state.hasHeader = guessHeader;
    }

    buildColumnSelectors();
    state.timeCol = 0;
    state.weightCol = Math.min(1, cc - 1);
    $('timeCol').value = String(state.timeCol);
    $('weightCol').value = String(state.weightCol);

    updateAll();
  }

  function buildColumnSelectors() {
    var tc = $('timeCol'), wc = $('weightCol');
    tc.innerHTML = ''; wc.innerHTML = '';
    for (var c = 0; c < state.colCount; c++) {
      var o1 = document.createElement('option'); o1.value = String(c); o1.textContent = colLabel(c);
      var o2 = document.createElement('option'); o2.value = String(c); o2.textContent = colLabel(c);
      tc.appendChild(o1); wc.appendChild(o2);
    }
  }

  function parseData() {
    var rows = state.rows;
    var tCol = state.timeCol, wCol = state.weightCol;
    var start = state.hasHeader ? 1 : 0;
    var out = [];
    for (var i = start; i < rows.length; i++) {
      var r = rows[i];
      if (!r) continue;
      var t = parseTime(r[tCol]);
      var w = parseWeight(r[wCol]);
      if (t == null || isNaN(w)) continue;
      out.push({ t: t, w: w });
    }
    out.sort(function (a, b) { return a.t.getTime() - b.t.getTime(); });
    state.parsed = out;
  }

  // 窗口波动(range-window)检测 + 事件段聚合
  function analyzeData(data) {
    var T = state.windowSec * 1000;
    var D = state.deltaKg;
    var n = data.length;

    var disp = new Array(n);
    var isEvent = new Array(n).fill(false);
    var eventType = new Array(n).fill(null);
    var stable = new Array(n).fill(true);

    var ranges = [];
    var segStart = null;
    var runs = [];
    var run = null;

    function openRun(type, t, w, winMin, winMax) {
      run = { type: type, startT: t, endT: t, minW: winMin, maxW: winMax };
    }
    function extendRun(t, w) {
      run.endT = t;
      if (w < run.minW) run.minW = w;
      if (w > run.maxW) run.maxW = w;
    }
    function closeRun() {
      if (!run) return;
      var from = run.type === 'load' ? run.minW : run.maxW;
      var to = run.type === 'load' ? run.maxW : run.minW;
      runs.push({
        type: run.type, t: run.startT,
        from: from, to: to,
        mag: run.maxW - run.minW
      });
      run = null;
    }

    var j = 0;
    for (var i = 0; i < n; i++) {
      var ti = data[i].t.getTime();
      var wi = data[i].w;
      while (ti - data[j].t.getTime() > T) j++;

      var minW = wi, maxW = wi;
      for (var k = j; k <= i; k++) {
        if (data[k].w < minW) minW = data[k].w;
        if (data[k].w > maxW) maxW = data[k].w;
      }
      var range = maxW - minW;

      if (range >= D) {
        var diff = wi - data[j].w;
        var type = diff >= 0 ? 'load' : 'unload';
        isEvent[i] = true;
        eventType[i] = type;
        stable[i] = false;
        disp[i] = [ti, wi];

        if (!run || run.type !== type) { closeRun(); openRun(type, ti, wi, minW, maxW); }
        else { extendRun(ti, wi); }

        if (segStart !== null) { ranges.push([segStart, ti]); segStart = null; }
      } else {
        stable[i] = true;
        disp[i] = [ti, wi];
        closeRun();
        if (segStart === null) segStart = ti;
      }
    }
    closeRun();
    if (segStart !== null && n) ranges.push([segStart, data[n - 1].t.getTime()]);

    return {
      isEvent: isEvent,
      eventType: eventType,
      stableFlags: stable,
      stableRanges: ranges,
      events: runs,
      displayData: disp
    };
  }

  function calcStats(data, events) {
    var n = data.length;
    var loads = events.filter(function (e) { return e.type === 'load'; });
    var unloads = events.filter(function (e) { return e.type === 'unload'; });
    return {
      count: events.length,
      loads: loads.length,
      unloads: unloads.length,
      net: n ? data[n - 1].w - data[0].w : 0,
      max: n ? Math.max.apply(null, data.map(function (d) { return d.w; })) : 0,
      min: n ? Math.min.apply(null, data.map(function (d) { return d.w; })) : 0,
      maxMag: events.length ? Math.max.apply(null, events.map(function (e) { return e.mag; })) : 0,
      points: n
    };
  }

  function compute() {
    var res = analyzeData(state.parsed);
    state.displayData = res.displayData;
    state.isEvent = res.isEvent;
    state.eventType = res.eventType;
    state.stableFlags = res.stableFlags;
    state.stableRanges = res.stableRanges;
    state.events = res.events;
    state.stats = calcStats(state.parsed, state.events);
  }

  // ---------- 渲染 ----------
  function initChart() {
    chart = echarts.init($('chart'));
    window.addEventListener('resize', function () { if (chart) chart.resize(); });
  }

  function renderChart(keepZoom) {
    if (!chart) return;
    var savedZoom = null;
    if (keepZoom) {
      var currentOpt = chart.getOption();
      if (currentOpt.dataZoom && currentOpt.dataZoom.length) {
        savedZoom = { start: currentOpt.dataZoom[0].start, end: currentOpt.dataZoom[0].end };
      }
    }
    var data = state.displayData;
    var series = {
      name: '重量',
      type: 'line',
      showSymbol: false,
      smooth: false,
      large: true,
      largeThreshold: 2000,
      animation: false,
      data: data,
      lineStyle: { width: 1.4, color: '#2f6fed' },
      itemStyle: { color: '#2f6fed' },
      areaStyle: { color: 'rgba(47,111,237,0.05)' }
    };

    if (state.highlightStable && state.stableRanges.length) {
      series.markArea = {
        silent: true,
        itemStyle: { color: 'rgba(154,163,178,0.06)' },
        data: state.stableRanges.map(function (r) { return [{ xAxis: r[0] }, { xAxis: r[1] }]; })
      };
    }

    if (state.events.length && state.events.length <= 300) {
      series.markPoint = {
        symbolSize: 9,
        data: state.events.map(function (e) {
          return {
            coord: [e.t, e.type === 'load' ? e.to : e.from],
            name: e.type === 'load' ? '上料' : '出料',
            symbol: e.type === 'load' ? 'triangle' : 'pin',
            itemStyle: { color: e.type === 'load' ? '#e0533d' : '#2f9e6f' }
          };
        }),
        label: { show: false }
      };
    }

    var opt = {
      backgroundColor: '#fff',
      grid: { left: 60, right: 24, top: 34, bottom: 70 },
      tooltip: {
        trigger: 'axis',
        axisPointer: { type: 'line', lineStyle: { color: '#9aa3b2' } },
        backgroundColor: 'rgba(255,255,255,0.96)',
        borderColor: '#e4e8f0',
        textStyle: { color: '#1f2733' },
        formatter: function (params) {
          var p = params[0];
          if (!p) return '';
          var t = new Date(p.value[0]);
          var w = p.value[1];
          var extra = '';
          var idx = p.dataIndex;
          if (idx != null && state.isEvent[idx]) {
            var evType = state.eventType[idx];
            extra = '<br/>' + (evType === 'load' ? '上料 ▲' : '出料 ▼');
          }
          return formatTime(t) + '<br/>重量：<b>' + w.toFixed(3) + ' kg</b>' + extra;
        }
      },
      toolbox: {
        right: 12, top: 0,
        feature: {
          dataZoom: { yAxisIndex: 'none', title: { zoom: '框选放大', back: '还原' } },
          restore: { title: '还原缩放' },
          saveAsImage: { title: '保存图片', name: 'weight_chart' }
        }
      },
      xAxis: {
        type: 'time',
        name: '时间',
        nameLocation: 'middle',
        nameGap: 34,
        axisLine: { lineStyle: { color: '#c7cedb' } },
        axisLabel: { color: '#6b7686' }
      },
      yAxis: {
        type: 'value',
        name: '重量 (kg)',
        scale: true,
        nameTextStyle: { color: '#6b7686' },
        axisLabel: { color: '#6b7686', formatter: function (v) { return v.toFixed(2); } },
        splitLine: { lineStyle: { color: '#eef1f6' } }
      },
      dataZoom: [
        { type: 'inside', filterMode: 'none', start: savedZoom ? savedZoom.start : 0, end: savedZoom ? savedZoom.end : 100 },
        { type: 'slider', filterMode: 'none', height: 22, bottom: 18, start: savedZoom ? savedZoom.start : 0, end: savedZoom ? savedZoom.end : 100 }
      ],
      series: [series]
    };
    chart.setOption(opt, true);
  }

  function renderStats(range) {
    var box = $('stats');
    var s;
    if (range && state.parsed.length) {
      var viewData = [];
      for (var i = 0; i < state.parsed.length; i++) {
        var tms = state.parsed[i].t.getTime();
        if (tms >= range[0] && tms <= range[1]) viewData.push(state.parsed[i]);
        else if (tms > range[1]) break;
      }
      var res = analyzeData(viewData);
      s = calcStats(viewData, res.events);
    } else {
      s = state.stats;
    }
    if (!s || !s.points) { box.innerHTML = '<div class="empty" style="padding:18px;">无有效数据</div>'; return; }
    function stat(k, v, cls) {
      return '<div class="stat"><div class="k">' + k + '</div><div class="v ' + (cls || '') + '">' + v + '</div></div>';
    }
    var suffix = range ? '（视图内）' : '';
    var html =
      stat('有效点数' + suffix, s.points) +
      stat('上料/出料事件' + suffix, s.loads + ' / ' + s.unloads) +
      stat('净重量变化' + suffix, (s.net >= 0 ? '+' : '') + s.net.toFixed(3) + ' kg', s.net >= 0 ? 'load' : 'unload') +
      stat('单次最大波动' + suffix, s.maxMag.toFixed(3) + ' kg') +
      stat('最大重量' + suffix, s.max.toFixed(3) + ' kg') +
      stat('最小重量' + suffix, s.min.toFixed(3) + ' kg');
    if (s.points && s.count === 0) {
      html += '<div style="padding:10px 12px; margin-top:8px; background:#fff8e6; border:1px solid #ffe58f; border-radius:8px; color:#8a6d1b; font-size:12px;">当前阈值下未识别到上料/出料事件，所有点被判为平稳段（抖动&lt;阈值）。如需看到事件，请尝试<b>降低重量差阈值</b>或<b>缩小时间窗口</b>。</div>';
    }
    box.innerHTML = html;
  }

  // 取图表当前 dataZoom 显示范围（毫秒），无缩放返回 null
  function getCurrentRange() {
    if (!chart || !state.parsed.length) return null;
    var opt = chart.getOption();
    var dz = (opt.dataZoom && opt.dataZoom[0]) || {};
    var n = state.parsed.length;
    var t0 = state.parsed[0].t.getTime();
    var t1 = state.parsed[n - 1].t.getTime();
    var s, e;
    if (dz.startValue != null && dz.endValue != null) {
      s = +dz.startValue; e = +dz.endValue;
    } else {
      var st = (dz.start == null ? 0 : +dz.start);
      var en = (dz.end == null ? 100 : +dz.end);
      s = t0 + (t1 - t0) * st / 100;
      e = t0 + (t1 - t0) * en / 100;
    }
    if (s > e) { var tmp = s; s = e; e = tmp; }
    return [s, e];
  }

  function renderPreview(range) {
    var box = $('preview');
    var rows = state.parsed;
    if (!rows.length) { box.innerHTML = '<div class="empty" style="padding:24px;">无有效数据</div>'; return; }

    // 按当前视图范围筛选
    var viewData = [];
    var viewIdx = [];
    if (range) {
      for (var i = 0; i < rows.length; i++) {
        var tms = rows[i].t.getTime();
        if (tms >= range[0] && tms <= range[1]) { viewData.push(rows[i]); viewIdx.push(i); }
        else if (tms > range[1]) break;
      }
    } else {
      for (var j = 0; j < rows.length; j++) { viewData.push(rows[j]); viewIdx.push(j); }
    }

    if (!viewData.length) {
      box.innerHTML = '<div class="empty" style="padding:24px;">当前视图范围内无数据</div>';
      return;
    }

    // 在当前视图内重新计算事件/平稳标记，保证与统计面板一致
    var viewRes = analyzeData(viewData);

    var limit = Math.min(viewData.length, 300);
    var html = '<table><thead><tr><th class="l">#</th><th class="l">时间</th><th>重量(kg)</th><th class="l">标记</th></tr></thead><tbody>';
    for (var k = 0; k < limit; k++) {
      var idx = viewIdx[k];
      var role, cls;
      if (viewRes.isEvent[k]) { role = viewRes.eventType[k] === 'load' ? '上料' : '出料'; cls = viewRes.eventType[k]; }
      else if (viewRes.stableFlags[k]) { role = '平稳'; cls = 'stable'; }
      else { role = '起点'; cls = 'start'; }
      html += '<tr><td class="l">' + (idx + 1) + '</td><td class="l">' + formatTime(viewData[k].t) +
        '</td><td>' + viewData[k].w.toFixed(3) + '</td><td class="l"><span class="pill ' + cls + '">' + role + '</span></td></tr>';
    }
    html += '</tbody></table>';
    if (range) {
      html += '<div class="hint" style="padding:6px 10px;">（显示当前放大区域前 ' + limit + ' 行，共 ' + viewData.length + ' 行）</div>';
    } else {
      html += '<div class="hint" style="padding:6px 10px;">（仅显示前 ' + limit + ' 行，共 ' + rows.length + ' 行）</div>';
    }
    box.innerHTML = html;
  }

  function updateAll(keepZoom) {
    parseData();
    compute();
    renderChart(keepZoom);
    var range = state.followZoom ? getCurrentRange() : null;
    renderStats(range);
    renderPreview(range);
  }

  // ---------- 导出 ----------
  function exportPng() {
    if (!chart) return;
    var url = chart.getDataURL({ type: 'png', pixelRatio: 2, backgroundColor: '#fff' });
    var a = document.createElement('a');
    a.href = url; a.download = 'weight_chart.png';
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
  }

  function exportCsv() {
    var rows = state.parsed;
    if (!rows.length) return;
    var out = [['时间', '重量(kg)', '类型']];
    for (var i = 0; i < rows.length; i++) {
      var role;
      if (state.isEvent[i]) role = state.eventType[i] === 'load' ? '上料' : '出料';
      else if (state.stableFlags[i]) role = '平稳';
      else role = '起点';
      out.push([formatTime(rows[i].t), rows[i].w.toFixed(4), role]);
    }
    var csv = '\uFEFF' + out.map(function (r) { return r.join(','); }).join('\r\n');
    downloadFile(csv, 'weight_data.csv', 'text/csv');
  }

  // ---------- 事件绑定 ----------
  function bind() {
    // 缩放/框选时实时刷新预览表为当前视图范围
    if (chart) {
      chart.on('datazoom', function () {
        var range = state.followZoom ? getCurrentRange() : null;
        renderStats(range);
        renderPreview(range);
      });
    }
    $('file').addEventListener('change', function (e) {
      if (e.target.files && e.target.files[0]) handleFile(e.target.files[0]);
    });
    $('sheet').addEventListener('change', function (e) { loadSheet(e.target.value); });
    $('hasHeader').addEventListener('change', function (e) {
      state.hasHeader = e.target.checked; updateAll();
    });
    $('timeCol').addEventListener('change', function (e) { state.timeCol = +e.target.value; updateAll(); });
    $('weightCol').addEventListener('change', function (e) { state.weightCol = +e.target.value; updateAll(); });
    $('windowSec').addEventListener('input', function (e) {
      var v = parseFloat(e.target.value);
      if (isNaN(v) || v <= 0) { v = 1; e.target.value = v; }
      state.windowSec = v;
      updateAll(true);
    });
    $('deltaKg').addEventListener('input', function (e) {
      var v = parseFloat(e.target.value); state.deltaKg = isNaN(v) || v < 0 ? 0 : v; updateAll(true);
    });
    $('highlightStable').addEventListener('change', function (e) { state.highlightStable = e.target.checked; updateAll(true); });
    $('followZoom').addEventListener('change', function (e) {
      state.followZoom = e.target.checked;
      var range = e.target.checked ? getCurrentRange() : null;
      renderStats(range);
      renderPreview(range);
    });
    $('resetZoom').addEventListener('click', function () {
      if (chart) chart.dispatchAction({ type: 'restore' });
    });
    $('exportPng').addEventListener('click', exportPng);
    $('exportCsv').addEventListener('click', exportCsv);
  }

  // ---------- 启动 ----------
  document.addEventListener('DOMContentLoaded', function () {
    initChart();
    bind();
  });
})();
