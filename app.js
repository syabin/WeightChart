/* 重量波动折线图分析 —— 纯前端离线版 */
(function () {
  'use strict';

  var APP_VERSION = '2026-09-22f';

  // ---------- 拖拽诊断日志（页面回显，便于定位「拖了没反应」）----------
  // 平时隐藏；出现 ✗ 类异常时自动现身；也可以点标题旁版本徽标手动开合。
  var dragLogs = [];
  function logDrag(msg, forceShow) {
    var t = new Date();
    var hh = ('0' + t.getHours()).slice(-2) + ':' + ('0' + t.getMinutes()).slice(-2) + ':' + ('0' + t.getSeconds()).slice(-2);
    dragLogs.push(hh + ' ' + msg);
    if (dragLogs.length > 4) dragLogs.shift();
    var el = document.getElementById('dragLog');
    if (el) {
      el.textContent = '拖拽日志：' + dragLogs.join(' ｜ ');
      if (forceShow) el.style.display = '';   // 异常时自动现身，不必手动点
    }
    try { console.log('[拖拽] ' + msg); } catch (e) { }
  }
  // 把 dataTransfer 的关键信息摊平成一行文本
  function dtInfo(e) {
    var dt = e && e.dataTransfer;
    if (!dt) return 'dataTransfer=无';
    var types = [];
    try { for (var i = 0; i < dt.types.length; i++) types.push(dt.types[i]); } catch (err) { }
    var kinds = [];
    try {
      for (var j = 0; j < dt.items.length; j++) kinds.push(dt.items[j].kind + '/' + (dt.items[j].type || '?'));
    } catch (err) { }
    var nf = dt.files ? dt.files.length : '?';
    return 'files=' + nf + ' types=[' + types.join(',') + '] items=[' + kinds.join(',') + ']';
  }

  // ---------- 全局状态 ----------
  var state = {
    workbook: null,
    rows: [],            // 当前工作表原始二维数组（含/不含表头行）
    hasHeader: false,
    timeCol: 0,
    weightCol: 1,
    auxCol: -1,         // -1 = 未启用
    windowSec: 10,
    deltaKg: 0.01,
    highlightStable: true,
    followZoom: true,
    stepLine: false,        // true = 阶梯显示（step:end，数值保持到下一采样点再垂直跳变）
    // 手动坐标轴：null = 跟随默认（无余量），数字 = 手动值（全程保留，不随缩放重置）
    manualWMin: null, manualWMax: null,
    manualAMin: null, manualAMax: null,
    showBaseLine: false,    // 基准线默认不显示，点工具栏「基准线」按钮才显示
    baselineTime: null,     // Date | null，在图上点选得到；基准重量自动取该时间点重量
    setWeight: null,        // Number | null，设定重量（用于计算偏差比例）
    pickMode: false,        // true = 正在等待用户在图上点选基准时间
    parsed: [],          // [{t:Date, w:Number, a:Number|null}]
    displayData: [],     // [[Number(ms), Number]] 用于 ECharts 绘制
    auxData: [],         // [[Number(ms), Number], ...] 辅助列绘制
    isEvent: [],         // Boolean[]
    eventType: [],       // 'load'|'unload'|null
    stableFlags: [],     // Boolean[] true=平稳(噪声)
    stableRanges: [],    // [[startMs,endMs], ...]
    events: [],          // [{type,t,from,to,mag}] 聚合后的事件段
    stats: null,
    colCount: 0
  };

  var chart = null;

  // 把数值统一按整 10 圆整（最小值向下取 10 的倍数、最大值向上取 10 的倍数），用于坐标轴默认上下界
  function roundUpToNice(v) {
    if (v <= 0) return 0;
    return Math.ceil(v / 10) * 10;
  }
  function roundDownToNice(v) {
    if (v <= 0) return 0;
    return Math.floor(v / 10) * 10;
  }

  // 解析基准重量：按基准时间在 parsed 里取最近数据点的重量（基准重量 = 基准时间的重量）
  function resolveBaseline() {
    if (!state.showBaseLine || state.baselineTime == null || !state.parsed.length) return null;
    var target = state.baselineTime.getTime();
    var best = null, bestD = Infinity;
    for (var i = 0; i < state.parsed.length; i++) {
      var d = Math.abs(state.parsed[i].t.getTime() - target);
      if (d < bestD) { bestD = d; best = state.parsed[i].w; }
    }
    return best;
  }

  // 退出点选模式（恢复默认光标）
  function disablePick() {
    state.pickMode = false;
    if (chart) { try { chart.getZr().setCursorStyle('default'); } catch (e) { } }
  }

  // 进入点选模式：光标变十字，等待用户在图上点击
  function enablePick() {
    if (!chart || !state.parsed.length) return;
    state.pickMode = true;
    try { chart.getZr().setCursorStyle('crosshair'); } catch (e) { }
    updateBaseLineUI();
  }

  // 在图上按像素位置点选基准：换算成时间 → 取最近数据点 → 基准时间 = 该点时间
  function pickBaselineAt(px, py) {
    if (!chart || !state.parsed.length) return;
    if (!chart.containPixel({ gridIndex: 0 }, [px, py])) return;   // 只在网格区内响应，避开工具栏/坐标轴
    var v = chart.convertFromPixel({ gridIndex: 0 }, [px, py]);
    if (!v || !isFinite(+v[0])) return;
    var tms = +v[0];
    var best = null, bestD = Infinity;
    for (var i = 0; i < state.parsed.length; i++) {
      var d = Math.abs(state.parsed[i].t.getTime() - tms);
      if (d < bestD) { bestD = d; best = state.parsed[i]; }
    }
    if (!best) return;
    state.baselineTime = best.t;
    state.showBaseLine = true;
    disablePick();
    renderChart(true);
    updateBaseLineUI();
  }

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

  // ---------- 数据解析 ----------
  function showFileName(name) {
    var box = $('fileName');
    var hint = $('dropHint');
    if (!box) return;
    if (name) {
      $('fileNameVal').textContent = name;
      box.style.display = '';
      if (hint) hint.classList.add('hidden');
    } else {
      box.style.display = 'none';
      if (hint) hint.classList.remove('hidden');
    }
  }

  function handleFile(file) {
    showFileName(file && file.name ? file.name : null);
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
        logDrag('✓ 解析完成：' + (file && file.name ? file.name : ''));
      } catch (err) {
        logDrag('✗ 解析失败：' + err.message, true);
        alert('读取文件失败：' + err.message);
      }
    };
    reader.onerror = function () {
      logDrag('✗ 文件读取失败（FileReader error）', true);
      alert('文件读取失败，请确认文件未被其他程序占用后重试。');
    };
    try {
      reader.readAsArrayBuffer(file);
    } catch (err) {
      logDrag('✗ readAsArrayBuffer 异常：' + err.message, true);
      alert('文件读取失败：' + err.message);
    }
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
    state.auxCol = -1;
    state.showBaseLine = false;
    state.baselineTime = null;
    disablePick();
    $('timeCol').value = String(state.timeCol);
    $('weightCol').value = String(state.weightCol);
    $('auxCol').value = String(state.auxCol);

    updateAll();
  }

  function buildColumnSelectors() {
    var tc = $('timeCol'), wc = $('weightCol'), ac = $('auxCol');
    tc.innerHTML = ''; wc.innerHTML = ''; ac.innerHTML = '';
    // 辅助列第一项为「无」
    var noOpt = document.createElement('option');
    noOpt.value = '-1'; noOpt.textContent = '无';
    ac.appendChild(noOpt);
    for (var c = 0; c < state.colCount; c++) {
      var o1 = document.createElement('option'); o1.value = String(c); o1.textContent = colLabel(c);
      var o2 = document.createElement('option'); o2.value = String(c); o2.textContent = colLabel(c);
      var o3 = document.createElement('option'); o3.value = String(c); o3.textContent = colLabel(c);
      tc.appendChild(o1); wc.appendChild(o2); ac.appendChild(o3);
    }
  }

  function parseData() {
    var rows = state.rows;
    var tCol = state.timeCol, wCol = state.weightCol, aCol = state.auxCol;
    var start = state.hasHeader ? 1 : 0;
    var out = [];
    for (var i = start; i < rows.length; i++) {
      var r = rows[i];
      if (!r) continue;
      var t = parseTime(r[tCol]);
      var w = parseWeight(r[wCol]);
      if (t == null || isNaN(w)) continue;
      var a = (aCol >= 0) ? parseWeight(r[aCol]) : null;
      out.push({ t: t, w: w, a: (aCol >= 0 && !isNaN(a)) ? a : null });
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
    // 满屏自适应布局：容器尺寸随页面变化时同步重算图表
    if (window.ResizeObserver) {
      try { new ResizeObserver(function () { if (chart) chart.resize(); }).observe($('chart')); } catch (e) { }
    }
    setTimeout(function () { if (chart) chart.resize(); }, 0);
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
    var auxActive = state.auxCol >= 0;
    var auxData = [];
    if (auxActive) {
      // 基于 parsed 构建辅助列数据（时间 + 辅助值，跳过 NaN）
      for (var pi = 0; pi < state.parsed.length; pi++) {
        var p = state.parsed[pi];
        if (p.a != null && !isNaN(p.a)) auxData.push([p.t.getTime(), p.a]);
      }
    }
    state.auxData = auxData;

    // 默认上下界（无余量，仅按数据本身）：直接取 displayData 的真实 min/max
    function rawMinMax(arr) {
      var min = Infinity, max = -Infinity;
      for (var i = 0; i < arr.length; i++) {
        var v = arr[i][1];
        if (v < min) min = v;
        if (v > max) max = v;
      }
      if (!isFinite(min) || !isFinite(max)) return null;
      return { min: min, max: max };
    }
    var wData = state.displayData;
    var wRaw = rawMinMax(wData);
    // 默认值：min = 数据最小值向下圆整；max = 数据最大值 + 其 5% 预留后向上圆整到 1/5/10/50/100 台阶
    var wMin, wMax;
    if (wRaw) {
      wMin = state.manualWMin != null ? state.manualWMin : roundDownToNice(wRaw.min);
      var wDefMax = roundUpToNice(wRaw.max + Math.abs(wRaw.max) * 0.05);
      wMax = state.manualWMax != null ? state.manualWMax : wDefMax;
    } else { wMin = 0; wMax = 0; }
    // 辅助轴
    var aMin, aMax;
    if (auxActive) {
      var aRaw = rawMinMax(auxData);
      aMin = aRaw ? (state.manualAMin != null ? state.manualAMin : roundDownToNice(aRaw.min)) : 0;
      aMax = aRaw ? (state.manualAMax != null ? state.manualAMax : roundUpToNice(aRaw.max)) : 0;
    }

    var series = {
      name: '重量',
      type: 'line',
      yAxisIndex: 0,
      showSymbol: false,
      smooth: false,
      step: state.stepLine ? 'end' : false,
      large: true,
      largeThreshold: 2000,
      animation: false,
      data: data,
      lineStyle: { width: 1.4, color: '#2f6fed' },
      itemStyle: { color: '#2f6fed' },
      areaStyle: { color: 'rgba(47,111,237,0.05)' }
    };

    var seriesArr = [series];
    var auxSeries = null;
    if (auxActive) {
      auxSeries = {
        name: '辅助',
        type: 'line',
        yAxisIndex: 1,
        showSymbol: false,
        smooth: false,
        step: state.stepLine ? 'end' : false,
        animation: false,
        z: 2,                 // 放在重量线（z:3）下层，避免遮挡
        data: auxData,
        lineStyle: { width: 1.4, color: 'rgba(255,122,26,0.5)' },
        itemStyle: { color: 'rgba(255,122,26,0.5)' }
      };
      seriesArr.push(auxSeries);
      series.z = 3;          // 重量线在上层
    }

    if (state.highlightStable && state.stableRanges.length) {
      series.markArea = {
        silent: true,
        itemStyle: { color: 'rgba(154,163,178,0.06)' },
        data: state.stableRanges.map(function (r) { return [{ xAxis: r[0] }, { xAxis: r[1] }]; })
      };
    }

    // 注：已按需求移除图上的「上料 ▲ / 出料 ▼」markPoint 标记（用户反馈不准确）。
    // 上料/出料判定仍在统计面板与数据预览表中保留。

    // 基准线：y = 解析出的基准重量，仅在 state.showBaseLine 为 true 时绘制（红色横线 50% 透明）
    var baseW = resolveBaseline();
    if (baseW != null) {
      series.markLine = {
        silent: true,
        symbol: ['none', 'none'],
        lineStyle: { color: 'rgba(255,59,48,0.5)', type: 'dashed', width: 1.6 },
        label: {
          show: true,
          position: 'insideEndTop',
          color: '#ff3b30',
          fontSize: 11,
          formatter: function () { return '基准 ' + baseW.toFixed(3) + ' kg'; }
        },
        data: [{ yAxis: baseW, name: '基准' }]
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
      grid: { left: 60, right: auxActive ? 64 : 24, top: 34, bottom: 70 },
      tooltip: {
        trigger: 'axis',
        axisPointer: { type: 'line', lineStyle: { color: '#9aa3b2' } },
        backgroundColor: 'rgba(255,255,255,0.96)',
        borderColor: '#e4e8f0',
        textStyle: { color: '#1f2733' },
        // 自动避开折线：根据悬浮数据点的像素位置决定 tooltip 放上方还是下方，并做容器边缘翻转
        position: function (point, params, dom, rect, size) {
          var cw = size.contentSize[0], ch = size.contentSize[1];
          var vw = size.viewSize[0], vh = size.viewSize[1];
          var x = point[0] + 16, y = point[1] - 18;
          try {
            var p0 = params && params[0];
            if (p0 && p0.data && chart) {
              var pt = chart.convertToPixel({ seriesIndex: p0.seriesIndex }, [p0.data[0], p0.data[1]]);
              if (pt) {
                var py = pt[1];
                if (Math.abs(point[1] - py) < ch + 24) {
                  // 鼠标太靠近线：线在上 → tooltip 放下方；线在下 → 放上方
                  if (py < point[1]) y = point[1] + 18;
                  else y = point[1] - ch - 18;
                }
              }
            }
          } catch (e) { }
          if (x + cw > vw - 8) x = point[0] - cw - 16;   // 右边放不下 → 翻到左侧
          if (x < 8) x = 8;
          if (y + ch > vh - 8) y = point[1] - ch - 18;   // 下边放不下 → 翻到上方
          if (y < 8) y = 8;
          return [x, y];
        },
        formatter: function (params) {
          if (!params || !params.length) return '';
          var p = params[0];
          var time = formatTime(new Date(p.value[0]));
          // 固定三行：时间 / 重量 / 辅助，无对应数据则该值显示 —
          var wVal = null, aVal = null;
          for (var i = 0; i < params.length; i++) {
            var pp = params[i];
            var v = pp.value && pp.value[1];
            if (v == null || isNaN(v)) continue;
            if (pp.seriesName === '辅助') aVal = (+v).toFixed(3);
            else wVal = (+v).toFixed(3);
          }
          var html = '<div style="min-width:120px;">' +
            '<div style="color:#6b7686;">' + time + '</div>' +
            '<div style="margin-top:4px;"><i style="display:inline-block;width:8px;height:8px;border-radius:2px;background:#2f6fed;margin-right:6px;"></i>重量' + (wVal != null ? '：<b>' + wVal + ' kg</b>' : '') + '</div>';
          if (auxActive) {
            html += '<div style="margin-top:3px;"><i style="display:inline-block;width:8px;height:8px;border-radius:2px;background:rgba(255,122,26,0.5);margin-right:6px;"></i>辅助' + (aVal != null ? '：<b>' + aVal + '</b>' : '') + '</div>';
          }
          // 基准线：差值 = 悬浮处重量 − 基准线重量（正红负绿）
          // 偏差比例 = (|差值| − 设定重量) / 设定重量
          // 排版：基准 / 差值 / 偏差 各占一行，避免挤在一行看不全
          var bw = resolveBaseline();
          if (bw != null && wVal != null) {
            var curW = parseFloat(wVal);
            var diff = curW - bw;
            var dc = diff >= 0 ? '#e0533d' : '#2f9e6f';   // 正红负绿
            var sign = diff >= 0 ? '+' : '';
            var rows = '<div>设定 ' + bw.toFixed(3) + ' kg<span style="color:#9aa3b2;">（' + formatTime(state.baselineTime) + '）</span></div>';
            rows += '<div style="margin-top:3px;">差值 <b style="color:' + dc + ';">' + sign + diff.toFixed(3) + ' kg</b></div>';
            if (state.setWeight != null && !isNaN(state.setWeight) && state.setWeight !== 0) {
              var gap = Math.abs(diff) - state.setWeight;
              var gsign = gap >= 0 ? '+' : '';
              rows += '<div style="margin-top:3px;">差距 <b>' + gsign + gap.toFixed(3) + ' kg</b></div>';
              var ratio = (Math.abs(diff) - state.setWeight) / state.setWeight;
              var rsign = ratio >= 0 ? '+' : '';
              rows += '<div style="margin-top:3px;">偏差 <b>' + rsign + (ratio * 100).toFixed(3) + '%</b></div>';
            }
            html += '<div style="margin-top:4px; border-top:1px dashed #e4e8f0; padding-top:4px;">' + rows + '</div>';
          }
          html += '</div>';
          return html;
        }
      },
      toolbox: {
        left: 'center', top: 0,
        itemSize: 16,
        itemGap: 8,
        feature: {
          dataZoom: { yAxisIndex: 'none', title: { zoom: '框选放大' } },
          saveAsImage: { title: '保存图片', name: 'weight_chart' },
          myMarkLine: {
            show: true,
            title: '基准线',
            name: '基准线',
            icon: 'path://M2,14 L16,14 M10,8 L10,10 M10,18 L10,20',
            onclick: function () {
              if (state.baselineTime == null) {
                // 还没基准 → 进入图上点选模式
                enablePick();
              } else {
                // 已有基准 → 切换显隐
                state.showBaseLine = !state.showBaseLine;
                if (!state.showBaseLine) disablePick();
                renderChart(true);
              }
              updateBaseLineUI();
            }
          }
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
      yAxis: auxActive ? [
        {
          type: 'value',
          name: '重量 (kg)',
          min: wMin,
          max: wMax,
          nameTextStyle: { color: '#6b7686' },
          axisLabel: { color: '#6b7686', formatter: function (v) { return v.toFixed(2); } },
          splitLine: { lineStyle: { color: '#eef1f6' } }
        },
        {
          type: 'value',
          name: colLabel(state.auxCol),
          min: aMin,
          max: aMax,
          position: 'right',
          nameTextStyle: { color: '#ff7a1a' },
          axisLabel: { color: '#ff7a1a', formatter: function (v) { return v.toFixed(2); } },
          splitLine: { show: false }
        }
      ] : [
        {
          type: 'value',
          name: '重量 (kg)',
          min: wMin,
          max: wMax,
          nameTextStyle: { color: '#6b7686' },
          axisLabel: { color: '#6b7686', formatter: function (v) { return v.toFixed(2); } },
          splitLine: { lineStyle: { color: '#eef1f6' } }
        }
      ],
      dataZoom: [
        { type: 'inside', filterMode: 'none', start: savedZoom ? savedZoom.start : 0, end: savedZoom ? savedZoom.end : 100 },
        { type: 'slider', filterMode: 'none', height: 22, bottom: 18, start: savedZoom ? savedZoom.start : 0, end: savedZoom ? savedZoom.end : 100 }
      ],
      series: seriesArr
    };
    chart.setOption(opt, true);
    // 图例里的「辅助折线」随开关切换
    var legendAux = $('legendAux');
    if (legendAux) legendAux.style.display = auxActive ? '' : 'none';
    // 同步坐标轴输入区默认值 + 越界标红
    syncAxisInputs(wRaw, aRaw);
    // 图例当前值（末点）
    updateLegendValues();
  }

  function updateLegendValues() {
    function rangeOf(arr) {
      if (!arr.length) return null;
      var mn = Infinity, mx = -Infinity;
      for (var i = 0; i < arr.length; i++) {
        var v = arr[i][1];
        if (v < mn) mn = v;
        if (v > mx) mx = v;
      }
      if (!isFinite(mn)) return null;
      return { mn: mn, mx: mx };
    }
    var lw = $('legendW');
    var la = $('legendA');
    var wr = rangeOf(state.displayData);
    if (lw) lw.textContent = wr ? wr.mn.toFixed(2) + '~' + wr.mx.toFixed(2) + ' kg' : '—';
    var ar = rangeOf(state.auxData);
    if (la) la.textContent = ar ? ar.mn.toFixed(3) + '~' + ar.mx.toFixed(3) : '—';
  }

  // 同步基准线提示文案（点选中 / 已设置 / 未设置）
  function updateBaseLineUI() {
    var hint = $('baselineHint');
    if (!hint) return;
    if (state.pickMode) {
      hint.innerHTML = '<b style="color:#2f6fed;">请在折线图上点击一下</b>选定基准时间（基准重量取该点重量）';
      return;
    }
    if (state.baselineTime) {
      var bw = resolveBaseline();
      if (state.showBaseLine) {
        hint.innerHTML = '● 基准 <b>' + formatTime(state.baselineTime) + '</b> · 重量 ' +
          (bw != null ? bw.toFixed(3) + ' kg' : '—') + '（点图表顶部图标可隐藏）';
      } else {
        hint.innerHTML = '基准已设但已隐藏（点图表顶部图标可显示）· ' + formatTime(state.baselineTime);
      }
    } else {
      hint.innerHTML = '点「点选基准」或图表顶部图标，然后在图上点一下选定基准时间；「设定重量」用于算偏差比例。';
    }
  }

  // 同步坐标轴手动输入框的值与越界标红状态
  function syncAxisInputs(wRaw, aRaw) {
    function fmt(v) { return v == null || !isFinite(v) ? '' : (Math.round(v * 100) / 100).toString(); }
    function setVal(id, val, raw, minField, maxField) {
      var el = $(id);
      if (!el) return;
      // 默认值 = 无余量（数据本身）
      if (el.dataset.synced !== '1') {
        el.value = fmt(val);
        el.dataset.synced = '1';
      }
      var v = parseFloat(el.value);
      var bad = false;
      if (!isNaN(v) && raw) {
        if (minField !== undefined && v > raw.min) bad = true;          // 手动最小 > 数据最小
        if (maxField !== undefined && v < raw.max) bad = true;         // 手动最大 < 数据最大
      }
      el.classList.toggle('axis-bad', bad);
    }
    // 重量（左轴）
    setVal('wMin', state.manualWMin, wRaw, 0, undefined);
    setVal('wMax', state.manualWMax, wRaw, undefined, 0);
    // 辅助（右轴）
    var auxBox = $('auxAxisField');
    if (auxBox) auxBox.style.display = state.auxCol >= 0 ? '' : 'none';
    if (state.auxCol >= 0) {
      setVal('aMin', state.manualAMin, aRaw, 0, undefined);
      setVal('aMax', state.manualAMax, aRaw, undefined, 0);
    }
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

    var limit = Math.min(viewData.length, 1000);
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

  // ---------- 坐标轴手动 min/max ----------
  function onAxisInput(id, stateKey) {
    return function (e) {
      var v = parseFloat(e.target.value);
      state[stateKey] = isNaN(v) ? null : v;   // 清空=默认
      e.target.dataset.synced = '1';          // 标记已同步，避免 renderChart 覆盖用户输入
      e.target.classList.remove('axis-bad');
      updateAll(true);
    };
  }
  function resetAxis(which) {
    var keys = which === 'w' ? ['manualWMin', 'manualWMax'] : ['manualAMin', 'manualAMax'];
    keys.forEach(function (k) { state[k] = null; });
    var ids = which === 'w' ? ['wMin', 'wMax'] : ['aMin', 'aMax'];
    ids.forEach(function (id) { var el = $(id); if (el) el.dataset.synced = '0'; });
    updateAll(true);
  }
  function clearAxisManual() {
    ['manualWMin', 'manualWMax', 'manualAMin', 'manualAMax'].forEach(function (k) { state[k] = null; });
    ['wMin', 'wMax', 'aMin', 'aMax'].forEach(function (id) {
      var el = $(id); if (el) { el.dataset.synced = '0'; el.classList.remove('axis-bad'); }
    });
  }

  // ---------- 事件绑定 ----------
  function bind() {
    // 缩放/框选时实时刷新预览表为当前视图范围；手动坐标轴值不再被缩放重置
    if (chart) {
      chart.on('datazoom', function () {
        var range = state.followZoom ? getCurrentRange() : null;
        renderStats(range);
        renderPreview(range);
      });
      // 点选基准：点选模式下在网格区内点击，取该处最近数据点时间为基准时间
      chart.getZr().on('click', function (e) {
        if (!state.pickMode) return;
        pickBaselineAt(e.offsetX, e.offsetY);
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
    $('auxCol').addEventListener('change', function (e) {
      state.auxCol = +e.target.value;
      updateAll(true);
    });
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
    // 坐标轴手动 min/max（仅全量视图生效，缩放时自动切默认）
    $('wMin').addEventListener('input', onAxisInput('wMin', 'manualWMin'));
    $('wMax').addEventListener('input', onAxisInput('wMax', 'manualWMax'));
    $('aMin').addEventListener('input', onAxisInput('aMin', 'manualAMin'));
    $('aMax').addEventListener('input', onAxisInput('aMax', 'manualAMax'));
    $('wAxisReset').addEventListener('click', function () { resetAxis('w'); });
    $('aAxisReset').addEventListener('click', function () { resetAxis('a'); });
    // 基准线：点选基准 / 清除 / 设定重量
    $('pickBaseline').addEventListener('click', function () {
      if (state.pickMode) { disablePick(); updateBaseLineUI(); return; }
      enablePick();
    });
    $('clearBaseline').addEventListener('click', function () {
      state.showBaseLine = false;
      state.baselineTime = null;
      disablePick();
      renderChart(true);
      updateBaseLineUI();
    });
    $('setWeight').addEventListener('input', function (e) {
      var v = parseFloat(e.target.value);
      state.setWeight = isNaN(v) ? null : v;   // tooltip 每次悬浮实时读取，无需重渲染
    });
    $('followZoom').addEventListener('change', function (e) {
      state.followZoom = e.target.checked;
      var range = e.target.checked ? getCurrentRange() : null;
      renderStats(range);
      renderPreview(range);
    });
    $('stepLine').addEventListener('change', function (e) {
      state.stepLine = e.target.checked;
      renderChart(true);
    });
    bindDropzone();
  }

  // ---------- 拖拽导入 ----------
  var dragDepth = 0;
  function bindDropzone() {
    var dz = $('dropzone');
    var accepts = /\.(xlsx|xls|csv)$/i;
    document.addEventListener('dragenter', function (e) {
      e.preventDefault(); e.stopPropagation();
      dragDepth++;
      var ok = hasFile(e);
      if (ok && dz) dz.classList.add('show');
      // 拖进来的不是文件对象（例如从 Excel 拖单元格）→ 直接把日志亮出来
      logDrag('dragenter ' + dtInfo(e) + (ok ? ' · 识别为文件 ✓' : ' · 未识别为文件 ✗'), !ok);
    });
    document.addEventListener('dragover', function (e) {
      e.preventDefault(); e.stopPropagation();
      if (hasFile(e) && dz) dz.classList.add('show');
    });
    document.addEventListener('dragleave', function (e) {
      e.preventDefault(); e.stopPropagation();
      dragDepth = Math.max(0, dragDepth - 1);
      if (dragDepth === 0 && dz) dz.classList.remove('show');
    });
    document.addEventListener('drop', function (e) {
      e.preventDefault(); e.stopPropagation();
      dragDepth = 0;
      if (dz) dz.classList.remove('show');
      try {
        var all = collectFiles(e);
        logDrag('drop ' + dtInfo(e) + ' → 取到 ' + all.length + ' 个文件');
        if (!all.length) {
          // 常见误操作：从 Excel 里拖选中的单元格内容 —— 那不是文件，页面上收不到文件
          logDrag('✗ 没有文件对象，已终止', true);
          alert('没有检测到文件。\n请从「文件资源管理器 / 访达」里把 .xlsx / .xls / .csv 文件拖进页面；\n从 Excel 里拖选中的单元格不算文件。');
          return;
        }
        var f = null;
        for (var i = 0; i < all.length; i++) { if (accepts.test(all[i].name || '')) { f = all[i]; break; } }
        if (!f) {
          logDrag('✗ 类型不支持：' + (all[0].name || '未知'), true);
          alert('暂不支持该文件类型，请拖入 .xlsx / .xls / .csv 文件。\n（收到：' + (all[0].name || '未知文件') + '）');
          return;
        }
        if (all.length > 1) {
          logDrag('多选 ' + all.length + ' 个，取「' + f.name + '」，其余忽略');
        }
        logDrag('→ 交给解析：' + f.name + (f.size ? '（' + Math.round(f.size / 1024) + ' KB）' : ''));
        handleFile(f);
      } catch (err) {
        logDrag('✗ drop 处理出错：' + err.message, true);
        alert('拖拽导入出错：' + err.message);
      }
    });

    // 收集拖入的文件：优先 dataTransfer.files；网盘等虚拟文件走 items.getAsFile()
    function collectFiles(e) {
      var dt = e.dataTransfer;
      if (!dt) return [];
      var out = [];
      if (dt.files && dt.files.length) {
        out = Array.prototype.slice.call(dt.files, 0);
      } else if (dt.items) {
        for (var i = 0; i < dt.items.length; i++) {
          var it = dt.items[i];
          if (it && it.kind === 'file' && it.getAsFile) {
            var f = it.getAsFile();
            if (f) out.push(f);
          }
        }
      }
      return out;
    }

    // 是否真的拖了文件：types 含 'Files' 或 items 里有 kind=file
    function hasFile(e) {
      var dt = e.dataTransfer;
      if (!dt) return false;
      if (dt.types && Array.prototype.indexOf.call(dt.types, 'Files') >= 0) return true;
      if (dt.items && dt.items.length) {
        for (var i = 0; i < dt.items.length; i++) {
          if (dt.items[i] && dt.items[i].kind === 'file') return true;
        }
      }
      return false;
    }

    // 兜底入口：在资源管理器 Ctrl+C 复制文件 → 页面里 Ctrl+V 粘贴导入
    // （Windows 下浏览器以管理员身份运行、或企业策略禁用跨窗口拖放时，拖拽会失效，粘贴仍可用）
    document.addEventListener('paste', function (e) {
      try {
        var cd = e.clipboardData;
        if (!cd || !cd.files || !cd.files.length) return;   // 粘贴的是文本/表格，不处理
        var f = null;
        for (var i = 0; i < cd.files.length; i++) {
          if (accepts.test(cd.files[i].name || '')) { f = cd.files[i]; break; }
        }
        if (!f) { logDrag('✗ 粘贴的文件类型不支持：' + (cd.files[0].name || '未知')); return; }
        logDrag('✓ 粘贴导入：' + f.name);
        handleFile(f);
      } catch (err) { logDrag('✗ 粘贴处理出错：' + err.message); }
    });
  }

  // ---------- 启动 ----------
  // 注：index.html 用动态 <script> 带时间戳加载本文件（破 CDN 缓存），
  //     动态脚本可能在 DOMContentLoaded 之后才执行，故需按 readyState 判断。
  function start() {
    initChart();
    bind();
    var v = document.getElementById('appVer');
    if (v) {
      v.textContent = 'v' + APP_VERSION;
      // 版本徽标点一下可显示/隐藏拖拽日志（默认隐藏，排查问题时再叫出来）
      v.title = '点击显示 / 隐藏拖拽日志';
      v.style.cursor = 'pointer';
      v.addEventListener('click', function () {
        var el = document.getElementById('dragLog');
        if (el) el.style.display = (el.style.display === 'none' ? '' : 'none');
      });
    }
    try { console.log('[WeightChart] app.js v' + APP_VERSION + ' 已加载，拖拽监听已就绪（点标题旁版本徽标可显示拖拽日志）'); } catch (e) { }
    // URL 带 debug 时日志常显，例如 index.html?debug —— 方便远程排查
    try {
      if (location.search.indexOf('debug') >= 0 || location.hash.indexOf('debug') >= 0) {
        var dl = document.getElementById('dragLog');
        if (dl) dl.style.display = '';
      }
    } catch (e) { }
  }
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', start);
  } else {
    start();
  }
})();
