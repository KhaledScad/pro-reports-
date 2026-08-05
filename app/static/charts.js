/* Dependency-free, multi-series SVG bar + pulse (line) charts for the compare pages. */
(function () {
  "use strict";

  var SVG_NS = "http://www.w3.org/2000/svg";
  var PALETTE = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#008300", "#4a3aa7", "#e34948"];
  var MAX_SERIES = 8;

  function fmtNum(n) {
    if (n === null || n === undefined || isNaN(n)) return "0";
    var abs = Math.abs(n);
    return n.toLocaleString(undefined, { maximumFractionDigits: abs >= 100 ? 0 : 2 });
  }

  function svgEl(tag, attrs) {
    var el = document.createElementNS(SVG_NS, tag);
    for (var k in attrs) el.setAttribute(k, attrs[k]);
    return el;
  }

  function clear(el) {
    while (el.firstChild) el.removeChild(el.firstChild);
  }

  function textEl(attrs, text) {
    var el = svgEl("text", attrs);
    el.textContent = text;
    return el;
  }

  function niceTicks(min, max, count) {
    if (min === max) { min -= 1; max += 1; }
    var range = max - min;
    var rawStep = range / count;
    var mag = Math.pow(10, Math.floor(Math.log(rawStep) / Math.LN10));
    var norm = rawStep / mag;
    var step = norm < 1.5 ? mag : norm < 3 ? 2 * mag : norm < 7 ? 5 * mag : 10 * mag;
    var niceMin = Math.floor(min / step) * step;
    var niceMax = Math.ceil(max / step) * step;
    var ticks = [];
    for (var v = niceMin; v <= niceMax + step * 0.5; v += step) ticks.push(Math.round(v * 1e6) / 1e6);
    return { ticks: ticks, min: niceMin, max: niceMax };
  }

  function barPath(x, w, yTop, yBottom, roundTop, r) {
    r = Math.min(r, w / 2, Math.max(yBottom - yTop, 0));
    if (r <= 0.01) {
      return "M" + x + "," + yBottom + " L" + x + "," + yTop + " L" + (x + w) + "," + yTop + " L" + (x + w) + "," + yBottom + " Z";
    }
    if (roundTop) {
      return ["M", x, yBottom, "L", x, yTop + r, "Q", x, yTop, x + r, yTop,
        "L", x + w - r, yTop, "Q", x + w, yTop, x + w, yTop + r, "L", x + w, yBottom, "Z"].join(" ");
    }
    return ["M", x, yTop, "L", x, yBottom - r, "Q", x, yBottom, x + r, yBottom,
      "L", x + w - r, yBottom, "Q", x + w, yBottom, x + w, yBottom - r, "L", x + w, yTop, "Z"].join(" ");
  }

  function renderTooltip(tooltipEl, data) {
    clear(tooltipEl);
    tooltipEl.appendChild(textLine("chart-tooltip-title", data.title));
    data.rows.forEach(function (row) {
      var rowEl = document.createElement("div");
      rowEl.className = "chart-tooltip-row";
      var key = document.createElement("span");
      key.className = "chart-tooltip-key";
      key.style.background = row.color;
      var name = document.createElement("span");
      name.className = "chart-tooltip-name";
      name.textContent = row.name;
      var val = document.createElement("span");
      val.className = "chart-tooltip-value";
      val.textContent = fmtNum(row.value);
      rowEl.appendChild(key);
      rowEl.appendChild(name);
      rowEl.appendChild(val);
      tooltipEl.appendChild(rowEl);
    });
  }

  function textLine(className, text) {
    var el = document.createElement("div");
    el.className = className;
    el.textContent = text;
    return el;
  }

  function positionTooltip(tooltipEl, wrapEl, clientX, clientY) {
    var wrapRect = wrapEl.getBoundingClientRect();
    var x = clientX - wrapRect.left + 14;
    var y = clientY - wrapRect.top + 14;
    tooltipEl.style.left = x + "px";
    tooltipEl.style.top = y + "px";
    tooltipEl.hidden = false;
    requestAnimationFrame(function () {
      var tw = tooltipEl.offsetWidth, th = tooltipEl.offsetHeight;
      var maxX = wrapRect.width - tw - 4;
      var maxY = wrapRect.height - th - 4;
      if (x > maxX) tooltipEl.style.left = Math.max(4, maxX) + "px";
      if (y > maxY) tooltipEl.style.top = Math.max(4, maxY) + "px";
    });
  }

  function attachMarkEvents(el, tooltipEl, wrapEl, getData) {
    function showAt(clientX, clientY) {
      renderTooltip(tooltipEl, getData());
      positionTooltip(tooltipEl, wrapEl, clientX, clientY);
    }
    el.addEventListener("pointerenter", function (e) { showAt(e.clientX, e.clientY); });
    el.addEventListener("pointermove", function (e) { showAt(e.clientX, e.clientY); });
    el.addEventListener("pointerleave", function () { tooltipEl.hidden = true; });
    el.addEventListener("focus", function () {
      var r = el.getBoundingClientRect();
      showAt(r.left + r.width / 2, r.top);
    });
    el.addEventListener("blur", function () { tooltipEl.hidden = true; });
  }

  function renderLegend(el, seriesKeys, colorOf, nameOf) {
    if (!el) return;
    clear(el);
    seriesKeys.forEach(function (key) {
      var item = document.createElement("span");
      item.className = "chart-legend-item";
      var swatch = document.createElement("span");
      swatch.className = "chart-legend-swatch";
      swatch.style.background = colorOf(key);
      var name = document.createElement("span");
      name.className = "chart-legend-name";
      name.textContent = nameOf(key);
      item.appendChild(swatch);
      item.appendChild(name);
      el.appendChild(item);
    });
  }

  var PCT_RE = /_pct$/;
  var MAGNITUDE_RATIO = 0.12;

  // Split series into chart groups so wildly different scales never share one
  // axis: percentages always get their own chart, and amount-type variables
  // split into "large" / "small" scale groups only when their peaks actually differ.
  function groupBarSeries(seriesKeys, seriesData) {
    var pctKeys = seriesKeys.filter(function (k) { return PCT_RE.test(k); });
    var amountKeys = seriesKeys.filter(function (k) { return !PCT_RE.test(k); });

    var large = [], small = [];
    if (amountKeys.length) {
      var peaks = {};
      amountKeys.forEach(function (k) {
        var m = 0;
        seriesData[k].forEach(function (v) { m = Math.max(m, Math.abs(v)); });
        peaks[k] = m;
      });
      var overallMax = Math.max.apply(null, amountKeys.map(function (k) { return peaks[k]; }));
      amountKeys.forEach(function (k) {
        if (overallMax === 0 || peaks[k] >= overallMax * MAGNITUDE_RATIO) large.push(k); else small.push(k);
      });
    }

    var groups = [];
    if (large.length) groups.push({ kind: "large", keys: large });
    if (small.length) groups.push({ kind: "small", keys: small });
    if (pctKeys.length) groups.push({ kind: "pct", keys: pctKeys });
    return groups;
  }

  function buildChartBox(container, title) {
    var box = document.createElement("div");
    box.className = "chart-box";
    var h4 = document.createElement("h4");
    h4.textContent = title;
    var legend = document.createElement("div");
    legend.className = "chart-legend";
    var wrap = document.createElement("div");
    wrap.className = "chart-canvas-wrap";
    var svg = svgEl("svg", { class: "chart-svg" });
    var tooltip = document.createElement("div");
    tooltip.className = "chart-tooltip";
    tooltip.hidden = true;
    wrap.appendChild(svg);
    wrap.appendChild(tooltip);
    box.appendChild(h4);
    box.appendChild(legend);
    box.appendChild(wrap);
    container.appendChild(box);
    return { svg: svg, legend: legend, tooltip: tooltip, wrap: wrap };
  }

  function drawYAxis(svg, ticks, min, max, padLeft, padRight, padTop, padBottom, width, height) {
    var usable = height - padTop - padBottom;
    ticks.forEach(function (tv) {
      var y = padTop + usable - ((tv - min) / (max - min)) * usable;
      svg.appendChild(svgEl("line", { x1: padLeft, y1: y, x2: width - padRight, y2: y, class: "chart-gridline" }));
      svg.appendChild(textEl({ x: padLeft - 8, y: y + 3, "text-anchor": "end", class: "chart-tick-label" }, fmtNum(tv)));
    });
  }

  function renderBarChart(opts) {
    var svg = opts.svg, labels = opts.labels, seriesKeys = opts.seriesKeys, seriesData = opts.seriesData;
    var colorOf = opts.colorOf, nameOf = opts.nameOf, width = opts.width, height = opts.height;
    clear(svg);
    svg.setAttribute("viewBox", "0 0 " + width + " " + height);
    svg.setAttribute("preserveAspectRatio", "xMidYMid meet");

    var padLeft = 56, padRight = 16, padTop = 16, padBottom = 34;
    var plotW = width - padLeft - padRight;

    var all = [];
    seriesKeys.forEach(function (k) { all = all.concat(seriesData[k]); });
    all.push(0);
    var domain = niceTicks(Math.min.apply(null, all), Math.max.apply(null, all), 4);
    var usable = height - padTop - padBottom;
    var y = function (v) { return padTop + usable - ((v - domain.min) / (domain.max - domain.min)) * usable; };
    var zeroY = y(0);

    drawYAxis(svg, domain.ticks, domain.min, domain.max, padLeft, padRight, padTop, padBottom, width, height);
    svg.appendChild(svgEl("line", { x1: padLeft, y1: zeroY, x2: width - padRight, y2: zeroY, class: "chart-baseline" }));

    var n = seriesKeys.length || 1;
    var groupSlot = plotW / Math.max(labels.length, 1);
    var maxBarW = 24, gap = 2;
    var barW = Math.min(maxBarW, (groupSlot * 0.72 - (n - 1) * gap) / n);
    barW = Math.max(barW, 3);
    var groupInnerW = n * barW + (n - 1) * gap;

    labels.forEach(function (label, i) {
      var groupX = padLeft + i * groupSlot + (groupSlot - groupInnerW) / 2;
      seriesKeys.forEach(function (key, si) {
        var v = seriesData[key][i] || 0;
        var barX = groupX + si * (barW + gap);
        var yTop = Math.min(y(v), zeroY), yBottom = Math.max(y(v), zeroY);
        var mark = svgEl("path", {
          d: barPath(barX, barW, yTop, yBottom, v >= 0, 4),
          fill: colorOf(key), class: "chart-mark", tabindex: "0",
          "aria-label": label + " " + nameOf(key) + " " + fmtNum(v),
        });
        attachMarkEvents(mark, opts.tooltip, opts.wrap, function () {
          return { title: label, rows: [{ name: nameOf(key), value: v, color: colorOf(key) }] };
        });
        svg.appendChild(mark);
      });
      svg.appendChild(textEl({ x: groupX + groupInnerW / 2, y: height - padBottom + 18, "text-anchor": "middle", class: "chart-axis-label" }, label));
    });
  }

  function renderPulseChart(opts) {
    var svg = opts.svg, labels = opts.labels, seriesKeys = opts.seriesKeys, seriesData = opts.seriesData;
    var colorOf = opts.colorOf, nameOf = opts.nameOf, width = opts.width, height = opts.height;
    clear(svg);
    svg.setAttribute("viewBox", "0 0 " + width + " " + height);
    svg.setAttribute("preserveAspectRatio", "xMidYMid meet");

    var padLeft = 56, padRight = 16, padTop = 16, padBottom = 34;
    var plotW = width - padLeft - padRight;

    var all = [];
    seriesKeys.forEach(function (k) { all = all.concat(seriesData[k]); });
    all.push(0);
    var domain = niceTicks(Math.min.apply(null, all), Math.max.apply(null, all), 4);
    var usable = height - padTop - padBottom;
    var y = function (v) { return padTop + usable - ((v - domain.min) / (domain.max - domain.min)) * usable; };
    var stepX = labels.length > 1 ? plotW / (labels.length - 1) : 0;
    var xAt = function (i) { return padLeft + i * stepX; };

    drawYAxis(svg, domain.ticks, domain.min, domain.max, padLeft, padRight, padTop, padBottom, width, height);
    svg.appendChild(svgEl("line", { x1: padLeft, y1: y(0), x2: width - padRight, y2: y(0), class: "chart-baseline" }));

    var crosshair = svgEl("line", { x1: padLeft, y1: padTop, x2: padLeft, y2: height - padBottom, class: "chart-crosshair", opacity: "0" });

    // Column hit targets (drawn first, under the lines/dots) so pointing anywhere
    // in a quarter's column shows every series at that quarter.
    labels.forEach(function (label, i) {
      var left = i === 0 ? padLeft : xAt(i) - stepX / 2;
      var right = i === labels.length - 1 ? width - padRight : xAt(i) + stepX / 2;
      var hit = svgEl("rect", {
        x: left, y: padTop, width: Math.max(right - left, 1), height: usable,
        class: "chart-hit", "pointer-events": "all", tabindex: "0",
        "aria-label": label,
      });
      function showColumn(clientX, clientY) {
        crosshair.setAttribute("x1", xAt(i));
        crosshair.setAttribute("x2", xAt(i));
        crosshair.setAttribute("opacity", "1");
        renderTooltip(opts.tooltip, {
          title: label,
          rows: seriesKeys.map(function (k) { return { name: nameOf(k), value: seriesData[k][i], color: colorOf(k) }; }),
        });
        positionTooltip(opts.tooltip, opts.wrap, clientX, clientY);
      }
      hit.addEventListener("pointerenter", function (e) { showColumn(e.clientX, e.clientY); });
      hit.addEventListener("pointermove", function (e) { showColumn(e.clientX, e.clientY); });
      hit.addEventListener("pointerleave", function () { crosshair.setAttribute("opacity", "0"); opts.tooltip.hidden = true; });
      hit.addEventListener("focus", function () {
        var r = hit.getBoundingClientRect();
        showColumn(r.left + r.width / 2, r.top);
      });
      hit.addEventListener("blur", function () { crosshair.setAttribute("opacity", "0"); opts.tooltip.hidden = true; });
      svg.appendChild(hit);
    });

    svg.appendChild(crosshair);

    seriesKeys.forEach(function (key) {
      var color = colorOf(key);
      var pts = seriesData[key].map(function (v, i) { return [xAt(i), y(v)]; });
      var d = pts.map(function (p, i) { return (i === 0 ? "M" : "L") + p[0].toFixed(1) + " " + p[1].toFixed(1); }).join(" ");
      svg.appendChild(svgEl("path", {
        d: d, fill: "none", stroke: color, "stroke-width": 2,
        "stroke-linejoin": "round", "stroke-linecap": "round", "pointer-events": "none",
      }));
    });

    seriesKeys.forEach(function (key) {
      var color = colorOf(key);
      seriesData[key].forEach(function (v, i) {
        var dot = svgEl("circle", {
          cx: xAt(i), cy: y(v), r: 5, fill: color, stroke: "#1f2937", "stroke-width": 1.5,
          class: "chart-mark", tabindex: "0", "aria-label": labels[i] + " " + nameOf(key) + " " + fmtNum(v),
        });
        attachMarkEvents(dot, opts.tooltip, opts.wrap, function () {
          return { title: labels[i], rows: [{ name: nameOf(key), value: v, color: color }] };
        });
        svg.appendChild(dot);
      });
    });

    labels.forEach(function (label, i) {
      svg.appendChild(textEl({ x: xAt(i), y: height - padBottom + 18, "text-anchor": "middle", class: "chart-axis-label" }, label));
    });
  }

  window.initTrendCharts = function (config) {
    var checklist = document.getElementById(config.checklistId);
    var barColumn = document.getElementById(config.barColumnId);
    var pulseSvg = document.getElementById(config.pulseSvgId);
    if (!checklist || !barColumn || !pulseSvg) return;

    var pulseWrap = pulseSvg.parentElement;
    var pulseTooltip = document.getElementById(config.pulseTooltipId);
    var pulseLegend = document.getElementById(config.pulseLegendId);
    var groupTitles = {
      large: config.groupLargeLabel, small: config.groupSmallLabel, pct: config.groupPctLabel,
    };
    var nameByKey = {};
    config.variables.forEach(function (v) { nameByKey[v.key] = v.label; });
    var nameOf = function (key) { return nameByKey[key] || key; };

    var checkboxes = Array.prototype.slice.call(checklist.querySelectorAll("input[type=checkbox]"));

    function checkedKeysInMasterOrder() {
      var checkedSet = {};
      checkboxes.forEach(function (cb) { if (cb.checked) checkedSet[cb.value] = true; });
      return config.variables.map(function (v) { return v.key; }).filter(function (k) { return checkedSet[k]; });
    }

    function colorMap(seriesKeys) {
      var map = {};
      seriesKeys.forEach(function (k, i) { map[k] = PALETTE[i % PALETTE.length]; });
      return map;
    }

    function updateChipVisuals(seriesKeys, colors) {
      checkboxes.forEach(function (cb) {
        var label = cb.closest(".var-chip");
        var dot = label.querySelector(".var-chip-dot");
        var checked = cb.checked;
        label.classList.toggle("is-checked", checked);
        dot.style.background = checked ? colors[cb.value] : "#e5e7eb";
        var disable = !checked && seriesKeys.length >= MAX_SERIES;
        cb.disabled = disable;
        label.classList.toggle("is-disabled", disable);
      });
    }

    function draw() {
      var seriesKeys = checkedKeysInMasterOrder();
      var colors = colorMap(seriesKeys);
      updateChipVisuals(seriesKeys, colors);
      var colorOf = function (k) { return colors[k]; };
      renderLegend(pulseLegend, seriesKeys, colorOf, nameOf);

      clear(barColumn);
      var groups = groupBarSeries(seriesKeys, config.series);
      groups.forEach(function (group) {
        var title = groups.length > 1 ? config.barChartLabel + " — " + groupTitles[group.kind] : config.barChartLabel;
        var refs = buildChartBox(barColumn, title);
        renderLegend(refs.legend, group.keys, colorOf, nameOf);
        renderBarChart({
          svg: refs.svg, wrap: refs.wrap, tooltip: refs.tooltip, width: refs.wrap.clientWidth || 480, height: 280,
          labels: config.labels, seriesKeys: group.keys, seriesData: config.series, colorOf: colorOf, nameOf: nameOf,
        });
      });

      var pulseWidth = pulseWrap.clientWidth || 480;
      renderPulseChart({
        svg: pulseSvg, wrap: pulseWrap, tooltip: pulseTooltip, width: pulseWidth, height: 280,
        labels: config.labels, seriesKeys: seriesKeys, seriesData: config.series, colorOf: colorOf, nameOf: nameOf,
      });
    }

    checklist.addEventListener("change", function (e) {
      var checkedCount = checkboxes.filter(function (cb) { return cb.checked; }).length;
      if (checkedCount === 0) {
        e.target.checked = true;
        return;
      }
      if (checkedCount > MAX_SERIES) {
        e.target.checked = false;
        return;
      }
      draw();
    });
    window.addEventListener("resize", draw);
    draw();
  };
})();
