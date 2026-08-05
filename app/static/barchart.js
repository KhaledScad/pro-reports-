/* Self-contained grouped bar chart: quarters on the x-axis, one bar per
   selected attribute per quarter. No external chart library — plain SVG. */
(function () {
  var COLORS = ["#d61f2c", "#0f766e", "#7c3aed", "#b45309", "#1d4ed8", "#15803d", "#be185d", "#525252"];
  var SVG_NS = "http://www.w3.org/2000/svg";

  function formatNumber(v) {
    var r = Math.round(v * 100) / 100;
    return r.toLocaleString();
  }

  function el(tag, attrs) {
    var e = document.createElementNS(SVG_NS, tag);
    for (var k in attrs) e.setAttribute(k, attrs[k]);
    return e;
  }

  function initBarChart(containerId, dataScriptId, attrsScriptId) {
    var container = document.getElementById(containerId);
    if (!container) return;
    var data = JSON.parse(document.getElementById(dataScriptId).textContent || "[]");
    var attrs = JSON.parse(document.getElementById(attrsScriptId).textContent || "[]");
    var controls = container.querySelector(".chart-controls");
    var plot = container.querySelector(".chart-svg-wrap");
    if (!controls || !plot) return;

    function selectedAttrs() {
      var keys = Array.prototype.slice
        .call(controls.querySelectorAll("input[type=checkbox]:checked"))
        .map(function (cb) { return cb.value; });
      return attrs.filter(function (a) { return keys.indexOf(a.key) !== -1; });
    }

    function render() {
      var chosen = selectedAttrs();
      plot.innerHTML = "";

      if (!chosen.length || !data.length) {
        var msg = document.createElement("p");
        msg.className = "muted";
        msg.textContent = plot.getAttribute("data-empty-msg") || "Select at least one attribute.";
        plot.appendChild(msg);
        return;
      }

      var W = Math.max(560, data.length * chosen.length * 44 + 90);
      var H = 320;
      var padL = 60, padR = 20, padT = 20, padB = 56;
      var plotW = W - padL - padR, plotH = H - padT - padB;

      var maxVal = 0, minVal = 0;
      data.forEach(function (d) {
        chosen.forEach(function (a) {
          var v = Number(d[a.key]) || 0;
          if (v > maxVal) maxVal = v;
          if (v < minVal) minVal = v;
        });
      });
      if (maxVal === 0 && minVal === 0) maxVal = 1;
      var range = maxVal - minVal || 1;
      function y(v) { return padT + (plotH * (maxVal - v)) / range; }
      var zeroY = y(0);

      var svg = el("svg", { viewBox: "0 0 " + W + " " + H, width: "100%", height: H, role: "img" });

      // gridlines + y-axis labels (0, max, min)
      [maxVal, 0, minVal].forEach(function (v) {
        var yy = y(v);
        svg.appendChild(el("line", { x1: padL, x2: W - padR, y1: yy, y2: yy, stroke: "#e5e7eb" }));
        var label = el("text", { x: padL - 8, y: yy + 4, "text-anchor": "end", "font-size": "11", fill: "#6b7280" });
        label.textContent = formatNumber(v);
        svg.appendChild(label);
      });

      var groupW = plotW / data.length;
      var barW = Math.max(6, (groupW - 12) / chosen.length);

      data.forEach(function (d, gi) {
        var gx = padL + gi * groupW;
        chosen.forEach(function (a, ai) {
          var v = Number(d[a.key]) || 0;
          var barX = gx + 6 + ai * barW;
          var barY = Math.min(y(v), zeroY);
          var barH = Math.max(0, Math.abs(y(v) - zeroY));
          var rect = el("rect", {
            x: barX, y: barY, width: Math.max(2, barW - 3), height: barH,
            fill: COLORS[ai % COLORS.length], rx: "2",
          });
          var title = document.createElementNS(SVG_NS, "title");
          title.textContent = a.label + " — " + d.periodLabel + ": " + formatNumber(v);
          rect.appendChild(title);
          svg.appendChild(rect);

          if (barH > 0) {
            var vlabel = el("text", {
              x: barX + (barW - 3) / 2, y: v >= 0 ? barY - 4 : barY + barH + 12,
              "text-anchor": "middle", "font-size": "10", fill: "#374151",
            });
            vlabel.textContent = formatNumber(v);
            svg.appendChild(vlabel);
          }
        });

        var xlabel = el("text", {
          x: gx + groupW / 2, y: H - padB + 20, "text-anchor": "middle",
          "font-size": "12", "font-weight": "700", fill: "#111827",
        });
        xlabel.textContent = d.periodLabel;
        svg.appendChild(xlabel);
      });

      plot.appendChild(svg);

      var legend = document.createElement("div");
      legend.className = "chart-legend";
      chosen.forEach(function (a, ai) {
        var item = document.createElement("span");
        item.className = "chart-legend-item";
        var dot = document.createElement("span");
        dot.className = "chart-legend-dot";
        dot.style.background = COLORS[ai % COLORS.length];
        item.appendChild(dot);
        item.appendChild(document.createTextNode(a.label));
        legend.appendChild(item);
      });
      plot.appendChild(legend);
    }

    controls.addEventListener("change", render);
    render();
  }

  window.initBarChart = initBarChart;
})();
