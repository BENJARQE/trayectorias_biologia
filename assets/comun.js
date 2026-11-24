// Utilidades compartidas (vanilla JS + SVG, sin dependencias). Expone el objeto global TM.
(function () {
  "use strict";
  const NS = "http://www.w3.org/2000/svg";
  const TM = (window.TM = {});

  TM.esOscuro = () => window.matchMedia("(prefers-color-scheme: dark)").matches;
  TM.color = (o) => (TM.esOscuro() ? o.colorOscuro : o.colorClaro);
  TM.fmt = (n) => Number(n).toLocaleString("es-CO");
  TM.pct = (n, t) => (t ? ((100 * n) / t).toFixed(1).replace(".", ",") + " %" : "—");
  TM.esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));

  TM.svg = function (tag, attrs, padre) {
    const e = document.createElementNS(NS, tag);
    for (const k in attrs || {}) e.setAttribute(k, attrs[k]);
    if (padre) padre.appendChild(e);
    return e;
  };
  TM.el = function (tag, props, padre) {
    const e = document.createElement(tag);
    Object.assign(e, props || {});
    if (padre) padre.appendChild(e);
    return e;
  };

  // ── Tooltip ────────────────────────────────────────────────────────────
  const tip = () => document.getElementById("tooltip");
  TM.mostrarTip = function (html, ev) {
    const t = tip();
    t.innerHTML = html;
    t.style.display = "block";
    TM.moverTip(ev);
  };
  TM.moverTip = function (ev) {
    const t = tip();
    let x = ev.clientX + 14, y = ev.clientY + 14;
    if (x + t.offsetWidth + 8 > window.innerWidth) x = ev.clientX - t.offsetWidth - 14;
    if (y + t.offsetHeight + 8 > window.innerHeight) y = ev.clientY - t.offsetHeight - 14;
    t.style.left = Math.max(4, x) + "px";
    t.style.top = Math.max(4, y) + "px";
  };
  TM.ocultarTip = () => { tip().style.display = "none"; };

  // ── Donut ──────────────────────────────────────────────────────────────
  function punto(cx, cy, r, ang) {
    const a = ((ang - 90) * Math.PI) / 180;
    return [cx + r * Math.cos(a), cy + r * Math.sin(a)];
  }
  function arco(cx, cy, r0, r1, a0, a1) {
    if (a1 - a0 >= 359.99) {  // círculo completo: dos medios arcos
      const m = a0 + 180;
      return arco(cx, cy, r0, r1, a0, m - 0.01) + " " + arco(cx, cy, r0, r1, m, a0 + 359.98);
    }
    const [x0, y0] = punto(cx, cy, r1, a0), [x1, y1] = punto(cx, cy, r1, a1);
    const [x2, y2] = punto(cx, cy, r0, a1), [x3, y3] = punto(cx, cy, r0, a0);
    const g = a1 - a0 > 180 ? 1 : 0;
    return `M${x0},${y0} A${r1},${r1} 0 ${g} 1 ${x1},${y1} L${x2},${y2} A${r0},${r0} 0 ${g} 0 ${x3},${y3} Z`;
  }
  /**
   * items: [{id, valor, color, opacidad?, etiqueta}]. opciones: {tam, interior, seleccionado, centro:[valor, rotulo], alClic, tip(item,total)}
   */
  TM.donut = function (contenedor, items, op) {
    const tam = op.tam || 300, c = tam / 2, r1 = tam / 2 - 6, r0 = r1 * (op.interior || 0.6);
    const total = items.reduce((s, i) => s + i.valor, 0);
    const svg = TM.svg("svg", { width: tam, height: tam, viewBox: `0 0 ${tam} ${tam}`, role: "img", "aria-label": op.aria || "Gráfico de torta" });
    let ang = 0;
    items.forEach((it) => {
      if (!it.valor) return;
      const a1 = ang + (it.valor / total) * 360;
      const p = TM.svg("path", {
        d: arco(c, c, r0, r1, ang, a1), fill: it.color, "fill-opacity": it.opacidad == null ? 1 : it.opacidad,
        stroke: "var(--panel)", "stroke-width": 2, class: "rebanada" + (op.seleccionado && op.seleccionado !== it.id ? " apagada" : ""),
        tabindex: 0, role: "button", "data-category": it.id, "aria-label": `${it.etiqueta}: ${TM.fmt(it.valor)} (${TM.pct(it.valor, total)})`,
      }, svg);
      p.addEventListener("mousemove", (e) => TM.mostrarTip(op.tip ? op.tip(it, total) : `<strong>${TM.esc(it.etiqueta)}</strong><br>${TM.fmt(it.valor)} · ${TM.pct(it.valor, total)}`, e));
      p.addEventListener("mouseleave", TM.ocultarTip);
      if (op.alClic) {
        p.addEventListener("click", () => op.alClic(it.id));
        p.addEventListener("keydown", (e) => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); op.alClic(it.id); } });
      }
      ang = a1;
    });
    if (op.centro) {
      TM.svg("text", { x: c, y: c + 4, class: "centro-valor" }, svg).textContent = op.centro[0];
      TM.svg("text", { x: c, y: c + 24, class: "centro-rotulo" }, svg).textContent = op.centro[1];
    }
    contenedor.appendChild(svg);
    return total;
  };

  // ── Lista de barras (empleadores / cargos) ─────────────────────────────
  TM.barras = function (contenedor, filas, color) {
    const max = Math.max(1, ...filas.map((f) => f.valor));
    filas.forEach((f) => {
      const fila = TM.el("div", { className: "fila-barra" }, contenedor);
      TM.el("div", { className: "nom", textContent: f.nombre }, fila);
      const pista = TM.el("div", { className: "pista" }, fila);
      const r = TM.el("div", { className: "relleno" }, pista);
      r.style.width = (100 * f.valor) / max + "%";
      r.style.background = f.color || color;
      TM.el("div", { className: "num", textContent: TM.fmt(f.valor) }, fila);
      fila.addEventListener("mousemove", (e) => TM.mostrarTip(f.tip, e));
      fila.addEventListener("mouseleave", TM.ocultarTip);
    });
  };

  // ── Selector de cohorte y unidad ───────────────────────────────────────
  TM.chips = function (contenedor, opciones, valor, alCambiar) {
    contenedor.innerHTML = "";
    opciones.forEach((o) => {
      const b = TM.el("button", { type: "button", className: "chip", textContent: o.label }, contenedor);
      b.setAttribute("aria-pressed", String(o.id === valor));
      b.addEventListener("click", () => alCambiar(o.id));
    });
  };

  TM.ordenar = function (lista, modo, valorDe, nombreDe) {
    const l = lista.slice();
    if (modo === "alfa") l.sort((a, b) => nombreDe(a).localeCompare(nombreDe(b), "es"));
    else if (modo === "asc") l.sort((a, b) => valorDe(a) - valorDe(b) || nombreDe(a).localeCompare(nombreDe(b), "es"));
    else l.sort((a, b) => valorDe(b) - valorDe(a) || nombreDe(a).localeCompare(nombreDe(b), "es"));
    return l;
  };

  window.matchMedia("(prefers-color-scheme: dark)").addEventListener("change", () => { if (TM.redibujar) TM.redibujar(); });
})();

