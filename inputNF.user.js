// ==UserScript==
// @name         input para associar NF
// @namespace    brunoAF
// @version      1.0
// @description  campo input pra associar NF
// @match        http://governanca.mongeral.seguros/Lists/BaseProvisaoPagamento/0204%20%20Proviso%20%20A%20pagar%20%20Por%20data%20limite%20para%20lanar%20PAGNET%202024%20e%20Por%20responsvel.aspx?InitialTabId=Ribbon%2EListItem&VisibilityContext=WSSTabPersistence
// @match        http://governanca.mongeral.seguros/Lists/BaseProvisaoPagamento/0204%20%20Proviso%20%20A%20pagar%20%20Por%20data%20limite%20para%20lanar%20PAGNET%202024%20e%20Por%20responsvel.aspx
// @run-at       document-idle
// @all-frames   true
// @grant        none
// ==/UserScript==
(function () {
  "use strict";

  // select id
  const idNF  = "#ctl00_m_g_dd0cce18_d6ca_4408_8dec_ee13516f0d19_ctl00_ListForm2_formFiller_FormView_ctl613_lookup94f201ca_cf75_4038_b660_f9b0504e87f5_Lookup";

  const idBox = "nfSearchWrap__onlyOne";
  const soNum   = (s) => (s || "").replace(/\D+/g, "");
  const esperar = (fn, ms = 150) => { let t; return (...a) => { clearTimeout(t); t = setTimeout(() => fn(...a), ms); }; };
  const disparar = (el) => {
    el.dispatchEvent(new Event("input",  { bubbles: true }));
    el.dispatchEvent(new Event("change", { bubbles: true }));
    if (window.jQuery) { try { window.jQuery(el).trigger("change"); } catch {} }
  };
  function porLabel(doc) {
    const nodes = [...doc.querySelectorAll("label, div, span, td, th")]
      .filter(n => (n.childElementCount ? n.innerText : n.textContent).trim() === "NF");
    for (const n of nodes) {
      const blocos = [
        n.closest(".nf-filler-control"),
        n.closest(".nf-filler-control-inner"),
        n.closest("tr"),
        n.parentElement,
        n.closest("div, section, fieldset")
      ].filter(Boolean);
      for (const b of blocos) {
        const mesmo = b.querySelector && b.querySelector("select");
        if (mesmo && mesmo.options?.length) return mesmo;

        const prox = b.nextElementSibling?.querySelector?.("select");
        if (prox && prox.options?.length) return prox;

        const up = b.parentElement;
        if (up) {
          const sib = up.querySelector("select");
          if (sib && sib.options?.length) return sib;
        }
      }
    }
    return null;
  }
  function porNum(doc) {
    const sec = [...doc.querySelectorAll("div,h2,h3,legend")]
      .find(e => /emiss[aã]o.*Nota Fiscal/i.test((e.textContent || "")));
    if (sec) {
      const box = sec.closest("div, section, fieldset") || sec.parentElement;
      if (box) {
        const sel = box.querySelector("select");
        if (sel) return sel;
      }
    }
    const bigs = [...doc.querySelectorAll("select")].filter(s =>
      [...s.options].some(o => /\d{6,}/.test((o.textContent || ""))));
    return bigs.at(-1) || null;
  }
  function selNF(doc) {
    return (
      doc.querySelector(idNF) ||
      porLabel(doc) ||
      porNum(doc)
    );
  }
  function criaBox(doc, sel) {
    if (!sel) return;

    const antigo = doc.getElementById(idBox);
    if (antigo) antigo.remove();

    const box = doc.createElement("div");
    box.id = idBox;
    box.style.display = "flex";
    box.style.gap = "6px";
    box.style.alignItems = "center";
    box.style.margin = "6px 0";

    const inp = doc.createElement("input");
    inp.type = "text";
    inp.placeholder = "Digite a NF… (só números)";
    inp.style.padding = "6px 8px";
    inp.style.width = "260px";
    inp.style.border = "1px solid #bbb";
    inp.style.borderRadius = "6px";
    const info = doc.createElement("span");
    info.style.fontSize = "12px";
    info.style.opacity = "0.75";
    box.appendChild(inp);
    box.appendChild(info);
    sel.parentNode.insertBefore(box, sel);

    let achados = [], idx = 0;

    function busca(qNum) {
      const lista = [];
      for (const o of sel.options) {
        const t = soNum(o.textContent);
        const attr = o.getAttribute("data-nfchoicevalue") || "";
        const after = attr.split("#")[1] || "";
        const a = soNum(after);
        const hay = [t, a].filter(Boolean);
        const comeca = hay.some(h => h.startsWith(qNum));
        const contem = comeca || hay.some(h => h.includes(qNum));
        if (qNum && contem) lista.push({ opt: o, score: comeca ? 0 : 1 });
      }
      lista.sort((a, b) => a.score - b.score);
      return lista.map(x => x.opt);
    }
    const atualizar = esperar(() => {
      const q = soNum(inp.value);
      if (!q) { achados = []; info.textContent = ""; return; }
      achados = busca(q); idx = 0;
      info.textContent = achados.length ? `${achados.length} resultado(s)` : "sem resultados";
      if (achados.length) {
        sel.value = achados[0].value;
        disparar(sel);
        sel.scrollIntoView({ block: "nearest" });
      }
    }, 120);
    inp.addEventListener("input", atualizar);
    inp.addEventListener("paste", () => setTimeout(atualizar, 0));
    inp.addEventListener("keydown", (e) => {
      if (!achados.length) return;
      if (e.key === "ArrowDown") { e.preventDefault(); idx = (idx + 1) % achados.length; }
      else if (e.key === "ArrowUp") { e.preventDefault(); idx = (idx - 1 + achados.length) % achados.length; }
      else if (e.key !== "Enter") { return; }
      sel.value = achados[idx].value;
      disparar(sel);
    });

    const obsSel = new MutationObserver(() => atualizar());
    obsSel.observe(sel, { childList: true, subtree: true });
  }
  function injeta(doc) {
    try {
      const sel = selNF(doc);
      if (sel) criaBox(doc, sel);
    } catch {}
  }
  injeta(document);
  const obsDoc = new MutationObserver(() => {
    injeta(document);
    document.querySelectorAll("iframe").forEach(ifr => {
      try {
        const doc2 = ifr.contentDocument || ifr.contentWindow?.document;
        if (doc2) injeta(doc2);
      } catch {}
    });
  });
  obsDoc.observe(document.documentElement, { childList: true, subtree: true });
  setInterval(() => injeta(document), 1500);
})();