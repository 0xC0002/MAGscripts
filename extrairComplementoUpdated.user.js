// ==UserScript==
// @name         exportar rateio pago e previsto
// @namespace    brunoAF
// @version      1.3
// @description  injeta dois botoes na ribbon: exportar (pago) e exportar (previsto)
// @match        http://governanca.mongeral.seguros/Lists/BaseProvisaoPagamento/*
// @grant        none
// @run-at       document-idle
// @require      https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js
// ==/UserScript==

(function () {
  'use strict';

  const WEBPART_TITULO_ALVO = "Provisão de Pagamento: Complemento (CC/SP)";
  const ID_BTN_PAGO     = 'btnExportarRateioPagNet_Pago';
  const ID_BTN_PREV     = 'btnExportarRateioPagNet_Prev';
  const ICON_SPRITE     = "/_layouts/1046/images/formatmap32x32.png";

  const PAGNET_HEADER = [
    "RAMO (TCODE0)","FILIAL(TCODE1)","CENTRO_CUSTO (TCODE2)",
    "CLIENTE/BENEFICIARIO (TCODE3)","FORNECEDOR/CORRETOR (TCODE4)","HIST_PADRAO (TCODE5)",
    "TCODE6","TCODE7","TCODE8","TCODE9","VLR_RATEIO"
  ];

  const norm = s => String(s||"").replace(/\s+/g," ").replace(/\n+/g," ").trim().toUpperCase();
  const parseMoney = s => {
    if (typeof s !== "string") return Number(s)||0;
    const n = Number(s.replace(/\s/g,"").replace(/\./g,"").replace(",","."));
    return Number.isFinite(n) ? n : 0;
  };

  const getDoc = () => {
    const dlg = document.querySelector("iframe.ms-dlgFrame, iframe[class*='DlgFrame']") || null;
    if (dlg?.contentWindow?.document) return dlg.contentWindow.document;
    return document;
  };

  function findComplementoTable(d){
    const wpHeader = Array.from(d.querySelectorAll(".ms-WPHeaderTd h3.ms-WPTitle, #WebPartTitleWPQ3"));
    const alvoHdr = wpHeader.find(h => norm(h.textContent).includes(norm(WEBPART_TITULO_ALVO)));
    if (alvoHdr){
      const wp = alvoHdr.closest("table")?.parentElement?.parentElement?.parentElement;
      const t  = wp?.querySelector("table.ms-listviewtable");
      if (t) return t;
    }
    return Array.from(d.querySelectorAll("table.ms-listviewtable")).find(tb=>{
      const headerText = norm(tb.querySelector("tr.ms-viewheadertr")?.innerText || tb.innerText || "");
      return headerText.includes("FILIALTCODE1") && headerText.includes("CENTRODECUSTO");
    }) || null;
  }

  // usarPago = true -> ValorRateadoPago; false -> ValorRateadoPrevisto
  function buildXLSX(usarPago){
    const d = getDoc();
    const table = findComplementoTable(d);
    if (!table){ alert("Tabela 'Complemento (CC/SP)' não encontrada."); return; }

    // mapear cabecalhos
    const ths = table.querySelectorAll("tr.ms-viewheadertr th, tr.ms-viewheadertr td");
    const col = {};
    ths.forEach((th, idx) => {
      const t = norm(th.innerText || th.textContent || "");
      if (t.includes("FILIALTCODE1"))         col.FILIAL = idx;
      if (t.includes("CENTRODECUSTO"))        col.CC     = idx;
      if (t.includes("VALORRATEADOPREVISTO")) col.PREV   = idx;
      if (t.includes("VALORRATEADOPAGO"))     col.PAGO   = idx;
      if (t === "PROVISAO")                   col.PROV   = idx;
      if (t === "ID")                         col.ID     = idx;
    });

    const idxValor = usarPago
      ? (col.PAGO ?? col.PREV)   // pref pago
      : (col.PREV ?? col.PAGO);  // pref previsto

    if (col.FILIAL==null || col.CC==null || idxValor==null){
      alert("não foi possível mapear FILIAL/CENTRO/VALOR nos cabeçalhos."); return;
    }

    // linhas uteis
    const trs = Array.from(table.querySelectorAll("tbody > tr"))
      .filter(tr => !tr.classList.contains("ms-viewheadertr"))
      .filter(tr => !/\bSOMA=/.test(norm(tr.innerText)));

    const out = [PAGNET_HEADER.slice()];
    trs.forEach(tr=>{
      const tds = Array.from(tr.children);
      const filial = (tds[col.FILIAL]?.innerText || "").trim();
      const cc     = (tds[col.CC]?.innerText || "").trim();
      const vStr   = (tds[idxValor]?.innerText || "").trim();
      const valor  = parseMoney(vStr);
      if (!filial && !cc && !valor) return;
      out.push(["", filial, cc, "", "", "", "", "", "", "", valor]);
    });

    if (out.length === 1){ alert("nenhuma linha encontrada para exportar."); return; }

    // nome do arquivo
    let ref = "complemento";
    if (trs[0]){
      const t0 = Array.from(trs[0].children);
      const prov = col.PROV!=null ? (t0[col.PROV]?.innerText || "").trim() : "";
      const id   = col.ID!=null   ? (t0[col.ID]?.innerText   || "").trim() : "";
      ref = (prov || id || ref).replace(/[^\w-]+/g,"_").slice(0,40) || ref;
    }

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(out);
    XLSX.utils.book_append_sheet(wb, ws, "Rateio");
    XLSX.writeFile(wb, `rateio_${usarPago ? 'pago' : 'previsto'}_${ref}.xlsx`);
  }
  function normLabelFromAnchor(a){
    const label = a.querySelector('.ms-cui-ctl-largelabel');
    return norm(label?.innerText || a.textContent || '');
  }
  function findEspelhoButton() {
    const anchors = Array.from(document.querySelectorAll('a.ms-cui-ctl-large'));
    return anchors.find(a => {
      const t = normLabelFromAnchor(a);
      return t.includes('ESPELHO') && t.includes('PAGNET');
    }) || anchors.find(a => normLabelFromAnchor(a) === 'SALVAR') || null;
  }

  function makeButton(id, labelHtml, clickHandler){
    const a = document.createElement('a');
    a.id = id;
    a.className = 'ms-cui-ctl-large';
    a.href = 'javascript:void(0);';
    a.role = 'button';
    a.style.marginLeft = '6px';
    a.innerHTML = `
      <span class="ms-cui-ctl-largeIconContainer">
        <span class="ms-cui-img-32by32 ms-cui-img-cont-float">
          <img alt="" src="${ICON_SPRITE}" style="top:-320px; left:-192px;">
        </span>
      </span>
      <span class="ms-cui-ctl-largelabel">${labelHtml}</span>
    `;
    a.addEventListener('click', clickHandler);
    return a;
  }

  function criarBotoes(){
    // evitar duplicados
    [ID_BTN_PAGO, ID_BTN_PREV].forEach(id => {
      const dups = document.querySelectorAll('#' + CSS.escape(id));
      dups.forEach((n, i) => { if (i > 0) n.remove(); });
    });
    if (document.getElementById(ID_BTN_PAGO) && document.getElementById(ID_BTN_PREV)) return;

    const botaoAlvo = findEspelhoButton();
    if (!botaoAlvo) return;

    const btnPago = makeButton(ID_BTN_PAGO, 'Exportar<br>(pago)',    () => buildXLSX(true));
    const btnPrev = makeButton(ID_BTN_PREV, 'Exportar<br>(previsto)',() => buildXLSX(false));

    botaoAlvo.parentElement.insertBefore(btnPago, botaoAlvo.nextSibling);
    botaoAlvo.parentElement.insertBefore(btnPrev, btnPago.nextSibling);
  }

  function observarRibbon(){
    const obs = new MutationObserver(() => {
      if (findEspelhoButton()){
        criarBotoes();
      }
    });
    obs.observe(document.body, { childList:true, subtree:true });
    criarBotoes();
  }

  window.addEventListener('load', observarRibbon);
})();
