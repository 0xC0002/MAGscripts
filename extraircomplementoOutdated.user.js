// ==UserScript==
// @name         exportar rateio p/ pagnet
// @namespace    brunoAF
// @version      1.1
// @description  injeta um botao na ribbon ao lado de "Espelho Pagnet" que extrai o completemento da provisao
// @match        http://governanca.mongeral.seguros/Lists/BaseProvisaoPagamento/*
// @grant        none
// @run-at       document-idle
// @require      https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js
// ==/UserScript==

(function () {
  'use strict';
  const USAR_VALOR_PAGO = false;
  const WEBPART_TITULO_ALVO = "Provisão de Pagamento: Complemento (CC/SP)";
  const ID_BTN = 'btnExportarRateioPagNetRibbon';

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
  function buildXLSX(){
    const d = getDoc();
    const table = findComplementoTable(d);
    if (!table){ alert("Tabela 'Complemento (CC/SP)' não encontrada."); return; }

    const ths = table.querySelectorAll("tr.ms-viewheadertr th, tr.ms-viewheadertr td");
    const col = {};
    ths.forEach((th, idx) => {
      const t = norm(th.innerText || th.textContent || "");
      if (t.includes("FILIALTCODE1")) col.FILIAL = idx;
      if (t.includes("CENTRODECUSTO")) col.CC = idx;
      if (t.includes("VALORRATEADOPREVISTO")) col.PREV = idx;
      if (t.includes("VALORRATEADOPAGO"))     col.PAGO = idx;
      if (t === "PROVISAO") col.PROV = idx;
      if (t === "ID")       col.ID = idx;
    });
    const idxValor = (USAR_VALOR_PAGO && col.PAGO!=null) ? col.PAGO : col.PREV;
    if (col.FILIAL==null || col.CC==null || idxValor==null){
      alert("nao foi possível mapear FILIAL/CENTRO/VALOR nos cabecalhos."); return;
    }
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
    XLSX.writeFile(wb, `rateio_${ref}.xlsx`);
  }

  function findEspelhoButton() {
    const anchors = Array.from(document.querySelectorAll('a.ms-cui-ctl-large'));
    let btn = anchors.find(a => {
      const label = a.querySelector('.ms-cui-ctl-largelabel');
      return label && norm(label.innerText) === 'ESPELHO PAGNET';
    });
    if (btn) return btn;

    btn = anchors.find(a => {
      const label = a.querySelector('.ms-cui-ctl-largelabel');
      const t = norm(label?.innerText || '');
      return t.includes('ESPELHO') && t.includes('PAGNET');
    });
    if (btn) return btn;
    btn = anchors.find(a => norm(a.querySelector('.ms-cui-ctl-largelabel')?.innerText || a.textContent) === 'SALVAR');
    return btn || null;
  }
  // botao ribbon
  function criarBotao(){
    const dups = document.querySelectorAll('#' + CSS.escape(ID_BTN));
    dups.forEach((n, i) => { if (i > 0) n.remove(); });
    if (document.getElementById(ID_BTN)) return;

    const botaoAlvo = findEspelhoButton();
    if (!botaoAlvo) return;

    const a = document.createElement('a');
    a.id = ID_BTN;
    a.className = 'ms-cui-ctl-large';
    a.href = 'javascript:void(0);';
    a.role = 'button';
    a.style.marginLeft = '6px';
    a.innerHTML = `
      <span class="ms-cui-ctl-largeIconContainer">
        <span class="ms-cui-img-32by32 ms-cui-img-cont-float">
          <img alt="" src="/_layouts/1046/images/formatmap32x32.png" style="top:-320px; left:-192px;">
        </span>
      </span>
      <span class="ms-cui-ctl-largelabel">Exportar<br>Rateio p/ PagNet</span>
    `;
    a.addEventListener('click', buildXLSX);
    botaoAlvo.parentElement.insertBefore(a, botaoAlvo.nextSibling);
  }
  function observarRibbon(){
    const obs = new MutationObserver(() => {
      const ok = !!findEspelhoButton();
      if (ok){
        criarBotao();
      }
    });
    obs.observe(document.body, { childList:true, subtree:true });
    // tentativa inicial
    criarBotao();
  }
  window.addEventListener('load', observarRibbon);
})();
