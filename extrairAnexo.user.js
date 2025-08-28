// ==UserScript==
// @name         extrair arquivos da provisao
// @namespace    http://tampermonkey.net/
// @version      1.0
// @description  o ID é a contagem - Alterar a URL do @match para a URL da sua fila a pagar
// @match        http://governanca.mongeral.seguros/Lists/BaseProvisaoPagamento/0204%20%20Proviso%20%20A%20pagar%20%20Por%20data%20limite%20para%20lanar%20PAGNET%202024%20e%20Por%20responsvel.aspx
// @grant        GM_download
// ==/UserScript==

(function () {
    'use strict';
    function injetarRibbon() {
        const grupos = document.querySelectorAll('.ms-cui-group');
        for (const grupo of grupos) {
            const tituloGrupo = grupo.querySelector('.ms-cui-groupTitle');
            if (tituloGrupo && tituloGrupo.textContent.trim() === 'Ações') {
                const container = grupo.querySelector('.ms-cui-row-onerow, .ms-cui-row');
                if (!container || container.querySelector('#btnBaixarRibbon')) return;
                const botao = document.createElement('a');
                botao.id = 'btnBaixarRibbon';
                botao.href = 'javascript:;';
                botao.role = 'button';
                botao.className = 'ms-cui-ctl-large';
                botao.style.marginLeft = '6px';
                botao.innerHTML = `
                    <span class="ms-cui-ctl-largeIconContainer">
                        <span class="ms-cui-img-32by32 ms-cui-img-cont-float">
                            <img alt="" src="/_layouts/1046/images/formatmap32x32.png" style="top: -320px; left: -192px;">
                        </span>
                    </span>
                    <span class="ms-cui-ctl-largelabel">Baixar<br>Anexos</span>
                `;
                botao.addEventListener('click', async () => {
                    const idProvisao = prompt("digite o número da contagem da provisão (ex: 55824):");
                    if (!idProvisao || !/^\d+$/.test(idProvisao)) {
                        return;
                    }
                    const linkProvisao = Array.from(document.querySelectorAll('a[href*="PageType=4"]'))
                        .find(a => a.href.includes(`ID=${idProvisao}`));

                    if (!linkProvisao) {
                        alert(`link da provisão com ID=${idProvisao} não encontrado.`);
                        return;
                    }

                    try {
                        const res = await fetch(linkProvisao.href);
                        const html = await res.text();
                        const parser = new DOMParser();
                        const doc = parser.parseFromString(html, "text/html");
                        const links = doc.querySelectorAll("#idAttachmentsTable a[href]");

                        if (!links.length) {
                            alert("nenhum anexo encontrado.");
                            return;
                        }   
                        for (const a of links) {
                            const fileUrl = a.href.startsWith("http")
                                ? a.href
                                : 'http://governanca.mongeral.seguros' + a.getAttribute("href");

                            const nomeArquivo = decodeURIComponent(fileUrl.split('/').pop());
                            console.log(`baixando: ${nomeArquivo}`);
                            GM_download(fileUrl, nomeArquivo);
                        }

                    } catch (e) {
                        console.error(e);
                        alert("erro ao baixar anexos");
                    }
                });
                container.appendChild(botao);
                return;
            }
        }
    }
    const observer = new MutationObserver(() => {
        injetarRibbon();
    });
    observer.observe(document.body, { childList: true, subtree: true });
    setTimeout(injetarRibbon, 3000);
})();
