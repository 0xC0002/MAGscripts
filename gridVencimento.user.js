// ==UserScript==
// @name         pagNET - vencimento na listagem dos titulos
// @namespace    http://tampermonkey.net/
// @version      v0.0.3
// @description  coluna de vencimento no pagNet
// @author       brunoMAG
// @match        https://pagnet.mongeral.seguros/CP/BuscaTitulo.aspx
// @icon         https://www.google.com/s2/favicons?sz=64&domain=mongeral.com
// @grant        GM_xmlhttpRequest
// @connect      pagnet.mongeral.seguros
// ==/UserScript==

(function() {
    'use strict';

    const headerRow = document.querySelector('.GridBuscaHeaderStyle');
    const vencHeader = document.createElement('th');
    vencHeader.innerText = 'Vencimento';
    headerRow.appendChild(vencHeader);

    const linhas = Array.from(document.querySelectorAll('tr.GridBuscaRowStyle, tr.GridBuscaAlternatingRowStyle'));

    linhas.forEach(linha => {
        const cell = document.createElement('td');
        cell.innerText = '...';
        linha.appendChild(cell);

        const link = linha.querySelector('a[href*="Mantitulo.aspx"]');
        if (link) {
            const url = "https://pagnet.mongeral.seguros/CP/" + link.getAttribute('href');

            GM_xmlhttpRequest({
                method: "GET",
                url: url,
                onload: function(response) {
                    const parser = new DOMParser();
                    const doc = parser.parseFromString(response.responseText, "text/html");
                    const vencimentoInput = doc.querySelector('#txtd_Venc');
                    const venc = vencimentoInput ? vencimentoInput.value.trim() : "N/A";
                    cell.innerText = venc;
                },
                onerror: function() {
                    cell.innerText = "Erro";
                }
            });
        } else {
            cell.innerText = "N/A";
        }
    });

})();
