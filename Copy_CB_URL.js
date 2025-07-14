// ==UserScript==
// @name         Copy Cyborg URL
// @namespace    http://tampermonkey.net/
// @version      1.2
// @description  
// @author       Luchito & Dilan
// @match        https://cyborg.deckard.com/listing/*/STR*
// @grant        none
// ==/UserScript==

(function () {
    function getIdentifier() {
        const pathSegments = window.location.pathname.split("/");
        return pathSegments[pathSegments.length - 1]; // Último segmento
    }

    function copyToClipboard(text) {
        navigator.clipboard.writeText(text).then(() => {
            alert("Copiado: " + text);
        }).catch(err => {
            console.error("Error al copiar:", err);
        });
    }

    function createButton() {
        const existingButton = document.getElementById("copyIdentifierButton");
        if (existingButton) return; // Evitar duplicados

        const button = document.createElement("button");
        button.id = "copyIdentifierButton";
        button.innerText = "Copiar ID";
        button.style.position = "fixed";
        button.style.bottom = "20px";
        button.style.right = "20px";
        button.style.padding = "10px";
        button.style.backgroundColor = "#007bff";
        button.style.color = "#fff";
        button.style.border = "none";
        button.style.borderRadius = "5px";
        button.style.cursor = "pointer";
        button.style.boxShadow = "2px 2px 10px rgba(0, 0, 0, 0.3)";
        button.onclick = () => copyToClipboard(getIdentifier());

        document.body.appendChild(button);
    }

    if (window.location.href.includes("/listing/")) {
        createButton();
    }
})();
