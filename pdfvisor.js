/* Renderiza un PDF a <canvas> con pdf.js, sin iframes ni visores del navegador
   (los hosts de Outlook bloquean la navegación a blob: dentro del panel). */
"use strict";

(function () {
  let libPromise = null;

  function cargarLib() {
    if (!libPromise) {
      const base = new URL("pdfjs/", document.currentScript ? document.currentScript.src : location.href).href;
      libPromise = import(base + "pdf.min.mjs").then((lib) => {
        lib.GlobalWorkerOptions.workerSrc = base + "pdf.worker.min.mjs";
        return lib;
      });
    }
    return libPromise;
  }

  /**
   * @param {Blob} blob            PDF
   * @param {HTMLElement} cont     contenedor; se vacía y se llena de <canvas>, uno por página
   * @param {object} [op]          { ancho: px disponibles, maxPaginas, onPagina(i, total) }
   * @returns {Promise<number>}    número total de páginas del documento
   */
  window.renderizarPdf = async function (blob, cont, op) {
    op = op || {};
    const lib = await cargarLib();
    const data = new Uint8Array(await blob.arrayBuffer());
    const doc = await lib.getDocument({ data }).promise;
    const total = doc.numPages;
    const limite = Math.min(total, op.maxPaginas || total);
    const ancho = Math.max(120, op.ancho || cont.clientWidth || 320);
    const dpr = Math.min(window.devicePixelRatio || 1, 2);

    cont.innerHTML = "";
    for (let i = 1; i <= limite; i++) {
      if (op.cancelado && op.cancelado()) break;
      const page = await doc.getPage(i);
      const base = page.getViewport({ scale: 1 });
      const vp = page.getViewport({ scale: ancho / base.width });
      const canvas = document.createElement("canvas");
      canvas.width = Math.floor(vp.width * dpr);
      canvas.height = Math.floor(vp.height * dpr);
      canvas.style.width = Math.floor(vp.width) + "px";
      canvas.style.height = Math.floor(vp.height) + "px";
      canvas.setAttribute("aria-label", "Página " + i + " de " + total);
      cont.appendChild(canvas);
      await page.render({
        canvas,
        viewport: vp,
        transform: dpr !== 1 ? [dpr, 0, 0, dpr, 0, 0] : null,
      }).promise;
      if (op.onPagina) op.onPagina(i, total);
    }
    return total;
  };
})();
