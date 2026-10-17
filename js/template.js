/* ═══════════════════════════════════════════════════════════════════
   BIBLIOTECA TEMPLATE — template.js
   Carica i partial condivisi (head, header, footer) in ogni pagina
   e inizializza i componenti interattivi.
════════════════════════════════════════════════════════════════════ */

(function () {
  'use strict';

  /* ─── Caricamento partial ─────────────────────────────────────── */

  /**
   * Recupera un file HTML e ne inserisce il contenuto
   * nell'elemento identificato da `idContenitore`.
   *
   * @param {string} percorso      - Percorso del file partial
   * @param {string} idContenitore - ID dell'elemento destinazione
   * @param {Function} [callback]  - Funzione chiamata dopo l'inserimento
   */
  function caricaPartial(percorso, idContenitore, callback) {
    var contenitore = document.getElementById(idContenitore);
    if (!contenitore) return;

    fetch(percorso)
      .then(function (risposta) {
        if (!risposta.ok) {
          throw new Error('Errore nel caricamento di ' + percorso);
        }
        return risposta.text();
      })
      .then(function (html) {
        contenitore.innerHTML = html;
        if (typeof callback === 'function') callback();
      })
      .catch(function (errore) {
        console.error('[template.js]', errore);
      });
  }

  /* ─── Tendina di navigazione ──────────────────────────────────── */

  /**
   * Inizializza la tendina dopo che l'header è stato iniettato nel DOM.
   * Deve essere chiamata come callback di caricaPartial('header').
   */
  function inizializzaTendina() {
    var tendina  = document.getElementById('tendina-nav');
    var trigger  = document.getElementById('trigger-nav');
    var pannello = document.getElementById('pannello-nav');

    if (!tendina || !trigger || !pannello) return;

    function impostaAperta(aperta) {
      tendina.classList.toggle('nav-tendina--aperta', aperta);
      trigger.setAttribute('aria-expanded', String(aperta));
    }

    /* Click sul trigger */
    trigger.addEventListener('click', function () {
      var eraAperta = tendina.classList.contains('nav-tendina--aperta');
      impostaAperta(!eraAperta);
    });

    /* Chiude cliccando fuori */
    document.addEventListener('click', function (evento) {
      if (!tendina.contains(evento.target)) {
        impostaAperta(false);
      }
    });

    /* Chiude con Escape */
    document.addEventListener('keydown', function (evento) {
      if (evento.key === 'Escape') {
        impostaAperta(false);
        trigger.focus();
      }
    });
  }

  /* ─── Avvio ───────────────────────────────────────────────────── */

  /* I partial vengono caricati in parallelo dove possibile.
     La tendina viene inizializzata solo dopo che l'header è pronto. */
  caricaPartial('includes/head.html',    'inc-head');
  caricaPartial('includes/header.html',  'inc-header', inizializzaTendina);
  caricaPartial('includes/footer.html',  'inc-footer');

})();