/* ═══════════════════════════════════════════════════════════════════
   BIBLIOTECA TEMPLATE — Codice.gs
   File principale Google Apps Script.
   Gestisce il routing, l'autenticazione e le funzioni server.
════════════════════════════════════════════════════════════════════ */

'use strict';

/* ─── Configurazione ────────────────────────────────────────────── */

var DOMINIO_AUTORIZZATO  = 'DOMINIOSCUOLA.education';
var UTENTE_BIBLIOTECARIO = 'UTENTEADMIN@DOMINIOSCUOLA.education';
var ID_FOGLIO            = 'GOOGLE_SHEET_ID';

/* ─── Punto di ingresso ─────────────────────────────────────────── */

/**
 * Risponde alle richieste GET alla web app.
 * Verifica che l'utente appartenga al dominio autorizzato,
 * poi serve la pagina richiesta.
 *
 * @param {Object} e - Oggetto evento GAS (e.parameter.pagina)
 * @returns {HtmlOutput}
 */
function doGet(e) {
  var utente = Session.getActiveUser().getEmail();

  /* Verifica dominio */
  if (!utente.endsWith('@' + DOMINIO_AUTORIZZATO)) {
    return HtmlService
      .createHtmlOutput('<p>Accesso non autorizzato.</p>')
      .setTitle('Accesso negato');
  }

  /* Routing */
  var pagina = (e.parameter && e.parameter.pagina) ? e.parameter.pagina : 'index';

  var paginheConsentite = [
    'index',
    'catalogo',
    'prestiti',
    'restituzione',
    'impostazioni',
    'gestione',
    'utenti',
    'guida'
  ];

  if (paginheConsentite.indexOf(pagina) === -1) {
    pagina = 'index';
  }

  /* Pagine riservate al bibliotecario */
  var paginheRiservate = ['impostazioni', 'gestione'];

  if (paginheRiservate.indexOf(pagina) !== -1 && utente !== UTENTE_BIBLIOTECARIO) {
    return HtmlService
      .createHtmlOutput('<p>Accesso riservato.</p>')
      .setTitle('Accesso negato');
  }

  /* Preparazione dati per i template che ne hanno bisogno */
  var template = HtmlService.createTemplateFromFile(pagina);

  if (pagina === 'catalogo') {
    template.datiCatalogo = leggiCatalogo();
  }

  if (pagina === 'prestiti') {
    template.datiPrestiti = {
      libri:  leggiLibriDisponibili(),
      classi: leggiClassi(),
      durate: leggiDurate(),
      anno:   leggiAnno()
    };
  }

  if (pagina === 'restituzione') {
    template.datiRestituzione = leggiLibriInPrestito();
  }

  if (pagina === 'utenti') {
    template.datiUtenti = {
      prestiti: leggiLibriInPrestito(),
      durate:   leggiDurate()
    };
  }

  if (pagina === 'gestione') {
    template.datiGestione = {
      catalogo: leggiCatalogo(),
      scaffali: leggiScaffali()
    };
  }

  if (pagina === 'impostazioni') {
    template.datiImpostazioni = leggiImpostazioni();
  }

  return template
    .evaluate()
    .setTitle('Biblioteca Scolastica')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

/* ─── Inclusione partial ────────────────────────────────────────── */

/**
 * Restituisce il contenuto di un file HTML del progetto.
 * Usata nei template con la sintassi: <?!= include('nomeFile') ?>
 *
 * @param {string} nomeFile - Nome del file senza estensione .html
 * @returns {string}
 */
function include(nomeFile) {
  return HtmlService.createHtmlOutputFromFile(nomeFile).getContent();
}

/* ─── Catalogo ──────────────────────────────────────────────────── */

/**
 * Legge tutte le righe del foglio Google Sheets e restituisce
 * un array di oggetti libro, uno per riga (dalla riga 2 in poi).
 * Aggiunge il flag booleano `inPrestito` senza esporre i dati utente.
 *
 * @returns {Array<Object>}
 */
function leggiCatalogo() {
  var foglio = SpreadsheetApp
    .openById(ID_FOGLIO)
    .getSheets()[0];

  var ultimaRiga    = foglio.getLastRow();
  var ultimaColonna = foglio.getLastColumn();

  if (ultimaRiga < 2) return [];

  var righe = foglio.getRange(2, 1, ultimaRiga - 1, ultimaColonna).getValues();
  var catalogo = [];

  for (var i = 0; i < righe.length; i++) {
    var r = righe[i];
    if (!r[0] && !r[1]) continue;

    catalogo.push({
      etichetta:  String(r[0]  || '').trim(),
      titolo:     String(r[1]  || '').trim(),
      volume:     String(r[2]  || '').trim(),
      autore:     String(r[3]  || '').trim(),
      coautore:   String(r[4]  || '').trim(),
      scaffale:   String(r[7]  || '').trim(),
      genere:     String(r[8]  || '').trim(),
      note:       String(r[9]  || '').trim(),
      inPrestito: !!(r[5] || r[6])
    });
  }

  return catalogo;
}

/**
 * Restituisce solo i libri non in prestito (colonne F e G vuote).
 * Usata dalla pagina Prestiti.
 *
 * @returns {Array<Object>}
 */
function leggiLibriDisponibili() {
  var foglio = SpreadsheetApp
    .openById(ID_FOGLIO)
    .getSheets()[0];

  var ultimaRiga    = foglio.getLastRow();
  var ultimaColonna = foglio.getLastColumn();

  if (ultimaRiga < 2) return [];

  var righe = foglio.getRange(2, 1, ultimaRiga - 1, ultimaColonna).getValues();
  var disponibili = [];

  for (var i = 0; i < righe.length; i++) {
    var r = righe[i];
    if (!r[0] && !r[1]) continue;

    /* Salta i libri già in prestito */
    if (r[5] || r[6]) continue;

    disponibili.push({
      etichetta: String(r[0] || '').trim(),
      titolo:    String(r[1] || '').trim(),
      volume:    String(r[2] || '').trim(),
      autore:    String(r[3] || '').trim(),
      coautore:  String(r[4] || '').trim()
    });
  }

  return disponibili;
}

/* ─── Prestiti ──────────────────────────────────────────────────── */

/**
 * Registra un prestito scrivendo nelle colonne F, G, J, K
 * della riga corrispondente all'etichetta.
 *
 * @param {Object} dati - {etichetta, data, utente, tipo, note}
 * @returns {{ok: boolean, errore: string}}
 */
function registraPrestito(dati) {
  try {
    var foglio = SpreadsheetApp
      .openById(ID_FOGLIO)
      .getSheets()[0];

    var ultimaRiga = foglio.getLastRow();
    if (ultimaRiga < 2) return { ok: false, errore: 'Catalogo vuoto.' };

    var colEtichette = foglio.getRange(2, 1, ultimaRiga - 1, 1).getValues();
    var rigaTrovata  = -1;

    for (var i = 0; i < colEtichette.length; i++) {
      if (String(colEtichette[i][0]).trim() === dati.etichetta.trim()) {
        rigaTrovata = i + 2; /* +2: riga 1 = intestazioni, array 0-based */
        break;
      }
    }

    if (rigaTrovata === -1) {
      return { ok: false, errore: 'Libro non trovato: ' + dati.etichetta };
    }

    /* Verifica che il libro non sia già in prestito */
    var dataPrestito = foglio.getRange(rigaTrovata, 6).getValue();
    if (dataPrestito) {
      return { ok: false, errore: 'Il libro è già in prestito.' };
    }

    /* Scrive i dati del prestito */
    foglio.getRange(rigaTrovata, 6).setValue(dati.data);   /* F — Data prestito */
    foglio.getRange(rigaTrovata, 7).setValue(dati.utente); /* G — Utente        */
    foglio.getRange(rigaTrovata, 10).setValue(dati.note);  /* J — Note          */
    foglio.getRange(rigaTrovata, 11).setValue(dati.tipo);  /* K — Tipo prestito */

    return { ok: true };

  } catch (err) {
    return { ok: false, errore: err.message };
  }
}

/* ─── Configurazione — PropertiesService ───────────────────────── */

/**
 * Restituisce la configurazione completa delle classi.
 * @returns {Array<{nome: string, alunni: number}>}
 */
function leggiClassi() {
  var stringa = PropertiesService.getScriptProperties()
                  .getProperty('CLASSI_REGISTRATE');
  if (!stringa) return [];
  try { return JSON.parse(stringa); }
  catch (err) { return []; }
}

/**
 * Restituisce l'elenco degli scaffali.
 * @returns {Array<string>}
 */
function leggiScaffali() {
  var stringa = PropertiesService.getScriptProperties()
                  .getProperty('SCAFFALI_REGISTRATI');
  if (!stringa) return [];
  try { return JSON.parse(stringa); }
  catch (err) { return []; }
}

/**
 * Restituisce l'anno scolastico corrente.
 * @returns {string}
 */
function leggiAnno() {
  return PropertiesService.getScriptProperties()
           .getProperty('ANNO_SCOLASTICO') || '';
}

/**
 * Restituisce le durate dei prestiti.
 * @returns {{breve: number, standard: number, lungo: number}}
 */
function leggiDurate() {
  var stringa = PropertiesService.getScriptProperties()
                  .getProperty('DURATE_PRESTITO');
  if (!stringa) return { breve: 0, standard: 0, lungo: 0 };
  try { return JSON.parse(stringa); }
  catch (err) { return { breve: 0, standard: 0, lungo: 0 }; }
}

/**
 * Restituisce tutte le impostazioni in un unico oggetto.
 * @returns {{anno: string, durate: Object, classi: Array, scaffali: Array}}
 */
function leggiImpostazioni() {
  return {
    anno:     leggiAnno(),
    durate:   leggiDurate(),
    classi:   leggiClassi(),
    scaffali: leggiScaffali()
  };
}

/**
 * Salva tutte le impostazioni ricevute dal client.
 * Chiamata tramite google.script.run dalla pagina impostazioni.html.
 *
 * @param {Object} dati - {anno, durate, classi, scaffali}
 * @returns {boolean} true se il salvataggio è riuscito
 */
function salvaImpostazioni(dati) {
  try {
    var utente = Session.getActiveUser().getEmail();
    if (utente !== UTENTE_BIBLIOTECARIO) {
      throw new Error('Accesso non autorizzato.');
    }

    var props = PropertiesService.getScriptProperties();

    props.setProperty('ANNO_SCOLASTICO',     String(dati.anno || '').trim());
    props.setProperty('DURATE_PRESTITO',     JSON.stringify(dati.durate   || {}));
    props.setProperty('CLASSI_REGISTRATE',   JSON.stringify(dati.classi   || []));
    props.setProperty('SCAFFALI_REGISTRATI', JSON.stringify(dati.scaffali || []));

    return true;
  } catch (err) {
    return false;
  }
}

/* ─── Restituzione ──────────────────────────────────────────────── */

/**
 * Restituisce solo i libri attualmente in prestito (F e G compilati).
 * Include i dati del prestito per mostrarli nella scheda.
 *
 * @returns {Array<Object>}
 */
function leggiLibriInPrestito() {
  var foglio = SpreadsheetApp
    .openById(ID_FOGLIO)
    .getSheets()[0];

  var ultimaRiga    = foglio.getLastRow();
  var ultimaColonna = foglio.getLastColumn();

  if (ultimaRiga < 2) return [];

  var righe = foglio.getRange(2, 1, ultimaRiga - 1, ultimaColonna).getValues();
  var inPrestito = [];

  for (var i = 0; i < righe.length; i++) {
    var r = righe[i];
    if (!r[0] && !r[1]) continue;

    /* Solo i libri con Data prestito o Utente compilati */
    if (!r[5] && !r[6]) continue;

    inPrestito.push({
      etichetta:    String(r[0]  || '').trim(),
      titolo:       String(r[1]  || '').trim(),
      volume:       String(r[2]  || '').trim(),
      autore:       String(r[3]  || '').trim(),
      coautore:     String(r[4]  || '').trim(),
      dataPrestito: r[5] ? Utilities.formatDate(new Date(r[5]), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
      utente:       String(r[6]  || '').trim(),
      note:         String(r[9]  || '').trim(),
      tipo:         String(r[10] || '').trim()
    });
  }

  return inPrestito;
}

/**
 * Registra una restituzione: azzera F, G e K, aggiorna J con le note.
 * Chiamata tramite google.script.run dalla pagina restituzione.html.
 *
 * @param {Object} dati - {etichetta, note}
 * @returns {{ok: boolean, errore: string}}
 */
function registraRestituzione(dati) {
  try {
    var foglio = SpreadsheetApp
      .openById(ID_FOGLIO)
      .getSheets()[0];

    var ultimaRiga = foglio.getLastRow();
    if (ultimaRiga < 2) return { ok: false, errore: 'Catalogo vuoto.' };

    var colEtichette = foglio.getRange(2, 1, ultimaRiga - 1, 1).getValues();
    var rigaTrovata  = -1;

    for (var i = 0; i < colEtichette.length; i++) {
      if (String(colEtichette[i][0]).trim() === dati.etichetta.trim()) {
        rigaTrovata = i + 2;
        break;
      }
    }

    if (rigaTrovata === -1) {
      return { ok: false, errore: 'Libro non trovato: ' + dati.etichetta };
    }

    /* Azzera Data prestito (F), Utente (G) e Tipo prestito (K) */
    foglio.getRange(rigaTrovata, 6).clearContent();   /* F — Data prestito  */
    foglio.getRange(rigaTrovata, 7).clearContent();   /* G — Utente         */
    foglio.getRange(rigaTrovata, 11).clearContent();  /* K — Tipo prestito  */

    /* Aggiorna le note (J) con il valore ricevuto */
    foglio.getRange(rigaTrovata, 10).setValue(dati.note || '');

    return { ok: true };

  } catch (err) {
    return { ok: false, errore: err.message };
  }
}

/* ─── Gestione libri ────────────────────────────────────────────── */

/**
 * Inserisce un nuovo libro in fondo al foglio.
 * @param {Object} dati
 * @returns {{ok: boolean, errore: string}}
 */
function inserisciLibro(dati) {
  try {
    var utente = Session.getActiveUser().getEmail();
    if (utente !== UTENTE_BIBLIOTECARIO) throw new Error('Accesso non autorizzato.');

    var foglio = SpreadsheetApp.openById(ID_FOGLIO).getSheets()[0];

    /* Verifica che l'etichetta non esista già */
    var ultimaRiga = foglio.getLastRow();
    if (ultimaRiga >= 2) {
      var etichette = foglio.getRange(2, 1, ultimaRiga - 1, 1).getValues();
      for (var i = 0; i < etichette.length; i++) {
        if (String(etichette[i][0]).trim() === dati.etichetta.trim()) {
          return { ok: false, errore: 'Etichetta già presente: ' + dati.etichetta };
        }
      }
    }

    foglio.appendRow([
      dati.etichetta, dati.titolo, dati.volume,
      dati.autore,    dati.coautore,
      '',             '',               /* F e G: Data prestito e Utente vuoti */
      dati.scaffale,  dati.genere, dati.note, ''  /* K: Tipo prestito vuoto */
    ]);

    return { ok: true };
  } catch (err) {
    return { ok: false, errore: err.message };
  }
}

/**
 * Modifica i campi di un libro esistente.
 * Non tocca mai Data prestito (F), Utente (G) e Tipo prestito (K).
 * @param {Object} dati
 * @returns {{ok: boolean, errore: string}}
 */
function modificaLibro(dati) {
  try {
    var utente = Session.getActiveUser().getEmail();
    if (utente !== UTENTE_BIBLIOTECARIO) throw new Error('Accesso non autorizzato.');

    var foglio     = SpreadsheetApp.openById(ID_FOGLIO).getSheets()[0];
    var ultimaRiga = foglio.getLastRow();
    if (ultimaRiga < 2) return { ok: false, errore: 'Catalogo vuoto.' };

    var etichette   = foglio.getRange(2, 1, ultimaRiga - 1, 1).getValues();
    var rigaTrovata = -1;

    for (var i = 0; i < etichette.length; i++) {
      if (String(etichette[i][0]).trim() === dati.etichetta.trim()) {
        rigaTrovata = i + 2;
        break;
      }
    }

    if (rigaTrovata === -1) return { ok: false, errore: 'Libro non trovato.' };

    /* Aggiorna solo le colonne modificabili */
    foglio.getRange(rigaTrovata, 2).setValue(dati.titolo);    /* B */
    foglio.getRange(rigaTrovata, 3).setValue(dati.volume);    /* C */
    foglio.getRange(rigaTrovata, 4).setValue(dati.autore);    /* D */
    foglio.getRange(rigaTrovata, 5).setValue(dati.coautore);  /* E */
    foglio.getRange(rigaTrovata, 8).setValue(dati.scaffale);  /* H */
    foglio.getRange(rigaTrovata, 9).setValue(dati.genere);    /* I */
    foglio.getRange(rigaTrovata, 10).setValue(dati.note);     /* J */

    return { ok: true };
  } catch (err) {
    return { ok: false, errore: err.message };
  }
}

/**
 * Elimina un libro dal foglio.
 * Blocca l'eliminazione se il libro è attualmente in prestito.
 * @param {string} etichetta
 * @returns {{ok: boolean, errore: string}}
 */
function eliminaLibro(etichetta) {
  try {
    var utente = Session.getActiveUser().getEmail();
    if (utente !== UTENTE_BIBLIOTECARIO) throw new Error('Accesso non autorizzato.');

    var foglio     = SpreadsheetApp.openById(ID_FOGLIO).getSheets()[0];
    var ultimaRiga = foglio.getLastRow();
    if (ultimaRiga < 2) return { ok: false, errore: 'Catalogo vuoto.' };

    var righe       = foglio.getRange(2, 1, ultimaRiga - 1, 7).getValues();
    var rigaTrovata = -1;

    for (var i = 0; i < righe.length; i++) {
      if (String(righe[i][0]).trim() === etichetta.trim()) {
        /* Verifica che non sia in prestito (F o G compilati) */
        if (righe[i][5] || righe[i][6]) {
          return { ok: false, errore: 'Impossibile eliminare: il libro è attualmente in prestito.' };
        }
        rigaTrovata = i + 2;
        break;
      }
    }

    if (rigaTrovata === -1) return { ok: false, errore: 'Libro non trovato.' };

    foglio.deleteRow(rigaTrovata);
    return { ok: true };
  } catch (err) {
    return { ok: false, errore: err.message };
  }
}