/* Pannello di Word: legge la selezione, mostra l'anteprima, inserisce il risultato. */
(function () {
  "use strict";

  var $ = function (id) { return document.getElementById(id); };
  var inWord = false;
  var ultimo = null;           // ultimo risultato valido
  var CHIAVE = "numeriInLettere.opzioni";

  function opzioni() {
    var fn = document.querySelector('input[name="fn"]:checked');
    return {
      tipo: $("tipo").value,
      formatoNumero: fn ? fn.value : "euro",
      formatoData: $("formatoData").value,
      decimaliEuro: $("decimaliEuro").checked,
      milioniUniti: $("milioniUniti").checked
    };
  }

  function salvaOpzioni() {
    try {
      var o = opzioni(); delete o.tipo;
      window.localStorage.setItem(CHIAVE, JSON.stringify(o));
    } catch (e) { /* memoria del browser non disponibile: si ignora */ }
  }

  function caricaOpzioni() {
    try {
      var o = JSON.parse(window.localStorage.getItem(CHIAVE) || "null");
      if (!o) return;
      var r = document.querySelector('input[name="fn"][value="' + o.formatoNumero + '"]');
      if (r) r.checked = true;
      if (o.formatoData) $("formatoData").value = o.formatoData;
      $("decimaliEuro").checked = !!o.decimaliEuro;
      $("milioniUniti").checked = !!o.milioniUniti;
    } catch (e) { /* nessuna preferenza salvata */ }
  }

  function stato(msg, tipo) {
    var s = $("stato");
    s.textContent = msg || "";
    s.className = "stato" + (tipo ? " " + tipo : "");
  }

  function aggiorna() {
    var testo = $("testo").value;
    var ris = $("risultato"), err = $("errore");
    ultimo = null;
    err.hidden = true;
    if (!testo.trim()) {
      ris.textContent = "—";
      ris.className = "risultato vuoto";
    } else {
      try {
        ultimo = NumeriInLettere.converti(testo, opzioni());
        ris.textContent = ultimo.completo;
        ris.className = "risultato";
      } catch (e) {
        ris.textContent = "—";
        ris.className = "risultato vuoto";
        err.textContent = e.message;
        err.hidden = false;
      }
    }
    $("copia").disabled = !ultimo;
    $("sostituisci").disabled = !(ultimo && inWord);
    $("aggiungi").disabled = !(ultimo && inWord);
  }

  // ---------- Word ----------

  function leggiSelezione() {
    if (!inWord) return;
    Word.run(function (ctx) {
      var sel = ctx.document.getSelection();
      sel.load("text");
      return ctx.sync().then(function () {
        var t = (sel.text || "").trim();
        if (t && t.length <= 60 && !/[\r\n]/.test(t)) {
          $("testo").value = t;
          aggiorna();
          stato("");
        }
      });
    }).catch(function () { /* selezione non leggibile: si lascia il campo com'è */ });
  }

  function scrivi(modalita) {
    if (!ultimo || !inWord) return;
    var risultato = ultimo;
    Word.run(function (ctx) {
      var sel = ctx.document.getSelection();
      sel.load("text");
      return ctx.sync().then(function () {
        var originale = sel.text || "";
        if (/[\r\n\u000b\u0007]/.test(originale)) {
          throw new Error("La selezione comprende un a capo: seleziona solo il numero o la data.");
        }
        var m = originale.match(/^(\s*)([\s\S]*?)(\s*)$/);
        var prima = m[1], nucleo = m[2], dopo = m[3];
        var testo;
        if (modalita === "aggiungi") {
          if (!nucleo) throw new Error("Seleziona nel documento il numero a cui aggiungere le lettere.");
          testo = prima + nucleo + " (" + risultato.lettere + ")" + dopo;
        } else {
          testo = prima + risultato.completo + dopo;
        }
        sel.insertText(testo, "Replace");
        return ctx.sync();
      });
    }).then(function () {
      stato(modalita === "aggiungi" ? "Lettere aggiunte nel documento." : "Selezione sostituita.", "ok");
    }).catch(function (e) {
      stato(e && e.message ? e.message : "Inserimento non riuscito.", "err");
    });
  }

  function copia() {
    if (!ultimo) return;
    var testo = ultimo.completo;
    var fallback = function () {
      var ta = document.createElement("textarea");
      ta.value = testo; document.body.appendChild(ta); ta.select();
      var ok = false;
      try { ok = document.execCommand("copy"); } catch (e) { ok = false; }
      document.body.removeChild(ta);
      stato(ok ? "Copiato negli appunti." : "Copia non riuscita: seleziona l'anteprima e premi ⌘C.", ok ? "ok" : "err");
    };
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(testo).then(function () { stato("Copiato negli appunti.", "ok"); }, fallback);
    } else fallback();
  }

  // ---------- avvio ----------

  function collega() {
    caricaOpzioni();
    $("testo").addEventListener("input", aggiorna);
    $("tipo").addEventListener("change", aggiorna);
    $("formatoData").addEventListener("change", function () { salvaOpzioni(); aggiorna(); });
    $("decimaliEuro").addEventListener("change", function () { salvaOpzioni(); aggiorna(); });
    $("milioniUniti").addEventListener("change", function () { salvaOpzioni(); aggiorna(); });
    Array.prototype.forEach.call(document.querySelectorAll('input[name="fn"]'), function (r) {
      r.addEventListener("change", function () { salvaOpzioni(); aggiorna(); });
    });
    $("sostituisci").addEventListener("click", function () { scrivi("sostituisci"); });
    $("aggiungi").addEventListener("click", function () { scrivi("aggiungi"); });
    $("copia").addEventListener("click", copia);
    aggiorna();
  }

  function avvio(info) {
    inWord = !!(info && window.Office && info.host === Office.HostType.Word);
    if (inWord) {
      Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, leggiSelezione);
      leggiSelezione();
    } else {
      stato("Aperto fuori da Word: puoi provare le conversioni; l'inserimento nel documento funziona solo in Word.");
    }
    aggiorna();
  }

  document.addEventListener("DOMContentLoaded", collega);
  if (window.Office && Office.onReady) Office.onReady(avvio);
  else document.addEventListener("DOMContentLoaded", function () { avvio(null); });
})();
