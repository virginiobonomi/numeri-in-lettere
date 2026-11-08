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

  // ---------- scorciatoia da tastiera ----------
  // Legge il numero (o data, o quota) che precede il cursore e inserisce subito dopo «(in lettere)».
  // Se c'è una selezione, aggiunge le lettere dopo la selezione.

  function lettereAlCursore(event) {
    var finito = function () {
      if (event && typeof event.completed === "function") event.completed();
    };
    var opz = opzioni();
    return Word.run(function (ctx) {
      var sel = ctx.document.getSelection();
      sel.load("text,isEmpty");
      return ctx.sync().then(function () {
        if (!sel.isEmpty) {
          var originale = sel.text || "";
          if (/[\r\n\u000b\u0007]/.test(originale)) {
            throw new Error("La selezione comprende un a capo: seleziona solo il numero o la data.");
          }
          var m = originale.match(/^(\s*)([\s\S]*?)(\s*)$/);
          if (!m[2]) throw new Error("Posiziona il cursore subito dopo un numero, una data o una quota.");
          var r = NumeriInLettere.lettereDopo(m[2], opz);
          var nuovo = sel.insertText(m[1] + m[2] + " (" + r.lettere + ")" + m[3], "Replace");
          nuovo.select("End");
          return ctx.sync();
        }
        var paragrafo = sel.paragraphs.getFirst();
        var prima = paragrafo.getRange("Start").expandTo(sel.getRange("Start"));
        prima.load("text");
        return ctx.sync().then(function () {
          var r = NumeriInLettere.lettereDopo(prima.text, opz);
          var inserito = sel.insertText(r.daInserire, "End");
          inserito.select("End");
          return ctx.sync();
        });
      });
    }).then(function () {
      stato("Lettere inserite nel documento.", "ok");
      finito();
    }, function (e) {
      // Nessun inserimento: si apre il pannello per mostrare il motivo.
      stato(e && e.message ? e.message : "Inserimento non riuscito.", "err");
      var apri = (window.Office && Office.addin && Office.addin.showAsTaskpane)
        ? Office.addin.showAsTaskpane() : null;
      return Promise.resolve(apri).catch(function () {}).then(finito);
    });
  }

  function mostraPannello(event) {
    var p = (window.Office && Office.addin && Office.addin.showAsTaskpane)
      ? Office.addin.showAsTaskpane() : null;
    return Promise.resolve(p).catch(function () {}).then(function () {
      if (event && typeof event.completed === "function") event.completed();
    });
  }

  try {
    if (window.Office && Office.actions && Office.actions.associate) {
      Office.actions.associate("InLettere", lettereAlCursore);
      Office.actions.associate("MostraPannello", mostraPannello);
    }
  } catch (e) { /* fuori da Word: comandi non disponibili */ }

  // ---------- scelta della scorciatoia da tastiera ----------

  var nuovaCombinazione = null;
  var confermaOccupata = false;   // secondo clic su una combinazione già usata = usala comunque
  // Su Mac i tasti si mostrano come simboli (⌘ ⌥ ⇧ ⌃), su Windows per nome (Ctrl, Alt, Maiusc).
  function suMac() {
    try {
      if (window.Office && Office.context && Office.context.diagnostics) {
        return Office.context.diagnostics.platform === "Mac";
      }
    } catch (e) { /* si usa il controllo del browser */ }
    return /Mac/i.test(navigator.platform || navigator.userAgent || "");
  }
  var NOMI_MAC = { Command: "⌘", Option: "⌥", Alt: "⌥", Shift: "⇧", Ctrl: "⌃", Control: "⌃" };
  var NOMI_WIN = { Command: "Ctrl", Ctrl: "Ctrl", Control: "Ctrl", Alt: "Alt", Option: "Alt", Shift: "Maiusc" };
  function nomeTasto(p) { return (suMac() ? NOMI_MAC : NOMI_WIN)[p] || p; }

  function inSimboli(combo) {
    return String(combo).split("+").map(function (p) {
      return "<kbd>" + nomeTasto(p) + "</kbd>";
    }).join(suMac() ? "" : "+");
  }

  function esitoScorciatoia(msg, tipo) {
    var e = $("esitoScorciatoia");
    e.textContent = msg || "";
    e.className = "nota" + (tipo ? " " + tipo : "");
  }

  function scorciatoieDisponibili() {
    try {
      return !!(inWord && Office.actions && Office.actions.getShortcuts &&
        Office.context.requirements.isSetSupported("KeyboardShortcuts", "1.1"));
    } catch (e) { return false; }
  }

  function mostraScorciatoiaAttuale() {
    var att = $("scorciatoiaAttuale");
    if (!inWord) { att.textContent = "Disponibile solo dentro Word."; return; }
    if (!scorciatoieDisponibili()) {
      att.innerHTML = "Questa versione di Word non gestisce le scorciatoie dei componenti aggiuntivi " +
        "(su Mac serve Word 16.88 o successivo; su Windows Word versione 2408 o successiva). Seleziona il numero e usa il pannello.";
      $("registra").disabled = true; $("ripristinaScorciatoia").disabled = true;
      return;
    }
    Office.actions.getShortcuts().then(function (s) {
      var c = s && s.InLettere;
      if (c) {
        att.innerHTML = "Attuale: " + inSimboli(c);
        $("tastoRapido").innerHTML = inSimboli(c);
      } else {
        att.textContent = "Nessuna scorciatoia attiva (la combinazione è stata assegnata a un'altra funzione). Scegline una qui sotto.";
        $("tastoRapido").textContent = "la scorciatoia che sceglierai in fondo al pannello";
      }
    }).catch(function () {
      att.textContent = "Impossibile leggere la scorciatoia attuale.";
    });
  }

  function registraTasti(ev) {
    ev.preventDefault();
    var m = ev.code.match(/^(?:Key([A-Z])|Digit([0-9]))$/);
    var tasto = m ? (m[1] || m[2]) : (ev.code === "Minus" ? "-" : null);
    var mac = suMac(), mod = [];
    if (ev.metaKey && mac) mod.push("Command");
    if (ev.ctrlKey) mod.push("Ctrl");
    if (ev.altKey) mod.push(mac ? "Option" : "Alt");
    if (ev.shiftKey) mod.push("Shift");
    var sep = mac ? " " : " + ";
    if (!tasto) { $("registra").value = mod.map(nomeTasto).join(sep) + (mod.length ? sep : "") + "…"; return; }
    if (!((ev.metaKey && mac) || ev.ctrlKey || ev.altKey)) {
      nuovaCombinazione = null; $("salvaScorciatoia").disabled = true;
      esitoScorciatoia(mac
        ? "Serve almeno ⌘, ⌥ o ⌃ insieme alla lettera (⇧ da solo non basta)."
        : "Serve almeno Ctrl o Alt insieme alla lettera (Maiusc da solo non basta).", "err");
      return;
    }
    nuovaCombinazione = mod.concat([tasto]).join("+");
    confermaOccupata = false;
    $("registra").value = mod.map(nomeTasto).concat([tasto]).join(sep);
    $("salvaScorciatoia").disabled = false;
    esitoScorciatoia("");
  }

  function salvaScorciatoia() {
    if (!nuovaCombinazione || !scorciatoieDisponibili()) return;
    var combo = nuovaCombinazione;
    var forza = confermaOccupata;
    var controllo = (!forza && Office.actions.areShortcutsInUse)
      ? Office.actions.areShortcutsInUse([combo]) : Promise.resolve([]);
    controllo.then(function (esito) {
      var occupata = esito && esito[0] && esito[0].inUse;
      if (occupata) {
        esitoScorciatoia("Questa combinazione risulta già usata. Provane un'altra, oppure premi di nuovo il pulsante per usarla comunque.", "err");
        confermaOccupata = true;
        return;
      }
      return Office.actions.replaceShortcuts({ InLettere: combo }).then(function () {
        esitoScorciatoia("Salvata. Torna nel documento, metti il cursore dopo un numero e prova.", "ok");
        mostraScorciatoiaAttuale();
      });
    }).catch(function (e) {
      esitoScorciatoia("Word non ha accettato la combinazione" + (e && e.message ? " (" + e.message + ")" : "") + ". Provane un'altra.", "err");
    });
  }

  function ripristinaScorciatoia() {
    if (!scorciatoieDisponibili()) return;
    Office.actions.replaceShortcuts({ InLettere: null }).then(function () {
      esitoScorciatoia("Ripristinata la combinazione predefinita.", "ok");
      mostraScorciatoiaAttuale();
    }).catch(function () { esitoScorciatoia("Ripristino non riuscito.", "err"); });
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
    $("registra").addEventListener("keydown", registraTasti);
    $("salvaScorciatoia").addEventListener("click", salvaScorciatoia);
    $("ripristinaScorciatoia").addEventListener("click", ripristinaScorciatoia);
    aggiorna();
  }

  function avvio(info) {
    inWord = !!(info && window.Office && info.host === Office.HostType.Word);
    if (inWord) {
      Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, leggiSelezione);
      leggiSelezione();
      mostraScorciatoiaAttuale();
    } else {
      mostraScorciatoiaAttuale();
      stato("Aperto fuori da Word: puoi provare le conversioni; l'inserimento nel documento funziona solo in Word.");
    }
    aggiorna();
  }

  document.addEventListener("DOMContentLoaded", collega);
  if (window.Office && Office.onReady) Office.onReady(avvio);
  else document.addEventListener("DOMContentLoaded", function () { avvio(null); });
})();
