/* ScholarKit taskpane logic. Personal-use build. */
"use strict";

var PB = null;              // phrasebank data
var LIBRARY = [];           // [{name, text, words}]
var NGRAM = 5;              // shingle length for similarity
var LS_KEY = "sk_key";
var LS_MODEL = "sk_model";
var LS_LIB = "sk_library";
var DEFAULT_MODEL = "gemini-2.5-flash";

/* ---------- boot ---------- */

Office.onReady(function () {
  initTabs();
  initPhrases();
  initRewrite();
  initSimilarity();
  initSettings();
});

function $(id) { return document.getElementById(id); }

function setStatus(msg, cls) {
  var el = $("status");
  el.textContent = msg || "";
  el.className = "status" + (cls ? " " + cls : "");
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, function (c) {
    return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c];
  });
}

function initTabs() {
  var tabs = document.querySelectorAll(".tab");
  tabs.forEach(function (t) {
    t.addEventListener("click", function () {
      tabs.forEach(function (x) { x.classList.remove("active"); });
      document.querySelectorAll(".panel").forEach(function (p) { p.classList.remove("active"); });
      t.classList.add("active");
      $("tab-" + t.dataset.tab).classList.add("active");
    });
  });
}

/* ---------- Word helpers ---------- */

function insertAtSelection(text) {
  return Word.run(function (ctx) {
    ctx.document.getSelection().insertText(text, Word.InsertLocation.replace);
    return ctx.sync();
  });
}

function getSelectionText() {
  return Word.run(function (ctx) {
    var r = ctx.document.getSelection();
    r.load("text");
    return ctx.sync().then(function () { return r.text || ""; });
  });
}

function getBodyText() {
  return Word.run(function (ctx) {
    var b = ctx.document.body;
    b.load("text");
    return ctx.sync().then(function () { return b.text || ""; });
  });
}

function highlightInDoc(snippet) {
  // Word search strings are capped at 255 chars; stay well under it.
  var needle = snippet.slice(0, 150).trim();
  return Word.run(function (ctx) {
    var results = ctx.document.body.search(needle, {
      matchCase: false, ignorePunct: true, ignoreSpace: true
    });
    results.load("items");
    return ctx.sync().then(function () {
      results.items.forEach(function (r) { r.font.highlightColor = "#FFF3BF"; });
      return ctx.sync().then(function () { return results.items.length; });
    });
  });
}

/* ---------- Phrasebank ---------- */

function initPhrases() {
  fetch("phrasebank.json")
    .then(function (r) {
      if (!r.ok) throw new Error("HTTP " + r.status);
      return r.json();
    })
    .then(function (data) {
      PB = data;
      var sel = $("phrase-group");
      var optAll = document.createElement("option");
      optAll.value = "";
      optAll.textContent = "All sections";
      sel.appendChild(optAll);
      data.groups.forEach(function (g) {
        var o = document.createElement("option");
        o.value = g.id;
        o.textContent = g.name;
        sel.appendChild(o);
      });
      sel.addEventListener("change", renderPhrases);
      $("phrase-search").addEventListener("input", renderPhrases);
      renderPhrases();
    })
    .catch(function (e) {
      $("phrase-list").innerHTML =
        '<p class="empty">Could not load phrasebank.json (' + escapeHtml(e.message) +
        "). Check that the file sits next to taskpane.html on your site.</p>";
    });
}

function renderPhrases() {
  if (!PB) return;
  var q = $("phrase-search").value.trim().toLowerCase();
  var gid = $("phrase-group").value;
  var list = $("phrase-list");
  list.innerHTML = "";
  var shown = 0;

  PB.groups.forEach(function (g) {
    if (gid && g.id !== gid) return;
    g.subs.forEach(function (sub) {
      var hits = sub.phrases.filter(function (p) {
        return !q || p.toLowerCase().indexOf(q) !== -1;
      });
      if (!hits.length) return;
      var head = document.createElement("div");
      head.className = "sub-head";
      head.textContent = g.name + " / " + sub.name;
      list.appendChild(head);
      hits.forEach(function (p) {
        var d = document.createElement("div");
        d.className = "phrase";
        d.innerHTML = escapeHtml(p).replace(/\[([^\]]+)\]/g, '<span class="ph">[$1]</span>');
        d.title = "Click to insert at cursor";
        d.addEventListener("click", function () {
          insertAtSelection(p + " ")
            .then(function () { setStatus("Inserted", "busy"); setTimeout(function () { setStatus(""); }, 1200); })
            .catch(function (e) { setStatus("Insert failed: " + e.message, "err"); });
        });
        list.appendChild(d);
        shown++;
      });
    });
  });

  if (!shown) list.innerHTML = '<p class="empty">No phrases match that search.</p>';
}

/* ---------- Gemini ---------- */

function geminiJson(prompt) {
  var key = localStorage.getItem(LS_KEY);
  var model = localStorage.getItem(LS_MODEL) || DEFAULT_MODEL;
  if (!key) return Promise.reject(new Error("No API key. Open the Settings tab and save your Gemini key."));
  var url = "https://generativelanguage.googleapis.com/v1beta/models/" +
    encodeURIComponent(model) + ":generateContent?key=" + encodeURIComponent(key);
  return fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      contents: [{ parts: [{ text: prompt }] }],
      generationConfig: { temperature: 0.7, responseMimeType: "application/json" }
    })
  }).then(function (res) {
    return res.json().then(function (data) {
      if (!res.ok) {
        var m = (data.error && data.error.message) || ("HTTP " + res.status);
        throw new Error(m);
      }
      var txt = "";
      try { txt = data.candidates[0].content.parts[0].text; }
      catch (e) { throw new Error("Empty response from model."); }
      txt = txt.replace(/^```(json)?/i, "").replace(/```$/m, "").trim();
      return JSON.parse(txt);
    });
  });
}

/* ---------- Rewrite ---------- */

var REWRITE_PROMPTS = {
  paraphrase: "Rewrite the text in formal academic English suitable for a research paper. Preserve the meaning exactly. Vary sentence structure and vocabulary. Do not add new claims or citations.",
  formal: "Rewrite the text in a more formal, scholarly register. Preserve the meaning exactly. Remove colloquialisms and contractions.",
  simplify: "Rewrite the text in clearer, simpler academic English. Shorter sentences, plainer vocabulary, same meaning. Do not lose any substantive content.",
  concise: "Rewrite the text to be significantly more concise while preserving every substantive point. Cut redundancy and filler."
};

function initRewrite() {
  $("rw-grab").addEventListener("click", function () {
    getSelectionText().then(function (t) {
      if (!t.trim()) { setStatus("Nothing selected in the document.", "err"); return; }
      $("rw-input").value = t.trim();
      setStatus("");
    }).catch(function (e) { setStatus(e.message, "err"); });
  });

  $("rw-go").addEventListener("click", function () {
    var text = $("rw-input").value.trim();
    if (!text) { setStatus("Add some text first.", "err"); return; }
    if (text.length > 4000) { setStatus("Keep it under ~4000 characters per pass.", "err"); return; }
    var mode = $("rw-mode").value;
    var prompt = "You are an academic writing assistant. " + REWRITE_PROMPTS[mode] +
      "\n\nReturn ONLY a JSON array of exactly 3 strings. Each string is one complete rewrite of the full text. No keys, no commentary, no markdown.\n\nText:\n" + text;
    runBusy($("rw-go"), "Rewriting...", geminiJson(prompt).then(function (arr) {
      if (!Array.isArray(arr)) throw new Error("Unexpected response shape.");
      renderRewriteOptions(arr.slice(0, 3));
    }));
  });

  $("wc-go").addEventListener("click", function () {
    var w = $("wc-input").value.trim();
    if (!w) return;
    var prompt = "Suggest 8 academic alternatives for the word or phrase below, suitable for a formal research paper. Single words or short phrases only. Return ONLY a JSON array of strings.\n\nWord or phrase: " + w;
    runBusy($("wc-go"), "Thinking...", geminiJson(prompt).then(function (arr) {
      if (!Array.isArray(arr)) throw new Error("Unexpected response shape.");
      var box = $("wc-results");
      box.innerHTML = "";
      arr.forEach(function (s) {
        var c = document.createElement("button");
        c.className = "chip";
        c.textContent = s;
        c.title = "Click to insert at cursor";
        c.addEventListener("click", function () {
          insertAtSelection(s).catch(function (e) { setStatus(e.message, "err"); });
        });
        box.appendChild(c);
      });
    }));
  });
}

function renderRewriteOptions(options) {
  var box = $("rw-results");
  box.innerHTML = "";
  options.forEach(function (opt, i) {
    var card = document.createElement("div");
    card.className = "option-card";
    var p = document.createElement("div");
    p.textContent = opt;
    var actions = document.createElement("div");
    actions.className = "opt-actions";
    actions.appendChild(makeBtn("Replace selection", function () {
      insertAtSelection(opt)
        .then(function () { setStatus("Replaced", "busy"); setTimeout(function () { setStatus(""); }, 1200); })
        .catch(function (e) { setStatus(e.message, "err"); });
    }));
    actions.appendChild(makeBtn("Copy", function () { copyText(opt); }));
    card.appendChild(p);
    card.appendChild(actions);
    box.appendChild(card);
  });
}

function makeBtn(label, fn) {
  var b = document.createElement("button");
  b.className = "btn small";
  b.textContent = label;
  b.addEventListener("click", fn);
  return b;
}

function copyText(t) {
  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(t).then(function () { setStatus("Copied", "busy"); setTimeout(function () { setStatus(""); }, 1000); });
  } else {
    var ta = document.createElement("textarea");
    ta.value = t;
    document.body.appendChild(ta);
    ta.select();
    document.execCommand("copy");
    document.body.removeChild(ta);
    setStatus("Copied", "busy");
    setTimeout(function () { setStatus(""); }, 1000);
  }
}

function runBusy(btn, label, promise) {
  var orig = btn.textContent;
  btn.disabled = true;
  btn.textContent = label;
  setStatus(label, "busy");
  promise
    .then(function () { setStatus(""); })
    .catch(function (e) { setStatus(e.message, "err"); })
    .finally(function () { btn.disabled = false; btn.textContent = orig; });
}

/* ---------- Similarity ---------- */

function initSimilarity() {
  if (window.pdfjsLib) {
    pdfjsLib.GlobalWorkerOptions.workerSrc =
      "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
  }
  loadLibrary();
  $("sim-files").addEventListener("change", function (ev) {
    var files = Array.prototype.slice.call(ev.target.files);
    files.reduce(function (chain, f) {
      return chain.then(function () { return addFileToLibrary(f); });
    }, Promise.resolve()).then(function () {
      ev.target.value = "";
      saveLibrary();
      renderLibrary();
      setStatus("");
    }).catch(function (e) { setStatus("Import failed: " + e.message, "err"); });
  });
  $("sim-check").addEventListener("click", checkSimilarity);
}

function addFileToLibrary(file) {
  setStatus("Reading " + file.name + "...", "busy");
  var name = file.name;
  var lower = name.toLowerCase();
  var p;
  if (lower.endsWith(".pdf")) p = extractPdf(file);
  else if (lower.endsWith(".docx")) p = extractDocx(file);
  else p = file.text();
  return p.then(function (text) {
    text = (text || "").trim();
    if (!text) throw new Error(name + " produced no text (scanned PDF? It needs OCR first).");
    LIBRARY.push({ name: name, text: text, words: tokenize(text) });
  });
}

function extractPdf(file) {
  return file.arrayBuffer().then(function (buf) {
    return pdfjsLib.getDocument({ data: buf }).promise.then(function (pdf) {
      var pages = [];
      for (var i = 1; i <= pdf.numPages; i++) pages.push(i);
      return pages.reduce(function (chain, n) {
        return chain.then(function (acc) {
          return pdf.getPage(n).then(function (page) {
            return page.getTextContent().then(function (tc) {
              return acc + tc.items.map(function (it) { return it.str; }).join(" ") + "\n";
            });
          });
        });
      }, Promise.resolve(""));
    });
  });
}

function extractDocx(file) {
  return file.arrayBuffer().then(function (buf) {
    return mammoth.extractRawText({ arrayBuffer: buf }).then(function (r) { return r.value; });
  });
}

function loadLibrary() {
  try {
    var raw = localStorage.getItem(LS_LIB);
    if (raw) {
      LIBRARY = JSON.parse(raw).map(function (d) {
        return { name: d.name, text: d.text, words: tokenize(d.text) };
      });
    }
  } catch (e) { LIBRARY = []; }
  renderLibrary();
}

function saveLibrary() {
  try {
    var slim = LIBRARY.map(function (d) { return { name: d.name, text: d.text }; });
    var json = JSON.stringify(slim);
    if (json.length > 4500000) {
      setStatus("Library too large to save between sessions; it will stay in memory for now.", "err");
      return;
    }
    localStorage.setItem(LS_LIB, json);
  } catch (e) {
    setStatus("Could not persist library (storage full). It will stay in memory for now.", "err");
  }
}

function renderLibrary() {
  var ul = $("sim-library");
  ul.innerHTML = "";
  if (!LIBRARY.length) {
    ul.innerHTML = '<li class="empty" style="border:none;background:none;">No reference documents yet.</li>';
    return;
  }
  LIBRARY.forEach(function (doc, i) {
    var li = document.createElement("li");
    var left = document.createElement("span");
    left.innerHTML = escapeHtml(doc.name) +
      ' <span class="meta">(' + doc.words.length.toLocaleString() + " words)</span>";
    var rm = document.createElement("button");
    rm.className = "rm";
    rm.textContent = "\u00d7";
    rm.title = "Remove";
    rm.addEventListener("click", function () {
      LIBRARY.splice(i, 1);
      saveLibrary();
      renderLibrary();
    });
    li.appendChild(left);
    li.appendChild(rm);
    ul.appendChild(li);
  });
}

function tokenize(t) {
  return t.toLowerCase().replace(/[^a-z0-9\s]/g, " ").split(/\s+/).filter(Boolean);
}

function buildShingleSet(words) {
  var set = {};
  for (var i = 0; i + NGRAM <= words.length; i++) {
    set[words.slice(i, i + NGRAM).join(" ")] = true;
  }
  return set;
}

function checkSimilarity() {
  if (!LIBRARY.length) { setStatus("Add at least one reference document first.", "err"); return; }
  var btn = $("sim-check");
  btn.disabled = true;
  setStatus("Checking document...", "busy");
  getBodyText().then(function (draft) {
    var words = tokenize(draft);
    if (words.length < NGRAM) throw new Error("The document is too short to check.");

    var globalMatched = new Array(words.length).fill(false);
    var perSource = [];

    LIBRARY.forEach(function (doc) {
      var refSet = buildShingleSet(doc.words);
      var matched = new Array(words.length).fill(false);
      for (var i = 0; i + NGRAM <= words.length; i++) {
        if (refSet[words.slice(i, i + NGRAM).join(" ")]) {
          for (var j = i; j < i + NGRAM; j++) { matched[j] = true; globalMatched[j] = true; }
        }
      }
      var count = matched.filter(Boolean).length;
      perSource.push({ name: doc.name, matched: matched, count: count });
    });

    var overall = globalMatched.filter(Boolean).length / words.length;
    renderSimilarity(overall, perSource, words);
    setStatus("");
  }).catch(function (e) {
    setStatus(e.message, "err");
  }).finally(function () { btn.disabled = false; });
}

function spansFromMask(mask) {
  var spans = [];
  var start = -1;
  for (var i = 0; i <= mask.length; i++) {
    if (i < mask.length && mask[i]) {
      if (start === -1) start = i;
    } else if (start !== -1) {
      spans.push([start, i - 1]);
      start = -1;
    }
  }
  return spans;
}

function renderSimilarity(overall, perSource, words) {
  var pct = Math.round(overall * 1000) / 10;
  var cls = pct < 5 ? "ok" : (pct < 15 ? "" : "high");
  var summary = $("sim-summary");
  summary.innerHTML =
    '<div class="score-card ' + cls + '">' +
    '<div class="big">' + pct + "%</div>" +
    "<div>of your draft overlaps your reference library in runs of " + NGRAM + "+ identical words.</div>" +
    "</div>";

  var box = $("sim-matches");
  box.innerHTML = "";
  var shown = 0;

  perSource.forEach(function (src) {
    if (!src.count) return;
    spansFromMask(src.matched).forEach(function (span) {
      if (shown >= 40) return;
      var snippet = words.slice(span[0], span[1] + 1).join(" ");
      var div = document.createElement("div");
      div.className = "match";
      div.innerHTML =
        '<div class="src">' + escapeHtml(src.name) + "</div>" +
        '<div><span class="snippet">' + escapeHtml(snippet) + "</span></div>";
      var hl = makeBtn("Highlight in document", function () {
        highlightInDoc(snippet).then(function (n) {
          setStatus(n ? "Highlighted " + n + " match(es)" : "Could not locate that exact passage (punctuation differences).", n ? "busy" : "err");
          if (n) setTimeout(function () { setStatus(""); }, 1500);
        }).catch(function (e) { setStatus(e.message, "err"); });
      });
      hl.style.marginTop = "6px";
      div.appendChild(hl);
      box.appendChild(div);
      shown++;
    });
  });

  if (!shown) {
    box.innerHTML = '<p class="hint">No overlapping passages of ' + NGRAM + '+ words found. Clean.</p>';
  } else if (shown >= 40) {
    var more = document.createElement("p");
    more.className = "hint";
    more.textContent = "Showing the first 40 matches. Fix these, then run the check again.";
    box.appendChild(more);
  }
}

/* ---------- Settings ---------- */

function initSettings() {
  $("set-key").value = localStorage.getItem(LS_KEY) || "";
  $("set-model").value = localStorage.getItem(LS_MODEL) || DEFAULT_MODEL;

  $("set-save").addEventListener("click", function () {
    localStorage.setItem(LS_KEY, $("set-key").value.trim());
    localStorage.setItem(LS_MODEL, $("set-model").value.trim() || DEFAULT_MODEL);
    $("set-msg").textContent = "Saved.";
    setTimeout(function () { $("set-msg").textContent = ""; }, 1500);
  });

  $("set-test").addEventListener("click", function () {
    localStorage.setItem(LS_KEY, $("set-key").value.trim());
    localStorage.setItem(LS_MODEL, $("set-model").value.trim() || DEFAULT_MODEL);
    $("set-msg").textContent = "Testing...";
    geminiJson('Return ONLY the JSON array ["ok"].')
      .then(function () { $("set-msg").textContent = "Connected. The Rewrite tab is ready."; })
      .catch(function (e) { $("set-msg").textContent = "Failed: " + e.message; });
  });
}
