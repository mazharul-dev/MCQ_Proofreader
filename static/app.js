const fileInput = document.querySelector("#fileInput");
const dropMirror = document.querySelector("#dropInputMirror");
const dropZone = document.querySelector("#dropZone");
const fileName = document.querySelector("#fileName");
const languageSelect = document.querySelector("#languageSelect");
const spellButton = document.querySelector("#spellButton");
const statusLine = document.querySelector("#status");
const summary = document.querySelector("#summary");
const warningsBox = document.querySelector("#warnings");
const cards = document.querySelector("#cards");
const metricTotal = document.querySelector("#metricTotal");
const metricDuplicates = document.querySelector("#metricDuplicates");
const metricSpell = document.querySelector("#metricSpell");

const state = {
  questions: [],
  duplicatePairs: [],
  warnings: [],
  spellMatches: new Map(),
  file: null,
};

const bnDigits = new Map([
  ["0", "০"],
  ["1", "১"],
  ["2", "২"],
  ["3", "৩"],
  ["4", "৪"],
  ["5", "৫"],
  ["6", "৬"],
  ["7", "৭"],
  ["8", "৮"],
  ["9", "৯"],
]);

fileInput.addEventListener("change", () => {
  const file = fileInput.files?.[0];
  if (file) parseFile(file);
});

dropMirror.addEventListener("change", () => {
  const file = dropMirror.files?.[0];
  if (file) parseFile(file);
});

dropZone.addEventListener("dragover", (event) => {
  event.preventDefault();
  dropZone.classList.add("drag-over");
});

dropZone.addEventListener("dragleave", () => {
  dropZone.classList.remove("drag-over");
});

dropZone.addEventListener("drop", (event) => {
  event.preventDefault();
  dropZone.classList.remove("drag-over");
  const file = [...event.dataTransfer.files].find((item) => item.name.toLowerCase().endsWith(".docx"));
  if (file) parseFile(file);
});

spellButton.addEventListener("click", async () => {
  if (!state.questions.length) return;

  setStatus("Spell checking with LanguageTool + Bangla wordlist ...", "busy");
  spellButton.disabled = true;
  state.spellMatches.clear();
  updateMetrics();
  renderCards();

  try {
    const segments = collectSpellSegments();
    const response = await fetch("/api/spellcheck", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ language: languageSelect.value, segments }),
    });
    const payload = await response.json();
    if (!response.ok) throw new Error(payload.detail || "Spell check failed.");

    for (const segment of payload.segments || []) {
      state.spellMatches.set(segment.id, segment.matches || []);
    }

    renderCards();
    runMathJax();
    updateMetrics();

    const totalMatches = spellIssueCount();
    const suffix = payload.errors?.length ? ` ${toBanglaNumber(payload.errors.length)}টি অংশ যাচাই করা যায়নি।` : "";
    setStatus(`${toBanglaNumber(totalMatches)}টি সম্ভাব্য বানান/ভাষাগত সমস্যা পাওয়া গেছে।${suffix}`);
    if (payload.errors?.length) renderWarnings([...state.warnings, ...payload.errors]);
  } catch (error) {
    setStatus(error.message, "error");
  } finally {
    spellButton.disabled = state.questions.length === 0;
  }
});

document.addEventListener("click", (event) => {
  const jump = event.target.closest("[data-jump]");
  if (!jump) return;
  const target = document.getElementById(jump.dataset.jump);
  if (!target) return;
  target.scrollIntoView({ behavior: "smooth", block: "start" });
  target.classList.add("pulse");
  window.setTimeout(() => target.classList.remove("pulse"), 1200);
});

window.addEventListener("load", runIcons);

async function parseFile(file) {
  setStatus(`Parsing ${file.name} ...`, "busy");
  clearState();
  state.file = file;
  fileName.textContent = `${file.name} · ${formatFileSize(file.size)}`;

  const body = new FormData();
  body.append("upload", file);

  try {
    const response = await fetch("/api/parse", { method: "POST", body });
    const payload = await response.json();
    if (!response.ok) throw new Error(payload.detail || "Could not parse DOCX.");

    state.questions = payload.questions || [];
    state.duplicatePairs = payload.duplicatePairs || [];
    state.warnings = payload.warnings || [];
    render();
    setStatus(`${toBanglaNumber(state.questions.length)}টি প্রশ্ন মেমরি থেকে লোড হয়েছে।`);
    spellButton.disabled = state.questions.length === 0;
    updateMetrics();
    runMathJax();
    runIcons();
  } catch (error) {
    setStatus(error.message, "error");
    renderEmpty("DOCX parsing failed.", "ফাইলের table format বা cell index মিলছে কি না দেখুন।");
    updateMetrics();
  }
}

function clearState() {
  state.questions = [];
  state.duplicatePairs = [];
  state.warnings = [];
  state.spellMatches.clear();
  spellButton.disabled = true;
  summary.hidden = true;
  warningsBox.hidden = true;
  updateMetrics();
}

function render() {
  renderSummary();
  renderWarnings(state.warnings);
  renderCards();
}

function renderSummary() {
  if (!state.duplicatePairs.length) {
    summary.hidden = true;
    summary.innerHTML = "";
    return;
  }

  const items = state.duplicatePairs
    .map((pair) => {
      const source = pair.source;
      const repeat = pair.repeat;
      return `<li>প্রশ্ন নম্বর ${escapeHtml(toBanglaNumber(source.serial))}-এর সাথে প্রশ্ন নম্বর ${escapeHtml(
        toBanglaNumber(repeat.serial)
      )} রিপিট হয়েছে। <button class="jump-link" type="button" data-jump="${escapeAttr(source.uid)}">প্রথম প্রশ্ন</button> <button class="jump-link" type="button" data-jump="${escapeAttr(
        repeat.uid
      )}">রিপিট প্রশ্ন</button></li>`;
    })
    .join("");

  summary.innerHTML = `<strong>Duplicate Warning</strong><ul class="summary-list">${items}</ul>`;
  summary.hidden = false;
}

function renderWarnings(warnings) {
  if (!warnings.length) {
    warningsBox.hidden = true;
    warningsBox.innerHTML = "";
    return;
  }

  warningsBox.innerHTML = `<strong>Warnings</strong><ul class="warning-list">${warnings
    .map((warning) => `<li>${escapeHtml(warning)}</li>`)
    .join("")}</ul>`;
  warningsBox.hidden = false;
}

function renderCards() {
  if (!state.questions.length) {
    renderEmpty("Preview stream is empty.", "DOCX upload করলে প্রশ্ন, equation, image, answer ও solve এখানে দেখা যাবে।");
    return;
  }

  cards.innerHTML = state.questions.map(renderCard).join("");
  runIcons();
}

function renderCard(question, index) {
  const duplicate = question.duplicateTargets?.length || question.duplicateOf;
  const duplicateBadges = renderDuplicateBadges(question);
  const options = arrangeOptions(question.options || [])
    .map((option) => renderOption(option, index))
    .join("");
  const answer = question.answerLabel
    ? `উত্তর: (${escapeHtml(question.answerLabel)}) ${renderParts(
        question.answerParts,
        question.answerText || "",
        segmentId(index, "answer")
      )}`
    : `উত্তর: ${escapeHtml(question.answerRaw || "নির্ধারণ করা যায়নি")}`;

  return `
    <article id="${escapeAttr(question.uid)}" class="mcq-card ${duplicate ? "duplicate" : ""}">
      <aside class="serial" aria-label="Question serial">
        <span class="serial-label">Serial</span>
        <span class="serial-number">${escapeHtml(toBanglaNumber(question.serial))}</span>
      </aside>
      <div class="question-shell">
        <header class="card-header">
          <div class="question" data-segment="${segmentId(index, "question")}">${renderParts(
            question.questionParts,
            question.question,
            segmentId(index, "question")
          )}</div>
          <div class="category">${escapeHtml(question.category || "No Category")}</div>
        </header>
        ${duplicateBadges}
        <div class="card-body">
          <div class="options-grid">${options}</div>
          <section class="answer-block">
            <p class="answer-title">সঠিক উত্তর</p>
            <div class="answer-text">${answer}</div>
          </section>
          <section class="explanation-block">
            <p class="explanation-title">ব্যাখ্যা</p>
            <div class="explanation-text" data-segment="${segmentId(index, "explanation")}">${renderParts(
              question.explanationParts,
              question.explanation || "",
              segmentId(index, "explanation")
            )}</div>
          </section>
        </div>
      </div>
    </article>
  `;
}

function renderDuplicateBadges(question) {
  const targets = question.duplicateTargets || [];
  if (!targets.length) return "";

  return `<div class="duplicate-row">${targets
    .map(
      (target) =>
        `<button type="button" class="dup-badge" data-jump="${escapeAttr(target.uid)}">রিপিট: ${escapeHtml(
          toBanglaNumber(target.serial)
        )}</button>`
    )
    .join("")}</div>`;
}

function renderOption(option, questionIndex) {
  const id = segmentId(questionIndex, `option-${option.key}`);

  return `
    <div class="option option-${escapeAttr(option.key)}">
      <div class="option-label">(${escapeHtml(option.label)})</div>
      <div class="option-text" data-segment="${id}">${renderParts(option.parts, option.text || "", id)}</div>
    </div>
  `;
}

function arrangeOptions(options) {
  return ["A", "B", "C", "D"].map((key) => options.find((option) => option.key === key) || { key, label: "", text: "" });
}

function collectSpellSegments() {
  const segments = [];

  state.questions.forEach((question, index) => {
    segments.push({ id: segmentId(index, "question"), text: question.question || "" });
    (question.options || []).forEach((option) => {
      segments.push({ id: segmentId(index, `option-${option.key}`), text: option.text || "" });
    });
    segments.push({ id: segmentId(index, "answer"), text: question.answerText || "" });
    segments.push({ id: segmentId(index, "explanation"), text: question.explanation || "" });
  });

  return segments;
}

function renderParts(parts, text, id) {
  if (!Array.isArray(parts) || !parts.length) {
    return renderText(text, id);
  }

  let cursor = 0;
  let html = "";

  for (const part of parts) {
    const partText = part.text || "";

    if (part.type === "text") {
      html += renderTextFragment(partText, id, cursor);
    } else {
      html += part.html || escapeHtml(partText);
    }

    cursor += partText.length;
  }

  return html;
}

function renderText(text, id) {
  return renderTextFragment(text || "", id, 0);
}

function renderTextFragment(text, id, baseOffset) {
  const value = text || "";
  const matches = state.spellMatches.get(id) || [];
  if (!matches.length) return escapeHtml(value);

  const mathRanges = findMathRanges(value);
  const cleanMatches = matches
    .map((match) => {
      const start = Math.max(baseOffset, match.offset);
      const end = Math.min(baseOffset + value.length, match.offset + match.length);
      return { ...match, offset: start - baseOffset, length: end - start };
    })
    .filter((match) => match.length > 0 && !intersectsAny(match.offset, match.offset + match.length, mathRanges))
    .sort((a, b) => a.offset - b.offset);
  let cursor = 0;
  let html = "";

  for (const match of cleanMatches) {
    const start = Math.max(0, match.offset);
    const end = Math.min(value.length, start + match.length);
    if (start < cursor || end <= start) continue;

    html += escapeHtml(value.slice(cursor, start));
    const tip = spellTip(match);
    html += `<span class="spell-error" tabindex="0" data-tip="${escapeAttr(tip)}">${escapeHtml(value.slice(start, end))}</span>`;
    cursor = end;
  }

  html += escapeHtml(value.slice(cursor));
  return html;
}

function spellTip(match) {
  const replacements = match.replacements?.length ? `Suggestion: ${match.replacements.join(", ")}` : "No suggestion";
  return `${match.message || match.shortMessage || "Possible spelling issue"} | ${replacements}`;
}

function findMathRanges(text) {
  const ranges = [];
  const patterns = [/\$\$[\s\S]*?\$\$/g, /\$[^$\n]+?\$/g, /\\\([\s\S]*?\\\)/g, /\\\[[\s\S]*?\\\]/g];

  for (const pattern of patterns) {
    let match;
    while ((match = pattern.exec(text)) !== null) {
      ranges.push([match.index, match.index + match[0].length]);
    }
  }

  return ranges;
}

function intersectsAny(start, end, ranges) {
  return ranges.some(([rangeStart, rangeEnd]) => start < rangeEnd && end > rangeStart);
}

function segmentId(questionIndex, part) {
  return `q${questionIndex}-${part}`;
}

function setStatus(message, tone = "") {
  statusLine.textContent = message;
  statusLine.classList.toggle("busy", tone === "busy");
  statusLine.classList.toggle("error", tone === "error");
}

function updateMetrics() {
  metricTotal.textContent = toBanglaNumber(state.questions.length);
  metricDuplicates.textContent = toBanglaNumber(state.duplicatePairs.length);
  metricSpell.textContent = toBanglaNumber(spellIssueCount());
}

function spellIssueCount() {
  return [...state.spellMatches.values()].reduce((sum, matches) => sum + matches.length, 0);
}

function renderEmpty(title, note = "") {
  cards.innerHTML = `
    <div class="empty-state">
      <div class="empty-state-inner">
        <i data-lucide="scan-text"></i>
        <div class="empty-state-title">${escapeHtml(title)}</div>
        <div class="empty-state-note">${escapeHtml(note)}</div>
      </div>
    </div>
  `;
  runIcons();
}

function runMathJax() {
  if (window.MathJax?.typesetPromise) {
    window.MathJax.typesetPromise([cards]).catch(() => {});
  }
}

function runIcons() {
  if (window.lucide?.createIcons) {
    window.lucide.createIcons();
  }
}

function formatFileSize(bytes) {
  if (!Number.isFinite(bytes) || bytes <= 0) return "0 KB";
  if (bytes < 1024 * 1024) return `${Math.max(1, Math.round(bytes / 1024))} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function toBanglaNumber(value) {
  return String(value ?? "").replace(/[0-9]/g, (digit) => bnDigits.get(digit) || digit);
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function escapeAttr(value) {
  return escapeHtml(value).replaceAll("`", "&#096;");
}

renderEmpty("Preview stream is empty.", "DOCX upload করলে প্রশ্ন, equation, image, answer ও solve এখানে দেখা যাবে।");
updateMetrics();
