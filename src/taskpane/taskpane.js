/* global Office, document, canExtractText, extractText, base64ToBytes */

const TYPE_META = {
  docx: { icon: "ms-Icon--WordDocument", bg: "var(--blue-bg)", fg: "var(--blue-fg)", label: "Word" },
  doc: { icon: "ms-Icon--WordDocument", bg: "var(--blue-bg)", fg: "var(--blue-fg)", label: "Word" },
  xlsx: { icon: "ms-Icon--ExcelDocument", bg: "var(--green-bg)", fg: "var(--green-fg)", label: "Excel" },
  xls: { icon: "ms-Icon--ExcelDocument", bg: "var(--green-bg)", fg: "var(--green-fg)", label: "Excel" },
  pptx: { icon: "ms-Icon--PowerPointDocument", bg: "var(--coral-bg)", fg: "var(--coral-fg)", label: "PowerPoint" },
  ppt: { icon: "ms-Icon--PowerPointDocument", bg: "var(--coral-bg)", fg: "var(--coral-fg)", label: "PowerPoint" },
  pdf: { icon: "ms-Icon--PDF", bg: "var(--red-bg)", fg: "var(--red-fg)", label: "PDF" },
  png: { icon: "ms-Icon--FileImage", bg: "var(--purple-bg)", fg: "var(--purple-fg)", label: "Image" },
  jpg: { icon: "ms-Icon--FileImage", bg: "var(--purple-bg)", fg: "var(--purple-fg)", label: "Image" },
  jpeg: { icon: "ms-Icon--FileImage", bg: "var(--purple-bg)", fg: "var(--purple-fg)", label: "Image" },
  default: { icon: "ms-Icon--Page", bg: "#F3F2F1", fg: "#605E5C", label: "File" },
};

let rawAttachments = [];
let view = "grid";
let query = "";
let activeType = "all";
let sortBy = "name-asc";
let selected = new Set();

// Content search: attachment id -> { status: "pending" | "done" | "unsupported" | "error", text, lower }
let contentIndex = new Map();
let indexGeneration = 0;
let indexingStarted = false;
const INDEX_CONCURRENCY = 2;

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    loadAttachments();
    wireStaticControls();
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, onItemChanged);
  }
});

function onItemChanged() {
  query = "";
  activeType = "all";
  selected = new Set();
  document.getElementById("searchInput").value = "";
  resetIndex();
  loadAttachments();
}

function wireStaticControls() {
  document.getElementById("gridBtn").addEventListener("click", () => setView("grid"));
  document.getElementById("listBtn").addEventListener("click", () => setView("list"));
  document.getElementById("searchInput").addEventListener("input", (e) => {
    query = e.target.value;
    if (query.trim()) startIndexing();
    updateSearchStatus();
    render();
  });
  document.getElementById("sortSelect").addEventListener("change", (e) => {
    sortBy = e.target.value;
    render();
  });
  document.getElementById("downloadSelectedBtn").addEventListener("click", downloadSelected);
}

function loadAttachments() {
  const item = Office.context.mailbox.item;

  // Read-mode items expose attachments directly.
  if (item.attachments) {
    rawAttachments = (item.attachments || []).filter((a) => !a.isInline);
    afterLoad();
    return;
  }

  // Compose-mode items require the async API.
  item.getAttachmentsAsync((result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      rawAttachments = (result.value || []).filter((a) => !a.isInline);
    }
    afterLoad();
  });
}

function afterLoad() {
  document.getElementById("attachmentCount").textContent = rawAttachments.length;
  document.getElementById("statusMessage").hidden = true;
  document.getElementById("galleryContainer").hidden = false;

  if (rawAttachments.length === 0) {
    document.getElementById("statusMessage").hidden = false;
    document.getElementById("statusMessage").textContent = "This message has no attachments.";
    document.getElementById("galleryContainer").hidden = true;
    document.getElementById("galleryContainer").innerHTML = "";
    document.getElementById("controls").hidden = true;
    document.getElementById("typeChips").innerHTML = "";
    updateBulkToolbar();
    return;
  }

  document.getElementById("controls").hidden = false;
  updateSearchStatus();
  renderChips();
  render();
}

function getExt(name) {
  const parts = name.split(".");
  return parts.length > 1 ? parts.pop().toLowerCase() : "";
}

function meta(name) {
  return TYPE_META[getExt(name)] || TYPE_META.default;
}

function formatSize(bytes) {
  if (!bytes && bytes !== 0) return "";
  if (bytes < 1024) return bytes + " B";
  if (bytes < 1024 * 1024) return Math.round(bytes / 1024) + " KB";
  return (bytes / (1024 * 1024)).toFixed(1) + " MB";
}

function renderChips() {
  const types = ["all", ...new Set(rawAttachments.map((a) => getExt(a.name)))];
  const chipsEl = document.getElementById("typeChips");
  chipsEl.innerHTML = "";
  types.forEach((t) => {
    const btn = document.createElement("button");
    btn.className = "chip" + (t === activeType ? " chip--active" : "");
    btn.textContent = t === "all" ? "All" : meta("x." + t).label;
    btn.addEventListener("click", () => {
      activeType = t;
      renderChips();
      render();
    });
    chipsEl.appendChild(btn);
  });
}

// --- Search (names + document contents) ---------------------------------

function searchTerms() {
  return query.toLowerCase().split(/\s+/).filter(Boolean);
}

function contentEntry(a) {
  const e = contentIndex.get(a.id);
  return e && e.status === "done" ? e : null;
}

// Every search word must appear in the file name or in the file's text.
function matchesQuery(a, terms) {
  if (terms.length === 0) return true;
  const name = a.name.toLowerCase();
  const entry = contentEntry(a);
  return terms.every((t) => name.includes(t) || (entry && entry.lower.includes(t)));
}

function snippetFor(a, terms) {
  const entry = contentEntry(a);
  if (!entry || terms.length === 0) return "";
  const name = a.name.toLowerCase();
  const term = terms.find((t) => !name.includes(t) && entry.lower.includes(t));
  if (!term || entry.lower.length !== entry.text.length) return "";
  const i = entry.lower.indexOf(term);
  const start = Math.max(0, i - 30);
  const end = Math.min(entry.text.length, i + term.length + 50);
  return (
    (start > 0 ? "…" : "") +
    escapeHtml(entry.text.slice(start, i)) +
    "<mark>" + escapeHtml(entry.text.slice(i, i + term.length)) + "</mark>" +
    escapeHtml(entry.text.slice(i + term.length, end)) +
    (end < entry.text.length ? "…" : "")
  );
}

function filteredSorted() {
  const terms = searchTerms();
  return rawAttachments
    .filter((a) => (activeType === "all" || getExt(a.name) === activeType) && matchesQuery(a, terms))
    .sort(compareAttachments);
}

function compareAttachments(a, b) {
  if (sortBy === "name-asc") return a.name.localeCompare(b.name);
  if (sortBy === "name-desc") return b.name.localeCompare(a.name);
  if (sortBy === "type") return meta(a.name).label.localeCompare(meta(b.name).label) || a.name.localeCompare(b.name);
  if (sortBy === "size-desc") return (b.size || 0) - (a.size || 0);
  return 0;
}

function resetIndex() {
  indexGeneration++;
  contentIndex = new Map();
  indexingStarted = false;
}

function fetchAttachmentBytes(attachment) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.getAttachmentContentAsync(attachment.id, (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        reject(new Error(result.error && result.error.message));
        return;
      }
      if (result.value.format !== Office.MailboxEnums.AttachmentContentFormat.Base64) {
        reject(new Error("Attachment content not available as base64"));
        return;
      }
      resolve(base64ToBytes(result.value.content));
    });
  });
}

// Reads attachment contents in the background the first time the user searches.
function startIndexing() {
  if (indexingStarted) return;
  indexingStarted = true;
  const gen = indexGeneration;

  rawAttachments.forEach((a) => {
    const readable = canExtractText(a.name, a.size);
    contentIndex.set(a.id, { status: readable ? "pending" : "unsupported", text: "", lower: "" });
  });

  const queue = rawAttachments.filter((a) => contentIndex.get(a.id).status === "pending");
  let next = 0;

  const runNext = () => {
    if (gen !== indexGeneration || next >= queue.length) return;
    const a = queue[next++];
    fetchAttachmentBytes(a)
      .then((bytes) => extractText(a.name, bytes))
      .then((text) => {
        if (gen !== indexGeneration) return;
        contentIndex.set(a.id, { status: "done", text, lower: text.toLowerCase() });
      })
      .catch(() => {
        if (gen !== indexGeneration) return;
        contentIndex.set(a.id, { status: "error", text: "", lower: "" });
      })
      .then(() => {
        if (gen !== indexGeneration) return;
        updateSearchStatus();
        if (query.trim()) render();
        runNext();
      });
  };

  for (let i = 0; i < INDEX_CONCURRENCY; i++) runNext();
}

function updateSearchStatus() {
  const el = document.getElementById("searchStatus");
  if (!indexingStarted || !query.trim()) {
    el.hidden = true;
    return;
  }
  const entries = rawAttachments.map((a) => contentIndex.get(a.id)).filter(Boolean);
  const pending = entries.filter((e) => e.status === "pending").length;
  const done = entries.filter((e) => e.status === "done").length;
  const nameOnly = entries.length - done - pending;

  el.hidden = false;
  if (pending > 0) {
    el.textContent = `Reading file contents… ${done} of ${done + pending} done`;
  } else if (nameOnly > 0) {
    el.textContent = `Searched names and contents of ${done} file${done === 1 ? "" : "s"}; ${nameOnly} matched by name only`;
  } else {
    el.textContent = "Searched names and contents";
  }
}

// --- Rendering -----------------------------------------------------------

function setView(next) {
  view = next;
  document.getElementById("gridBtn").classList.toggle("view-btn--active", next === "grid");
  document.getElementById("gridBtn").setAttribute("aria-pressed", next === "grid");
  document.getElementById("listBtn").classList.toggle("view-btn--active", next === "list");
  document.getElementById("listBtn").setAttribute("aria-pressed", next === "list");
  render();
}

function buildCard(a, terms) {
  const m = meta(a.name);
  const isSelected = selected.has(a.id);
  const snippet = snippetFor(a, terms);

  const card = document.createElement("div");
  card.className = "file-card" + (view === "list" ? " list-row" : "") + (isSelected ? " file-card--selected" : "");

  const checkbox = document.createElement("div");
  checkbox.className = "file-card__checkbox";
  checkbox.innerHTML = isSelected ? '<i class="ms-Icon ms-Icon--CheckMark" aria-hidden="true"></i>' : "";

  const icon = document.createElement("div");
  icon.className = "file-icon";
  icon.style.background = m.bg;
  icon.innerHTML = `<i class="ms-Icon ${m.icon}" style="font-size:16px;color:${m.fg};" aria-hidden="true"></i>`;

  const textWrap = document.createElement("div");
  textWrap.className = "file-text";
  textWrap.innerHTML =
    `<p class="file-name" title="${escapeHtml(a.name)}">${escapeHtml(a.name)}</p>` +
    `<p class="file-meta">${formatSize(a.size)}</p>` +
    (snippet ? `<p class="file-snippet">${snippet}</p>` : "");

  card.appendChild(checkbox);
  card.appendChild(icon);
  card.appendChild(textWrap);

  card.addEventListener("click", () => {
    if (selected.has(a.id)) selected.delete(a.id);
    else selected.add(a.id);
    render();
  });

  return card;
}

function render() {
  const container = document.getElementById("galleryContainer");
  container.className = "gallery-container " + view;
  const items = filteredSorted();

  if (items.length === 0) {
    container.innerHTML = '<div class="empty-state">No attachments match</div>';
    updateBulkToolbar();
    return;
  }

  const terms = searchTerms();
  container.innerHTML = "";
  items.forEach((a) => container.appendChild(buildCard(a, terms)));
  updateBulkToolbar();
}

function updateBulkToolbar() {
  const toolbar = document.getElementById("bulkToolbar");
  if (selected.size > 0) {
    toolbar.hidden = false;
    document.getElementById("selectionCount").textContent = selected.size + " selected";
  } else {
    toolbar.hidden = true;
  }
}

// --- Download ------------------------------------------------------------

function downloadSelected() {
  const chosen = rawAttachments.filter((a) => selected.has(a.id));
  chosen.forEach((a) => downloadAttachment(a));
}

function downloadAttachment(attachment) {
  Office.context.mailbox.item.getAttachmentContentAsync(attachment.id, (result) => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) return;

    const content = result.value;

    if (content.format !== Office.MailboxEnums.AttachmentContentFormat.Base64) {
      // Url format: content.content is a URL Outlook can resolve directly.
      window.open(content.content, "_blank");
      return;
    }

    const url = URL.createObjectURL(new Blob([base64ToBytes(content.content)]));
    const link = document.createElement("a");
    link.href = url;
    link.download = attachment.name;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  });
}

function escapeHtml(str) {
  const div = document.createElement("div");
  div.textContent = str;
  return div.innerHTML.replace(/"/g, "&quot;").replace(/'/g, "&#39;");
}
