/* global Office, document */

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

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    loadAttachments();
    wireStaticControls();
  }
});

function wireStaticControls() {
  document.getElementById("gridBtn").addEventListener("click", () => setView("grid"));
  document.getElementById("listBtn").addEventListener("click", () => setView("list"));
  document.getElementById("searchInput").addEventListener("input", (e) => {
    query = e.target.value;
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
    return;
  }

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

function filteredSorted() {
  let items = rawAttachments.filter((a) => {
    const matchesType = activeType === "all" || getExt(a.name) === activeType;
    const matchesQuery = a.name.toLowerCase().includes(query.toLowerCase());
    return matchesType && matchesQuery;
  });

  items = items.slice().sort((a, b) => {
    if (sortBy === "name-asc") return a.name.localeCompare(b.name);
    if (sortBy === "name-desc") return b.name.localeCompare(a.name);
    if (sortBy === "type") return meta(a.name).label.localeCompare(meta(b.name).label) || a.name.localeCompare(b.name);
    if (sortBy === "size-desc") return (b.size || 0) - (a.size || 0);
    return 0;
  });

  return items;
}

function setView(next) {
  view = next;
  document.getElementById("gridBtn").classList.toggle("view-btn--active", next === "grid");
  document.getElementById("gridBtn").setAttribute("aria-pressed", next === "grid");
  document.getElementById("listBtn").classList.toggle("view-btn--active", next === "list");
  document.getElementById("listBtn").setAttribute("aria-pressed", next === "list");
  render();
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

  container.innerHTML = "";
  items.forEach((a) => {
    const m = meta(a.name);
    const isSelected = selected.has(a.id);

    const card = document.createElement("div");
    card.className = "file-card" + (view === "list" ? " list-row" : "") + (isSelected ? " file-card--selected" : "");

    const checkbox = document.createElement("div");
    checkbox.className = "file-card__checkbox";
    checkbox.innerHTML = isSelected ? '<i class="ms-Icon ms-Icon--CheckMark" aria-hidden="true"></i>' : "";

    const icon = document.createElement("div");
    icon.className = "file-icon";
    icon.style.background = m.bg;
    icon.innerHTML = `<i class="ms-Icon ${m.icon}" style="font-size:16px;color:${m.fg};" aria-hidden="true"></i>`;

    if (view === "list") {
      const textWrap = document.createElement("div");
      textWrap.className = "file-text";
      textWrap.innerHTML = `<p class="file-name">${escapeHtml(a.name)}</p><p class="file-meta">${formatSize(a.size)}</p>`;
      card.appendChild(checkbox);
      card.appendChild(icon);
      card.appendChild(textWrap);
    } else {
      card.appendChild(checkbox);
      card.appendChild(icon);
      card.innerHTML += `<p class="file-name">${escapeHtml(a.name)}</p><p class="file-meta">${formatSize(a.size)}</p>`;
    }

    card.addEventListener("click", () => {
      if (selected.has(a.id)) selected.delete(a.id);
      else selected.add(a.id);
      updateBulkToolbar();
      render();
    });

    container.appendChild(card);
  });

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

function downloadSelected() {
  const chosen = rawAttachments.filter((a) => selected.has(a.id));
  chosen.forEach((a) => downloadAttachment(a));
}

function downloadAttachment(attachment) {
  Office.context.mailbox.item.getAttachmentContentAsync(attachment.id, (result) => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) return;

    const content = result.value;
    let blob;

    if (content.format === Office.MailboxEnums.AttachmentContentFormat.Base64) {
      const byteChars = atob(content.content);
      const byteNumbers = new Array(byteChars.length);
      for (let i = 0; i < byteChars.length; i++) byteNumbers[i] = byteChars.charCodeAt(i);
      blob = new Blob([new Uint8Array(byteNumbers)]);
    } else {
      // Url format: content.content is a URL Outlook can resolve directly.
      window.open(content.content, "_blank");
      return;
    }

    const url = URL.createObjectURL(blob);
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
  return div.innerHTML;
}
