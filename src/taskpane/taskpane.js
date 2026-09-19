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

// Thread mode: aggregate attachments across the whole conversation, grouped by sender.
// Uses EWS (makeEwsRequestAsync) since Office.js's item.attachments only covers the open
// message. Exchange-only — not available for POP/IMAP or some consumer Outlook.com accounts.
let threadMode = false;
let threadGroups = null; // [{ senderName, senderEmail, items: [{ name, size, sourceAttachment }] }]
let threadStatus = ""; // "" | "loading" | "error"

const BUILD = "3";
let itemChangedCount = 0;
let handlerStatus = "not registered";

function updateDebug() {
  const el = document.getElementById("debugInfo");
  if (!el) return;
  const item = Office.context.mailbox.item;
  const subject = item && item.subject ? item.subject : "(none)";
  el.textContent = `build ${BUILD} | item-changed listener: ${handlerStatus} | switches seen: ${itemChangedCount} | subject: ${subject}`;
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    loadAttachments();
    wireStaticControls();
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, onItemChanged, (r) => {
      handlerStatus = r.status === Office.AsyncResultStatus.Succeeded ? "registered" : "FAILED: " + (r.error && r.error.message);
      updateDebug();
    });
    updateDebug();
  }
});

function onItemChanged() {
  itemChangedCount++;
  query = "";
  activeType = "all";
  selected = new Set();
  document.getElementById("searchInput").value = "";

  if (threadMode) {
    loadThread();
  }
  loadAttachments();
  updateDebug();
}

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
  document.getElementById("threadToggle").addEventListener("change", (e) => setThreadMode(e.target.checked));
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

  items = items.slice().sort(compareAttachments);

  return items;
}

function compareAttachments(a, b) {
  if (sortBy === "name-asc") return a.name.localeCompare(b.name);
  if (sortBy === "name-desc") return b.name.localeCompare(a.name);
  if (sortBy === "type") return meta(a.name).label.localeCompare(meta(b.name).label) || a.name.localeCompare(b.name);
  if (sortBy === "size-desc") return (b.size || 0) - (a.size || 0);
  return 0;
}

function setView(next) {
  view = next;
  document.getElementById("gridBtn").classList.toggle("view-btn--active", next === "grid");
  document.getElementById("gridBtn").setAttribute("aria-pressed", next === "grid");
  document.getElementById("listBtn").classList.toggle("view-btn--active", next === "list");
  document.getElementById("listBtn").setAttribute("aria-pressed", next === "list");
  render();
}

function buildCard(a, opts) {
  const options = opts || {};
  const m = meta(a.name);
  const isSelected = options.selectable !== false && selected.has(a.id);

  const card = document.createElement("div");
  card.className =
    "file-card" +
    (view === "list" ? " list-row" : "") +
    (isSelected ? " file-card--selected" : "") +
    (options.selectable === false ? " file-card--disabled" : "");
  if (options.disabledNote) card.title = options.disabledNote;

  const checkbox = document.createElement("div");
  checkbox.className = "file-card__checkbox";
  if (options.selectable !== false) {
    checkbox.innerHTML = isSelected ? '<i class="ms-Icon ms-Icon--CheckMark" aria-hidden="true"></i>' : "";
  }

  const icon = document.createElement("div");
  icon.className = "file-icon";
  icon.style.background = m.bg;
  icon.innerHTML = `<i class="ms-Icon ${m.icon}" style="font-size:16px;color:${m.fg};" aria-hidden="true"></i>`;

  const textWrap = document.createElement("div");
  textWrap.className = "file-text";
  textWrap.innerHTML = `<p class="file-name">${escapeHtml(a.name)}</p><p class="file-meta">${formatSize(a.size)}</p>`;

  card.appendChild(checkbox);
  card.appendChild(icon);
  card.appendChild(textWrap);

  if (options.selectable !== false) {
    card.addEventListener("click", () => {
      if (selected.has(a.id)) selected.delete(a.id);
      else selected.add(a.id);
      updateBulkToolbar();
      render();
    });
  }

  return card;
}

function render() {
  const container = document.getElementById("galleryContainer");
  container.className = "gallery-container " + view;

  if (threadMode) {
    renderThreadView(container);
    return;
  }

  const items = filteredSorted();

  if (items.length === 0) {
    container.innerHTML = '<div class="empty-state">No attachments match</div>';
    updateBulkToolbar();
    return;
  }

  container.innerHTML = "";
  items.forEach((a) => container.appendChild(buildCard(a)));
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

// --- Thread mode ---------------------------------------------------------

function setThreadMode(enabled) {
  threadMode = enabled;
  document.getElementById("sortSelect").disabled = enabled;
  if (enabled) {
    loadThread();
  } else {
    render();
  }
}

function loadThread() {
  threadStatus = "loading";
  threadGroups = null;
  render();

  const item = Office.context.mailbox.item;
  const conversationId = item.conversationId;

  if (!conversationId || !Office.context.mailbox.makeEwsRequestAsync) {
    threadStatus = "error";
    render();
    return;
  }

  const soap = buildGetConversationItemsRequest(conversationId);
  Office.context.mailbox.makeEwsRequestAsync(soap, (result) => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      threadStatus = "error";
      render();
      return;
    }
    try {
      threadGroups = parseConversationAttachments(result.value);
      threadStatus = "";
    } catch (e) {
      threadStatus = "error";
    }
    render();
  });
}

function buildGetConversationItemsRequest(conversationId) {
  const escapedId = escapeXml(conversationId);
  return (
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" ' +
    'xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
    "<soap:Header><t:RequestServerVersion Version=\"Exchange2013\" /></soap:Header>" +
    "<soap:Body>" +
    '<GetConversationItems xmlns="http://schemas.microsoft.com/exchange/services/2006/messages" ' +
    'xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
    "<ItemShape>" +
    "<t:BaseShape>IdOnly</t:BaseShape>" +
    "<t:AdditionalProperties>" +
    '<t:FieldURI FieldURI="message:Sender" />' +
    '<t:FieldURI FieldURI="item:Attachments" />' +
    "</t:AdditionalProperties>" +
    "</ItemShape>" +
    "<Conversations>" +
    "<t:Conversation>" +
    `<t:ConversationId Id="${escapedId}" />` +
    "</t:Conversation>" +
    "</Conversations>" +
    "</GetConversationItems>" +
    "</soap:Body>" +
    "</soap:Envelope>"
  );
}

function escapeXml(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function firstChildByLocalName(node, localName) {
  for (const child of node.childNodes) {
    if (child.nodeType === 1 && child.localName === localName) return child;
  }
  return null;
}

// Matches a thread-view attachment back to a live attachment on the open message
// (by name + size), since the open item is the only one we can download from directly
// via Office.js — everything else in the thread is read-only until you open that email.
function matchLiveAttachment(name, size) {
  return rawAttachments.find((a) => a.name === name && (a.size || 0) === (size || 0)) || null;
}

function parseConversationAttachments(xmlString) {
  const doc = new DOMParser().parseFromString(xmlString, "text/xml");
  const groups = new Map();

  const messageNodes = Array.from(doc.getElementsByTagNameNS("*", "Message")).concat(
    Array.from(doc.getElementsByTagNameNS("*", "MeetingMessage"))
  );

  messageNodes.forEach((msgNode) => {
    const senderNode = firstChildByLocalName(msgNode, "Sender");
    let senderName = "Unknown sender";
    let senderEmail = "";
    if (senderNode) {
      const mailboxNode = firstChildByLocalName(senderNode, "Mailbox");
      if (mailboxNode) {
        const nameNode = firstChildByLocalName(mailboxNode, "Name");
        const emailNode = firstChildByLocalName(mailboxNode, "EmailAddress");
        if (nameNode && nameNode.textContent) senderName = nameNode.textContent;
        if (emailNode && emailNode.textContent) senderEmail = emailNode.textContent;
      }
    }

    const attachmentsNode = firstChildByLocalName(msgNode, "Attachments");
    if (!attachmentsNode) return;

    Array.from(attachmentsNode.childNodes)
      .filter((n) => n.nodeType === 1 && n.localName === "FileAttachment")
      .forEach((att) => {
        const isInlineNode = firstChildByLocalName(att, "IsInline");
        if (isInlineNode && isInlineNode.textContent === "true") return;

        const nameNode = firstChildByLocalName(att, "Name");
        const sizeNode = firstChildByLocalName(att, "Size");
        const name = nameNode ? nameNode.textContent : "(unnamed attachment)";
        const size = sizeNode ? parseInt(sizeNode.textContent, 10) : 0;

        const key = senderEmail || senderName;
        if (!groups.has(key)) groups.set(key, { senderName, senderEmail, items: [] });
        groups.get(key).items.push({ name, size, sourceAttachment: matchLiveAttachment(name, size) });
      });
  });

  return Array.from(groups.values())
    .filter((g) => g.items.length > 0)
    .sort((a, b) => a.senderName.localeCompare(b.senderName));
}

function renderThreadView(container) {
  container.innerHTML = "";

  if (threadStatus === "loading") {
    container.innerHTML = '<div class="empty-state">Loading the whole thread&hellip;</div>';
    updateBulkToolbar();
    return;
  }

  if (threadStatus === "error" || !threadGroups) {
    container.innerHTML =
      '<div class="empty-state">Couldn’t load the full thread (this needs an Exchange mailbox). ' +
      "Turn off “Group by sender” to see just this message.</div>";
    updateBulkToolbar();
    return;
  }

  const lowerQuery = query.toLowerCase();
  let anyRendered = false;

  threadGroups.forEach((group) => {
    const items = group.items
      .filter((it) => activeType === "all" || getExt(it.name) === activeType)
      .filter((it) => it.name.toLowerCase().includes(lowerQuery))
      .slice()
      .sort(compareAttachments);

    if (items.length === 0) return;
    anyRendered = true;

    const heading = document.createElement("div");
    heading.className = "thread-group-heading";
    heading.textContent = `${group.senderName} (${items.length})`;
    container.appendChild(heading);

    const groupEl = document.createElement("div");
    groupEl.className = "gallery-container " + view;
    items.forEach((it) => {
      if (it.sourceAttachment) {
        groupEl.appendChild(buildCard(it.sourceAttachment));
      } else {
        groupEl.appendChild(
          buildCard(it, { selectable: false, disabledNote: "Open that email to download this file" })
        );
      }
    });
    container.appendChild(groupEl);
  });

  if (!anyRendered) {
    container.innerHTML = '<div class="empty-state">No attachments match</div>';
  }

  updateBulkToolbar();
}
