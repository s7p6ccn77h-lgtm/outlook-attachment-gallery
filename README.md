# Attachment Gallery — Outlook add-in

A task pane add-in that replaces Outlook's native attachment strip with a
searchable, filterable, sortable gallery. Matches the mockup: grid/list
toggle, file-type color coding, type filter chips, sort by name/type/size,
multi-select, and bulk download.

## What's here

```
manifest.xml              add-in manifest (classic XML) — declares the task
                           pane and the ribbon button that opens it
src/taskpane/taskpane.html the pane's markup
src/taskpane/taskpane.css  styling (matches the mockup's look)
src/taskpane/taskpane.js   Office.js logic — reads live attachments off the
                           open email, renders the gallery, handles
                           search/filter/sort/select/download
assets/                    placeholder icons (swap for real branding)
package.json               local dev scripts
```

## Run it locally

This has to run on your machine — it needs a real Outlook client (desktop,
web, or new Outlook) and a trusted local HTTPS server, neither of which
exist in this chat environment.

1. **Install dependencies**
   ```
   npm install
   ```

2. **Trust a local dev certificate** (one-time; required because Outlook
   only loads add-ins over HTTPS)
   ```
   npm run dev-certs
   ```

3. **Serve the files over HTTPS on port 3000**
   ```
   npm start
   ```
   Leave this running — `manifest.xml` points at `https://localhost:3000`.

4. **Sideload into Outlook**, in a second terminal:
   ```
   npm run sideload
   ```
   This opens Outlook and installs the add-in for your account automatically.
   If you'd rather do it by hand: in Outlook on the web, go to
   **Settings → Manage add-ins → My add-ins → Add a custom add-in → Add from
   file**, and pick `manifest.xml`.

5. Open any received email with attachments. You'll see an **Attachments**
   group with a **Gallery view** button on the ribbon — click it to open the
   pane.

## How the data binding works

- `Office.context.mailbox.item.attachments` gives the attachment list
  directly when reading a message; `getAttachmentsAsync` is the fallback for
  compose mode.
- Inline images (signatures, embedded logos) are filtered out via
  `attachment.isInline` so they don't clutter the gallery.
- Downloads use `getAttachmentContentAsync`, which returns the file as
  base64 (or a URL for some providers) — that's decoded into a `Blob` and
  triggered as a browser download.

## Group by sender (whole thread)

The toggle above the gallery pulls in attachments from every message in the
current conversation, not just the one you have open, and groups them by
sender. It works differently from the rest of the add-in:

- It uses EWS (`Office.context.mailbox.makeEwsRequestAsync` with a
  `GetConversationItems` request keyed on `item.conversationId`), since
  `item.attachments` only ever covers the open message.
- **Exchange only.** POP/IMAP accounts and some consumer Outlook.com
  configurations don't support EWS from an add-in; the toggle shows a clear
  error state ("Couldn't load the full thread...") rather than failing
  silently.
- Files that belong to the **open** message stay fully selectable and
  downloadable, same as always. Files from **other** messages in the thread
  are shown so you can see what's there, but are view-only — Office.js can
  only fetch attachment *content* for the item you currently have open, not
  arbitrary other messages. Opening that email lets you download it.
- This is the one part of the add-in that hasn't been tested against a real
  mailbox (no way to run a live Exchange session from where this was built).
  The SOAP request shape and response parsing follow Microsoft's documented
  EWS schema, but if it doesn't work first try, the likely fix is in
  `parseConversationAttachments()` / `buildGetConversationItemsRequest()` in
  `taskpane.js` — that's the first place to add a `console.log` of the raw
  EWS response and compare against what actually came back.

## Known limitations to fix before shipping

- **Permissions**: manifest requests `ReadWriteItem`. If you add features
  like moving attachments to OneDrive, you'll need broader Graph permissions
  (and an Azure AD app registration) — Office.js alone can't call Graph.
- **Compose-mode download** isn't wired up — `getAttachmentContentAsync`
  behaves differently before a message is sent, and most galleries only need
  read mode anyway.
- **No backend/auth yet** — this is a pure client-side add-in. Fine for
  personal use or internal deployment; before AppSource you'd also want the
  newer JSON "unified manifest" format, which AppSource now prefers over
  this classic XML one.
- **Hosting**: `localhost:3000` only works for local testing — see "Next
  steps" below for moving to GitHub Pages.

## Next steps

- Host on GitHub Pages and update every URL in `manifest.xml` from
  `localhost:3000` to the real domain
- Save-to-OneDrive via Microsoft Graph — deliberately not built yet; needs
  an Azure AD app registration (your own Microsoft/Azure admin login, not
  something that can be set up on your behalf) before the code side is
  worth writing
