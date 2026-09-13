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

## Known limitations to fix before shipping

- **Icons are placeholders** — swap `assets/icon-*.png` for real artwork at
  16/32/64/80/128px.
- **Permissions**: manifest requests `ReadWriteItem`. If you add features
  like moving attachments to OneDrive, you'll need broader Graph permissions
  and a proper backend — Office.js alone can't call Graph.
- **Compose-mode download** isn't wired up — `getAttachmentContentAsync`
  behaves differently before a message is sent, and most galleries only need
  read mode anyway.
- **No backend/auth yet** — this is a pure client-side add-in. Fine for
  personal use or internal deployment; before AppSource you'd also want the
  newer JSON "unified manifest" format, which AppSource now prefers over
  this classic XML one.
- **Hosting**: `localhost:3000` only works for local testing. For real
  deployment, host the `src/taskpane` files (HTTPS, valid cert) and update
  every URL in `manifest.xml` to match.

## Next steps

- Swap in real icons and a hosted domain
- Add a "group by sender" view
- Wire up drag-to-reorder or save-to-OneDrive via Microsoft Graph
