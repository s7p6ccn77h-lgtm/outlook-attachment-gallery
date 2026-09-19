# Attachment Gallery — Outlook add-in

A task pane add-in that replaces Outlook's native attachment strip with a
searchable, filterable, sortable gallery. Matches the mockup: grid/list
toggle, file-type color coding, type filter chips, sort by name/type/size,
multi-select, bulk download, and search across both file names and the
text inside the documents.

Hosted on GitHub Pages at
https://s7p6ccn77h-lgtm.github.io/outlook-attachment-gallery/ — repo:
https://github.com/s7p6ccn77h-lgtm/outlook-attachment-gallery

## What's here

```
manifest.xml              add-in manifest (classic XML) — declares the task
                           pane and the ribbon button that opens it
src/taskpane/taskpane.html the pane's markup
src/taskpane/taskpane.css  styling (matches the mockup's look)
src/taskpane/taskpane.js   Office.js logic — reads live attachments off the
                           open email, renders the gallery, handles
                           search/filter/sort/select/download
src/taskpane/extract.js    pulls plain text out of attachments for content
                           search (docx/xlsx/pptx via JSZip, pdf via PDF.js)
assets/                    icons (blue rounded square, gallery-grid glyph)
package.json               local dev scripts
```

## Install it in Outlook

`manifest.xml` now points at the hosted GitHub Pages URLs above, so no local
server is needed to try it:

- **Outlook on the web**: open the add-ins dialog (the apps grid icon on
  the left rail or ribbon, or https://aka.ms/olksideload), then **My
  add-ins → Custom add-ins → Add a custom add-in → Add from file**, and pick
  `manifest.xml` from this repo. Personal accounts don't offer "Add from
  URL". Add-ins installed this way sync to the desktop apps after a
  restart.
- **Outlook for Mac**: `npm run sideload` does not support it
  ("Sideload to the Outlook app is not supported"), so install via the web
  route above and restart the app.
- After changing `manifest.xml`, remove and re-add the add-in — Outlook
  caches the old manifest. Bump `?v=N` on the URLs and the manifest
  `<Version>` to defeat cached copies of the pane itself.

Then open any received email with attachments — you'll see an
**Attachments** group with a **Gallery view** button on the ribbon.

### Making changes

Any edit to `src/taskpane/*` needs to be pushed to `main` before Outlook
sees it (GitHub Pages redeploys automatically, usually within a minute).
For faster local iteration, run `npm run dev-certs` then `npm start` to
serve over `https://localhost:3000`, and temporarily repoint the URLs in
`manifest.xml` back to `localhost:3000` while you sideload it — just
remember to point them back at the GitHub Pages URLs (or `git checkout
manifest.xml`) before pushing.

## How the data binding works

- `Office.context.mailbox.item.attachments` gives the attachment list
  directly when reading a message; `getAttachmentsAsync` is the fallback for
  compose mode.
- Inline images (signatures, embedded logos) are filtered out via
  `attachment.isInline` so they don't clutter the gallery.
- Downloads use `getAttachmentContentAsync`, which returns the file as
  base64 (or a URL for some providers) — that's decoded into a `Blob` and
  triggered as a browser download.

## Searching inside documents

Typing in the search box matches file names immediately. The first time you
search, the pane also reads each attachment's contents in the background
(two at a time) and re-filters as they finish; matches from inside a file
show a highlighted snippet on the card. Every word you type must appear
somewhere in the name or contents (case-insensitive).

- Readable: `.docx`, `.xlsx` (cell text and sheet names), `.pptx` (slides
  and notes), text-based `.pdf` (first 150 pages), and `.txt`, `.csv`,
  `.tsv`, `.md`, `.json`, `.xml`, `.log`, `.html`.
- Name-only: images, scanned PDFs (no text layer, would need OCR), legacy
  `.doc`/`.xls`/`.ppt`, password-protected files, and anything over 25 MB.
- Everything happens inside the task pane — file contents are never sent
  anywhere. JSZip and PDF.js are loaded from cdnjs the first time you
  search, so the pane needs internet access for that.
- Only the open message's attachments are read, and the index is discarded
  when you switch emails.

## Known limitations to fix before shipping

- **Pane can't follow email switches on personal outlook.live.com.** The
  manifest declares `SupportsPinning`, but personal Outlook on the web
  offers no pin control, and Outlook freezes the pane's open email
  (neither `ItemChanged` nor polling `item.itemId` ever sees a change).
  Close and reopen **Gallery view** for each email there. Pinning is
  supported in Outlook for Mac, Windows, and work/school Outlook on the
  web, where the refresh-on-switch logic should work.
- **Permissions**: manifest requests `ReadItem`, which is enough to list
  and read the open message's attachments. Moving attachments to OneDrive
  would need Graph permissions and an Azure AD app registration — Office.js
  alone can't call Graph.
- **Compose-mode download** isn't wired up — `getAttachmentContentAsync`
  behaves differently before a message is sent, and most galleries only need
  read mode anyway.
- **No backend/auth yet** — this is a pure client-side add-in. Fine for
  personal use or internal deployment; before AppSource you'd also want the
  newer JSON "unified manifest" format, which AppSource now prefers over
  this classic XML one.

## Next steps

- Save-to-OneDrive via Microsoft Graph — deliberately not built yet; needs
  an Azure AD app registration (your own Microsoft/Azure admin login, not
  something that can be set up on your behalf) before the code side is
  worth writing
