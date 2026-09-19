# Attachment Gallery — Outlook add-in

A task pane add-in that replaces Outlook's native attachment strip with a
searchable, filterable, sortable gallery. Matches the mockup: grid/list
toggle, file-type color coding, type filter chips, sort by name/type/size,
multi-select, bulk download, and a whole-thread "group by sender" view.

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
                           search/filter/sort/select/download, and the
                           EWS-based whole-thread grouping
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

- **Pane can't follow email switches on personal outlook.live.com.** The
  manifest declares `SupportsPinning`, but personal Outlook on the web
  offers no pin control, and Outlook freezes the pane's open email
  (neither `ItemChanged` nor polling `item.itemId` ever sees a change).
  Close and reopen **Gallery view** for each email there. Pinning is
  supported in Outlook for Mac, Windows, and work/school Outlook on the
  web, where the refresh-on-switch logic should work.
- **Permissions**: manifest requests `ReadWriteMailbox`, which
  `makeEwsRequestAsync` (whole-thread view) requires. Moving attachments to
  OneDrive would additionally need Graph permissions and an Azure AD app
  registration — Office.js alone can't call Graph.
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
