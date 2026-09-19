# Third-party notices

The add-in bundles the following unmodified libraries in `src/taskpane/vendor/`
so that no third-party server ever sees attachment content. Each is used under
its own license, not the license in `LICENSE`.

| Library | Version | License | Full license text | Source |
| --- | --- | --- | --- | --- |
| JSZip | 3.10.1 | MIT (JSZip is dual-licensed MIT or GPLv3; used here under MIT) | `src/taskpane/vendor/LICENSE-jszip.txt` | https://github.com/Stuk/jszip/tree/v3.10.1 |
| PDF.js (Mozilla), "legacy" build | 6.3.289 | Apache-2.0 | `src/taskpane/vendor/LICENSE-pdfjs.txt` | https://github.com/mozilla/pdf.js/tree/v6.3.289 |

JSZip was downloaded from cdnjs (`cdnjs.cloudflare.com/ajax/libs/`); PDF.js
comes from the `pdfjs-dist@6.3.289` npm package (`legacy/build/`). All are
byte-for-byte the published builds. SHA-256 checksums for verification:

```
acc7e41455a80765b5fd9c7ee1b8078a6d160bbbca455aeae854de65c947d59e  jszip.min.js
f401927e692efc7735e0cd528c490d0dd31b7f0972c122b7040df805be45cce4  pdf.min.mjs
a33cfe728c584fdba4fcc1fd54bcdc2f9f2f13889ddbb5b2bd1d0f8cbe49b84e  pdf.worker.min.mjs
```

JSZip's own distribution also embeds `pako` (MIT / Zlib); see its license file
for the complete notices.

## Loaded at runtime, not bundled

`office.js` (Microsoft) is loaded from `https://appsforoffice.microsoft.com`
because Outlook requires the add-in to use Microsoft's hosted copy. It is
subject to Microsoft's own terms.
