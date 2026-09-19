# Third-party notices

The add-in bundles the following unmodified libraries in `src/taskpane/vendor/`
so that no third-party server ever sees attachment content. Each is used under
its own license, not the license in `LICENSE`.

| Library | Version | License | Full license text | Source |
| --- | --- | --- | --- | --- |
| JSZip | 3.10.1 | MIT (JSZip is dual-licensed MIT or GPLv3; used here under MIT) | `src/taskpane/vendor/LICENSE-jszip.txt` | https://github.com/Stuk/jszip/tree/v3.10.1 |
| PDF.js (Mozilla) | 3.11.174 | Apache-2.0 | `src/taskpane/vendor/LICENSE-pdfjs.txt` | https://github.com/mozilla/pdf.js/tree/v3.11.174 |

Files were downloaded from cdnjs (`cdnjs.cloudflare.com/ajax/libs/`) and are
byte-for-byte the published builds. SHA-256 checksums for verification:

```
acc7e41455a80765b5fd9c7ee1b8078a6d160bbbca455aeae854de65c947d59e  jszip.min.js
5b5799e6f8c680663207ac5b42ee14eed2a406fa7af48f50c154f0c0b1566946  pdf.min.js
feabdf309770ed24bba31a5467836cdc8cf639c705af27d52b585b041bb8527b  pdf.worker.min.js
```

JSZip's own distribution also embeds `pako` (MIT / Zlib); see its license file
for the complete notices.

## Loaded at runtime, not bundled

`office.js` (Microsoft) is loaded from `https://appsforoffice.microsoft.com`
because Outlook requires the add-in to use Microsoft's hosted copy. It is
subject to Microsoft's own terms.
