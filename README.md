# FitPrint

**FitPrint** is a free, open-source Microsoft Outlook add-in that prints the
currently selected email cleanly — and automatically **scales oversized images
down to the page width** so they are no longer cut off.

Modern Outlook renders messages with the Word engine and no longer offers a
"shrink to fit page" option. As a result, wide or high-resolution images get
clipped at the right edge when you print. FitPrint fixes that.

It is a **client-side Office.js web add-in** (no VSTO/COM, no backend), so it
runs in **New Outlook, classic Outlook, Outlook on the web, and Outlook for Mac**.

## What it does

When you click **Print this email**, FitPrint:

1. Reads the message body as HTML (`item.body.getAsync`).
2. Embeds inline images: it loads the message attachments
   (`getAttachmentsAsync` + `getAttachmentContentAsync`) and replaces the
   `cid:` references in the body with base64 `data:` URIs.
3. Shrinks large images with a `<canvas>` (long edge max ~1600 px, JPEG quality
   ~0.85) and bakes in the correct **EXIF rotation** so photos are not sideways.
4. Adds a **From / To / Date / Subject** header block.
5. Injects print CSS (notably `img { max-width: 100%; height: auto }`).
6. Opens a top-level **print dialog** and calls `window.print()` so the normal
   system print dialog with preview appears.

The email is only ever **read, never modified** (the manifest requests
`ReadItem` permission only). A single broken image never aborts the print — each
image is processed in its own `try/catch` and falls back to the original.

## Project structure

```
manifest.xml              Add-in manifest (classic XML, ReadItem, Mailbox 1.8)
src/
  taskpane/
    taskpane.html/.css/.js  UI + pipeline orchestration
    office-helpers.js       Promise wrappers around Office.js callbacks
    images.js               Embed cid: images, resize + EXIF orientation
    build-document.js       Header block, print CSS, document assembly
  dialog/
    print.html / print.js   Top-level print window (renders + prints)
  commands/                 Generated command file (unused placeholder)
```

## Prerequisites

- [Node.js](https://nodejs.org/) (LTS) and npm
- Outlook (New, classic, web, or Mac)

## Install & run (development)

```bash
npm install
npm run build      # production build (optional sanity check)
npm run dev-server # serves the add-in at https://localhost:3000
```

The first run installs a local HTTPS development certificate; accept the trust
prompt so the browser/Outlook can load `https://localhost:3000`.

### Sideload the add-in

`npm start` tries to sideload automatically, but on many machines that path now
requires a Microsoft 365 cloud sign-in and fails with `401`. The reliable way is
to **sideload the manifest manually** while the dev server is running:

**New Outlook / Outlook on the web**

1. Run `npm run dev-server` and keep it running.
2. In Outlook, open any email, then go to **Apps** (or **Get Add-ins**) →
   **My add-ins** → **Custom add-ins** → **Add a custom add-in** →
   **Add from file…**
3. Select `manifest.xml` from this folder and confirm.

**Classic Outlook on Windows** — sideload from a trusted catalog (a shared
folder): see the Microsoft docs on
[sideloading Outlook add-ins](https://learn.microsoft.com/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing).

Once installed, open or select an email, click the **FitPrint** button in the
ribbon to open the task pane, then click **Print this email**.

## Deployment

The add-in is fully static and is intended to be hosted on **GitHub Pages**.
The production build rewrites the URLs in the manifest from
`https://localhost:3000/` to the GitHub Pages location
(`https://tzero78.github.io/FitPrint/`). Publish the contents of the `dist/`
folder and distribute the production `manifest.xml`.

## License

[MIT](LICENSE)
