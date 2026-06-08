<p align="center">
  <img src="assets/icon-128.png" width="96" height="96" alt="FitPrint logo" />
</p>

<h1 align="center">FitPrint</h1>

<p align="center">
  <a href="https://github.com/TZERO78/FitPrint/actions/workflows/deploy.yml"><img src="https://github.com/TZERO78/FitPrint/actions/workflows/deploy.yml/badge.svg" alt="Deploy to GitHub Pages" /></a>
</p>

**FitPrint** is a free, open-source Microsoft Outlook add-in that prints the
currently selected email cleanly — and automatically **scales oversized images
down to the page width** so they are no longer cut off.

Modern Outlook renders messages with the Word engine and no longer offers a
"shrink to fit page" option. As a result, wide or high-resolution images get
clipped at the right edge when you print. FitPrint fixes that.

It is a **client-side Office.js web add-in** (no VSTO/COM, no backend), so it
runs in **New Outlook, classic Outlook, Outlook on the web, and Outlook for Mac**.
It is hosted on GitHub Pages: **<https://tzero78.github.io/FitPrint/>**

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

## Install (for users)

FitPrint is already hosted on GitHub Pages, so there is **nothing to build or
run** — you only sideload the hosted manifest once.

The manifest to install is:

```
https://tzero78.github.io/FitPrint/manifest.xml
```

**New Outlook / Outlook on the web**

1. Open any email.
2. Go to **Apps** (or **Get Add-ins**) → **My add-ins** →
   **Custom add-ins** → **Add a custom add-in** → **Add from URL…**
3. Paste the manifest URL above and confirm.

**Classic Outlook on Windows**

1. Open a received email in its own window (double-click it).
2. On the ribbon choose **All Apps / Get Add-ins**.
3. **My add-ins** → **Custom Addins** → **Add a custom add-in** →
   **Add from file…** (or **Add from URL…**), then point it at the manifest
   above (or a local copy of `manifest.xml`) and confirm the security prompt.

If your organization has disabled custom add-ins, sideload from a trusted
network share instead — see the Microsoft docs on
[sideloading Outlook add-ins](https://learn.microsoft.com/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing).

Once installed, open or select an email, click the **FitPrint** button in the
ribbon to open the task pane, then click **Print this email**.

## Develop locally

Prerequisites: [Node.js](https://nodejs.org/) (LTS) and npm.

```bash
npm install
npm run build      # production build (optional sanity check)
npm run dev-server # serves the add-in at https://localhost:3000
```

The first run installs a local HTTPS development certificate; accept the trust
prompt so Outlook can load `https://localhost:3000`.

To test your local build, sideload the **local** `manifest.xml` (which points at
`https://localhost:3000/`) using the same "Add from file" steps as above while
the dev server is running.

> Note: `npm start` tries to sideload automatically, but on many machines that
> path now requires a Microsoft 365 cloud sign-in and fails with `401`.
> Sideloading the manifest manually (as above) is the reliable way.

### Project structure

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
.github/workflows/
  deploy.yml                Build + deploy to GitHub Pages on every push to main
```

## Deployment

The add-in is fully static. Every push to `main` triggers the
`Deploy to GitHub Pages` GitHub Actions workflow, which runs `npm run build` and
publishes the `dist/` folder to GitHub Pages. The production build automatically
rewrites the URLs in the manifest from `https://localhost:3000/` to
`https://tzero78.github.io/FitPrint/`, so the hosted `manifest.xml` is ready to
distribute as-is.

## License

[MIT](LICENSE)
