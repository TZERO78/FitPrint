<p align="center">
  <img src="assets/icon-128.png" width="96" height="96" alt="OpenMailPrint logo" />
</p>

<h1 align="center">OpenMailPrint</h1>

<p align="center">
  <a href="https://github.com/TZERO78/OpenMailPrint/actions/workflows/deploy.yml"><img src="https://github.com/TZERO78/OpenMailPrint/actions/workflows/deploy.yml/badge.svg" alt="Deploy to GitHub Pages" /></a>
</p>

**OpenMailPrint** is a free, open-source Microsoft Outlook add-in that prints the
currently selected email cleanly — and automatically **scales oversized images
down to the page width** so they are no longer cut off.

Modern Outlook renders messages with the Word engine and no longer offers a
"shrink to fit page" option. As a result, wide or high-resolution images get
clipped at the right edge when you print. OpenMailPrint fixes that.

It is a **client-side Office.js web add-in** (no VSTO/COM, no backend), so it
runs in **New Outlook, classic Outlook, Outlook on the web, and Outlook for Mac**.
It is hosted on GitHub Pages: **<https://tzero78.github.io/OpenMailPrint/>**

## What it does

When you click **Print this email**, OpenMailPrint:

1. Reads the message body as HTML (`item.body.getAsync`).
2. Embeds inline images: it reads the message attachments (the synchronous
   `item.attachments` property in read mode) and `getAttachmentContentAsync`,
   then replaces the `cid:` references in the body with base64 `data:` URIs.
3. Shrinks large images with a `<canvas>` (long edge max ~1600 px, JPEG quality
   ~0.85) and bakes in the correct **EXIF rotation** so photos are not sideways.
   Remote images that cannot be inlined are replaced with a placeholder so no
   tracking pixels load while printing.
4. Sanitizes the body with **DOMPurify** before it is rendered — stripping
   scripts, inline event handlers, `javascript:` links and structural injection
   (`base`, `meta`, `form`, framing tags).
5. Adds a header block with **From / To / Cc / Date / Subject** (empty fields,
   e.g. a missing Cc, are omitted).
6. Injects print CSS (notably `img { max-width: 100%; height: auto }`).
7. Opens a top-level **print dialog** and calls `window.print()` so the normal
   system print dialog with preview appears.

The email is only ever **read, never modified** (the manifest requests
`ReadItem` permission only). A single broken image never aborts the print — each
image is processed in its own `try/catch` and falls back to the original.

## Install (for users)

OpenMailPrint is already hosted on GitHub Pages, so there is **nothing to build or
run** — you only sideload the hosted manifest once.

The manifest to install is:

```
https://tzero78.github.io/OpenMailPrint/manifest.xml
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

Once installed, open or select an email, click the **OpenMailPrint** button in the
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
    sanitize.js             DOMPurify sanitizing of the mail body
  dialog/
    print.html / print.js   Top-level print window (renders + prints)
  shared/
    wait-for-images.js      Helper shared by the task pane and the dialog
.github/workflows/
  deploy.yml                Build + deploy to GitHub Pages on every push to main
```

## Deployment

The add-in is fully static. Every push to `main` triggers the
`Deploy to GitHub Pages` GitHub Actions workflow, which audits the shipped
dependencies (`npm audit --omit=dev --audit-level=high`, failing the deploy on a
high-severity advisory), runs `npm run build`, and publishes the `dist/` folder
to GitHub Pages. The production build automatically
rewrites the URLs in the manifest from `https://localhost:3000/` to
`https://tzero78.github.io/OpenMailPrint/`, so the hosted `manifest.xml` is ready to
distribute as-is.

## Changelog

### v1.2.2
- **Privacy:** external (`http(s)`) images are no longer fetched while preparing the
  print — only already-inlined `data:` images are processed, so no remote server is
  contacted (no IP/timestamp leak); remote images render as a placeholder.
- Hardening: re-sanitize at every `innerHTML` sink, bounds-check incoming print-dialog
  chunks, make the address formatter self-escaping, and guard against an empty body.

### v1.2.1
- Renamed the project from **FitPrint** to **OpenMailPrint** (clearer, and avoids
  collision with unrelated "FitPrint" products). No functional change; the add-in
  ID is unchanged so existing installs keep working.

### v1.2.0
- **Security hardening:** the email body is now sanitized with
  [DOMPurify](https://github.com/cure53/DOMPurify) before it is rendered; a
  Content-Security-Policy and a `no-referrer` policy are enforced on both pages;
  remote images that cannot be inlined are replaced with a placeholder so no
  tracking pixels load while printing; and the deploy fails on high-severity
  advisories in shipped dependencies (`npm audit --omit=dev`).
- Removed the unused generated command file, dropped IE 11 from the build
  targets, and de-duplicated the image-loading helper.

### v1.1.0
- Header block now includes **Cc**; empty header fields are omitted.
- Inline images are embedded correctly (read `item.attachments`).
- Reliable, acknowledged document transfer to the print dialog.
- Content-hashed bundles so updates are never served from a stale cache.

### v1.0.0
- Initial release: read the email, embed inline `cid:` images, downscale large
  images with EXIF orientation, add a header block and print via a dialog.

## Acknowledgements

OpenMailPrint sanitizes untrusted email HTML with
[**DOMPurify**](https://github.com/cure53/DOMPurify) by [Cure53](https://cure53.de/) —
a huge thank you for this excellent, battle-tested library.

## License

[MIT](LICENSE)
