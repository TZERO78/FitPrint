/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 *
 * FitPrint - task pane orchestration
 *
 * When the user clicks "Print this email", we run a small pipeline:
 *   1. Read the message body as HTML.
 *   2. Embed inline images (replace cid: references with base64 data).
 *   3. Shrink oversized images and fix their EXIF rotation.
 *   4. Add a From/To/Date/Subject header and print CSS.
 *   5. Open a print dialog (a top-level window) and call window.print() there.
 *
 * The message is only ever READ, never modified.
 */

/* global Office, document, window, DOMParser */

import { getBodyAsync, openPrintDialog } from "./office-helpers.js";
import { embedInlineImages, resizeAndOrientImages } from "./images.js";
import { buildHeaderHtml, buildPrintableDocument } from "./build-document.js";

// Same key is read by the print dialog (src/dialog/print.js). Both pages are
// served from the same origin, so they share localStorage.
const STORAGE_KEY = "fitprint:doc";

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

/** Show a short status message in the task pane. */
function setStatus(text) {
  document.getElementById("status").textContent = text;
}

/** Collect the header fields. In read mode these item properties are synchronous. */
function readHeaderData() {
  const item = Office.context.mailbox.item;
  return {
    from: item.from, // EmailAddressDetails { displayName, emailAddress }
    to: item.to, // Array<EmailAddressDetails>
    subject: item.subject, // string
    date: item.dateTimeCreated, // Date
  };
}

/** Resolve once every <img> inside `root` has finished loading (or failed). */
function waitForImages(root) {
  const images = Array.from(root.querySelectorAll("img"));
  return Promise.all(
    images.map((img) =>
      img.complete
        ? Promise.resolve()
        : new Promise((resolve) => {
            img.onload = resolve;
            img.onerror = resolve;
          })
    )
  );
}

/**
 * Fallback used only if the print dialog cannot be opened: render the prepared
 * document into the task pane itself and print from here. The print CSS in
 * taskpane.css hides the FitPrint controls while printing.
 */
async function printInTaskPane(fullDocHtml) {
  const parsed = new DOMParser().parseFromString(fullDocHtml, "text/html");
  parsed.querySelectorAll("style").forEach((s) => document.head.appendChild(s.cloneNode(true)));

  const preview = document.getElementById("preview");
  preview.innerHTML = parsed.body ? parsed.body.innerHTML : fullDocHtml;

  document.body.classList.add("fp-printing");
  await waitForImages(preview);
  window.print();
  document.body.classList.remove("fp-printing");
}

/**
 * Button handler: run the full prepare-and-print pipeline.
 */
export async function run() {
  const preview = document.getElementById("preview");
  preview.innerHTML = "";

  try {
    setStatus("Reading email body...");
    const rawHtml = await getBodyAsync(Office.CoercionType.Html);

    setStatus("Embedding inline images...");
    const withImages = await embedInlineImages(rawHtml);

    setStatus("Resizing large images...");
    const resized = await resizeAndOrientImages(withImages);

    setStatus("Building printable document...");
    const headerHtml = buildHeaderHtml(readHeaderData());
    const fullDoc = buildPrintableDocument(headerHtml, resized);

    // Show a preview in the task pane so the user sees what will be printed.
    const parsed = new DOMParser().parseFromString(fullDoc, "text/html");
    preview.innerHTML = parsed.body ? parsed.body.innerHTML : fullDoc;

    // Hand the document to the print dialog via localStorage (no size limit).
    try {
      window.localStorage.setItem(STORAGE_KEY, fullDoc);
    } catch (e) {
      // Storage might be full for very large emails - fall back to task-pane print.
      setStatus("Preparing print (fallback)...");
      await printInTaskPane(fullDoc);
      setStatus("Print dialog opened.");
      return;
    }

    setStatus("Opening print dialog...");
    try {
      await openPrintDialog();
      setStatus("Print dialog opened. If nothing appeared, allow pop-ups and try again.");
    } catch (e) {
      // displayDialogAsync failed (blocked/unsupported) - print from the task pane.
      await printInTaskPane(fullDoc);
      setStatus("Print dialog opened.");
    }
  } catch (error) {
    const message = error && error.message ? error.message : String(error);
    setStatus("Could not prepare the email for printing: " + message);
  }
}
