/*
 * FitPrint - print dialog
 *
 * This page runs in a separate, top-level window opened by the task pane.
 * It reads the prepared HTML document (stored in localStorage by the task pane,
 * same origin) and prints it. Printing from a top-level window is reliable
 * across New/Classic Outlook, Outlook on the web and Mac.
 */

/* global Office, document, window, DOMParser */

// Must match the key the task pane writes to (src/taskpane/taskpane.js).
const STORAGE_KEY = "fitprint:doc";

Office.onReady(() => {
  wireButtons();
  renderAndPrint();
});

/** Hook up the Print and Close buttons. */
function wireButtons() {
  document.getElementById("fp-print").onclick = () => window.print();
  document.getElementById("fp-close").onclick = () => {
    // Ask the task pane to close this dialog (it owns the dialog object).
    try {
      Office.context.ui.messageParent(JSON.stringify({ type: "close" }));
    } catch (e) {
      // If messaging is unavailable, try to close the window directly.
      window.close();
    }
  };
}

/** Resolve once every <img> inside `root` has loaded (or failed). */
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

/** Read the prepared document, render it into this page, then print. */
function renderAndPrint() {
  const mount = document.getElementById("fp-mount");

  let docHtml = "";
  try {
    docHtml = window.localStorage.getItem(STORAGE_KEY) || "";
  } catch (e) {
    docHtml = "";
  }

  if (!docHtml) {
    mount.innerHTML =
      "<p>Could not load the prepared email. Please close this window and click " +
      "<b>Print this email</b> again.</p>";
    return;
  }

  // The stored value is a full HTML document. Copy its <style> blocks into our
  // <head> and its body markup into the mount point.
  const parsed = new DOMParser().parseFromString(docHtml, "text/html");
  parsed.querySelectorAll("style").forEach((s) => document.head.appendChild(s.cloneNode(true)));
  mount.innerHTML = parsed.body ? parsed.body.innerHTML : docHtml;

  // Wait for images so the printout is not blank, then open the print dialog.
  waitForImages(mount).then(() => {
    window.focus();
    window.print();
  });
}
