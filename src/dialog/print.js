/*
 * OpenMailPrint - print dialog
 *
 * This page runs in a separate, top-level window opened by the task pane.
 * The task pane streams the prepared HTML document to us over the Office dialog
 * messaging channel (we cannot share localStorage with the task pane, because in
 * Outlook on the web the task pane runs in a partitioned iframe). Once we have
 * the full document we render it and print. Printing from a top-level window is
 * reliable across New/Classic Outlook, Outlook on the web and Mac.
 */

/* global Office, document, window, DOMParser */

import { waitForImages } from "../shared/wait-for-images.js";

// Buffer for the incoming document chunks.
let chunks = [];

Office.onReady(() => {
  wireButtons();

  // Register the receiver for messages from the task pane, then tell it we are
  // ready. The order matters: only signal "ready" once we can receive data.
  Office.context.ui.addHandlerAsync(
    Office.EventType.DialogParentMessageReceived,
    onParentMessage,
    () => {
      Office.context.ui.messageParent(JSON.stringify({ type: "ready" }));
    }
  );
});

/** Hook up the Print and Close buttons. */
function wireButtons() {
  document.getElementById("fp-print").onclick = () => window.print();
  document.getElementById("fp-close").onclick = () => {
    // Ask the task pane to close this dialog (it owns the dialog object).
    try {
      Office.context.ui.messageParent(JSON.stringify({ type: "close" }));
    } catch (e) {
      window.close();
    }
  };
}

/** Handle a single message from the task pane (the document arrives in chunks). */
function onParentMessage(arg) {
  let msg;
  try {
    msg = JSON.parse(arg.message);
  } catch (e) {
    return;
  }

  if (msg.type === "begin") {
    chunks = new Array(msg.total);
  } else if (msg.type === "chunk") {
    chunks[msg.index] = msg.data;
    // Acknowledge this chunk so the task pane sends the next one.
    Office.context.ui.messageParent(JSON.stringify({ type: "ack", index: msg.index }));
  } else if (msg.type === "end") {
    renderAndPrint(chunks.join(""));
  }
}

/** Render the assembled document into this page, then print. */
function renderAndPrint(docHtml) {
  const mount = document.getElementById("fp-mount");

  if (!docHtml) {
    mount.innerHTML =
      "<p>Could not load the prepared email. Please close this window and click " +
      "<b>Print this email</b> again.</p>";
    return;
  }

  // The document is a full HTML page. Copy its <style> blocks into our <head>
  // and its body markup into the mount point.
  const parsed = new DOMParser().parseFromString(docHtml, "text/html");
  parsed.querySelectorAll("style").forEach((s) => document.head.appendChild(s.cloneNode(true)));
  mount.innerHTML = parsed.body ? parsed.body.innerHTML : docHtml;

  // Wait for images so the printout is not blank, then open the print dialog.
  waitForImages(mount).then(() => {
    window.focus();
    window.print();
  });
}
