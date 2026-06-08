/*
 * FitPrint - Office.js helpers
 *
 * Office.js still uses the old callback style: you pass a function that receives
 * an "asyncResult". These tiny wrappers turn those calls into Promises so the
 * rest of the code can use clean async/await with try/catch.
 *
 * Everything here only READS the message. FitPrint never modifies the email.
 */

/* global Office, window */

/**
 * Read the current message body.
 * @param {Office.CoercionType} coercionType e.g. Office.CoercionType.Html
 * @returns {Promise<string>}
 */
export function getBodyAsync(coercionType) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.getAsync(coercionType, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(result.error);
      }
    });
  });
}

/**
 * List the attachments of the current message.
 *
 * In READ mode (which is what FitPrint runs in) the attachment list is exposed
 * as the synchronous property item.attachments. Note: getAttachmentsAsync() is a
 * COMPOSE-mode API and is NOT available here - calling it throws.
 *
 * @returns {Array<Office.AttachmentDetails>}
 */
export function getReadModeAttachments() {
  return Office.context.mailbox.item.attachments || [];
}

/**
 * Download the content of a single attachment.
 * For inline images Outlook returns the bytes as a Base64 string.
 * @param {string} attachmentId from AttachmentDetails.id
 * @returns {Promise<Office.AttachmentContent>} { content, format }
 */
export function getAttachmentContentAsync(attachmentId) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.getAttachmentContentAsync(attachmentId, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(result.error);
      }
    });
  });
}

// Size of each message chunk sent to the dialog. Dialog messages are limited in
// size, and in Outlook on the web the task pane runs in a partitioned iframe so
// we cannot share localStorage with the dialog - we stream the document instead.
const CHUNK_SIZE = 16000;

/**
 * Open a separate, top-level window (an Office dialog), stream the prepared HTML
 * to it over the Office dialog messaging channel, and let it print. A top-level
 * window prints reliably across New/Classic Outlook, Outlook on the web and Mac.
 *
 * Flow:
 *   1. We open print.html.
 *   2. The dialog registers its receiver and messages us "ready".
 *   3. We send the document as "begin" + many "chunk" + "end" messages.
 *   4. The dialog reassembles the HTML, renders it and calls window.print().
 *
 * @param {string} fullDocHtml the complete printable HTML document
 * @returns {Promise<void>} resolves when the dialog is closed
 */
export function printViaDialog(fullDocHtml) {
  return new Promise((resolve, reject) => {
    // Build the dialog URL relative to THIS page so it works both at the dev
    // server root (https://localhost:3000/print.html) and under a GitHub Pages
    // project subpath (https://<user>.github.io/FitPrint/print.html).
    const url = new URL("print.html", window.location.href).href;

    Office.context.ui.displayDialogAsync(
      url,
      { height: 70, width: 60, displayInIframe: false },
      (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          reject(result.error);
          return;
        }
        const dialog = result.value;
        const total = Math.ceil(fullDocHtml.length / CHUNK_SIZE) || 1;
        let started = false;

        // Send one chunk; the dialog acknowledges each one before we send the
        // next. This acknowledged flow guarantees every chunk arrives in order,
        // so large base64 images are never truncated in transit.
        const sendChunk = (index) => {
          try {
            const part = fullDocHtml.substr(index * CHUNK_SIZE, CHUNK_SIZE);
            dialog.messageChild(JSON.stringify({ type: "chunk", index, data: part }));
          } catch (e) {
            // Messaging not available - give up so the caller can fall back to
            // printing from the task pane.
            dialog.close();
            reject(e);
          }
        };

        // Messages coming back from the dialog.
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
          let msg;
          try {
            msg = JSON.parse(arg.message);
          } catch (e) {
            return;
          }
          if (msg.type === "ready") {
            if (started) {
              return;
            }
            started = true;
            // eslint-disable-next-line no-console
            console.log("[FitPrint] sending document:", fullDocHtml.length, "chars in", total, "chunks");
            try {
              dialog.messageChild(JSON.stringify({ type: "begin", total }));
            } catch (e) {
              dialog.close();
              reject(e);
              return;
            }
            sendChunk(0);
          } else if (msg.type === "ack") {
            const next = msg.index + 1;
            if (next < total) {
              sendChunk(next);
            } else {
              try {
                dialog.messageChild(JSON.stringify({ type: "end" }));
              } catch (e) {
                /* ignore - the dialog already has every chunk */
              }
            }
          } else if (msg.type === "close") {
            dialog.close();
            resolve();
          }
        });

        // Fired when the dialog is closed (e.g. the user closes the window).
        dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
          resolve();
        });
      }
    );
  });
}
