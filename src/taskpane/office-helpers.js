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
 * Requires Mailbox requirement set 1.8 (declared in the manifest).
 * @returns {Promise<Array<Office.AttachmentDetails>>}
 */
export function getAttachmentsAsync() {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.getAttachmentsAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value || []);
      } else {
        reject(result.error);
      }
    });
  });
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

/**
 * Open a separate, top-level window (an Office dialog) that will show the
 * print preview and trigger the browser's print dialog. A top-level window
 * prints reliably across New/Classic Outlook, Outlook on the web and Mac.
 *
 * The prepared HTML is handed over via localStorage (same origin), so there is
 * no message-size limit to worry about.
 *
 * @returns {Promise<Office.Dialog>}
 */
export function openPrintDialog() {
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
        // The dialog can ask us to close it (its "Close" button).
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
          try {
            const msg = JSON.parse(arg.message);
            if (msg && msg.type === "close") {
              dialog.close();
            }
          } catch (e) {
            dialog.close();
          }
        });
        resolve(dialog);
      }
    );
  });
}
