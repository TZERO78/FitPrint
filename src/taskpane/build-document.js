/*
 * OpenMailPrint - build the printable document
 *
 * Takes the email's header fields plus the prepared body HTML and assembles a
 * complete, self-contained HTML document with a header block and print CSS.
 */

/**
 * Print stylesheet injected into the document we hand to the print dialog.
 * The key rule for OpenMailPrint is `img { max-width: 100% }` so oversized images
 * are never clipped at the page edge.
 */
export const PRINT_CSS = `
  @page { margin: 16mm; }
  * { box-sizing: border-box; }
  body { font-family: Calibri, Arial, sans-serif; color: #000; margin: 0; line-height: 1.4; }
  .fp-doc { width: 100%; }
  .fp-headers { border-bottom: 1px solid #999; margin-bottom: 14px; padding-bottom: 8px; }
  .fp-headers .fp-row { margin: 2px 0; }
  .fp-headers .fp-label { font-weight: bold; display: inline-block; min-width: 64px; }
  .fp-body { word-wrap: break-word; }

  /* The whole point of OpenMailPrint: keep images within the printable page width. */
  img { max-width: 100% !important; height: auto !important; image-orientation: from-image; }
  table { max-width: 100% !important; }

  @media print {
    img { max-width: 100% !important; height: auto !important; }
    a { color: #000; text-decoration: none; }
  }
`;

/** Escape text so it is safe to place inside HTML. */
function escapeHtml(value) {
  return String(value == null ? "" : value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/** Format a single EmailAddressDetails as "Display Name <email>" (or just the email). */
function formatAddress(addr) {
  if (!addr) {
    return "";
  }
  return addr.displayName ? `${addr.displayName} <${addr.emailAddress}>` : addr.emailAddress || "";
}

/** Join a list of recipients into one comma-separated string. */
function formatRecipients(list) {
  if (!list || !list.length) {
    return "";
  }
  return list.map(formatAddress).join(", ");
}

/**
 * Build the header block shown above the email (From / To / Cc / Date / Subject).
 * Empty fields are omitted entirely - e.g. a message without Cc shows no Cc row.
 *
 * @param {{from: object, to: Array, cc: Array, subject: string, date: Date}} header
 * @returns {string} HTML
 */
export function buildHeaderHtml(header) {
  // [label, value] pairs in the order they should appear.
  const fields = [
    ["From", formatAddress(header.from)],
    ["To", formatRecipients(header.to)],
    ["Cc", formatRecipients(header.cc)],
    ["Date", header.date ? header.date.toLocaleString() : ""],
    ["Subject", header.subject],
  ];

  const rows = fields
    // Keep only fields that actually have a value.
    .filter(([, value]) => value != null && String(value).trim() !== "")
    .map(
      ([label, value]) =>
        `<div class="fp-row"><span class="fp-label">${label}:</span> ${escapeHtml(value)}</div>`
    )
    .join("\n    ");

  return `
  <div class="fp-headers">
    ${rows}
  </div>`;
}

/**
 * Assemble the full, self-contained HTML document for printing.
 * Note: bodyHtml is the email's own HTML and is intentionally NOT escaped - we
 * want it rendered. Only the header fields above are escaped.
 *
 * @param {string} headerHtml from buildHeaderHtml()
 * @param {string} bodyHtml prepared email body (images embedded + resized)
 * @returns {string} complete HTML document
 */
export function buildPrintableDocument(headerHtml, bodyHtml) {
  return `<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8" />
  <title>OpenMailPrint</title>
  <style>${PRINT_CSS}</style>
</head>
<body>
  <div class="fp-doc">
    ${headerHtml}
    <div class="fp-body">${bodyHtml}</div>
  </div>
</body>
</html>`;
}
