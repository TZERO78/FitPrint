/*
 * FitPrint - image handling
 *
 * Two jobs live here:
 *   1) embedInlineImages: replace "cid:" references in the body with the actual
 *      image bytes (as base64 data: URIs), read from the message's attachments.
 *   2) resizeAndOrientImages: shrink oversized images with a <canvas> and bake in
 *      the correct EXIF rotation so photos don't print sideways or huge.
 *
 * Both are robust: a single broken image is caught and skipped, never aborting
 * the whole print (the original src is simply left untouched).
 */

/* global Office, document, window, Image, DOMParser, URL, createImageBitmap, fetch */

import { getReadModeAttachments, getAttachmentContentAsync } from "./office-helpers.js";

/**
 * Replace every <img src="cid:..."> in the HTML with a base64 data: URI built
 * from the matching inline attachment.
 *
 * Office.js does not expose an attachment's Content-ID directly, so we match the
 * "cid" token against the attachment file name (which is how Outlook usually
 * names inline images), with an order-based fallback as a last resort.
 *
 * @param {string} html the message body as HTML
 * @returns {Promise<string>} HTML with inline images embedded
 */
export async function embedInlineImages(html) {
  // In READ mode the attachment list is a synchronous property (item.attachments).
  // getAttachmentsAsync is a COMPOSE-mode API and throws here.
  let attachments = [];
  try {
    attachments = getReadModeAttachments();
  } catch (e) {
    // No attachments available - just return the body unchanged.
    return html;
  }

  // Prefer attachments flagged as inline; if none are flagged, fall back to any
  // image attachment (some servers don't set the inline flag correctly).
  const isImage = (a) => /^image\//i.test(a.contentType || "");
  const inline = attachments.filter((a) => a.isInline && isImage(a));
  const candidates = inline.length ? inline : attachments.filter(isImage);
  // eslint-disable-next-line no-console
  console.log(
    "[FitPrint] attachments:",
    attachments.length,
    "| inline images:",
    inline.length,
    "| candidates:",
    candidates.length
  );
  if (!candidates.length) {
    return html;
  }

  // Download each candidate once and remember it by (lower-cased) file name.
  const byName = {};
  const inOrder = [];
  for (const att of candidates) {
    try {
      const content = await getAttachmentContentAsync(att.id);
      if (content && content.format === Office.MailboxEnums.AttachmentContentFormat.Base64) {
        const uri = `data:${att.contentType};base64,${content.content}`;
        const key = (att.name || "").toLowerCase();
        byName[key] = uri;
        inOrder.push({ key, uri });
      }
    } catch (e) {
      // Skip this one image, keep going with the rest.
    }
  }

  // Replace src="cid:..." and src='cid:...' (single or double quotes).
  return html.replace(/src\s*=\s*(["'])cid:([^"']+)\1/gi, (match, quote, cidRaw) => {
    const cid = decodeURIComponent(cidRaw).replace(/[<>]/g, "").toLowerCase();
    // The cid is often "name@host" - the part before "@" tends to be the file name.
    const cidName = cid.split("@")[0];

    // 1) exact file-name match, 2) prefix match, 3) consume next image in order.
    let uri = byName[cidName] || byName[cid];
    if (!uri) {
      const hit = inOrder.find(
        (r) => r.key === cidName || cidName.indexOf(r.key) === 0 || r.key.indexOf(cidName) === 0
      );
      uri = hit && hit.uri;
    }
    if (!uri && inOrder.length) {
      uri = inOrder.shift().uri;
    }

    return uri ? `src=${quote}${uri}${quote}` : match;
  });
}

/**
 * Load image bytes for a given <img> src.
 * Handles data: URIs and http(s) URLs; returns null for anything we can't read
 * (e.g. a remote image blocked by CORS), so the caller keeps the original.
 * @param {string} src
 * @returns {Promise<Blob|null>}
 */
async function srcToBlob(src) {
  try {
    if (src.startsWith("data:")) {
      const res = await fetch(src);
      return await res.blob();
    }
    if (/^https?:/i.test(src)) {
      const res = await fetch(src, { mode: "cors" });
      if (!res.ok) {
        return null;
      }
      return await res.blob();
    }
  } catch (e) {
    // CORS / network error - leave the original image as-is.
  }
  return null;
}

/**
 * Decode a Blob into something we can draw on a canvas.
 * Prefers createImageBitmap with imageOrientation "from-image" so that EXIF
 * rotation is applied automatically; falls back to a plain <img> element.
 * @param {Blob} blob
 * @returns {Promise<{source: CanvasImageSource, width: number, height: number, close: Function}|null>}
 */
async function decodeImage(blob) {
  try {
    if (window.createImageBitmap) {
      const bmp = await createImageBitmap(blob, { imageOrientation: "from-image" });
      return { source: bmp, width: bmp.width, height: bmp.height, close: () => bmp.close && bmp.close() };
    }
  } catch (e) {
    // Fall through to the <img> fallback below.
  }

  const url = URL.createObjectURL(blob);
  try {
    const img = await new Promise((resolve, reject) => {
      const el = new Image();
      el.onload = () => resolve(el);
      el.onerror = reject;
      el.src = url;
    });
    return { source: img, width: img.naturalWidth, height: img.naturalHeight, close: () => URL.revokeObjectURL(url) };
  } catch (e) {
    URL.revokeObjectURL(url);
    return null;
  }
}

/**
 * Shrink oversized images and bake in EXIF orientation.
 * Images whose long edge is already <= maxEdge are left untouched (modern
 * browsers honor EXIF orientation on render via the image-orientation CSS rule
 * we inject in the print stylesheet).
 *
 * @param {string} html body HTML (inline images should already be embedded)
 * @param {number} [maxEdge=1600] target maximum length of the longer edge, in px
 * @param {number} [quality=0.85] JPEG quality for re-encoded photos
 * @returns {Promise<string>} HTML with oversized images replaced by smaller ones
 */
export async function resizeAndOrientImages(html, maxEdge = 1600, quality = 0.85) {
  const doc = new DOMParser().parseFromString(html, "text/html");
  const images = Array.from(doc.images);

  for (const img of images) {
    try {
      const src = img.getAttribute("src");
      if (!src) {
        continue;
      }

      const blob = await srcToBlob(src);
      if (!blob || !/^image\//.test(blob.type)) {
        continue;
      }

      const decoded = await decodeImage(blob);
      if (!decoded) {
        continue;
      }

      const longEdge = Math.max(decoded.width, decoded.height);
      if (longEdge <= maxEdge) {
        // Small enough already - don't re-encode (avoids quality loss).
        decoded.close();
        continue;
      }

      // Draw the (correctly oriented) image onto a smaller canvas.
      const scale = maxEdge / longEdge;
      const w = Math.round(decoded.width * scale);
      const h = Math.round(decoded.height * scale);
      const canvas = document.createElement("canvas");
      canvas.width = w;
      canvas.height = h;
      canvas.getContext("2d").drawImage(decoded.source, 0, 0, w, h);
      decoded.close();

      // Keep PNGs as PNG (preserves transparency); everything else becomes JPEG.
      const isPng = /png/i.test(blob.type);
      const dataUrl = isPng ? canvas.toDataURL("image/png") : canvas.toDataURL("image/jpeg", quality);

      img.setAttribute("src", dataUrl);
      // Drop fixed width/height attributes so the print CSS can scale to the page.
      img.removeAttribute("width");
      img.removeAttribute("height");
    } catch (e) {
      // Any failure here just keeps the original image - printing never breaks.
    }
  }

  // Preserve <style> blocks from the original body, then return the body markup.
  const styles = Array.from(doc.querySelectorAll("style")).map((s) => s.outerHTML).join("\n");
  const body = doc.body ? doc.body.innerHTML : html;
  return styles + body;
}
