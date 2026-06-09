/*
 * OpenMailPrint - HTML sanitizing
 *
 * The email body is attacker-controllable HTML (anyone can send you mail), and
 * we render it via innerHTML both in the task-pane preview and in the print
 * dialog. We run it through DOMPurify first to strip script, inline event
 * handlers, javascript:/data:text-html URLs, and structural injection vectors
 * (<base>, <meta http-equiv="refresh">, <form action>, framing tags).
 *
 * Inline <style> blocks and embedded data: image URIs are intentionally kept -
 * they are needed to render the email faithfully and are safe in this context
 * (DOMPurify allows data: URIs only on media tags like <img>).
 */

import DOMPurify from "dompurify";

export function sanitizeMailHtml(html) {
  return DOMPurify.sanitize(html, {
    FORCE_BODY: true, // treat input as body content even if it starts with e.g. <style>
    ALLOW_UNKNOWN_PROTOCOLS: false, // only known-safe URL schemes survive
    ADD_TAGS: ["style"], // keep email styling
    // Drop tags that can redirect, reframe, exfiltrate or auto-refresh the page.
    FORBID_TAGS: ["base", "meta", "form", "iframe", "object", "embed"],
    FORBID_ATTR: ["ping"], // <a ping> beacons
  });
}
