/*
 * FitPrint - shared DOM helper
 *
 * Used by both the task pane (preview / fallback print) and the print dialog,
 * so it lives in one place instead of being duplicated in each entry.
 */

/* global Promise */

/**
 * Resolve once every <img> inside `root` has finished loading (or failed), so a
 * printout is never blank because an image had not decoded yet.
 * @param {ParentNode} root
 * @returns {Promise<void[]>}
 */
export function waitForImages(root) {
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
