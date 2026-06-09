/*
 * Generate the OpenMailPrint PNG icons from assets/logo.svg.
 *
 * Run with:  node tools/generate-icons.mjs
 * Requires "sharp" (install once with:  npm install sharp --no-save).
 *
 * The Office manifest references icon-16/32/64/80/128.png; logo-filled.png is
 * used as the header logo in the task pane.
 */

import sharp from "sharp";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";

const assetsDir = join(dirname(fileURLToPath(import.meta.url)), "..", "assets");
const svg = readFileSync(join(assetsDir, "logo.svg"));

// filename -> pixel size (square)
const targets = {
  "icon-16.png": 16,
  "icon-32.png": 32,
  "icon-64.png": 64,
  "icon-80.png": 80,
  "icon-128.png": 128,
  "logo-filled.png": 128,
};

for (const [name, size] of Object.entries(targets)) {
  await sharp(svg)
    .resize(size, size)
    .png()
    .toFile(join(assetsDir, name));
  console.log(`wrote ${name} (${size}x${size})`);
}
