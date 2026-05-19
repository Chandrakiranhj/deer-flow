#!/usr/bin/env node
/**
 * dom_export.cjs — Render a composed VEPIP deck HTML to a fully-editable PPTX.
 *
 * Pipeline:
 *   1. puppeteer-core launches the system Chrome (we already located it from
 *      Python; passed in via --chrome).
 *   2. Loads the composed HTML deck file (one document with N stacked
 *      1920×1080 <section id="slide-N"> elements).
 *   3. Waits for fonts to load (Inter, Playfair Display, etc.) so dom-to-pptx
 *      reads correct text metrics from getComputedStyle.
 *   4. Injects the dom-to-pptx UMD bundle (exposes window.domToPptx.exportToPptx).
 *   5. Calls exportToPptx(['#slide-1','#slide-2',…], {skipDownload:true,
 *      layout:'LAYOUT_WIDE', autoEmbedFonts:true}) — returns a Blob.
 *   6. Reads Blob bytes as base64, sends back to Node, writes the .pptx file.
 *
 * Usage:
 *   node dom_export.cjs --chrome <path> --html <deck.html> --out <deck.pptx>
 */
const fs = require("fs");
const path = require("path");
const puppeteer = require("puppeteer-core");

function parseArgs(argv) {
  const out = {};
  for (let i = 2; i < argv.length; i++) {
    const a = argv[i];
    if (a.startsWith("--") && i + 1 < argv.length) {
      out[a.slice(2)] = argv[i + 1];
      i++;
    }
  }
  return out;
}

async function main() {
  const args = parseArgs(process.argv);
  if (!args.chrome || !args.html || !args.out) {
    console.error("Usage: dom_export.cjs --chrome <path> --html <file> --out <file>");
    process.exit(2);
  }
  if (!fs.existsSync(args.chrome)) {
    console.error(`Chrome not found at ${args.chrome}`);
    process.exit(3);
  }
  if (!fs.existsSync(args.html)) {
    console.error(`HTML file not found at ${args.html}`);
    process.exit(4);
  }

  // dom-to-pptx UMD bundle path. The package's "exports" field restricts
  // require.resolve to its declared subpaths, so we locate the package root
  // via require.resolve("dom-to-pptx") then build the bundle path manually.
  const pkgEntry = require.resolve("dom-to-pptx");
  // pkgEntry is .../node_modules/dom-to-pptx/dist/dom-to-pptx.cjs; walk up
  // to the package root and append the bundle file.
  const pkgRoot = path.dirname(path.dirname(pkgEntry));
  const bundlePath = path.join(pkgRoot, "dist", "dom-to-pptx.bundle.js");
  if (!fs.existsSync(bundlePath)) {
    throw new Error(`dom-to-pptx UMD bundle not found at ${bundlePath}`);
  }

  const browser = await puppeteer.launch({
    executablePath: args.chrome,
    headless: "new",
    args: [
      "--no-sandbox",
      "--disable-dev-shm-usage",
      "--disable-gpu",
      "--hide-scrollbars",
      // Allow file:// URLs to load images via the same protocol.
      "--allow-file-access-from-files",
    ],
    defaultViewport: { width: 1920, height: 1080, deviceScaleFactor: 1 },
  });

  try {
    const page = await browser.newPage();
    // Surface any browser-side errors back to the Node stderr so the Python
    // orchestrator's traceback is helpful.
    page.on("console", (msg) => {
      const t = msg.type();
      if (t === "error" || t === "warning") {
        console.error(`[browser ${t}]`, msg.text());
      }
    });
    page.on("pageerror", (err) => {
      console.error("[browser pageerror]", err.message);
    });

    const fileUrl = "file://" + path.resolve(args.html).replace(/\\/g, "/");
    await page.goto(fileUrl, { waitUntil: "networkidle0", timeout: 60000 });

    // Wait for webfonts (Inter, Playfair Display) so dom-to-pptx reads correct
    // text metrics. Without this, line-heights and widths fall back to system
    // fonts and the exported shapes mis-align.
    await page.evaluate(async () => {
      if (document.fonts && document.fonts.ready) {
        await document.fonts.ready;
      }
      // Give one extra animation frame for any post-font layout settle.
      await new Promise((r) => requestAnimationFrame(() => r()));
    });

    // Inject dom-to-pptx UMD bundle. Exposes window.domToPptx.exportToPptx.
    await page.addScriptTag({ path: bundlePath });

    // Count emitted slide sections and build selectors.
    const selectors = await page.evaluate(() => {
      const nodes = Array.from(document.querySelectorAll('[id^="slide-"]'));
      return nodes.map((n) => "#" + n.id);
    });
    if (selectors.length === 0) {
      throw new Error("No slide sections found in HTML (expected #slide-1, #slide-2, …)");
    }

    // Drive dom-to-pptx → returns a Blob; we read it as base64 to ship back.
    const base64 = await page.evaluate(async (sels) => {
      if (!window.domToPptx || typeof window.domToPptx.exportToPptx !== "function") {
        throw new Error("dom-to-pptx UMD bundle did not expose window.domToPptx.exportToPptx");
      }
      const blob = await window.domToPptx.exportToPptx(sels, {
        skipDownload: true,
        autoEmbedFonts: true,
        // Keep our inline SVG charts (donut, etc.) as editable PowerPoint
        // vectors instead of rasterizing them. Sharp at any zoom.
        svgAsVector: true,
        // LAYOUT_WIDE = 13.333 × 7.5 inches, same dimensions as our previous
        // python-pptx output; PowerPoint's "Widescreen (16:9)" default.
        layout: "LAYOUT_WIDE",
      });
      if (!blob) throw new Error("exportToPptx returned no blob");
      return await new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onloadend = () => {
          const dataUrl = String(reader.result || "");
          const comma = dataUrl.indexOf(",");
          resolve(comma >= 0 ? dataUrl.slice(comma + 1) : "");
        };
        reader.onerror = () => reject(reader.error || new Error("FileReader failed"));
        reader.readAsDataURL(blob);
      });
    }, selectors);

    if (!base64) {
      throw new Error("Blob -> base64 returned empty payload");
    }

    const buf = Buffer.from(base64, "base64");
    fs.mkdirSync(path.dirname(path.resolve(args.out)), { recursive: true });
    fs.writeFileSync(args.out, buf);
    console.log(`WROTE: ${args.out} (${buf.length} bytes, ${selectors.length} slides)`);
  } finally {
    await browser.close();
  }
}

main().catch((err) => {
  console.error("dom_export failed:", err && err.stack ? err.stack : err);
  process.exit(1);
});
