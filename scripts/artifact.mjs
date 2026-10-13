// Turns dist/ (npm run build) into a page claude.ai can publish as an Artifact:
// the page's own content without the html/head/body wrapper the publisher adds,
// and every other file beside it. Prints the files map for the Artifact tool.
//   node scripts/artifact.mjs <outdir>
import fs from "node:fs";
import path from "node:path";

const DIST = path.resolve(path.dirname(new URL(import.meta.url).pathname), "..", "dist");
const OUT = path.resolve(process.argv[2] || "artifact");
fs.rmSync(OUT, { recursive: true, force: true });
fs.mkdirSync(OUT, { recursive: true });

const html = fs.readFileSync(path.join(DIST, "index.html"), "utf8");
const head = (/<head>([\s\S]*?)<\/head>/i.exec(html) || [])[1] || "";
const body = (/<body>([\s\S]*?)<\/body>/i.exec(html) || [])[1] || "";
const title = (/<title>[\s\S]*?<\/title>/i.exec(head) || [""])[0];
const styles = [...head.matchAll(/<style>[\s\S]*?<\/style>/gi)].map(m => m[0]).join("\n");
const links = [...head.matchAll(/<link rel="stylesheet"[^>]*>/gi)].map(m => m[0].replace(/href="\.\//, 'href="')).join("\n");
const scripts = [...head.matchAll(/<script type="module"[^>]*><\/script>/gi)].map(m => m[0].replace(/src="\.\//, 'src="')).join("\n");
fs.writeFileSync(path.join(OUT, "index.html"), [title, styles, links, body.trim(), scripts].filter(Boolean).join("\n") + "\n");

const files = {};
const copy = rel => {
  const from = path.join(DIST, rel), to = path.join(OUT, rel);
  fs.mkdirSync(path.dirname(to), { recursive: true });
  fs.copyFileSync(from, to);
  files[rel.split(path.sep).join("/")] = to;
};
const walk = dir => {
  for(const f of fs.readdirSync(path.join(DIST, dir))){
    const rel = path.join(dir, f);
    if(fs.statSync(path.join(DIST, rel)).isDirectory()) walk(rel);
    else if(/\.xlsx$/i.test(rel)){
      /* claude.ai serves no .xlsx beside a page: the page reads the same bytes as base64 text (main.js, sampleBytes) */
      const to = path.join(OUT, rel + ".b64.txt");
      fs.writeFileSync(to, fs.readFileSync(path.join(DIST, rel)).toString("base64"));
      files[(rel + ".b64.txt").split(path.sep).join("/")] = to;
    }
    else if(rel !== "index.html") copy(rel);
  }
};
walk("");
fs.writeFileSync(path.join(OUT, "files.json"), JSON.stringify(files, null, 1));
const big = Object.entries(files).map(([k, v]) => [k, fs.statSync(v).size]).sort((a, b) => b[1] - a[1]);
console.log(`${Object.keys(files).length} files beside the page; largest:`, big.slice(0, 4).map(([k, s]) => `${k} ${(s / 1048576).toFixed(1)}MB`).join(", "));
console.log("page:", fs.readFileSync(path.join(OUT, "index.html"), "utf8").length, "characters");
