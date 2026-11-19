#!/usr/bin/env node
import fs from "node:fs";
import path from "node:path";
import { spawnSync } from "node:child_process";

const usage = () => {
  console.error("Usage: node scripts/audit_html.mjs <report.html> [--no-scripts] [--fail-on-warning] [--template]");
  process.exit(2);
};

const args = process.argv.slice(2);
let file;
let noScripts = false;
let failOnWarning = false;
let templateMode = false;
for (let index = 0; index < args.length; index += 1) {
  const arg = args[index];
  if (arg === "--no-scripts") {
    noScripts = true;
  } else if (arg === "--fail-on-warning") {
    failOnWarning = true;
  } else if (arg === "--template") {
    templateMode = true;
  } else if (arg.startsWith("--") || file) {
    usage();
  } else {
    file = arg;
  }
}
if (!file) usage();

let html;
try {
  html = fs.readFileSync(file, "utf8");
} catch (error) {
  console.error(`Cannot read ${file}: ${error.message}`);
  process.exit(2);
}

const markup = html.replace(/<!--[\s\S]*?-->/g, "");
const count = pattern => (markup.match(pattern) || []).length;
const results = [];
const check = (name, pass, severity, detail) => results.push({ name, pass: Boolean(pass), severity, detail });
const escapeRegExp = value => value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
const attrValue = (tag, name) => {
  const pattern = new RegExp("(?:^|\\s)" + escapeRegExp(name) + "\\s*=\\s*(?:\"([^\"]*)\"|'([^']*)'|([^\\s\"'=<>]+))", "i");
  const match = tag.match(pattern);
  return match ? (match[1] ?? match[2] ?? match[3] ?? "") : undefined;
};
const allStartTags = [...markup.matchAll(/<([a-zA-Z][\w:-]*)\b[^>]*>/g)].map(match => match[0]);
const tagsNamed = name => allStartTags.filter(tag => new RegExp(`^<${escapeRegExp(name)}\\b`, "i").test(tag));
const htmlTag = tagsNamed("html")[0] || "";
const metaTags = tagsNamed("meta");
const anchorTags = tagsNamed("a");
const linkTags = tagsNamed("link");
const scriptTags = tagsNamed("script");

check("doctype", /^<!doctype html>/i.test(html.trimStart()), "error", "Use HTML5 doctype");
check("language", Boolean(attrValue(htmlTag, "lang")?.trim()), "error", "Set the document language");
check("viewport", metaTags.some(tag => attrValue(tag, "name")?.toLowerCase() === "viewport"), "error", "Add a responsive viewport meta tag");
check("title", /<title>\s*[^<]{4,}\s*<\/title>/i.test(markup), "error", "Use a descriptive title");
check("single h1", count(/<h1\b/gi) === 1, "error", `Found ${count(/<h1\b/gi)} h1 elements`);
check("section headings", count(/<h2\b/gi) >= 1, "error", `Found ${count(/<h2\b/gi)} h2 elements`);
check("header landmark", count(/<header\b/gi) === 1, "warning", `Found ${count(/<header\b/gi)} header elements`);
check("main landmark", count(/<main\b/gi) === 1, "error", `Found ${count(/<main\b/gi)} main elements`);
check("navigation", count(/<nav\b/gi) >= 1, "warning", "Add a reading map for multi-section reports");
check("footer landmark", count(/<footer\b/gi) === 1, "warning", `Found ${count(/<footer\b/gi)} footer elements`);

const skipTag = anchorTags.find(tag => ["#main", "#main-content", "#content"].includes(attrValue(tag, "href")?.toLowerCase()));
check("skip link", Boolean(skipTag), "warning", "Add a skip-to-content link");
if (skipTag) {
  const targetId = attrValue(skipTag, "href").slice(1);
  const target = allStartTags.find(tag => attrValue(tag, "id") === targetId) || "";
  const focusable = /^-?\d+$/.test(attrValue(target, "tabindex") || "");
  check("focusable skip target", focusable, "warning", focusable ? `#${targetId} is focusable` : `Add tabindex="-1" to #${targetId}`);
}
check("visible focus", /:focus-visible/.test(html), "warning", "Define a visible keyboard focus style");
check("reduced motion", /prefers-reduced-motion/i.test(html), "warning", "Respect reduced-motion preference");
check("print stylesheet", /@media\s+print/i.test(html), "warning", "Add print rules for internal sharing");
// Desktop is the default target. Breakpoints are optional; required mobile behavior is verified in-browser when in scope.
check("wide-content overflow", /overflow-x\s*:\s*auto/i.test(html) || count(/<table\b/gi) === 0, "warning", "Protect wide tables or diagrams from clipping");

const stripMarkup = value => value
  .replace(/<(script|style)\b[^>]*>[\s\S]*?<\/\1>/gi, " ")
  .replace(/<[^>]+>/g, " ")
  .replace(/&(?:nbsp|ensp|emsp|thinsp);/gi, " ")
  .replace(/&[a-z0-9#]+;/gi, "x")
  .replace(/\s+/g, " ")
  .trim();
const h1Text = stripMarkup(markup.match(/<h1\b[^>]*>([\s\S]*?)<\/h1>/i)?.[1] || "");
const h2Texts = [...markup.matchAll(/<h2\b[^>]*>([\s\S]*?)<\/h2>/gi)].map(match => stripMarkup(match[1]));
const mainMatch = markup.match(/<main\b([^>]*)>([\s\S]*?)<\/main>/i);
const mainStartTag = tagsNamed("main")[0] || "";
const mainText = stripMarkup(mainMatch?.[2] || "");
check("non-empty h1", h1Text.length >= 4, "error", h1Text || "Empty or missing h1 text");
check("non-empty h2", h2Texts.length > 0 && h2Texts.every(text => text.length >= 2), "error", h2Texts.every(Boolean) ? `${h2Texts.length} populated h2 elements` : "Empty h2 text found");
check("substantive static main", mainText.length >= 120, "error", `Found ${mainText.length} visible text characters in main`);
const mainStyle = attrValue(mainStartTag, "style") || "";
const mainHidden = /(?:^|\s)hidden(?:\s|=|>|$)/i.test(mainStartTag) || attrValue(mainStartTag, "aria-hidden")?.toLowerCase() === "true" || /(?:display\s*:\s*none|visibility\s*:\s*hidden)/i.test(mainStyle);
check("static main is visible", !mainHidden, "error", mainHidden ? "main is hidden without JavaScript" : "main is visible by default");

if (!templateMode) {
  const placeholderPatterns = [
    /YYYY(?:-MM(?:-DD)?)?/i,
    /字段内容/,
    /来源名称/,
    /文档标题/,
    /报告标题/,
    /章节标题/,
    /步骤名/,
    /为什么重要。?/,
    /写在这里/,
    /待填写/,
    /\bXXX-000\b/i,
    /lorem ipsum/i
  ];
  const placeholders = placeholderPatterns.filter(pattern => pattern.test(markup)).map(pattern => pattern.source);
  check("no template placeholders", placeholders.length === 0, "error", placeholders.join(", ") || "No known placeholders");
  const hasDemoMarkers = /\bdata-demo=/i.test(markup);
  check("no data-demo markers", !hasDemoMarkers, "error", hasDemoMarkers ? "Remove template demo markers before delivery" : "No data-demo markers");
}

const isEmbeddedRef = value => {
  const ref = value.trim().toLowerCase();
  return !ref || ref.startsWith("data:") || ref.startsWith("#") || ref === "about:blank";
};
const externalScripts = scriptTags
  .map(tag => attrValue(tag, "src"))
  .filter(Boolean)
  .filter(value => !isEmbeddedRef(value));
const externalStyles = linkTags
  .filter(tag => (attrValue(tag, "rel") || "").toLowerCase().split(/\s+/).includes("stylesheet"))
  .map(tag => attrValue(tag, "href"))
  .filter(value => value && !isEmbeddedRef(value));

const cssBlocks = [...html.matchAll(/<style\b[^>]*>([\s\S]*?)<\/style>/gi)].map(match => match[1]);
const inlineStyles = allStartTags.map(tag => attrValue(tag, "style")).filter(Boolean);
const css = [...cssBlocks, ...inlineStyles].join("\n");
const cssImports = [...css.matchAll(/@import\s+([^;]+);?/gi)].map(match => match[1].trim());
const cssUrls = [...css.matchAll(/url\(\s*(["']?)(.*?)\1\s*\)/gi)]
  .map(match => match[2])
  .filter(value => !isEmbeddedRef(value));

const mainSectionTags = [...(mainMatch?.[2] || "").matchAll(/<section\b[^>]*>/gi)].map(match => match[0]);
const sectionClassCounts = new Map();
for (const tag of mainSectionTags) {
  const classes = (attrValue(tag, "class") || "").split(/\s+/).filter(Boolean);
  for (const className of classes) sectionClassCounts.set(className, (sectionClassCounts.get(className) || 0) + 1);
}
const sharedSectionClasses = [...sectionClassCounts.entries()].filter(([, classCount]) => classCount >= 2);
const hiddenSectionSignals = [];
const progressiveGate = /(?:^|[\s>+~])(?:html|body)?\.(?:js|focus-mode|tabs-enabled|is-enhanced|enhanced)\b|\[data-(?:enhanced|mode|view)(?:[=\]\s])/i;
const cssWithoutComments = cssBlocks.join("\n").replace(/\/\*[\s\S]*?\*\//g, "");
for (const rule of cssWithoutComments.matchAll(/([^{}]+)\{([^{}]*)\}/g)) {
  if (!/(?:display\s*:\s*none|visibility\s*:\s*hidden)/i.test(rule[2])) continue;
  for (const selector of rule[1].split(",").map(value => value.trim()).filter(Boolean)) {
    if (selector.startsWith("@") || progressiveGate.test(selector)) continue;
    for (const [className, classCount] of sharedSectionClasses) {
      if (new RegExp(`\\.${escapeRegExp(className)}(?![\\w-])`).test(selector)) hiddenSectionSignals.push(`.${className} (${classCount} sections via ${selector})`);
    }
  }
}
for (const tag of mainSectionTags) {
  const style = attrValue(tag, "style") || "";
  if (/(?:display\s*:\s*none|visibility\s*:\s*hidden)/i.test(style)) hiddenSectionSignals.push(`#${attrValue(tag, "id") || "unnamed-section"} (inline style)`);
}
const uniqueHiddenSectionSignals = [...new Set(hiddenSectionSignals)];
check("major sections visible without JS", uniqueHiddenSectionSignals.length === 0, "warning", uniqueHiddenSectionSignals.join(", ") || `${mainSectionTags.length} section(s) have no default-hidden signal`);

const assetRefs = [];
for (const match of markup.matchAll(/<(img|source|video|audio|iframe|embed|object|track|input|use|image|feimage)\b[^>]*>/gi)) {
  const tag = match[0];
  const name = match[1].toLowerCase();
  const names = name === "object" ? ["data"] : ["use", "image", "feimage"].includes(name) ? ["href", "xlink:href"] : name === "video" ? ["src", "poster"] : ["img", "source"].includes(name) ? ["src", "srcset"] : ["src"];
  for (const attribute of names) {
    const value = attrValue(tag, attribute);
    if (value && !isEmbeddedRef(value)) assetRefs.push(`${name}[${attribute}]=${value}`);
  }
}
for (const tag of linkTags) {
  const rel = attrValue(tag, "rel")?.toLowerCase() || "";
  if (!/(?:^|\s)(?:icon|preload|prefetch|modulepreload|manifest|apple-touch-icon)(?:\s|$)/.test(rel)) continue;
  const value = attrValue(tag, "href");
  if (value && !isEmbeddedRef(value)) assetRefs.push(`link[rel=${rel}]=${value}`);
}

check("no external scripts", externalScripts.length === 0, "error", externalScripts.join(", ") || "Self-contained");
check("no external styles", externalStyles.length === 0, "error", externalStyles.join(", ") || "Self-contained");
check("no css @import", cssImports.length === 0, "error", cssImports.join(", ") || "No @import");
check("no external css url()", cssUrls.length === 0, "error", cssUrls.join(", ") || "No external CSS assets");
check("no external media assets", assetRefs.length === 0, "error", assetRefs.join(", ") || "No external media assets");

const ids = allStartTags.map(tag => attrValue(tag, "id")).filter(Boolean);
const duplicateIds = [...new Set(ids.filter((id, index) => ids.indexOf(id) !== index))];
check("unique ids", duplicateIds.length === 0, "error", duplicateIds.join(", ") || "All ids are unique");

const idSet = new Set(ids);
const anchors = anchorTags
  .map(tag => attrValue(tag, "href"))
  .filter(value => value?.startsWith("#"))
  .map(value => value.slice(1));
const missingAnchors = [...new Set(anchors.filter(anchor => anchor && !idSet.has(anchor)))];
check("internal anchors", missingAnchors.length === 0, "error", missingAnchors.join(", ") || "All targets exist");

const balanceErrors = [];
const voidTags = new Set(["area", "base", "br", "col", "embed", "hr", "img", "input", "link", "meta", "param", "source", "track", "wbr"]);
const balanceSource = markup.replace(/<(script|style)\b([^>]*)>[\s\S]*?<\/\1\s*>/gi, (_match, tag, attributes) => `<${tag}${attributes}></${tag}>`);
const stack = [];
for (const match of balanceSource.matchAll(/<\s*(\/?)\s*([a-zA-Z][\w:-]*)\b([^<>]*?)>/g)) {
  const closing = Boolean(match[1]);
  const tag = match[2].toLowerCase();
  const selfClosing = /\/\s*>$/.test(match[0]) || voidTags.has(tag);
  if (!closing && !selfClosing) {
    stack.push(tag);
  } else if (closing) {
    if (stack.at(-1) === tag) {
      stack.pop();
    } else {
      const found = stack.lastIndexOf(tag);
      balanceErrors.push(`Unexpected </${tag}>; open stack ends with <${stack.at(-1) || "none"}>`);
      if (found >= 0) stack.splice(found);
    }
  }
}
if (stack.length) balanceErrors.push(`Unclosed tags: ${stack.map(tag => `<${tag}>`).join(", ")}`);
check("balanced html tags", balanceErrors.length === 0, "error", balanceErrors.join("; ") || "Tags are balanced");

const inlineScripts = [...markup.matchAll(/<script\b([^>]*)>([\s\S]*?)<\/script>/gi)]
  .filter(match => !attrValue(`<script ${match[1]}>`, "src"))
  .filter(match => {
    const type = attrValue(`<script ${match[1]}>`, "type")?.toLowerCase();
    return !type || type === "text/javascript" || type === "application/javascript" || type === "module";
  })
  .map(match => ({ source: match[2], module: attrValue(`<script ${match[1]}>`, "type")?.toLowerCase() === "module" }));
const scriptErrors = [];
const networkScripts = [];
inlineScripts.forEach(({ source, module }, index) => {
  try {
    if (module) {
      const parsed = spawnSync(process.execPath, ["--check", "--input-type=module"], { input: source, encoding: "utf8" });
      if (parsed.status !== 0) throw new Error(parsed.error?.message || parsed.stderr.trim() || "Module syntax check failed");
    } else {
      new Function(source);
    }
  } catch (error) {
    scriptErrors.push(`script ${index + 1}: ${error.message}`);
  }
  if (/\b(?:fetch|XMLHttpRequest|WebSocket|EventSource)\s*\(/.test(source) || /\bnavigator\.sendBeacon\s*\(/.test(source) || /\bimport\s*\(/.test(source) || /\bserviceWorker\.register\s*\(/.test(source) || (module && /\b(?:import|export)\s+(?:[^;]*?\bfrom\s*)?["']/.test(source))) {
    networkScripts.push(`script ${index + 1}`);
  }
});
check("inline script syntax", scriptErrors.length === 0, "error", scriptErrors.join("; ") || `${inlineScripts.length} inline script(s) parsed`);
check("no script network APIs", networkScripts.length === 0, "error", networkScripts.join(", ") || "No recognized network APIs or module dependencies (static heuristic)");

const cssHidesMain = /(?:^|[}\s])(?:body|main|#main|#main-content|#content)\s*\{[^}]*\b(?:display\s*:\s*none|visibility\s*:\s*hidden)/i.test(css);
check("css keeps main visible", !cssHidesMain, "error", cssHidesMain ? "CSS hides body/main by default" : "CSS keeps static content visible");

if (noScripts) {
  const anyScript = count(/<script\b/gi);
  check("zero script tags", anyScript === 0, "error", anyScript ? `Found ${anyScript} script tag(s)` : "Zero-JS confirmed");
}
console.log(`\nHTML structural audit [scripts: ${noScripts ? "none" : "progressive"}${templateMode ? ", template" : ""}]: ${path.resolve(file)}\n`);
for (const item of results) {
  const mark = item.pass ? "PASS" : item.severity === "error" ? "FAIL" : "WARN";
  console.log(`${mark.padEnd(4)}  ${item.name.padEnd(30)} ${item.detail}`);
}

const errors = results.filter(item => !item.pass && item.severity === "error");
const warnings = results.filter(item => !item.pass && item.severity === "warning");
console.log(`\n${errors.length} error(s), ${warnings.length} warning(s)\n`);
process.exit(errors.length || (failOnWarning && warnings.length) ? 1 : 0);
