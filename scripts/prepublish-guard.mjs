#!/usr/bin/env node
/**
 * Publish gate.
 *
 * Runs as the last step of `prepublishOnly`, after the build, so it inspects the
 * artefact that is actually about to be uploaded. It refuses the two mistakes that
 * are only visible once the package is on the registry — ownership metadata that
 * still points at a placeholder, and a `main`/`exports` map that promises files the
 * build never wrote — plus the packaging mistakes that are merely embarrassing
 * (sources or tests inside the tarball, no licence, no README).
 *
 * It never talks to the registry; `npm publish --dry-run` remains the rehearsal.
 */

import { existsSync, readFileSync, statSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const problems = [];

const pkg = JSON.parse(readFileSync(resolve(root, "package.json"), "utf8"));

// --- ownership metadata ------------------------------------------------------

/** Owner names people forget to replace before their first release. */
const PLACEHOLDER =
  /\b(owner|your[-_ ]?(user|name|org|scope|username)|username|example\.(com|org|net)|todo|fixme|changeme)\b/i;

const links = {
  repository:
    typeof pkg.repository === "string" ? pkg.repository : (pkg.repository?.url ?? undefined),
  homepage: pkg.homepage,
  "bugs.url": pkg.bugs?.url,
};

for (const [field, value] of Object.entries(links)) {
  // An absent link is a deliberate choice; a placeholder is not.
  if (value === undefined) continue;
  if (PLACEHOLDER.test(value)) {
    problems.push(`${field} still points at a placeholder (${value}) — set the real URL or drop the field`);
  } else if (!/^(https?:\/\/|git\+https?:\/\/|github:)/.test(value)) {
    problems.push(`${field} is not an http(s) or git+ URL (${value})`);
  }
}

for (const field of ["name", "version", "description", "license"]) {
  if (typeof pkg[field] !== "string" || !pkg[field]) {
    problems.push(`package.json has no "${field}"`);
  }
}

if (!/^\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?$/.test(pkg.version ?? "")) {
  problems.push(`version "${pkg.version}" is not a semver release version`);
}

if (pkg.private === true) {
  problems.push("package.json is marked private");
}

// --- what the tarball is allowed to contain ----------------------------------

const files = Array.isArray(pkg.files) ? pkg.files : [];
const NOT_SHIPPABLE = /^(src|tests?|examples?|scripts|node_modules|\*)/;

if (files.length === 0) {
  problems.push('"files" is missing — everything, including sources and tests, would be uploaded');
}

for (const entry of files) {
  if (NOT_SHIPPABLE.test(entry)) {
    problems.push(`"files" includes "${entry}", which would ship sources or tests`);
  } else if (entry !== "dist" && !existsSync(resolve(root, entry))) {
    problems.push(`"files" lists "${entry}", which does not exist`);
  }
}

for (const required of ["dist", "README.md", "LICENSE"]) {
  if (!files.includes(required)) {
    problems.push(`"files" does not include "${required}"`);
  } else if (required !== "dist" && !existsSync(resolve(root, required))) {
    problems.push(`"files" includes "${required}", but it is not in the working tree`);
  }
}

// --- what the entry points promise -------------------------------------------

const entries = new Set();
const collect = (value) => {
  if (typeof value === "string") {
    if (value.startsWith("./")) entries.add(value);
    return;
  }
  if (value && typeof value === "object") {
    for (const nested of Object.values(value)) collect(nested);
  }
};

for (const field of ["main", "module", "types", "unpkg", "jsdelivr"]) {
  collect(pkg[field]);
}
collect(pkg.exports);

for (const target of [...entries].sort()) {
  const path = resolve(root, target);
  if (!existsSync(path)) {
    problems.push(`${target} is promised by main/exports but missing — did the build run?`);
  } else if (target.endsWith(".css") && statSync(path).size === 0) {
    problems.push(`${target} is empty`);
  }
}

for (const required of ["./dist/index.js", "./dist/index.cjs", "./dist/index.d.ts", "./dist/styles.css"]) {
  if (!entries.has(required)) {
    problems.push(`${required} is no longer reachable through main/module/types/exports`);
  }
}

// --- verdict -----------------------------------------------------------------

if (problems.length > 0) {
  process.stderr.write(
    `publish gate: refusing to publish ${pkg.name}@${pkg.version} — fix these first:\n` +
      problems.map((problem) => `  • ${problem}\n`).join(""),
  );
  process.exit(1);
}

process.stdout.write(
  `publish gate: ${pkg.name}@${pkg.version} — metadata, "files" allowlist and ${entries.size} entry points check out\n`,
);
