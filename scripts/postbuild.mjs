#!/usr/bin/env node
/**
 * Post-build tidy-up.
 *
 * The ESM build keeps its source map (that is what bundlers and bundler-based
 * consumers resolve), but the CJS and UMD maps triple the unpacked package size
 * for very little benefit, so they are removed.
 */

import { rmSync } from "node:fs";

const redundant = ["dist/index.cjs.map", "dist/index.umd.cjs.map"];

for (const file of redundant) {
  rmSync(file, { force: true });
}

process.stdout.write(`postbuild: removed ${redundant.join(", ")}\n`);
