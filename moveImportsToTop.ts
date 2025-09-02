#! /usr/bin/env node

import { readFileSync, writeFileSync } from "fs";

const lines = readFileSync("dist/main.js", "utf-8").split("\n");

writeFileSync(
  "dist/main.js",
  [
    ...lines.filter((l) => l.trim().startsWith("import ")),
    ...lines.filter((l) => !l.trim().startsWith("import ")),
  ].join("\n")
);
