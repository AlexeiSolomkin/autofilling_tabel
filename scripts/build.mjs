// Сборка: src/*.ts → один самодостаточный dist/index.html.
import { build } from "esbuild";
import { mkdirSync, readFileSync, writeFileSync } from "node:fs";

const result = await build({
  entryPoints: ["src/main.ts"],
  bundle: true,
  format: "iife",
  target: "es2020",
  minify: true,
  write: false,
  legalComments: "none",
  logLevel: "warning",
});

const js = result.outputFiles[0].text.replace(/<\/script/gi, "<\\/script");
const template = readFileSync("src/index.html", "utf8");
const html = template.replace("/*BUNDLE*/", () => js);
mkdirSync("dist", { recursive: true });
writeFileSync("dist/index.html", html);
console.log(`dist/index.html собран: ${(html.length / 1024).toFixed(0)} КБ`);
