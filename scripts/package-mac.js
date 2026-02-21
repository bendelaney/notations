const { execFileSync } = require("node:child_process");
const fs = require("node:fs");
const path = require("node:path");

const ROOT = path.resolve(__dirname, "..");
const RESOURCES_DIR = path.join(
  ROOT,
  "release",
  "Notations-darwin-arm64",
  "Notations.app",
  "Contents",
  "Frameworks",
  "Electron Framework.framework",
  "Versions",
  "A",
  "Resources"
);

function runPackager() {
  const args = [
    "@electron/packager",
    ".",
    "Notations",
    "--platform=darwin",
    "--arch=arm64",
    "--icon=assets/icons/notations.icns",
    "--overwrite",
    "--out=release",
    "--asar",
    "--prune=true",
    "--junk=true",
    "--ignore=^/release(/|$)",
    "--ignore=^/Notations-darwin-arm64(/|$)",
    "--ignore=^/mockups(/|$)",
    "--ignore=^/assets/paper_textures(/|$)",
    "--ignore=^/assets/icons/MyIcon\\.icns$",
    "--ignore=^/assets/icons/MyIcon\\.iconset(/|$)",
    "--ignore=^/assets/icons/notations\\.icon(/|$)",
    "--ignore=^/\\.DS_Store$"
  ];

  execFileSync("npx", args, { cwd: ROOT, stdio: "inherit" });
}

function getLocaleKeepList() {
  const raw = process.env.NOTATIONS_KEEP_LOCALES;
  if (!raw || !raw.trim()) {
    return new Set(["en.lproj", "en_GB.lproj"]);
  }

  return new Set(
    raw
      .split(",")
      .map((value) => value.trim())
      .filter(Boolean)
      .map((value) => (value.endsWith(".lproj") ? value : `${value}.lproj`))
  );
}

function pruneElectronLocales() {
  if (!fs.existsSync(RESOURCES_DIR)) return;

  const keepLocales = getLocaleKeepList();
  for (const entry of fs.readdirSync(RESOURCES_DIR, { withFileTypes: true })) {
    if (!entry.isDirectory() || !entry.name.endsWith(".lproj")) continue;
    if (keepLocales.has(entry.name)) continue;
    fs.rmSync(path.join(RESOURCES_DIR, entry.name), { recursive: true, force: true });
  }
}

runPackager();
pruneElectronLocales();
