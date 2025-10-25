#!/usr/bin/env node
const fs = require("fs");
const path = require("path");
const readline = require("readline");

const DEFAULT = "https://diyaa.in";
const envPath = path.resolve(__dirname, "..", ".env");

function isValidUrl(u) {
  try {
    // allow plain hostnames by prefixing https:// if missing
    new URL(u.startsWith("http") ? u : `https://${u}`);
    return true;
  } catch (e) {
    return false;
  }
}

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout,
});
rl.question(`Base URL [${DEFAULT}]: `, (answer) => {
  let val = (answer && answer.trim()) || DEFAULT;
  if (!isValidUrl(val)) {
    // try to prefix https:// and revalidate
    if (isValidUrl(`https://${val}`)) val = `https://${val}`;
    else {
      console.error("Invalid URL. Aborting.");
      process.exit(2);
    }
  }

  const content = `BASE_URL=${val}\n`;
  try {
    fs.writeFileSync(envPath, content, { encoding: "utf8" });
    console.log(`Wrote ${envPath}`);
    console.log(`Using BASE_URL=${val}`);
  } catch (e) {
    console.error("Failed to write .env file:", e.message);
    process.exit(1);
  }
  rl.close();
});
