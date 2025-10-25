# Playwright tests for diyaa.shop

This folder contains a minimal Playwright test that navigates to `https://diyaa.in/login`.

## Building a Windows executable

You can build a standalone Windows exe from `run.js` using `pkg`.

1. Install dev dependencies:

```powershell
npm install
```

2. Build the exe:

```powershell
npm run build:exe
```

This produces `diyaa-shop.exe` in the project root. Note: Playwright's browser binaries are large and may not be fully bundled by `pkg`. It's recommended to keep the `browsers/` directory alongside the exe or set `PLAYWRIGHT_BROWSERS_PATH` to a known location.

Usage notes:

- To run tests against a custom base URL, run:

- Set the base URL interactively and run tests:

  npm run test:with-baseurl

  This runs `npm run set-baseurl` which prompts for a Base URL (default: `https://diyaa.in`), writes it to a `.env` file, then runs `playwright test`.

- Or run the prompt alone to only set the base URL:

  npm run set-baseurl

The Playwright config reads BASE_URL from the `.env` file and falls back to `https://diyaa.in`.

Prerequisites

- Node.js (16+)

Install and run (PowerShell):

```powershell
npm init -y; npm i -D @playwright/test; npx playwright install
npx playwright test tests/login.spec.js
```

The test will open a browser and assert the URL contains `/login`. Adjust as needed.
