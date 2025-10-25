// runtime base URL (strip trailing slash for consistent joining)
const BASE_URL = (process.env.BASE_URL || "https://diyaa.in").replace(
  /\/$/,
  ""
);

class LoginPage {
  /**
   * @param {import('@playwright/test').Page} page
   */
  constructor(page) {
    this.page = page;
    // default selectors - can be extended
    this.usernameSelectors = [
      'input[name="username"]',
      'input[name="email"]',
      'input[type="email"]',
      'input[type="text"]',
    ];
    this.passwordSelectors = [
      'input[name="password"]',
      'input[type="password"]',
    ];
    this.submitSelectors = [
      'button[type="submit"]',
      'button:has-text("Log in")',
      'button:has-text("Login")',
      'input[type="submit"]',
    ];
  }

  async goto() {
    await this.page.goto(`${BASE_URL}/login`);
    await this.page.waitForLoadState("networkidle");
  }

  async _findVisible(selectors) {
    for (const sel of selectors) {
      const el = await this.page.$(sel);
      if (el) {
        try {
          if (await el.isVisible()) return sel;
        } catch (e) {
          // ignore and continue
        }
      }
    }
    return null;
  }

  async fillUsername(username, options = { delay: 120 }) {
    const sel = await this._findVisible(this.usernameSelectors);
    if (!sel) throw new Error("Username field not found");
    await this.page.click(sel);
    await this.page.type(sel, username, options);
    return sel;
  }

  async fillPassword(password, options = { delay: 120 }) {
    const sel = await this._findVisible(this.passwordSelectors);
    if (!sel) throw new Error("Password field not found");
    await this.page.click(sel);
    await this.page.type(sel, password, options);
    return sel;
  }

  async submit() {
    const sel = await this._findVisible(this.submitSelectors);
    if (sel) {
      await this.page.click(sel);
    } else {
      // fallback: press Enter in password field if available
      const passSel = await this._findVisible(this.passwordSelectors);
      if (passSel) await this.page.press(passSel, "Enter");
    }
  }

  // Navigate the left sidebar and open Orders > Order List
  async selectOrderList() {
    // Wait for the sidebar menu to be present on the master page
    await this.page.waitForSelector("ul.sidebar-menu", {
      state: "visible",
      timeout: 10000,
    });

    // Try to expand the "Orders" treeview by clicking its anchor
    const ordersAnchor = await this.page.$(
      'ul.sidebar-menu a:has-text("Orders")'
    );
    if (ordersAnchor) {
      try {
        await ordersAnchor.click();
      } catch (e) {
        // ignore click errors and continue to find the link
      }
    }

    // Click the Order List link. Try common href and text-based selectors.
    const orderListHref = `${BASE_URL}/inventory/order_list`;
    const orderListSelector = `a[href="${orderListHref}"], a[href="/inventory/order_list"], ul.treeview-menu a:has-text("Order List")`;
    await this.page.waitForSelector(orderListSelector, {
      state: "visible",
      timeout: 10000,
    });
    await this.page.click(orderListSelector);

    // Wait for navigation or the page to load
    try {
      await this.page.waitForLoadState("networkidle");
    } catch (e) {
      // ignore load state timeout
    }
  }

  /**
   * Check address popup content in `#addressShowBody` before closing it.
   * - waits briefly for the element
   * - verifies there's at least one address character (letter or digit)
   * - extracts up to 6 digits following the text 'pincode:' (case-insensitive) per row
   * - attempts to close the popup by common close buttons or Escape
   * @returns {{foundAddress: boolean, pincode: string|null, pincodes: Array, rawText: string|null}}
   */
  async handleAddressPopup() {
    const selector = "#addressShowBody";
    // wait briefly for the popup to appear; if it doesn't, return defaults
    try {
      await this.page.waitForSelector(selector, {
        state: "visible",
        timeout: 2000,
      });
    } catch (e) {
      return { foundAddress: false, pincode: null, rawText: null };
    }

    const el = await this.page.$(selector);
    if (!el) return { foundAddress: false, pincode: null, rawText: null };

    const rawText = (await el.innerText()).trim();

    // check for at least one address character (letter or digit)
    const hasAddressChar = /[A-Za-z0-9]/.test(rawText);

    // find 'Pincode:' case-insensitive and capture up to 6 digits after the colon per line
    let pincode = null; // keep first found for backward compatibility
    const pincodes = [];
    const lines = rawText.split(/\r?\n/);
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i] || "";
      const m = line.match(/Pincode\s*:\s*(\d{1,6})/i);
      if (m && m[1]) {
        const code = m[1].trim();
        pincodes.push({ row: i + 1, pincode: code, line: line.trim() });
        // set the single pincode if not set yet (backwards compatibility)
        if (!pincode) pincode = code;
        // log each found pincode per row
        // eslint-disable-next-line no-console
        console.log(`[handleAddressPopup] row ${i + 1} pincode: ${code}`);
      }
    }

    // try common close controls inside the popup
    const closeSelectors = [
      'button:has-text("Close")',
      `${selector} .close`,
      `${selector} .btn-close`,
      `${selector} button.close`,
      `${selector} button:has-text("×")`,
    ];
    for (const cs of closeSelectors) {
      try {
        const closeEl = await this.page.$(cs);
        if (closeEl) {
          try {
            await closeEl.click();
            // give the UI a moment to update
            await this.page.waitForTimeout(200);
            break;
          } catch (e) {
            // ignore click errors and continue
          }
        }
      } catch (e) {
        // ignore selector errors
      }
    }

    // as a fallback, press Escape to close modal/popup
    try {
      await this.page.keyboard.press("Escape");
      await this.page.waitForTimeout(150);
    } catch (e) {
      // ignore
    }

    // If no pincodes found, log summary
    if (pincodes.length === 0) {
      // eslint-disable-next-line no-console
      console.log("[handleAddressPopup] no pincodes found in popup");
    }

    return { foundAddress: hasAddressChar, pincode, pincodes, rawText };
  }
}

module.exports = { LoginPage };
