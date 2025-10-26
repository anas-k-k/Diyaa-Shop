const path = require("path");
const fs = require("fs");
let xlsx;
try {
  xlsx = require("xlsx");
} catch (e) {
  // If xlsx isn't installed, we'll fail later when attempting to read files.
  xlsx = null;
}

// Optional list of order IDs to process during a run. If this Set is
// non-empty, only orders whose IDs appear in this Set will be processed.
// If empty, all orders will be processed. Order IDs are stored as strings
// for consistent comparison.
const ORDERS_TO_PROCESS = new Set([]);

// Extract pincode(s) from a raw text blob.
// Returns an array of numeric pincodes as strings (e.g. ['689672']).
// Matches formats like 'Pincode : 689672', 'Pincode:689672', or plain 6-digit numbers.
function extractPincode(rawText) {
  if (!rawText || typeof rawText !== "string") return [];

  // normalize and search for 6-digit sequences which are typical pincodes
  const candidates = [];

  // First try to find patterns like 'Pincode *: *123456'
  const labelled = rawText.match(/Pincode\s*[:\-]?\s*(\d{4,6})/gi);
  if (labelled) {
    for (const m of labelled) {
      const num = m.match(/(\d{4,6})/);
      if (num) candidates.push(num[1]);
    }
  }

  // Fallback: find any 4-6 digit sequences (some pincodes may be 4-6 digits depending on locale)
  if (!candidates.length) {
    const any = rawText.match(/\b\d{4,6}\b/g);
    if (any) candidates.push(...any);
  }

  // dedupe while keeping order
  return [...new Set(candidates)];
}

// Extract state from a raw text blob.
// Returns the state name as a string (e.g. 'Tamil Nadu').
// Matches formats like 'State : Tamil Nadu', 'State:Tamil Nadu', etc.
function extractState(rawText) {
  if (!rawText || typeof rawText !== "string") return null;

  // Look for patterns like 'State *: *value'
  const stateMatch = rawText.match(/State\s*[:\-]?\s*([^,\n\r]+)/i);
  if (stateMatch && stateMatch[1]) {
    return stateMatch[1].trim();
  }

  return null;
}

// Read first-column values from the first sheet of an Excel file and return a Set of strings
function readPincodesFromExcel(absPath) {
  if (!xlsx) return new Set();
  try {
    if (!fs.existsSync(absPath)) return new Set();
    const wb = xlsx.readFile(absPath);
    const sheetName = wb.SheetNames && wb.SheetNames[0];
    if (!sheetName) return new Set();
    const sheet = wb.Sheets[sheetName];
    const rows = xlsx.utils.sheet_to_json(sheet, { header: 1 });
    const set = new Set();
    for (const r of rows) {
      if (!r || r.length === 0) continue;
      const v = String(r[0]).trim();
      if (v) set.add(v);
    }
    return set;
  } catch (e) {
    return new Set();
  }
}

// Cache for Excel lookups to avoid re-reading files repeatedly
// Dynamic cache that can store any carrier name as key
const _excelCache = {
  carriers: new Map(), // Map<carrierName, Set<pincode>>
  lastLoaded: new Map(), // Map<carrierName, timestamp>
  // reload interval in ms (optional) - set to 60s to allow occasional refresh
  reloadInterval: 60 * 1000,
};

// Load Excel cache for a specific carrier
function loadExcelCacheForCarrier(carrierName) {
  if (!carrierName) return null;

  const now = Date.now();
  const normalizedCarrier = carrierName.trim();

  // Check if we have a recent cache for this carrier
  const lastLoadTime = _excelCache.lastLoaded.get(normalizedCarrier);
  if (
    lastLoadTime &&
    now - lastLoadTime < _excelCache.reloadInterval &&
    _excelCache.carriers.has(normalizedCarrier)
  ) {
    return _excelCache.carriers.get(normalizedCarrier);
  }

  // Load fresh data for this carrier
  const dataDir = path.join(process.cwd(), "data");
  const carrierFileName = `${normalizedCarrier}.xlsx`;
  const carrierFilePath = path.join(dataDir, carrierFileName);

  const pincodeSet = readPincodesFromExcel(carrierFilePath);

  // Cache the result
  _excelCache.carriers.set(normalizedCarrier, pincodeSet);
  _excelCache.lastLoaded.set(normalizedCarrier, now);

  return pincodeSet;
}

// Legacy function for backward compatibility - now loads all known carriers
function loadExcelCaches() {
  // Load commonly used carriers (DTDC, Delhivery) for backward compatibility
  const commonCarriers = ["DTDC", "Delhivery"];
  commonCarriers.forEach((carrier) => {
    loadExcelCacheForCarrier(carrier);
  });
}

// Get cached pincode data for a specific carrier
function getCarrierPincodes(carrierName) {
  if (!carrierName) return null;
  const normalizedCarrier = carrierName.trim();

  // Try to get from cache first
  if (_excelCache.carriers.has(normalizedCarrier)) {
    const lastLoadTime = _excelCache.lastLoaded.get(normalizedCarrier);
    const now = Date.now();

    // If cache is still valid, return it
    if (lastLoadTime && now - lastLoadTime < _excelCache.reloadInterval) {
      return _excelCache.carriers.get(normalizedCarrier);
    }
  }

  // Cache miss or expired, load fresh data
  return loadExcelCacheForCarrier(normalizedCarrier);
}

// runtime base URL (strip trailing slash)
const BASE_URL = (process.env.BASE_URL || "https://diyaa.in").replace(
  /\/$/,
  ""
);

class OrderListPage {
  /**
   * @param {import('@playwright/test').Page} page
   */
  constructor(page) {
    this.page = page;
    // table and button selectors
    this.tableSelector = "table#example";
    // the per-row button as provided in the user request for first row
    // generalize to any row by using tbody tr td button with the same classes
    this.rowButtonSelector = "td.sorting_1 > button.address-show-btn";
    // common popup/modal close selectors to try
    this.popupCloseSelectors = [
      "#addressShowModal > div > div > div.modal-footer > button",
      ".modal:visible button.close",
      '.modal:visible button:has-text("Close")',
      ".modal:visible .close",
      ".bootbox-close-button",
      ".swal2-close",
      'button[aria-label="Close"]',
      'button:has-text("OK")',
      'button:has-text("Close")',
    ];
  }

  // Lightweight local handler for the address popup. Mirrors the behavior of
  // LoginPage.handleAddressPopup but kept here to avoid cross-file coupling.
  async handleAddressPopup(rowIndex = null, orderId = null) {
    const selector = "#addressShowBody";
    try {
      await this.page.waitForSelector(selector, {
        state: "visible",
        timeout: 1500,
      });
    } catch (e) {
      return { foundAddress: false, pincode: null, rawText: null };
    }

    const el = await this.page.$(selector);
    if (!el) return { foundAddress: false, pincode: null, rawText: null };

    const rawText = (await el.innerText()).trim();
    const hasAddressChar = /[A-Za-z0-9]/.test(rawText);

    // If the raw text is too short, close the popup and skip processing for this row
    const textLength = rawText.length;
    if (textLength < 150) {
      // log which row was skipped and the length
      // eslint-disable-next-line no-console
      console.log(
        `Skipping row ${rowIndex != null ? rowIndex : "?"} (orderId=${
          orderId || "N/A"
        }) - address text too short (${textLength} chars)`
      );

      // attempt to close modal using preferred close button or Escape
      const preferredClose =
        "#addressShowModal > div > div > div.modal-footer > button";
      try {
        const closeBtn = await this.page.$(preferredClose);
        if (closeBtn) {
          await closeBtn.click();
          await this.page.waitForTimeout(150);
        } else {
          await this.page.keyboard.press("Escape");
          await this.page.waitForTimeout(100);
        }
      } catch (e) {
        try {
          await this.page.keyboard.press("Escape");
          await this.page.waitForTimeout(100);
        } catch (ee) {
          // ignored
        }
      }

      return {
        foundAddress: false,
        pincode: null,
        rawText,
        textLength,
        orderId,
      };
    }

    // try to extract pincode(s) and state from the raw text and log them
    const pincodes = extractPincode(rawText);
    const pincode = pincodes.length ? pincodes[0] : null;
    const state = extractState(rawText);

    // eslint-disable-next-line no-console
    console.log(
      `Extracted pincode: ${pincode}, state: ${state} (orderId=${
        orderId || "N/A"
      })`
    );

    // First attempt: click the modal footer close button (preferred selector)
    const preferredClose =
      "#addressShowModal > div > div > div.modal-footer > button";
    try {
      const closeBtn = await this.page.$(preferredClose);
      if (closeBtn) {
        await closeBtn.click();
        // give modal a moment to close
        await this.page.waitForTimeout(150);
      } else {
        // fallback to pressing Escape if preferred button not found
        await this.page.keyboard.press("Escape");
        await this.page.waitForTimeout(100);
      }
    } catch (e) {
      // if clicking fails for any reason, fallback to Escape
      try {
        await this.page.keyboard.press("Escape");
        await this.page.waitForTimeout(100);
      } catch (e) {
        // ignored
      }
    }

    return {
      foundAddress: hasAddressChar,
      pincode,
      pincodes,
      state,
      rawText,
      orderId,
    };
  }

  // Open a new tab for the order sync page, click Sync with Shiprocket,
  // interact with the modal (select courier DTDC, choose radio, wait), then close.
  // This method is defensive and will return quickly if elements are not found.
  async syncShiprocketForOrder(
    orderId,
    {
      waitMs = 2500,
      pincode = null,
      state = null,
      paymentType = null,
      paymentStatus = null,
    } = {}
  ) {
    if (!orderId) return { synced: false, reason: "no-order-id" };

    const targetUrl = `${BASE_URL}/inventory/order/${orderId}`;
    // open new tab
    const context = this.page.context();
    const newPage = await context.newPage();
    try {
      await newPage.goto(targetUrl, {
        waitUntil: "domcontentloaded",
        timeout: 15000,
      });

      // wait for the sync button and click it
      try {
        await newPage.waitForSelector("#sync_shiprocket", {
          state: "visible",
          timeout: 4000,
        });

        // Set up dialog handler before clicking the sync button
        let dialogAppeared = false;
        const dialogHandler = async (dialog) => {
          dialogAppeared = true;
          console.log(
            `Dialog appeared for order ${orderId}: "${dialog.message()}"`
          );
          try {
            await dialog.dismiss(); // Dismiss the dialog
          } catch (e) {
            // ignore dialog dismiss failures
          }
        };

        // Listen for any dialog that might appear
        newPage.once("dialog", dialogHandler);

        await newPage.click("#sync_shiprocket");

        // Wait a moment to see if dialog appears
        await newPage.waitForTimeout(1000);

        // If dialog appeared, skip this row
        if (dialogAppeared) {
          console.log(
            `Skipping order ${orderId} due to dialog appearance. Closing tab and continuing to next row.`
          );
          return { synced: false, reason: "dialog-appeared", skipped: true };
        }

        // Remove the dialog listener if no dialog appeared
        newPage.removeListener("dialog", dialogHandler);
      } catch (e) {
        // couldn't find or click sync button
        return { synced: false, reason: "no-sync-button" };
      }

      // Wait for the logistics modal to appear (the selector for the dropdown wrapper)
      const dropdownWrapper =
        "#logisticsModal > div > div > div.modal-body > div > div > div.col-md-9 > div > span > span.selection > span";
      // track which carrier we selected for reporting (declare here so it's
      // visible later outside the dropdown-selection try/catch)
      let selectedCarrier = null;
      try {
        await newPage.waitForSelector(dropdownWrapper, {
          state: "visible",
          timeout: 5000,
        });
        // click to expand
        await newPage.click(dropdownWrapper);

        // select carrier based on CARRIER_OVERRIDE from .env
        let selected = false;
        try {
          // Helper to try selecting an option by visible text (case-insensitive)
          // with pincode validation against Excel files
          const trySelectByText = async (text, pincode = null) => {
            // Step 1: Check if there is an Excel file with carrier name in data folder
            if (pincode && text) {
              try {
                const dataDir = path.join(process.cwd(), "data");
                const carrierFileName = `${text.trim()}.xlsx`;
                const carrierFilePath = path.join(dataDir, carrierFileName);

                // Step 2: Check if Excel file exists
                if (fs.existsSync(carrierFilePath)) {
                  console.log(
                    `Found Excel file for carrier: ${text} at ${carrierFilePath}`
                  );

                  // Step 3: Load Excel cache for this specific carrier
                  const carrierPincodes = loadExcelCacheForCarrier(text.trim());

                  // Step 4: Check if pincode is present in the Excel file
                  if (
                    carrierPincodes &&
                    carrierPincodes.has(String(pincode).trim())
                  ) {
                    console.log(
                      `Pincode ${pincode} found in ${text} Excel file. Proceeding with carrier selection.`
                    );
                    // Proceed with carrier selection as pincode is valid
                  } else {
                    // Step 5: Pincode not found, don't select carrier and close popup
                    console.log(
                      `Pincode ${pincode} NOT found in ${text} Excel file. Skipping carrier selection.`
                    );
                    return false; // Don't proceed with selection
                  }
                } else {
                  console.log(
                    `No Excel file found for carrier: ${text}. Proceeding with default behavior.`
                  );
                  // Step 2: File doesn't exist, proceed as current code (no validation)
                }
              } catch (e) {
                console.warn(
                  `Error checking Excel file for carrier ${text}:`,
                  e.message
                );
                // On error, proceed with default behavior
              }
            }

            const selCandidates = [
              `#select2-logistics-results li.select2-results__option`,
              `ul.select2-results__options li.select2-results__option`,
              `#logisticsModal .select2-results__option`,
              `#logisticsModal .dropdown-menu li`,
              `#logisticsModal li`,
            ];
            for (const sel of selCandidates) {
              try {
                const locator = newPage
                  .locator(sel)
                  .filter({ hasText: new RegExp(text, "i") });
                const count = await locator.count();
                if (count > 0) {
                  await locator
                    .first()
                    .click({ timeout: 2000, force: true })
                    .catch(() => {});
                  await newPage.waitForTimeout(250);
                  return true;
                }
              } catch (e) {
                // ignore
              }
            }
            // fallback: try find in DOM under #logisticsModal (case-insensitive)
            try {
              const found = await newPage.evaluate((txt) => {
                const modal = document.querySelector("#logisticsModal");
                if (!modal) return false;
                const items = Array.from(
                  modal.querySelectorAll("li, option, div")
                );
                const match = items.find(
                  (i) =>
                    i.innerText &&
                    i.innerText.trim().toLowerCase() === txt.toLowerCase()
                );
                if (match) {
                  try {
                    match.click();
                  } catch (e) {
                    /* ignore */
                  }
                  return true;
                }
                return false;
              }, text);
              if (found) return true;
            } catch (e) {
              // ignore
            }
            return false;
          };

          // Always use CARRIER_OVERRIDE from .env if provided
          const carrierOverride = (process.env.CARRIER_OVERRIDE || "").trim();

          if (carrierOverride) {
            // Validate carrier-specific conditions before selection
            let shouldSelectCarrier = true;
            let skipReason = "";

            const carrierLower = carrierOverride.toLowerCase();
            const stateLower = (state || "").toLowerCase();
            const paymentTypeLower = (paymentType || "").toLowerCase();
            const paymentStatusLower = (paymentStatus || "").toLowerCase();

            // Condition 1: STCourier validation
            if (carrierLower === "stcourier") {
              if (
                stateLower !== "tamil nadu" ||
                !paymentTypeLower.includes("prepaid") ||
                paymentStatusLower !== "success"
              ) {
                shouldSelectCarrier = false;
                skipReason = `STCourier conditions not met: state='${state}' (should be 'Tamil Nadu'), paymentType='${paymentType}' (should contain 'prepaid'), paymentStatus='${paymentStatus}' (should be 'success')`;
              }
            }

            // Condition 2: Shiprocket validation
            if (carrierLower === "shiprocket") {
              // Shiprocket is eligible if:
              // (Any State AND COD AND Success) OR (Prepaid AND NOT Tamil Nadu AND Success)
              const isCODWithSuccess =
                paymentTypeLower.includes("cod") &&
                paymentStatusLower === "success";
              const isPrepaidNotTamilNaduWithSuccess =
                paymentTypeLower.includes("prepaid") &&
                stateLower !== "tamil nadu" &&
                paymentStatusLower === "success";

              const isEligible =
                isCODWithSuccess || isPrepaidNotTamilNaduWithSuccess;

              if (!isEligible) {
                shouldSelectCarrier = false;
                skipReason = `Shiprocket conditions not met: Either (COD + Success) or (Prepaid + NOT Tamil Nadu + Success) required. Current: state='${state}', paymentType='${paymentType}', paymentStatus='${paymentStatus}'`;
              }
            }

            if (!shouldSelectCarrier) {
              console.log(
                `Skipping carrier selection for order ${orderId}: ${skipReason}`
              );
              // Close the sync popup and return early
              await this.CloseSyncPopup(newPage);
              return {
                synced: false,
                reason: "carrier-conditions-not-met",
                skipReason,
              };
            }

            // Use the exact CARRIER_OVERRIDE value from .env to select from dropdown
            // Pass the pincode from the function parameters for validation
            selected = await trySelectByText(carrierOverride, pincode);
            if (selected) {
              selectedCarrier = carrierOverride;
            } else {
              // carrier not found in dropdown or pincode validation failed - log warning
              console.warn(
                `Carrier '${carrierOverride}' not selected. Either not found in dropdown or pincode validation failed.`
              );
            }
          }

          // If CARRIER_OVERRIDE is not set or selection failed, log warning
          if (!selected && !carrierOverride) {
            console.warn(
              "No CARRIER_OVERRIDE set in .env file. Please set CARRIER_OVERRIDE to 'DTDC' or 'Delhivery'."
            );
          } else if (!selected && carrierOverride) {
            console.warn(
              `Failed to select carrier '${carrierOverride}'. Please check if the carrier is available in the dropdown.`
            );
          }
        } catch (e) {
          // ignore selection failures
        }
      } catch (e) {
        // dropdown didn't appear or selection failed - continue to try next steps
      }

      // select radio #chk_lst_yes if present - use evaluate fallback to avoid hang
      try {
        const found = await newPage
          .waitForSelector("#chk_lst_yes", {
            state: "visible",
            timeout: 3000,
          })
          .catch(() => null);
        if (found) {
          // set checked via DOM and dispatch events
          await newPage.evaluate(() => {
            const el = document.querySelector("#chk_lst_yes");
            if (!el) return;
            try {
              el.checked = true;
              el.dispatchEvent(new Event("input", { bubbles: true }));
              el.dispatchEvent(new Event("change", { bubbles: true }));
            } catch (ee) {
              // ignore
            }
          });
        }
      } catch (e) {
        // ignore if radio not found
      }

      // wait a few seconds to allow any async popup process to run
      await newPage.waitForTimeout(waitMs);

      // Special handling for order 1599: perform logistic sync, fetch and save
      try {
        // normalize orderId for comparison
        const orderNumeric = Number(orderId);
        if (orderId !== "" || orderNumeric !== 0) {
          // 1) click on submit button with selector #logistic_sync
          try {
            // First wait for the button to be visible
            await newPage.waitForSelector("#logistic_sync", {
              state: "visible",
              timeout: 3000,
            });

            // Then wait for the button to become enabled (not disabled)
            await newPage.waitForFunction(
              () => {
                const btn = document.querySelector("#logistic_sync");
                return btn && !btn.disabled && !btn.hasAttribute("disabled");
              },
              { timeout: 5000 }
            );

            // accept any native confirm/alert dialog that may appear when submitting
            newPage.once("dialog", async (dialog) => {
              try {
                await dialog.accept();
              } catch (ee) {
                // ignore dialog accept failures
              }
            });
            await newPage.click("#logistic_sync");
          } catch (e) {
            // if logistic_sync not found or doesn't become enabled, continue - non-fatal
          }

          // 2) wait for the process to complete — detect modal close or wait a bit
          try {
            // Wait for any modal under #logisticsModal to disappear, or timeout
            await newPage.waitForSelector("#logisticsModal", {
              state: "detached",
              timeout: 8000,
            });
          } catch (e) {
            // fallback: short fixed wait to allow process to complete
            await newPage.waitForTimeout(2500);
          }

          // close popup by clicking #SyncClose if present
          await this.CloseSyncPopup(newPage);

          // Check if CARRIER_OVERRIDE is not "shiprocket" before executing fetch, GST, and save operations
          const carrierOverride = (process.env.CARRIER_OVERRIDE || "").trim();
          if (carrierOverride.toLowerCase() !== "shiprocket") {
            // 3) once the popup is closed, click on fetch button on the main page
            try {
              const fetchSel =
                "body > div.wrapper > div.content-wrapper > section > div.row > div > div.row.col-mb-4 > div:nth-child(3) > div:nth-child(1) > button";
              await newPage.waitForSelector(fetchSel, {
                state: "visible",
                timeout: 5000,
              });
              // some actions trigger a native confirmation dialog; accept it if shown
              newPage.once("dialog", async (dialog) => {
                try {
                  await dialog.accept();
                } catch (ee) {
                  // ignore
                }
              });
              await newPage.click(fetchSel);
              // wait for fetch to run
              await newPage.waitForTimeout(3000);
            } catch (e) {
              // fallback small wait if selector not found
              await newPage.waitForTimeout(1500);
            }

            // 4) generate GST invoice if required, then click on save with selector #save_order
            try {
              await newPage.waitForSelector("#save_order", {
                state: "visible",
                timeout: 5000,
              });
              // accept confirm/alert if the save triggers one
              newPage.once("dialog", async (dialog) => {
                try {
                  await dialog.accept();
                } catch (ee) {
                  // ignore
                }
              });
              await newPage.click("#save_order");
            } catch (e) {
              // if save button not found, ignore
            }

            // 5) wait for save to complete — look for save button to become disabled or just wait
            try {
              // wait a bit for save operation to complete
              await newPage.waitForTimeout(3000);
            } catch (e) {
              // noop
            }
          } else {
            console.log(
              "Skipping fetch, GST generation, and save operations for Shiprocket carrier"
            );
          }
        }
      } catch (e) {
        // don't let errors here break the main flow
      }

      // additional processing...

      // close popup by clicking #SyncClose if present
      await this.CloseSyncPopup(newPage);

      return { synced: true, carrier: selectedCarrier };
    } catch (e) {
      return { synced: false, reason: e.message, carrier: null };
    } finally {
      // ensure tab is closed
      try {
        await newPage.close();
      } catch (e) {
        // ignore
      }
    }
  }

  async CloseSyncPopup(newPage) {
    try {
      const closeSel = "#SyncClose";
      await newPage.waitForSelector(closeSel, {
        state: "visible",
        timeout: 3000,
      });
      await newPage.click(closeSel);
    } catch (e) {
      // fallback: try pressing Escape
      try {
        await newPage.keyboard.press("Escape");
      } catch (ee) {
        // ignore
      }
    }
  }

  // Wait for the order list table to be visible
  async waitForTable(timeout = 10000) {
    await this.page.waitForSelector(`${this.tableSelector} tbody tr`, {
      state: "visible",
      timeout,
    });
  }

  // Click each row's address button, wait for popup, then close it.
  // This method is defensive: it tries several close selectors and will
  // timeout gracefully per row instead of failing the whole run.
  async clickEachRowAddressPopup({ perRowTimeout = 5000 } = {}) {
    await this.waitForTable();
    // get all rows currently in the table
    const allRows = await this.page.$$(`${this.tableSelector} tbody tr`);

    // Filter rows to only include those with "New" badge
    const rows = [];
    for (const row of allRows) {
      try {
        // Check if this row has a "New" badge in the Name column (second column)
        const nameCell = await row.$("td:nth-child(2)");
        if (nameCell) {
          const badgeText = await nameCell
            .$eval(".badge", (el) => el.textContent.trim())
            .catch(() => null);
          if (badgeText === "New") {
            rows.push(row);
          }
        }
      } catch (e) {
        // If we can't check the badge, skip this row
        continue;
      }
    }

    // Log filtering results
    console.log(`Total rows in table: ${allRows.length}`);
    console.log(`Rows with "New" badge: ${rows.length}`);

    // If PROCESS_COUNT env var is set to a positive integer, treat it as
    // the maximum number of rows to process in this run. Otherwise process
    // all rows as before.
    const envCount = parseInt(process.env.PROCESS_COUNT, 10);
    const maxToProcess =
      Number.isInteger(envCount) && envCount > 0 ? envCount : null;
    if (maxToProcess) {
      // eslint-disable-next-line no-console
      console.log(
        `PROCESS_COUNT set: will stop after ${maxToProcess} successful records or end of table (whichever comes first)`
      );
    }
    // counter for how many rows we've successfully processed
    let successfullyProcessedCount = 0;
    let totalRowsAttempted = 0; // Track total rows attempted
    const processed = new Map(); // Use Map to store carrier -> orders dynamically
    const errors = []; // Array to store orders that had errors during processing
    for (let i = 0; i < rows.length; i++) {
      totalRowsAttempted = i + 1; // Update the count as we process each row
      // stop early if we've reached the PROCESS_COUNT limit for successful records
      // or if we've reached the end of all records
      if (maxToProcess && successfullyProcessedCount >= maxToProcess) {
        // eslint-disable-next-line no-console
        console.log(
          `Reached PROCESS_COUNT limit (${maxToProcess}) for successful records, stopping further processing.`
        );
        break;
      }
      const row = rows[i];
      try {
        // Attempt to extract an order id from the address button cell first
        // Selector pattern used by the UI: `#example > tbody > tr:nth-child(1) > td.sorting_1 > button.btn.btn-link.address-show-btn`
        let orderId = null;
        let paymentType = null;
        let paymentStatus = null;

        try {
          const addrBtn = await row.$(
            "td.sorting_1 > button.address-show-btn, td.sorting_1 > a.address-show-btn"
          );
          if (addrBtn) {
            // common attributes where an id might be stored
            const attrCandidates = [
              "data-order-id",
              "data-id",
              "data-order",
              "title",
              "aria-label",
            ];
            for (const attr of attrCandidates) {
              try {
                const v = await addrBtn.getAttribute(attr);
                if (v) {
                  orderId = v.trim();
                  break;
                }
              } catch (e) {
                // ignore attribute read errors
              }
            }

            if (!orderId) {
              try {
                const btnText = (await addrBtn.innerText()).trim();
                if (btnText) orderId = btnText;
              } catch (e) {
                // ignore
              }
            }
          }
        } catch (e) {
          // ignore
        }

        // If not found on the button, fallback to row attribute or common cells
        if (!orderId) {
          try {
            const dataAttr = await row.getAttribute("data-order-id");
            if (dataAttr) orderId = dataAttr.trim();
          } catch (e) {
            // ignore
          }
        }

        if (!orderId) {
          const orderCell = await row.$("td.order-id, th.order-id");
          if (orderCell) {
            try {
              orderId = (await orderCell.innerText()).trim();
            } catch (e) {
              // ignore
            }
          }
        }

        if (!orderId) {
          // fallback to first td text
          const firstTd = await row.$("td:first-child");
          if (firstTd) {
            try {
              orderId = (await firstTd.innerText()).trim();
            } catch (e) {
              // ignore
            }
          }
        }

        // Extract Payment Type and Payment Status from table cells
        // Based on the HTML structure: Payment Type is column 9 (index 8), Payment Status is column 10 (index 9)
        try {
          const allCells = await row.$$("td");
          if (allCells.length >= 10) {
            // Payment Type is at index 8 (9th column)
            try {
              paymentType = (await allCells[8].innerText()).trim();
            } catch (e) {
              // ignore
            }
            // Payment Status is at index 9 (10th column)
            try {
              paymentStatus = (await allCells[9].innerText()).trim();
            } catch (e) {
              // ignore
            }
          }
        } catch (e) {
          // ignore table cell extraction errors
        }
        // If ORDERS_TO_PROCESS is non-empty, only process rows whose
        // orderId is listed there. Otherwise process all orders.
        if (ORDERS_TO_PROCESS.size > 0) {
          const shouldProcess =
            orderId && ORDERS_TO_PROCESS.has(String(orderId).trim());
          if (!shouldProcess) {
            // eslint-disable-next-line no-console
            console.log(
              `Skipping order ${
                orderId || "N/A"
              } because it's not listed in ORDERS_TO_PROCESS`
            );
            continue;
          }
        }

        // find the button within the row using the relative selector
        const btn = await row.$(this.rowButtonSelector);
        if (!btn) {
          // try fallback: any button with address-show-btn within the row
          const fallback = await row.$(
            "button.address-show-btn, a.address-show-btn"
          );
          if (!fallback) {
            // nothing to click on this row
            continue;
          }
          await fallback.click();
        } else {
          await btn.click();
        }

        // wait a short while for popup to appear
        // Try to detect a modal or popup by waiting for either a .modal element
        // or an element that wasn't present before. We'll wait for a short fixed delay
        // then attempt to close using known selectors.
        await this.page.waitForTimeout(1000); // give popup a chance to appear

        // If a centralized address popup handler exists (from LoginPage), use it.
        // This delegates to LoginPage.handleAddressPopup which will inspect and close
        // the #addressShowBody popup if present. It's safe to call and will return
        // quickly if the popup doesn't exist.
        let handleResult = null;
        try {
          // pass 1-based row index and orderId for clearer logs
          handleResult = await this.handleAddressPopup(i + 1, orderId);
        } catch (e) {
          // ignore errors from the delegated handler and continue with local logic
        }

        // Track whether this row was successfully processed
        let rowProcessedSuccessfully = false;

        // If we extracted a pincode and have an orderId, attempt to sync via Shiprocket in a new tab.
        try {
          const pincode = handleResult && handleResult.pincode;
          const state = handleResult && handleResult.state;
          if (pincode && orderId) {
            // run sync flow for this order; keep it quick and non-blocking per row
            // awaiting here ensures sequential per-row behavior; if you want parallel,
            // you could spawn without await but ensure resource limits.
            const result = await this.syncShiprocketForOrder(orderId, {
              waitMs: 2500,
              pincode,
              state,
              paymentType,
              paymentStatus,
            });

            // Check if sync was successful
            if (result && result.synced) {
              try {
                // Always use CARRIER_OVERRIDE as the definitive carrier name
                // This ensures consistency regardless of what the sync operation returns
                const carrierOverride = (
                  process.env.CARRIER_OVERRIDE || ""
                ).trim();
                let carrier = carrierOverride || result.carrier;

                if (carrier) {
                  if (!processed.has(carrier)) {
                    processed.set(carrier, []);
                  }
                  processed.get(carrier).push({
                    orderId,
                    pincode,
                    state: (handleResult && handleResult.state) || "N/A",
                    paymentType: paymentType || "N/A",
                    paymentStatus: paymentStatus || "N/A",
                  });
                  // Mark as successfully processed since we have a carrier and sync succeeded
                  rowProcessedSuccessfully = true;
                } else {
                  // No carrier identified but sync was successful - add to errors for investigation
                  errors.push({
                    orderId,
                    pincode,
                    error: "Sync successful but no carrier identified",
                  });
                  // Don't mark as successful since no carrier was identified
                }
              } catch (e) {
                // Error processing successful sync result
                errors.push({
                  orderId,
                  pincode,
                  error: `Error processing sync result: ${e.message}`,
                });
                // Don't mark as successful due to error
              }
            } else {
              // Sync failed - check if it was due to dialog appearance
              const errorReason = result
                ? result.reason
                : "Unknown sync failure";
              const skipReason =
                result && result.skipReason ? ` - ${result.skipReason}` : "";

              // Special handling for dialog-appeared case
              if (result && result.reason === "dialog-appeared") {
                console.log(
                  `Row ${
                    i + 1
                  } (Order ID: ${orderId}) - Skipped due to browser dialog. Pincode: ${pincode}, State: ${
                    state || "N/A"
                  }, Payment Type: ${paymentType || "N/A"}, Payment Status: ${
                    paymentStatus || "N/A"
                  }`
                );
                errors.push({
                  orderId,
                  pincode,
                  error: `Skipped due to browser dialog appearance`,
                  state: state || "N/A",
                  paymentType: paymentType || "N/A",
                  paymentStatus: paymentStatus || "N/A",
                });
              } else {
                errors.push({
                  orderId,
                  pincode,
                  error: `Sync failed: ${errorReason}${skipReason}`,
                });
              }
              // Don't mark as successful since sync failed
            }
          } else {
            // Missing pincode or orderId - add to errors
            errors.push({
              orderId: orderId || "Unknown",
              pincode: (handleResult && handleResult.pincode) || "Unknown",
              error: "Missing pincode or order ID for processing",
            });
          }
        } catch (e) {
          // Exception during sync attempt - add to errors
          errors.push({
            orderId: orderId || "Unknown",
            pincode: (handleResult && handleResult.pincode) || "Unknown",
            error: `Exception during sync: ${e.message}`,
          });
          // eslint-disable-next-line no-console
          console.warn(
            `row ${i + 1}: error during syncShiprocket - ${e.message}`
          );
        }

        // Per-row logging so user sees immediate progress for each processed row
        try {
          const pcode = (handleResult && handleResult.pincode) || null;
          const state = (handleResult && handleResult.state) || null;
          // determine carrier if we recorded it in processed arrays
          let carrier = null;
          for (const [carrierName, orders] of processed) {
            if (orders.find((x) => x.orderId === orderId)) {
              carrier = carrierName;
              break;
            }
          }
          // eslint-disable-next-line no-console
          console.log(
            `Row ${i + 1}: order=${orderId || "N/A"}, pincode=${
              pcode || "N/A"
            }, state=${state || "N/A"}, paymentType=${
              paymentType || "N/A"
            }, paymentStatus=${paymentStatus || "N/A"}, carrier=${
              carrier || "N/A"
            }`
          );
        } catch (e) {
          // ignore logging errors per row
        }

        // Only increment counter for successfully processed rows
        if (rowProcessedSuccessfully) {
          try {
            successfullyProcessedCount += 1;
          } catch (e) {
            // ignore
          }

          // Check if we've reached the configured maximum successful records
          if (maxToProcess && successfullyProcessedCount >= maxToProcess) {
            // eslint-disable-next-line no-console
            console.log(
              `Reached PROCESS_COUNT limit (${maxToProcess}) for successful records after row ${
                i + 1
              }. Stopping processing.`
            );
            break;
          }
        }

        // short pause before next row to stabilize DOM

        await this.page.waitForTimeout(200);
      } catch (e) {
        // continue to next row; do not fail the whole loop
        // but log to console for debugging
        // eslint-disable-next-line no-console
        console.warn(`row ${i + 1}: error handling popup - ${e.message}`);
      }
    }
    // After processing all rows, print a summary and write it to logs
    try {
      // Calculate totals
      let totalSuccessful = 0;
      for (const ordersList of processed.values()) {
        totalSuccessful += ordersList.length;
      } // Print processing summary
      console.log(`\nPROCESSING SUMMARY`);
      console.log("================================");
      console.log(`Total rows attempted: ${totalRowsAttempted}`);
      console.log(`Successfully processed: ${totalSuccessful}`);
      console.log(`Errors/Skipped: ${errors.length}`);
      if (maxToProcess) {
        console.log(`PROCESS_COUNT limit: ${maxToProcess}`);
        console.log(
          `Limit reached: ${totalSuccessful >= maxToProcess ? "YES" : "NO"}`
        );
      }

      // Print consolidated summary of all successful orders
      if (totalSuccessful > 0) {
        console.log(`\nSUCCESSFUL ORDERS SUMMARY (${totalSuccessful})`);
        console.log("============================================");
        let orderIndex = 1;
        for (const [carrierName, ordersList] of processed) {
          for (const item of ordersList) {
            console.log(
              `${orderIndex}. Order: ${item.orderId} | Pincode: ${item.pincode} | State: ${item.state} | Payment: ${item.paymentType} | Status: ${item.paymentStatus} | Carrier: ${carrierName}`
            );
            orderIndex++;
          }
        }
      }

      // Print summary for each carrier dynamically (successful processing only)
      for (const [carrierName, ordersList] of processed) {
        console.log(`\nProcessed on ${carrierName} (${ordersList.length})`);
        console.log("--------------------------------");
        for (let i = 0; i < ordersList.length; i++) {
          const item = ordersList[i];
          console.log(
            `${i + 1}. Order: ${item.orderId}, Pincode: ${
              item.pincode
            }, State: ${item.state}, Payment Type: ${
              item.paymentType
            }, Payment Status: ${item.paymentStatus}`
          );
        }
      }

      // Print errors section if there are any
      if (errors.length > 0) {
        // Separate dialog-skipped rows from other errors
        const dialogSkipped = errors.filter((e) =>
          e.error.includes("browser dialog")
        );
        const otherErrors = errors.filter(
          (e) => !e.error.includes("browser dialog")
        );

        if (dialogSkipped.length > 0) {
          console.log(
            `\nSkipped due to Browser Dialog (${dialogSkipped.length})`
          );
          console.log("--------------------------------------------------");
          for (let i = 0; i < dialogSkipped.length; i++) {
            const item = dialogSkipped[i];
            console.log(
              `${i + 1}. Order: ${item.orderId}, Pincode: ${
                item.pincode
              }, State: ${item.state || "N/A"}, Payment Type: ${
                item.paymentType || "N/A"
              }, Payment Status: ${item.paymentStatus || "N/A"}`
            );
          }
        }

        if (otherErrors.length > 0) {
          console.log(`\nProcessing Errors (${otherErrors.length})`);
          console.log("--------------------------------");
          for (let i = 0; i < otherErrors.length; i++) {
            const item = otherErrors[i];
            console.log(
              `${i + 1}. Order: ${item.orderId}, Pincode: ${
                item.pincode
              }, Error: ${item.error}`
            );
          }
        }
      }

      // write to logs directory
      try {
        const logsDir = path.join(process.cwd(), "logs");
        if (!fs.existsSync(logsDir)) fs.mkdirSync(logsDir, { recursive: true });
        const ts = new Date().toISOString().replace(/[:.]/g, "-");
        const filename = path.join(logsDir, `summary-${ts}.txt`);
        const lines = [];

        // Calculate totals for log file
        let totalSuccessful = 0;
        for (const ordersList of processed.values()) {
          totalSuccessful += ordersList.length;
        }

        // Write processing summary to log
        lines.push("PROCESSING SUMMARY");
        lines.push("================================");
        lines.push(`Total rows attempted: ${totalRowsAttempted}`);
        lines.push(`Successfully processed: ${totalSuccessful}`);
        lines.push(`Errors/Skipped: ${errors.length}`);
        if (maxToProcess) {
          lines.push(`PROCESS_COUNT limit: ${maxToProcess}`);
          lines.push(
            `Limit reached: ${totalSuccessful >= maxToProcess ? "YES" : "NO"}`
          );
        }
        lines.push(""); // Add empty line

        // Write consolidated summary of all successful orders to log
        if (totalSuccessful > 0) {
          lines.push(`SUCCESSFUL ORDERS SUMMARY (${totalSuccessful})`);
          lines.push("============================================");
          let orderIndex = 1;
          for (const [carrierName, ordersList] of processed) {
            for (const item of ordersList) {
              lines.push(
                `${orderIndex}. Order: ${item.orderId} | Pincode: ${item.pincode} | State: ${item.state} | Payment: ${item.paymentType} | Status: ${item.paymentStatus} | Carrier: ${carrierName}`
              );
              orderIndex++;
            }
          }
          lines.push(""); // Add empty line
        }

        // Write summary for each carrier dynamically (successful processing only)
        for (const [carrierName, ordersList] of processed) {
          lines.push(`Processed on ${carrierName} (${ordersList.length})`);
          lines.push("--------------------------------");
          for (let i = 0; i < ordersList.length; i++) {
            const item = ordersList[i];
            lines.push(
              `${i + 1}. Order: ${item.orderId}, Pincode: ${
                item.pincode
              }, State: ${item.state}, Payment Type: ${
                item.paymentType
              }, Payment Status: ${item.paymentStatus}`
            );
          }
          lines.push(""); // Add empty line between carriers
        }

        // Write errors section if there are any
        if (errors.length > 0) {
          // Separate dialog-skipped rows from other errors in log file too
          const dialogSkipped = errors.filter((e) =>
            e.error.includes("browser dialog")
          );
          const otherErrors = errors.filter(
            (e) => !e.error.includes("browser dialog")
          );

          if (dialogSkipped.length > 0) {
            lines.push(
              `Skipped due to Browser Dialog (${dialogSkipped.length})`
            );
            lines.push("--------------------------------------------------");
            for (let i = 0; i < dialogSkipped.length; i++) {
              const item = dialogSkipped[i];
              lines.push(
                `${i + 1}. Order: ${item.orderId}, Pincode: ${
                  item.pincode
                }, State: ${item.state || "N/A"}, Payment Type: ${
                  item.paymentType || "N/A"
                }, Payment Status: ${item.paymentStatus || "N/A"}`
              );
            }
            lines.push(""); // Add empty line
          }

          if (otherErrors.length > 0) {
            lines.push(`Processing Errors (${otherErrors.length})`);
            lines.push("--------------------------------");
            for (let i = 0; i < otherErrors.length; i++) {
              const item = otherErrors[i];
              lines.push(
                `${i + 1}. Order: ${item.orderId}, Pincode: ${
                  item.pincode
                }, Error: ${item.error}`
              );
            }
            lines.push(""); // Add empty line after errors
          }
        }

        fs.writeFileSync(filename, lines.join("\n"));
        console.log(`Summary written to ${filename}`);
      } catch (e) {
        // ignore file write errors
      }
    } catch (e) {
      // Log the error instead of ignoring it completely
      console.error("Error in summary generation:", e.message);
      console.error("Stack trace:", e.stack);
    }
  }
}

module.exports = {
  OrderListPage,
  extractPincode,
  extractState,
  loadExcelCacheForCarrier,
  getCarrierPincodes,
  loadExcelCaches,
};
