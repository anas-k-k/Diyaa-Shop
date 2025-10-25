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
const _excelCache = {
  DTDC: null, // Set or null
  Delhivery: null,
  lastLoaded: 0,
  // reload interval in ms (optional) - set to 60s to allow occasional refresh
  reloadInterval: 60 * 1000,
};

function loadExcelCaches() {
  const now = Date.now();
  if (
    _excelCache.lastLoaded &&
    now - _excelCache.lastLoaded < _excelCache.reloadInterval &&
    _excelCache.DTDC !== null &&
    _excelCache.Delhivery !== null
  ) {
    return;
  }
  const dataDir = path.join(process.cwd(), "data");
  _excelCache.DTDC = readPincodesFromExcel(path.join(dataDir, "DTDC.xlsx"));
  _excelCache.Delhivery = readPincodesFromExcel(
    path.join(dataDir, "Delhivery.xlsx")
  );
  // Delhivery cache is a Set; use .add to insert a value (avoid .push which is for arrays)
  try {
    if (
      _excelCache.Delhivery &&
      typeof _excelCache.Delhivery.add === "function"
    ) {
      // _excelCache.Delhivery.add("682021"); // manually add known pincode
    }
  } catch (e) {
    // ignore any errors when adding to cache
  }
  _excelCache.lastLoaded = now;
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

    // try to extract pincode(s) from the raw text and log the main one
    const pincodes = extractPincode(rawText);
    const pincode = pincodes.length ? pincodes[0] : null;

    // eslint-disable-next-line no-console
    console.log(
      `Extracted pincode from address popup: ${pincode} (orderId=${
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
      rawText,
      orderId,
    };
  }

  // Open a new tab for the order sync page, click Sync with Shiprocket,
  // interact with the modal (select courier DTDC, choose radio, wait), then close.
  // This method is defensive and will return quickly if elements are not found.
  async syncShiprocketForOrder(orderId, { waitMs = 2500 } = {}) {
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
        await newPage.click("#sync_shiprocket");
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
          const trySelectByText = async (text) => {
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
            // Use the exact CARRIER_OVERRIDE value from .env to select from dropdown
            selected = await trySelectByText(carrierOverride);
            if (selected) {
              selectedCarrier = carrierOverride;
            } else {
              // carrier not found in dropdown - log warning
              console.warn(
                `Carrier '${carrierOverride}' not found in dropdown. No carrier selected.`
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
            // Click the "Generate" GST button for all carriers if available
            try {
              // click the generate GST button if visible
              const genSel = "#gen_gst_invoice";
              const gstNbSel = "#gst_invoice_nb";
              const genEl = await newPage
                .waitForSelector(genSel, {
                  state: "visible",
                  timeout: 3000,
                })
                .catch(() => null);
              if (genEl) {
                // Clicking may trigger a native browser dialog (confirm/alert).
                // Accept it explicitly to match manual behaviour seen in the UI.
                newPage.once("dialog", async (dialog) => {
                  try {
                    await dialog.accept();
                  } catch (ee) {
                    // ignore accept failures
                  }
                });
                console.log(
                  `GST Invoice Generation Triggered for carrier: ${
                    selectedCarrier || "Unknown"
                  }`
                );
                // clicking may trigger an async process that sets the value of #gst_invoice_nb
                await genEl.click().catch(() => {});

                // wait for #gst_invoice_nb to have a non-empty value (up to 8s)
                const start = Date.now();
                const timeout = 8000;
                let populated = false;
                while (Date.now() - start < timeout) {
                  try {
                    const val = await newPage.evaluate((sel) => {
                      const e = document.querySelector(sel);
                      return e ? e.value || e.innerText || "" : "";
                    }, gstNbSel);
                    if (val && String(val).trim().length) {
                      populated = true;
                      break;
                    }
                  } catch (ee) {
                    // ignore evaluation errors and retry
                  }
                  // short sleep
                  await newPage.waitForTimeout(250);
                }
                // if not populated, continue anyway - non-fatal
              }
            } catch (e) {
              // ignore failures in GST generation; proceed to save
            }

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
    const rows = await this.page.$$(`${this.tableSelector} tbody tr`);
    // If PROCESS_COUNT env var is set to a positive integer, treat it as
    // the maximum number of rows to process in this run. Otherwise process
    // all rows as before.
    const envCount = parseInt(process.env.PROCESS_COUNT, 10);
    const maxToProcess =
      Number.isInteger(envCount) && envCount > 0 ? envCount : null;
    if (maxToProcess) {
      // eslint-disable-next-line no-console
      console.log(
        `PROCESS_COUNT set: will process at most ${maxToProcess} rows`
      );
    }
    // counter for how many rows we've actually processed (click + handle)
    let processedRowCount = 0;
    const processed = new Map(); // Use Map to store carrier -> orders dynamically
    const errors = []; // Array to store orders that had errors during processing
    for (let i = 0; i < rows.length; i++) {
      // stop early if we've reached the PROCESS_COUNT limit
      if (maxToProcess && processedRowCount >= maxToProcess) {
        // eslint-disable-next-line no-console
        console.log(
          `Reached PROCESS_COUNT limit (${maxToProcess}), stopping further processing.`
        );
        break;
      }
      const row = rows[i];
      try {
        // Attempt to extract an order id from the address button cell first
        // Selector pattern used by the UI: `#example > tbody > tr:nth-child(1) > td.sorting_1 > button.btn.btn-link.address-show-btn`
        let orderId = null;
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

        // If we extracted a pincode and have an orderId, attempt to sync via Shiprocket in a new tab.
        try {
          const pincode = handleResult && handleResult.pincode;
          if (pincode && orderId) {
            // run sync flow for this order; keep it quick and non-blocking per row
            // awaiting here ensures sequential per-row behavior; if you want parallel,
            // you could spawn without await but ensure resource limits.
            const result = await this.syncShiprocketForOrder(orderId, {
              waitMs: 2500,
              pincode,
            });

            // Check if sync was successful
            if (result && result.synced) {
              try {
                // Prefer the carrier returned by the sync flow if present
                let carrier = result.carrier;

                // If sync didn't return a carrier, use the CARRIER_OVERRIDE value directly
                if (!carrier) {
                  const carrierOverride = (
                    process.env.CARRIER_OVERRIDE || ""
                  ).trim();
                  if (carrierOverride) {
                    carrier = carrierOverride; // Use the exact value from .env
                  }
                }

                if (carrier) {
                  if (!processed.has(carrier)) {
                    processed.set(carrier, []);
                  }
                  processed.get(carrier).push({ orderId, pincode });
                } else {
                  // No carrier identified but sync was successful - add to errors for investigation
                  errors.push({
                    orderId,
                    pincode,
                    error: "Sync successful but no carrier identified",
                  });
                }
              } catch (e) {
                // Error processing successful sync result
                errors.push({
                  orderId,
                  pincode,
                  error: `Error processing sync result: ${e.message}`,
                });
              }
            } else {
              // Sync failed - add to errors
              const errorReason = result
                ? result.reason
                : "Unknown sync failure";
              errors.push({
                orderId,
                pincode,
                error: `Sync failed: ${errorReason}`,
              });
            }
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
            }, carrier=${carrier || "N/A"}`
          );
        } catch (e) {
          // ignore logging errors per row
        }

        // short pause before next row to stabilize DOM
        // increment processed counter only for rows that actually reached
        // this point (i.e. were clicked and handled)
        try {
          processedRowCount += 1;
        } catch (e) {
          // ignore
        }

        // if we've reached the configured maximum, break out early
        if (maxToProcess && processedRowCount >= maxToProcess) break;

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
      // Print summary for each carrier dynamically (successful processing only)
      for (const [carrierName, ordersList] of processed) {
        console.log(`\nProcessed on ${carrierName} (${ordersList.length})`);
        console.log("--------------------------------");
        for (let i = 0; i < ordersList.length; i++) {
          const item = ordersList[i];
          console.log(
            `${i + 1}. Order :${item.orderId}, Pincode: ${item.pincode}`
          );
        }
      }

      // Print errors section if there are any
      if (errors.length > 0) {
        console.log(`\nProcessing Errors (${errors.length})`);
        console.log("--------------------------------");
        for (let i = 0; i < errors.length; i++) {
          const item = errors[i];
          console.log(
            `${i + 1}. Order :${item.orderId}, Pincode: ${
              item.pincode
            }, Error: ${item.error}`
          );
        }
      }

      // write to logs directory
      try {
        const logsDir = path.join(process.cwd(), "logs");
        if (!fs.existsSync(logsDir)) fs.mkdirSync(logsDir, { recursive: true });
        const ts = new Date().toISOString().replace(/[:.]/g, "-");
        const filename = path.join(logsDir, `summary-${ts}.txt`);
        const lines = [];

        // Write summary for each carrier dynamically (successful processing only)
        for (const [carrierName, ordersList] of processed) {
          lines.push(`Processed on ${carrierName} (${ordersList.length})`);
          lines.push("--------------------------------");
          for (let i = 0; i < ordersList.length; i++) {
            const item = ordersList[i];
            lines.push(
              `${i + 1}. Order :${item.orderId}, Pincode: ${item.pincode}`
            );
          }
          lines.push(""); // Add empty line between carriers
        }

        // Write errors section if there are any
        if (errors.length > 0) {
          lines.push(`Processing Errors (${errors.length})`);
          lines.push("--------------------------------");
          for (let i = 0; i < errors.length; i++) {
            const item = errors[i];
            lines.push(
              `${i + 1}. Order :${item.orderId}, Pincode: ${
                item.pincode
              }, Error: ${item.error}`
            );
          }
          lines.push(""); // Add empty line after errors
        }

        fs.writeFileSync(filename, lines.join("\n"));
        console.log(`Summary written to ${filename}`);
      } catch (e) {
        // ignore file write errors
      }
    } catch (e) {
      // ignore logging errors
    }
  }
}

module.exports = { OrderListPage, extractPincode };
