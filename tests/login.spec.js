require("dotenv").config();
const { test, expect } = require("@playwright/test");
const { LoginPage } = require("./pages/loginPage");
const { OrderListPage } = require("./pages/orderListPage");

test("navigate to diyaa.in/login and type credentials using page object", async ({
  page,
}) => {
  test.setTimeout(0);
  const login = new LoginPage(page);
  await login.goto();
  await expect(page).toHaveURL(/.*\/login/);

  const username = process.env.LOGIN_USERNAME;
  const password = process.env.LOGIN_PASSWORD;

  if (!username || !password) {
    throw new Error(
      "Missing LOGIN_USERNAME and/or LOGIN_PASSWORD in environment. Add them to a .env file or set env vars."
    );
  }

  await login.fillUsername(username, { delay: 120 });
  await login.fillPassword(password, { delay: 120 });
  await login.submit();
  // wait for master page to load after successful login
  await page.waitForURL(/.*\/master/, { timeout: 0 });

  // open Orders -> Order List from the left sidebar and verify navigation
  await login.selectOrderList();
  await page.waitForURL(/.*\/inventory\/order_list/, { timeout: 0 });
  await expect(page).toHaveURL(/.*\/inventory\/order_list/);
  // basic visibility check for the Order List page - check the page header
  await expect(page.locator('h3.box-title:has-text("Order List")')).toBeVisible(
    { timeout: 5000 }
  );
  // select "All" from the page length dropdown to show all items (value -1)
  const pageLength = page.locator('select[name="example_length"]');
  await expect(pageLength).toBeVisible({ timeout: 5000 });
  await pageLength.selectOption({ value: "-1" });
  // verify the select has the All value
  await expect(pageLength).toHaveValue("-1");
  // verify the table has at least one row (tbody tr)
  const rows = page.locator("table#example tbody tr");
  await expect(rows.first()).toBeVisible({ timeout: 5000 });
  // create OrderListPage and click each row's address button to open/close popups
  const orders = new OrderListPage(page);
  await orders.clickEachRowAddressPopup({ perRowTimeout: 3000 });
});
