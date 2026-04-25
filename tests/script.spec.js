import { test, expect } from '@playwright/test';
import excel from 'exceljs';

test('script', async ({ browser }) => {

    const context = await browser.newContext({
    permissions: ['geolocation']
    });
    const page = await context.newPage();

    await page.goto('https://www.flipkart.com');

    await page.locator('[role="button"]').click();
    await page.locator('[src="https://static-assets-web.flixcart.com/apex-static/images/svgs/L1Nav/home-final.svg"]').click();
    await page.locator('[style="filter: none; opacity: 1; transition: filter 0.5s ease-in-out, opacity 0.5s ease-in-out; width: 100%; margin: auto; display: block; object-fit: contain; aspect-ratio: 45 / 64;"]').first().click();
    await page.waitForTimeout(1000);
    await page.locator('//*[@title="4★ & above"]/div/label/div[@class="ybaCDx"]').click();
    await page.waitForTimeout(1000);
    await page.getByText('Price -- Low to High').click();

    await Promise.all([
        context.waitForEvent('page'),
        page.locator('[class="lWX0_T"]').nth(4).click()
    ])

    const pages = context.pages();
    const page2 = pages[1];
    
    const title = await page2.locator('h1[class="v1zwn21l v1zwn26 _1psv1zeb9 _1psv1ze0"]').textContent();
    const price = await page2.locator('//div[@class="v1zwn21l v1zwn20 _1psv1zeb9 _1psv1ze0"]').textContent();

    await page2.getByText('Add to cart').click();
    await page2.screenshot({path: 'screenshots/screenshot.png'});

    expect(title).not.toBeNull();
    expect(title?.trim().length).toBeGreaterThan(0);
    expect(price).not.toBeNull();
    expect(price).toMatch(/₹\s?\d+/);

    let book = new excel.Workbook();
    await book.xlsx.readFile("C:/Users/arush/playwright-assesments/25-04-2026/Excel Sheet/products.xlsx");

    let sheet = await book.getWorksheet("Sheet2");

    if (!sheet) {
        sheet = await book.addWorksheet("Sheet2");
    }

    sheet.addRow([title]);
    sheet.addRow([price]);

    await book.xlsx.writeFile("C:/Users/arush/playwright-assesments/25-04-2026/Excel Sheet/products.xlsx");
});