const puppeteer = require('puppeteer');

(async () => {
  const browser = await puppeteer.launch({
    headless: 'new', // Opt into the new headless mode
  });
  const page = await browser.newPage();

  await page.goto('https://prod-vnv-inquisito-ui.dexcomdev.com');
  await page.screenshot({ path: 'example.png' });

  await browser.close();
})();