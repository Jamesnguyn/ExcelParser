const puppeteer = require('puppeteer');

(async () => {
  const browser = await puppeteer.launch({
    headless: 'new', // Opt into the new headless mode
  });
  const page = await browser.newPage();

  // Set a custom user agent
  await page.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36');

  await page.goto('https://prod-vnv-inquisito-ui.dexcomdev.com');
  // await page.goto('https://youtube.com');

  // Wait for navigation to complete
  await page.waitForNavigation();

  await page.screenshot({ path: 'example.png' });

  await browser.close();
})();