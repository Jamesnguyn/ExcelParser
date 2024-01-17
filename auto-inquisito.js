const puppeteer = require('puppeteer');

(async () => {
  const browser = await puppeteer.launch({
    headless: 'new', // Opt into the new headless mode
  });
  const page = await browser.newPage();

  // Set a custom user agent
  await page.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36');

  await page.goto('https://prod-vnv-inquisito-ui.dexcomdev.com');

  await page.waitForSelector('.login-button-container');
  await page.click('.login-button-container');

  // Wait for the username input field to appear
  await page.waitForSelector('input[name="identifier"]');
  // Type into the username input field
  await page.type('input[name="identifier"]', 'test@dexcom.com');

  // Wait for the password input field to appear
  await page.waitForSelector('input[name="credentials.passcode"]');
  // Type into the password input field
  await page.type('input[name="credentials.passcode"]', 'Hello, Puppeteer!');

  await page.click('span.eyeicon.visibility-16.button-show');

  // Wait for the submit button to appear
  await page.waitForSelector('input.button.button-primary[type="submit"][value="Sign in"][data-type="save"]');

  await page.evaluate(() => {
    document.querySelector('input.button.button-primary[type="submit"][value="Sign in"][data-type="save"]').click();
  });

  // await page.click('input.button.button-primary[type="submit"][value="Sign in"][data-type="save"]');

  new Promise(r => setTimeout(r, 3000));

  await page.screenshot({ path: 'example.png' });

  await browser.close();
})();