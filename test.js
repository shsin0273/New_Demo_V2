const { chromium: chromiumExtra } = require('playwright-extra');
const StealthPlugin = require('puppeteer-extra-plugin-stealth');
chromiumExtra.use(StealthPlugin());

(async () => {
  const browser = await chromiumExtra.launch({
    headless: false,
    slowMo: 300,
    args: [
      '--no-sandbox',
      '--disable-setuid-sandbox',
      '--disable-blink-features=AutomationControlled',
      '--window-size=390,844',
    ],
    proxy: {
      server: 'http://ca.proxy-jet.io:1010',
      username: '241120H9lZQ-resi-US',
      password: 'w3ILawmH11RrtYV'
    }
  });

  const pool = [
    { ua: 'Mozilla/5.0 (Linux; Android 14; SAMSUNG SM-S918B) AppleWebKit/537.36 (KHTML, like Gecko) SamsungBrowser/24.0 Chrome/117.0.0.0 Mobile Safari/537.36', w: 2.0 },
    { ua: 'Mozilla/5.0 (Linux; Android 14; Pixel 8) AppleWebKit/537.36 (KHTML, like Gecko) SamsungBrowser/25.0 Chrome/121.0.0.0 Mobile Safari/537.36', w: 2.0 },
    { ua: 'Mozilla/5.0 (Linux; Android 13; SM-A546B) AppleWebKit/537.36 (KHTML, like Gecko) SamsungBrowser/23.0 Chrome/115.0.0.0 Mobile Safari/537.36', w: 1.5 },
    { ua: 'Mozilla/5.0 (Linux; Android 14; Pixel 8 Pro) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Mobile Safari/537.36', w: 1.5 },
    { ua: 'Mozilla/5.0 (Linux; Android 13; SM-G991B) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Mobile Safari/537.36', w: 1.0 },
  ];
  const total = pool.reduce((s, x) => s + x.w, 0);
  let r = Math.random() * total;
  let selectedUA = pool[pool.length - 1].ua;
  for (const x of pool) { r -= x.w; if (r <= 0) { selectedUA = x.ua; break; } }

  console.log('Selected UA:', selectedUA);

  const context = await browser.newContext({
    userAgent: selectedUA,
    viewport: { width: 390, height: 844 },
    deviceScaleFactor: 3,
    isMobile: true,
    hasTouch: true,
    locale: 'en-US',
    timezoneId: 'America/New_York',
  });

  const page = await context.newPage();

  await context.setExtraHTTPHeaders({
  'Sec-CH-UA': '"Samsung Browser";v="24", "Not-A.Brand";v="99"',
  'Sec-CH-UA-Mobile': '?1',
  'Sec-CH-UA-Platform': '"Android"',
  'Sec-CH-UA-Platform-Version': '"14.0"',
});

  await context.addInitScript(() => {
    Object.defineProperty(navigator, 'platform', { get: () => 'Linux armv8l' });
    Object.defineProperty(navigator, 'maxTouchPoints', { get: () => 5 });
  });

  console.log('Opening tracking URL...');

  async function openWithRetry(retries = 3) {
    for (let i = 0; i < retries; i++) {
      try {
        await page.goto('https://offers.track2all.xyz/c?s=nGh7NWZg', {
          waitUntil: 'domcontentloaded',
          timeout: 60000
        });

        await page.waitForTimeout(8000);

        const finalUrl = page.url();
        console.log(`Attempt ${i + 1}:`, finalUrl);

        if (!finalUrl.includes('chrome-error')) {
          return true;
        }

      } catch (err) {
        console.log('Error:', err.message);
      }

      console.log('Retrying...');
      await page.waitForTimeout(5000);
    }

    return false;
  }

  await openWithRetry();

  console.log('Final URL:', page.url());
  console.log('Browser will stay open. Close it manually.');

  await new Promise(() => {});
})();