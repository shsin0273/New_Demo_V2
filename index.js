const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const csv = require('csv-parser');
const createCsvWriter = require('csv-writer').createObjectCsvWriter;

const file = path.join(__dirname, 'FiveK_Demo_leads_Unclaim.xlsx');
const leadsFile = path.join(__dirname, 'leads.csv');
const outDir = path.join(__dirname, 'output');
if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });

const url = 'https://track.stats36245.xyz/c?s=1M2bXZal';
const MAX_ATTEMPTS = 5;
const PAGE_TIMEOUT = 15000;

function pick(v) {
  return v == null ? '' : String(v).trim();
}

function normalizePhone(v) {
  const d = pick(v).replace(/\D/g, '').slice(0, 10);
  if (d.length < 10) return d;
  return `${d.slice(0, 3)}-${d.slice(3, 6)}-${d.slice(6, 10)}`;
}

function normalizeZip(v) {
  return pick(v).replace(/\D/g, '').slice(0, 5);
}

function parseDOB(v) {
  if (v == null || v === '') return { month: '', day: '', year: '' };

  if (v instanceof Date && !isNaN(v)) {
    return {
      month: String(v.getMonth() + 1),
      day: String(v.getDate()),
      year: String(v.getFullYear())
    };
  }

  if (typeof v === 'number') {
    const excelEpoch = new Date(Date.UTC(1899, 11, 30));
    const d = new Date(excelEpoch.getTime() + v * 86400000);
    return {
      month: String(d.getUTCMonth() + 1),
      day: String(d.getUTCDate()),
      year: String(d.getUTCFullYear())
    };
  }

  const s = String(v).trim();

  let m = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})$/);
  if (m) return { month: m[2], day: m[3], year: m[1] };

  m = s.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})$/);
  if (m) return { month: m[1], day: m[2], year: m[3] };

  const d = new Date(s);
  if (!isNaN(d)) {
    return {
      month: String(d.getMonth() + 1),
      day: String(d.getDate()),
      year: String(d.getFullYear())
    };
  }

  return { month: '', day: '', year: '' };
}

function sleep(ms) {
  return new Promise(r => setTimeout(r, ms));
}

function randomItem(arr) {
  if (!arr || !arr.length) return null;
  return arr[Math.floor(Math.random() * arr.length)];
}

function ensureLeadsCsvHeaders() {
  if (!fs.existsSync(leadsFile)) {
    const headers = ['FIRST', 'LAST', 'EMAIL', 'PHONE1', 'ZIP', 'ADDRESS', 'DOB', 'PROCESSED_AT'];
    fs.writeFileSync(leadsFile, headers.join(',') + '\n');
  }
}

async function loadExistingLeads() {
  const existingLeads = new Set();
  if (!fs.existsSync(leadsFile)) return existingLeads;

  return new Promise((resolve, reject) => {
    const leads = [];
    fs.createReadStream(leadsFile)
      .pipe(csv())
      .on('data', (row) => {
        const key = `${pick(row.FIRST)}|${pick(row.LAST)}|${pick(row.EMAIL)}|${pick(row.PHONE1)}`.toLowerCase().trim();
        if (key) leads.push(key);
      })
      .on('end', () => {
        leads.forEach(key => existingLeads.add(key));
        resolve(existingLeads);
      })
      .on('error', reject);
  });
}

async function addProcessedLead(row) {
  const csvWriter = createCsvWriter({
    path: leadsFile,
    header: [
      { id: 'FIRST', title: 'FIRST' },
      { id: 'LAST', title: 'LAST' },
      { id: 'EMAIL', title: 'EMAIL' },
      { id: 'PHONE1', title: 'PHONE1' },
      { id: 'ZIP', title: 'ZIP' },
      { id: 'ADDRESS', title: 'ADDRESS' },
      { id: 'DOB', title: 'DOB' },
      { id: 'PROCESSED_AT', title: 'PROCESSED_AT' }
    ],
    append: true
  });

  const record = {
    FIRST: pick(row.FIRST || row.first || row.First),
    LAST: pick(row.LAST || row.last || row.Last),
    EMAIL: pick(row.EMAIL || row.email || row.Email),
    PHONE1: pick(row.PHONE1 || row.phone || row.Phone),
    ZIP: pick(row.ZIP || row.zip || row.Zip),
    ADDRESS: pick(row.ADDRESS || row.address || row.Address),
    DOB: pick(row.DOB || row.dob || row.Dob),
    PROCESSED_AT: new Date().toISOString()
  };

  await csvWriter.writeRecords([record]);
  console.log(`✓ Added to leads.csv: ${record.FIRST} ${record.LAST}`);
}

function isLeadProcessed(row, existingLeads) {
  const first = pick(row.FIRST || row.first || row.First);
  const last = pick(row.LAST || row.last || row.Last);
  const email = pick(row.EMAIL || row.email || row.Email);
  const phone = pick(row.PHONE1 || row.phone || row.Phone);
  const key = `${first}|${last}|${email}|${phone}`.toLowerCase().trim();
  return existingLeads.has(key);
}

async function pageReady(page) {
  const markers = [
    '#first',
    '#submit-btn',
    'select[name="dobmonth"]',
    'select[name="dobday"]',
    'select[name="dobyear"]',
    '.checkbox-sub-container',
    '.cq1answers.buttons',
    '.offer-yes-btn',
    'input[type="checkbox"]',
    'input[type="radio"]',
    '#flow-skip',
    'button.btn-flow'
  ];

  let lastBodyLen = -1;
  let staticCount = 0;
  let redirectAttempts = 0;
  const MAX_STATIC_DEAD = 5;
  const MAX_WAIT = 60;
  const MAX_REDIRECTS = 3;

  await sleep(1000);
  const readyStart = Date.now();

  for (let i = 0; i < MAX_WAIT; i++) {
    const waitSecs = Math.floor((Date.now() - readyStart) / 1000);
    process.stdout.write(`\r  Waiting for page... ${waitSecs}s`);
    const currentUrl = page.url();

    if (redirectAttempts < MAX_REDIRECTS) {
      const redirectUrl = await page.evaluate(() => {
        const url = window.location.href;
        if (/metarefresh|shors\.site/i.test(url)) {
          try {
            const match = url.match(/[?&]t=([^&]+)/);
            if (match) return atob(decodeURIComponent(match[1]));
          } catch {}
        }
        const meta = document.querySelector('meta[http-equiv="refresh"], meta[http-equiv="Refresh"]');
        if (meta) {
          const content = meta.getAttribute('content') || '';
          const m = content.match(/url=(.+)/i);
          if (m) return m[1].trim().replace(/^['"]|['"]$/g, '');
        }
        return null;
      }).catch(() => null);

      if (redirectUrl) {
        redirectAttempts++;
        console.log(`  Redirect detected (attempt ${redirectAttempts}/${MAX_REDIRECTS}), navigating to: ${redirectUrl.slice(0, 80)}...`);
        await page.goto(redirectUrl, { waitUntil: 'domcontentloaded', timeout: 15000 }).catch(() => {});
        await sleep(1500);
        lastBodyLen = -1;
        staticCount = 0;
        continue;
      }
    } else if (/metarefresh|shors\.site/i.test(currentUrl)) {
      console.log('Max redirect attempts reached, bailing');
      return false;
    }

    for (const sel of markers) {
      const loc = page.locator(sel);
      if (await loc.count().catch(() => 0)) {
        const visible = await loc.first().isVisible().catch(() => false);
        if (visible) {
          process.stdout.write(`\r`);
          return true;
        }
      }
    }

    const bodyLen = await page.evaluate(() => document.body ? document.body.innerHTML.length : 0).catch(() => 0);

    if (bodyLen === lastBodyLen) staticCount++;
    else staticCount = 0;
    lastBodyLen = bodyLen;

    if (staticCount >= MAX_STATIC_DEAD && bodyLen < 200 && !/metarefresh|shors\.site/i.test(currentUrl)) {
      process.stdout.write(`\r`);
      console.log('Page is static and empty → bailing early');
      return false;
    }

    await sleep(1000);
  }
  process.stdout.write(`\r`);
  return false;
}

async function closeOverlay(page) {
  await page.evaluate(() => {
    const popup = document.getElementById('popup_offer_desktop');
    if (popup) {
      popup.style.display = 'none';
      popup.style.visibility = 'hidden';
      popup.style.pointerEvents = 'none';
    }
    const consent = document.getElementById('consent_modal');
    if (consent) {
      consent.style.display = 'none';
      consent.style.visibility = 'hidden';
      consent.style.pointerEvents = 'none';
    }
  }).catch(() => {});
}

async function setPhone(page, value) {
  const phone = page.locator('#phone');
  await phone.scrollIntoViewIfNeeded().catch(() => {});
  await phone.click({ force: true }).catch(() => {});
  await page.keyboard.press('Control+A').catch(() => {});
  await page.keyboard.press('Backspace').catch(() => {});
  await page.keyboard.type(value, { delay: 40 });
}

async function fillRow(page, row) {
  const first = pick(row.FIRST || row.first || row.First);
  const last = pick(row.LAST || row.last || row.Last);
  const email = pick(row.EMAIL || row.email || row.Email);
  const phone = normalizePhone(row.PHONE1 || row.phone || row.Phone);
  const zip = normalizeZip(row.ZIP || row.zip || row.Zip);
  const address = pick(row.ADDRESS || row.address || row.Address);

  const dob = parseDOB(row.DOB || row.dob || row.Dob);
  const month = dob.month;
  const day = dob.day;
  const year = dob.year;

  await page.fill('#first', first).catch(() => {});
  await page.fill('#last', last).catch(() => {});
  await page.fill('#address', address).catch(() => {});
  await page.fill('#zip', zip).catch(() => {});
  await page.fill('#email', email).catch(() => {});
  await setPhone(page, phone);

  if (month) await page.selectOption('select[name="dobmonth"]', month).catch(() => {});
  if (day) await page.selectOption('select[name="dobday"]', day).catch(() => {});
  if (year) await page.selectOption('select[name="dobyear"]', year).catch(() => {});

  await closeOverlay(page);

  const agree = page.locator('input[name="iagree"]');
  if (await agree.count()) {
    await agree.first().check({ force: true }).catch(async () => {
      await page.evaluate(() => {
        const e = document.querySelector('input[name="iagree"]');
        if (e) {
          e.checked = true;
          e.dispatchEvent(new Event('input', { bubbles: true }));
          e.dispatchEvent(new Event('change', { bubbles: true }));
        }
      }).catch(() => {});
    });
  }
}

async function createBrowser() {
  return chromium.launch({
    headless: false,
    proxy: {
      server: 'http://ca.proxy-jet.io:1010',
      username: '241120H9lZQ-resi-US',
      password: 'w3ILawmH11RrtYV'
    }
  });
}

async function logVisibleCandidates(page, label, selector) {
  const loc = page.locator(selector);
  const count = await loc.count().catch(() => 0);
  console.log(`-- ${label}: found ${count}`);
  for (let i = 0; i < Math.min(count, 15); i++) {
    const el = loc.nth(i);
    const visible = await el.isVisible().catch(() => false);
    const txt = pick(await el.textContent().catch(() => ''));
    if (visible) console.log(`   [${i}] ${txt.slice(0, 140)}`);
  }
}

async function clickPrimaryFormContinue(page) {
  const selectors = [
    '#submit-btn',
    'input#submit-btn',
    'input[value="CONTINUE"]',
    'input[type="submit"][value="CONTINUE"]',
    '#submit-btn[value="CONTINUE"]'
  ];

  for (const selector of selectors) {
    const btn = page.locator(selector);
    const count = await btn.count().catch(() => 0);
    if (!count) continue;

    try {
      await btn.first().scrollIntoViewIfNeeded().catch(() => {});
      await sleep(300);
      await btn.first().click({ force: true, timeout: 5000 }).catch(async () => {
        await btn.first().click({ timeout: 5000 }).catch(async () => {
          await page.evaluate(sel => {
            const el = document.querySelector(sel);
            if (el) el.click();
          }, selector).catch(() => {});
        });
      });
      await page.waitForNavigation({ waitUntil: 'domcontentloaded', timeout: 15000 }).catch(() => {});
      return true;
    } catch {}
  }

  await page.evaluate(() => {
    const form = document.querySelector('form') || document.querySelector('#submit-btn')?.closest('form');
    if (form) form.submit();
    else {
      const btn = document.querySelector('#submit-btn');
      if (btn) btn.click();
    }
  }).catch(() => {});

  await sleep(2000);
  return true;
}

async function getPageSignature(page) {
  try {
    return await page.evaluate(() => {
      // Use only VISIBLE interactive elements so we don't get fooled by
      // stacked coreg panels that stay in DOM after being completed
      const visibleInputs = Array.from(
        document.querySelectorAll('input, button, a, select, textarea')
      ).filter(el => {
        const r = el.getBoundingClientRect();
        return r.width > 0 && r.height > 0 &&
          getComputedStyle(el).visibility !== 'hidden' &&
          getComputedStyle(el).display !== 'none';
      });
      const url = location.href;
      const title = document.title || '';
      // Capture IDs/values of visible inputs as the signature
      const ids = visibleInputs.map(el => el.id || el.name || el.tagName).join(',');
      const checked = Array.from(document.querySelectorAll('input[type="checkbox"]:checked')).map(e => e.id || e.value).join(',');
      return JSON.stringify({ url, title, ids, checked, count: visibleInputs.length });
    });
  } catch {
    return '';
  }
}

async function detectPageType(page) {
  const data = await page.evaluate(() => {
    return {
      formFields: document.querySelectorAll(
        '#first, #last, #email, #phone, #address, #zip, select[name="dobmonth"], select[name="dobday"], select[name="dobyear"]'
      ).length,
      unlockBtn: (document.querySelectorAll('a.btn.btn-primary').length > 0 &&
        document.body.innerText.includes('Unlock')) ? 1 : 0,
      coregContainer: document.querySelectorAll(
        '[class*="customquestions"], .checkbox-main-container-single, .checkbox-sub-container'
      ).length,
      coregBtn: document.querySelectorAll(
        'button.coreg-continue-button, [id^="checkbox-coreg-button-"]'
      ).length,
      checkboxes: document.querySelectorAll('input[type="checkbox"]').length,
      radios: document.querySelectorAll('input[type="radio"]').length,
      yesAnchors: document.querySelectorAll(
        'a.offer-yes-btn-container, a.offer-yes-btn, a[onclick*="submit_lead"], a[onclick*="set_user_attributes"]'
      ).length,
      yesButtons: document.querySelectorAll(
        '.offer-yes-btn, .cq-next-yes, div[onclick*="submit_lead"]'
      ).length,
    };
  }).catch(() => ({
    formFields: 0, unlockBtn: 0, coregContainer: 0, coregBtn: 0,
    checkboxes: 0, radios: 0, yesAnchors: 0, yesButtons: 0
  }));

  console.log('  [detect]', JSON.stringify(data));

  if (data.formFields > 0) return 'form';
  if (data.unlockBtn > 0) return 'unlock';
  if (data.coregContainer > 0 || data.coregBtn > 0 || data.checkboxes > 0) return 'coreg';
  if (data.radios > 0) return 'radio';
  if (data.yesAnchors > 0 || data.yesButtons > 0) return 'yesno';
  return 'unknown';
}

async function clickAndVerify(page, fn, waitMs = 1200) {
  const before = await getPageSignature(page);
  await fn().catch(() => {});
  await sleep(waitMs);
  const after = await getPageSignature(page);
  return before !== after;
}


async function clickFlowSkipIfPresent(page) {
  const btn = page.locator('#flow-skip');
  const count = await btn.count().catch(() => 0);
  if (!count) return false;
  const visible = await btn.first().isVisible().catch(() => false);
  if (!visible) return false;

  console.log('  → #flow-skip detected — clicking it now');
  const fired = await page.evaluate(() => {
    const btn = document.querySelector('#flow-skip');
    if (!btn) return false;
    btn.scrollIntoView({ block: 'center' });
    const fn = btn.getAttribute('onclick') || '';
    if (fn) {
      try { eval(fn.replace(/^javascript:/i, '').replace(/window\.scrollTo[^;]+;?/g, '')); return true; } catch {}
    }
    ['mousedown', 'mouseup', 'click'].forEach(t =>
      btn.dispatchEvent(new MouseEvent(t, { bubbles: true, cancelable: true, view: window }))
    );
    btn.click();
    return true;
  }).catch(() => false);

  if (!fired) {
    await btn.first().click({ force: true }).catch(() => {});
  }

  await Promise.race([
    page.waitForNavigation({ waitUntil: 'domcontentloaded', timeout: 10000 }),
    sleep(10000)
  ]).catch(() => {});
  await sleep(800);
  return true;
}

async function handleYesNoOfferPage(page) {
  const selectors = [
    'a.offer-yes-btn-container',
    'a#continue9136',
    'a.offer-yes-btn-container#continue9136',
    '.offer-yes-btn-container',
    '.offer-yes-btn',
    '.offer-yes-btn-container .offer-yes-btn',
    'a[onclick*="submit_lead"]',
    'a[onclick*="set_user_attributes"]',
    'text=Yes'
  ];

  for (const selector of selectors) {
    const loc = page.locator(selector);
    const count = await loc.count().catch(() => 0);
    for (let i = 0; i < count; i++) {
      const el = loc.nth(i);
      if (!(await el.isVisible().catch(() => false))) continue;

      const changed = await clickAndVerify(page, async () => {
        await el.scrollIntoViewIfNeeded().catch(() => {});
        await el.click({ force: true, timeout: 5000 }).catch(async () => {
          await el.evaluate(node => {
            node.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window }));
            if (node.tagName === 'A' || node.tagName === 'BUTTON' || node.tagName === 'DIV') node.click();
          }).catch(() => {});
        });
      }, 1500);

      if (changed) return true;
    }
  }

  return false;
}


async function clickContinueBtn(page) {
  const allBtns = page.locator(
    'button.coreg-continue-button, button[id^="checkbox-coreg-button-"], [id^="checkbox-coreg-button-"], .coreg-continue-button'
  );
  const btnCount = await allBtns.count().catch(() => 0);
  let continueSel = null;
  for (let i = 0; i < btnCount; i++) {
    const b = allBtns.nth(i);
    if (await b.isVisible().catch(() => false)) { continueSel = b; break; }
  }

  if (continueSel) {
    console.log('  → Clicking coreg-continue-button via Playwright');
    await continueSel.scrollIntoViewIfNeeded().catch(() => {});
    await continueSel.click({ force: true, timeout: 5000 }).catch(() => {});
    await Promise.race([
      page.waitForNavigation({ waitUntil: 'domcontentloaded', timeout: 10000 }),
      sleep(10000)
    ]).catch(() => {});
    return true;
  }

  const fired = await page.evaluate(() => {
    const sel = 'button.coreg-continue-button, [id^="checkbox-coreg-button-"], .coreg-continue-button';
    const btn = Array.from(document.querySelectorAll(sel)).find(b => {
      const r = b.getBoundingClientRect();
      return r.width > 0 && r.height > 0;
    });
    if (!btn) return false;
    btn.scrollIntoView({ block: 'center' });
    const fn = btn.getAttribute('onclick') || '';
    if (fn) {
      try { eval(fn.replace(/return\s+false\s*;?/g, '')); } catch {}
    }
    ['mousedown', 'mouseup', 'click'].forEach(type => {
      btn.dispatchEvent(new MouseEvent(type, { bubbles: true, cancelable: true, view: window }));
    });
    btn.click();
    return true;
  }).catch(() => false);

  if (fired) {
    console.log('  → Fired coreg-continue-button via evaluate');
    await Promise.race([
      page.waitForNavigation({ waitUntil: 'domcontentloaded', timeout: 10000 }),
      sleep(10000)
    ]).catch(() => {});
    return true;
  }

  console.log('  ⚠️ No coreg-continue-button found');
  return false;
}

async function handleCoregPage(page) {

  let cqOptions = [];
  for (let attempt = 0; attempt < 3; attempt++) {
    const loc = page.locator('.cq1answers.buttons .offer-yes-btn, .cq1answers.buttons .cq-next-yes');
    const count = await loc.count().catch(() => 0);
    for (let i = 0; i < count; i++) {
      const el = loc.nth(i);
      if (!(await el.isVisible().catch(() => false))) continue;
      const text = pick(await el.textContent().catch(() => ''));
      if (!text || /none of the above|none|no thanks/i.test(text)) continue;
      cqOptions.push({ text, el });
    }
    if (cqOptions.length) break;
    await sleep(1000);
  }

  if (cqOptions.length) {
    const chosen = randomItem(cqOptions);
    console.log(`  → [CQ button] Clicking: "${chosen.text.trim().slice(0, 60)}"`);
    await chosen.el.scrollIntoViewIfNeeded().catch(() => {});
    const clicked = await chosen.el.click({ force: true, timeout: 4000 })
      .then(() => true)
      .catch(async () =>
        chosen.el.evaluate(node => {
          const fn = node.getAttribute('onclick') || '';
          if (fn) {
            try {
              eval(fn.replace(/window\.scrollTo[^;]*;?/g, '').replace(/return\s+false\s*;?/g, ''));
              return true;
            } catch {}
          }
          ['mousedown', 'mouseup', 'click'].forEach(t =>
            node.dispatchEvent(new MouseEvent(t, { bubbles: true, cancelable: true, view: window }))
          );
          node.click();
          return true;
        }).catch(() => false)
      );
    if (clicked) console.log(`    ✓ Clicked: "${chosen.text.trim().slice(0, 60)}"`);
    await Promise.race([
      page.waitForNavigation({ waitUntil: 'domcontentloaded', timeout: 10000 }),
      sleep(10000)
    ]).catch(() => {});
    return true;
  }

  const visibleInputs = page.locator(
    'input[type="checkbox"], .checkbox-sub-container input[type="checkbox"], .checkbox-main-container-single input[type="checkbox"]'
  );
  const count = await visibleInputs.count().catch(() => 0);
  const cbOptions = [];

  for (let i = 0; i < count; i++) {
    const el = visibleInputs.nth(i);
    if (!(await el.isVisible().catch(() => false))) continue;
    const isChecked = await el.isChecked().catch(() => false);
    if (isChecked) continue;
    const id = await el.getAttribute('id').catch(() => '');
    let text = '';
    if (id) text = await page.locator(`label[for="${id}"]`).textContent().catch(() => '');
    if (!text) {
      text = await el.evaluate(node => {
        const label = node.closest('label');
        if (label) return label.innerText || '';
        const p = node.parentElement;
        return p ? (p.innerText || '') : '';
      }).catch(() => '');
    }
    text = pick(text);
    if (/none of the above|none|no thanks/i.test(text)) continue;
    if (text && text.length < 200) cbOptions.push({ text, el });
  }

  if (!cbOptions.length) {
    console.log('  No options found — clicking continue anyway...');
    const clicked = await clickContinueBtn(page);
    if (!clicked) {
      // fallback: try yes/anchor buttons
      const yesLoc = page.locator('a.offer-yes-btn-container, .offer-yes-btn, a[onclick*="submit_lead"], a[onclick*="set_user_attributes"]');
      const yesCount = await yesLoc.count().catch(() => 0);
      for (let i = 0; i < yesCount; i++) {
        const el = yesLoc.nth(i);
        if (await el.isVisible().catch(() => false)) {
          console.log('  → Clicking yes/anchor fallback');
          await el.click({ force: true }).catch(() => {});
          await Promise.race([
            page.waitForNavigation({ waitUntil: 'domcontentloaded', timeout: 10000 }),
            sleep(10000)
          ]).catch(() => {});
          return true;
        }
      }
    }
    return clicked;
  }

  const chosen = randomItem(cbOptions);
  console.log(`  → [Checkbox] Checking: "${chosen.text.trim().slice(0, 60)}"`);
  await chosen.el.scrollIntoViewIfNeeded().catch(() => {});
  const clicked = await chosen.el.click({ force: true, timeout: 3000 })
    .then(() => true)
    .catch(async () =>
      chosen.el.check({ force: true, timeout: 3000 })
        .then(() => true)
        .catch(async () =>
          chosen.el.evaluate(node => {
            node.checked = true;
            node.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window }));
            node.dispatchEvent(new Event('change', { bubbles: true }));
            return true;
          }).catch(() => false)
        )
    );
  if (clicked) console.log(`    ✓ Checked: "${chosen.text.trim().slice(0, 60)}"`);

  await sleep(600);
  await clickContinueBtn(page);
  return true;
}

async function handleRadioPage(page) {
  const inputs = page.locator('input[type="radio"]');
  const count = await inputs.count().catch(() => 0);
  const options = [];

  for (let i = 0; i < count; i++) {
    const el = inputs.nth(i);
    if (!(await el.isVisible().catch(() => false))) continue;

    const id = await el.getAttribute('id').catch(() => '');
    let text = '';
    if (id) text = await page.locator(`label[for="${id}"]`).textContent().catch(() => '');
    if (!text) {
      text = await el.evaluate(node => {
        const label = node.closest('label');
        if (label) return label.innerText || '';
        const p = node.parentElement;
        return p ? (p.innerText || '') : '';
      }).catch(() => '');
    }

    text = pick(text);
    if (text && text.length < 140) options.push({ text, el });
  }

  if (!options.length) return false;

  const chosen = randomItem(options);
  if (!chosen) return false;

  const targets = [
    chosen.el,
    page.getByLabel(chosen.text, { exact: false }),
    page.locator(`label:has-text("${chosen.text.replace(/"/g, '\\"')}")`)
  ];

  for (const target of targets) {
    try {
      if (!(await target.count().catch(() => 0))) continue;
      const el = target.first();
      await el.scrollIntoViewIfNeeded().catch(() => {});
      await el.click({ force: true, timeout: 5000 }).catch(async () => {
        await el.check({ force: true, timeout: 5000 }).catch(() => {});
      });
      return true;
    } catch {}
  }

  return false;
}

async function processFollowUpPages(page) {
  let lastSig = '';
  let stuck = 0;

  for (let step = 0; step < 40; step++) {
    await sleep(1200);

    const sig = await getPageSignature(page);
    if (sig === lastSig) stuck++;
    else { stuck = 0; lastSig = sig; }

    if (stuck >= 3) {
      console.log('Page did not change after several checks, stopping.');
      return false;
    }

    const skipped = await clickFlowSkipIfPresent(page);
    if (skipped) continue;

    const type = await detectPageType(page);
    console.log('Follow-up page type:', type);

    if (type === 'coreg') {
      const ok = await handleCoregPage(page);
      if (!ok) return false;
      continue;
    }

    if (type === 'radio') {
      const ok = await handleRadioPage(page);
      if (!ok) return false;
      continue;
    }

    if (type === 'yesno') {
      const ok = await handleYesNoOfferPage(page);
      if (!ok) return false;
      continue;
    }

    const content = await page.content().catch(() => '');
    if (/thank you|submitted|success|complete/i.test(content)) return true;

    if ((await page.locator('input, button, a, select').count().catch(() => 0)) === 0) {
      return false;
    }

    // unknown type but page has elements — wait for page to settle and retry
    console.log('  Unknown type with elements present, waiting for page to settle...');
    await sleep(2000);
    continue;
  }

  return false;
}

async function runAttempt(row, attemptNo, existingLeads) {
  const browser = await createBrowser();
  const context = await browser.newContext();
  const page = await context.newPage();
  page.setDefaultTimeout(PAGE_TIMEOUT);
  page.setDefaultNavigationTimeout(PAGE_TIMEOUT);

  try {
    process.stdout.write('Navigating... 0s');
    const navStart = Date.now();
    const navTimer = setInterval(() => {
      const secs = Math.floor((Date.now() - navStart) / 1000);
      process.stdout.write(`\rNavigating... ${secs}s`);
    }, 1000);

    const response = await page.goto(url, {
      waitUntil: 'domcontentloaded',
      timeout: 10000
    }).catch(err => {
      clearInterval(navTimer);
      const secs = Math.floor((Date.now() - navStart) / 1000);
      process.stdout.write(`\r`);
      console.log(`💥 Navigation failed after ${secs}s: ${err.message}`);
      return null;
    });

    clearInterval(navTimer);
    const navSecs = Math.floor((Date.now() - navStart) / 1000);
    process.stdout.write(`\r`);
    console.log(`Navigated in ${navSecs}s`);

    if (!response) {
      await browser.close();
      return false;
    }

    const pageStatus = response.status();
    if (pageStatus >= 400) {
      console.log(`HTTP ${pageStatus} → closing session immediately`);
      await browser.close();
      return false;
    }

    const earlyContent = await page.content().catch(() => '');
    if (/can't be reached|cannot be reached|err_tunnel|err_|site can't|refused to connect|connection timed out|dns_probe|not available|access denied|403 forbidden|404 not found|zero sized reply|squid|could not be retrieved|failed to retrieve|the requested url|no data received|this site can|network error|http error 502|http error 503|http error 500|502 bad gateway|503 service|500 internal|is currently unable|unable to handle this request/i.test(earlyContent)) {
      console.log('Error page detected → closing session immediately');
      await browser.close();
      return false;
    }

    await pageReady(page);
    await sleep(2500);

    const type = await detectPageType(page);
    console.log('Page type detected:', type);

    if (type !== 'form' && type !== 'unlock') {
      console.log('Wrong page → closing session');
      await browser.close();
      return false;
    }

    const unlock = page.locator('a.btn.btn-primary', { hasText: 'Unlock Now' });
    if (await unlock.count()) {
      await unlock.first().click().catch(() => {});
      await sleep(2000);
      await pageReady(page);
    }

    await fillRow(page, row);
    const continued = await clickPrimaryFormContinue(page);
    await pageReady(page);
    await sleep(2500);

    const leadKey = `${pick(row.FIRST || row.first || row.First)}|${pick(row.LAST || row.last || row.Last)}|${pick(row.EMAIL || row.email || row.Email)}|${pick(row.PHONE1 || row.phone || row.Phone)}`.toLowerCase().trim();

    if (continued && !existingLeads.has(leadKey)) {
      await addProcessedLead(row);
      existingLeads.add(leadKey);
    }

    const finished = await processFollowUpPages(page);

    await page.screenshot({
      path: path.join(outDir, `run-${Date.now()}-attempt-${attemptNo}.png`),
      fullPage: true
    }).catch(() => {});

    console.log(finished ? 'SUCCESS run completed' : 'FLOW ENDED / INCOMPLETE');
    await browser.close();
    return existingLeads.has(leadKey) ? true : finished;
  } catch (err) {
    console.log('ERROR:', err.message);
    try {
      await browser.close();
    } catch {}
    return false;
  }
}

async function cleanupBrowsers() {
  try {
    const { execSync } = require('child_process');
    execSync('pkill -f "chromium|chrome" 2>/dev/null || true', { stdio: 'ignore' });
  } catch {}
  await sleep(1000);
}

async function main() {
  ensureLeadsCsvHeaders();

  console.log('Loading existing leads from leads.csv...');
  const existingLeads = await loadExistingLeads();
  console.log(`Found ${existingLeads.size} existing leads`);

  const wb = XLSX.readFile(file, { cellDates: true });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: '', raw: true });
  console.log(`Processing ${rows.length} rows from "${wb.SheetNames[0]}"`);

  let successCount = 0;
  let skipCount = 0;

  for (let i = 0; i < rows.length; i++) {
    console.log(`\n=== PROCESSING ROW ${i + 1}/${rows.length} ===`);

    const row = rows[i];
    if (isLeadProcessed(row, existingLeads)) {
      console.log('⏭️ SKIPPED: Already in leads.csv');
      skipCount++;
      continue;
    }

    let success = false;
    for (let attempt = 1; attempt <= MAX_ATTEMPTS; attempt++) {
      console.log(`Attempt ${attempt}`);
      success = await runAttempt(row, attempt, existingLeads);
      if (success) break;
      console.log('Retrying with new session...');
    }

    if (success) {
      successCount++;
    } else {
      console.log('FAILED row after max attempts');
    }

    await cleanupBrowsers();
  }

  console.log('\n=== SUMMARY ===');
  console.log(`✅ Processed: ${successCount}`);
  console.log(`⏭️  Skipped: ${skipCount}`);
  console.log('DONE ALL ROWS');
}

main().catch(err => {
  console.error(err);
  process.exit(1);
});