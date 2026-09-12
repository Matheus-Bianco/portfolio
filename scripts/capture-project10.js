/**
 * Captura prints do Painel SEDUC-RS para o projeto 10.
 * Uso: node scripts/capture-project10.js
 */
const path = require('path');
const fs = require('fs');
const puppeteer = require('puppeteer');
const sharp = require('sharp');

const URL = 'https://indicadores.educacao.rs.gov.br/';
const OUT_DIR = path.join(__dirname, '..', 'assets', 'images', 'projects', 'project10');

const SHOTS = [
  { file: '1.jpg', tab: null, ready: '.home-wrap, .home-banner', waitMap: false, note: 'hub' },
  { file: '2.jpg', tab: '#tab-acesso', ready: '#tab-acesso', waitMap: true, note: 'acesso' },
  { file: '3.jpg', tab: '#tab-fluxo', ready: '#tab-fluxo', waitMap: true, note: 'fluxo' },
  { file: '4.jpg', tab: '#tab-escolas', ready: '#tab-escolas', waitMap: true, note: 'escolas' },
];

async function waitReady(page, extraMs = 1500) {
  await page.waitForFunction(() => {
    const loading = document.getElementById('loading');
    const home = document.querySelector('.home-wrap, .home-banner, .home-banner-title');
    const loadedFlag = window.S && window.S.loaded;
    return !!(loadedFlag || home) && !(loading && loading.offsetParent !== null && loading.textContent.includes('Carregando'));
  }, { timeout: 90000 });
  await new Promise((r) => setTimeout(r, extraMs));
}

async function clickTab(page, selector) {
  await page.waitForSelector(selector, { timeout: 15000 });
  await page.evaluate((sel) => {
    const el = document.querySelector(sel);
    if (el) el.click();
  }, selector);
  await new Promise((r) => setTimeout(r, 2500));
  if (selector !== '#tab-home') {
    await page.waitForFunction(() => !document.body.classList.contains('sidebar-hidden') || document.querySelector('.leaflet-container, canvas, .section-sticky, .kpi-strip'), { timeout: 20000 }).catch(() => {});
  }
}

async function waitMap(page) {
  await page.waitForSelector('.leaflet-container, canvas', { timeout: 25000 }).catch(() => {});
  await new Promise((r) => setTimeout(r, 2200));
}

async function capture(page, dest) {
  const raw = path.join(OUT_DIR, dest.replace('.jpg', '-raw.png'));
  await page.screenshot({ path: raw, type: 'png', fullPage: false });
  await sharp(raw)
    .resize(1200, null, { withoutEnlargement: true })
    .jpeg({ quality: 86 })
    .toFile(path.join(OUT_DIR, dest));
  fs.unlinkSync(raw);
  console.log('ok', dest);
}

async function main() {
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const browser = await puppeteer.launch({
    headless: 'new',
    args: ['--no-sandbox', '--disable-setuid-sandbox', '--window-size=1600,900'],
    defaultViewport: { width: 1600, height: 900, deviceScaleFactor: 1 },
  });
  const page = await browser.newPage();
  page.setDefaultTimeout(90000);
  await page.goto(URL, { waitUntil: 'networkidle2', timeout: 90000 });
  await waitReady(page, 2000);

  for (const shot of SHOTS) {
    if (shot.tab) await clickTab(page, shot.tab);
    if (shot.waitMap) await waitMap(page);
    await page.evaluate(() => {
      const el = document.querySelector('.main-scroll, #main-content, .main');
      if (el) el.scrollTop = 0;
      window.scrollTo(0, 0);
    });
    await new Promise((r) => setTimeout(r, 600));
    await capture(page, shot.file);
  }

  await browser.close();
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
