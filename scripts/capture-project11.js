/**
 * Captura prints do Painel de Dados Abertos de Joinville para o projeto 11.
 * Uso: node scripts/capture-project11.js
 */
const path = require('path');
const fs = require('fs');
const puppeteer = require('puppeteer');
const sharp = require('sharp');

const URL = 'https://matheus-bianco.github.io/painel-indicadores-joinville/';
const OUT_DIR = path.join(__dirname, '..', 'assets', 'images', 'projects', 'project11');

const SHOTS = [
  { file: '1.jpg', tab: null, waitMap: false },
  { file: '2.jpg', tab: '#tab-censo-ibge', waitMap: false },
  { file: '3.jpg', tab: '#tab-acesso', waitMap: true },
  { file: '4.jpg', tab: '#tab-escolas', waitMap: true },
];

async function waitReady(page) {
  await page.waitForFunction(() => {
    const home = document.querySelector('.home-wrap, .home-banner, .home-banner-title');
    return !!(window.S && window.S.loaded) || !!home;
  }, { timeout: 90000 });
  await new Promise((r) => setTimeout(r, 2000));
}

async function clickTab(page, selector) {
  await page.waitForSelector(selector, { timeout: 15000 });
  await page.evaluate((sel) => document.querySelector(sel)?.click(), selector);
  await new Promise((r) => setTimeout(r, 2800));
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
    args: ['--no-sandbox', '--disable-setuid-sandbox'],
    defaultViewport: { width: 1600, height: 900, deviceScaleFactor: 1 },
  });
  const page = await browser.newPage();
  page.setDefaultTimeout(90000);
  await page.goto(URL, { waitUntil: 'networkidle2', timeout: 90000 });
  await waitReady(page);

  for (const shot of SHOTS) {
    if (shot.tab) await clickTab(page, shot.tab);
    if (shot.waitMap) {
      await page.waitForSelector('.leaflet-container, canvas', { timeout: 25000 }).catch(() => {});
      await new Promise((r) => setTimeout(r, 2200));
    }
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
