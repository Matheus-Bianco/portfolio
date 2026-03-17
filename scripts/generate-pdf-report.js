/**
 * Gera um PDF contínuo do portfólio em formato de relatório.
 * Usa o template portfolio-report.html (layout otimizado para documento).
 * Uso: npm run generate-pdf
 */

const path = require('path');
const fs = require('fs');
const http = require('http');
const puppeteer = require('puppeteer');

const PORT = 3847;
const REPORT_URL = `http://localhost:${PORT}/portfolio-report.html`;

function serveStatic(port) {
  return new Promise((resolve) => {
    const server = http.createServer((req, res) => {
      let filePath = req.url === '/' ? 'portfolio-report.html' : req.url.replace(/^\//, '');
      filePath = path.join(__dirname, '..', filePath);
      if (!path.extname(filePath)) filePath += '.html';
      if (!fs.existsSync(filePath)) {
        res.writeHead(404);
        res.end();
        return;
      }
      const ext = path.extname(filePath);
      const types = {
        '.html': 'text/html',
        '.css': 'text/css',
        '.js': 'application/javascript',
        '.png': 'image/png',
        '.jpg': 'image/jpeg',
        '.jpeg': 'image/jpeg',
        '.gif': 'image/gif',
        '.svg': 'image/svg+xml',
        '.ico': 'image/x-icon',
        '.woff': 'font/woff',
        '.woff2': 'font/woff2',
      };
      res.setHeader('Content-Type', types[ext] || 'application/octet-stream');
      res.end(fs.readFileSync(filePath));
    });
    server.listen(port, () => resolve(server));
  });
}

async function generatePdf() {
  const rootDir = path.join(__dirname, '..');
  const outputPath = path.join(rootDir, 'portfolio-relatorio.pdf');

  console.log('Iniciando servidor local...');
  const server = await serveStatic(PORT);

  let browser;
  try {
    browser = await puppeteer.launch({
      headless: 'new',
      args: ['--no-sandbox', '--disable-setuid-sandbox'],
    });

    const page = await browser.newPage();
    await page.setViewport({ width: 794, height: 1123 }); // A4 proporção
    await page.goto(REPORT_URL, { waitUntil: 'networkidle0', timeout: 15000 });
    await new Promise((r) => setTimeout(r, 500));

    await page.pdf({
      path: outputPath,
      format: 'A4',
      printBackground: true,
      margin: { top: '15mm', right: '12mm', bottom: '15mm', left: '12mm' },
    });

    console.log(`\nPDF gerado com sucesso: ${outputPath}`);
  } finally {
    if (browser) await browser.close();
    server.close();
  }
}

generatePdf().catch((err) => {
  console.error('Erro:', err.message);
  process.exit(1);
});
