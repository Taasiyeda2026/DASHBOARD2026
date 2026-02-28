const http = require('http');
const fs = require('fs');
const path = require('path');

const פורט = 5000;
const סוגי_קובץ = {
  '.html': 'text/html',
  '.css': 'text/css',
  '.js': 'application/javascript',
  '.json': 'application/json',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.svg': 'image/svg+xml',
  '.ico': 'image/x-icon',
};

const שרת = http.createServer((בקשה, תגובה) => {
  let נתיב_קובץ = path.join(__dirname, בקשה.url === '/' ? 'index.html' : בקשה.url);
  const סיומת = path.extname(נתיב_קובץ).toLowerCase();
  const סוג_תוכן = סוגי_קובץ[סיומת] || 'application/octet-stream';

  fs.readFile(נתיב_קובץ, (שגיאה, נתונים) => {
    if (שגיאה) {
      תגובה.writeHead(404, { 'Content-Type': 'text/plain', 'Cache-Control': 'no-cache' });
      תגובה.end('לא נמצא');
      return;
    }
    תגובה.writeHead(200, { 'Content-Type': סוג_תוכן, 'Cache-Control': 'no-cache, no-store, must-revalidate' });
    תגובה.end(נתונים);
  });
});

שרת.listen(פורט, '0.0.0.0', () => {
  console.log(`השרת פועל בכתובת http://0.0.0.0:${פורט}`);
});
