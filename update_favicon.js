const fs = require('fs');

const code = fs.readFileSync('Code.gs', 'utf8');
const base64 = fs.readFileSync('/tmp/favicon_b64.txt', 'utf8').trim();

const newCode = code.replace(
  /\.setFaviconUrl\('.*?'\)/,
  `.setFaviconUrl('data:image/png;base64,${base64}')`
);

fs.writeFileSync('Code.gs', newCode);
console.log('Favicon updated successfully.');
