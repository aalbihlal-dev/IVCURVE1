// build.js — run once before every Netlify deploy
// Usage: node build.js
// What it does: replaces __BUILDTIME__ in sw.js with current Unix timestamp
//               so every deploy gets a unique cache version automatically

const fs = require('fs');
const path = require('path');

const swPath = path.join(__dirname, 'sw.js');
const buildTime = Date.now().toString();

let sw = fs.readFileSync(swPath, 'utf8');

// Replace any existing timestamp or placeholder
sw = sw.replace(/iv-[\d]+/, `iv-${buildTime}`)
       .replace(/iv-__BUILDTIME__/, `iv-${buildTime}`);

fs.writeFileSync(swPath, sw);
console.log(`✅ sw.js versioned: iv-${buildTime}`);
