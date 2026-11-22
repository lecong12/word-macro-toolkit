const fs = require('fs');
const path = require('path');

const dir = path.join(__dirname, '..', 'src', 'working');
const files = fs.readdirSync(dir).filter(f => f.endsWith('.bas')).sort();

console.log('Total modules:', files.length);
console.log('------------------------------------------------------------');
console.log('File'.padEnd(16) + 'VB_Name'.padEnd(14) + 'Procedures');
console.log('------------------------------------------------------------');

files.forEach(f => {
  const content = fs.readFileSync(path.join(dir, f), 'latin1');
  const attr = content.match(/Attribute\s+VB_Name\s*=\s*"([^"]+)"/i);
  const name = attr ? attr[1] : '(missing)';
  const procs = content.match(/^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Sub|Function)\s+(\w+)/gim) || [];
  const procNames = procs.map(p => p.trim().split(/\s+/).pop().replace(/\(.*$/, '')).join(', ');
  console.log(f.padEnd(16) + name.padEnd(14) + procNames);
});
console.log('------------------------------------------------------------');
