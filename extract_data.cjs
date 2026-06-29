// Extracts all arrays from bda-payback/src/data.js and moneyball/src/App.jsx
// and writes a unified data.json
const fs = require("fs");
const path = require("path");

function extractArray(src, varName) {
  const re = new RegExp(`(?:export )?const ${varName}\\s*=\\s*(\\[.*?\\]);`, "s");
  const m = src.match(re);
  if (!m) return null;
  return JSON.parse(m[1]);
}

const payback = fs.readFileSync("C:/Users/ECP/bda-payback/src/data.js", "utf8");
const moneyball = fs.readFileSync("C:/Users/ECP/moneyball/src/App.jsx", "utf8");

const data = {
  // Rep lists
  BDA_REPS:        extractArray(payback, "BDA_REPS"),
  SBC_REPS:        extractArray(payback, "SBC_REPS"),
  BDA_KNOWN_REPS:  extractArray(payback, "BDA_KNOWN_REPS"),
  SBC_KNOWN_REPS:  extractArray(payback, "SBC_KNOWN_REPS"),
  BDA_ACTIVE:      extractArray(payback, "BDA_ACTIVE"),
  BDA_INACTIVE:    extractArray(payback, "BDA_INACTIVE"),
  SBC_ACTIVE:      extractArray(payback, "SBC_ACTIVE"),
  SBC_INACTIVE:    extractArray(payback, "SBC_INACTIVE"),
  // Date-indexed RPR
  BDA_DATE_RPR:    extractArray(payback, "BDA_DATE_RPR"),
  SBC_DATE_RPR:    extractArray(payback, "SBC_DATE_RPR"),
  // Month-indexed RPR
  BDA_MONTH_RPR:   extractArray(payback, "BDA_MONTH_RPR"),
  SBC_MONTH_RPR:   extractArray(payback, "SBC_MONTH_RPR"),
  // Cost arrays
  BDA_DATE_CCOST:  extractArray(payback, "BDA_DATE_CCOST"),
  SBC_DATE_CCOST:  extractArray(payback, "SBC_DATE_CCOST"),
  BDA_MONTH_CCOST: extractArray(payback, "BDA_MONTH_CCOST"),
  SBC_MONTH_CCOST: extractArray(payback, "SBC_MONTH_CCOST"),
  // Payback multiple arrays
  BDA_DATE_PC:     extractArray(payback, "BDA_DATE_PC"),
  SBC_DATE_PC:     extractArray(payback, "SBC_DATE_PC"),
  BDA_MONTH_PC:    extractArray(payback, "BDA_MONTH_PC"),
  SBC_MONTH_PC:    extractArray(payback, "SBC_MONTH_PC"),
};

// Sanity check
for (const [k, v] of Object.entries(data)) {
  if (v === null) console.warn(`WARNING: ${k} not found`);
  else console.log(`${k}: ${Array.isArray(v) ? v.length + " rows" : typeof v}`);
}

const outPath = path.join(__dirname, "data.json");
fs.writeFileSync(outPath, JSON.stringify(data, null, 2));
console.log(`\nWrote ${outPath}`);
