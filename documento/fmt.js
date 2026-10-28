// Utilidades de formato y acceso a los resultados del modelo
const fs = require("fs");
const path = require("path");
const R = JSON.parse(fs.readFileSync(path.join(__dirname, "..", "modelo", "resultados.json"), "utf8"));

function miles(n, dec = 0) {
  const neg = n < 0;
  let s = Math.abs(n).toFixed(dec);
  let [ent, frac] = s.split(".");
  ent = ent.replace(/\B(?=(\d{3})+(?!\d))/g, ".");
  return (neg ? "−" : "") + ent + (frac ? "," + frac : "");
}
const eur = (n, dec = 0) => miles(n, dec) + " €";
const pct = (x, dec = 1) => miles(x * 100, dec) + " %";
const num = (n, dec = 1) => miles(n, dec);

module.exports = { R, miles, eur, pct, num };
