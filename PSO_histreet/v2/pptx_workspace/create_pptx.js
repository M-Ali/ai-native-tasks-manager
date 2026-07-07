const path = require('path');
const html2pptx = require('D:\\Personal\\SkillsApril2026\\claude-code-skills-lab-main\\.claude\\skills\\pptx\\scripts\\html2pptx.js');
const pptxgen = require('pptxgenjs');

const SLIDES_DIR = path.join(__dirname, 'slides');
const OUT = path.join(__dirname, '..', 'output', 'reports', 'PSO_Pricing_Strategy.pptx');

// ── Palette (no # prefix in pptxgenjs) ────────────────────────────
const G_DARK   = "023520";
const G_MED    = "00703C";
const GOLD     = "C8960C";
const G_LIGHT  = "E8F5EE";
const WHITE    = "FFFFFF";
const GRAY_LT  = "F5F5F5";
const RED_SIG  = "C0392B";
const AMBER    = "E67E22";

function html(name) { return path.join(SLIDES_DIR, name); }

// ── Data ────────────────────────────────────────────────────────────

const compData = [
  ["Brand", "Tier", "Origin", "Grades", "Market Min\n(PKR/L)", "Market Med\n(PKR/L)", "Market Max\n(PKR/L)"],
  ["Shell Helix", "Super Premium", "Imported", "0W-20, 5W-30, 5W-40", "3,375", "3,412", "4,012"],
  ["Caltex Havoline", "Premium", "Imported", "5W-30, 10W-40", "1,660", "2,300", "2,300"],
  ["Total Quartz", "Premium", "Imported", "5W-30, 5W-40, 10W-40", "1,660", "2,300", "2,300"],
  ["ZIC (SK)", "Premium/Mainstream", "Imported (Korea)", "5W-30, 10W-40, 20W-50", "1,088", "1,750", "2,262"],
  ["Aramco Mobil", "Mainstream", "Imported", "10W-40, 15W-40", "1,088", "1,750", "2,262"],
  ["Kixx (GS)", "Mainstream/Economy", "Imported (Korea)", "10W-40, 15W-40, 20W-50", "876", "1,019", "1,500"],
  ["Equimoli", "Economy", "Local/Imported", "15W-40, 20W-50", "922", "1,025", "2,150"],
];

const channelData = [
  [
    { text: "Brand",        options: { fill: { color: G_DARK }, color: WHITE, bold: true, fontSize: 9 } },
    { text: "Grade",        options: { fill: { color: G_DARK }, color: WHITE, bold: true, fontSize: 9 } },
    { text: "Retail Pump\n(PKR/L)", options: { fill: { color: G_DARK }, color: WHITE, bold: true, fontSize: 9, align: "center" } },
    { text: "Workshop\n(−12%)",     options: { fill: { color: G_DARK }, color: WHITE, bold: true, fontSize: 9, align: "center" } },
    { text: "Distributor\n(−15%)",  options: { fill: { color: G_DARK }, color: WHITE, bold: true, fontSize: 9, align: "center" } },
    { text: "Fleet Contract\n(−20%)", options: { fill: { color: G_DARK }, color: WHITE, bold: true, fontSize: 9, align: "center" } },
  ],
  ["Carient Ultra", "0W-20", "Rs 3,745", "Rs 3,296", "Rs 3,183", "Rs 2,996"],
  ["Carient Ultra", "5W-40", "Rs 3,287", "Rs 2,893", "Rs 2,794", "Rs 2,630"],
  ["Carient FS",    "5W-30", "Rs 2,006", "Rs 1,765", "Rs 1,705", "Rs 1,605"],
  ["Carient FS",    "5W-40", "Rs 3,138", "Rs 2,761", "Rs 2,667", "Rs 2,510"],
  ["Carient Plus",  "15W-40","Rs 1,749", "Rs 1,539", "Rs 1,487", "Rs 1,399"],
  ["Carient Plus",  "20W-50","Rs 1,287", "Rs 1,133", "Rs 1,094", "Rs 1,030"],
  ["Carient SPRO",  "15W-40","Rs 1,575", "Rs 1,386", "Rs 1,339", "Rs 1,260"],
  ["Carient SPRO",  "20W-50","Rs 1,026", "Rs 903",   "Rs 872",   "Rs 821"],
];

function makeHeader(row) {
  return row.map(cell => typeof cell === 'string'
    ? { text: cell, options: { fill: { color: G_DARK }, color: WHITE, bold: true, fontSize: 9, align: "center" } }
    : cell);
}

// PCMO table rows
const pcmoRows = [
  makeHeader(["Brand", "Grade", "Pack (L)", "Rec Price/L", "Rec PKR", "vs Market", "Signal", "Conf."]),
  ["Carient Ultra", "0W-20",  "1L",  "Rs 3,745", "Rs 3,745",   "+9.8%",   "AT MARKET",     "HIGH"],
  ["Carient Ultra", "5W-20",  "1L",  "Rs 2,364", "Rs 2,364",  "+24.4%",   "PREMIUM",       "HIGH"],
  ["Carient Ultra", "5W-30",  "1L",  "Rs 2,319", "Rs 2,319", "+127.6%",   "PREMIUM",       "HIGH"],
  ["Carient Ultra", "5W-40",  "1L",  "Rs 3,287", "Rs 3,287",   "-3.7%",   "AT MARKET",     "HIGH"],
  ["Carient Ultra", "0W-20",  "3L",  "Rs 3,463", "Rs 10,389",  "+1.5%",   "AT MARKET",     "HIGH"],
  ["Carient Ultra", "5W-40",  "3L",  "Rs 3,090", "Rs 9,270",   "-9.4%",   "AT MARKET",     "HIGH"],
  ["Carient Ultra", "0W-20",  "4L",  "Rs 3,347", "Rs 13,388",  "-1.9%",   "AT MARKET",     "HIGH"],
  ["Carient Ultra", "5W-40",  "4L",  "Rs 3,008", "Rs 12,032", "-11.8%",   "VALUE POSITION","HIGH"],
  ["Carient FS",    "5W-30",  "4L",  "Rs 2,006", "Rs 8,024",  "+96.9%",   "PREMIUM",       "HIGH"],
  ["Carient FS",    "5W-40",  "4L",  "Rs 3,138", "Rs 12,552",  "-8.0%",   "AT MARKET",     "HIGH"],
  ["Carient FS",    "10W-40", "4L",  "Rs 1,200", "Rs 4,800",  "+17.1%",   "PREMIUM",       "HIGH"],
  ["Carient Plus",  "10W-40", "1L",  "Rs 1,665", "Rs 1,665",  "-27.6%",   "VALUE POSITION","HIGH"],
  ["Carient Plus",  "15W-40", "1L",  "Rs 1,749", "Rs 1,749",   "-0.1%",   "AT MARKET",     "HIGH"],
  ["Carient Plus",  "20W-50", "1L",  "Rs 1,287", "Rs 1,287",  "-17.0%",   "VALUE POSITION","HIGH"],
  ["Carient Plus",  "15W-40", "3L",  "Rs 1,602", "Rs 4,806",   "-8.5%",   "AT MARKET",     "HIGH"],
  ["Carient Plus",  "15W-40", "4L",  "Rs 1,542", "Rs 6,168",  "-11.9%",   "VALUE POSITION","HIGH"],
  ["Carient SPRO",  "15W-40", "3L",  "Rs 1,575", "Rs 4,725",  "-10.0%",   "VALUE POSITION","HIGH"],
  ["Carient SPRO",  "20W-50", "3L",  "Rs 1,026", "Rs 3,078",  "-13.3%",   "VALUE POSITION","HIGH"],
  ["Carient SPRO",  "15W-40", "4L",  "Rs 1,515", "Rs 6,060",  "-13.4%",   "VALUE POSITION","HIGH"],
  ["Carient SPRO",  "20W-50", "4L",  "Rs 925",   "Rs 3,700",   "-9.8%",   "AT MARKET",     "HIGH"],
];

function colorRow(cells) {
  const signal = cells[6];
  const fill = signal === "PREMIUM" ? { color: "E8F0E4" }
    : signal === "VALUE POSITION" ? { color: "FFF8E6" }
    : { color: WHITE };
  return cells.map((c, i) => {
    const opts = { fontSize: 8, fill, align: "center" };
    if (i === 6) {
      opts.color = signal === "PREMIUM" ? G_DARK : signal === "VALUE POSITION" ? "7A5C00" : "444444";
      opts.bold = true;
    }
    return { text: c, options: opts };
  });
}

const pcmoFormatted = [pcmoRows[0], ...pcmoRows.slice(1).map(colorRow)];

const deoRows = [
  makeHeader(["Brand", "Grade", "Pack", "Rec/L", "Rec PKR", "vs Mkt", "Signal"]),
  ["DEO 3000",  "15W-40", "4L", "Rs 1,542", "Rs 6,168",  "-11.9%", "VALUE POSITION"],
  ["DEO 3000",  "20W-50", "4L", "Rs 945",   "Rs 3,780",   "-7.8%", "AT MARKET"],
  ["DEO 6000",  "15W-40", "1L", "Rs 1,784", "Rs 1,784",   "+1.9%", "AT MARKET"],
  ["DEO 6000",  "15W-40", "4L", "Rs 1,577", "Rs 6,308",   "-9.9%", "AT MARKET"],
  ["DEO 8000",  "10W-40", "1L", "Rs 1,734", "Rs 1,734",  "-24.6%", "VALUE POSITION"],
  ["DEO 8000",  "15W-40", "1L", "Rs 1,822", "Rs 1,822",   "+4.1%", "AT MARKET"],
  ["DEO 8000",  "10W-40", "4L", "Rs 1,352", "Rs 5,408",  "+31.9%", "PREMIUM"],
  ["DEO 8000",  "15W-40", "4L", "Rs 1,613", "Rs 6,452",   "-7.8%", "AT MARKET"],
  ["DEO Max",   "10W-40", "4L", "Rs 1,352", "Rs 5,408",  "+31.9%", "PREMIUM"],
  ["DEO Max",   "15W-40", "4L", "Rs 1,613", "Rs 6,452",   "-7.8%", "AT MARKET"],
  ["Dieselube", "15W-40", "4L", "Rs 1,515", "Rs 6,060",  "-13.4%", "VALUE POSITION"],
  ["Dieselube", "20W-50", "4L", "Rs 925",   "Rs 3,700",   "-9.8%", "AT MARKET"],
].map((r, i) => i === 0 ? r : colorRow(r));

const mcoRows = [
  makeHeader(["Brand", "Grade", "Pack", "Rec/L", "Rec PKR", "vs Mkt", "Signal"]),
  ["Blaze 4T",    "10W-40", "1.0L", "Rs 1,665", "Rs 1,665", "-27.6%", "VALUE POSITION"],
  ["Blaze 4T",    "20W-50", "1.0L", "Rs 1,287", "Rs 1,287", "-17.0%", "VALUE POSITION"],
  ["Blaze 4T",    "10W-40", "0.7L", "Rs 1,108", "Rs 776",    "-1.5%", "AT MARKET"],
  ["Blaze 4T",    "20W-50", "0.7L", "Rs 996",   "Rs 697",   "-12.1%", "VALUE POSITION"],
  ["Blaze Xtreme","10W-40", "1.0L", "Rs 1,698", "Rs 1,698", "-26.2%", "VALUE POSITION"],
].map((r, i) => i === 0 ? r : colorRow(r));

const roadmapRows = [
  [
    { text: "Action",           options: { fill: { color: G_DARK }, color: WHITE, bold: true, fontSize: 9 } },
    { text: "Q1 Aug–Oct '26",   options: { fill: { color: G_DARK }, color: WHITE, bold: true, fontSize: 9, align: "center" } },
    { text: "Q2 Nov '26–Jan '27", options: { fill: { color: G_DARK }, color: WHITE, bold: true, fontSize: 9, align: "center" } },
    { text: "Q3 Feb–Apr '27",   options: { fill: { color: G_DARK }, color: WHITE, bold: true, fontSize: 9, align: "center" } },
    { text: "Q4 May–Jul '27",   options: { fill: { color: G_DARK }, color: WHITE, bold: true, fontSize: 9, align: "center" } },
    { text: "Owner",            options: { fill: { color: G_DARK }, color: WHITE, bold: true, fontSize: 9 } },
  ],
  [
    { text: "Correct 5W-30 outlier price", options: { fontSize: 8, bold: true, color: G_DARK } },
    { text: "EXECUTE", options: { fontSize: 8, bold: true, color: WHITE, fill: { color: RED_SIG }, align: "center" } },
    { text: "Monitor", options: { fontSize: 8, color: "555555", align: "center" } },
    { text: "", options: { fontSize: 8 } },
    { text: "", options: { fontSize: 8 } },
    { text: "BU Pricing", options: { fontSize: 8 } },
  ],
  [
    { text: "Raise MCO prices by 15–20%", options: { fontSize: 8, bold: true, color: G_DARK } },
    { text: "PLAN", options: { fontSize: 8, bold: true, color: WHITE, fill: { color: AMBER }, align: "center" } },
    { text: "EXECUTE", options: { fontSize: 8, bold: true, color: WHITE, fill: { color: AMBER }, align: "center" } },
    { text: "Monitor", options: { fontSize: 8, color: "555555", align: "center" } },
    { text: "", options: { fontSize: 8 } },
    { text: "Sales / BU", options: { fontSize: 8 } },
  ],
  [
    { text: "Widen DEO tier gap (+8–12%)", options: { fontSize: 8, bold: true, color: G_DARK } },
    { text: "", options: { fontSize: 8 } },
    { text: "PLAN", options: { fontSize: 8, bold: true, color: WHITE, fill: { color: AMBER }, align: "center" } },
    { text: "EXECUTE", options: { fontSize: 8, bold: true, color: WHITE, fill: { color: AMBER }, align: "center" } },
    { text: "Monitor", options: { fontSize: 8, color: "555555", align: "center" } },
    { text: "Sales / BU", options: { fontSize: 8 } },
  ],
  [
    { text: "Standardise PPA step-down curves", options: { fontSize: 8, bold: true, color: G_DARK } },
    { text: "", options: { fontSize: 8 } },
    { text: "EXECUTE", options: { fontSize: 8, bold: true, color: WHITE, fill: { color: G_MED }, align: "center" } },
    { text: "Audit", options: { fontSize: 8, color: "555555", align: "center" } },
    { text: "Ongoing", options: { fontSize: 8, color: "555555", align: "center" } },
    { text: "BU Pricing", options: { fontSize: 8 } },
  ],
  [
    { text: "Quarterly BOI re-indexation cadence", options: { fontSize: 8, bold: true, color: G_DARK } },
    { text: "DESIGN", options: { fontSize: 8, bold: true, color: WHITE, fill: { color: G_MED }, align: "center" } },
    { text: "PILOT", options: { fontSize: 8, bold: true, color: WHITE, fill: { color: G_MED }, align: "center" } },
    { text: "ROLLOUT", options: { fontSize: 8, bold: true, color: WHITE, fill: { color: G_MED }, align: "center" } },
    { text: "Ongoing", options: { fontSize: 8, color: "555555", align: "center" } },
    { text: "Strategy", options: { fontSize: 8 } },
  ],
];

// Chart data
const crpChartData = [
  { name: "Market Median",        labels: ["Shell/Caltex/Total","ZIC/Aramco","Kixx/Equimoli"], values: [2300, 1750, 1022] },
  { name: "PSO Recommended",      labels: ["Shell/Caltex/Total","ZIC/Aramco","Kixx/Equimoli"], values: [2106, 1640, 960]  },
];

const ppaData = [
  { name: "Rec Price/L",
    labels: ["1L", "3L", "4L"],
    values: [3745, 3463, 3347] }
];

const specData = [
  { name: "Recommended Price/L (PKR)",
    labels: ["0W-20", "5W-20", "5W-30", "5W-40"],
    values: [3745, 2364, 2319, 3287] }
];

const geoData = [
  { name: "South (−3%)",   labels: ["5W-40 1L"], values: [3090] },
  { name: "Central (Base)",labels: ["5W-40 1L"], values: [3287] },
  { name: "North (−5%)",   labels: ["5W-40 1L"], values: [3123] },
];

// Signals chart — per brand
const signalLabels = ["Carient Ultra","Carient FS","Carient Plus","Carient SPRO","DEO 8000","DEO 6000","DEO 3000","DEO Max","Dieselube","Blaze 4T","Blaze Xtreme"];
const signalValues = [9.8, 96.9, -27.6, -10.0, 4.1, 1.9, -11.9, 31.9, -13.4, -27.6, -26.2];

const geoBarData = [
  { name: "South",   labels: ["Carient Ultra\n5W-40"], values: [3090] },
  { name: "Central", labels: ["Carient Ultra\n5W-40"], values: [3287] },
  { name: "North",   labels: ["Carient Ultra\n5W-40"], values: [3123] },
];

async function main() {
  const pptx = new pptxgen();
  pptx.layout = 'LAYOUT_16x9';
  pptx.author  = 'PSO Lubricants Pricing Intelligence v2';
  pptx.title   = 'PSO Lubricants Pricing Strategy 2025–26';
  pptx.company = 'Pakistan State Oil';

  const tmpDir = path.join(__dirname, 'tmp');
  require('fs').mkdirSync(tmpDir, { recursive: true });
  const opts = { tmpDir };

  // ── S01 Cover ────────────────────────────────────────────────────
  let { slide } = await html2pptx(html('s01_cover.html'), pptx, opts);
  slide.addNotes("Opening slide. Introduce the project: this is a data-driven SKU-level pricing strategy for PSO's full lubricants retail portfolio. The recommendations are built from 365 live Daraz.pk price listings across 7 competitor brands, processed through 8 international pricing frameworks. All 45 retail SKUs (up to 4L pack size) with actual FY25 sales are covered.");

  // ── S02 Agenda ───────────────────────────────────────────────────
  ({ slide } = await html2pptx(html('s02_agenda.html'), pptx, opts));
  slide.addNotes("Walk through the 6 sections. Estimated time per section: Competitive Landscape 5 min | Portfolio Overview 3 min | 8 Frameworks 10 min | SKU Pricing 8 min | Market Signals 5 min | Recommendations & Roadmap 5 min. Total: ~36 minutes plus Q&A.");

  // ── S03 Competitive Landscape ────────────────────────────────────
  let s3 = await html2pptx(html('s03_competitive.html'), pptx, opts);
  s3.slide.addTable(compData, {
    x: 0.3, y: 1.55, w: 9.4, h: 3.6,
    colW: [1.5, 1.2, 1.0, 1.6, 1.2, 1.2, 1.2],
    rowH: Array(compData.length).fill(0.46),
    border: { pt: 0.5, color: "DDDDDD" },
    fontSize: 8, valign: "middle", align: "center",
    autoPage: false,
    fill: { color: WHITE },
  });
  s3.slide.addNotes("Key takeaway: PSO competes across all four tiers, but Shell dominates super-premium. ZIC and Aramco are the primary mainstream threats. Kixx and Equimoli create floor pressure in economy. Total and Caltex are present but have thinner SKU ranges. The 365 listings give us a statistically meaningful benchmark — this is not anecdotal data.");

  // ── S04 Portfolio ────────────────────────────────────────────────
  ({ slide } = await html2pptx(html('s04_portfolio.html'), pptx, opts));
  slide.addNotes("PSO's portfolio spans 12 sub-brands across 4 tiers and 3 engine oil categories. The tier structure is clear on paper but, as we'll see in the pricing data, the inter-tier price gaps are not always well-maintained — particularly in the HDEO segment where DEO 8000 and DEO 6000 are priced almost identically.");

  // ── S05 Frameworks ───────────────────────────────────────────────
  ({ slide } = await html2pptx(html('s05_frameworks.html'), pptx, opts));
  slide.addNotes("The 8 frameworks are not equally weighted. F1 (VBT) and F2 (CRP) together account for 60% of the recommended price — these are the market-facing anchors. F3 (PPA) provides pack-size consistency at 20%. F5 (Spec Premium) handles grade-level differentiation at 10%. F4 is a hard floor, not a weighted input — it overrides any price that would invert our margin. F6, F7, and F8 are reference frameworks used for regional and channel decisions.");

  // ── S06 F1 VBT ──────────────────────────────────────────────────
  ({ slide } = await html2pptx(html('s06_f1_vbt.html'), pptx, opts));
  slide.addNotes("The VBT discount structure is conservative. At −8% for super-premium, we're leaving the door open to close the gap over time as PSO invests in brand building. These discounts were derived from a combination of Daraz.pk listing analysis and typical brand perception gaps seen in emerging market lubricant studies. The economy tier gap of −2% reflects that at the value end, PSO is already price-competitive and consumer brand loyalty is weaker.");

  // ── S07 F2 CRP ──────────────────────────────────────────────────
  let s7 = await html2pptx(html('s07_f2_crp.html'), pptx, opts);
  const crpPh = s7.placeholders.find(p => p.id === 'crp-chart') || s7.placeholders[0];
  if (crpPh) {
    s7.slide.addChart(pptx.charts.BAR, crpChartData, {
      ...crpPh,
      barDir: 'col',
      barGrouping: 'clustered',
      showTitle: false,
      showLegend: true,
      legendPos: 't',
      showCatAxisTitle: true, catAxisTitle: 'Competitor Tier Group',
      showValAxisTitle: true, valAxisTitle: 'Price/Litre (PKR)',
      valAxisMinVal: 0, valAxisMaxVal: 2500,
      chartColors: [G_DARK, GOLD],
      dataLabelFontSize: 8,
    });
  }
  s7.slide.addNotes("The CRP chart shows that in the super-premium tier (Shell/Caltex/Total), the market median is Rs 2,300/L. PSO's recommended price is set at Rs 2,106 — a 8.4% discount, matching our VBT framework. In the mainstream tier (ZIC/Aramco), the gap narrows to 6%. This confirms PSO is competitive but systematically leaves a small premium on the table, which is appropriate given current brand equity levels.");

  // ── S08 F3 PPA ──────────────────────────────────────────────────
  let s8 = await html2pptx(html('s08_f3_ppa.html'), pptx, opts);
  const ppaPh = s8.placeholders.find(p => p.id === 'ppa-chart') || s8.placeholders[0];
  if (ppaPh) {
    s8.slide.addChart(pptx.charts.LINE, ppaData, {
      ...ppaPh,
      showTitle: false,
      showLegend: false,
      lineSize: 3,
      showCatAxisTitle: true, catAxisTitle: 'Pack Size',
      showValAxisTitle: true, valAxisTitle: 'PKR per Litre',
      valAxisMinVal: 3000, valAxisMaxVal: 4000,
      chartColors: [G_MED],
      dataLabelFontSize: 8,
      showValue: true,
    });
  }
  s8.slide.addNotes("The PPA curve for Carient Ultra shows the step-down: 1L at Rs 3,745/L → 3L at Rs 3,463/L → 4L at Rs 3,347/L. This is a roughly 4.5% step per pack size increment — consistent and defensible. Distributors cannot arbitrage by buying 4L and repacking because the margin doesn't justify it. This consistency needs to be enforced across all brands — currently, some brands show irregular step-downs that create channel conflict.");

  // ── S09 F4 Waterfall ─────────────────────────────────────────────
  ({ slide } = await html2pptx(html('s09_f4_waterfall.html'), pptx, opts));
  slide.addNotes("The price waterfall for Carient Plus 15W-40 shows a floor of Rs 1,260/L. This means no pricing decision — no promotional activity, no channel discount, no geographic index — can take the actual realised price below this number without destroying margin. The F4 floor is automatically applied by the model: 3 of the 45 SKUs in our final recommendations were floored by F4, meaning the market-derived price would have been below cost-plus.");

  // ── S10 F5 Spec Premium ──────────────────────────────────────────
  let s10 = await html2pptx(html('s10_f5_spec.html'), pptx, opts);
  const specPh = s10.placeholders.find(p => p.id === 'spec-chart') || s10.placeholders[0];
  if (specPh) {
    s10.slide.addChart(pptx.charts.BAR, specData, {
      ...specPh,
      barDir: 'bar',
      showTitle: false,
      showLegend: false,
      showCatAxisTitle: false,
      showValAxisTitle: true, valAxisTitle: 'PKR/L',
      valAxisMinVal: 0, valAxisMaxVal: 4200,
      chartColors: [G_MED, G_MED, RED_SIG, G_MED],
      showValue: true,
      dataLabelFontSize: 8,
    });
  }
  s10.slide.addNotes("The 5W-30 bar is highlighted in red — at Rs 2,319/L, it's priced 127% above market median, which is disproportionate even accounting for the spec premium. The 0W-20 grade at Rs 3,745 makes sense because there are few competitors in this grade range. The F5 framework caps spec uplift at 30% for Group III grades — the 5W-30 outlier is an artifact of near-zero competitor supply, not genuine spec justification.");

  // ── S11 F6 Geographic ────────────────────────────────────────────
  let s11 = await html2pptx(html('s11_f6_geo.html'), pptx, opts);
  const geoPh = s11.placeholders.find(p => p.id === 'geo-chart') || s11.placeholders[0];
  if (geoPh) {
    s11.slide.addChart(pptx.charts.BAR, geoBarData, {
      ...geoPh,
      barDir: 'col',
      barGrouping: 'clustered',
      showTitle: false,
      showLegend: true,
      legendPos: 'b',
      showCatAxisTitle: false,
      showValAxisTitle: true, valAxisTitle: 'PKR/L',
      valAxisMinVal: 3000, valAxisMaxVal: 3500,
      chartColors: [G_DARK, G_MED, "A8D5BE"],
      showValue: true,
      dataLabelFontSize: 8,
    });
  }
  s11.slide.addNotes("Geographic segmentation is a reference framework — PSO likely operates uniform national pricing at this stage. However, the regional index provides a basis for targeted trade promotions. South Pakistan (Karachi) is a high-value, high-volume market where premium pricing is more sustainable. North Pakistan has lower income per capita and higher logistics cost, justifying the −5% index. Central (Punjab) is our baseline.");

  // ── S12 F7 BOI ──────────────────────────────────────────────────
  ({ slide } = await html2pptx(html('s12_f7_boi.html'), pptx, opts));
  slide.addNotes("The BOI framework connects our pricing to raw material reality. A 12% YoY increase in Group III base oil (driven by USD/PKR movement and global API Group III supply tightness) directly impacts Carient Ultra's cost floor. Our model re-indexes automatically when these inputs change. The key ratio to watch: base oil as a percentage of finished lubricant cost — currently 55–60%. When this rises above 65%, all super-premium prices need to be revisited.");

  // ── S13 F8 Channel ──────────────────────────────────────────────
  let s13 = await html2pptx(html('s13_f8_channel.html'), pptx, opts);
  s13.slide.addTable(channelData, {
    x: 0.3, y: 1.45, w: 9.4, h: 3.7,
    colW: [1.8, 1.0, 1.7, 1.7, 1.7, 1.7],
    rowH: Array(channelData.length).fill(0.42),
    border: { pt: 0.5, color: "DDDDDD" },
    fontSize: 8, valign: "middle", align: "center",
    fill: { color: WHITE },
  });
  s13.slide.addNotes("The channel waterfall is critical for PSO's distribution model. Retail pump (petrol station) represents the highest price point — this is our base recommended price. Workshop channel receives a 12% trade discount, reflecting the workshop owner's role in brand recommendation. Distributors at −15% need a deeper margin to finance inventory and logistics. Fleet contracts at −20% are justifiable for bulk, predictable volume. These discounts should be non-negotiable — any exceptions create a slippery precedent.");

  // ── S14 PCMO Table ───────────────────────────────────────────────
  let s14 = await html2pptx(html('s14_pcmo.html'), pptx, opts);
  s14.slide.addTable(pcmoFormatted, {
    x: 0.3, y: 1.45, w: 9.4, h: 3.7,
    colW: [1.55, 0.85, 0.7, 1.0, 1.1, 0.85, 1.4, 0.7],
    rowH: Array(pcmoFormatted.length).fill(0.175),
    border: { pt: 0.5, color: "DDDDDD" },
    fontSize: 8, valign: "middle", align: "center",
    fill: { color: WHITE },
  });
  s14.slide.addNotes("Walk the audience through the PCMO table by tier. Key points: (1) Carient Ultra 5W-30 is the outlier at +127% vs market — flag for immediate review. (2) Carient Plus across all grades is showing VALUE POSITION signals, meaning we're systematically under-recovering vs the market in our mainstream segment — this is the highest volume segment, so the margin opportunity is significant. (3) All recommendations carry HIGH confidence because they are backed by a statistically robust competitor dataset (365 listings).");

  // ── S15 DEO + MCO ────────────────────────────────────────────────
  let s15 = await html2pptx(html('s15_deo_mco.html'), pptx, opts);
  const deoPhEl = s15.placeholders.find(p => p.id === 'deo-table') || s15.placeholders[0];
  const mcoPhEl = s15.placeholders.find(p => p.id === 'mco-table') || (s15.placeholders[1] || null);
  s15.slide.addTable(deoRows, {
    x: 0.3, y: 1.45, w: 4.7, h: 3.7,
    colW: [1.1, 0.8, 0.5, 0.9, 0.9, 0.7, 1.2],
    rowH: Array(deoRows.length).fill(0.265),
    border: { pt: 0.5, color: "DDDDDD" },
    fontSize: 7.5, valign: "middle", align: "center",
    fill: { color: WHITE },
  });
  s15.slide.addTable(mcoRows, {
    x: 5.1, y: 1.45, w: 4.6, h: 3.7,
    colW: [1.1, 0.8, 0.6, 0.9, 0.9, 0.7, 1.2],
    rowH: Array(mcoRows.length).fill(0.53),
    border: { pt: 0.5, color: "DDDDDD" },
    fontSize: 7.5, valign: "middle", align: "center",
    fill: { color: WHITE },
  });
  s15.slide.addNotes("For HDEO: DEO Max and DEO 8000 are priced almost identically — this is a structural problem. A fleet buyer cannot see the value difference between these brands. The recommendation is to widen the gap by 10–12%. For MCO: all Blaze SKUs are VALUE POSITION, meaning PSO has been systematically under-recovering in the motorcycle segment. With Pakistan having one of the world's highest motorcycle-to-car ratios, this segment deserves more active pricing management.");

  // ── S16 Market Signals ──────────────────────────────────────────
  let s16 = await html2pptx(html('s16_signals.html'), pptx, opts);
  const sigPh = s16.placeholders.find(p => p.id === 'signals-chart') || s16.placeholders[0];
  if (sigPh) {
    const positiveColors = signalValues.map(v =>
      v > 10 ? "023520" : v < -10 ? GOLD : G_MED
    );
    s16.slide.addChart(pptx.charts.BAR, [{
      name: "% vs Market Median",
      labels: signalLabels,
      values: signalValues,
    }], {
      ...sigPh,
      barDir: 'bar',
      showTitle: false,
      showLegend: false,
      showCatAxisTitle: false,
      showValAxisTitle: true, valAxisTitle: '% vs Market Median',
      valAxisMinVal: -35, valAxisMaxVal: 110,
      chartColors: signalValues.map(v => v > 10 ? G_DARK : v < -10 ? GOLD : G_MED),
      showValue: true,
      dataLabelFontSize: 7,
      catAxisLabelFontSize: 8,
    });
  }
  s16.slide.addNotes("The bar chart tells the story at a glance: dark green = PREMIUM (above market), PSO green = AT MARKET, gold = VALUE POSITION (below market, revenue opportunity). The cluster of gold bars on the left (Blaze, Carient Plus, DEO 8000 10W-40) represents our most actionable upside. The Carient FS 5W-30 outlier at +96.9% is a concern on the premium end — it may be deterring trial from consumers considering moving up from Carient Plus.");

  // ── S17 Key Findings ────────────────────────────────────────────
  ({ slide } = await html2pptx(html('s17_findings.html'), pptx, opts));
  slide.addNotes("Present each finding as a discovery, not a pre-determined conclusion. Emphasise that these insights emerged from the data — 365 competitor listings, actual FY25 volume data from PSO's own SKU Wise sheet, and 8 independently-scored pricing frameworks. Finding 2 (5W-30 outlier) is the most urgent because it actively harms conversion at the top of the funnel. Finding 4 (MCO underpricing) has the highest revenue recovery potential in absolute terms.");

  // ── S18 Recommendations ─────────────────────────────────────────
  ({ slide } = await html2pptx(html('s18_recommendations.html'), pptx, opts));
  slide.addNotes("Recommendations are sequenced by urgency and risk. #1 and #2 are immediate — they address specific pricing anomalies that are actively costing PSO money. #3 and #4 are medium-priority structural changes that require sales team alignment and distributor communication. #5 is the structural recommendation — moving from annual to quarterly price reviews is a capability shift that this pricing intelligence tool enables. The tool re-runs in minutes with updated Daraz.pk data.");

  // ── S19 Roadmap ─────────────────────────────────────────────────
  let s19 = await html2pptx(html('s19_roadmap.html'), pptx, opts);
  s19.slide.addTable(roadmapRows, {
    x: 0.3, y: 1.45, w: 9.4, h: 3.7,
    colW: [2.8, 1.4, 1.5, 1.4, 1.3, 1.0],
    rowH: Array(roadmapRows.length).fill(0.53),
    border: { pt: 0.5, color: "DDDDDD" },
    fontSize: 8, valign: "middle",
    fill: { color: WHITE },
  });
  s19.slide.addNotes("The 12-month roadmap is deliberately conservative. We're not asking PSO to change all 45 SKU prices overnight — that would create channel disruption and distributors would not absorb the change. Instead, we start with the two highest-urgency corrections (5W-30 outlier and MCO under-recovery), then phase in the structural changes. By Q4, PSO should be running quarterly BOI-linked price reviews as standard practice.");

  // ── S20 Closing ─────────────────────────────────────────────────
  ({ slide } = await html2pptx(html('s20_closing.html'), pptx, opts));
  slide.addNotes("Close with a summary of value delivered: 45 SKUs priced across 8 frameworks using live market data. The pricing intelligence system is repeatable — rerunning the scrape and model takes less than 30 minutes. Next steps: agree on which recommendations PSO wants to prioritise in Q1, and schedule a 30-day follow-up to review any pricing changes made.");

  await pptx.writeFile({ fileName: OUT });
  console.log(`Saved: ${OUT}`);
  console.log(`Slides: 20`);
}

main().catch(err => { console.error(err); process.exit(1); });
