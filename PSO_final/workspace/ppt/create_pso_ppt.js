/**
 * PSO OMC Analytics â€” PowerPoint Report Generator
 * 10 Months FY26 | Pakistan State Oil
 *
 * CONSTRAINT: Borders/backgrounds/shadows only on <div>, not on <p>/<h1>/etc.
 */

const fs   = require('fs');
const path = require('path');
const NPM  = 'C:/Users/Muhamad.Ali/AppData/Roaming/npm/node_modules';
const pptxgen   = require(path.join(NPM, 'pptxgenjs'));
const html2pptx = require('D:/Personal/SkillsApril2026/claude-code-skills-lab-main/.claude/skills/pptx/scripts/html2pptx.js');

const WS     = 'D:/Personal/PSO_final/workspace/ppt';
const OUTPUT = 'D:/Personal/PSO_final/reports/PSO_Report_10M_FY26.pptx';

function w(name, html) {
  const fp = path.join(WS, `${name}.html`);
  fs.writeFileSync(fp, html, 'utf8');
  return fp;
}

// Shared header
const HDR = (sub) => `
  <div style="background:#1B2A4A;padding:12pt 24pt 10pt;display:flex;align-items:center;justify-content:space-between;">
    <p style="color:#C9A030;font-size:11pt;font-weight:bold;margin:0;">PSO OMC ANALYTICS</p>
    <p style="color:#9BB0C8;font-size:9pt;margin:0;">${sub} &nbsp;|&nbsp; 10M FY26</p>
  </div>`;

// Section title: border on wrapping div, text in inner p
const ST = (txt, color='#C9A030') =>
  `<div style="border-bottom:2pt solid ${color};padding-bottom:3pt;margin-bottom:8pt;">
     <p style="color:#1B2A4A;font-size:11pt;font-weight:bold;margin:0;">${txt}</p>
   </div>`;

// Numbered finding row
const FIND = (n, html) =>
  `<div style="display:flex;align-items:flex-start;gap:6pt;margin-bottom:6pt;">
     <div style="background:#1B2A4A;min-width:16pt;height:16pt;border-radius:2pt;display:flex;align-items:center;justify-content:center;">
       <p style="color:#C9A030;font-size:9pt;font-weight:bold;margin:0;">${n}</p>
     </div>
     <p style="color:#2C3E50;font-size:9pt;line-height:1.35;margin:0;">${html}</p>
   </div>`;

// â”€â”€ SLIDE 01 â€” Title â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
const s01 = `<!DOCTYPE html><html><head><style>
  html{background:#1B2A4A;}
  body{width:720pt;height:405pt;margin:0;padding:0;font-family:Arial,sans-serif;
       display:flex;flex-direction:column;background:#1B2A4A;}
</style></head><body>
  <div style="height:6pt;background:#C9A030;width:100%;"></div>
  <div style="flex:1;display:flex;flex-direction:column;justify-content:center;padding:0 60pt;">
    <p style="color:#9BB0C8;font-size:10pt;letter-spacing:2pt;margin:0 0 8pt;">PAKISTAN STATE OIL â€” RETAIL BUSINESS DIVISION</p>
    <p style="color:#C9A030;font-size:42pt;font-weight:bold;margin:0 0 6pt;line-height:1.1;">PSO OMC Analytics</p>
    <p style="color:#DDEEFF;font-size:16pt;margin:0 0 20pt;">Performance &amp; Strategic Insight Report</p>
    <div style="width:60pt;height:3pt;background:#C9A030;margin:0 0 16pt;"></div>
    <p style="color:#9BB0C8;font-size:10pt;margin:0;">Period: 10 Months FY26 &nbsp;|&nbsp; Data: 27,344 transactions &nbsp;|&nbsp; 21,334 Retail rows</p>
  </div>
  <div style="padding:12pt 24pt;display:flex;justify-content:space-between;align-items:center;border-top:1px solid #2D4060;">
    <p style="color:#9BB0C8;font-size:9pt;margin:0;">CONFIDENTIAL â€” For Management Use Only</p>
    <p style="color:#C9A030;font-size:9pt;margin:0;">June 2026</p>
  </div>
</body></html>`;

// â”€â”€ SLIDE 02 â€” Portfolio â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
const s02 = `<!DOCTYPE html><html><head><style>
  html{background:#F5F7FA;}
  body{width:720pt;height:405pt;margin:0;padding:0;font-family:Arial,sans-serif;
       display:flex;flex-direction:column;background:#F5F7FA;}
</style></head><body>
  ${HDR('Portfolio Overview')}
  <div style="flex:1;display:flex;padding:12pt 20pt 8pt;gap:12pt;">
    <div style="flex:0 0 42%;display:flex;flex-direction:column;gap:10pt;">
      <div style="background:#FFFFFF;border-radius:4pt;padding:10pt 12pt;border-left:4pt solid #C9A030;">
        <p style="color:#1B2A4A;font-size:20pt;font-weight:bold;margin:0;">PKR +23.6%</p>
        <p style="color:#6B7D8A;font-size:8pt;margin:0;">GROSS REVENUE (GRS) YoY</p>
        <p style="color:#E07B39;font-size:9pt;font-weight:bold;margin:2pt 0 0;">Price inflation driving revenue â€” volume is down</p>
      </div>
      <div style="background:#FFFFFF;border-radius:4pt;padding:10pt 12pt;border-left:4pt solid #B22222;">
        <p style="color:#B22222;font-size:20pt;font-weight:bold;margin:0;">-1.85%</p>
        <p style="color:#6B7D8A;font-size:8pt;margin:0;">TOTAL VOLUME YoY</p>
        <p style="color:#B22222;font-size:9pt;font-weight:bold;margin:2pt 0 0;">Real market demand is shrinking</p>
      </div>
      <div style="background:#FFFFFF;border-radius:4pt;padding:10pt 12pt;border-left:4pt solid #B22222;">
        <p style="color:#B22222;font-size:20pt;font-weight:bold;margin:0;">-4.86%</p>
        <p style="color:#6B7D8A;font-size:8pt;margin:0;">NET MARGIN YoY</p>
        <p style="color:#B22222;font-size:9pt;font-weight:bold;margin:2pt 0 0;">NMgn/Ltr: PKR 7.573 â†’ 7.342</p>
      </div>
      <div style="background:#FFFFFF;border-radius:4pt;padding:10pt 12pt;border-left:4pt solid #E07B39;">
        <p style="color:#E07B39;font-size:20pt;font-weight:bold;margin:0;">-5.67%</p>
        <p style="color:#6B7D8A;font-size:8pt;margin:0;">LUBE NET MARGIN / LITRE YoY</p>
        <p style="color:#E07B39;font-size:9pt;font-weight:bold;margin:2pt 0 0;">239.5 â†’ 225.9 PKR/ltr despite vol growth</p>
      </div>
    </div>
    <div style="flex:1;background:#FFFFFF;border-radius:4pt;padding:10pt;border:1px solid #D8E0E8;">
      <p style="color:#6B7D8A;font-size:8pt;margin:0 0 4pt;font-weight:bold;">GRS % CHANGE BY CHANNEL â€” CY vs LY</p>
      <div id="chart1" class="placeholder" class="placeholder" style="width:100%;height:210pt;background:#EEF2F7;border-radius:3pt;"></div>
    </div>
  </div>
</body></html>`;

// â”€â”€ SLIDE 03 â€” Retail Segments â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
const s03 = `<!DOCTYPE html><html><head><style>
  html{background:#F5F7FA;}
  body{width:720pt;height:405pt;margin:0;padding:0;font-family:Arial,sans-serif;
       display:flex;flex-direction:column;background:#F5F7FA;}
</style></head><body>
  ${HDR('Retail Business â€” Segment Split')}
  <div style="flex:1;display:flex;padding:12pt 20pt 8pt;gap:14pt;">
    <div style="flex:0 0 38%;display:flex;flex-direction:column;">
      ${ST('RETAIL SEGMENT KPIs')}
      <div style="background:#FFFFFF;border-radius:4pt;padding:10pt 12pt;margin-bottom:8pt;border-left:4pt solid #2E5B9A;">
        <p style="color:#6B7D8A;font-size:8pt;font-weight:bold;margin:0 0 2pt;">PETROL (PMG + R95)</p>
        <p style="color:#1B2A4A;font-size:14pt;font-weight:bold;margin:0;">3,469 ML</p>
        <p style="color:#B22222;font-size:9pt;margin:2pt 0 0;">Vol: -1.33% &nbsp;|&nbsp; NMgn: -7.88%</p>
        <p style="color:#2C3E50;font-size:9pt;margin:1pt 0 0;">NMgn/ltr: 7.555 â†’ <span style="color:#B22222;font-weight:bold;">7.054</span></p>
      </div>
      <div style="background:#FFFFFF;border-radius:4pt;padding:10pt 12pt;margin-bottom:8pt;border-left:4pt solid #E07B39;">
        <p style="color:#6B7D8A;font-size:8pt;font-weight:bold;margin:0 0 2pt;">DIESEL (HSD + LDO)</p>
        <p style="color:#1B2A4A;font-size:14pt;font-weight:bold;margin:0;">2,603 ML</p>
        <p style="color:#B22222;font-size:9pt;margin:2pt 0 0;">Vol: -2.58% &nbsp;|&nbsp; NMgn: -1.33%</p>
        <p style="color:#2C3E50;font-size:9pt;margin:1pt 0 0;">NMgn/ltr: 6.524 â†’ <span style="color:#1A7A44;font-weight:bold;">6.608</span></p>
      </div>
      <div style="background:#FFFFFF;border-radius:4pt;padding:10pt 12pt;margin-bottom:8pt;border-left:4pt solid #C9A030;">
        <p style="color:#6B7D8A;font-size:8pt;font-weight:bold;margin:0 0 2pt;">LUBRICANTS (ALL LUBES)</p>
        <p style="color:#1B2A4A;font-size:14pt;font-weight:bold;margin:0;">13.3 ML</p>
        <p style="color:#1A7A44;font-size:9pt;margin:2pt 0 0;">Vol: +7.60% &nbsp;|&nbsp; NMgn: +1.51%</p>
        <p style="color:#B22222;font-size:9pt;margin:1pt 0 0;">NMgn/ltr: 239.5 â†’ <span style="font-weight:bold;">225.9</span> (mix erosion)</p>
      </div>
    </div>
    <div style="flex:1;">
      <p style="color:#6B7D8A;font-size:8pt;margin:0 0 4pt;font-weight:bold;">RETAIL VOLUME CY vs LY BY SEGMENT (Million Litres)</p>
      <div id="chart1" class="placeholder" class="placeholder" style="width:100%;height:240pt;background:#EEF2F7;border-radius:3pt;"></div>
    </div>
  </div>
</body></html>`;

// â”€â”€ SLIDE 04 â€” Diesel â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
const s04 = `<!DOCTYPE html><html><head><style>
  html{background:#F5F7FA;}
  body{width:720pt;height:405pt;margin:0;padding:0;font-family:Arial,sans-serif;
       display:flex;flex-direction:column;background:#F5F7FA;}
</style></head><body>
  ${HDR('Diesel Analysis')}
  <div style="flex:1;display:flex;padding:12pt 20pt 8pt;gap:14pt;">
    <div style="flex:0 0 40%;display:flex;flex-direction:column;">
      ${ST('KEY FINDINGS')}
      ${FIND(1,'<b>Central leads but declining:</b> 1,156 ML, <span style="color:#B22222;font-weight:bold;">-4.71% YoY</span> â€” largest region losing volume fastest.')}
      ${FIND(2,'<b>North: only bright spot:</b> 942 ML, <span style="color:#1A7A44;font-weight:bold;">+1.68%</span> â€” the single growing diesel region.')}
      ${FIND(3,'<b>South in retreat:</b> 505 ML, <span style="color:#B22222;font-weight:bold;">-5.16%</span> â€” just 8.3% of national diesel share.')}
      ${FIND(4,'<b>Top 24 cities = 51%</b> of total diesel volume. Karachi, Lahore, Multan, Rawalpindi dominate.')}
      ${FIND(5,'<b>NMgn/ltr pressure:</b> North worst â€” 6.668 â†’ <span style="color:#B22222;font-weight:bold;">6.187 (-7.2%)</span>. City hotspots: Okara -20%, Rawalpindi -16.7%.')}
      <div style="background:#FEF9EC;border-left:3pt solid #C9A030;padding:6pt 8pt;border-radius:2pt;">
        <p style="color:#5D4E00;font-size:8pt;margin:0;">South has the HIGHEST diesel NMgn/ltr (7.466) despite the steepest volume decline â€” better pricing discipline.</p>
      </div>
    </div>
    <div style="flex:1;">
      <p style="color:#6B7D8A;font-size:8pt;margin:0 0 4pt;font-weight:bold;">DIESEL VOLUME BY REGION â€” CY vs LY (Million Litres)</p>
      <div id="chart1" class="placeholder" class="placeholder" style="width:100%;height:240pt;background:#EEF2F7;border-radius:3pt;"></div>
    </div>
  </div>
</body></html>`;

// â”€â”€ SLIDE 05 â€” Petrol â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
const s05 = `<!DOCTYPE html><html><head><style>
  html{background:#F5F7FA;}
  body{width:720pt;height:405pt;margin:0;padding:0;font-family:Arial,sans-serif;
       display:flex;flex-direction:column;background:#F5F7FA;}
</style></head><body>
  ${HDR('Petrol Analysis')}
  <div style="flex:1;display:flex;padding:12pt 20pt 8pt;gap:14pt;">
    <div style="flex:0 0 40%;display:flex;flex-direction:column;">
      ${ST('KEY FINDINGS')}
      ${FIND(1,'<b>ALL regions losing NMgn/ltr</b> despite GRS up 20-28%. Revenue growth = price inflation only.')}
      ${FIND(2,'<b>Central:</b> <span style="color:#B22222;font-weight:bold;">-2.5% vol, -10.5% NMgn</span>. NMgn/ltr 7.319 â†’ 6.723. Largest loss: PKR 1.23B.')}
      ${FIND(3,'<b>North:</b> +2.2% vol but NMgn/ltr 7.471 â†’ <span style="color:#B22222;font-weight:bold;">6.622 (-11.4%)</span>. Growing volume at margin cost.')}
      ${FIND(4,'<b>South:</b> -3.7% vol but highest NMgn/ltr at <span style="color:#1A7A44;font-weight:bold;">7.973</span> â€” minimal discounting (PKR 0.28/ltr vs 1.20 Central).')}
      ${FIND(5,'<b>R95 margin collapse:</b> Volume growing +7-11% but NMgn/ltr falling 14-27%. Zero discounts â€” primary margin erosion.')}
      <div style="background:#FFF0F0;border-left:3pt solid #B22222;padding:6pt 8pt;border-radius:2pt;">
        <p style="color:#7A0000;font-size:8pt;margin:0;"><b>Central Punjab cluster:</b> Sahiwal -23.5%, Kasur -16.7%, Sheikhupura -14.8%, Gujranwala -10.9%, Rawalpindi -9.8%. Combined ~57 ML lost.</p>
      </div>
    </div>
    <div style="flex:1;">
      <p style="color:#6B7D8A;font-size:8pt;margin:0 0 4pt;font-weight:bold;">PETROL NET MARGIN / LITRE BY REGION â€” CY vs LY (PKR)</p>
      <div id="chart1" class="placeholder" class="placeholder" style="width:100%;height:240pt;background:#EEF2F7;border-radius:3pt;"></div>
    </div>
  </div>
</body></html>`;

// â”€â”€ SLIDE 06 â€” R95 Critical Signal â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
const s06 = `<!DOCTYPE html><html><head><style>
  html{background:#F5F7FA;}
  body{width:720pt;height:405pt;margin:0;padding:0;font-family:Arial,sans-serif;
       display:flex;flex-direction:column;background:#F5F7FA;}
</style></head><body>
  <div style="background:#B22222;padding:12pt 24pt 10pt;display:flex;align-items:center;justify-content:space-between;">
    <p style="color:#FFFFFF;font-size:11pt;font-weight:bold;margin:0;">PSO OMC ANALYTICS â€” CRITICAL SIGNAL</p>
    <p style="color:#FFAAAA;font-size:9pt;margin:0;">R95 Premium Petrol â€” Growing Volume, Collapsing Margin &nbsp;|&nbsp; 10M FY26</p>
  </div>
  <div style="flex:1;display:flex;padding:12pt 20pt 8pt;gap:14pt;">
    <div style="flex:0 0 38%;display:flex;flex-direction:column;">
      <div style="background:#1B2A4A;border-radius:4pt;padding:10pt;margin-bottom:8pt;text-align:center;">
        <p style="color:#B22222;font-size:26pt;font-weight:bold;margin:0;">-27%</p>
        <p style="color:#9BB0C8;font-size:8pt;margin:0;">R95 NMgn/Ltr DECLINE â€” NORTH REGION</p>
      </div>
      <div style="background:#FFFFFF;border-radius:4pt;padding:8pt;margin-bottom:6pt;border-left:3pt solid #B22222;">
        <p style="color:#6B7D8A;font-size:8pt;margin:0 0 2pt;">CENTRAL R95 NMgn/Ltr</p>
        <p style="color:#B22222;font-size:11pt;font-weight:bold;margin:0;">9.487 <span style="color:#6B7D8A;font-size:9pt;font-weight:normal;">vs 12.493 LY</span></p>
      </div>
      <div style="background:#FFFFFF;border-radius:4pt;padding:8pt;margin-bottom:6pt;border-left:3pt solid #B22222;">
        <p style="color:#6B7D8A;font-size:8pt;margin:0 0 2pt;">NORTH R95 NMgn/Ltr</p>
        <p style="color:#B22222;font-size:11pt;font-weight:bold;margin:0;">7.966 <span style="color:#6B7D8A;font-size:9pt;font-weight:normal;">vs 10.885 LY</span></p>
      </div>
      <div style="background:#FFFFFF;border-radius:4pt;padding:8pt;margin-bottom:6pt;border-left:3pt solid #B22222;">
        <p style="color:#6B7D8A;font-size:8pt;margin:0 0 2pt;">SOUTH R95 NMgn/Ltr</p>
        <p style="color:#B22222;font-size:11pt;font-weight:bold;margin:0;">17.360 <span style="color:#6B7D8A;font-size:9pt;font-weight:normal;">vs 20.102 LY</span></p>
      </div>
      <div style="background:#FFF0F0;border-left:3pt solid #B22222;padding:6pt 8pt;border-radius:2pt;">
        <p style="color:#7A0000;font-size:8pt;margin:0;">Zero discounts on R95 â€” this is PRIMARY MARGIN compression. Volume is growing (+7-11%) while the product becomes less profitable. Pricing or cost review urgently needed.</p>
      </div>
    </div>
    <div style="flex:1;">
      <p style="color:#6B7D8A;font-size:8pt;margin:0 0 4pt;font-weight:bold;">R95 NET MARGIN / LITRE BY REGION â€” CY vs LY (PKR)</p>
      <div id="chart1" class="placeholder" class="placeholder" style="width:100%;height:240pt;background:#EEF2F7;border-radius:3pt;"></div>
    </div>
  </div>
</body></html>`;

// â”€â”€ SLIDE 07 â€” Central Punjab Cluster â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
const s07 = `<!DOCTYPE html><html><head><style>
  html{background:#F5F7FA;}
  body{width:720pt;height:405pt;margin:0;padding:0;font-family:Arial,sans-serif;
       display:flex;flex-direction:column;background:#F5F7FA;}
</style></head><body>
  ${HDR('Central Punjab â€” Petrol Volume Decline Cluster')}
  <div style="flex:1;display:flex;padding:12pt 20pt 8pt;gap:14pt;">
    <div style="flex:0 0 40%;display:flex;flex-direction:column;">
      <div style="background:#B22222;border-radius:3pt;padding:6pt 10pt;margin-bottom:6pt;">
        <p style="color:#FFFFFF;font-size:9pt;margin:0;font-weight:bold;">5 Adjacent Cities â€” Combined 57 ML Volume Loss</p>
      </div>
      ${ST('CITY-LEVEL DECLINES')}
      <div style="background:#FFFFFF;border-radius:3pt;padding:6pt 10pt;margin-bottom:4pt;border-left:3pt solid #B22222;">
        <p style="color:#6B7D8A;font-size:7pt;font-weight:bold;margin:0;">SAHIWAL</p>
        <p style="color:#1B2A4A;font-size:10pt;font-weight:bold;margin:0;">35.4 ML <span style="color:#B22222;">â†“ from 46.3 ML</span></p>
        <p style="color:#B22222;font-size:8pt;margin:0;">-23.5% YoY</p>
      </div>
      <div style="background:#FFFFFF;border-radius:3pt;padding:6pt 10pt;margin-bottom:4pt;border-left:3pt solid #B22222;">
        <p style="color:#6B7D8A;font-size:7pt;font-weight:bold;margin:0;">KASUR</p>
        <p style="color:#1B2A4A;font-size:10pt;font-weight:bold;margin:0;">35.4 ML <span style="color:#B22222;">â†“ from 42.5 ML</span></p>
        <p style="color:#B22222;font-size:8pt;margin:0;">-16.7% YoY</p>
      </div>
      <div style="background:#FFFFFF;border-radius:3pt;padding:6pt 10pt;margin-bottom:4pt;border-left:3pt solid #B22222;">
        <p style="color:#6B7D8A;font-size:7pt;font-weight:bold;margin:0;">SHEIKHUPURA</p>
        <p style="color:#1B2A4A;font-size:10pt;font-weight:bold;margin:0;">35.0 ML <span style="color:#B22222;">â†“ from 41.0 ML</span></p>
        <p style="color:#B22222;font-size:8pt;margin:0;">-14.8% YoY</p>
      </div>
      <div style="background:#FFFFFF;border-radius:3pt;padding:6pt 10pt;margin-bottom:4pt;border-left:3pt solid #B22222;">
        <p style="color:#6B7D8A;font-size:7pt;font-weight:bold;margin:0;">GUJRANWALA</p>
        <p style="color:#1B2A4A;font-size:10pt;font-weight:bold;margin:0;">90.9 ML <span style="color:#B22222;">â†“ from 102.0 ML</span></p>
        <p style="color:#B22222;font-size:8pt;margin:0;">-10.9% YoY</p>
      </div>
      <div style="background:#FFFFFF;border-radius:3pt;padding:6pt 10pt;margin-bottom:4pt;border-left:3pt solid #B22222;">
        <p style="color:#6B7D8A;font-size:7pt;font-weight:bold;margin:0;">RAWALPINDI</p>
        <p style="color:#1B2A4A;font-size:10pt;font-weight:bold;margin:0;">122.8 ML <span style="color:#B22222;">â†“ from 136.1 ML</span></p>
        <p style="color:#B22222;font-size:8pt;margin:0;">-9.8% YoY</p>
      </div>
    </div>
    <div style="flex:1;">
      <p style="color:#6B7D8A;font-size:8pt;margin:0 0 4pt;font-weight:bold;">PETROL VOLUME â€” CENTRAL PUNJAB CITIES (CY vs LY, Million Litres)</p>
      <div id="chart1" class="placeholder" class="placeholder" style="width:100%;height:245pt;background:#EEF2F7;border-radius:3pt;"></div>
    </div>
  </div>
</body></html>`;

// â”€â”€ SLIDE 08 â€” Lubes Mix Problem â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
const s08 = `<!DOCTYPE html><html><head><style>
  html{background:#F5F7FA;}
  body{width:720pt;height:405pt;margin:0;padding:0;font-family:Arial,sans-serif;
       display:flex;flex-direction:column;background:#F5F7FA;}
</style></head><body>
  ${HDR('Lubricants â€” The Mix Shift Problem')}
  <div style="flex:1;display:flex;padding:12pt 20pt 8pt;gap:14pt;">
    <div style="flex:0 0 42%;display:flex;flex-direction:column;">
      <div style="background:#C9A030;border-radius:3pt;padding:8pt 10pt;margin-bottom:8pt;">
        <p style="color:#1B2A4A;font-size:12pt;font-weight:bold;margin:0;">Total Volume UP +7.6% â€” But Wrong Products</p>
        <p style="color:#4A3500;font-size:8pt;margin:2pt 0 0;">LOW GRADE surge is replacing premium categories</p>
      </div>
      <div style="background:#FFFFFF;border-radius:3pt;padding:6pt 10pt;margin-bottom:4pt;border-left:3pt solid #B22222;">
        <p style="color:#1B2A4A;font-size:9pt;font-weight:bold;margin:0;">LOW GRADE <span style="color:#6B7D8A;font-weight:normal;font-size:8pt;">7.204 ML CY</span></p>
        <p style="color:#1A7A44;font-size:9pt;font-weight:bold;margin:2pt 0 0;">+24.4% YoY &nbsp;|&nbsp; NMgn/ltr: PKR 116</p>
      </div>
      <div style="background:#FFFFFF;border-radius:3pt;padding:6pt 10pt;margin-bottom:4pt;border-left:3pt solid #2E5B9A;">
        <p style="color:#1B2A4A;font-size:9pt;font-weight:bold;margin:0;">DEO <span style="color:#6B7D8A;font-weight:normal;font-size:8pt;">3.895 ML CY</span></p>
        <p style="color:#B22222;font-size:9pt;font-weight:bold;margin:2pt 0 0;">-6.7% YoY &nbsp;|&nbsp; NMgn/ltr: PKR 348</p>
      </div>
      <div style="background:#FFFFFF;border-radius:3pt;padding:6pt 10pt;margin-bottom:4pt;border-left:3pt solid #E07B39;">
        <p style="color:#1B2A4A;font-size:9pt;font-weight:bold;margin:0;">MCO <span style="color:#6B7D8A;font-weight:normal;font-size:8pt;">1.218 ML CY</span></p>
        <p style="color:#B22222;font-size:9pt;font-weight:bold;margin:2pt 0 0;">-11.1% YoY &nbsp;|&nbsp; NMgn/ltr: PKR 269</p>
      </div>
      <div style="background:#FFFFFF;border-radius:3pt;padding:6pt 10pt;margin-bottom:4pt;border-left:3pt solid #1A7A44;">
        <p style="color:#1B2A4A;font-size:9pt;font-weight:bold;margin:0;">PCMO <span style="color:#6B7D8A;font-weight:normal;font-size:8pt;">0.954 ML CY</span></p>
        <p style="color:#B22222;font-size:9pt;font-weight:bold;margin:2pt 0 0;">-4.7% YoY &nbsp;|&nbsp; NMgn/ltr: PKR 501</p>
      </div>
      <div style="background:#FFF0F0;border-left:3pt solid #B22222;padding:6pt 8pt;border-radius:2pt;margin-top:6pt;">
        <p style="color:#7A0000;font-size:8pt;margin:0;"><b>Zero discounts across ALL categories.</b> This is a product mix / demand structure problem. PSO is following the market toward low-margin products.</p>
      </div>
    </div>
    <div style="flex:1;">
      <p style="color:#6B7D8A;font-size:8pt;margin:0 0 4pt;font-weight:bold;">LUBE CATEGORY VOLUMES â€” CY vs LY (Million Litres)</p>
      <div id="chart1" class="placeholder" class="placeholder" style="width:100%;height:240pt;background:#EEF2F7;border-radius:3pt;"></div>
    </div>
  </div>
</body></html>`;

// â”€â”€ SLIDE 09 â€” Lubes Margin â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
const s09 = `<!DOCTYPE html><html><head><style>
  html{background:#F5F7FA;}
  body{width:720pt;height:405pt;margin:0;padding:0;font-family:Arial,sans-serif;
       display:flex;flex-direction:column;background:#F5F7FA;}
</style></head><body>
  ${HDR('Lubricants â€” Why Mix Matters')}
  <div style="flex:1;display:flex;padding:12pt 20pt 8pt;gap:14pt;">
    <div style="flex:0 0 38%;display:flex;flex-direction:column;">
      <div style="background:#1B2A4A;border-radius:4pt;padding:10pt;margin-bottom:8pt;text-align:center;">
        <p style="color:#C9A030;font-size:22pt;font-weight:bold;margin:0;">PKR 232M</p>
        <p style="color:#9BB0C8;font-size:8pt;margin:0;">NET MARGIN LOST per ML shifted from DEO to LOW GRADE</p>
      </div>
      <p style="color:#1B2A4A;font-size:9pt;font-weight:bold;margin:0 0 4pt;">NET MARGIN / LITRE BY CATEGORY (CY)</p>
      <div style="background:#FFFFFF;border-radius:3pt;padding:6pt 10pt;margin-bottom:4pt;display:flex;justify-content:space-between;align-items:center;">
        <p style="color:#1B2A4A;font-size:9pt;font-weight:bold;margin:0;">PCMO</p>
        <p style="color:#2E5B9A;font-size:11pt;font-weight:bold;margin:0;">PKR 500.9</p>
      </div>
      <div style="background:#FFFFFF;border-radius:3pt;padding:6pt 10pt;margin-bottom:4pt;display:flex;justify-content:space-between;align-items:center;">
        <p style="color:#1B2A4A;font-size:9pt;font-weight:bold;margin:0;">DEO</p>
        <p style="color:#2E5B9A;font-size:11pt;font-weight:bold;margin:0;">PKR 348.1</p>
      </div>
      <div style="background:#FFFFFF;border-radius:3pt;padding:6pt 10pt;margin-bottom:4pt;display:flex;justify-content:space-between;align-items:center;">
        <p style="color:#1B2A4A;font-size:9pt;font-weight:bold;margin:0;">MCO</p>
        <p style="color:#2E5B9A;font-size:11pt;font-weight:bold;margin:0;">PKR 268.9</p>
      </div>
      <div style="background:#FFF0F0;border-radius:3pt;padding:6pt 10pt;margin-bottom:4pt;display:flex;justify-content:space-between;align-items:center;">
        <p style="color:#B22222;font-size:9pt;font-weight:bold;margin:0;">LOW GRADE</p>
        <p style="color:#B22222;font-size:11pt;font-weight:bold;margin:0;">PKR 116.0</p>
      </div>
      <div style="background:#FEF9EC;border-left:3pt solid #C9A030;padding:6pt 8pt;border-radius:2pt;margin-top:6pt;">
        <p style="color:#4A3500;font-size:8pt;margin:0;">PCMO generates <b>4.3x</b> more margin/ltr than LOW GRADE. Yet PCMO total volume is only 0.95 ML â€” the smallest premium category.</p>
      </div>
    </div>
    <div style="flex:1;">
      <p style="color:#6B7D8A;font-size:8pt;margin:0 0 4pt;font-weight:bold;">NET MARGIN / LITRE BY LUBE CATEGORY â€” CY (PKR per Litre)</p>
      <div id="chart1" class="placeholder" class="placeholder" style="width:100%;height:240pt;background:#EEF2F7;border-radius:3pt;"></div>
    </div>
  </div>
</body></html>`;

// â”€â”€ SLIDE 10 â€” Regional Matrix (div grid, no table) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
const ROW = (region, seg, vol, volChg, volBg, nmgnChg, nmgnBg, ltrCY, ltrLY, sig, sigBg) =>
  `<div style="display:flex;align-items:center;padding:4pt 6pt;border-bottom:1pt solid #E8EEF4;font-size:8pt;">
    <div style="flex:0 0 70pt;"><p style="font-weight:bold;color:#1B2A4A;font-size:8pt;margin:0;">${region}</p></div>
    <div style="flex:0 0 55pt;"><p style="color:#6B7D8A;font-size:8pt;margin:0;">${seg}</p></div>
    <div style="flex:0 0 60pt;text-align:center;"><p style="color:#2C3E50;font-size:8pt;margin:0;">${vol}</p></div>
    <div style="flex:0 0 55pt;padding:0 2pt;"><div style="background:${volBg};border-radius:2pt;padding:2pt 4pt;text-align:center;"><p style="font-weight:bold;font-size:8pt;margin:0;">${volChg}</p></div></div>
    <div style="flex:0 0 55pt;padding:0 2pt;"><div style="background:${nmgnBg};border-radius:2pt;padding:2pt 4pt;text-align:center;"><p style="font-weight:bold;font-size:8pt;margin:0;">${nmgnChg}</p></div></div>
    <div style="flex:0 0 55pt;text-align:center;"><p style="color:#2C3E50;font-size:8pt;margin:0;">${ltrCY}</p></div>
    <div style="flex:0 0 55pt;text-align:center;"><p style="color:#2C3E50;font-size:8pt;margin:0;">${ltrLY}</p></div>
    <div style="flex:1;padding:0 2pt;"><div style="background:${sigBg};border-radius:2pt;padding:2pt 4pt;text-align:center;"><p style="font-weight:bold;font-size:8pt;margin:0;">${sig}</p></div></div>
   </div>`;

const s10 = `<!DOCTYPE html><html><head><style>
  html{background:#F5F7FA;}
  body{width:720pt;height:405pt;margin:0;padding:0;font-family:Arial,sans-serif;
       display:flex;flex-direction:column;background:#F5F7FA;}
</style></head><body>
  ${HDR('Regional Performance Matrix')}
  <div style="flex:1;display:flex;flex-direction:column;padding:10pt 20pt 8pt;">
    ${ST('REGION Ã— SEGMENT â€” VOLUME & MARGIN CHANGES YoY')}
    <div style="display:flex;gap:12pt;margin-bottom:6pt;">
      <div style="display:flex;align-items:center;gap:4pt;"><div style="width:10pt;height:10pt;border-radius:50%;background:#FDECEA;"></div><p style="color:#6B7D8A;font-size:8pt;margin:0;">Both declining</p></div>
      <div style="display:flex;align-items:center;gap:4pt;"><div style="width:10pt;height:10pt;border-radius:50%;background:#FEF9EC;"></div><p style="color:#6B7D8A;font-size:8pt;margin:0;">Mixed signal</p></div>
      <div style="display:flex;align-items:center;gap:4pt;"><div style="width:10pt;height:10pt;border-radius:50%;background:#E8F8EE;"></div><p style="color:#6B7D8A;font-size:8pt;margin:0;">Both growing</p></div>
    </div>
    <div style="background:#1B2A4A;display:flex;padding:4pt 6pt;border-radius:3pt 3pt 0 0;">
      <div style="flex:0 0 70pt;"><p style="color:#C9A030;font-size:8pt;font-weight:bold;margin:0;">Region</p></div>
      <div style="flex:0 0 55pt;"><p style="color:#C9A030;font-size:8pt;font-weight:bold;margin:0;">Segment</p></div>
      <div style="flex:0 0 60pt;text-align:center;"><p style="color:#9BB0C8;font-size:8pt;margin:0;">Vol CY (ML)</p></div>
      <div style="flex:0 0 55pt;text-align:center;"><p style="color:#9BB0C8;font-size:8pt;margin:0;">Vol Chg%</p></div>
      <div style="flex:0 0 55pt;text-align:center;"><p style="color:#9BB0C8;font-size:8pt;margin:0;">NMgn Chg%</p></div>
      <div style="flex:0 0 55pt;text-align:center;"><p style="color:#9BB0C8;font-size:8pt;margin:0;">NMgn/Ltr CY</p></div>
      <div style="flex:0 0 55pt;text-align:center;"><p style="color:#9BB0C8;font-size:8pt;margin:0;">NMgn/Ltr LY</p></div>
      <div style="flex:1;text-align:center;"><p style="color:#9BB0C8;font-size:8pt;margin:0;">Signal</p></div>
    </div>
    ${ROW('Central','Petrol','1,557.9','-2.5%','#FDECEA','-10.5%','#FDECEA','6.72','7.32','CONCERN','#FDECEA')}
    ${ROW('Central','Diesel','1,156.1','-4.7%','#FDECEA','-0.7%','#FEF9EC','6.41','6.15','CONCERN','#FDECEA')}
    ${ROW('Central','Lubes','6.3','+10.6%','#E8F8EE','+0.7%','#FEF9EC','198','218','MIX ISSUE','#FEF9EC')}
    ${ROW('North','Petrol','1,109.9','+2.2%','#E8F8EE','-9.4%','#FDECEA','6.62','7.47','MARGIN RISK','#FEF9EC')}
    ${ROW('North','Diesel','941.9','+1.7%','#E8F8EE','-5.7%','#FDECEA','6.19','6.67','MARGIN RISK','#FEF9EC')}
    ${ROW('North','Lubes','3.4','+4.9%','#E8F8EE','+3.3%','#E8F8EE','261','265','HEALTHY','#E8F8EE')}
    ${ROW('South','Petrol','801.5','-3.7%','#FDECEA','-5.4%','#FDECEA','7.97','8.12','CONCERN','#FDECEA')}
    ${ROW('South','Diesel','505.1','-5.2%','#FDECEA','-0.7%','#FEF9EC','7.47','7.13','CONCERN','#FDECEA')}
    ${ROW('South','Lubes','3.6','+5.1%','#E8F8EE','+0.9%','#FEF9EC','241','252','MIX ISSUE','#FEF9EC')}
  </div>
</body></html>`;

// â”€â”€ SLIDE 11 â€” City Concentration â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
const s11 = `<!DOCTYPE html><html><head><style>
  html{background:#F5F7FA;}
  body{width:720pt;height:405pt;margin:0;padding:0;font-family:Arial,sans-serif;
       display:flex;flex-direction:column;background:#F5F7FA;}
</style></head><body>
  ${HDR('City Concentration â€” Pareto Analysis')}
  <div style="flex:1;display:flex;padding:12pt 20pt 8pt;gap:14pt;">
    <div style="flex:0 0 36%;display:flex;flex-direction:column;">
      <div style="background:#1B2A4A;border-radius:4pt;padding:10pt;margin-bottom:8pt;text-align:center;">
        <p style="color:#C9A030;font-size:28pt;font-weight:bold;margin:0;">17</p>
        <p style="color:#9BB0C8;font-size:8pt;margin:0;">CITIES = 50.6% OF TOTAL RETAIL VOLUME</p>
      </div>
      <div style="background:#FFFFFF;border-radius:3pt;padding:6pt 10pt;margin-bottom:5pt;border-left:3pt solid #2E5B9A;">
        <p style="color:#6B7D8A;font-size:7pt;font-weight:bold;margin:0;">TOP 2 CITIES ALONE</p>
        <p style="color:#1B2A4A;font-size:9pt;margin:2pt 0 0;">Karachi (10.6%) + Lahore (9.1%) = <b>19.7%</b> of national retail</p>
      </div>
      <div style="background:#FFFFFF;border-radius:3pt;padding:6pt 10pt;margin-bottom:5pt;border-left:3pt solid #2E5B9A;">
        <p style="color:#6B7D8A;font-size:7pt;font-weight:bold;margin:0;">TOP 30 CITIES</p>
        <p style="color:#1B2A4A;font-size:9pt;margin:2pt 0 0;">Cover <b>63.6%</b> of retail volume. ~400+ cities = 36.4%.</p>
      </div>
      <div style="background:#FFFFFF;border-radius:3pt;padding:6pt 10pt;margin-bottom:5pt;border-left:3pt solid #B22222;">
        <p style="color:#6B7D8A;font-size:7pt;font-weight:bold;margin:0;">DECLINING TOP CITIES</p>
        <p style="color:#1B2A4A;font-size:9pt;margin:2pt 0 0;"><span style="color:#B22222;font-weight:bold;">Rawalpindi -11.2%</span>, Bahawalpur -15.7%, Okara -15.8%, Sahiwal -18.3%</p>
      </div>
      <div style="background:#FFFFFF;border-radius:3pt;padding:6pt 10pt;margin-bottom:5pt;border-left:3pt solid #1A7A44;">
        <p style="color:#6B7D8A;font-size:7pt;font-weight:bold;margin:0;">GROWING TOP CITIES</p>
        <p style="color:#1B2A4A;font-size:9pt;margin:2pt 0 0;">Karachi +0.7%, Lahore +2.9%, Peshawar +2.9%, Gujrat +7.6%</p>
      </div>
    </div>
    <div style="flex:1;">
      <p style="color:#6B7D8A;font-size:8pt;margin:0 0 4pt;font-weight:bold;">TOP 12 CITIES BY TOTAL RETAIL VOLUME (Million Litres CY)</p>
      <div id="chart1" class="placeholder" class="placeholder" style="width:100%;height:250pt;background:#EEF2F7;border-radius:3pt;"></div>
    </div>
  </div>
</body></html>`;

// â”€â”€ SLIDE 12 â€” Strategic Implications â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
const s12 = `<!DOCTYPE html><html><head><style>
  html{background:#F5F7FA;}
  body{width:720pt;height:405pt;margin:0;padding:0;font-family:Arial,sans-serif;
       display:flex;flex-direction:column;background:#F5F7FA;}
  .ic{flex:1;background:#FFFFFF;border-radius:4pt;padding:14pt;display:flex;flex-direction:column;}
  ul{padding-left:14pt;margin:0;}
  li{color:#2C3E50;font-size:8.5pt;margin-bottom:5pt;line-height:1.4;}
</style></head><body>
  ${HDR('Strategic Implications')}
  <div style="flex:1;display:flex;gap:12pt;padding:10pt 20pt 8pt;">
    <div class="ic">
      <div style="border-bottom:2pt solid #B22222;padding-bottom:4pt;margin-bottom:8pt;">
        <p style="color:#1B2A4A;font-size:11pt;font-weight:bold;margin:0;">THE REAL PROBLEM</p>
      </div>
      <ul>
        <li>Revenue growth (+23.6%) is entirely <span style="color:#B22222;font-weight:bold;">price inflation</span> â€” real volume is shrinking</li>
        <li>Lubricant issue is a <span style="color:#B22222;font-weight:bold;">mix shift</span>, not volume loss: LOW GRADE +24.4% while DEO/PCMO/MCO all decline</li>
        <li>R95 margin has <span style="color:#B22222;font-weight:bold;">collapsed 24-27%</span> with zero discounts â€” cost or pricing structure problem</li>
        <li>Central Punjab: <span style="color:#B22222;font-weight:bold;">5 cities lost ~57 ML</span> petrol volume â€” possibly outlet loss or competitive entry</li>
      </ul>
    </div>
    <div class="ic">
      <div style="border-bottom:2pt solid #C9A030;padding-bottom:4pt;margin-bottom:8pt;">
        <p style="color:#1B2A4A;font-size:11pt;font-weight:bold;margin:0;">THE OPPORTUNITY</p>
      </div>
      <ul>
        <li>South's <span style="color:#C9A030;font-weight:bold;">low-discount model</span> generates highest NMgn/ltr (7.97 Petrol, 7.47 Diesel). Replicable elsewhere.</li>
        <li>Islamabad's balanced lube mix (DEO â‰ˆ PCMO â‰ˆ LOW GRADE) can be exported to <span style="color:#C9A030;font-weight:bold;">Lahore and Faisalabad</span></li>
        <li>PCMO: PKR 501/ltr margin, only 0.95 ML volume. <span style="color:#C9A030;font-weight:bold;">4.3x</span> more profitable than LOW GRADE</li>
        <li>South PCMO NMgn/ltr is PKR 518 â€” highest margin pocket in the country</li>
      </ul>
    </div>
    <div class="ic">
      <div style="border-bottom:2pt solid #1A7A44;padding-bottom:4pt;margin-bottom:8pt;">
        <p style="color:#1B2A4A;font-size:11pt;font-weight:bold;margin:0;">THE HIDDEN RISK</p>
      </div>
      <ul>
        <li>North is the <span style="color:#1A7A44;font-weight:bold;">only volume-growing</span> region (Petrol +2.2%, Diesel +1.7%)</li>
        <li>But North NMgn/ltr is <span style="color:#B22222;font-weight:bold;">declining fastest</span>: Diesel 6.668â†’6.187, Petrol 7.471â†’6.622</li>
        <li>North diesel discount is PKR <span style="color:#B22222;font-weight:bold;">1.683/ltr</span> â€” highest nationally â€” for only +1.7% volume growth</li>
        <li>Risk: Volume growth is being <span style="color:#B22222;font-weight:bold;">bought at margin cost</span> and may not be sustainable</li>
      </ul>
    </div>
  </div>
</body></html>`;

// â”€â”€ SLIDE 13 â€” Priority Actions â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
const s13 = `<!DOCTYPE html><html><head><style>
  html{background:#F5F7FA;}
  body{width:720pt;height:405pt;margin:0;padding:0;font-family:Arial,sans-serif;
       display:flex;flex-direction:column;background:#F5F7FA;}
  .action{flex:1;background:#FFFFFF;border-radius:4pt;padding:10pt 12pt;border-top:3pt solid #B22222;}
  .p2i{flex:1;background:#F5F7FA;border-radius:3pt;padding:7pt 9pt;border-left:3pt solid #E07B39;}
</style></head><body>
  <div style="background:#B22222;padding:12pt 24pt 10pt;display:flex;align-items:center;justify-content:space-between;">
    <p style="color:#FFFFFF;font-size:11pt;font-weight:bold;margin:0;">PRIORITIZED ACTION LIST</p>
    <p style="color:#FFAAAA;font-size:9pt;margin:0;">Priority 1 â€” Act This Quarter &nbsp;|&nbsp; 10M FY26</p>
  </div>
  <div style="flex:1;display:flex;flex-direction:column;padding:10pt 20pt 8pt;">
    <p style="color:#B22222;font-size:10pt;font-weight:bold;margin:0 0 8pt;text-transform:uppercase;letter-spacing:1pt;">Priority 1 â€” Immediate (This Quarter)</p>
    <div style="display:flex;gap:10pt;margin-bottom:10pt;">
      <div class="action">
        <p style="color:#B22222;font-size:18pt;font-weight:bold;margin:0;">1</p>
        <p style="color:#1B2A4A;font-size:9pt;font-weight:bold;margin:4pt 0;">Stop LOW GRADE Lube Mix Erosion</p>
        <p style="color:#5D6D7E;font-size:8pt;margin:0 0 3pt;"><b>Problem:</b> LOW GRADE 54.2% of lube vol at PKR 116/ltr. Each ML shifted from DEO costs PSO ~PKR 232M margin.</p>
        <p style="color:#2C3E50;font-size:8pt;margin:0 0 3pt;"><b>Action:</b> Set LOW GRADE volume ceiling. Incentivise DEO/PCMO/MCO sales in Central region first.</p>
        <div style="background:#FEF9EC;border-radius:2pt;padding:3pt 5pt;margin-top:3pt;">
          <p style="color:#7A5800;font-size:7pt;margin:0;">KPI: Premium lube share â€” halt decline below 45.8%</p>
        </div>
      </div>
      <div class="action">
        <p style="color:#B22222;font-size:18pt;font-weight:bold;margin:0;">2</p>
        <p style="color:#1B2A4A;font-size:9pt;font-weight:bold;margin:4pt 0;">Arrest Central Petrol NMgn Collapse</p>
        <p style="color:#5D6D7E;font-size:8pt;margin:0 0 3pt;"><b>Problem:</b> Central Petrol NMgn fell PKR 1.23B YoY (-10.5%). Largest single-region loss in the portfolio.</p>
        <p style="color:#2C3E50;font-size:8pt;margin:0 0 3pt;"><b>Action:</b> Audit Central petrol dealer discounts. Benchmark vs South model (PKR 0.28/ltr â†’ 7.97 NMgn).</p>
        <div style="background:#FEF9EC;border-radius:2pt;padding:3pt 5pt;margin-top:3pt;">
          <p style="color:#7A5800;font-size:7pt;margin:0;">KPI: Central Petrol NMgn/ltr direction toward 7.0+</p>
        </div>
      </div>
      <div class="action">
        <p style="color:#B22222;font-size:18pt;font-weight:bold;margin:0;">3</p>
        <p style="color:#1B2A4A;font-size:9pt;font-weight:bold;margin:4pt 0;">Fix R95 Pricing</p>
        <p style="color:#5D6D7E;font-size:8pt;margin:0 0 3pt;"><b>Problem:</b> R95 NMgn/ltr collapsed 24-27% with zero discounts. Primary margin destruction at scale.</p>
        <p style="color:#2C3E50;font-size:8pt;margin:0 0 3pt;"><b>Action:</b> Review R95 ex-depot pricing and cost inputs. Volume growing (+7-11%) â€” pricing correction feasible.</p>
        <div style="background:#FEF9EC;border-radius:2pt;padding:3pt 5pt;margin-top:3pt;">
          <p style="color:#7A5800;font-size:7pt;margin:0;">KPI: R95 NMgn/ltr recovery (Central &gt;12, North &gt;10)</p>
        </div>
      </div>
    </div>
    <p style="color:#1B2A4A;font-size:10pt;font-weight:bold;margin:0 0 6pt;text-transform:uppercase;letter-spacing:1pt;">Priority 2 â€” This Half-Year</p>
    <div style="display:flex;gap:10pt;">
      <div class="p2i"><p style="color:#1B2A4A;font-size:8pt;font-weight:bold;margin:0 0 2pt;">Recover Central Punjab Cluster</p><p style="color:#5D6D7E;font-size:7.5pt;margin:0;">5 cities lost ~100 ML fuel volume. Outlet audit + competitor mapping for Rawalpindi, Bahawalpur, Okara, Sahiwal, Kasur.</p></div>
      <div class="p2i"><p style="color:#1B2A4A;font-size:8pt;font-weight:bold;margin:0 0 2pt;">Review North Diesel Discount ROI</p><p style="color:#5D6D7E;font-size:7.5pt;margin:0;">PKR 1.683/ltr discount for only +1.7% volume growth. South model at PKR 0.40/ltr is the benchmark.</p></div>
      <div class="p2i"><p style="color:#1B2A4A;font-size:8pt;font-weight:bold;margin:0 0 2pt;">Protect South Margin Premium</p><p style="color:#5D6D7E;font-size:7.5pt;margin:0;">South has highest NMgn/ltr (Petrol 7.97, Diesel 7.47) at minimal discount. Do NOT import discount culture from North.</p></div>
    </div>
  </div>
</body></html>`;

// â”€â”€ SLIDE 14 â€” Strategic Next Steps â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
const s14 = `<!DOCTYPE html><html><head><style>
  html{background:#F5F7FA;}
  body{width:720pt;height:405pt;margin:0;padding:0;font-family:Arial,sans-serif;
       display:flex;flex-direction:column;background:#F5F7FA;}
  .si{background:#FFFFFF;border-radius:3pt;padding:8pt 10pt;margin-bottom:6pt;border-left:3pt solid #C9A030;}
  .ni{background:#F5F7FA;border-radius:3pt;padding:7pt 9pt;margin-bottom:5pt;display:flex;gap:8pt;align-items:flex-start;}
</style></head><body>
  <div style="background:#1A7A44;padding:12pt 24pt 10pt;display:flex;align-items:center;justify-content:space-between;">
    <p style="color:#FFFFFF;font-size:11pt;font-weight:bold;margin:0;">STRATEGIC ACTIONS &amp; NEXT STEPS</p>
    <p style="color:#AAE0BB;font-size:9pt;margin:0;">Priority 3 â€” Next FY &nbsp;|&nbsp; 10M FY26</p>
  </div>
  <div style="flex:1;display:flex;gap:14pt;padding:10pt 20pt 8pt;">
    <div style="flex:1;display:flex;flex-direction:column;">
      <div style="border-bottom:2pt solid #1A7A44;padding-bottom:3pt;margin-bottom:8pt;">
        <p style="color:#1A7A44;font-size:10pt;font-weight:bold;margin:0;text-transform:uppercase;letter-spacing:1pt;">Priority 3 â€” Strategic (Next FY)</p>
      </div>
      <div class="si">
        <p style="color:#1B2A4A;font-size:9pt;font-weight:bold;margin:0 0 2pt;">Upgrade Lube Premium Mix: Lahore, Faisalabad, Multan</p>
        <p style="color:#5D6D7E;font-size:8pt;margin:0;">Islamabad model (DEO â‰ˆ LOW GRADE â‰ˆ PCMO) vs Lahore (LOW GRADE 3.2x over DEO). Trade marketing, mechanic engagement, consumer education.</p>
        <p style="color:#1A7A44;font-size:7pt;font-weight:bold;margin:3pt 0 0;">KPI: DEO share in these 3 cities from ~17% â†’ 25%+ by FY27</p>
      </div>
      <div class="si">
        <p style="color:#1B2A4A;font-size:9pt;font-weight:bold;margin:0 0 2pt;">Expand South PCMO â€” Highest Margin Pocket</p>
        <p style="color:#5D6D7E;font-size:8pt;margin:0;">South PCMO NMgn/ltr = PKR 518 (best in country) at only 0.342 ML. Most underleveraged high-margin combination in the dataset.</p>
        <p style="color:#1A7A44;font-size:7pt;font-weight:bold;margin:3pt 0 0;">KPI: South PCMO volume +30% next FY</p>
      </div>
      <div class="si">
        <p style="color:#1B2A4A;font-size:9pt;font-weight:bold;margin:0 0 2pt;">Understand the North Growth Model</p>
        <p style="color:#5D6D7E;font-size:8pt;margin:0;">Only region growing fuel volumes â€” but losing NMgn/ltr fastest. Determine if growth is sustainable before replicating in Central/South.</p>
        <p style="color:#1A7A44;font-size:7pt;font-weight:bold;margin:3pt 0 0;">KPI: North NMgn/ltr stabilisation while holding volume positive</p>
      </div>
    </div>
    <div style="flex:1;display:flex;flex-direction:column;">
      <div style="border-bottom:2pt solid #1A7A44;padding-bottom:3pt;margin-bottom:8pt;">
        <p style="color:#1A7A44;font-size:10pt;font-weight:bold;margin:0;text-transform:uppercase;letter-spacing:1pt;">Next Steps</p>
      </div>
      <div class="ni"><div style="background:#1B2A4A;min-width:18pt;height:18pt;border-radius:3pt;display:flex;align-items:center;justify-content:center;"><p style="color:#C9A030;font-size:9pt;font-weight:bold;margin:0;">1</p></div><p style="color:#2C3E50;font-size:8.5pt;line-height:1.3;margin:0;"><b>Outlet-level audit</b> in Central Punjab 5-city cluster: verify active PSO outlet count vs LY and map competitor activity</p></div>
      <div class="ni"><div style="background:#1B2A4A;min-width:18pt;height:18pt;border-radius:3pt;display:flex;align-items:center;justify-content:center;"><p style="color:#C9A030;font-size:9pt;font-weight:bold;margin:0;">2</p></div><p style="color:#2C3E50;font-size:8.5pt;line-height:1.3;margin:0;"><b>R95 cost structure review</b> with pricing team â€” identify root cause of 24-27% primary margin collapse</p></div>
      <div class="ni"><div style="background:#1B2A4A;min-width:18pt;height:18pt;border-radius:3pt;display:flex;align-items:center;justify-content:center;"><p style="color:#C9A030;font-size:9pt;font-weight:bold;margin:0;">3</p></div><p style="color:#2C3E50;font-size:8.5pt;line-height:1.3;margin:0;"><b>Lube sales incentive redesign</b> â€” move from volume-based to margin-mix-weighted targets in Central region</p></div>
      <div class="ni"><div style="background:#1B2A4A;min-width:18pt;height:18pt;border-radius:3pt;display:flex;align-items:center;justify-content:center;"><p style="color:#C9A030;font-size:9pt;font-weight:bold;margin:0;">4</p></div><p style="color:#2C3E50;font-size:8.5pt;line-height:1.3;margin:0;"><b>Add population data layer</b> â€” match city volumes against household/vehicle data to size the lube premium opportunity</p></div>
      <div class="ni"><div style="background:#1B2A4A;min-width:18pt;height:18pt;border-radius:3pt;display:flex;align-items:center;justify-content:center;"><p style="color:#C9A030;font-size:9pt;font-weight:bold;margin:0;">5</p></div><p style="color:#2C3E50;font-size:8.5pt;line-height:1.3;margin:0;"><b>Monthly refresh</b> â€” pipeline ready; set AI_PROVIDER=anthropic with API key for automated monthly runs</p></div>
    </div>
  </div>
</body></html>`;

// â”€â”€ Chart Data â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
const channelGRS   = { labels:['Chemicals','Power Proj.','Marine','PR Strat.','Aviation','Head Office','Key Accts','Retail','LNG','Agency'], values:[424.7,113.0,55.6,52.5,39.0,43.4,34.8,23.6,-11.8,-9.8] };
const retailSegVol = { labels:['Petrol','Diesel','Lubes'], cy:[3469.3,2602.0,13.3], ly:[3517.2,2672.1,12.4] };
const dieselRegVol = { labels:['Central','North','South'], cy:[1156.1,941.9,505.1], ly:[1213.3,926.3,532.5] };
const petrolNMgn   = { labels:['Central','North','South'], cy:[6.723,6.622,7.973], ly:[7.319,7.471,8.119] };
const r95NMgn      = { labels:['Central','North','South'], cy:[9.487,7.966,17.360], ly:[12.493,10.885,20.102] };
const punjabCities = { labels:['Rawalpindi','Gujranwala','Sheikhupura','Kasur','Sahiwal'], cy:[122.8,90.9,35.0,35.4,35.4], ly:[136.1,102.0,41.0,42.5,46.3] };
const lubeCatVol   = { labels:['LOW GRADE','DEO','MCO','PCMO','OTHERS'], cy:[7.204,3.895,1.218,0.954,0.027], ly:[5.790,4.175,1.369,1.001,0.026] };
const lubeMgn      = { labels:['PCMO','DEO','MCO','LOW GRADE'], values:[500.9,348.1,268.9,116.0], colors:['1A7A44','2E5B9A','E07B39','B22222'] };
const topCities    = { labels:['Karachi','Lahore','Multan','Rawalpindi','Faisalabad','Islamabad','Peshawar','Gujranwala','Bahawalpur','Gujrat','Sialkot','R.Y.Khan'], values:[644.9,555.4,197.9,188.2,188.1,186.1,182.5,130.4,112.6,105.9,100.1,89.3] };

// â”€â”€ Main â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
async function main() {
  const pptx = new pptxgen();
  pptx.layout = 'LAYOUT_16x9';
  pptx.title  = 'PSO OMC Analytics â€” 10M FY26';
  pptx.author = 'PSO Analytics Pipeline';

  const slides = [
    ['slide01-title',s01],['slide02-metrics',s02],['slide03-retail',s03],
    ['slide04-diesel',s04],['slide05-petrol',s05],['slide06-r95',s06],
    ['slide07-punjab',s07],['slide08-lubes-mix',s08],['slide09-lubes-mgn',s09],
    ['slide10-matrix',s10],['slide11-cities',s11],['slide12-implications',s12],
    ['slide13-actions',s13],['slide14-next',s14],
  ];

  const results = [];
  for (const [name, html] of slides) {
    const fp = w(name, html);
    const r  = await html2pptx(fp, pptx, { tmpDir: WS });
    results.push(r);
    process.stdout.write(`  slide ${name} done\n`);
  }

  const addChart = (slideResult, chartType, data, opts) => {
    const ph = slideResult.placeholders[0];
    if (!ph) { console.log('  (no placeholder)'); return; }
    slideResult.slide.addChart(chartType, data, { ...ph, ...opts });
  };

  addChart(results[1], pptx.charts.BAR, [{
    name:'GRS Chg %', labels:channelGRS.labels, values:channelGRS.values
  }], { barDir:'bar', showTitle:false, showLegend:false, showValAxisTitle:true,
        valAxisTitle:'GRS % Change YoY',
        chartColors:channelGRS.values.map(v=>v>=0?'2E5B9A':'B22222'),
        valAxisMinVal:-20, valAxisMaxVal:450,
        showValue:true, dataLabelColor:'333333', dataLabelFontSize:7 });

  addChart(results[2], pptx.charts.BAR, [
    { name:'CY', labels:retailSegVol.labels, values:retailSegVol.cy },
    { name:'LY', labels:retailSegVol.labels, values:retailSegVol.ly }
  ], { barDir:'col', barGrouping:'clustered', showTitle:false, showLegend:true, legendPos:'b',
       showValAxisTitle:true, valAxisTitle:'Volume (ML)', chartColors:['2E5B9A','D8E0E8'],
       showValue:true, dataLabelColor:'333333', dataLabelFontSize:7, valAxisMinVal:0 });

  addChart(results[3], pptx.charts.BAR, [
    { name:'CY', labels:dieselRegVol.labels, values:dieselRegVol.cy },
    { name:'LY', labels:dieselRegVol.labels, values:dieselRegVol.ly }
  ], { barDir:'col', barGrouping:'clustered', showTitle:false, showLegend:true, legendPos:'b',
       showValAxisTitle:true, valAxisTitle:'Volume (ML)', chartColors:['E07B39','D8E0E8'],
       showValue:true, dataLabelColor:'333333', dataLabelFontSize:8,
       valAxisMinVal:0, valAxisMaxVal:1400 });

  addChart(results[4], pptx.charts.BAR, [
    { name:'CY NMgn/Ltr', labels:petrolNMgn.labels, values:petrolNMgn.cy },
    { name:'LY NMgn/Ltr', labels:petrolNMgn.labels, values:petrolNMgn.ly }
  ], { barDir:'col', barGrouping:'clustered', showTitle:false, showLegend:true, legendPos:'b',
       showValAxisTitle:true, valAxisTitle:'NMgn/Ltr (PKR)', chartColors:['B22222','D8E0E8'],
       showValue:true, dataLabelColor:'333333', dataLabelFontSize:8,
       valAxisMinVal:5.5, valAxisMaxVal:9.0, valAxisMajorUnit:0.5 });

  addChart(results[5], pptx.charts.BAR, [
    { name:'CY NMgn/Ltr', labels:r95NMgn.labels, values:r95NMgn.cy },
    { name:'LY NMgn/Ltr', labels:r95NMgn.labels, values:r95NMgn.ly }
  ], { barDir:'col', barGrouping:'clustered', showTitle:false, showLegend:true, legendPos:'b',
       showValAxisTitle:true, valAxisTitle:'R95 NMgn/Ltr (PKR)', chartColors:['B22222','FFAAAA'],
       showValue:true, dataLabelColor:'333333', dataLabelFontSize:8,
       valAxisMinVal:0, valAxisMaxVal:25 });

  addChart(results[6], pptx.charts.BAR, [
    { name:'CY Volume', labels:punjabCities.labels, values:punjabCities.cy },
    { name:'LY Volume', labels:punjabCities.labels, values:punjabCities.ly }
  ], { barDir:'bar', barGrouping:'clustered', showTitle:false, showLegend:true, legendPos:'b',
       showValAxisTitle:true, valAxisTitle:'Volume (ML)', chartColors:['B22222','FFCCCC'],
       showValue:true, dataLabelColor:'333333', dataLabelFontSize:8,
       valAxisMinVal:0, valAxisMaxVal:160 });

  addChart(results[7], pptx.charts.BAR, [
    { name:'CY Volume', labels:lubeCatVol.labels, values:lubeCatVol.cy },
    { name:'LY Volume', labels:lubeCatVol.labels, values:lubeCatVol.ly }
  ], { barDir:'col', barGrouping:'clustered', showTitle:false, showLegend:true, legendPos:'b',
       showValAxisTitle:true, valAxisTitle:'Volume (ML)', chartColors:['C9A030','D8E0E8'],
       showValue:true, dataLabelColor:'333333', dataLabelFontSize:7,
       valAxisMinVal:0, valAxisMaxVal:8 });

  addChart(results[8], pptx.charts.BAR, [{
    name:'NMgn/Ltr (PKR)', labels:lubeMgn.labels, values:lubeMgn.values
  }], { barDir:'col', showTitle:false, showLegend:false,
        showValAxisTitle:true, valAxisTitle:'NMgn/Ltr (PKR)', chartColors:lubeMgn.colors,
        showValue:true, dataLabelColor:'333333', dataLabelFontSize:8,
        valAxisMinVal:0, valAxisMaxVal:600 });

  addChart(results[10], pptx.charts.BAR, [{
    name:'Retail Volume (ML)', labels:topCities.labels, values:topCities.values
  }], { barDir:'bar', showTitle:false, showLegend:false,
        showValAxisTitle:true, valAxisTitle:'Volume (ML)', chartColors:['2E5B9A'],
        showValue:true, dataLabelColor:'333333', dataLabelFontSize:7,
        valAxisMinVal:0, valAxisMaxVal:700 });

  await pptx.writeFile({ fileName: OUTPUT });
  console.log(`\nSaved: ${OUTPUT}`);
}

main().catch(e => { console.error(e); process.exit(1); });

