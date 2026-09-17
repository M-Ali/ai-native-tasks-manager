# SOFAR

The running record of the conversation that shapes this project: what was
asked, what was decided and what is still open. Add a new dated session at the
bottom and don't rewrite earlier ones.

---

## Session 1 — 2026-09-17

### What the user asked

1. "I am trying to create some sort of specialized tool for automobile
   industry, being a marketing communication agencies group we have multiple
   automobile companies including Suzuki and Hyundai; what sort of analytical
   tools can we create for them"
2. "Can you create a python app using uv, create AGENTS.md, HISTORY.md,
   README.md, and also SOFAR.md (this one keeps the project my chat details)"

### What was proposed (answer to 1)

**Settle first: competing clients.** The group handles Suzuki and Hyundai,
which compete directly. Build two layers:
- a shared layer on public data (PAMA sales, PakWheels, social, YouTube,
  search trends) that any client may see;
- a private layer for each client's own data (CRM, dealer enquiries, GA4,
  media spend), never pooled.

This is also a selling point: the group sees the whole category, which a
single-brand agency can't.

**Seven tools, ranked by value for effort:**
1. Monthly market monitor: PAMA sales plus brands that aren't PAMA members;
   share by segment, month-on-month and year-on-year. *Build first*, because
   it's cheap, used monthly, and the other tools plug into it.
2. Used-car price and "on money" tracker (PakWheels / OLX): how well each model
   holds its value, the premium paid for quick delivery, days a listing stays
   up. Resale value is the strongest purchase driver.
3. Consumer voice tracker: complaints and wants by model from reviews,
   YouTube, Facebook groups and forums (reuses the `comment-analysis` skill).
4. Share of voice vs share of market: Meta Ad Library, social, Google Trends,
   message themes and gaps (extends `category-creative-scan` /
   `competitor-comms-audit`).
5. Launch tracker: the first 90 days of each new model, compared with
   earlier launches.
6. Demand drivers model: interest rates, State Bank auto financing, exchange
   rate, fuel and price rises, to separate economic effects from marketing
   effects.
7. Leads to sales, one client at a time: GA4, lead forms, test drives,
   dealer CRM.

**Suggested phases:** Phase 1 = tools 1–3 as one monthly "Auto Category Pulse";
Phase 2 = tools 4–5; Phase 3 = tools 6–7 as paid analytics work.

**Risk noted:** scraping PakWheels and OLX raises terms-of-service and blocking
issues. Check before tool 2 becomes something clients rely on.

### What was built (answer to 2)

- uv app `autopulse` in `D:\personal\BHT_Sep26` with `validate` and `monitor`
  (tool 1). Details are in HISTORY.md.
- The user didn't pick a starting tool, so the build started with the market
  monitor, as the proposal recommended.

### Open questions for the user

- Which clients are in scope besides Suzuki and Hyundai?
- Is past PAMA or model-level sales data available? No real data is in the
  project yet.
- Can either client share its own data (GA4, leads, dealer CRM) for tool 7?
- Output format: xlsx only, or also a monthly PPTX / Word report per client?
- Should tool 2's scraping go ahead once the terms of service are checked?
