# Where to collect, and what goes wrong in each source

Ranked by evidentiary value. Prefer sources that are **dated, complete and re-checkable** — an
evaluator or client may ask where a number came from, and "we looked around" is not an answer.

## Contents

- [Ad transparency libraries](#ad-transparency-libraries) — the spine
- [Video platforms](#video-platforms)
- [Press](#press)
- [Television](#television)
- [Out of home](#out-of-home)
- [Owned assets and the purchase funnel](#owned-assets-and-the-purchase-funnel)
- [Earned and sentiment](#earned-and-sentiment)
- [Market share data](#market-share-data)
- [Provenance discipline](#provenance-discipline)

## Ad transparency libraries

**Meta Ad Library** (`facebook.com/ads/library`) is the backbone of most audits. Searchable by
advertiser, shows every currently active ad with its launch date, all creative variants, and
the platforms it runs on.

What it gives you: creative volume per brand, launch dates, variant counts, targeting country.

What it does not: **spend, for commercial advertisers.** Spend and impression ranges appear
only for political and social-issue ads in some jurisdictions. Do not extrapolate spend from
what you can see here.

Gotchas:
- Brands run under multiple page names — search variants, local-language spellings, and
  subsidiary or Takaful-window entities separately, or you will undercount a competitor badly.
- The library shows *active* ads. Ads that stopped before you looked are invisible, which biases
  toward whatever is running the week you collect. State your collection date.
- Variant inflation: one concept may appear as thirty near-identical ads. Decide up front
  whether the row is the *concept* or the *asset* and apply it consistently. Concept-level is
  usually the more meaningful unit; note the variant count in a column instead.

**Google Ads Transparency Center** covers Search, Display and YouTube by advertiser. Same
strengths and same spend limitation.

**TikTok Creative Center** and **LinkedIn ad libraries** are worth checking where the category
uses them — LinkedIn especially for B2B, corporate and SME propositions.

## Video platforms

Brand YouTube channels give TVCs and long-form with **upload dates and view counts**. View
counts are a legitimate, citable engagement measure — unlike spend, they are published.

Sort by date for the audit window and by popularity to see what the brand itself pushed. Note
where paid promotion has obviously inflated views; flag rather than exclude.

## Press

In markets where print still carries category weight, press cannot be skipped — and in
procurement contexts that gate on press-release accreditation, evaluators expect it covered.

Sources: newspaper e-papers for the window, agency release records, and the client's own
monitoring if they share it. Code the same way as digital; print skews heavily toward
`brand_corporate` and `heritage_scale`, which is itself a finding.

## Television

Proper TV monitoring is bought, not scraped — in Pakistan that means Medialogic or Gallup
Pakistan TAM. If you have access, use it for genuine share-of-voice weighting.

If you do not: say so explicitly, use YouTube uploads and press as the proxy, and label it a
proxy. An honest limitation stated up front is far stronger than an invented GRP table that
falls apart under one question.

## Out of home

Photograph what you see; there is rarely a database. Treat OOH as qualitative colour on
message territory, not as a countable channel, unless the client supplies site data.

## Owned assets and the purchase funnel

Underrated, and often where the sharpest findings are. For each competitor, actually attempt
the journey: website, quote or lead form, app store listing, WhatsApp or call-centre entry
point.

Record how many steps to a quote, whether pricing is visible, whether the flow completes on
mobile, what happens after submission. In categories where advertising promises ease and the
funnel does not deliver it, that gap is a strategic wedge — and it is evidence nobody else in
the pitch will have bothered to gather.

## Earned and sentiment

Complaint patterns, regulator records, review sites and social replies show where category
trust actually breaks. In protection and financial categories the recurring theme is usually
settlement — whether the company pays when it matters.

This is the raw material for the **claim-to-reality gap**: what the advertising promises set
against what customers report receiving.

Handle with care. Report patterns and cite sources; do not repeat unverified allegations about
a named competitor in a client document.

## Market share data

Needed for the share-of-voice-vs-share-of-market comparison. Use primary sources and cite them
per figure in `market_share.csv`:

- The insurance or sector regulator's published industry statistics (in Pakistan, SECP)
- Company annual reports and financial statements
- Rating agency reports (PACRA, VIS) which often summarise sector position
- Industry association publications

Do not use a number you cannot cite. One unsourced share figure invites the client's commercial
team to dismiss the whole section — and they will know the real figure.

## Provenance discipline

Every row carries a `source_url`. For anything transient — an ad library entry, a social post —
screenshot it into a dated folder alongside the workbook. Ads come down; your evidence should
not disappear with them.

Record the **collection date** on the sheet's README. An audit is a snapshot, and saying so is a
strength, not a hedge.
