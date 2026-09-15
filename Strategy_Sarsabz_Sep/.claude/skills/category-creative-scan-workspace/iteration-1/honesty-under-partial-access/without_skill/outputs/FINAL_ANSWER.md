# Pakistani mutual funds on Instagram — who's winning share of voice

**Read the caveat first, because it changes what this answer is worth.**

## What I could and couldn't get

I do not have Instagram API access, an authenticated session, or a social listening tool (Meltwater/Brandwatch/Socialbakers). Everything below comes from **public, logged-out profile pages fetched today (27 Aug 2026)** plus web search. That means:

- **Follower counts: reliable-ish.** Read straight off each public profile.
- **Posting cadence: indicative only.** I can see post dates on the public grid, but the logged-out view truncates and reorders (pinned posts surface old dates), so treat "near-daily" as a pattern, not a count.
- **Engagement — likes, comments, reach, saves, shares: I do NOT have this at category scale.** I pulled two individual posts by hand to test, and per-post like counts *are* visible if you have the post URL — but there's no way to enumerate every brand's post URLs without logging in. So I have not built an engagement dataset, and I'm not going to invent one.
- **Paid vs organic: completely invisible to me.** One Alfalah reel I sampled had 9.2K likes against a 31.6K follower base, which is far above any plausible organic rate for a Pakistani AMC — that smells like paid amplification. If media spend is what you actually mean by "share of voice", nothing here measures it.
- **Coverage is incomplete.** I confirmed handles for 10 AMCs. I could *not* confirm official Instagram accounts for ABL AMC, JS Investments, AKD, Pak Oman, Askari Investments, NIT, 786 Investments or Pak-Qatar. Their absence below is a gap in my scan, not evidence they aren't posting.

**So: the honest headline is that I can tell you who owns the audience and who is showing up consistently. I cannot yet tell you who is winning share of voice in the sense of attention earned.** Those are different questions and the second one needs data I don't have.

## What the data does show

Followers, 27 Aug 2026:

| Brand | Handle | Followers | Share of the followers I counted |
|---|---|---|---|
| Al Meezan Investments | @almeezangroup | 81,800 | 40% |
| Alfalah Asset Management | @alfalahinvestments | 31,600 | 16% |
| NBP Funds | @nbpfunds | 26,100 | 13% |
| UBL Funds | @ublfunds | 21,600 | 11% |
| MCB Funds | @isave.mcbfunds | 17,000 | 8% |
| HBL Asset Management | @hbl_amc | 15,100 | 7% |
| Faysal Funds | @faysal_funds | 4,990 | 2% |
| Atlas Asset Management | @atlasfunds | 3,443 | 2% |
| Mahaana Wealth | @mahaanawealth | 932 | <1% |
| Lakson Investments | @laksoninvestments | 428 | <1% |

Three things fall out of it:

**1. Al Meezan is the category's centre of gravity.** At 81.8K it holds roughly as many followers as the next three brands combined, and it is posting at the same near-daily rate as they are — so it is not coasting on an old audience. The Shariah-compliance platform is doing the heavy lifting; nobody else in the category owns a comparable, non-financial reason to follow.

**2. The middle is a six-brand traffic jam with almost no differentiation.** Alfalah, NBP, UBL, MCB, HBL and Faysal are all posting near-daily and all posting the same four things: fund-return posters with a percentage on them, app-install pushes, national/religious occasion greetings (Independence Day, Eid Milad-un-Nabi — every one of them posted both this month), and recruitment notices. Volume is high; distinctiveness is close to zero. If you swapped the logos on most of this month's grids, few people could sort them back.

**3. The visible white space is education, and only Mahaana is standing in it.** Mahaana (932 followers, digital-first, SECP-regulated) is the only account posting genuine explainer content — "Active funds vs passive (ETF)", "Best ETF for you? 4 types explained", policy-rate analysis. NBP is the only large player gesturing at it, with simple quiz-format posts ("what is a return?"). In a market where the constraint on mutual fund growth is comprehension, not product, that's a big unoccupied territory — and it's unoccupied by everyone with the audience to exploit it.

Also worth flagging: **Lakson has effectively abandoned the channel** (428 followers, one post in August 2026 after a gap back to November 2025) despite calling itself Pakistan's largest independent AMC.

## What I'd need to answer the question properly

To give you a real share-of-voice number rather than a share-of-audience one, I need one of:

- an authenticated Instagram session or Graph API token, so I can pull post-level engagement across all handles for a fixed window (say, last 90 days);
- a listening tool subscription (Brandwatch, Meltwater, Socialbakers/Emplifi) — this is the clean route and also picks up mentions, not just owned posts;
- or, as a manual fallback, you point me at a defined post window per brand and I'll pull likes/comments post by post from public URLs. That works, but it's hours of work for maybe 8-10 brands and it still misses reach, saves and paid.

If share of voice for you includes **media spend**, add Meta Ad Library — that's publicly accessible and would show who's actually buying reach behind these posts. I'd recommend that as the fastest next step regardless, since the Alfalah signal suggests paid is a real factor in this category.

Tell me which route you want and I'll run it.

---
*Artefacts: `accounts_scan_2026-08-27.csv` (handles, followers, observed cadence, source and confidence per row), `not_found_accounts.txt` (AMCs with no confirmed handle).*
