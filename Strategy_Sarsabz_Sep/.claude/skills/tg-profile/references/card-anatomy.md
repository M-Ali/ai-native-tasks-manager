# The card, block by block

One audience, one 16:9 slide (13.333 x 7.5in). The layout below is what `scripts/build_tg_card.py` draws. It follows the
persona cards agencies already use in Pakistani pitches — identity across the top, mindset and pain points down the middle,
triggers and media on the right, sources along the bottom.

| Block | Position (in) | Spec key | Notes |
|---|---|---|---|
| Header band | 0, 0 → 13.33 x 1.05 | `name`, `subtitle`, `tag` | Name and role big; `subtitle` carries age, location, land or income; `tag` is for a status flag such as PROVISIONAL |
| Who he is | 0.45, 1.25 → 4.0 x 1.5 | `household` (list) | Family status, home, land, literacy — each line traced in `sources` |
| Interests and affinity index | 0.45, 2.9 → 4.0 x 3.45 | `interests.items[]` = `{label, index, users}`, `interests.source`, `interests.gap` | Show index **and** users. With no dataset, `gap` prints one line saying so |
| Behavioural mindset | 4.65, 1.25 → 4.6 x 1.5 | `mindset` (list) | How they decide and what they fear, in the corpus' register |
| Pain points | 4.65, 2.95 → 4.6 x 3.4 | `pain_points[]` = `{title, detail, evidence}` | About six fit. `evidence` is checked but not printed — it lives in the spec so anyone can verify a line |
| Buying triggers | 9.45, 1.25 → 3.45 x 2.6 | `triggers[]` = `{title, detail, evidence}` | What moves them to act, including season |
| Where to reach him | 9.45, 3.95 → 3.45 x 2.4 | `media.items[]` = `{channel, note}`, `media.source` | Qualitative unless a survey is cited. A `%` here without `media.source` stops the build |
| Conclusion band | 0.45, 6.5 → 12.45 x 0.55 | `conclusion` | The one sentence you want remembered |
| Sources line | 0.45, 7.14 | `sources[]` | Files, pages, dataset dates |

## The insurance-deck card this was distilled from

A one-slide persona used in an earlier pitch (`data/reference/insurance middle age.pptx`, Aug 2026) carried: name and role,
age and city, family status and property type, top interests with affinity indices (i115-i137), behavioural mindset, four
pain points, a buying-triggers table, and a media-consumption strip with icons for TV, radio and OOH plus Facebook, YouTube,
Instagram and one more platform, each with a percentage.

Two things to copy and one to avoid:

- **Copy** the structure — identity, mindset, pains, triggers, interests, media — and the affinity-index notation (`i137`),
  which media planners read at a glance.
- **Copy** the habit of giving each pain point a title and an explanatory sentence, so the card briefs rather than labels.
- **Avoid** its unsourced media percentages. "(65%), (46%), (5%)" appear with no survey named, which is why this skill's
  builder refuses a percentage without a source.

## Adding photography

The builder draws no photograph, because a stock face implies an evidence base that is not there and a real farmer's face
needs consent. If the client supplies imagery, add it in PowerPoint after the build, or extend the builder with a
`photo` key pointing at a file the client owns.
