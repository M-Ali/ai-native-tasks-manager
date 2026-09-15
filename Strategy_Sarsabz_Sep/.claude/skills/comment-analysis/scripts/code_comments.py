"""Code comments against a theme and brand lexicon, and count what is there.

    python code_comments.py <dir>/comments.csv --lexicon lexicon.yaml --out <dir>

Writes coded.csv (one row per comment, matched themes and brands), theme_counts.csv,
brand_counts.csv, brand_sentiment.csv, cooccurrence.csv and coding_report.md.

WHY DETERMINISTIC, WHEN A MODEL COULD READ THEM

Two reasons, and they are about defensibility rather than capability:

1. **The counts get quoted.** "118 comments raise after-sales" ends up on a slide and in
   a client's head. It must be reproducible on a rerun and traceable to the exact
   comments that matched, so anyone can check it by reading them.
2. **A model reading 1,000 comments resamples.** Run it twice, get two different numbers,
   and the analysis cannot be defended in a room.

So: the ARITHMETIC is mechanical here, and the JUDGEMENT is yours. This script tells you
how often a subject comes up and how it skews. It cannot tell you what it means, which
comment is the telling one, or what the brand should do - that comes from reading the
comments the script points you at. Both halves are needed; neither substitutes.

SENTIMENT IS DIRECTIONAL, NOT MEASURED

`brand_sentiment.csv` counts comments containing explicit positive or negative markers
near a brand mention. That is a crude proxy and the report says so. It is useful for
"Oshan is talked about warmly, MG is not", never for "MG has 10.7% negative sentiment"
stated as a measurement. Sarcasm, code-switching and Roman Urdu all defeat it, which is
precisely why the qualitative read is not optional.

The lexicon ships with a general vocabulary. ALWAYS extend it for the category before
running - twenty minutes with the client's own product language - because an unmatched
theme is invisible, and coverage below ~50% usually means the lexicon, not the corpus,
is wrong.
"""

from __future__ import annotations

import argparse
import csv
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

POS = [r"\bbest\b", r"\bgreat\b", r"\bgood\b", r"\bexcellent\b", r"\bamazing\b",
       r"\blove\b", r"\bperfect\b", r"\bsuperb\b", r"\bawesome\b", r"\breliable\b",
       r"\bsmooth\b", r"\bfast\b", r"\bworth\b", r"\brecommend\b", r"\bsatisfied\b",
       r"\bachi?a?\b", r"\bbehtar\b", r"\bzabardast\b", r"\bshandar\b", r"\bmast\b",
       r"\btheek\b", r"\bsahi\b", r"\bbrilliant\b", r"\bhappy\b"]
NEG = [r"\bworst\b", r"\bbad\b", r"\bpathetic\b", r"\bpoor\b", r"\bterrible\b",
       r"\bhorrible\b", r"\bawful\b", r"\bscam\b", r"\bfraud\b", r"\bwaste\b",
       r"\bnever\b", r"\bhate\b", r"\bslow\b", r"\bissue\b", r"\bproblem\b",
       r"\bcomplain(t|ts|ing)?\b", r"\bdisappoint", r"\buseless\b", r"\bghati?ya\b",
       r"\bbakwas\b", r"\bfazool\b", r"\bfazol\b", r"\bkharab\b", r"\bmasla\b",
       r"\blanat\b", r"\bdhoka\b", r"\bbekar\b", r"\bnot working\b", r"\bdoesn.?t work\b"]

# A negated positive is a negative: "not good", "no service", "kabhi sahi nahi".
NEGATORS = r"(?:not|no|never|n't|nahi|nai|nhi|kabhi nahi|bilkul nahi)\W+(?:\w+\W+){0,2}"


def compile_terms(terms, where: str = "") -> list[re.Pattern]:
    out = []
    for t in terms or []:
        # YAML double-quoting eats backslashes: "\bptcl\b" arrives as two BACKSPACE
        # characters and matches nothing, silently, which is how a client brand once
        # read as 11 mentions in a corpus holding 89. Fail loudly instead.
        if any(ord(c) < 32 for c in t):
            bad = "".join(f"\\x{ord(c):02x}" if ord(c) < 32 else c for c in t)
            sys.exit(
                f"Control character in pattern {bad!r}{' under ' + where if where else ''}.\n"
                "This is YAML double-quoting processing the backslash escapes. Use "
                "SINGLE quotes for regex in the lexicon: '\\bword\\b', not \"\\bword\\b\"."
            )
        # a bare word is matched whole; anything with regex metachars is used as-is
        pat = t if re.search(r"[\\\[\](){}|+*?^$]", t) else rf"\b{re.escape(t)}\b"
        try:
            out.append(re.compile(pat, re.I))
        except re.error as exc:
            sys.exit(f"Bad pattern {t!r}{' under ' + where if where else ''}: {exc}")
    return out


def load_lexicon(path: Path | None) -> dict:
    if path is None:
        sys.exit("--lexicon is required. Copy assets/lexicon.example.yaml and adapt it.")
    import yaml
    data = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    if not data.get("themes"):
        sys.exit(f"{path}: no `themes:` block.")
    return data


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("comments", type=Path)
    ap.add_argument("--lexicon", type=Path, required=True)
    ap.add_argument("--out", type=Path, required=True)
    ap.add_argument("--window", type=int, default=90,
                    help="chars either side of a brand mention searched for sentiment "
                         "markers; the whole comment is used when it is shorter")
    args = ap.parse_args()

    lex = load_lexicon(args.lexicon)
    themes = {k: compile_terms(v, f"themes.{k}") for k, v in (lex.get("themes") or {}).items()}
    brands = {k: compile_terms(v, f"brands.{k}") for k, v in (lex.get("brands") or {}).items()}
    pos = compile_terms(lex.get("positive_extra", [])) + [re.compile(p, re.I) for p in POS]
    neg = compile_terms(lex.get("negative_extra", [])) + [re.compile(p, re.I) for p in NEG]
    negated = [re.compile(NEGATORS + p.pattern, re.I) for p in pos]

    args.out.mkdir(parents=True, exist_ok=True)
    with args.comments.open(encoding="utf-8", newline="") as fh:
        rows = list(csv.DictReader(fh))
    if not rows:
        sys.exit(f"{args.comments}: no rows")

    theme_ct: Counter = Counter()
    brand_ct: Counter = Counter()
    brand_pos: Counter = Counter()
    brand_neg: Counter = Counter()
    pair_ct: Counter = Counter()
    coded = []
    matched_any = 0

    for r in rows:
        t = r["text"]
        hit_themes = [name for name, pats in themes.items() if any(p.search(t) for p in pats)]
        hit_brands = [name for name, pats in brands.items() if any(p.search(t) for p in pats)]
        if hit_themes:
            matched_any += 1
        for th in hit_themes:
            theme_ct[th] += 1
        for b in hit_brands:
            brand_ct[b] += 1
        for th in hit_themes:
            for b in hit_brands:
                pair_ct[(b, th)] += 1

        # sentiment in the neighbourhood of each brand mention
        for b in hit_brands:
            span = t
            for p in brands[b]:
                m = p.search(t)
                if m:
                    span = t[max(0, m.start() - args.window): m.end() + args.window]
                    break
            has_negated = any(p.search(span) for p in negated)
            is_neg = any(p.search(span) for p in neg) or has_negated
            is_pos = any(p.search(span) for p in pos) and not has_negated
            if is_neg:
                brand_neg[b] += 1
            elif is_pos:
                brand_pos[b] += 1

        coded.append({**r, "themes": "|".join(hit_themes), "brands": "|".join(hit_brands)})

    def dump(name, header, data):
        with (args.out / name).open("w", encoding="utf-8", newline="") as fh:
            w = csv.writer(fh)
            w.writerow(header)
            w.writerows(data)

    dump("theme_counts.csv", ["theme", "comments", "share_pct"],
         [[k, v, round(100 * v / len(rows), 1)] for k, v in theme_ct.most_common()])
    dump("brand_counts.csv", ["brand", "mentions", "share_pct"],
         [[k, v, round(100 * v / len(rows), 1)] for k, v in brand_ct.most_common()])
    dump("brand_sentiment.csv",
         ["brand", "mentions", "explicit_positive", "positive_pct",
          "explicit_negative", "negative_pct", "unmarked"],
         [[b, n, brand_pos[b], round(100 * brand_pos[b] / n, 1),
           brand_neg[b], round(100 * brand_neg[b] / n, 1),
           n - brand_pos[b] - brand_neg[b]]
          for b, n in brand_ct.most_common()])
    dump("cooccurrence.csv", ["brand", "theme", "comments"],
         [[b, th, n] for (b, th), n in pair_ct.most_common()])
    with (args.out / "coded.csv").open("w", encoding="utf-8", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=list(coded[0].keys()))
        w.writeheader()
        w.writerows(coded)

    cov = round(100 * matched_any / len(rows), 1)
    lines = [
        "# Coding report", "",
        f"- Comments coded: **{len(rows)}**",
        f"- Matched at least one theme: **{matched_any} ({cov}%)**",
        f"- Themes in lexicon: {len(themes)} | brands: {len(brands)}",
        "",
        "Unmatched comments are mostly short affirmations, greetings and off-topic "
        "chatter that carry no codeable subject. Stating coverage is what makes the "
        "matched counts credible.",
        "",
    ]
    if cov < 50:
        lines += [
            f"> **Coverage is {cov}%, which is low.** That usually means the lexicon does "
            "not fit this category's vocabulary rather than that the corpus is empty. "
            "Read 40-50 unmatched comments and extend the lexicon before relying on "
            "these counts.", "",
        ]
    lines += [
        "## Sentiment method", "",
        "Explicit positive/negative markers within "
        f"{args.window} characters of a brand mention, with negated positives "
        '("not good", "sahi nahi") counted as negative. This is **directional, not '
        "measured**: sarcasm, code-switching and Roman Urdu spelling variants all defeat "
        "it, and most mentions carry no explicit marker at all. Use it to say which "
        "brands are discussed warmly, never to publish a sentiment percentage as fact.",
    ]
    (args.out / "coding_report.md").write_text("\n".join(lines) + "\n", encoding="utf-8")

    print(f"{args.out}/coded.csv + 4 count file(s)")
    print(f"  theme coverage {cov}% | {len(brand_ct)} brand(s) mentioned")
    for b, n in brand_ct.most_common(8):
        print(f"    {b:<16}{n:>5} mentions  (+{brand_pos[b]} / -{brand_neg[b]})")


if __name__ == "__main__":
    main()
