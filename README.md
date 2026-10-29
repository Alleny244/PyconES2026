# Zero-Copy Data Sharing Between Python and Rust with PyO3 and Arrow

Talk materials for **PyCon ES 2026** — 8 Nov, Aula B7, 35 min + 5 min Q&A.
Allen Y — Associate Software Engineer, Red Hat.

## What's here

| Path | What it is |
|---|---|
| `slides/pycones2026-zero-copy-python-rust.pptx` | The deck — 36 slides, speaker notes on every slide |
| `slides/build_deck.py` | Generates the deck; reads every number from `bench/results.json` |
| `slides/assets/` | Speaker headshot (circular crop) and the LinkedIn QR |
| `demo/` | The Rust extension (`arrowbridge`) — PyO3 + Arrow C Data Interface |
| `bench/bench.py` | Correctness checks and the benchmarks |
| `bench/results.json` | Measured results the deck is built from |

No figure in the deck is typed by hand. Re-run the benchmark, re-run the
build, and the slides update.

## Setup

```bash
python3 -m venv .venv
.venv/bin/pip install maturin pyarrow numpy pandas python-pptx
cd demo && ../.venv/bin/maturin develop --release
```

Requires a Rust toolchain (`cargo`). `arrow` must stay pinned to the version
`pyo3-arrow` depends on — see Pitfall 1 in the talk.

## Run it

```bash
cd bench
../.venv/bin/python bench.py verify   # correctness + the zero-copy pointer proof
../.venv/bin/python bench.py all      # full suite, writes results.json
```

Rebuild the deck after a benchmark run:

```bash
cd slides && ../.venv/bin/python build_deck.py
```

## The live demo

The moment worth rehearsing is slide 20 — showing that the pointer is
literally the same on both sides:

```python
import pyarrow as pa, arrowbridge
arr = pa.array([1.0, 2.0, 3.0])
arr.buffers()[1].address        # what pyarrow owns
arrowbridge.arrow_data_ptr(arr) # what Rust reads
```

## Measured on

Intel Core Ultra 7 165U · Fedora 44 · Python 3.12.9 · pyarrow 26.0.0 ·
PyO3 0.30 · arrow-rs 59.3. Numbers will differ on your machine; re-run
the benchmark before presenting.

## Theme

PyCon ES 2026 branding on a white background. Colors and fonts are taken
from the conference site's own CSS (`--color-pycon-*` custom properties):

| Token | Hex | Used for |
|---|---|---|
| pycon-black | `#1d1d1b` | body ink |
| pycon-gray | `#4a4a48` | secondary ink |
| pycon-black-50 / -25 | `#8e8e8d` / `#c6c6c6` | labels, rules |
| pycon-orange | `#ff8200` | accent bars (large shapes) |
| pycon-mute | `#cc6600` | accent text (contrast-safe on white) |
| pycon-red | `#f9423a` | "copies the data" in charts |

Fonts: **Outfit** (sans) and **JetBrains Mono** (code) — the same two the
site uses. Install both on the presenting machine, or embed them in
PowerPoint (File → Options → Save → Embed fonts).

```bash
mkdir -p ~/.local/share/fonts && cd ~/.local/share/fonts
curl -sLO "https://github.com/google/fonts/raw/main/ofl/outfit/Outfit%5Bwght%5D.ttf"
curl -sLO "https://github.com/google/fonts/raw/main/ofl/jetbrainsmono/JetBrainsMono%5Bwght%5D.ttf"
fc-cache -f
```

### A note on the chart colors

The brand is three warm hues (orange, red, yellow). That cannot carry a
four-series chart: the raw brand colors fail the chroma, colorblind-
separation and contrast checks on white, and yellow sits at 1.6:1 against
the background. So the deck keeps brand orange as the *accent* and uses
validated brand-adjacent steps for *data*:

- bars — `#f9423a` / `#2f5fa8`, encoding the strategy (copying vs sharing)
  rather than four arbitrary categories; passes every check
- lines — `#f9423a,#2f5fa8,#00895e,#c66a00`, which passes with the
  green/orange pair in the colorblind warn band, so those series also
  carry dash patterns and distinct markers

## Built to the PyConES speaker guide

| Guide requirement | How the deck meets it |
|---|---|
| Titles 36–44pt | 36pt, 44pt on section dividers |
| Main text 28–32pt | 28pt body, 32pt for emphasis lines |
| Code and chart labels 24–28pt | 24pt code, 24pt data labels, 22pt axes |
| Nothing readable below 24pt | only footers (14pt) and method captions / kickers (18pt) |
| White background, dark text | `#ffffff` with `#1d1d1b` ink (15.9:1) |
| One idea per slide | 39 slides for 35 minutes |
| No information by colour alone | chart bars carry a text legend *and* category labels; lines carry dash patterns and distinct marker shapes; the comparison table says yes/no in words |
| Alt text for sharing | every chart and diagram described |
| Materials link at start and end | title slide, every footer, closing slide |
| QR accompanied by a written URL | closing slide shows the QR *and* `linkedin.com/in/allen-yesudasan` at 24pt |
| Demo fallback | slide 22 shows the saved pointer output, so the live demo is optional |

A `MAXCOL` assertion in `build_deck.py` fails the build if any code line is
too long to fit at 24pt, so code can never silently wrap.

Speaker notes on all 39 slides carry a cumulative timing cue (`[12:10]`…)
plus what to say out loud for each chart and diagram, per the guide's
accessibility section.

### Timing

| Section | Slides | Time |
|---|---|---|
| Problem and context | 1–5 | 0:00–4:00 |
| Parts 01 and 02 | 6–13 | 4:00–11:15 |
| Part 03 — Arrow | 14–30 | 11:15–28:30 |

Part 03 leads with the payoff: slide 16 shows the identical pointer as a
teaser, then slides 17–22 explain how it is possible. Don't explain first.
| Part 04 — pitfalls | 31–37 | 28:30–34:10 |
| Takeaways and close | 38–39 | 34:10–35:00 |
| Q&A | — | 5 min |

The running example is `sensors.parquet`, 5 million readings. It is named
on slide 2 and called back on slides 22 and 33 — use the name each time.

If you run long, the first things to cut are slides 20 (the dunder),
30 (RecordBatch) and 37 (ecosystem) — none of them carry the argument.

## Before you present

- Replace the repo URL (`REPO` in `build_deck.py`) with your own — the
  placeholder is long enough to wrap onto two lines on the closing slide.
- Install the two fonts above, or embed them.
- Re-run `bench.py all` on the presenting laptop and rebuild, so the
  numbers you quote are the numbers on screen.
