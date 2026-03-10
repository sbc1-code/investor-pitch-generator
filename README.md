# Investor Pitch Deck Generator

Python script that generates professional investor pitch decks.

**[Live Preview](https://sbc1-code.github.io/investor-pitch-generator/)** -- see the 13-slide output rendered as HTML.

## What It Does

Generates a 16:9 PPTX investor pitch deck with 13 slides:

1. Cover with branded stats bar
2. Market opportunity overview
3. Company background with key metrics
4. Multi-market comparison (3 columns)
5. Demand drivers grid (2x3)
6. Tenant demand showcase
7. Investment strategy (dual-column split)
8. Return engine (5 numbered levers)
9. Fund structure and terms
10. Risk management framework
11. Competitive advantage summary
12. The ask (dark slide with terms recap)
13. Contact slide

Built for industrial real estate funds, but the structure works for any institutional pitch. All content is defined in Python data structures, so swapping in your own data is straightforward.

## Usage

```bash
pip install python-pptx
python3 build_investor_deck.py
```

Output: `./output/Investor_Pitch_Deck.pptx`

## Customization

1. Update the brand constants at the top of the script (`ACCENT`, `SECONDARY`, `DARK_SLATE`, etc.)
2. Place your logo files in `./assets/`
3. Modify the content data structures (stats, cities, drivers, etc.) for your fund
4. Run the script

## Features

- Configurable brand colors and fonts
- Reusable helper functions: `add_accent_bar()`, `add_section_label()`, `add_title()`, `add_body()`, `add_bullet_list()`, `add_stat_box()`, `add_kv_row()`
- Logo watermark placement on content slides
- Dark and light slide backgrounds
- Professional typography with Inter font
- Data-driven: all content lives in Python data structures

## Requirements

- Python 3.8+
- python-pptx

## License

MIT
