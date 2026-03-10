# Investor Pitch Deck Generator

Python script that generates a professional 13-slide investor pitch deck using `python-pptx`. Built for industrial real estate funds, but the structure and techniques apply to any institutional pitch.

## What It Generates

A 16:9 PPTX presentation with:

- Cover slide with branded stats bar
- Market opportunity overview
- Company background with key metrics
- Multi-market comparison (3 columns)
- Demand drivers grid (2x3)
- Tenant demand showcase
- Investment strategy (dual-column split)
- Return engine (5 numbered levers)
- Fund structure and terms
- Risk management framework
- Competitive advantage summary
- The ask (dark slide with terms recap)
- Contact slide

## Features

- Configurable brand colors and fonts at the top of the file
- Reusable helper functions: `add_accent_bar()`, `add_section_label()`, `add_title()`, `add_body()`, `add_bullet_list()`, `add_stat_box()`, `add_kv_row()`
- Logo watermark placement on content slides
- Dark and light slide backgrounds
- Professional typography with Inter font
- Data-driven: all content is defined in Python data structures, easy to swap

## Usage

```bash
pip install python-pptx
python3 build_investor_deck.py
```

Output: `./output/Investor_Pitch_Deck.pptx`

## Customization

1. Update the brand constants at the top (`ACCENT`, `SECONDARY`, `DARK_SLATE`, etc.)
2. Place your logo files in `./assets/`
3. Modify the content data structures (stats, cities, drivers, etc.) for your fund
4. Run the script

## Directory Structure

```
investor-pitch-generator/
  build_investor_deck.py    # Main script
  assets/                   # Place logo PNGs here
    logo_dark.png
    logo_white.png
    icon_accent.png
  output/                   # Generated deck lands here
```

## Requirements

- Python 3.8+
- python-pptx

## License

MIT
