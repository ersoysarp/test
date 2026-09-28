# test

## BlueCurrency.pptx

Editable 16:9 PowerPoint of the four-page deck. Open `BlueCurrency.pptx` in PowerPoint, Keynote or Google Slides — every title, box and chip is a native shape, not a screenshot. Vendor evidence that lives as hover notes on page 4 of the HTML is in the speaker notes of slide 4 (View → Notes). To rebuild after HTML changes: `python3 build_pptx.py`.

## BlueCurrency.html

The Blue Currency investment pitch deck. Open the file in any browser: arrow keys, swipe or the Prev/Next buttons move between the four pages, and Ctrl/Cmd+P exports to PDF at 1280x760 per page.

### Vendor bubbles

Hovering or tapping any vendor chip on pages 3 and 4 opens a card showing how AI-native the product is, which of the six modules it covers, and one line on what it does. Every card is generated from the single `VENDORS` table near the bottom of the HTML, so editing a vendor there updates it on both pages.

### Adding real vendor logos

Page 3 shows each player as a name chip. Drop a logo file into a `logos/` folder next to the HTML and it appears in front of the name, `svg` first, then `png`:

```
logos/kreo.svg
logos/oracle.png
```

The file name must match the chip's `data-v` value (`kreo`, `buildots`, `costify`, `zebel`, `conwize`, `ediphi`, `doxel`, `openspace`, `alice`, `nplan`, `bluebeam`, `planswift`, `ribcostx`, `ribcandy`, `sage`, `causeway`, `excel`, `oracle`, `procore`, `ribitwo`, `cleopatra`, `sapariba`, `autodesk`, `ineight`, `trimble`, `jaggaer`). Chips with no matching file just show the name, so a partial set is fine.
