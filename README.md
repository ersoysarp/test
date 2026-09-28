# test

## BlueCurrency.pptx

Editable 16:9 PowerPoint of the four-page deck. Open `BlueCurrency.pptx` in PowerPoint, Keynote or Google Slides — every title, box and chip is a native shape, not a screenshot. Vendor evidence that lives as hover notes on page 4 of the HTML is in the speaker notes of slide 4 (View → Notes). To rebuild after HTML changes: `python3 build_pptx.py`.

## BlueCurrency.html

The Blue Currency investment pitch deck. Open the file in any browser: arrow keys, swipe or the Prev/Next buttons move between the four pages, and Ctrl/Cmd+P exports to PDF at 1280x760 per page.

### Adding real vendor logos

Page 3 shows each player as a chip with a monogram tile. Drop a logo file into a `logos/` folder next to the HTML and the monogram is replaced automatically — `svg` is tried first, then `png`:

```
logos/kreo.svg
logos/oracle.png
```

The file name must match the chip's `data-logo` value (`kreo`, `buildots`, `costify`, `zebel`, `conwize`, `ediphi`, `doxel`, `openspace`, `alice`, `nplan`, `bluebeam`, `planswift`, `ribcostx`, `ribcandy`, `sage`, `causeway`, `excel`, `oracle`, `procore`, `ribitwo`, `cleopatra`, `sapariba`, `autodesk`, `ineight`, `trimble`, `jaggaer`). Chips with no matching file keep their monogram, so a partial set is fine.
