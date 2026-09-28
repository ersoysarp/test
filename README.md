# test

## BlueCurrency.html

The Blue Currency investment pitch deck. Open the file in any browser: arrow keys, swipe or the Prev/Next buttons move between the four pages, and Ctrl/Cmd+P exports to PDF at 1280x760 per page.

### Adding real vendor logos

Page 3 shows each player as a chip with a monogram tile. Drop a logo file into a `logos/` folder next to the HTML and the monogram is replaced automatically — `svg` is tried first, then `png`:

```
logos/kreo.svg
logos/oracle.png
```

The file name must match the chip's `data-logo` value (`kreo`, `buildots`, `costify`, `zebel`, `conwize`, `ediphi`, `doxel`, `openspace`, `alice`, `nplan`, `bluebeam`, `planswift`, `ribcostx`, `ribcandy`, `sage`, `causeway`, `excel`, `oracle`, `procore`, `ribitwo`, `cleopatra`, `sapariba`, `autodesk`, `ineight`, `trimble`, `jaggaer`). Chips with no matching file keep their monogram, so a partial set is fine.
