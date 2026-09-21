# CAT PPT Plugin — PowerPoint add-in

Windows VSTO add-in for classic desktop PowerPoint (365 / 2021 / 2019 / 2016). The **CAT** ribbon tab adds tools for alignment, text boxes, tables, icons, and a floating **Auto format** toolbar. Most actions also have keyboard shortcuts.

**GitHub:** [github.com/SuisoRa/CAT-PPT-Plugin](https://github.com/SuisoRa/CAT-PPT-Plugin)

## Requirements

- Windows 10/11, classic PowerPoint (not web)
- Visual Studio 2022 with **.NET desktop development** and **Office/SharePoint development**
- .NET Framework 4.8

## Build and install

1. Open `src/CAT/CAT.sln` in Visual Studio.
2. Build **Release** (or F5 for Debug).
3. Register the add-in: run `installer/install-addin.ps1`, or set `Manifest` in `installer/register-addin.reg` to your `CAT.vsto` path and merge the reg file.
4. Restart PowerPoint.

Tool icons load from PNGs in `assets/icons` (embedded in `CAT.dll` on build). Names must match control ids, e.g. `catTableCreator.png`.

Settings: `HKCU\Software\CAT` (icon folder path, **Align by first**, Auto format toolbar position).

---

## Ribbon tools

### Auto format
Opens the floating Auto format toolbar (toggle). Esc closes it; position is remembered.

| Toolbar control id | Action |
|---|---|
| `catAutoFormat` | Open/close toolbar |
| `catAfLinesV` / `catAfLinesH` | Grey 0.5 pt lines in vertical / horizontal gaps |
| `catAfBracketLeft/Right/Top/Bottom` | Brace on that side |
| `catAfCalloutUp` | Callout pointing up |
| `catAfKeyTakeaway` | Light grey fill, no outline |
| `catAfNumbers` / `catAfLetters` | Numbered / letter circles |
| `catAfRectangle` | Picture circle → rectangle |
| `catAfSquare` | Make selection square (1:1) |
| `catAfTitle` / `catAfSubtitle` / `catAfFootnote` / `catAfSource` | Apply layout placeholder (fixed element) |

### Table Creator (`catTableCreator`)
Opens a task window:

- **Create table** — rows × columns as text boxes (dummy cells show **text** in theme Text 1).
- **Create table from clipboard** — Excel/TSV from clipboard; size detected automatically.
- **Convert native table → text boxes** — replaces selected PowerPoint table.
- **Table has header** — first row 14 pt bold; body 12 pt; 2 pt gaps between cells; horizontal row rules only.

Layout: table width = title width; height = 70% of space below title; cells grouped.

### Icon library (`catIconLibrary`)
Browse SVGs under `assets/icon-library` (fixed path). Search by tags, accent tint, optional enclosed circle. Single-click inserts vector SVG (ungroup/convert in PowerPoint).

### Align order
Stacked in one column:

| Control id | Action |
|---|---|
| `catAlignByFirst` | When checked, align shortcuts use **first-selected** shape as anchor |
| `catMakeVertical` | Line/connector: height 0 pt, top of bounds fixed |
| `catMakeHorizontal` | Line/connector: width 0 pt, left of bounds fixed |

### Select similar (ribbon only)
| Control id | Action |
|---|---|
| `catSelectSimilar` | Match fill, outline, and type on slide |
| `catSelectSimilarOptions` | Choose which properties to match |

### Copy for Excel (ribbon only)
`catCopyExcel` — copies selected shape text as a tab-separated grid (rows by vertical position). Shows row × column count in the confirmation dialog.

### Grouping
| Control id | Shortcut | Action |
|---|---|---|
| `catAlignRows` | Ctrl+Alt+Shift+R | Align rows (top anchor) and group as table rows |
| `catAlignCols` | Ctrl+Alt+Shift+C | Align columns (left anchor) and group as table columns |

### Distributing
| Control id | Shortcut | Action |
|---|---|---|
| `catDistH` | Alt+Shift+H | Distribute horizontally |
| `catDistV` | Alt+Shift+V | Distribute vertically |

### Position
| Control id | Shortcut | Action |
|---|---|---|
| `catCopyPos` | Ctrl+1 | Copy positions |
| `catPastePos` | Ctrl+2 | Paste positions |

### Sizing
| Control id | Shortcut | Action |
|---|---|---|
| `catSameH` | Ctrl+Shift+E | Same height |
| `catSameW` | Ctrl+Alt+E | Same width |
| `catSameSize` | Alt+Z | Same size |

### Text box properties
| Control id | Shortcut | Action |
|---|---|---|
| `catResize` | Ctrl+8 | Resize shape to fit text |
| `catNoResize` | Ctrl+Shift+8 | Do not autofit |
| `catSplitJoin` | Ctrl+Alt+J | Split / join text boxes |
| `catWrap` | Ctrl+7 | Word wrap on |
| `catNoWrap` | Ctrl+Shift+7 | Word wrap off |
| `catLineSpacing` | Ctrl+Shift+L | List line spacing |
| `catFixTextBox` | — | No fill/outline; margins 0; text top-aligned; theme Text 1 |

### Reset fixed
| Control id | Shortcut | Action |
|---|---|---|
| `catReset` | Ctrl+Alt+R | Reset fixed placeholders to layout |

### Shape support
| Control id | Action |
|---|---|
| `catSwapPosition` | Swap top-left of **exactly two** selected shapes (selection order) |
| `catSameCornerRadius` | Apply first shape’s corner adjustment to others (rounded shapes) |

### Help
| Control id | Action |
|---|---|
| `catHelp` | Popup list of keyboard shortcuts |

---

## Keyboard shortcuts (no ribbon button)

| Combo | Action |
|---|---|
| `Alt+G` | Top-align selection, then group |
| `Ctrl+Alt+↑ / ↓ / ← / →` | Align top / bottom / left / right |
| `Ctrl+Alt+C` / `Ctrl+Alt+M` | Align center / middle (one shape → slide; see **Align by first**) |
| `Alt+Shift+H` / `V` | Distribute horizontally / vertically |
| `Alt+]` / `Alt+Shift+]` | Bring forward / to front |
| `Alt+[` / `Alt+Shift+[` | Send backward / to back |
| `Ctrl+Alt+Delete` or `Ctrl+Shift+Delete` | **Nothing selected:** delete all shapes on slide. **Selection:** delete everything except selection |
| `Ctrl+5` | Zoom 100% |
| `Ctrl+M` | New slide (same layout; blocked if Alt held for align) |
| `Alt+Q` | Insert text box (12 pt, no margins, title-aligned) |
| `Ctrl+0` | Yellow sticky (top-right, half in/out of slide) |
| `Ctrl+Alt+T` | Paste unformatted text |
| `Alt+Shift+A` | Cycle theme accent fill |
| `Alt+Shift+← / →` | Decrease / increase list level |
| `Ctrl+3` / `Ctrl+4` | Normal view / Slide Sorter |

Rebind in `src/CAT/Core.cs` → `ShortcutMap.Bindings`.

---

## Project layout

```
CAT.sln
src/CAT/           Add-in code (Core, Commands, UI, AutoFormat, TableCreator, Icon library)
assets/icons/      Ribbon & toolbar PNGs (cat*.png)
assets/icon-library/  SVG icon catalog for Icon library
installer/         register-addin.ps1, register-addin.reg
```

The keyboard hook runs while PowerPoint is the foreground app; shortcuts are not limited to the CAT tab.
