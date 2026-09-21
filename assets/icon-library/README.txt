Icon library (SVG)
====================

Place your icons here:

  icon-library/
    FamilyName/
      tag1-tag2-tag3.svg
      arrow-up-blue.svg

- **Family name** = subfolder name (shown in the left list).
- **Tags** = parts of the file name before `.svg`, split on `-`, `_`, or space.
- Search matches any tag (e.g. `arrow blue` finds files whose tags contain both words).

Open **CAT → Icon library** in PowerPoint. Icons are read from this folder automatically (no user path setting).
Double-click a tile to insert the SVG on the current slide. Choose accent colour or “enclosed in accent circle” in the dialog.

Easier tagging later: we can add an optional `tags.txt` sidecar or a small JSON manifest per family if filenames get unwieldy — say if you want that.
