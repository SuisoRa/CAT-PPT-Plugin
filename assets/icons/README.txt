Ribbon and Auto Format toolbar icons are PNG files named after each control id.

Examples:
  catAutoFormat.png
  catAfLinesV.png
  catAfNumbers.png
  catAlignRows.png

Use PNG with transparency. 32x32 at 100% DPI, or 64x64 for high-DPI displays.

Folder resolution (first match wins):
  1. HKCU\Software\CAT\IconFolderPath if set and the folder exists
  2. icons\ next to CAT.dll (populated on build from this folder)
  3. assets\icons under the project tree when running from Visual Studio

There are no built-in drawn fallbacks — each visible control needs a matching PNG.
