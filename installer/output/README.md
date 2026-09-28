# Built installers (not stored in git)

These files are produced by `installer/build-installer.ps1`. They are **gitignored** because `CAT-Setup.exe` embeds prerequisite installers and exceeds GitHub’s 100 MB per-file limit.

Download official builds from **[GitHub Releases](https://github.com/SuisoRa/CAT-PPT-Plugin/releases)** (`CAT-Setup.exe` recommended for new PCs).

To build locally:

```powershell
.\installer\build-installer.ps1
```

Outputs in this folder:

| File | Purpose |
|------|---------|
| `CAT-Setup.exe` | Burn bootstrapper (.NET 4.8 + VSTO + CAT MSI) |
| `CAT.msi` | CAT add-in only (prerequisites already installed) |
