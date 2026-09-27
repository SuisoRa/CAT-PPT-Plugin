# Offline prerequisites (Burn bundle)

Place these **exact filenames** in this folder before running `installer\build-installer.ps1`:

| File | Purpose | Download |
|------|---------|----------|
| `ndp48-x86-x64-allos-enu.exe` | .NET Framework 4.8 | [Microsoft .NET 4.8 offline installer](https://dotnet.microsoft.com/download/dotnet-framework/net48) |
| `vstor_redist.exe` | VSTO 2010 Runtime (x86/x64) | [Visual Studio 2010 Tools for Office Runtime](https://www.microsoft.com/download/details.aspx?id=56961) (or winget `Microsoft.VSTOR`; current CDN: `…/5d24f8f8-efbb-4b63-aa33-3785e3104713/vstor_redist.exe`) |

These files are **not** committed to git (large binaries). The bootstrapper installs them only when registry checks show they are missing.

**Not included:** Desktop Microsoft PowerPoint must already be installed on the target PC.
