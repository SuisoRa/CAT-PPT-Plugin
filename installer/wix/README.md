# WiX installer for CAT (per-user VSTO add-in)

## What gets installed

| Location | Contents |
|----------|----------|
| `%LocalAppData%\Programs\CAT\` | `CAT.dll`, `CAT.vsto`, dependencies, `icons\`, `text-block\` |
| `%LocalAppData%\Programs\CAT\icon-library\` | SVG catalog (**replace files here to update icons without reinstalling**) |

PowerPoint registration: `HKCU\Software\Microsoft\Office\PowerPoint\Addins\CAT`

## Build the installer (on your dev PC)

1. Install **WiX Toolset v3.14** ([wixtoolset.org](https://wixtoolset.org/)).
2. Download offline prerequisites into `installer/prereqs/` (see `prereqs/README.md`).
3. Build **Release** in Visual Studio (manifest signing uses the cert thumbprint in `CAT.csproj`).
4. From repo root:

   ```powershell
   .\installer\build-installer.ps1
   ```

Outputs:

- `installer/output/CAT.msi` — add-in only (assumes .NET 4.8 + VSTO already installed)
- `installer/output/CAT-Setup.exe` — Burn bundle: installs prerequisites if missing, then the MSI

## Deploy to a fresh PC

1. Copy `CAT-Setup.exe` (and optionally the whole `output` folder) to a share or USB.
2. Run **CAT-Setup.exe** (per-user; no admin required for the add-in MSI step; prerequisite EXEs may prompt for elevation).
3. Restart PowerPoint.

Desktop **PowerPoint** must already be installed; the bundle does not install Office.

## Version bumps

Update `ProductVersion` in `Product.wxs` and `Bundle.wxs`, and `AssemblyVersion` in `ThisAddIn.cs`, then rebuild the installer.
