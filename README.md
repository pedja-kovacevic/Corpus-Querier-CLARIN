# Corpus Querier

Corpus Querier is a Windows desktop application for running CQL frequency queries against corpora hosted by CLARIN.SI. It processes existing CQL spreadsheets or generates queries from words and reusable `{WORD}` / `{LEMMA}` templates.

## For users

Download the appropriate package from the latest GitHub release:

- **Windows:** `Corpus-Querier-Setup.exe`
- **Ubuntu, Debian or Linux Mint:** `corpus-querier_1.0.0_amd64.deb`
- **Other x86-64 Linux distributions:** `Corpus-Querier-x86_64.AppImage`

The Windows installer creates Start Menu and optional Desktop shortcuts and adds Corpus Querier to Windows **Installed apps**, from which it can be uninstalled normally.

Python, PowerShell, VS Code, and manual dependency installation are not required.

### Linux installation

On Ubuntu, Debian or Linux Mint, double-click the `.deb` file and open it with the system Software Installer. It can also be installed from a terminal with:

```bash
sudo apt install ./corpus-querier_1.0.0_amd64.deb
```

The AppImage does not install anything. Mark it executable and open it:

```bash
chmod +x Corpus-Querier-x86_64.AppImage
./Corpus-Querier-x86_64.AppImage
```

Some distributions require the FUSE 2 compatibility package to launch AppImages.

> Windows SmartScreen may warn about an unsigned installer downloaded from the internet. Code signing can remove this warning, but it requires a commercial signing certificate.

## For maintainers: automatic GitHub build

The included GitHub Actions workflow can build the Windows installer without requiring a Windows development environment.

1. Commit and push all project files, including the hidden `.github` folder.
2. Open **Actions** and run both **Build Windows installer** and **Build Linux packages**.
3. Download their Windows and Linux artifacts when the runs finish.

To publish a release automatically, create and push a version tag:

```bash
git tag v1.0.0
git push origin v1.0.0
```

GitHub will build the Windows installer, Linux AppImage and Debian package, create the release, and attach all three downloads.

## For maintainers: local Windows build

Install Python 3 and [Inno Setup 6](https://jrsoftware.org/isdl.php), then right-click `build_installer.ps1` and choose **Run with PowerShell**, or run:

```powershell
powershell -ExecutionPolicy Bypass -File .\build_installer.ps1
```

The finished installer is written to:

```text
installer-output\Corpus-Querier-Setup.exe
```

## Installer behavior

- Installs per user under `%LOCALAPPDATA%\Programs\Corpus Querier`.
- Does not request administrator rights.
- Includes Python and all runtime dependencies.
- Creates Start Menu and optional Desktop shortcuts.
- Registers a standard Windows uninstaller.
- Preserves the app's Excel input/output and query behavior.

## License

MIT
