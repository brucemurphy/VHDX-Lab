# VHDX Lab

VHDX Lab is a .NET 8 WPF utility for power users who need a friendlier way to prep Windows-to-Go style VHDX images and manage the boot entries they rely on. It automates the repetitive DISM and BCDEdit steps, adds clear progress feedback, and wraps everything in a polished dashboard with live status indicators.

## Highlights
- **Guided VHDX onboarding** – choose a source `.vhdx`, select any fixed disk as the destination, copy it there, and monitor throughput/percentage in real time.
- **Optional driver injection** – export the current machine's drivers with DISM, inject them into the mounted VHDX, and automatically clean up the staging folders.
- **Boot entry automation** – create a new BCDEdit entry, point `device`/`osdevice` to the VHDX, enable HAL detection, and surface success or failure details.
- **BCD inventory management** – list every boot entry, inspect the raw BCDEdit block, and remove non-protected entries directly from the UI.
- **Visual feedback** – DISM-style activity light, dual-line progress readout, and adaptive copy monitoring keep long-running operations transparent.
- **Dynamic wallpaper** – pulls the Bing daily image and metadata for a more engaging workspace.

## Requirements
- Windows 10/11 with administrator privileges (the app relaunches elevated when needed).
- DISM, BCDEdit, and Hyper-V/VHD features available on the host system.
- [.NET 8 SDK](https://dotnet.microsoft.com/) and Visual Studio 2022 (17.8+) with the .NET desktop workload.

## Build & Run
1. Clone the repo and restore dependencies:
   ```powershell
   git clone https://github.com/brucemurphy/VHDX-Lab.git
   cd "VHDX Lab"
   dotnet restore
   ```
2. Build or publish:
   ```powershell
   dotnet publish "VHDX Lab/VHDX Lab.vbproj" -c Release
   ```
3. Launch `VHDX Lab.exe` from `VHDX Lab\bin\Release\net8.0-windows\publish` **as Administrator**.

> ⚠️ The app writes to the system BCD store and mounts images. Test in a lab environment first and ensure you have recovery media.

## Publishing to GitHub
The `.gitignore` is configured so the published binaries under `VHDX Lab/bin/Release/net8.0-windows/publish` can be committed. After running `dotnet publish`, stage that directory, commit, and push to `origin/master` to surface the latest build on GitHub.

## Contributing / Feedback
Bug reports, suggestions, and pull requests are welcome. Open an issue on GitHub if you hit workflow snags or would like to see new automation scenarios (e.g., unattended provisioning, Azure-hosted images, etc.).
