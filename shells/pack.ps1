$ErrorActionPreference = "Stop"

# Build GUI executable
python -m PyInstaller --noconsole --onefile --name "SUDA-Net-Daemon" --icon "resources\suda-logo.png" --add-data "resources;resources" --add-data "config.json;." .\gui.py

# Copy extra files
Copy-Item -Path "config.json" -Destination "dist\config.json" -Force
Copy-Item -Path "README.md" -Destination "dist\README.md" -Force

# Zip release
$compress = @{
    Path = "dist\SUDA-Net-Daemon.exe", "dist\config.json", "dist\README.md"
    CompressionLevel = "Fastest"
    DestinationPath = "dist\SUDA-Net-Daemon.zip"
}
Compress-Archive @compress -Force
