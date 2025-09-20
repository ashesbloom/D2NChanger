# D2Nchanger (Wallpaper Scheduler)

D2Nchanger brings the elegant, time-based wallpaper scheduling found in macOS to the Windows desktop. Windows lacks native support for this kind of seamless, automated wallpaper transition, and this project aims to fill that gap.

### Why D2Nchanger?

macOS offers beautiful dynamic wallpapers that change throughout the day. When these HEIC files are extracted on Windows, you get a series of high-quality static images representing different times of day (e.g., morning, noon, night). D2Nchanger allows you to schedule these images to create a similar dynamic wallpaper experience on Windows.

You can set any image (JPEG, PNG, BMP) for a specific time frame, and the application will automatically transition them for you, all while using a minimal amount of system resources (typically 40-50MB of RAM).

### Future Upgrades

-   **Wallpaper Stacks:** Allow users to select a folder or multiple images for each time period (morning, afternoon, etc.). The application would then cycle through these images for a more varied and seamless transition throughout the day.

---

## Build & Run

### Prerequisites
- Windows
- Python 3.7 or higher

### Quick build (from project root):

1. Create and activate a virtual environment (optional but recommended):
   ```powershell
   python -m venv venv
   .\venv\Scripts\Activate.ps1
   ```

2. Install requirements:
   ```powershell
   python -m pip install --upgrade pip
   python -m pip install pyinstaller pillow pystray pywin32
   ```

3. Build the executable with PyInstaller (one-file, windowed):
   ```powershell
   python -m PyInstaller --noconsole --onefile --name D2Nchanger --icon D2Nchanger_icon.ico D2Nchanger.py
   ```

4. The built executable will be in `dist\D2Nchanger.exe`.

## Notes
- The app optionally uses `win32com` to create a startup shortcut. If `pywin32` is not available at runtime, that functionality is disabled.
- The generated executable is not code-signed. If distributing widely, consider code signing and creating an installer.
- The app uses `pystray` for the system tray icon. If issues arise, ensure your environment supports it.

## License
- MIT License. See `LICENSE` for details.
