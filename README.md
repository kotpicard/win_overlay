# Windows Overlay Tool

A desktop overlay application for Windows that can display text over your screen and optionally integrate with Microsoft PowerPoint.  
The overlay and its behavior (fonts, colors, hotkeys, text sources) are fully configurable through a simple interface.

## Features
- Transparent overlay window that stays on top of all other windows.
- Customizable font, size, and color for displayed text.
- Supports text sources from:
  - Plain text files.
  - PowerPoint presentations.
- Configurable hotkeys for:
  - Showing/hiding the overlay.
  - Navigating between overlay texts.
  - Controlling PowerPoint slides.
  - Starting/stopping a logging timer.
- Option to automatically synchronize overlay text changes with PowerPoint slide changes.
- Logging support for capturing overlay activity.

## Executables
Prebuilt `.exe` files are available so you don’t need Python installed.

- **`overlay_configuration.exe`**  
  Opens the configuration UI. Use this tool to:
  - Select your text source (PowerPoint file or text file).
  - Configure overlay font, size, color, and position.
  - Set up hotkeys for overlay and PowerPoint controls.
  - Save your settings to `settings.json`.

- **`overlay_controller.exe`**  
  Runs the overlay based on your saved settings. It will:
  - Display overlay text on screen.
  - Listen for your configured hotkeys.
  - Control PowerPoint slides if configured.

Download both from the [Releases](https://github.com/kotpicard/win_overlay/releases) page.

## Installation

### Option 1: Use the executables (recommended)
1. Download the latest `.exe` files from the [Releases](https://github.com/kotpicard/win_overlay/releases) page.
2. Run `overlay_configuration.exe` to set up your preferences.
3. Run `overlay_controller.exe` to launch the overlay.

### Option 2: Run from source
1. Clone this repository:
   ```bash
   git clone https://github.com/kotpicard/win_overlay.git
   cd win_overlay
2. Run `setup.bat` to install dependencies.
3. Run `overlay_configuration.py` to create the `settings.json` file.
4. Run `overlay_controller.py`.

## Quick Start

1. Download and run `overlay_configuration.exe`
2. Configure your settings in the GUI:
   - Choose text source (PowerPoint file or text file)
   - Set font, size, and color
   - Click to set overlay position
   - Configure hotkeys
3. Click "Save All Settings"
4. Run `overlay_controller.exe` to start the overlay

## Configuration Options

### General Settings

- **Text Source**: Choose between PowerPoint presentations or plain text files
- **PowerPoint Control**: 
  - Synchronize overlay text changes with slide navigation
  - Auto-hide overlay when PowerPoint is open
- **Logging**: Enable activity logging with custom file location

### Text Settings

- **Font Selection**: Choose any installed system font
- **Font Size**: Adjustable from 0-100 points
- **Color**: Full color picker with named color support
- **Position**: Click anywhere on screen to set overlay position

### Hotkey Settings

Customize all keyboard shortcuts:
- **Overlay Controls**: Show/Hide overlay, Next/Previous text
- **PowerPoint Controls**: Navigate slides, Show/Hide PowerPoint
- **System Controls**: Start logging timer, Quit application

## Default Hotkeys

| Action | Default Shortcut |
|--------|------------------|
| Show Overlay | `Ctrl + Shift + W` |
| Hide Overlay | `Ctrl + Shift + Q` |
| Next Text | `Ctrl + Shift + T` |
| Previous Text | `Ctrl + Shift + R` |
| Show PowerPoint | `Ctrl + Shift + G` |
| Hide PowerPoint | `Ctrl + Shift + A` |
| Next Slide | `Ctrl + Shift + F` |
| Previous Slide | `Ctrl + Shift + D` |
| Start Timer | `Ctrl + Shift + S` |
| Quit | `Ctrl + Shift + X` |

## Usage Modes

### PowerPoint Mode
- Load a PowerPoint file as your text source
- Overlay displays slide content
- Navigate slides with hotkeys
- Optional synchronization between overlay text and slide changes

### Text File Mode
- Use a plain text file with one line per overlay text
- Navigate through text entries with hotkeys
- Simpler setup for basic text overlays

## File Structure

When you save settings, the application creates`settings.json` - a configuration file with all your preferences.


## System Requirements

- Windows operating system
- PowerPoint installation (for PowerPoint mode)

## Technical Details

The application consists of two main components:
- **Configuration Tool** (`overlay_configuration.exe`): GUI for setting up preferences
- **Overlay Controller** (`overlay_controller.exe`): Runtime overlay display system

## License

This project is available under the GNU GPLv3 License. See the LICENSE file for details.

## Author

Created by [kotpicard](https://github.com/kotpicard)

---

**Note**: This application is designed for Windows only.

