# Desktop Profile Manager

A powerful Windows desktop management application that allows you to create, save, and switch between different desktop configurations with custom layouts, icon arrangements, themes, and desktop folders.

![Desktop Profile Manager](https://img.shields.io/badge/Platform-Windows-blue) ![Version](https://img.shields.io/badge/Version-1.0-green) ![License](https://img.shields.io/badge/License-MIT-yellow)

![demo](./media/demo.gif)

## Features

- **Multiple Desktop Profiles** - Create unlimited desktop configurations for different workflows
- **Icon Layout Preservation** - Save and restore exact desktop icon positions
- **Custom Desktop Folders** - Use any folder as your desktop for each profile  
- **Windows Theme Integration** - Automatically apply different themes per profile
- **Temporary Desktop Mode** - Quickly switch to any folder without saving changes
- **System Tray Integration** - Background operation with quick access from tray
- **Context Menu Integration** - Right-click folders for instant profile creation

## Installation

```bash
# Clone the repository
git clone https://github.com/Gigio-Bagigio/desktop-manager.git
cd desktop-manager

# Install dependencies
pip install -r requirements.txt

# Run the application
python app.py
```

**Requirements:** Windows 10/11

## Quick Start

### Creating Your First Profile
1. Launch the application
2. Enter a name for your profile (e.g., "Work", "Gaming", "Design")
3. Click **"Create Profile"**
4. Select the folder you want to use as desktop for this profile
5. Your current desktop layout and theme are automatically saved!

### Switching Between Profiles
- **From GUI:** Select a profile and click **"Upload Profile"**
- **From System Tray:** Right-click tray icon → Profiles → [Profile Name] → Load

### Using Temporary Desktop
Perfect for quick tasks without permanent changes:
- Click **"Open Temporary Desktop"** and select any folder
- Right-click any folder → **"Open as Temporary Desktop"**

## Use Cases

- **Work Setup** - Separate profiles for different projects with relevant files on desktop
- **Gaming** - Clean desktop with game shortcuts and tools
- **Development** - Project-specific desktops with relevant documentation and tools  
- **Presentation Mode** - Clean, professional desktop for screen sharing

##  Advanced Features

### Command Line Usage
```bash
# Open folder as temporary desktop
python app.py --temp "C:\Path\To\Folder"

# Create new profile from folder
python app.py --new "C:\Path\To\Folder"
```

## How It Works

Each profile consists of three components:

1. **Icon Layout** - Exact positions of desktop icons (saved in Windows registry format)
2. **Desktop Folder** - The folder that acts as your desktop
3. **Windows Theme** - Visual theme applied with the profile

When you switch profiles, the application:
- Updates Windows registry to change desktop folder location
- Restores saved icon positions 
- Applies the associated theme
- Restarts Windows Explorer to apply changes

### File Structure
Data is stored in `C:\Users\[Username]\DesktopProfiles\`:
```
DesktopProfiles/
├── active_profile.txt              # Currently active profile
├── [profile_name].reg             # Desktop icon positions  
├── [profile_name].theme           # Windows theme file
├── [profile_name]_desktop.txt     # Desktop folder path
└── .temp_active                   # Temporary profile marker
```

##  Troubleshooting

### Profile Not Loading
- Wait for Windows Explorer to fully restart
- Ensure the desktop folder path still exists  
- Try running as administrator

### Theme Not Applying
- Windows may require manual theme confirmation
- Check if theme file exists in profile folder
- Some Windows versions have theme restrictions

### Context Menu Missing
- Re-run registration as administrator: `--register`
- Check Windows registry permissions
- Restart Windows Explorer: `taskkill /f /im explorer.exe && start explorer.exe`

### Emergency Reset
If something goes wrong, restore default desktop:
1. Launch the application
2. Click **"Restore Default Desktop"**
3. Or manually delete the `DesktopProfiles` folder

## Privacy & Security

- **100% Local** - No network communication, all data stays on your computer
- **Registry Safe** - Only modifies user-specific registry keys (no system changes)
- **Fully Reversible** - All changes can be undone
- **Non-Destructive** - Original files and settings are preserved

## Contributing

Contributions welcome! 

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Support

If this project helped you, please consider giving it a star! 

For issues or feature requests, please [open an issue](../../issues).
