# HELACAD Customization 2026

A VB.NET-based AutoCAD customization plugin designed for structural engineering workflows at Herold Engineering Ltd. (HEL). This plugin integrates seamlessly with AutoCAD 2026 to provide enhanced productivity tools, automated layer management, drawing setup, and specialized palettes for wood and masonry design.

## Features

- **Drawing Startup Automation**: Automated setup of layers, text styles, dimension styles, and system variables for new drawings
- **Layer Management**: Intelligent layer generation and management with HEL-specific naming conventions
- **Title Block Integration**: Dynamic title blocks with automatic line adjustment
- **Palette System**: Dedicated palettes for:
  - Wood construction routines
  - Masonry unit specifications
  - Drawing issue tracking
- **Command Integration**: Custom AutoCAD commands for streamlined workflows
- **Menu System**: Integrated HEL-ACAD menus and toolbars
- **Plot Style Management**: Automatic configuration of plot styles and printer paths

## Prerequisites

- **AutoCAD 2026** (64-bit)
- **.NET Framework 8.0** or later
- **Windows 10/11** (64-bit)
- **Visual Studio 2022** (for development)

## Installation

1. **Build the Project**:
   ```bash
   # Open the solution in Visual Studio
   # Build in Release|x64 configuration
   ```

2. **Deploy the DLL**:
   - Copy `HELACADCUST2026.dll` to your AutoCAD plugins directory
   - Or use the bundled installer for network deployment

3. **Configure Trusted Paths**:
   - Add the plugin directory to AutoCAD's trusted paths
   - Default trusted paths are set automatically on load

4. **Load the Plugin**:
   - The plugin loads automatically when AutoCAD starts
   - Verify loading in the command line: "HEL Customization Loaded"

## Usage

### Basic Workflow

1. **Start AutoCAD**: The customization loads automatically
2. **New Drawing**: Run `DWGStartup` command to initialize HEL standards
3. **Layer Setup**: Use `laygen` command for layer generation
4. **Palettes**: Access wood/masonry routines via `Timber`, `Masonry`, or `Issues` commands

### Key Commands

- `DWGStartup`: Initialize new drawings with HEL standards
- `laygen`: Generate and manage drawing layers
- `typetag`: Launch type tag form
- `Timber`: Load wood construction palette
- `Masonry`: Load masonry units palette
- `Issues`: Load drawing issues palette
- `Titlelines`: Adjust dynamic title block lines
- `FontUpdate`: Update text styles and fonts

### Menu Integration

The customization includes:
- HEL-ACAD-2023 menu system
- SteelPlus menu integration
- Custom toolbars and ribbons

## Project Structure

```
HELACADCUST/
├── ACADCUST.vb              # Main customization class
├── cls*.vb                  # Utility and routine classes
├── frm*.vb                  # Windows Forms
├── usc*.vb                  # User Controls (Palettes)
├── HELACADCUST2026.vbproj   # Project file
└── My Project/              # Application settings
```

## Development

### Building

```bash
# Restore NuGet packages
dotnet restore

# Build for x64 Release
dotnet build --configuration Release --platform x64
```

### Dependencies

- AutoCAD .NET API assemblies
- Windows Forms
- WPF (for UI components)
- COM Interop for AutoCAD integration

## Contributing

This is an internal HEL project. For contributions:

1. Fork the repository
2. Create a feature branch
3. Make changes and test thoroughly
4. Submit a pull request

## License

Internal use only - Herold Engineering Ltd.

## Support

For support or issues:
- Contact the HEL IT department
- Check internal documentation
- Report bugs through the HEL issue tracking system

## Version History

- **2026**: Updated for AutoCAD 2026 compatibility, Converted to .NET 8.0
- **2023**: Updated for AutoCAD 2023 compatibility
- **2019**: Enhanced palette system
- **2016**: Initial VB.NET conversion