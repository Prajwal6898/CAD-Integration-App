# CAD Integration Application

A WPF application that provides a compact, vertical interface for integrating with AutoCAD and ZWCAD through COM automation.

## Features

- **Multi-CAD Support**: Connects to both AutoCAD and ZWCAD
- **Version Flexibility**: Supports multiple AutoCAD versions (2025, 2024.1, 2024, and generic)
- **Compact UI**: Vertical layout (320x750) positioned on the right side of the screen
- **COM Automation**: Robust connection handling with retry logic and error management
- **Basic CAD Operations**: Drawing lines, inserting blocks, and document information retrieval

## Architecture

### Main Components

- **MainWindow.xaml**: Welcome screen with vertical layout
- **PopupWindow.xaml**: Main application interface
- **AutoCADConnection.cs**: Core COM automation class for CAD integration

### Connection Strategy

1. **Priority Order**:
   - Connect to existing AutoCAD instance
   - Connect to existing ZWCAD instance
   - Start new AutoCAD instance
   - Start new ZWCAD instance

2. **Supported AutoCAD Versions**:
   - AutoCAD 2025 (ProgID: AutoCAD.Application.25)
   - AutoCAD 2024.1 (ProgID: AutoCAD.Application.24.1)
   - AutoCAD 2024 (ProgID: AutoCAD.Application.24)
   - Generic AutoCAD (ProgID: AutoCAD.Application)

3. **ZWCAD Support**:
   - ZWCAD (ProgID: ZwCAD.Application)

## Requirements

- .NET Framework 4.7.2 or higher
- Windows operating system
- AutoCAD or ZWCAD installed
- Visual Studio 2019 or higher (for development)

## Building

```bash
# Using MSBuild
MSBuild.exe CAD.sln /p:Configuration=Release

# Using dotnet CLI
dotnet build CAD\CAD.csproj --configuration Release
```

## Usage

1. Launch the application
2. Click "Launch CAD Interface" on the welcome screen
3. The application will automatically attempt to connect to available CAD software
4. Use the interface to perform CAD operations

## UI Layout

- **Dimensions**: 320x750 pixels
- **Position**: Right side of screen (Left: 1380, Top: 80)
- **Style**: Borderless with drop shadow
- **Layout**: Vertical Grid structure with header and content sections

## Development

### Project Structure

```
CAD/
├── CAD/
│   ├── MainWindow.xaml          # Welcome screen
│   ├── PopupWindow.xaml         # Main interface
│   ├── AutoCADConnection.cs     # COM automation
│   ├── App.xaml                 # Application entry
│   └── Properties/              # Assembly info
└── CAD.sln                      # Solution file
```

### Key Classes

- `AutoCADConnection`: Handles all CAD software communication
- `MainWindow`: Welcome screen with launch button
- `PopupWindow`: Main application interface with CAD operations

## Error Handling

- COM exception handling for unavailable CAD instances
- Retry logic for application startup delays
- Proper COM object cleanup and disposal
- Connection status monitoring

## License

This project is for internal use and CAD integration purposes.