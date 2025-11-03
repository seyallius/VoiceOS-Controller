# PowerPoint Voice Controller
A Go application that allows you to control PowerPoint presentations using English voice commands.

## Features
- **Voice Control**: Control PowerPoint with voice commands
- **English Language Support**: Commands in English
- **PowerPoint Automation**: Open, close, and navigate presentations via COM (direct integration)
- **Static File Path**: Configurable path for PowerPoint files
- **Slideshow Support**: Automatically starts slideshow on open

## Supported Commands
### English Commands:
- `open project` - Open PowerPoint file and start slideshow
- `close project` - Close the opened presentation and quit PowerPoint
- `next slide` or `next` - Go to next slide (COM or keyboard fallback)
- `previous slide` or `previous` - Go to previous slide (COM or keyboard fallback)
- `stop` - Stop the program

## Requirements
- Windows OS
- Microsoft PowerPoint installed
- Go 1.21 or later (for building from source)
- Microphone for voice input

## Installation
1. Download the `PowerPointVoiceController.exe` executable
2. Make sure you have PowerPoint installed on your system
3. Create a PowerPoint file named `pro.pptx` on your desktop, or modify the static path in the configuration (in main.go)

## Usage
1. Run the executable: `PowerPointVoiceController.exe` or use `run.bat`
2. The application will start and wait for commands
3. Speak one of the supported commands into your microphone
4. The application will execute the corresponding PowerPoint action

## Configuration
The default PowerPoint file path is set to `C:\Users\Al-Khalsi\Desktop\pro.pptx`. You can modify this path directly in the source code if needed.

## How It Works
The application uses Windows Speech API (SAPI) through PowerShell for voice recognition. It then uses COM automation (via go-ole) to directly control PowerPoint – opening files, navigating slides, and closing. This ensures reliable connection without process mismatches.

## Development
To build from source:
```bash
go mod init powerpoint-voice-controller
go get github.com/go-ole/go-ole
go build -o PowerPointVoiceController.exe .