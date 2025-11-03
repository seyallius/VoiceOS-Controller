# PowerPoint Voice Controller

A Go application that allows you to control PowerPoint presentations using English voice commands.

## Features

- **Voice Control**: Control PowerPoint with voice commands
- **English Language Support**: Commands in English
- **PowerPoint Automation**: Open, close, and navigate presentations
- **Static File Path**: Configurable path for PowerPoint files

## Supported Commands

### English Commands:
- `open project` - Open PowerPoint file
- `close project` - Close PowerPoint
- `next slide` or `next` - Go to next slide
- `previous slide` or `previous` - Go to previous slide
- `stop` - Stop the program

## Requirements

- Windows OS
- Microsoft PowerPoint installed
- Go 1.21 or later (for building from source)

## Installation

1. Download the `PowerPointVoiceController.exe` executable
2. Make sure you have PowerPoint installed on your system
3. Create a PowerPoint file named `pro.pptx` on your desktop, or modify the static path in the configuration

## Usage

1. Run the executable: `PowerPointVoiceController.exe` or use `run.bat`
2. The application will start and attempt to open PowerPoint file directly
3. Speak one of the supported commands into your microphone
4. The application will execute the corresponding PowerPoint action

## Configuration

The default PowerPoint file path is set to `C:\Users\Al-Khalsi\Desktop\pro.pptx`. You can modify this path directly in the source code if needed.

## How It Works

The application uses Windows Speech API (SAPI) through COM objects to recognize voice commands. It then uses Windows API calls to simulate keyboard input for PowerPoint navigation (RIGHT/LEFT arrow keys) and process management to open/close PowerPoint.

## Development

To build from source:

```bash
go build -o PowerPointVoiceController.exe .