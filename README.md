# Voice OS Controller for Windows

Control your Windows PC entirely by voice.

## Features
- Launch any application
- Close applications
- Control system volume
- Lock, shutdown, restart, sleep
- Take screenshots
- Switch windows
- Type text via voice

## Dynamic Configuration
All commands are defined in `config/commands.json`.  
Add, remove, or change commands without recompiling.

## Requirements
- Windows 10/11
- Microphone
- (Optional) nircmd for volume control

## Installation
1. Download or build `voice-os-controller.exe`
2. Edit `config/commands.json` to add your own commands
3. Run `run.bat` or `voice-os-controller.exe`

## Usage
- Speak a command from your config file
- Say `stop` to exit

## Building from Source
```bash
go mod tidy
go build -o voice-os-controller.exe