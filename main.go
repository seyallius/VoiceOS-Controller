package main

import (
    "fmt"
    "log"
    "os"
    "os/signal"
    "strings"
    "syscall"
    "time"

    "powerpoint-voice-controller/powerpoint"
    "powerpoint-voice-controller/speech"
)

// Configuration for the PowerPoint Voice Controller
type Config struct {
    PowerPointFilePath string // Static path to the PowerPoint file
    Commands           map[string]CommandType
}

// CommandType defines the type of commands
type CommandType int

const (
    OpenProject CommandType = iota
    CloseProject
    NextSlide
    PreviousSlide
    Stop
)

// Default configuration
var DefaultConfig = Config{
    PowerPointFilePath: `C:\Users\Al-Khalsi\Desktop\pro.pptx`, // Default static path
    Commands: map[string]CommandType{
        // English commands
        "open project":    OpenProject,
        "close project":   CloseProject,
        "next slide":      NextSlide,
        "previous slide":  PreviousSlide,
        "next":            NextSlide,
        "previous":        PreviousSlide,
        "stop":            Stop,
    },
}

func main() {
    fmt.Println("🎤 Voice Controller for PowerPoint - Go Version")
    fmt.Println("=================================================")
    fmt.Println("")

    var pptCtrl *powerpoint.Controller
    var err error

    // Initialize speech recognition
    speechRec, err := speech.NewRecognizer()
    if err != nil {
        log.Fatalf("❌ Error initializing speech recognition: %v", err)
    }
    defer speechRec.Close()

    // Try to connect to existing PowerPoint instance
    pptCtrl, err = powerpoint.NewController()
    if err != nil {
        fmt.Println("ℹ️  PowerPoint is not currently running")
        fmt.Println("   Use 'open project' command to start it")
    } else {
        fmt.Println("✅ Connected to existing PowerPoint instance")
        defer pptCtrl.Close()
    }

    fmt.Println("✅ System initialized successfully")
    fmt.Println("🎤 Voice recognition ready...")

    // Signal handling for graceful shutdown
    signalChan := make(chan os.Signal, 1)
    signal.Notify(signalChan, os.Interrupt, syscall.SIGTERM)

    // Main command handling loop
    go func() {
        for {
            select {
            case command, ok := <-speechRec.Commands:
                if !ok {
                    return
                }
                handleCommand(command, &pptCtrl)
            case <-signalChan:
                fmt.Println("\n🛑 Exit signal received. Shutting down...")
                if pptCtrl != nil {
                    pptCtrl.Close()
                }
                os.Exit(0)
            }
        }
    }()

    // Start listening for commands
    if err := speechRec.StartListening(); err != nil {
        log.Fatalf("❌ Error starting to listen: %v", err)
    }

    // Keep the program running
    select {}
}

func handleCommand(command string, pptCtrl **powerpoint.Controller) {
    fmt.Printf("\n🔊 Command received: '%s'\n", command)

    // Normalize the command
    normalizedCmd := strings.TrimSpace(strings.ToLower(command))

    // Check against configured commands
    if commandType, exists := DefaultConfig.Commands[normalizedCmd]; exists {
        switch commandType {
        case OpenProject:
            fmt.Println("🚀 Opening PowerPoint project...")
            handleOpenProject(pptCtrl)
            
        case CloseProject:
            fmt.Println("🔒 Closing PowerPoint project...")
            handleCloseProject(pptCtrl)
            
        case NextSlide:
            fmt.Println("➡️  Going to next slide...")
            handleNextSlide(pptCtrl)
            
        case PreviousSlide:
            fmt.Println("⬅️  Going to previous slide...")
            handlePreviousSlide(pptCtrl)
            
        case Stop:
            fmt.Println("🛑 Stopping program...")
            handleStop(pptCtrl)
        }
    } else {
        fmt.Printf("❌ Unknown command: '%s'\n", command)
        fmt.Println("💡 Available commands: 'open project', 'close project', 'next slide', 'previous slide', 'stop'")
    }
}

func handleOpenProject(pptCtrl **powerpoint.Controller) {
    // Close existing PowerPoint if any
    if *pptCtrl != nil {
        (*pptCtrl).Close()
        *pptCtrl = nil
        time.Sleep(2 * time.Second)
    }

    // Try to open PowerPoint with file
    fmt.Printf("📂 Opening file: %s\n", DefaultConfig.PowerPointFilePath)
    
    if err := powerpoint.OpenPowerPointWithFile(DefaultConfig.PowerPointFilePath); err != nil {
        fmt.Printf("❌ Error opening PowerPoint: %v\n", err)
        return
    }

    fmt.Println("⏳ Waiting for PowerPoint to start...")
    time.Sleep(5 * time.Second)

    // Try to connect to PowerPoint
    newPptCtrl, err := powerpoint.NewController()
    if err != nil {
        fmt.Printf("⚠️  PowerPoint opened but could not connect: %v\n", err)
        fmt.Println("💡 You can still use 'next slide' and 'previous slide' commands")
    } else {
        *pptCtrl = newPptCtrl
        fmt.Println("✅ Successfully connected to PowerPoint!")
    }
}

func handleCloseProject(pptCtrl **powerpoint.Controller) {
    if *pptCtrl != nil {
        if err := (*pptCtrl).ClosePowerPoint(); err != nil {
            fmt.Printf("❌ Error closing PowerPoint: %v\n", err)
        } else {
            fmt.Println("✅ PowerPoint closed successfully")
            *pptCtrl = nil
        }
    } else {
        if err := powerpoint.ClosePowerPointProcess(); err != nil {
            fmt.Printf("❌ Error closing PowerPoint process: %v\n", err)
        } else {
            fmt.Println("✅ PowerPoint process closed")
        }
    }
}

func handleNextSlide(pptCtrl **powerpoint.Controller) {
    if *pptCtrl != nil {
        if err := (*pptCtrl).NextSlide(); err != nil {
            fmt.Printf("❌ Error going to next slide: %v\n", err)
        } else {
            fmt.Println("✅ Moved to next slide")
        }
    } else {
        fmt.Println("❌ PowerPoint not available. Use 'open project' first.")
    }
}

func handlePreviousSlide(pptCtrl **powerpoint.Controller) {
    if *pptCtrl != nil {
        if err := (*pptCtrl).PreviousSlide(); err != nil {
            fmt.Printf("❌ Error going to previous slide: %v\n", err)
        } else {
            fmt.Println("✅ Moved to previous slide")
        }
    } else {
        fmt.Println("❌ PowerPoint not available. Use 'open project' first.")
    }
}

func handleStop(pptCtrl **powerpoint.Controller) {
    fmt.Println("👋 Shutting down...")
    if *pptCtrl != nil {
        (*pptCtrl).Close()
    }
    os.Exit(0)
}