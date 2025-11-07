package main

import (
    "fmt"
    "log"
    "os"
    "os/signal"
    "strings"
    "syscall"
    "time"
    "path/filepath"
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
        "open project":   OpenProject,
        "close project":  CloseProject,
        "next":           NextSlide,      // Swapped: next -> previous slide
        "previous":       PreviousSlide,          // Swapped: previous -> next slide
        "stop":           Stop,
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
        fmt.Println("ℹ️ PowerPoint is not currently running")
        fmt.Println(" Use 'open project' command to start it")
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
            handleNextSlide(pptCtrl)
        case PreviousSlide:
            handlePreviousSlide(pptCtrl)
        case Stop:
            fmt.Println("🛑 Stopping program...")
            handleStop(pptCtrl)
        }
    } else {
        fmt.Printf("❌ Unknown command: '%s'\n", command)
        fmt.Println("💡 Available commands: 'open project', 'close project', 'next', 'previous', 'stop'")
    }
}

func handleOpenProject(pptCtrl **powerpoint.Controller) {
    // Close existing PowerPoint if any
    if *pptCtrl != nil {
        (*pptCtrl).Close()
        *pptCtrl = nil
        time.Sleep(2 * time.Second)
    }

    // Check if file exists
    if _, err := os.Stat(DefaultConfig.PowerPointFilePath); os.IsNotExist(err) {
        fmt.Printf("❌ PowerPoint file not found: %s\n", DefaultConfig.PowerPointFilePath)
        return
    }

    fmt.Printf("📂 Opening file: %s\n", DefaultConfig.PowerPointFilePath)

    // Open PowerPoint with file (with fallback)
    newPptCtrl, err := powerpoint.OpenPowerPointWithFile(DefaultConfig.PowerPointFilePath)
    if err != nil {
        fmt.Printf("⚠️ Could not connect via COM after open: %v\n", err)
        fmt.Println("💡 File should still be open; use 'next/previous' with keyboard fallback")
        // Don't set pptCtrl to nil – fallback might have opened it
        return
    }

    *pptCtrl = newPptCtrl
    fmt.Println("✅ Successfully connected to PowerPoint!")
}

func handleCloseProject(pptCtrl **powerpoint.Controller) {
    filename := filepath.Base(DefaultConfig.PowerPointFilePath)
    filename = strings.TrimSuffix(filename, filepath.Ext(filename)) // e.g., "pro"

    if *pptCtrl != nil {
        if err := (*pptCtrl).ClosePowerPoint(); err != nil {
            fmt.Printf("❌ Error closing via COM: %v\n", err)
        } else {
            fmt.Println("✅ Closed via COM")
            *pptCtrl = nil
            return
        }
        *pptCtrl = nil
    }
    // Always try process close if COM failed or no connection
    fmt.Printf("🔒 Attempting to close presentation '%s'...\n", filename)
    if err := powerpoint.ClosePowerPointProcess(filename); err != nil {
        fmt.Printf("❌ Process close failed: %v\n", err)
    } else {
        fmt.Println("✅ Specific presentation closed")
    }
}

func handleNextSlide(pptCtrl **powerpoint.Controller) {
    fmt.Println("➡️ Going to next slide...")

    if *pptCtrl != nil {
        if err := (*pptCtrl).NextSlide(); err != nil {
            fmt.Printf("⚠️ COM navigation failed: %v, falling back to keyboard\n", err)
            if err := powerpoint.SendRightArrow(); err != nil {
                fmt.Printf("❌ Keyboard fallback failed: %v\n", err)
            } else {
                fmt.Println("✅ Moved to next slide via keyboard")
            }
        } else {
            fmt.Println("✅ Moved to next slide via COM")
        }
    } else {
        fmt.Println("⚠️ No COM connection, using keyboard fallback...")
        if err := powerpoint.SendRightArrow(); err != nil {
            fmt.Printf("❌ Keyboard fallback failed: %v\n", err)
        } else {
            fmt.Println("✅ Moved to next slide via keyboard")
        }
    }
}

func handlePreviousSlide(pptCtrl **powerpoint.Controller) {
    fmt.Println("⬅️ Going to previous slide...")

    if *pptCtrl != nil {
        if err := (*pptCtrl).PreviousSlide(); err != nil {
            fmt.Printf("⚠️ COM navigation failed: %v, falling back to keyboard\n", err)
            if err := powerpoint.SendLeftArrow(); err != nil {
                fmt.Printf("❌ Keyboard fallback failed: %v\n", err)
            } else {
                fmt.Println("✅ Moved to previous slide via keyboard")
            }
        } else {
            fmt.Println("✅ Moved to previous slide via COM")
        }
    } else {
        fmt.Println("⚠️ No COM connection, using keyboard fallback...")
        if err := powerpoint.SendLeftArrow(); err != nil {
            fmt.Printf("❌ Keyboard fallback failed: %v\n", err)
        } else {
            fmt.Println("✅ Moved to previous slide via keyboard")
        }
    }
}

func handleStop(pptCtrl **powerpoint.Controller) {
    fmt.Println("👋 Shutting down...")
    if *pptCtrl != nil {
        (*pptCtrl).Close()
    }
    os.Exit(0)
}