package powerpoint

import (
    "fmt"
    "os"
    "os/exec"
    "strings"
    "syscall"
    "time"

    "github.com/go-ole/go-ole"
    "github.com/go-ole/go-ole/oleutil"
)

// Default configuration
var DefaultConfig = struct {
    PowerPointFilePath string
}{
    PowerPointFilePath: `C:\Users\Al-Khalsi\Desktop\pro.pptx`,
}

type Controller struct {
    pptApp *ole.IDispatch
}

// NewController creates a new PowerPoint controller
func NewController() (*Controller, error) {
    // Initialize COM
    err := ole.CoInitialize(0)
    if err != nil {
        // Check if it's just "COM already initialized" warning
        if err.Error() != "CoInitialize has already been called." {
            return nil, fmt.Errorf("failed to initialize COM: %v", err)
        }
        // If COM is already initialized, we can continue
    }

    // Try to get an existing PowerPoint application instance
    unknown, err := oleutil.CreateObject("PowerPoint.Application")
    if err != nil {
        ole.CoUninitialize()
        return nil, fmt.Errorf("failed to connect to PowerPoint: %v", err)
    }

    pptApp, err := unknown.QueryInterface(ole.IID_IDispatch)
    if err != nil {
        unknown.Release()
        ole.CoUninitialize()
        return nil, fmt.Errorf("failed to get PowerPoint interface: %v", err)
    }

    // Make PowerPoint visible
    oleutil.PutProperty(pptApp, "Visible", true)

    fmt.Println("✅ Connected to PowerPoint application")

    return &Controller{pptApp: pptApp}, nil
}

// NextSlide goes to the next slide
func (c *Controller) NextSlide() error {
    if c.pptApp == nil {
        return fmt.Errorf("PowerPoint application not initialized")
    }

    // Try multiple methods to advance to next slide

    // Method 1: Try to get active presentation and slide show
    activePres, err := oleutil.GetProperty(c.pptApp, "ActivePresentation")
    if err == nil {
        defer activePres.Clear()
        
        slideShowWindow, err := oleutil.GetProperty(activePres.ToIDispatch(), "SlideShowWindow")
        if err == nil {
            defer slideShowWindow.Clear()
            
            if slideShowWindow.Value() != nil {
                view, err := oleutil.GetProperty(slideShowWindow.ToIDispatch(), "View")
                if err == nil {
                    defer view.Clear()
                    _, err = oleutil.CallMethod(view.ToIDispatch(), "Next")
                    if err == nil {
                        return nil
                    }
                }
            }
        }
    }

    // Method 2: Try to use active window
    activeWindow, err := oleutil.GetProperty(c.pptApp, "ActiveWindow")
    if err == nil {
        defer activeWindow.Clear()
        
        if activeWindow.Value() != nil {
            view, err := oleutil.GetProperty(activeWindow.ToIDispatch(), "View")
            if err == nil {
                defer view.Clear()
                _, err = oleutil.CallMethod(view.ToIDispatch(), "Next")
                if err == nil {
                    return nil
                }
            }
        }
    }

    // Method 3: Send right arrow key as fallback
    fmt.Println("⚠️  Using keyboard fallback for next slide")
    return c.sendKeyCommand(0x27) // RIGHT arrow key
}

// PreviousSlide goes to the previous slide
func (c *Controller) PreviousSlide() error {
    if c.pptApp == nil {
        return fmt.Errorf("PowerPoint application not initialized")
    }

    // Try multiple methods to go to previous slide

    // Method 1: Try to get active presentation and slide show
    activePres, err := oleutil.GetProperty(c.pptApp, "ActivePresentation")
    if err == nil {
        defer activePres.Clear()
        
        slideShowWindow, err := oleutil.GetProperty(activePres.ToIDispatch(), "SlideShowWindow")
        if err == nil {
            defer slideShowWindow.Clear()
            
            if slideShowWindow.Value() != nil {
                view, err := oleutil.GetProperty(slideShowWindow.ToIDispatch(), "View")
                if err == nil {
                    defer view.Clear()
                    _, err = oleutil.CallMethod(view.ToIDispatch(), "Previous")
                    if err == nil {
                        return nil
                    }
                }
            }
        }
    }

    // Method 2: Try to use active window
    activeWindow, err := oleutil.GetProperty(c.pptApp, "ActiveWindow")
    if err == nil {
        defer activeWindow.Clear()
        
        if activeWindow.Value() != nil {
            view, err := oleutil.GetProperty(activeWindow.ToIDispatch(), "View")
            if err == nil {
                defer view.Clear()
                _, err = oleutil.CallMethod(view.ToIDispatch(), "Previous")
                if err == nil {
                    return nil
                }
            }
        }
    }

    // Method 3: Send left arrow key as fallback
    fmt.Println("⚠️  Using keyboard fallback for previous slide")
    return c.sendKeyCommand(0x25) // LEFT arrow key
}

// OpenPowerPointFile opens a specific PowerPoint file
func (c *Controller) OpenPowerPointFile(filePath string) error {
    if c.pptApp == nil {
        return fmt.Errorf("PowerPoint application not initialized")
    }

    // Check if file exists
    if _, err := os.Stat(filePath); os.IsNotExist(err) {
        return fmt.Errorf("PowerPoint file not found: %s", filePath)
    }

    presentations, err := oleutil.GetProperty(c.pptApp, "Presentations")
    if err != nil {
        return fmt.Errorf("error accessing presentations: %v", err)
    }
    defer presentations.Clear()

    // Open the presentation file
    presentation, err := oleutil.CallMethod(presentations.ToIDispatch(), "Open", filePath)
    if err != nil {
        return fmt.Errorf("error opening presentation: %v", err)
    }
    defer presentation.Clear()

    fmt.Printf("✅ Opened presentation: %s\n", filePath)

    // Try to start slide show
    pres := presentation.ToIDispatch()
    slideShowSettings, err := oleutil.GetProperty(pres, "SlideShowSettings")
    if err != nil {
        fmt.Printf("⚠️  Could not start slideshow: %v\n", err)
        return nil
    }
    defer slideShowSettings.Clear()

    _, err = oleutil.CallMethod(slideShowSettings.ToIDispatch(), "Run")
    if err != nil {
        fmt.Printf("⚠️  Could not run slideshow: %v\n", err)
        // This is not critical - presentation is still open
    } else {
        fmt.Println("✅ Slideshow started")
    }

    return nil
}

// StartSlideShow starts the slide show for the active presentation
func (c *Controller) StartSlideShow() error {
    if c.pptApp == nil {
        return fmt.Errorf("PowerPoint application not initialized")
    }

    activePres, err := oleutil.GetProperty(c.pptApp, "ActivePresentation")
    if err != nil {
        return fmt.Errorf("no active presentation: %v", err)
    }
    defer activePres.Clear()

    slideShowSettings, err := oleutil.GetProperty(activePres.ToIDispatch(), "SlideShowSettings")
    if err != nil {
        return fmt.Errorf("error getting slideshow settings: %v", err)
    }
    defer slideShowSettings.Clear()

    _, err = oleutil.CallMethod(slideShowSettings.ToIDispatch(), "Run")
    if err != nil {
        return fmt.Errorf("error starting slideshow: %v", err)
    }

    fmt.Println("✅ Slideshow started")
    return nil
}

// OpenPowerPointWithFile opens PowerPoint with a specific file using Windows start command
func OpenPowerPointWithFile(filePath string) error {
    // Check if file exists
    if _, err := os.Stat(filePath); os.IsNotExist(err) {
        return fmt.Errorf("PowerPoint file not found: %s", filePath)
    }

    fmt.Printf("🔍 Opening PowerPoint file: %s\n", filePath)

    // Use PowerShell to open the file (more reliable than cmd)
    cmd := exec.Command("powershell", "-Command", "Start-Process", filePath)
    if err := cmd.Run(); err != nil {
        // Fallback to cmd
        cmd = exec.Command("cmd", "/C", "start", "", filePath)
        if err := cmd.Run(); err != nil {
            return fmt.Errorf("error opening PowerPoint file: %v", err)
        }
    }

    fmt.Println("✅ PowerPoint file opened successfully")
    return nil
}

// OpenPowerPoint opens PowerPoint application directly
func OpenPowerPoint() error {
    fmt.Println("🔍 Starting PowerPoint application...")
    
    // Try to open PowerPoint executable directly
    cmd := exec.Command("powerpnt.exe")
    if err := cmd.Start(); err != nil {
        return fmt.Errorf("error starting PowerPoint: %v", err)
    }

    fmt.Println("✅ PowerPoint application started")
    return nil
}

// ClosePowerPoint closes the PowerPoint application through COM
func (c *Controller) ClosePowerPoint() error {
    if c.pptApp != nil {
        fmt.Println("🔒 Closing PowerPoint via COM...")
        
        // Try to quit gracefully
        _, err := oleutil.CallMethod(c.pptApp, "Quit")
        if err != nil {
            fmt.Printf("⚠️  Error quitting PowerPoint via COM: %v\n", err)
        }
        
        c.pptApp.Release()
        c.pptApp = nil
        ole.CoUninitialize()
        
        fmt.Println("✅ PowerPoint closed via COM")
        return nil
    }
    return fmt.Errorf("PowerPoint controller not initialized")
}

// IsPowerPointRunning checks if PowerPoint is running
func IsPowerPointRunning() bool {
    cmd := exec.Command("tasklist", "/FI", "IMAGENAME eq POWERPNT.EXE")
    output, err := cmd.Output()
    if err != nil {
        return false
    }
    
    outputStr := string(output)
    // Check if PowerPoint process is in the tasklist output
    return len(outputStr) > 0 && 
           !strings.Contains(outputStr, "No tasks are running") &&
           strings.Contains(outputStr, "POWERPNT.EXE")
}

// ClosePowerPointProcess closes PowerPoint completely using taskkill
func ClosePowerPointProcess() error {
    fmt.Println("🔒 Closing PowerPoint process...")
    
    if !IsPowerPointRunning() {
        fmt.Println("ℹ️  PowerPoint is not running")
        return nil
    }

    // Try graceful close first
    cmd := exec.Command("taskkill", "/IM", "POWERPNT.EXE")
    if err := cmd.Run(); err == nil {
        time.Sleep(3 * time.Second)
        if !IsPowerPointRunning() {
            fmt.Println("✅ PowerPoint closed gracefully")
            return nil
        }
    }

    // Force close if needed
    fmt.Println("⚠️  Force closing PowerPoint...")
    cmd = exec.Command("taskkill", "/F", "/IM", "POWERPNT.EXE")
    if err := cmd.Run(); err != nil {
        return fmt.Errorf("error closing PowerPoint: %v", err)
    }

    time.Sleep(2 * time.Second)
    if !IsPowerPointRunning() {
        fmt.Println("✅ PowerPoint closed forcefully")
        return nil
    }

    return fmt.Errorf("could not close PowerPoint - process still running")
}

// sendKeyCommand sends keyboard commands to PowerPoint using Win32 API
func (c *Controller) sendKeyCommand(keyCode uint16) error {
    user32 := syscall.NewLazyDLL("user32.dll")
    keybdEvent := user32.NewProc("keybd_event")
    
    // Key down
    _, _, err := keybdEvent.Call(uintptr(keyCode), 0, 0, 0)
    if err != nil {
        // Check if it's the "success" error
        if err.Error() != "The operation completed successfully." {
            return fmt.Errorf("error sending key down: %v", err)
        }
    }
    
    time.Sleep(100 * time.Millisecond)
    
    // Key up
    _, _, err = keybdEvent.Call(uintptr(keyCode), 0, 2, 0)
    if err != nil {
        // Check if it's the "success" error
        if err.Error() != "The operation completed successfully." {
            return fmt.Errorf("error sending key up: %v", err)
        }
    }
    
    return nil
}

// GetActivePresentationName gets the name of the active presentation
func (c *Controller) GetActivePresentationName() (string, error) {
    if c.pptApp == nil {
        return "", fmt.Errorf("PowerPoint application not initialized")
    }

    activePres, err := oleutil.GetProperty(c.pptApp, "ActivePresentation")
    if err != nil {
        return "", fmt.Errorf("no active presentation: %v", err)
    }
    defer activePres.Clear()

    name, err := oleutil.GetProperty(activePres.ToIDispatch(), "Name")
    if err != nil {
        return "", fmt.Errorf("error getting presentation name: %v", err)
    }
    defer name.Clear()

    return name.ToString(), nil
}

// GetCurrentSlideNumber gets the current slide number
func (c *Controller) GetCurrentSlideNumber() (int, error) {
    if c.pptApp == nil {
        return 0, fmt.Errorf("PowerPoint application not initialized")
    }

    activeWindow, err := oleutil.GetProperty(c.pptApp, "ActiveWindow")
    if err != nil {
        return 0, fmt.Errorf("error getting active window: %v", err)
    }
    defer activeWindow.Clear()

    view, err := oleutil.GetProperty(activeWindow.ToIDispatch(), "View")
    if err != nil {
        return 0, fmt.Errorf("error getting view: %v", err)
    }
    defer view.Clear()

    slide, err := oleutil.GetProperty(view.ToIDispatch(), "CurrentShowPosition")
    if err != nil {
        // Alternative method for normal view
        selection, err := oleutil.GetProperty(activeWindow.ToIDispatch(), "Selection")
        if err != nil {
            return 0, fmt.Errorf("error getting selection: %v", err)
        }
        defer selection.Clear()

        slideRange, err := oleutil.GetProperty(selection.ToIDispatch(), "SlideRange")
        if err != nil {
            return 0, fmt.Errorf("error getting slide range: %v", err)
        }
        defer slideRange.Clear()

        slideNumber, err := oleutil.GetProperty(slideRange.ToIDispatch(), "SlideNumber")
        if err != nil {
            return 0, fmt.Errorf("error getting slide number: %v", err)
        }
        defer slideNumber.Clear()

        return int(slideNumber.Val), nil
    }

    return int(slide.Val), nil
}

// Close cleans up the controller
func (c *Controller) Close() {
    if c.pptApp != nil {
        fmt.Println("🧹 Cleaning up PowerPoint controller...")
        c.pptApp.Release()
        c.pptApp = nil
    }
    ole.CoUninitialize()
}

// Helper function to check if file exists
func fileExists(filename string) bool {
    info, err := os.Stat(filename)
    if os.IsNotExist(err) {
        return false
    }
    return !info.IsDir()
}