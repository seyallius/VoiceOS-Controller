package powerpoint

import (
    "fmt"
    "os"
    "os/exec"
    "strings"
	"path/filepath"
    "syscall"
    "time"
    "unsafe"

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

// NewController creates a new PowerPoint controller (for existing instance)
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

// OpenPowerPointWithFile opens PowerPoint with a specific file with fallback to exec.Command if COM fails
func OpenPowerPointWithFile(filePath string) (*Controller, error) {
    // Check if file exists
    if _, err := os.Stat(filePath); os.IsNotExist(err) {
        return nil, fmt.Errorf("PowerPoint file not found: %s", filePath)
    }

    fmt.Printf("🔍 Trying to open PowerPoint file via COM: %s\n", filePath)

    // First, try COM method
    var pptApp *ole.IDispatch
    var err error

    // Initialize COM
    err = ole.CoInitialize(0)
    if err != nil && err.Error() != "CoInitialize has already been called." {
        fmt.Printf("⚠️ COM init failed, falling back to exec: %v\n", err)
        return fallbackOpenPowerPointWithFile(filePath)
    }

    // Create PowerPoint application object via COM
    unknown, comErr := oleutil.CreateObject("PowerPoint.Application")
    if comErr != nil {
        fmt.Printf("⚠️ COM object creation failed: %v, falling back to exec\n", comErr)
        ole.CoUninitialize()
        return fallbackOpenPowerPointWithFile(filePath)
    }
    defer unknown.Release()

    pptApp, err = unknown.QueryInterface(ole.IID_IDispatch)
    if err != nil {
        fmt.Printf("⚠️ COM interface failed: %v, falling back to exec\n", err)
        ole.CoUninitialize()
        return fallbackOpenPowerPointWithFile(filePath)
    }

    // Make PowerPoint visible
    oleutil.PutProperty(pptApp, "Visible", true)

    // Get Presentations collection
    presentations, presErr := oleutil.GetProperty(pptApp, "Presentations")
    if presErr != nil {
        fmt.Printf("⚠️ Presentations access failed: %v, falling back to exec\n", presErr)
        pptApp.Release()
        ole.CoUninitialize()
        return fallbackOpenPowerPointWithFile(filePath)
    }
    defer presentations.Clear()

    // Open the presentation file via COM
    presentation, openErr := oleutil.CallMethod(presentations.ToIDispatch(), "Open", filePath)
    if openErr != nil {
        fmt.Printf("⚠️ COM open failed: %v, falling back to exec\n", openErr)
        pptApp.Release()
        ole.CoUninitialize()
        return fallbackOpenPowerPointWithFile(filePath)
    }
    defer presentation.Clear()

    fmt.Printf("✅ Opened presentation via COM: %s\n", filePath)

    // Optional: Start slideshow
    pres := presentation.ToIDispatch()
    slideShowSettings, ssErr := oleutil.GetProperty(pres, "SlideShowSettings")
    if ssErr == nil {
        defer slideShowSettings.Clear()
        _, runErr := oleutil.CallMethod(slideShowSettings.ToIDispatch(), "Run")
        if runErr != nil {
            fmt.Printf("⚠️ Could not start slideshow: %v\n", runErr)
        } else {
            fmt.Println("✅ Slideshow started via COM")
        }
    } else {
        fmt.Printf("⚠️ Could not get slideshow settings: %v\n", ssErr)
    }

    fmt.Println("✅ PowerPoint file opened and connected via COM")
    return &Controller{pptApp: pptApp}, nil
}

// fallbackOpenPowerPointWithFile uses exec.Command to open the file (old method) and tries to connect via NewController
func fallbackOpenPowerPointWithFile(filePath string) (*Controller, error) {
    fmt.Printf("🔍 Falling back to exec method for: %s\n", filePath)

    // Use PowerShell to open the file (more reliable than cmd)
    cmd := exec.Command("powershell", "-Command", "Start-Process", filePath)
    if err := cmd.Run(); err != nil {
        // Fallback to cmd
        cmd = exec.Command("cmd", "/C", "start", "", filePath)
        if err := cmd.Run(); err != nil {
            return nil, fmt.Errorf("error opening PowerPoint file via fallback: %v", err)
        }
    }

    fmt.Println("✅ PowerPoint file opened via fallback exec")

    // Wait for PowerPoint to start
    fmt.Println("⏳ Waiting for PowerPoint to start...")
    time.Sleep(5 * time.Second)

    // Start slideshow with F5 key
    fmt.Println("🎬 Starting slideshow via F5...")
    if err := SendF5(); err != nil {
        fmt.Printf("⚠️ Could not start slideshow: %v\n", err)
    } else {
        fmt.Println("✅ Slideshow started via F5")
    }

    // Try to connect to PowerPoint via COM
    pptCtrl, connectErr := NewController()
    if connectErr != nil {
        fmt.Printf("⚠️ PowerPoint opened via fallback but could not connect via COM: %v\n", connectErr)
        fmt.Println("💡 You can still use 'next slide' and 'previous slide' via keyboard fallback")
        return nil, connectErr
    }

    return pptCtrl, nil
}

// NextSlide goes to the next slide
func (c *Controller) NextSlide() error {
    if c.pptApp == nil {
        return fmt.Errorf("PowerPoint application not initialized")
    }

    // Method 1: Try slideshow view
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

    // Method 2: Try normal view via active window
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

    // Method 3: Fallback to keyboard (right arrow)
    fmt.Println("⚠️ Using keyboard fallback for next slide")
    return c.sendKeyCommand(0x27) // VK_RIGHT
}

// PreviousSlide goes to the previous slide
func (c *Controller) PreviousSlide() error {
    if c.pptApp == nil {
        return fmt.Errorf("PowerPoint application not initialized")
    }

    // Method 1: Try slideshow view
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

    // Method 2: Try normal view via active window
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

    // Method 3: Fallback to keyboard (left arrow)
    fmt.Println("⚠️ Using keyboard fallback for previous slide")
    return c.sendKeyCommand(0x25) // VK_LEFT
}

// ClosePowerPoint closes the PowerPoint application through COM (closes presentation first)
func (c *Controller) ClosePowerPoint() error {
    if c.pptApp == nil {
        return fmt.Errorf("PowerPoint controller not initialized")
    }

    fmt.Println("🔒 Closing PowerPoint via COM...")

    // Close active presentation first
    activePres, err := oleutil.GetProperty(c.pptApp, "ActivePresentation")
    if err == nil {
        defer activePres.Clear()
        _, closeErr := oleutil.CallMethod(activePres.ToIDispatch(), "Close")
        if closeErr != nil {
            fmt.Printf("⚠️ Warning: Could not close presentation: %v\n", closeErr)
        }
    }

    // Quit PowerPoint
    _, err = oleutil.CallMethod(c.pptApp, "Quit")
    if err != nil {
        fmt.Printf("⚠️ Error quitting PowerPoint via COM: %v\n", err)
        return err
    }

    c.pptApp.Release()
    c.pptApp = nil
    ole.CoUninitialize()

    fmt.Println("✅ PowerPoint closed via COM")
    return nil
}

// IsPowerPointRunning checks if PowerPoint is running
func IsPowerPointRunning() bool {
    psCommand := `
        $count = (Get-Process | Where-Object { $_.MainWindowTitle -like "*PowerPoint*" } | Measure-Object).Count
        $count
    `
    cmd := exec.Command("powershell", "-Command", psCommand)
    output, err := cmd.Output()
    if err != nil {
        return false
    }
    countStr := strings.TrimSpace(string(output))
    var count int
    fmt.Sscanf(countStr, "%d", &count)
    return count > 0
}

// ClosePowerPointProcess closes only the PowerPoint presentation with the given filename (via window title)
func ClosePowerPointProcess(filename string) error {
    if filename == "" {
        return fmt.Errorf("no filename provided")
    }

    fmt.Printf("🔒 Closing specific PowerPoint presentation '%s' (fallback via window title)...\n", filename)

    // First, check if it's running with this filename
    checkPsCommand := fmt.Sprintf(`
        $procs = Get-Process | Where-Object { $_.MainWindowTitle -like "*%s*" -or $_.MainWindowTitle -like "*%s.pptx*" }
        $procs | Select-Object Id, ProcessName, MainWindowTitle | ForEach-Object { Write-Output "PID: $($_.Id) | Name: $($_.ProcessName) | Title: $($_.MainWindowTitle)" }
        $count = ($procs | Measure-Object).Count
        Write-Output "Count: $count"
    `, filename, filename)
    checkCmd := exec.Command("powershell", "-Command", checkPsCommand)
    output, checkErr := checkCmd.CombinedOutput()
    outputStr := string(output)
    fmt.Printf("🔍 Check output: %s\n", outputStr) // Debug

    if checkErr != nil {
        fmt.Printf("⚠️ Check error: %v\n", checkErr)
    }

    // Parse count from output
    lines := strings.Split(outputStr, "\n")
    var count int
    for _, line := range lines {
        if strings.HasPrefix(strings.TrimSpace(line), "Count:") {
            fmt.Sscanf(line, "Count: %d", &count)
            break
        }
    }
    if count == 0 {
        fmt.Println("ℹ️ No matching PowerPoint presentation found")
        return nil
    }

    // Graceful close: CloseMainWindow for each matching process
    gracefulPsCommand := fmt.Sprintf(`
        $procs = Get-Process | Where-Object { $_.MainWindowTitle -like "*%s*" -or $_.MainWindowTitle -like "*%s.pptx*" }
        foreach ($proc in $procs) {
            try {
                $proc.CloseMainWindow() | Out-Null
                Start-Sleep -Seconds 2
                if (!$proc.HasExited) {
                    Write-Output "Graceful close failed for PID $($proc.Id), forcing..."
                    $proc.Kill()
                } else {
                    Write-Output "Graceful close succeeded for PID $($proc.Id)"
                }
            } catch {
                Write-Output "Error closing PID $($proc.Id): $_"
                $proc.Kill()
            }
        }
        "Graceful close attempted"
    `, filename, filename)
    gracefulCmd := exec.Command("powershell", "-Command", gracefulPsCommand)
    gOutput, gErr := gracefulCmd.CombinedOutput()
    fmt.Printf("🔍 Graceful close output: %s\n", string(gOutput)) // Debug
    if gErr != nil {
        fmt.Printf("⚠️ Graceful close error: %v\n", gErr)
    } else {
        fmt.Println("✅ Graceful close attempted")
    }

    // Wait
    time.Sleep(3 * time.Second)

    // Force close if still running
    forcePsCommand := fmt.Sprintf(`
        $procs = Get-Process | Where-Object { $_.MainWindowTitle -like "*%s*" -or $_.MainWindowTitle -like "*%s.pptx*" }
        foreach ($proc in $procs) {
            try {
                Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
                Write-Output "Force closed PID $($proc.Id)"
            } catch {
                Write-Output "Failed to force close PID $($proc.Id): $_"
            }
        }
        "Force close attempted"
    `, filename, filename)
    forceCmd := exec.Command("powershell", "-Command", forcePsCommand)
    fOutput, fErr := forceCmd.CombinedOutput()
    fmt.Printf("🔍 Force close output: %s\n", string(fOutput)) // Debug
    if fErr != nil {
        fmt.Printf("❌ Force close error: %v\n", fErr)
        return fmt.Errorf("error force closing specific presentation: %v", fErr)
    }

    time.Sleep(2 * time.Second)

    // Final check
    finalCheckPsCommand := fmt.Sprintf(`
        $procs = Get-Process | Where-Object { $_.MainWindowTitle -like "*%s*" -or $_.MainWindowTitle -like "*%s.pptx*" }
        $count = ($procs | Measure-Object).Count
        Write-Output "Final count: $count"
    `, filename, filename)
    finalCmd := exec.Command("powershell", "-Command", finalCheckPsCommand)
    finalOutput, _ := finalCmd.Output()
    finalCountStr := strings.TrimSpace(string(finalOutput))
    var finalCount int
    fmt.Sscanf(finalCountStr, "Final count: %d", &finalCount)
    if finalCount == 0 {
        fmt.Println("✅ Specific PowerPoint presentation closed successfully")
        return nil
    }

    return fmt.Errorf("could not close specific PowerPoint presentation - still running (final count: %d)", finalCount)
}

// sendKeyCommand sends keyboard commands to PowerPoint using Win32 API
func (c *Controller) sendKeyCommand(keyCode uint16) error {
    user32 := syscall.NewLazyDLL("user32.dll")
    keybdEvent := user32.NewProc("keybd_event")

    // Key down
    _, _, err := keybdEvent.Call(uintptr(keyCode), 0, 0, 0)
    if err != nil && err.Error() != "The operation completed successfully." {
        return fmt.Errorf("error sending key down: %v", err)
    }

    time.Sleep(100 * time.Millisecond)

    // Key up
    _, _, err = keybdEvent.Call(uintptr(keyCode), 0, 2, 0)
    if err != nil && err.Error() != "The operation completed successfully." {
        return fmt.Errorf("error sending key up: %v", err)
    }

    return nil
}

// FocusPowerPoint focuses the PowerPoint window using Win32 API by title (more reliable)
func FocusPowerPoint(filename string) error {
    if filename == "" {
        filename = "pro" // Default for fallback
    }

    user32 := syscall.NewLazyDLL("user32.dll")
    findWindow := user32.NewProc("FindWindowW")
    setForeground := user32.NewProc("SetForegroundWindow")

    // Find PowerPoint window by title containing filename or "PowerPoint"
    titlePattern := fmt.Sprintf("*%s*", filename) // e.g., "*pro*"
    titlePtr, _ := syscall.UTF16PtrFromString(titlePattern)
    hwnd, _, _ := findWindow.Call(0, uintptr(unsafe.Pointer(titlePtr))) // lpClassName = nil, lpWindowName = title

    if hwnd == 0 {
        // Fallback to general "PowerPoint" title
        titlePtr, _ = syscall.UTF16PtrFromString("*PowerPoint*")
        hwnd, _, _ = findWindow.Call(0, uintptr(unsafe.Pointer(titlePtr)))
        if hwnd == 0 {
            return fmt.Errorf("PowerPoint window not found by title '%s' or '*PowerPoint*'", titlePattern)
        }
    }

    // Set foreground
    _, _, err := setForeground.Call(hwnd)
    if err != nil && err != syscall.Errno(0) {
        return fmt.Errorf("failed to focus PowerPoint: %v", err)
    }

    time.Sleep(300 * time.Millisecond) // Longer pause for focus and stability
    return nil
}

// SendF5 sends F5 key to start slideshow
func SendF5() error {
    filename := filepath.Base(DefaultConfig.PowerPointFilePath)
    filename = strings.TrimSuffix(filename, filepath.Ext(filename))
    if err := FocusPowerPoint(filename); err != nil {
        fmt.Printf("⚠️ Could not focus PowerPoint for F5: %v\n", err)
    }

    user32 := syscall.NewLazyDLL("user32.dll")
    keybdEvent := user32.NewProc("keybd_event")

    // F5 down (VK_F5 = 0x74)
    _, _, err := keybdEvent.Call(uintptr(0x74), 0, 0, 0)
    if err != nil && err.Error() != "The operation completed successfully." {
        return fmt.Errorf("error sending F5 down: %v", err)
    }

    time.Sleep(150 * time.Millisecond) // Slightly longer for F5

    // F5 up
    _, _, err = keybdEvent.Call(uintptr(0x74), 0, 2, 0)
    if err != nil && err.Error() != "The operation completed successfully." {
        return fmt.Errorf("error sending F5 up: %v", err)
    }

    return nil
}

// SendRightArrow sends the right arrow key globally (for next slide fallback)
func SendRightArrow() error {
    filename := filepath.Base(DefaultConfig.PowerPointFilePath)
    filename = strings.TrimSuffix(filename, filepath.Ext(filename))
    if err := FocusPowerPoint(filename); err != nil {
        fmt.Printf("⚠️ Could not focus PowerPoint for right arrow: %v\n", err)
    }

    user32 := syscall.NewLazyDLL("user32.dll")
    keybdEvent := user32.NewProc("keybd_event")

    // Key down (VK_RIGHT = 0x27)
    _, _, err := keybdEvent.Call(uintptr(0x27), 0, 0, 0)
    if err != nil && err.Error() != "The operation completed successfully." {
        return fmt.Errorf("error sending right key down: %v", err)
    }

    time.Sleep(150 * time.Millisecond) // Longer sleep to prevent rapid key events

    // Key up
    _, _, err = keybdEvent.Call(uintptr(0x27), 0, 2, 0)
    if err != nil && err.Error() != "The operation completed successfully." {
        return fmt.Errorf("error sending right key up: %v", err)
    }

    return nil
}

// SendLeftArrow sends the left arrow key globally (for previous slide fallback)
func SendLeftArrow() error {
    filename := filepath.Base(DefaultConfig.PowerPointFilePath)
    filename = strings.TrimSuffix(filename, filepath.Ext(filename))
    if err := FocusPowerPoint(filename); err != nil {
        fmt.Printf("⚠️ Could not focus PowerPoint for left arrow: %v\n", err)
    }

    user32 := syscall.NewLazyDLL("user32.dll")
    keybdEvent := user32.NewProc("keybd_event")

    // Key down (VK_LEFT = 0x25)
    _, _, err := keybdEvent.Call(uintptr(0x25), 0, 0, 0)
    if err != nil && err.Error() != "The operation completed successfully." {
        return fmt.Errorf("error sending left key down: %v", err)
    }

    time.Sleep(150 * time.Millisecond) // Longer sleep

    // Key up
    _, _, err = keybdEvent.Call(uintptr(0x25), 0, 2, 0)
    if err != nil && err.Error() != "The operation completed successfully." {
        return fmt.Errorf("error sending left key up: %v", err)
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
        c.ClosePowerPoint() // Use the full close method
    }
}

// Helper function to check if file exists
func fileExists(filename string) bool {
    info, err := os.Stat(filename)
    if os.IsNotExist(err) {
        return false
    }
    return !info.IsDir()
}