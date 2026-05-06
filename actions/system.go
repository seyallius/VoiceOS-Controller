package actions

import (
	"fmt"
	"os/exec"
	"syscall"
	"unsafe"
)

func SystemCommand(cmdType string) {
	switch cmdType {
	case "lock":
		LockWorkstation()
	case "shutdown":
		Shutdown()
	case "restart":
		Restart()
	case "sleep":
		Sleep()
	default:
		fmt.Printf("⚠️ Unknown system command: %s\n", cmdType)
	}
}

func LockWorkstation() {
	fmt.Println("🔒 Locking workstation...")
	user32 := syscall.NewLazyDLL("user32.dll")
	lock := user32.NewProc("LockWorkStation")
	lock.Call()
}

func Shutdown() {
	fmt.Println("🛑 Shutting down in 10 seconds...")
	exec.Command("shutdown", "/s", "/t", "10").Run()
}

func Restart() {
	fmt.Println("🔄 Restarting in 10 seconds...")
	exec.Command("shutdown", "/r", "/t", "10").Run()
}

func Sleep() {
	fmt.Println("💤 Sleeping...")
	powrprof := syscall.NewLazyDLL("powrprof.dll")
	setSuspend := powrprof.NewProc("SetSuspendState")
	setSuspend.Call(0, 0, 0)
}

func TakeScreenshot() {
	fmt.Println("📸 Taking screenshot...")
	exec.Command("powershell", "-Command", `
Add-Type -AssemblyName System.Windows.Forms
$screen = [System.Windows.Forms.SystemInformation]::VirtualScreen
$bitmap = New-Object System.Drawing.Bitmap $screen.Width, $screen.Height
$graphics = [System.Drawing.Graphics]::FromImage($bitmap)
$graphics.CopyFromScreen($screen.X, $screen.Y, 0, 0, $bitmap.Size)
$path = "$env:USERPROFILE\Desktop\screenshot_$(Get-Date -Format 'yyyyMMdd_HHmmss').png"
$bitmap.Save($path)
Write-Output "Screenshot saved to $path"
`).Run()
	fmt.Println("✅ Screenshot saved on Desktop")
}