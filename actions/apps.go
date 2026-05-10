package actions

import (
	"fmt"
	"os/exec"
	"strings"
	"time"
)

func LaunchApp(appName string) {
	fmt.Printf("🚀 Launching: %s\n", appName)
	cmd := exec.Command("cmd", "/C", "start", "", appName)
	if err := cmd.Run(); err != nil {
		cmd = exec.Command("powershell", "-Command", "Start-Process", appName)
		if err := cmd.Run(); err != nil {
			fmt.Printf("❌ Failed to launch %s: %v\n", appName, err)
			return
		}
	}
	fmt.Printf("✅ %s launched\n", appName)
}

func CloseApp(appName string) {
	fmt.Printf("🔒 Closing: %s\n", appName)
	psCommand := fmt.Sprintf(`
$proc = Get-Process | Where-Object { $_.ProcessName -like "*%s*" -or $_.MainWindowTitle -like "*%s*" }
if ($proc) { $proc.CloseMainWindow() ; Start-Sleep -Seconds 1 ; if (!$proc.HasExited) { Stop-Process -Id $proc.Id -Force } }
`, appName, appName)
	cmd := exec.Command("powershell", "-Command", psCommand)
	if err := cmd.Run(); err != nil {
		fmt.Printf("❌ Close failed: %v\n", err)
		return
	}
	time.Sleep(500 * time.Millisecond)
	fmt.Println("✅ Closed")
}