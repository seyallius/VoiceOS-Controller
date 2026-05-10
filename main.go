package main

import (
	"encoding/json"
	"fmt"
	"log"
	"os"
	"os/signal"
	"strings"
	"syscall"
	"voice-os-controller/actions"
	"voice-os-controller/voice"
)

type CommandConfig struct {
	Commands map[string]CommandAction `json:"commands"`
}

type CommandAction struct {
	Action string `json:"action"`
	Target string `json:"target,omitempty"`
	Value  string `json:"value,omitempty"`
}

var config CommandConfig

func loadConfig() error {
	file, err := os.ReadFile("config/commands.json")
	if err != nil {
		return err
	}
	return json.Unmarshal(file, &config)
}

func main() {
	fmt.Println("🎙️ Voice OS Controller - Windows")
	fmt.Println("=================================")
	fmt.Println("💡 Say commands from config/commands.json")
	fmt.Println("🛑 Say 'stop' or press Ctrl+C to exit")
	fmt.Println()

	if err := loadConfig(); err != nil {
		log.Fatalf("❌ Failed to load config: %v", err)
	}

	recognizer, err := voice.NewRecognizer(config)
	if err != nil {
		log.Fatalf("❌ Speech init error: %v", err)
	}
	defer recognizer.Close()

	fmt.Println("✅ System ready. Listening...")

	sigChan := make(chan os.Signal, 1)
	signal.Notify(sigChan, os.Interrupt, syscall.SIGTERM)

	go func() {
		for {
			select {
			case cmd := <-recognizer.Commands:
				executeCommand(cmd)
			case <-sigChan:
				fmt.Println("\n👋 Shutting down...")
				os.Exit(0)
			}
		}
	}()

	if err := recognizer.StartListening(); err != nil {
		log.Fatalf("❌ Listening error: %v", err)
	}

	select {}
}

func executeCommand(cmd string) {
	cmd = strings.ToLower(strings.TrimSpace(cmd))
	fmt.Printf("\n🔊 Recognized: '%s'\n", cmd)

	if cmd == "stop" {
		fmt.Println("🛑 Stopping...")
		os.Exit(0)
	}

	actionDef, exists := config.Commands[cmd]
	if !exists {
		fmt.Printf("❌ Unknown command: '%s'\n", cmd)
		return
	}

	switch actionDef.Action {
	case "launch_app":
		actions.LaunchApp(actionDef.Target)
	case "close_app":
		actions.CloseApp(actionDef.Target)
	case "volume":
		actions.VolumeControl(actionDef.Value)
	case "system":
		actions.SystemCommand(actionDef.Value)
	case "screenshot":
		actions.TakeScreenshot()
	case "switch_window":
		actions.SwitchWindow()
	case "type":
		actions.TypeText(actionDef.Value)
	default:
		fmt.Printf("⚠️ Unknown action type: %s\n", actionDef.Action)
	}
}