package voice

import (
	"encoding/json"
	"fmt"
	"os/exec"
	"strings"
	"sync"
	"time"
)

type Recognizer struct {
	Commands   chan string
	isListening bool
	mutex      sync.Mutex
	config     map[string]interface{}
}

func NewRecognizer(cfg interface{}) (*Recognizer, error) {
	jsonBytes, _ := json.Marshal(cfg)
	var parsed map[string]interface{}
	json.Unmarshal(jsonBytes, &parsed)
	return &Recognizer{
		Commands:    make(chan string, 10),
		isListening: false,
		config:      parsed,
	}, nil
}

func (r *Recognizer) StartListening() error {
	r.mutex.Lock()
	r.isListening = true
	r.mutex.Unlock()

	fmt.Println("🎤 Windows Speech Recognition active")

	go func() {
		for r.isListening {
			cmdList := r.getCommandList()
			if len(cmdList) == 0 {
				time.Sleep(1 * time.Second)
				continue
			}

			psCommand := r.buildPowerShellGrammar(cmdList)
			out, err := exec.Command("powershell", "-Command", psCommand).Output()
			if err != nil {
				time.Sleep(1 * time.Second)
				continue
			}

			recognized := strings.ToLower(strings.TrimSpace(string(out)))
			if recognized != "" {
				fmt.Printf("🎯 Recognized: %s\n", recognized)
				select {
				case r.Commands <- recognized:
				default:
					fmt.Println("⚠️ Command channel full")
				}
			}
			time.Sleep(500 * time.Millisecond)
		}
	}()
	return nil
}

func (r *Recognizer) getCommandList() []string {
	cmds, ok := r.config["commands"].(map[string]interface{})
	if !ok {
		return []string{"stop"}
	}
	keys := make([]string, 0, len(cmds)+1)
	for k := range cmds {
		keys = append(keys, k)
	}
	keys = append(keys, "stop")
	return keys
}

func (r *Recognizer) buildPowerShellGrammar(commands []string) string {
	choices := strings.Join(commands, "\";\n        $choices.Add(\"")
	return fmt.Sprintf(`
Add-Type -AssemblyName System.Speech
$recognizer = New-Object System.Speech.Recognition.SpeechRecognitionEngine
$choices = New-Object System.Speech.Recognition.Choices
$choices.Add("%s")
$gb = New-Object System.Speech.Recognition.GrammarBuilder
$gb.Append($choices)
$grammar = New-Object System.Speech.Recognition.Grammar($gb)
$recognizer.LoadGrammar($grammar)
$recognizer.SetInputToDefaultAudioDevice()
$result = $recognizer.Recognize()
if ($result) { Write-Output $result.Text }
`, choices)
}

func (r *Recognizer) StopListening() {
	r.mutex.Lock()
	defer r.mutex.Unlock()
	r.isListening = false
}

func (r *Recognizer) Close() {
	r.StopListening()
	close(r.Commands)
}