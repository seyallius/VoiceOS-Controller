package speech

import (
    "fmt"
    "os/exec"
    "strings"
    "sync"
    "time"
)

type Recognizer struct {
    Commands    chan string
    isListening bool
    mutex       sync.Mutex
}

func NewRecognizer() (*Recognizer, error) {
    return &Recognizer{
        Commands:    make(chan string, 10),
        isListening: false,
    }, nil
}

func (r *Recognizer) StartListening() error {
    r.mutex.Lock()
    r.isListening = true
    r.mutex.Unlock()

    fmt.Println("🎤 Using Windows Speech Recognition (command mode)...")
    fmt.Println("💡 Say: 'open project', 'close project', 'next slide', 'previous slide', or 'stop'")

    go func() {
        for r.isListening {
            psCommand := `
                Add-Type -AssemblyName System.Speech;
                $recognizer = New-Object System.Speech.Recognition.SpeechRecognitionEngine;

                # Define allowed commands
                $choices = New-Object System.Speech.Recognition.Choices;
                $choices.Add("open project");
                $choices.Add("close project");
                $choices.Add("next slide");
                $choices.Add("previous slide");
                $choices.Add("stop");

                # Build grammar
                $gb = New-Object System.Speech.Recognition.GrammarBuilder;
                $gb.Append($choices);
                $grammar = New-Object System.Speech.Recognition.Grammar($gb);

                $recognizer.LoadGrammar($grammar);
                $recognizer.SetInputToDefaultAudioDevice();
                $result = $recognizer.Recognize();
                if ($result) { Write-Output $result.Text }
            `
            cmd := exec.Command("powershell", "-Command", psCommand)
            out, err := cmd.Output()
            if err != nil {
                fmt.Printf("⚠️  Speech recognition error: %v\n", err)
                time.Sleep(2 * time.Second)
                continue
            }

            recognized := strings.ToLower(strings.TrimSpace(string(out)))
            if recognized != "" {
                fmt.Printf("🎯 Voice recognized: %s\n", recognized)
                r.sendCommand(recognized)
            } else {
                fmt.Println("🕓 (no speech detected)")
            }

            time.Sleep(1 * time.Second)
        }
    }()

    return nil
}

func (r *Recognizer) sendCommand(text string) {
    select {
    case r.Commands <- text:
    default:
        fmt.Println("⚠️ Command channel busy, skipping:", text)
    }
}

func (r *Recognizer) StopListening() {
    r.mutex.Lock()
    r.isListening = false
    r.mutex.Unlock()
}

func (r *Recognizer) Close() {
    r.StopListening()
    close(r.Commands)
}