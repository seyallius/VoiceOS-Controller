package actions

import (
	"fmt"
	"syscall"
	"time"
	"unicode/utf16"
)

func TypeText(text string) {
	fmt.Printf("⌨️ Typing: %s\n", text)
	for _, ch := range text {
		sendChar(ch)
		time.Sleep(50 * time.Millisecond)
	}
}

func sendChar(r rune) {
	user32 := syscall.NewLazyDLL("user32.dll")
	keybd := user32.NewProc("keybd_event")
	unicodeStr := string(r)
	chars := utf16.Encode([]rune(unicodeStr))
	for _, c := range chars {
		keybd.Call(uintptr(c), 0, 0, 0)
		keybd.Call(uintptr(c), 0, 2, 0)
	}
}