; Initialize state
global MikeMuted := false

; Speak a message to indicate the script is running
Run, powershell -Command "(New-Object -ComObject SAPI.SPVoice).Speak('Program started.');",, Hide

; Hotkey definition for Left Windows + Shift + M to toggle microphone mute
#+M::ToggleMikeMute()

; Hotkey definition for Left Windows + Shift + Q to exit the script
#+Q::ExitScript()

; Function to toggle mic mute and display status
ToggleMikeMute() {
    global MikeMuted
    MikeMuted := !MikeMuted
    
    ; Mute or unmute microphone using Nircmd (Ensure nircmd.exe is in the same directory as the script)
    if (MikeMuted) {
        RunWait, nircmd.exe mutesysvolume 1 Microphone,, Hide
    } else {
        RunWait, nircmd.exe mutesysvolume 0 Microphone,, Hide
    }

    ; Speak status using Text-to-Speech
    SpeakMikeStatus()
}

; Function to speak current mute status using Text-to-Speech
SpeakMikeStatus() {
    global MikeMuted
    
    ; Determine the status message
    statusMessage := MikeMuted ? "Microphone off" : "Microphone on"

    ; Escape single quotes in the status message
    statusMessage := StrReplace(statusMessage, "'", "''")

    ; Speak status without showing PowerShell window
    Run, powershell -Command "(New-Object -ComObject SAPI.SPVoice).Speak('`"%statusMessage%`"');",, Hide
}

Return

; Function to gracefully exit the script
ExitScript() {
    Run, powershell -Command "(New-Object -ComObject SAPI.SPVoice).Speak('Exiting program');",, Hide
    ExitApp
}
