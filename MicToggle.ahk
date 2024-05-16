; Initialize state
global MikeMuted := false

; Speak a message to indicate the script is running
Run, powershell -Command "(New-Object -ComObject SAPI.SPVoice).Speak('Mike script is running.');",, Hide

; Hotkey definition for Left Windows + Shift + M
#+M::ToggleMikeMute()

; Function to toggle mic mute and display status
ToggleMikeMute() {
    ; Toggle muted state
    global MikeMuted
    MikeMuted := !MikeMuted
    
    ; Mute or unmute microphone using Nircmd (Ensure nircmd.exe is in the same directory as the script)
    if (MikeMuted)
        RunWait, nircmd.exe mutesysvolume 1 Microphone,, Hide
    else
        RunWait, nircmd.exe mutesysvolume 0 Microphone,, Hide
    
    ; Speak status using Text-to-Speech
    SpeakMikeStatus()
}

; Function to speak current mute status using Text-to-Speech
SpeakMikeStatus() {
    global MikeMuted
    
    ; Speak status without showing PowerShell window
    Run, powershell -Command "(New-Object -ComObject SAPI.SPVoice).Speak('Microphone is ' + ('muted' * !$MikeMuted + 'unmuted' * $MikeMuted));",, Hide
}

Return
