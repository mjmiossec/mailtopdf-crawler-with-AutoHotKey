SetTitleMatchMode, 2 ; Partial title matching

varname = 1

; Activate the Mail app
WinActivate, Mail

; Wait for the app to be active
WinWaitActive, Mail

Loop, 900 ; Repeat the following actions X times X being the number of emails you want to copy from first email pre-selected
{

    ; Press Ctrl + E to focus on the search box
    ; Send, ^e
    ; Sleep, 500 ; Adjust sleep time as needed

    ; Type DOWN arrow key to select the first email
    Send, {Down}
    Sleep, 3000

    ; Press Enter to open the selected email
    
    Send, {Enter}
    Sleep, 6000 ; Adjust sleep time as needed

    ; Press Ctrl + P to open the Print dialog
    Send, ^p
    Sleep, 1500 ; Wait for the Print dialog to appear

    Loop, 7
    {
    Send, {Tab}
    Sleep, 100 ; Optional: add a small delay between keypresses
    }

    ; Press Enter to go to save as
    Send, {Enter}

    ; Modify the following line to set the desired file path and name
    Sleep, 1000
    Send, {Enter}
    Sleep, 1000
    Send, mail_batch1_%varname%
    varname++

    ; Press Enter to confirm the save
    Send, {Enter}

    ; Wait for the Save As dialog to close
    WinWaitClose, Save As
    Sleep, 1000

    ; Press Ctrl + W to close the email
    ; Send, ^w
    Send {Esc} ; really close that email
    Sleep, 5000 ; Adjust sleep time as needed
}

MsgBox, Emails saved successfully for realsies!
