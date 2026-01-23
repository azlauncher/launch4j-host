Option Explicit

Dim fso, shell, RL_DIR, msg
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

RL_DIR = shell.ExpandEnvironmentStrings("%LOCALAPPDATA%\RuneLite")
msg = ""

' Check if all three versions exist
If fso.FileExists(RL_DIR & "\RuneLite.exe") And fso.FileExists(RL_DIR & "\RuneLiteV.exe") And fso.FileExists(RL_DIR & "\RuneLiteA.exe") Then
    ' Delete RuneLiteA.exe and rerun
    fso.DeleteFile RL_DIR & "\RuneLiteA.exe"
    shell.Run """" & WScript.ScriptFullName & """", 0, False
    WScript.Quit
End If

' If V version exists, make it active
If fso.FileExists(RL_DIR & "\RuneLiteV.exe") Then
    If fso.FileExists(RL_DIR & "\RuneLite.exe") Then
        fso.MoveFile RL_DIR & "\RuneLite.exe", RL_DIR & "\RuneLiteA.exe"
    End If
    fso.MoveFile RL_DIR & "\RuneLiteV.exe", RL_DIR & "\RuneLite.exe"
    msg = "The RuneLite launcher will now open vanilla RuneLite"

' If A version exists, make it active
ElseIf fso.FileExists(RL_DIR & "\RuneLiteA.exe") Then
    If fso.FileExists(RL_DIR & "\RuneLite.exe") Then
        fso.MoveFile RL_DIR & "\RuneLite.exe", RL_DIR & "\RuneLiteV.exe"
    End If
    fso.MoveFile RL_DIR & "\RuneLiteA.exe", RL_DIR & "\RuneLite.exe"
    msg = "The RuneLite launcher will now open azolite"

Else
    msg = "Error: no additional executable files found"
End If

' Show message box
If msg <> "" Then
    MsgBox msg, 64, "RuneLite launcher"
End If