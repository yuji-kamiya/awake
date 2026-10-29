Option Explicit

Const APP_TITLE = "awake"
Const INTERVAL_MS = 60000
Const KEEP_ALIVE_KEY = "{F16}"
Const WORKER_ARGUMENT = "--worker"

Dim errorMessage

On Error Resume Next
Main
If Err.Number <> 0 Then
    errorMessage = "Operation failed." & vbCrLf & vbCrLf _
        & "File: " & WScript.ScriptFullName & vbCrLf _
        & "Source: " & Err.Source & vbCrLf _
        & "Code: " & Err.Number & vbCrLf & Err.Description
    On Error GoTo 0
    MsgBox errorMessage, vbOKOnly + vbCritical, APP_TITLE
    WScript.Quit 1
End If
On Error GoTo 0

Sub Main()
    Dim wmi, command, workers

    If WScript.Arguments.Count = 1 Then
        If WScript.Arguments(0) = WORKER_ARGUMENT Then
            KeepAwake
            Exit Sub
        End If
    End If
    If WScript.Arguments.Count <> 0 Then
        Err.Raise vbObjectError + 1, "Main", "Invalid arguments. Run without arguments."
    End If

    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    command = BuildWorkerCommand(WScript.ScriptFullName)
    Set workers = FindWorkers(wmi, command)

    If workers.Count > 0 Then
        If MsgBox("Awake is running. Stop it?", _
            vbYesNo + vbQuestion, APP_TITLE) = vbYes Then StopWorkers workers
        Exit Sub
    End If

    If MsgBox("Start awake?", vbYesNo + vbQuestion, APP_TITLE) <> vbYes Then Exit Sub

    ' Another instance may have started while the dialog was open.
    Set workers = FindWorkers(wmi, command)
    If workers.Count = 0 Then Call StartWorker(wmi, command)
End Sub

Sub KeepAwake()
    Dim shell
    Set shell = CreateObject("WScript.Shell")
    Do
        shell.SendKeys KEEP_ALIVE_KEY
        WScript.Sleep INTERVAL_MS
    Loop
End Sub

Function BuildWorkerCommand(scriptPath)
    Dim fso, hostPath
    Set fso = CreateObject("Scripting.FileSystemObject")
    hostPath = fso.BuildPath(fso.GetParentFolderName(WScript.FullName), "wscript.exe")
    BuildWorkerCommand = Chr(34) & hostPath & Chr(34) & " //nologo " _
        & Chr(34) & scriptPath & Chr(34) & " " & WORKER_ARGUMENT
End Function

Function FindWorkers(wmi, command)
    Dim workers, processes, process
    Set workers = CreateObject("Scripting.Dictionary")
    Set processes = wmi.ExecQuery("SELECT * FROM Win32_Process WHERE Name = 'wscript.exe'")

    For Each process In processes
        If Not IsNull(process.CommandLine) Then
            ' Match the full command to exclude the controller and other scripts.
            If StrComp(process.CommandLine, command, vbTextCompare) = 0 Then
                workers.Add CStr(process.ProcessId), process
            End If
        End If
    Next
    Set FindWorkers = workers
End Function

Function StartWorker(wmi, command)
    Dim processClass, result, processId
    Set processClass = wmi.Get("Win32_Process")
    result = processClass.Create(command, Null, Null, processId)
    If result <> 0 Then
        Err.Raise vbObjectError + 2, "StartWorker", _
            "Could not start worker. Status: " & result & vbCrLf & command
    End If
    StartWorker = processId
End Function

Sub StopWorkers(workers)
    Dim process, result
    For Each process In workers.Items
        result = process.Terminate(0)
        If result <> 0 Then
            Err.Raise vbObjectError + 3, "StopWorkers", _
                "Could not stop worker. PID: " & process.ProcessId & " / Status: " & result
        End If
    Next
End Sub
