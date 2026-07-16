' launcher.vbs v2.0.0 — Collects ALL selected files from Explorer via COM, then launches doc-convert
' Bug fix 2026-05-19: previously iterated all Explorer windows and grabbed the
' SelectedItems of the FIRST one with any selection — which could be the wrong
' window if the user had another Explorer window open. The prior version
' workaround (lines 63-73) detected the mismatch but lost multi-select. The
' new logic matches the right-clicked path (%1) against each window's
' selection and uses ALL files from the correct window.
' Usage: wscript.exe launcher.vbs "%1"

Dim fso, wshShell
Set fso = CreateObject("Scripting.FileSystemObject")
Set wshShell = CreateObject("WScript.Shell")

Dim rightClickedPath
If WScript.Arguments.Count > 0 Then
    rightClickedPath = WScript.Arguments(0)
Else
    rightClickedPath = ""
End If

' Use a lock file so only the first instance does the work
Dim tempDir, lockFile
tempDir = wshShell.ExpandEnvironmentStrings("%TEMP%")
lockFile = tempDir & "\doc_convert.lock"

On Error Resume Next
If fso.FileExists(lockFile) Then
    Dim lockAge
    lockAge = DateDiff("s", fso.GetFile(lockFile).DateLastModified, Now)
    If lockAge < 5 Then
        WScript.Quit
    End If
End If

Dim lockHandle
Set lockHandle = fso.CreateTextFile(lockFile, True)
If Err.Number <> 0 Then
    WScript.Quit
End If
lockHandle.Close
On Error GoTo 0

' Small delay to let Explorer finish launching all instances
WScript.Sleep 400

' Find the Explorer window whose selection contains the right-clicked path
Dim shellApp, wnd, selectedItems
Dim files
Set files = CreateObject("Scripting.Dictionary")

On Error Resume Next
Set shellApp = CreateObject("Shell.Application")
For Each wnd In shellApp.Windows
    If InStr(1, TypeName(wnd.Document), "ShellFolderView", vbTextCompare) > 0 Then
        Set selectedItems = wnd.Document.SelectedItems
        If Not selectedItems Is Nothing Then
            If selectedItems.Count > 0 Then
                Dim item, matchFound
                matchFound = False
                For Each item In selectedItems
                    If StrComp(item.Path, rightClickedPath, vbTextCompare) = 0 Then
                        matchFound = True
                        Exit For
                    End If
                Next
                If matchFound Then
                    For Each item In selectedItems
                        If Not files.Exists(item.Path) Then
                            files.Add item.Path, True
                        End If
                    Next
                    Exit For
                End If
            End If
        End If
    End If
Next
On Error GoTo 0

' Fallback: use %1 alone if no window matched
If files.Count = 0 And rightClickedPath <> "" Then
    If fso.FileExists(rightClickedPath) Then
        files.Add rightClickedPath, True
    End If
End If

' Clean up lock
On Error Resume Next
fso.DeleteFile lockFile, True
On Error GoTo 0

If files.Count = 0 Then WScript.Quit

' Write file list to temp file
Dim collectFile, cf, key
collectFile = tempDir & "\doc_convert_batch.txt"
Set cf = fso.CreateTextFile(collectFile, True)
For Each key In files.Keys
    cf.WriteLine key
Next
cf.Close

' Launch PowerShell converter
Dim scriptPath, cmd
scriptPath = fso.GetParentFolderName(WScript.ScriptFullName) & "\doc-convert.ps1"
cmd = "powershell.exe -ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -File """ & scriptPath & """ -ListFile """ & collectFile & """"
wshShell.Run cmd, 0, False
