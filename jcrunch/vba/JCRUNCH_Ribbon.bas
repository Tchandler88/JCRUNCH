Attribute VB_Name = "JCRUNCH_Ribbon"
Option Explicit

' ═══════════════════════════════════════════════
' CONSTANTS
' ═══════════════════════════════════════════════

Private Const CONFIG_SHEET  As String = "_Config"
Private Const PATH_CELL     As String = "A1"
Private Const STATUS_CELL   As String = "A2"
Private Const SENTINEL_FILE As String = "jcrunch_done.tmp"
Private Const POLL_INTERVAL As Long   = 2    ' seconds between polls
Private Const MAX_WAIT      As Long   = 300  ' 5 minute timeout

' ═══════════════════════════════════════════════
' BUTTON 1 — Browse for AEM package zip
' ═══════════════════════════════════════════════

Public Sub BrowsePackage()
    Dim fd As FileDialog
    Dim ws As Worksheet
    Dim selectedPath As String

    ' Get or create config sheet
    Set ws = GetConfigSheet()
    If ws Is Nothing Then
        MsgBox "Cannot access _Config sheet. Please contact support.", _
               vbCritical, "JCRUNCH"
        Exit Sub
    End If

    ' Open file dialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Select AEM Package (.zip)"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "AEM Package", "*.zip"
        .Filters.Add "All Files", "*.*"
        If .Show = -1 Then
            selectedPath = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With

    ' Write path to config sheet
    ws.Range(PATH_CELL).Value = selectedPath
    ws.Range(STATUS_CELL).Value = "Package selected: " & _
        Mid(selectedPath, InStrRev(selectedPath, "\") + 1)

    MsgBox "Package selected:" & vbCrLf & selectedPath & vbCrLf & _
           vbCrLf & "Click 'Run JCRUNCH' to process.", _
           vbInformation, "JCRUNCH — Package Ready"
End Sub

' ═══════════════════════════════════════════════
' BUTTON 2 — Run JCRUNCH pipeline
' ═══════════════════════════════════════════════

Public Sub RunJCRUNCH()
    Dim ws           As Worksheet
    Dim zipPath      As String
    Dim wbPath       As String
    Dim jcrunchDir   As String
    Dim pythonExe    As String
    Dim scriptPath   As String
    Dim sentinelPath As String
    Dim cmd          As String
    Dim startTime    As Double
    Dim elapsed      As Long

    ' Validate config sheet
    Set ws = GetConfigSheet()
    If ws Is Nothing Then
        MsgBox "Cannot access _Config sheet.", vbCritical, "JCRUNCH"
        Exit Sub
    End If

    ' Get zip path
    zipPath = Trim(ws.Range(PATH_CELL).Value)
    If Len(zipPath) = 0 Then
        MsgBox "No package selected." & vbCrLf & _
               "Click 'Browse Package' first.", _
               vbExclamation, "JCRUNCH"
        Exit Sub
    End If

    If Not FileExists(zipPath) Then
        MsgBox "Package file not found:" & vbCrLf & zipPath, _
               vbCritical, "JCRUNCH"
        Exit Sub
    End If

    ' Paths
    wbPath       = ThisWorkbook.FullName
    jcrunchDir   = ThisWorkbook.Path & "\jcrunch"
    scriptPath   = jcrunchDir & "\jcrunch.py"
    sentinelPath = ThisWorkbook.Path & "\" & SENTINEL_FILE

    ' Validate jcrunch.py exists
    If Not FileExists(scriptPath) Then
        MsgBox "jcrunch.py not found at:" & vbCrLf & scriptPath & vbCrLf & _
               vbCrLf & "Ensure the jcrunch folder is in the same " & _
               "directory as this workbook.", _
               vbCritical, "JCRUNCH"
        Exit Sub
    End If

    ' Find Python
    pythonExe = FindPython()
    If Len(pythonExe) = 0 Then
        MsgBox "Python not found on PATH." & vbCrLf & _
               "Install Python 3.8+ and ensure it is on your system PATH.", _
               vbCritical, "JCRUNCH"
        Exit Sub
    End If

    ' Clean up old sentinel
    If FileExists(sentinelPath) Then Kill sentinelPath

    ' Save workbook before run
    ThisWorkbook.Save

    ' Update status
    ws.Range(STATUS_CELL).Value = "Running JCRUNCH... please wait."
    Application.StatusBar = "JCRUNCH: Processing package..."
    DoEvents

    ' Build command
    ' cd into jcrunch dir so relative imports work, then run script
    ' Write sentinel on completion via && echo done > sentinel
    cmd = "cmd /c cd /d """ & jcrunchDir & """ && """ & pythonExe & """ """ & _
          scriptPath & """ --package """ & zipPath & _
          """ --workbook """ & wbPath & _
          """ && echo done > """ & sentinelPath & """"

    ' Run
    On Error GoTo RunError
    Shell cmd, vbNormalFocus
    On Error GoTo 0

    ' Poll for sentinel file (completion signal)
    startTime = Timer
    Do
        DoEvents
        Sleep 2000
        elapsed = CLng(Timer - startTime)

        ws.Range(STATUS_CELL).Value = "Running... " & elapsed & "s elapsed"
        Application.StatusBar = "JCRUNCH: Processing... " & elapsed & "s"

        If FileExists(sentinelPath) Then
            GoTo RunComplete
        End If

        If elapsed >= MAX_WAIT Then
            ws.Range(STATUS_CELL).Value = "Timeout after " & MAX_WAIT & "s"
            Application.StatusBar = False
            MsgBox "JCRUNCH did not complete within " & MAX_WAIT & " seconds." & _
                   vbCrLf & "Check the terminal window for errors.", _
                   vbExclamation, "JCRUNCH — Timeout"
            Exit Sub
        End If
    Loop

RunComplete:
    ' Clean up sentinel
    If FileExists(sentinelPath) Then Kill sentinelPath

    ' Reload workbook data
    Application.StatusBar = "JCRUNCH: Reloading workbook..."
    DoEvents
    ThisWorkbook.RefreshAll

    ws.Range(STATUS_CELL).Value = "JCRUNCH complete — " & Now()
    Application.StatusBar = False

    MsgBox "JCRUNCH complete!" & vbCrLf & _
           "All phase sheets have been populated." & vbCrLf & _
           vbCrLf & "Package: " & _
           Mid(zipPath, InStrRev(zipPath, "\") + 1), _
           vbInformation, "JCRUNCH — Done"
    Exit Sub

RunError:
    Application.StatusBar = False
    ws.Range(STATUS_CELL).Value = "Error: " & Err.Description
    MsgBox "Error running JCRUNCH:" & vbCrLf & Err.Description, _
           vbCritical, "JCRUNCH — Error"
End Sub

' ═══════════════════════════════════════════════
' HELPERS
' ═══════════════════════════════════════════════

Private Function GetConfigSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(CONFIG_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        ' Create it hidden
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets( _
                 ThisWorkbook.Sheets.Count))
        ws.Name = CONFIG_SHEET
        ws.Visible = xlSheetVeryHidden
        ws.Range(PATH_CELL).Value = ""
        ws.Range(STATUS_CELL).Value = "Ready"
    End If
    Set GetConfigSheet = ws
End Function

Private Function FileExists(path As String) As Boolean
    FileExists = (Len(Dir(path)) > 0)
End Function

Private Function FindPython() As String
    ' Try common python executables in order
    Dim candidates(2) As String
    candidates(0) = "python"
    candidates(1) = "python3"
    candidates(2) = "py"

    Dim i       As Integer
    Dim tmpFile As String

    tmpFile = Environ("TEMP") & "\jcrunch_pycheck.tmp"

    For i = 0 To 2
        On Error Resume Next
        Shell "cmd /c " & candidates(i) & " --version > """ & _
              tmpFile & """ 2>&1", vbHide
        Application.Wait Now + TimeValue("0:00:01")
        If FileExists(tmpFile) Then
            FindPython = candidates(i)
            Kill tmpFile
            On Error GoTo 0
            Exit Function
        End If
        On Error GoTo 0
    Next i

    FindPython = ""
End Function

Private Sub Sleep(milliseconds As Long)
    Application.Wait Now + milliseconds / 86400000
End Sub

' ═══════════════════════════════════════════════
' RIBBON CALLBACK STUBS
' (called by CustomUI when buttons are clicked)
' ═══════════════════════════════════════════════

Public Sub OnBrowsePackage(control As IRibbonControl)
    BrowsePackage
End Sub

Public Sub OnRunJCRUNCH(control As IRibbonControl)
    RunJCRUNCH
End Sub
