from __future__ import annotations

from pathlib import Path

from . import __version__ as _ADDIN_VERSION


def generate_vba_module(*, meta_sheet_name: str, module_name: str = "ClearMLDatasetExcelAddin") -> str:
    mod = module_name.replace('"', '""')
    meta = meta_sheet_name.replace('"', '""')
    ver = str(_ADDIN_VERSION).replace('"', '""')
    return f'''Attribute VB_Name = "{mod}"
Option Explicit

Private Const META_SHEET As String = "{meta}"
Private Const LOG_FILENAME As String = "clearml_dataset_excel_addin.log"
Private Const ADDIN_VERSION As String = "{ver}"

Public Sub ClearMLDatasetExcel_Run()
    Dim targetWb As Workbook
    Set targetWb = ResolveTargetWorkbook()
    If targetWb Is Nothing Then
        MsgBox "Could not find the condition workbook (missing _meta sheet).", vbCritical
        Exit Sub
    End If

    Dim enabled As String
    enabled = LCase$(MetaValue(targetWb, "addin_enabled"))
    If enabled <> "true" And enabled <> "1" And enabled <> "yes" Then
        MsgBox "Add-in is disabled (addin.enabled=false in YAML).", vbExclamation
        Exit Sub
    End If

    Dim expectedVersion As String
    expectedVersion = MetaValue(targetWb, "addin_version")
    If expectedVersion <> "" And expectedVersion <> ADDIN_VERSION Then
        MsgBox "Add-in version mismatch." & vbCrLf & _
               "Template expects: " & expectedVersion & vbCrLf & _
               "Macro/Add-in: " & ADDIN_VERSION & vbCrLf & _
               "Please regenerate template/add-in.", vbExclamation
    End If

    Dim wbPath As String
    wbPath = targetWb.FullName

    Dim wbDir As String
    wbDir = targetWb.Path
    If wbDir = "" Then
        MsgBox "Please save the workbook before running.", vbExclamation
        Exit Sub
    End If

    Dim specFile As String
    specFile = MetaValue(targetWb, "addin_spec_filename")
    If specFile = "" Then
        MsgBox "Missing meta: addin_spec_filename", vbCritical
        Exit Sub
    End If

    Dim specPath As String
#If Mac Then
    specPath = JoinPathPosix(wbDir, specFile)
#Else
    specPath = JoinPath(wbDir, specFile)
#End If

    Dim cmdTemplate As String
    cmdTemplate = SelectCommandTemplate(targetWb)
    If cmdTemplate = "" Then
        cmdTemplate = "clearml-dataset-excel run --spec ""${{SPEC}}"" --excel ""${{EXCEL}}"""
    End If

    Dim cmd As String
    ' Support both placeholder styles (single/double braces).
    cmd = Replace(cmdTemplate, "${{{{SPEC}}}}", specPath)
    cmd = Replace(cmd, "${{{{EXCEL}}}}", wbPath)
    cmd = Replace(cmd, "${{SPEC}}", specPath)
    cmd = Replace(cmd, "${{EXCEL}}", wbPath)

#If Mac Then
    RunOnMac cmd, wbDir
#Else
    RunOnWindows cmd, wbDir
#End If
End Sub

Public Sub ClearMLDatasetExcel_Run_Ribbon(ByVal control As Object)
    ClearMLDatasetExcel_Run
End Sub


Private Function SelectCommandTemplate(ByVal wb As Workbook) As String
    Dim targetOS As String
    targetOS = LCase$(MetaValue(wb, "addin_target_os"))

    Dim isMac As Boolean
    isMac = (InStr(1, Application.OperatingSystem, "Mac", vbTextCompare) > 0)

    Dim useMac As Boolean
    If targetOS = "" Or targetOS = "auto" Then
        useMac = isMac
    ElseIf targetOS = "mac" Then
        useMac = True
    ElseIf targetOS = "windows" Or targetOS = "win" Then
        useMac = False
    Else
        useMac = isMac
    End If

    Dim cmd As String
    If useMac Then
        cmd = MetaValue(wb, "addin_command_mac")
    Else
        cmd = MetaValue(wb, "addin_command_windows")
    End If

    If cmd = "" Then
        cmd = MetaValue(wb, "addin_command")
    End If

    SelectCommandTemplate = cmd
End Function


Private Sub RunOnWindows(ByVal cmd As String, ByVal wbDir As String)
    Dim sh As Object
    On Error GoTo EH
    Set sh = CreateObject("WScript.Shell")
    Dim logPath As String
    logPath = JoinPath(wbDir, LOG_FILENAME)

    Dim fullCmd As String
    fullCmd = "cmd.exe /c cd /d " & CmdQuote(wbDir) & " && (" & cmd & ") > " & CmdQuote(logPath) & " 2>&1"
    sh.Run fullCmd, 0, False
    MsgBox "Started. Log: " & logPath, vbInformation
    Exit Sub
EH:
    MsgBox "Failed to run command on Windows:" & vbCrLf & cmd & vbCrLf & vbCrLf & Err.Description, vbCritical
End Sub

Private Function CmdQuote(ByVal s As String) As String
    CmdQuote = Chr(34) & Replace(s, Chr(34), Chr(34) & Chr(34)) & Chr(34)
End Function


#If Mac Then
Private Sub RunOnMac(ByVal cmd As String, ByVal wbDir As String)
    On Error GoTo EH
    If wbDir = "" Then
        MsgBox "Please save the workbook before running.", vbExclamation
        Exit Sub
    End If

    ' Run under a login shell so PATH is set correctly (Homebrew/python.org installs).
    ' Write stdout/stderr to a log file next to the workbook for troubleshooting.
    Dim logPath As String
    logPath = JoinPathPosix(wbDir, LOG_FILENAME)

    Dim wrapped As String
    wrapped = "(" & cmd & ") > " & ShellQuote(logPath) & " 2>&1 &"

    Dim shCmd As String
    shCmd = "/bin/zsh -lc " & ShellQuote(wrapped)

    Call Shell(shCmd)
    MsgBox "Started. Log: " & logPath, vbInformation
    Exit Sub
EH:
    MsgBox "Failed to run command on Mac:" & vbCrLf & cmd & vbCrLf & vbCrLf & Err.Description, vbCritical
End Sub

Private Function ShellQuote(ByVal s As String) As String
    Dim t As String
    t = Replace(s, Chr(39), Chr(39) & Chr(34) & Chr(39) & Chr(34) & Chr(39))
    ShellQuote = Chr(39) & t & Chr(39)
End Function

Private Function JoinPathPosix(ByVal dirPath As String, ByVal fileName As String) As String
    If dirPath = "" Then
        JoinPathPosix = fileName
        Exit Function
    End If
    If Right$(dirPath, 1) = "/" Then
        JoinPathPosix = dirPath & fileName
    Else
        JoinPathPosix = dirPath & "/" & fileName
    End If
End Function
#End If


Private Function MetaValue(ByVal wb As Workbook, ByVal key As String) As String
    Dim ws As Worksheet
    Dim found As Range
    On Error GoTo EH
    Set ws = ResolveMetaWorksheet(wb)
    If ws Is Nothing Then
        MetaValue = ""
        Exit Function
    End If
    Set found = ws.Columns(1).Find(What:=key, LookIn:=xlValues, LookAt:=xlWhole)
    If found Is Nothing Then
        MetaValue = ""
        Exit Function
    End If
    MetaValue = CStr(found.Offset(0, 1).Value)
    Exit Function
EH:
    MetaValue = ""
End Function


Private Function ResolveMetaWorksheet(ByVal wb As Workbook) As Worksheet
    ' Prefer explicit META_SHEET name, but fall back to scanning for schema_version key.
    Dim ws As Worksheet
    Dim candidate As Worksheet
    Dim v As Variant

    On Error Resume Next
    Set ws = wb.Worksheets(META_SHEET)
    On Error GoTo 0
    If Not ws Is Nothing Then
        Set ResolveMetaWorksheet = ws
        Exit Function
    End If

    For Each candidate In wb.Worksheets
        On Error Resume Next
        v = candidate.Cells(1, 1).Value
        On Error GoTo 0
        If LCase$(CStr(v)) = "schema_version" Then
            Set ResolveMetaWorksheet = candidate
            Exit Function
        End If
    Next candidate

    Set ResolveMetaWorksheet = Nothing
End Function


Private Function ResolveTargetWorkbook() As Workbook
    ' Works for both:
    ' - workbook-embedded macros (ThisWorkbook is the condition workbook)
    ' - .xlam add-in macros (ThisWorkbook is the add-in; use ActiveWorkbook)
    On Error Resume Next
    If WorkbookHasMeta(ThisWorkbook) Then
        Set ResolveTargetWorkbook = ThisWorkbook
        Exit Function
    End If
    If Not ActiveWorkbook Is Nothing Then
        If WorkbookHasMeta(ActiveWorkbook) Then
            Set ResolveTargetWorkbook = ActiveWorkbook
            Exit Function
        End If
    End If
    Set ResolveTargetWorkbook = Nothing
End Function


Private Function WorkbookHasMeta(ByVal wb As Workbook) As Boolean
    If wb Is Nothing Then
        WorkbookHasMeta = False
        Exit Function
    End If

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(META_SHEET)
    On Error GoTo 0
    If Not ws Is Nothing Then
        WorkbookHasMeta = True
        Exit Function
    End If

    For Each ws In wb.Worksheets
        On Error Resume Next
        If LCase$(CStr(ws.Cells(1, 1).Value)) = "schema_version" Then
            WorkbookHasMeta = True
            Exit Function
        End If
        On Error GoTo 0
    Next ws

    WorkbookHasMeta = False
End Function


Private Function JoinPath(ByVal dirPath As String, ByVal fileName As String) As String
    If dirPath = "" Then
        JoinPath = fileName
        Exit Function
    End If
    Dim sep As String
    sep = Application.PathSeparator
    If Right$(dirPath, 1) = sep Then
        JoinPath = dirPath & fileName
    Else
        JoinPath = dirPath & sep & fileName
    End If
End Function


'''


def write_vba_module(path: str | Path, *, meta_sheet_name: str, module_name: str = "ClearMLDatasetExcelAddin") -> Path:
    out = Path(path).expanduser().resolve()
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(generate_vba_module(meta_sheet_name=meta_sheet_name, module_name=module_name), encoding="utf-8")
    return out
