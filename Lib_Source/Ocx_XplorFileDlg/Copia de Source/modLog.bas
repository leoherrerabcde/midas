Attribute VB_Name = "modLog"
Option Explicit

Public GV_hFile     As Integer
Public GV_Path_Iconos   As String
Public GV_File_Btn_Nsuperior    As String

Sub OpenLogFile()

Dim lv_LogFile      As String
Dim lvPath          As String

On Error GoTo Exit_OpenLogFile

    If GV_hFile = 0 Then
        GV_hFile = FreeFile
        lv_LogFile = "OpenDlg_" + Format(Now(), "YYYY_MM_SS_hh_mm_ss") + ".log"
        lvPath = Retroceder_Path(App.Path)
        lvPath = lvPath & "\Dbg"
        Open lvPath & "\" & lv_LogFile For Output As GV_hFile
        Print #GV_hFile, "App.Title : " & App.Title
        Print #GV_hFile, "App.Path : " & App.Path
        Print #GV_hFile, "Time Start: " & Format(Now(), "hh:mm:ss")
        Print #GV_hFile, vbCrLf
    End If
Exit_OpenLogFile:

End Sub

Sub WriteLogFile(lsStr As String)

    OpenLogFile
On Error Resume Next
    Print #GV_hFile, lsStr
    
End Sub


Sub CloseLogFile()

On Error GoTo Exit_CloseLogFile

    If GV_hFile Then
        Print #GV_hFile, vbCrLf
        Print #GV_hFile, "App.Title :" & App.Title
        Print #GV_hFile, "App.Path :" & App.Path
        Print #GV_hFile, "Time End: " & Format(Now(), "hh:mm:ss")
        Close GV_hFile
    End If
    
Exit_CloseLogFile:
        GV_hFile = 0

End Sub
