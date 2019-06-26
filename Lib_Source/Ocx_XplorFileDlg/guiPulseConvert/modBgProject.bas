Attribute VB_Name = "modBgproject"
'---------------------------------------------------------------------------------------
' Module    : modBgproject
' Author    : Leo Herrera
' Date      : 01/01/2013
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

Function FindIndexBgProject(IndexTrVw As Long) As Long

Dim i           As Long

    With BackGroundProjectList
        For i = 0 To .Count - 1
            With BackGroundProjectList.ProjectList(i)
                If .IndexTrVw = IndexTrVw Then
                    FindIndexBgProject = i
                    Exit Function
                End If
            End With
        Next
    End With
    FindIndexBgProject = -1
        
End Function

' ListView
' - Project name
' Ubicaciòn
' Ubicaciòn Salida
' Estado Actual
' File 1:
Function GetEstadoBgProject(Index As Long)

    With BackGroundProjectList
        With BackGroundProjectList.ProjectList(Index)
            Select Case .ProjectState
                Case Is = Message_Header_Const.MSG_START_PROJECT
                    GetEstadoBgProject = "Comenzando del Proceso."
                Case Is = Message_Header_Const.MSG_FILE_START
                    GetEstadoBgProject = "Comenzando Generación de Archivo."
                Case Is = Message_Header_Const.MSG_STATUS
                    GetEstadoBgProject = .IndexPulse + 1 & " pulsos de la hoja " & .IndexSheet + 1
                Case Is = Message_Header_Const.MSG_SAVING_FILE
                    GetEstadoBgProject = "Guardando Archivo en Disco"
                Case Is = Message_Header_Const.MSG_XLS_FILE_READY
                    GetEstadoBgProject = "Archivo Guardado"
                Case Is = Message_Header_Const.MSG_END_PROJECT
                    GetEstadoBgProject = "Archivos Excel Generados."
            End Select
        End With
    End With
    
End Function

Sub ShowGralParameters(LstVw As ListView, Index As Long)

Dim i               As Integer
Dim LstItm          As ListItem

    With BackGroundProjectList
        With BackGroundProjectList.ProjectList(Index)
            LstVw.ListItems.Clear
            For i = 1 To 5
                LstVw.ListItems.Add
            Next
            LstVw.ListItems(1).Text = "Nombre Proyecto"
            LstVw.ListItems(1).ListSubItems.Add , , .ProjectName
            LstVw.ListItems(2).Text = "Ubicación"
            LstVw.ListItems(2).ListSubItems.Add , , .ProjectPath
            LstVw.ListItems(3).Text = "Ubicación Salida"
            LstVw.ListItems(3).ListSubItems.Add , , .OutputPath
            LstVw.ListItems(4).Text = "Cantidad de Archivos"
            LstVw.ListItems(4).ListSubItems.Add , , .OutFilesCount
            LstVw.ListItems(5).Text = "Estado"
            LstVw.ListItems(5).ListSubItems.Add , , GetEstadoBgProject(Index)
            For i = 1 To .OutFilesCount
                Set LstItm = LstVw.ListItems.Add(, , .OutFiles(i - 1))
                If .OutFiles(i - 1) = "" Then
                    LstItm.Text = "File " & i
                End If
                If i > .IndexSpread + 1 Then
                    LstItm.ListSubItems.Add , , "Pendiente"
                Else
                    If i <= .IndexSpread Then
                        LstItm.ListSubItems.Add , , "Guardado."
                    Else
                        LstItm.ListSubItems.Add , , .IndexSheet + 1 & _
                                " de " & .OutFilesSheetCount(i) + 1 & _
                                " hojas."
                    End If
                End If
            Next
            AutoAjusteColumnWidth LstVw
        End With
    End With
    
End Sub

Sub RefreshBgProject(LstVw As ListView, Index As Long)

Dim lvCount         As Long

    lvCount = LstVw.ListItems.Count - 5
    
    LstVw.ListItems(5).ListSubItems(1).Text = GetEstadoBgProject(Index)
    With BackGroundProjectList
        With BackGroundProjectList.ProjectList(Index)
            Select Case .ProjectState
                Case Is = Message_Header_Const.MSG_FILE_START
                    LstVw.ListItems(.IndexSpread + 6).Text = .OutFiles(.IndexSpread)
                    LstVw.ListItems(.IndexSpread + 6).ListSubItems(1).Text = "Creando Hojas."
                Case Is = Message_Header_Const.MSG_STATUS
                    If .IndexSpread < lvCount Then
                        'LstVw.ListItems(.IndexSpread + 6).Text = .OutFiles(.IndexSpread)
                        LstVw.ListItems(.IndexSpread + 6).ListSubItems(1).Text = _
                                .IndexSheet + 1 & " de " & _
                                .OutFilesSheetCount(.IndexSpread) & " hojas."
                    Else
                        LstVw.ListItems.Add , , "Wrong"
                    End If
                Case Is = Message_Header_Const.MSG_SAVING_FILE
                Case Is = Message_Header_Const.MSG_END_PROJECT
            End Select
        End With
    End With
    AutoAjusteColumnWidth LstVw
    
End Sub

Sub ShowBgProject(LstVw As ListView, Index As Long)

Dim i, n              As Integer
Dim LstItm          As ListItem
Dim IndexBg         As Long

    IndexBg = FindIndexBgProject(Index)
    ShowGralParameters LstVw, IndexBg
    
End Sub
