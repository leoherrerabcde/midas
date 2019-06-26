VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmConfigSpreadSheets 
   Caption         =   "Previsualización de Pulsos"
   ClientHeight    =   6960
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13980
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6960
   ScaleWidth      =   13980
   Begin VB.Frame FrameListView 
      Caption         =   "Pre Visualizacion"
      Height          =   4455
      Left            =   7800
      TabIndex        =   20
      Top             =   1680
      Width           =   6015
      Begin MSComctlLib.ListView LstVwPwd 
         Height          =   3135
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   5530
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame FrameControles 
      Caption         =   "Controles"
      Height          =   1575
      Left            =   7800
      TabIndex        =   18
      Top             =   120
      Width           =   6015
      Begin MSComctlLib.ProgressBar ProgressBarSpreadFile 
         Height          =   255
         Left            =   3360
         TabIndex        =   25
         Top             =   720
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComctlLib.ProgressBar ProgressBarCreaSheet 
         Height          =   255
         Left            =   3360
         TabIndex        =   22
         Top             =   360
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton cmdSaveXls 
         Caption         =   "Generar Xls"
         Height          =   975
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblSaveState 
         Caption         =   "Creando..."
         Height          =   255
         Left            =   1800
         TabIndex        =   26
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblProgressSpread 
         Caption         =   "Archivo Xls:"
         Height          =   255
         Left            =   1800
         TabIndex        =   24
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblProgressSheet 
         Caption         =   "Hoja:"
         Height          =   255
         Left            =   1800
         TabIndex        =   23
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame FrameConfigSpreadSheet 
      Caption         =   "Configuracion Hoja de Calculo"
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      Begin VB.Frame FrameSheetDetails 
         Caption         =   "Detalle por Hoja de Calculo"
         Height          =   4455
         Left            =   120
         TabIndex        =   7
         Top             =   2160
         Width           =   7215
         Begin VB.TextBox txtTimeEnd 
            Height          =   285
            Left            =   5160
            TabIndex        =   32
            Text            =   "txtTimeEnd"
            Top             =   480
            Width           =   1935
         End
         Begin VB.TextBox txtTimeStart 
            Height          =   285
            Left            =   3120
            TabIndex        =   31
            Text            =   "txtTimeStart"
            Top             =   480
            Width           =   1935
         End
         Begin VB.TextBox txtTotalPulses 
            Height          =   285
            Left            =   1680
            TabIndex        =   28
            Text            =   "txtTotalPulses"
            Top             =   480
            Width           =   1335
         End
         Begin MSComctlLib.TreeView TrVwSpreadFiles 
            Height          =   3255
            Left            =   120
            TabIndex        =   11
            Top             =   840
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   5741
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   441
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            FullRowSelect   =   -1  'True
            HotTracking     =   -1  'True
            SingleSel       =   -1  'True
            Appearance      =   1
         End
         Begin VB.TextBox txtSpreadCount 
            Height          =   285
            Left            =   120
            TabIndex        =   10
            Text            =   "txtSheetCount"
            Top             =   480
            Width           =   1215
         End
         Begin MSComctlLib.ListView LstVwConfigSpreadSheets 
            Height          =   3255
            Left            =   2640
            TabIndex        =   8
            Top             =   840
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   5741
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Campo"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Valor"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label lblEndTime 
            Caption         =   "Tiempo Termino:"
            Height          =   255
            Left            =   5160
            TabIndex        =   30
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label lblStartTime 
            Caption         =   "Tiempo Inicio:"
            Height          =   255
            Left            =   3120
            TabIndex        =   29
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblTotalPulses 
            Caption         =   "Pulsos Totales:"
            Height          =   255
            Left            =   1680
            TabIndex        =   27
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label lblSpreadQty 
            Caption         =   "Hojas Totales:"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame FrameSpreadSheetTypeConfig 
         Caption         =   "Tipo de Configuracion"
         Height          =   1935
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   7215
         Begin VB.CommandButton cmdCancel 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   240
            TabIndex        =   2
            Top             =   1440
            Width           =   1335
         End
         Begin VB.CommandButton cmdAccept 
            Caption         =   "Aceptar"
            Height          =   375
            Left            =   4680
            TabIndex        =   1
            Top             =   1440
            Width           =   1335
         End
         Begin VB.TextBox txtInterval 
            Height          =   375
            Left            =   4200
            TabIndex        =   17
            Text            =   "txtInterval"
            Top             =   960
            Width           =   1815
         End
         Begin VB.TextBox txtPulsesPerSheet 
            Height          =   375
            Left            =   1920
            TabIndex        =   15
            Text            =   "txtPulsesPerSheet"
            Top             =   960
            Width           =   1935
         End
         Begin VB.TextBox txtSheetCountPerFile 
            Height          =   375
            Left            =   240
            TabIndex        =   14
            Text            =   "txtSheetCountPerFile"
            Top             =   960
            Width           =   1335
         End
         Begin VB.OptionButton OptionSpredSheetTypeConfig 
            Caption         =   "por Intervalo de Tiempo"
            Height          =   375
            Index           =   2
            Left            =   4080
            TabIndex        =   6
            Top             =   360
            Width           =   2055
         End
         Begin VB.OptionButton OptionSpredSheetTypeConfig 
            Caption         =   "por Cantidad de Pulsos"
            Height          =   375
            Index           =   1
            Left            =   1920
            TabIndex        =   5
            Top             =   360
            Width           =   1935
         End
         Begin VB.OptionButton OptionSpredSheetTypeConfig 
            Caption         =   "Por Archivos"
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   4
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label lblInterval 
            Caption         =   "Intervalo por Hoja[seg]:"
            Height          =   255
            Left            =   4200
            TabIndex        =   16
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label lblHojas 
            Caption         =   "Hojas por Archivo:"
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label lblPulsos 
            Caption         =   "Pulsos:"
            Height          =   255
            Left            =   1920
            TabIndex        =   12
            Top             =   720
            Width           =   735
         End
      End
   End
End
Attribute VB_Name = "frmConfigSpreadSheets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private PV_LvtVw_Instance_Count     As Long
Private PV_Busy                     As Boolean

Sub InitForm()

    With Me
        .txtInterval = 0
        .txtPulsesPerSheet = 0
        .txtSpreadCount = ""
        .txtSheetCountPerFile = 1
        .OptionSpredSheetTypeConfig(0).Value = True
        .TrVwSpreadFiles.Nodes.Clear
        .LstVwConfigSpreadSheets.ListItems.Clear
        '.cmdAccept.SetFocus
    End With
    
End Sub

Sub InitWorkSpace()

Dim Interval        As Double
Dim Pulses          As Long
Dim Sheets          As Long

    Pulses = 0
    Pulse_Sheets_Per_Pulses Pulses
    Interval = 0
    Pulse_Sheets_Per_Interval Interval
    Pulse_Sheets_Per_File Sheets
    Pulse_CreateWorkSpace
    
    Me.ShowProjectInfo
    Me.LoadSpreadStruct Me.TrVwSpreadFiles
    
End Sub

Sub SetConfigByFiles()

Dim Interval        As Double
Dim Pulses          As Long
    
    With Me
        .txtInterval.Enabled = False
        .txtPulsesPerSheet.Enabled = False
        SetPulsesPerSheet .txtPulsesPerSheet
        Pulses = 0
        Pulse_Sheets_Per_Pulses Pulses
        Interval = 0
        Pulse_Sheets_Per_Interval Interval
    End With
    
End Sub
            
Sub SetConfigByPulses()

Dim Interval         As Double
Dim Pulses          As Long

    With Me
        .txtInterval.Enabled = False
        .txtPulsesPerSheet.Enabled = True
        'Pulses = 0
        'Pulse_Sheets_Per_Pulses Pulses
        Interval = 0
        Pulse_Sheets_Per_Interval Interval
        .SetPulsesPerSheet .txtPulsesPerSheet
    End With
    
End Sub
            
Sub SetSheetPerFile(txtSheet As TextBox)

Dim lSheets As Long

    If IsNumeric(txtSheet.Text) Then
        lSheets = Val(txtSheet.Text)
        If lSheets > 0 Then
            Pulse_Sheets_Per_File lSheets
        End If
    End If

End Sub

Sub SetPulsesPerSheet(txtPulses As TextBox)

Dim lPulses As Long

    If IsNumeric(txtPulses.Text) Then
        lPulses = Val(txtPulses.Text)
        If lPulses > 0 Then
            Pulse_Sheets_Per_Pulses lPulses
        End If
    End If

End Sub

Sub SetInterval(txtInterval As TextBox)

Dim Interval As Double

    If IsNumeric(txtInterval.Text) = True Then
        Interval = Val(txtInterval.Text)
        If Interval > 0 Then
            Pulse_Sheets_Per_Interval Interval
        End If
    End If
    
End Sub

Sub setConfigByInterval()

    With Me
        .txtInterval.Enabled = True
        .txtPulsesPerSheet.Enabled = False
        SetInterval .txtInterval
    End With
    
End Sub

Sub ClearAllTreeView()

Dim lvControl
Dim TrVw               As TreeView
Dim PreFixName          As String

    With Me
        PreFixName = "TrVw"
        For Each lvControl In .Controls
            If InStr(lvControl.Name, PreFixName) = 1 Then
                Set TrVw = lvControl
                TrVw.Visible = False
                TrVw.Nodes.Clear
                TrVw.Visible = True
            End If
        Next
    End With
    
End Sub
        
Sub ClearAllListView()

Dim lvControl
Dim LstVw               As ListView
Dim PreFixName          As String

    With Me
        PreFixName = "LstVw"
        For Each lvControl In .Controls
            If InStr(lvControl.Name, PreFixName) = 1 Then
                Set LstVw = lvControl
                LstVw.Enabled = False
                LstVw.ListItems.Clear
                LstVw.Enabled = True
            End If
        Next
    End With

End Sub

Sub ShowProjectInfo()

Dim SpreadCount         As Long
Dim PulsesQty           As Long
Dim TimeStart           As Double
Dim TimeEnd             As Double

    With Me
        SpreadCount = Pulse_GetSpreadFileCount
        Pulse_GetProjectInfo PulsesQty, TimeStart, TimeEnd
        .txtSpreadCount = SpreadCount
        .txtTotalPulses = PulsesQty
        .txtTimeStart = Format(TimeStart / 24000 / 3600, "hh:mm:ss")
        .txtTimeEnd = Format(TimeEnd / 2400 / 3600, "hh:mm:ss")
    End With
    
End Sub

Sub LoadSpreadInfo(TrVw As TreeView, LstVw As ListView)

Dim i               As Long
Dim IndexSpread     As Long
Dim IndexSheet      As Long
Dim IndexRoot       As Long

    With TrVw.SelectedItem
        If .Child Is Nothing Then
            IndexSheet = Val(.Tag)
            IndexRoot = .Parent.Index
            IndexSpread = Val(TrVw.Nodes(IndexRoot).Tag)
            ShowSheetInfo Me.LstVwConfigSpreadSheets, IndexSpread, IndexSheet
            ShowSheet Me.LstVwPWD, IndexSpread, IndexSheet
        Else
            IndexSpread = Val(.Tag)
            ShowSpreadInfo Me.LstVwConfigSpreadSheets, IndexSpread
        End If
    End With
    
End Sub

Sub ShowSheetInfo(LstVw As ListView, IndexSpread As Long, IndexSheet As Long)

Dim LstItm      As ListItem
Dim PlsQty      As Long
Dim TimeIni     As Double
Dim TimeEnd     As Double

    With LstVw
        .ListItems.Clear
        Pulse_GetSheetInfo IndexSpread, IndexSheet, PlsQty, TimeIni, TimeEnd
        Set LstItm = .ListItems.Add(, , "Pulse Qty")
        LstItm.ListSubItems.Add , , Trim$(Str(PlsQty))
        Set LstItm = .ListItems.Add(, , "Tiempo Inicio")
        LstItm.ListSubItems.Add , , Trim$(Str(TimeIni / 1000))
        Set LstItm = .ListItems.Add(, , "Tiempo Fin")
        LstItm.ListSubItems.Add , , Trim$(Str(TimeEnd / 1000))
    End With
    AutoAjusteColumnWidth LstVw
    
End Sub

Sub ShowSpreadInfo(LstVw As ListView, IndexSpread As Long)

Dim LstItm      As ListItem
Dim PlsQty      As Long
Dim TimeIni     As Double
Dim TimeEnd     As Double
Dim SheetCount  As Long

    With LstVw
        .ListItems.Clear
        Pulse_GetSpreadFileInfo IndexSpread, PlsQty, TimeIni, TimeEnd
        SheetCount = Pulse_GetSheetCount(IndexSpread)
        Set LstItm = .ListItems.Add(, , "Pulse Qty")
        LstItm.ListSubItems.Add , , Trim$(Str(PlsQty))
        Set LstItm = .ListItems.Add(, , "Tiempo Inicio")
        LstItm.ListSubItems.Add , , Trim$(Str(TimeIni / 1000))
        Set LstItm = .ListItems.Add(, , "Tiempo Fin")
        LstItm.ListSubItems.Add , , Trim$(Str(TimeEnd / 1000))
        Set LstItm = .ListItems.Add(, , "Tiempo Fin")
        LstItm.ListSubItems.Add , , Trim$(Str(TimeEnd / 1000))
        AutoAjusteColumnWidth LstVw
    End With
    
End Sub

Sub Load_Pwd_ColumnHeader(LstVw As ListView)

Dim lsStr       As String
Dim i           As Long

    lsStr = Space(261)
    With LstVw.ColumnHeaders
        .Clear
        .Add , , "Num"
        For i = 0 To Pulse_Field_Count - 1
            Pulse_Field_Header i, lsStr
             .Add , , lsStr
        Next
    End With
    AutoAjusteColumnWidth LstVw
    
End Sub

Sub ShowSheet(LstVw As ListView, ByVal IndexSpread As Long, ByVal IndexSheet As Long)

Dim lPulseCount     As Long
Dim i, j            As Long
Dim ldArray()       As Double
Dim InvalidateCount As Long
Dim Limit           As Long
Dim lCols           As Integer
Dim lvInstance      As Long

    PV_LvtVw_Instance_Count = PV_LvtVw_Instance_Count + 1
    lvInstance = PV_LvtVw_Instance_Count
'    If Me.cmdSaveXls.Enabled = False Then
'        Do
'            DoEvents
'            If lvInstance < PV_LvtVw_Instance_Count Then
'                Exit Sub
'            End If
'        Loop Until Me.cmdSaveXls.Enabled = True
'    End If
    InvalidateRect LstVw.hwnd, 0&, 0&
    GV_Mdi.mnuPjtExportSpreadSheet.Enabled = False   'CalcMnuExportState(False)
    Me.cmdSaveXls.Enabled = False
    With LstVw
        lCols = Pulse_Field_Count
        ReDim ldArray(lCols - 1)
        .ListItems.Clear
        InvalidateCount = 31
        lPulseCount = Pulse_GetSheetPulseCount(IndexSpread, IndexSheet)
        'Limit = lCountSheet / 2 - 1
        For i = 0 To lPulseCount - 1
            Pulse_GetPwd IndexSpread, IndexSheet, i, ldArray(0)
            AddDoubleItemListView LstVw, ldArray, i + 1, True
            ValidateRect LstVw.hwnd, 0&
            If ((i + 1) And InvalidateCount) = 0 Then
                InvalidateRect LstVw.hwnd, 0&, 0&
                DoEvents
                If lvInstance <> PV_LvtVw_Instance_Count Then
                    Exit For
                End If
            End If
        Next
    End With
    InvalidateRect LstVw.hwnd, 0&, 0&
    AutoAjusteColumnWidth LstVw
    If lvInstance = PV_LvtVw_Instance_Count Then
        PV_LvtVw_Instance_Count = 0
    End If
    Me.cmdSaveXls.Enabled = True
    GV_Mdi.mnuPjtExportSpreadSheet.Enabled = True
    
End Sub

Sub LoadSpreadStruct(TrVw As TreeView)

Dim lCountSpread    As Long
Dim lCountSheet     As Long
Dim i, j            As Long
Dim lvNod           As Node
Dim IndexPrev       As Long
Dim IndexParent     As Long
Dim lvStr           As String

    With TrVw
        .Nodes.Clear
        If Me.OptionSpredSheetTypeConfig.Item(0).Value = True Then
            lCountSpread = Pulse_GetSpreadFileCount
            IndexPrev = 0
            For i = 0 To lCountSpread - 1
                If i = 0 Then
                    Set lvNod = .Nodes.Add(, tvwFirst, , "Archivo Xls " & Trim$(Str(i + 1)))
                Else
                    Set lvNod = .Nodes.Add(IndexParent, tvwNext, , "Archivo Xls " & Trim$(Str(i + 1)))
                End If
                ValidateRect .hwnd, 0&
                lvNod.Tag = i
                IndexPrev = lvNod.Index
                IndexParent = IndexPrev
                lCountSheet = Pulse_GetSheetCount(i)
                For j = 0 To lCountSheet - 1
                    lvStr = Space(260)
                    Pulse_Get_File i, j, lvStr
                    If j = 0 Then
                        Set lvNod = .Nodes.Add(IndexParent, tvwChild, , lvStr)
                    Else
                        Set lvNod = .Nodes.Add(IndexPrev, tvwNext, , lvStr)
                    End If
                    lvNod.Tag = j
                    IndexPrev = lvNod.Index
                Next
            Next
        Else
            lCountSpread = Pulse_GetSpreadFileCount
            For i = 0 To lCountSpread - 1
                Set lvNod = .Nodes.Add(, , , "Archivo Xls " & Trim$(Str(i + 1)))
                lvNod.Tag = i
                ValidateRect .hwnd, 0&
            Next
            For i = 0 To lCountSpread - 1
                lCountSheet = Pulse_GetSheetCount(i)
                For j = 0 To lCountSheet - 1
                    Set lvNod = .Nodes.Add(i + 1, tvwChild, , "Hoja " & Trim$(Str(j + 1)))
                    lvNod.Tag = j
                    ValidateRect .hwnd, 0&
                Next
            Next
        End If
        InvalidateRect .hwnd, 0&, 0&
    End With

End Sub

Private Sub cmdAccept_Click()

    With Me
        ClearAllTreeView
        ClearAllListView
        Pulse_CreateWorkSpace
'        If .OptionSpredSheetTypeConfig(0).Value = True Then
'        Else
'            If .OptionSpredSheetTypeConfig(1).Value = True Then
'            Else
'                If .OptionSpredSheetTypeConfig(2).Value = True Then
'                End If
'            End If
'        End If
        .ShowProjectInfo
        .LoadSpreadStruct .TrVwSpreadFiles
    End With
    
End Sub

Private Sub cmdSaveXls_Click()

Dim lCountSpread    As Long
Dim lCountSheet     As Long
Dim i, j            As Long

    With Me
    
        GV_clsExportSpreadSheet.SetLoad
        
        Exit Sub
        
        .cmdSaveXls.Enabled = False
        lCountSpread = Pulse_GetSpreadFileCount
        .ProgressBarSpreadFile.Max = lCountSpread
        .ProgressBarSpreadFile.Min = 0
        .ProgressBarCreaSheet.Min = 0
        For i = 0 To lCountSpread - 1
            .ProgressBarSpreadFile.Value = i
            DoEvents
        Next
    End With
    
End Sub

Private Sub Form_Load()
    
    GV_Mdi.Set_Status_MnuProject ConfigOutput, False
    InitForm
    Load_Pwd_ColumnHeader Me.LstVwPWD
    InitWorkSpace
    
End Sub

Private Sub Form_Resize()

    With Me
        .FrameListView.Width = .ScaleWidth - .FrameListView.Left - 60
        .FrameListView.Height = .ScaleHeight - .FrameListView.Top - 60
        .LstVwPWD.Width = .FrameListView.Width - 2 * .LstVwPWD.Left
        .LstVwPWD.Height = .FrameListView.Height - 2 * .LstVwPWD.Top
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    GV_clsConfigSpreadSheet.ClearLoaded
    'GV_Mdi.Restore_Visible_MnuProject PreviewOutput
    GV_Mdi.ProjectMnuUpdate
    
End Sub

Private Sub OptionSpredSheetTypeConfig_Click(Index As Integer)

    Select Case Index
        Case Is = 0
            SetConfigByFiles
        Case Is = 1
            SetConfigByPulses
        Case Is = 2
            setConfigByInterval
    End Select
    
End Sub


Private Sub TrVwSpreadFiles_NodeCheck(ByVal Node As MSComctlLib.Node)

    Me.LoadSpreadInfo Me.TrVwSpreadFiles, Me.LstVwConfigSpreadSheets

End Sub

Private Sub TrVwSpreadFiles_NodeClick(ByVal Node As MSComctlLib.Node)

    Me.LoadSpreadInfo Me.TrVwSpreadFiles, Me.LstVwConfigSpreadSheets

End Sub

Private Sub txtInterval_Change()

Dim Interval             As Double

    With Me
        SetInterval .txtInterval
    End With

End Sub

Private Sub txtPulsesPerSheet_Change()

Dim lPulses             As Long

    With Me
        .SetPulsesPerSheet .txtPulsesPerSheet
    End With
    
End Sub

Private Sub txtSheetCountPerFile_Change()

Dim lFiles              As Long

    SetSheetPerFile Me.txtSheetCountPerFile
    
End Sub
