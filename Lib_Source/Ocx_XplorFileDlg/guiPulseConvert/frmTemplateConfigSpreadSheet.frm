VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTemplateConfigSpreadSheet 
   Caption         =   "Form1"
   ClientHeight    =   7245
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14640
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7245
   ScaleWidth      =   14640
   Begin VB.Timer tmrMoveMouse 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   0
      Top             =   120
   End
   Begin VB.Frame FrameListView 
      Caption         =   "Pre Visualizacion"
      Height          =   4095
      Left            =   6480
      TabIndex        =   42
      Top             =   3120
      Width           =   8055
      Begin VB.CommandButton cmdUp 
         Caption         =   "Aba&jo"
         Height          =   255
         Index           =   1
         Left            =   5280
         TabIndex        =   45
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdUp 
         Caption         =   "Arri&ba"
         Height          =   255
         Index           =   0
         Left            =   2160
         TabIndex        =   44
         Top             =   240
         Width           =   615
      End
      Begin MSComctlLib.ListView LstVwPwd 
         Height          =   3375
         Left            =   120
         TabIndex        =   43
         Top             =   600
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   5953
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
      Begin MSComctlLib.Toolbar toolbarControles 
         Height          =   390
         Left            =   120
         TabIndex        =   47
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         ImageList       =   "ImageListControles"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   3
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   1
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   2
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   4
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameSheetDetails 
      Caption         =   "Detalle por Hoja de Calculo"
      Height          =   4455
      Left            =   120
      TabIndex        =   31
      Top             =   2280
      Width           =   6255
      Begin VB.TextBox txtSpreadCount 
         Height          =   285
         Left            =   120
         TabIndex        =   36
         Text            =   "txtSheetCount"
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtTotalPulses 
         Height          =   285
         Left            =   1560
         TabIndex        =   34
         Text            =   "txtTotalPulses"
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtTimeStart 
         Height          =   285
         Left            =   2880
         TabIndex        =   33
         Text            =   "txtTimeStart"
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtTimeEnd 
         Height          =   285
         Left            =   4320
         TabIndex        =   32
         Text            =   "txtTimeEnd"
         Top             =   480
         Width           =   1215
      End
      Begin MSComctlLib.TreeView TrVwSpreadFiles 
         Height          =   3255
         Left            =   120
         TabIndex        =   35
         Top             =   840
         Width           =   2895
         _ExtentX        =   5106
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
      Begin MSComctlLib.ListView LstVwConfigSpreadSheets 
         Height          =   3255
         Left            =   3120
         TabIndex        =   37
         Top             =   840
         Width           =   3015
         _ExtentX        =   5318
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
      Begin VB.Label lblSpreadQty 
         Caption         =   "Archivos Totales:"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblTotalPulses 
         Caption         =   "Pulsos Totales:"
         Height          =   255
         Left            =   1560
         TabIndex        =   40
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblStartTime 
         Caption         =   "Tiempo Inicio:"
         Height          =   255
         Left            =   2880
         TabIndex        =   39
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblEndTime 
         Caption         =   "Tiempo Termino:"
         Height          =   255
         Left            =   4320
         TabIndex        =   38
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdAccept 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   5400
      TabIndex        =   9
      Top             =   6840
      Width           =   975
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   6840
      Width           =   975
   End
   Begin VB.Frame frameConfigColums 
      Caption         =   "Configuración Columnas de Salida"
      Height          =   3135
      Left            =   6480
      TabIndex        =   7
      Top             =   0
      Width           =   8055
      Begin VB.CommandButton cmdCancelNewColumnConfig 
         Caption         =   "&Cancelar"
         Height          =   255
         Left            =   3960
         TabIndex        =   18
         Top             =   720
         Width           =   855
      End
      Begin VB.CommandButton cmdAddNewColumnConfig 
         Caption         =   "&Agregar"
         Height          =   255
         Left            =   5760
         TabIndex        =   17
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtNewColumConfig 
         Height          =   315
         Left            =   4080
         TabIndex        =   14
         Text            =   "txtNewColumConfig"
         Top             =   480
         Width           =   2655
      End
      Begin VB.ComboBox cboColumnConfig 
         Height          =   315
         Left            =   3960
         TabIndex        =   13
         Text            =   "cboColumnConfig"
         Top             =   360
         Width           =   2655
      End
      Begin MSComctlLib.ListView LstVwHideColumns 
         Height          =   615
         Left            =   120
         TabIndex        =   12
         Top             =   2160
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   1085
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.Toolbar toolbarColumnConfig 
         Height          =   390
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         ImageList       =   "ImageListConfigColumn"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "hide"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   3
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   1
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   4
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   1
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  EndProperty
               EndProperty
            EndProperty
         EndProperty
         Enabled         =   0   'False
      End
      Begin MSComctlLib.ListView LstVwColumnConfig 
         Height          =   615
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   1085
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ImageList ImageListConfigColumn 
         Left            =   120
         Top             =   240
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTemplateConfigSpreadSheet.frx":0000
               Key             =   "add"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTemplateConfigSpreadSheet.frx":015A
               Key             =   "hide"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTemplateConfigSpreadSheet.frx":02B4
               Key             =   "move_to_left"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTemplateConfigSpreadSheet.frx":084E
               Key             =   "move_to_right"
            EndProperty
         EndProperty
      End
      Begin VB.Label lbl 
         Caption         =   "Selección de Plantilla:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   1680
         TabIndex        =   20
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label lblWarningColumnConfig 
         Caption         =   "Nombre Existente."
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3960
         TabIndex        =   19
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label lbl 
         Caption         =   "Campos Ocultos:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label lblOutput 
         Caption         =   "Campos de Salida:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   2175
      End
   End
   Begin VB.Frame FrameSpreadSheetTypeConfig 
      Caption         =   "Tipo de Configuracion"
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6255
      Begin VB.CommandButton cmdApply 
         Caption         =   "A&plicar"
         Height          =   195
         Left            =   360
         TabIndex        =   46
         Top             =   2040
         Width           =   735
      End
      Begin VB.PictureBox pictureSpreadConfigControls 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   1095
         Left            =   240
         ScaleHeight     =   1095
         ScaleWidth      =   5895
         TabIndex        =   21
         Top             =   360
         Width           =   5895
         Begin VB.TextBox txtInterval 
            Height          =   375
            Left            =   3960
            TabIndex        =   27
            Text            =   "txtInterval"
            Top             =   600
            Width           =   1815
         End
         Begin VB.TextBox txtPulsesPerSheet 
            Height          =   375
            Left            =   1680
            TabIndex        =   26
            Text            =   "txtPulsesPerSheet"
            Top             =   600
            Width           =   1935
         End
         Begin VB.TextBox txtSheetCountPerFile 
            Height          =   375
            Left            =   0
            TabIndex        =   25
            Text            =   "txtSheetCountPerFile"
            Top             =   600
            Width           =   1335
         End
         Begin VB.OptionButton OptionSpredSheetTypeConfig 
            Caption         =   "por Intervalo de Tiempo"
            Height          =   375
            Index           =   2
            Left            =   3840
            TabIndex        =   24
            Top             =   0
            Width           =   2055
         End
         Begin VB.OptionButton OptionSpredSheetTypeConfig 
            Caption         =   "por Cantidad de Pulsos"
            Height          =   375
            Index           =   1
            Left            =   1680
            TabIndex        =   23
            Top             =   0
            Width           =   1935
         End
         Begin VB.OptionButton OptionSpredSheetTypeConfig 
            Caption         =   "Por Archivos"
            Height          =   375
            Index           =   0
            Left            =   0
            TabIndex        =   22
            Top             =   0
            Width           =   1215
         End
         Begin VB.Label lblInterval 
            Caption         =   "Intervalo por Hoja[seg]:"
            Height          =   255
            Left            =   3960
            TabIndex        =   30
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label lblHojas 
            Caption         =   "Hojas por Archivo:"
            Height          =   255
            Left            =   0
            TabIndex        =   29
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label lblPulsos 
            Caption         =   "Pulsos:"
            Height          =   255
            Left            =   1680
            TabIndex        =   28
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.CommandButton cmdCancelNewConfigType 
         Caption         =   "&Cancelar"
         Height          =   255
         Left            =   3000
         TabIndex        =   5
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox txtNewConfigtype 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3240
         TabIndex        =   4
         Text            =   "txtNewConfigtype"
         Top             =   1680
         Width           =   2895
      End
      Begin VB.CommandButton cmdAddNewConfigType 
         Caption         =   "&Agregar"
         Height          =   255
         Left            =   5160
         TabIndex        =   3
         Top             =   1920
         Width           =   855
      End
      Begin VB.ComboBox cboConfigType 
         Height          =   315
         Left            =   3120
         TabIndex        =   1
         Text            =   "cboConfigType"
         Top             =   1560
         Width           =   2895
      End
      Begin MSComctlLib.ImageList ImageListControles 
         Left            =   2400
         Top             =   1680
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTemplateConfigSpreadSheet.frx":0DE8
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTemplateConfigSpreadSheet.frx":17FA
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTemplateConfigSpreadSheet.frx":220C
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTemplateConfigSpreadSheet.frx":27A6
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label lblWarning 
         Caption         =   "Nombre Existente."
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Plantilla Tipo de Configuracion:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   1560
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmTemplateConfigSpreadSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmTemplateConfigSpreadSheet
' Author    : Leo Herrera
' Date      : 29/11/2012
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

Private PV_ConfigSheetList          As typeConfigSpreadSheetList
Private PV_ColumnConfigList         As typeColumnsConfigList
Private PV_ActualColumnConfig       As typeConfigSheetColumns

Private PV_ColumnOut_Index          As Integer
Private PV_ColumnHide_Index         As Integer
Private PV_LstVwOutput              As Boolean
Private PV_ColHdrOut                As MSComctlLib.ColumnHeader
Private PV_ColHdrHide               As MSComctlLib.ColumnHeader

Private PV_LvtVw_Instance_Count     As Long
Private PV_Init_Form                As Boolean
Private PV_Index_WrkSpc             As Integer

Private WithEvents mClickOnSpreadConfig        As clsDetectClickOnControl
Attribute mClickOnSpreadConfig.VB_VarHelpID = -1
Private WithEvents mClickOnColumnConfig        As clsDetectClickOnControl
Attribute mClickOnColumnConfig.VB_VarHelpID = -1

Private Const DEFAULT_OPTION = "<default>"

Dim WithEvents m_clsNewItem             As clsCboNew
Attribute m_clsNewItem.VB_VarHelpID = -1
Dim WithEvents m_clsColumnConfig        As clsCboNew
Attribute m_clsColumnConfig.VB_VarHelpID = -1

Dim PV_PreView                          As Boolean
Dim PV_Index_PreVw                      As Long
Dim PV_Count_PreVw                      As Long

Private PV_Height_Min                   As Long
Private PV_Cbo_Width                    As Long
Private PV_NewSpreadConfig              As Boolean

Private PV_Init_Busy                    As Boolean

Sub InitForm()

    With Me
        PV_Height_Min = 7755
        PV_Cbo_Width = .cboColumnConfig.Width
        PV_Index_WrkSpc = -1
        .txtInterval = 0
        .txtPulsesPerSheet = 0
        .txtSpreadCount = ""
        .txtSheetCountPerFile = 1
        .OptionSpredSheetTypeConfig(0).Value = True
        .TrVwSpreadFiles.Nodes.Clear
        .LstVwConfigSpreadSheets.ListItems.Clear
        '.cmdAccept.SetFocus
        Set mClickOnSpreadConfig = New clsDetectClickOnControl
        Set mClickOnColumnConfig = New clsDetectClickOnControl
        mClickOnSpreadConfig.SetControls .frameConfigColums, _
                                            .tmrMoveMouse
        mClickOnColumnConfig.SetControls .FrameSpreadSheetTypeConfig, _
                                            .tmrMoveMouse
        PV_PreView = False
        PV_Index_PreVw = 0
        PV_Count_PreVw = 512
    End With
    
End Sub

Sub InitWorkSpace()

Dim Interval        As Double
Dim Pulses          As Long
Dim Sheets          As Long

'    Pulses = 0
'    Pulse_Sheets_Per_Pulses Pulses
'    Interval = 0
'    Pulse_Sheets_Per_Interval Interval
'    Pulse_Sheets_Per_File Sheets
'    Pulse_CreateWorkSpace
    
    Me.ShowProjectInfo
    Me.LoadSpreadStruct Me.TrVwSpreadFiles
    
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

Sub GenerateIntermediaData()

    With Me
        ClearAllTreeView
        ClearAllListView
        m_Project.Set_IntermediaFileCount
        If m_Project.GetIntermediateDataReady = False Then
            Pulse_CreateWorkSpace
            m_Project.SetIntermediateDataReady
        Else
            m_Project.LoadWorkSpace
        End If
        .ShowProjectInfo
        .LoadSpreadStruct .TrVwSpreadFiles
    End With

End Sub

Sub ShowProjectInfo()

Dim SpreadCount         As Long
Dim PulsesQty           As Long
Dim TimeStart           As Double
Dim TimeEnd             As Double

    With Me
        If m_Project.GetIntermediateDataReady = True Then
            SpreadCount = Pulse_GetSpreadFileCount
            Pulse_GetProjectInfo PulsesQty, TimeStart, TimeEnd
            .txtSpreadCount = SpreadCount
            .txtTotalPulses = PulsesQty
            .txtTimeStart = Format(TimeStart / 24000 / 3600, "hh:mm:ss")
            .txtTimeEnd = Format(TimeEnd / 2400 / 3600, "hh:mm:ss")
        Else
            .txtSpreadCount = ""
            .txtTotalPulses = ""
            .txtTimeStart = ""
            .txtTimeEnd = ""
        End If
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
            'ShowSheet Me.LstVwPwd, IndexSpread, IndexSheet
            Me.toolbarControles.Enabled = modViewData.Init_Index(Me.LstVwPWD, _
                                            IndexSpread, IndexSheet, 0)
            modViewData.RefreshLstVwData Me.LstVwPWD
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
        'Set LstItm = .ListItems.Add(, , "Tiempo Fin")
        'LstItm.ListSubItems.Add , , Trim$(Str(TimeEnd / 1000))
        AutoAjusteColumnWidth LstVw
    End With
    
End Sub

Sub Load_Pwd_ColumnHeader(LstVw As ListView)

Dim lsStr       As String
Dim i           As Long

    lsStr = Space(261)
    With LstVw.ColumnHeaders
        .Clear
        '.Add , , "Num"
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
Dim lvIniPreVw      As Long
Dim lvEndPreVw      As Long

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
    GV_Mdi.mnuPjtExportSpreadSheet.Enabled = False  ' CalcMnuExportState(False)
    'Me.cmdSaveXls.Enabled = False
    With LstVw
        lvIniPreVw = PV_Index_PreVw
        lvEndPreVw = lvIniPreVw + PV_Count_PreVw
        lCols = Pulse_Field_Count
        ReDim ldArray(lCols - 1)
        .ListItems.Clear
        InvalidateCount = 31
        lPulseCount = Pulse_GetSheetPulseCount(IndexSpread, IndexSheet)
        'Limit = lCountSheet / 2 - 1
        If lvEndPreVw > lPulseCount Then
            lvEndPreVw = lPulseCount
        End If
        'For i = 0 To lPulseCount - 1
        For i = 0 To lvEndPreVw - 1
            Pulse_GetPwd IndexSpread, IndexSheet, i, ldArray(0)
            'AddDoubleItemListView LstVw, ldArray, i + 1, True
            AddDoubleItemListViewWithFilter LstVw, _
                                            ldArray, _
                                            PV_ActualColumnConfig, _
                                            i + 1, _
                                            True
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
    'Me.cmdSaveXls.Enabled = True
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
                lvStr = "Archivo Xls " & Trim$(Str(i + 1))
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
                    'lvStr = Space(260)
                    'Pulse_Get_File i, j, lvStr
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

Private Sub cmdUp_Click(Index As Integer)

Dim i           As Long

    i = PV_Index_PreVw
    If Index = 0 Then
        'UP
        i = i - PV_Count_PreVw
    Else
        ' DOWN
        i = i + PV_Count_PreVw
    End If
    
End Sub

Private Sub m_clsColumnConfig_OnClickNewItem(Index As Integer)

    Me.cboColumnConfig.ListIndex = Index
    
End Sub

Private Sub m_clsNewItem_OnClickNewItem(Index As Integer)

    'Me.GetSpreadConfigFromCbo Me.cboConfigType
    Me.cboConfigType.ListIndex = Index
    Me.GetSpreadConfigFromCbo Me.cboConfigType
    
End Sub

Private Sub toolbarControles_ButtonClick(ByVal Button As MSComctlLib.Button)

    With Me
        modViewData.ProcessScrollData .LstVwPWD, Button.Index
    End With

End Sub

Private Sub TrVwSpreadFiles_NodeCheck(ByVal Node As MSComctlLib.Node)

    Me.LoadSpreadInfo Me.TrVwSpreadFiles, Me.LstVwConfigSpreadSheets

End Sub

Private Sub TrVwSpreadFiles_NodeClick(ByVal Node As MSComctlLib.Node)

    Me.LoadSpreadInfo Me.TrVwSpreadFiles, Me.LstVwConfigSpreadSheets

End Sub
'
'
Sub MoveToLeft()

Dim lvTxt           As String
Dim lvIndex         As Integer

    If PV_LstVwOutput = False Then
        Exit Sub
    End If
    If PV_ColumnOut_Index > 1 Then
        With Me.LstVwColumnConfig.ColumnHeaders
            lvTxt = .Item(PV_ColumnOut_Index).Text
            lvIndex = .Item(PV_ColumnOut_Index).Tag
            .Item(PV_ColumnOut_Index).Text = .Item(PV_ColumnOut_Index - 1).Text
            .Item(PV_ColumnOut_Index).Tag = .Item(PV_ColumnOut_Index - 1).Tag
            .Item(PV_ColumnOut_Index - 1).Text = lvTxt
            .Item(PV_ColumnOut_Index - 1).Tag = lvIndex
        End With
        PV_ColumnOut_Index = PV_ColumnOut_Index - 1
        AutoAjusteColumnWidth Me.LstVwColumnConfig
    End If
    
End Sub

Sub MoveToRight()

Dim lvTxt           As String
Dim lvIndex         As Integer

    If PV_LstVwOutput = False Then
        Exit Sub
    End If
    If PV_ColumnOut_Index < Me.LstVwColumnConfig.ColumnHeaders.Count - 1 Then
        With Me.LstVwColumnConfig.ColumnHeaders
            lvTxt = .Item(PV_ColumnOut_Index).Text
            lvIndex = .Item(PV_ColumnOut_Index).Tag
            .Item(PV_ColumnOut_Index).Text = .Item(PV_ColumnOut_Index + 1).Text
            .Item(PV_ColumnOut_Index).Tag = .Item(PV_ColumnOut_Index + 1).Tag
            .Item(PV_ColumnOut_Index + 1).Text = lvTxt
            .Item(PV_ColumnOut_Index + 1).Tag = lvIndex
        End With
        PV_ColumnOut_Index = PV_ColumnOut_Index + 1
        AutoAjusteColumnWidth Me.LstVwColumnConfig
    End If

End Sub

Sub MoveToHide()

    With Me
        If PV_ColumnOut_Index < 0 Or _
            PV_ColumnOut_Index > .LstVwColumnConfig.ColumnHeaders.Count _
            Then
            Exit Sub
        End If
        .LstVwHideColumns.ColumnHeaders.Add , , .LstVwColumnConfig.ColumnHeaders(PV_ColumnOut_Index).Text
        PV_ColumnHide_Index = .LstVwHideColumns.ColumnHeaders.Count
        .LstVwHideColumns.ColumnHeaders(PV_ColumnHide_Index).Tag = .LstVwColumnConfig.ColumnHeaders(PV_ColumnOut_Index).Tag
        .LstVwHideColumns.ColumnHeaders(PV_ColumnHide_Index).Width = .LstVwColumnConfig.ColumnHeaders(PV_ColumnOut_Index).Width
        .LstVwColumnConfig.ColumnHeaders.Remove PV_ColumnOut_Index
    End With
    
End Sub

Sub MoveToOutput()

    With Me
        If PV_ColumnOut_Index < 0 Or _
            PV_ColumnOut_Index > .LstVwColumnConfig.ColumnHeaders.Count _
            Then
            Exit Sub
        End If
        If PV_ColumnHide_Index < 0 Or _
            PV_ColumnHide_Index > .LstVwHideColumns.ColumnHeaders.Count _
            Then
            Exit Sub
        End If
        .LstVwColumnConfig.ColumnHeaders.Add PV_ColumnOut_Index, , _
                            .LstVwColumnConfig.ColumnHeaders(PV_ColumnHide_Index).Text
        .LstVwColumnConfig.ColumnHeaders(PV_ColumnOut_Index).Tag = .LstVwColumnConfig.ColumnHeaders(PV_ColumnHide_Index).Tag
        .LstVwColumnConfig.ColumnHeaders(PV_ColumnOut_Index).Width = .LstVwColumnConfig.ColumnHeaders(PV_ColumnHide_Index).Width
        .LstVwHideColumns.ColumnHeaders.Remove PV_ColumnHide_Index
    End With

End Sub

Sub Add_NewColumnConfigToList()

    ReDim Preserve PV_ColumnConfigList.Config(PV_ColumnConfigList.Count)
    Me.AddColumnConfigToTemplate PV_ColumnConfigList.Count
    'Debug.Print cboColumnConfig.ListIndex
    'Me.cboColumnConfig.ListIndex = PV_ColumnConfigList.Count
    PV_ColumnConfigList.Count = PV_ColumnConfigList.Count + 1
    
End Sub

Sub Add_NewSpreadConfigToList()

    ReDim Preserve PV_ConfigSheetList.ConfigList(PV_ConfigSheetList.Count)
    Me.AddSpreadConfigToTemplate PV_ConfigSheetList.Count
    PV_ConfigSheetList.Count = PV_ConfigSheetList.Count + 1
    
End Sub

Sub GetColumnConfigFromControls(LstVwOutput As ListView, _
                                    LstVwHide As ListView)

Dim i           As Integer
Dim Index       As Integer

    With PV_ActualColumnConfig
        .ColumnConfigName = Me.cboColumnConfig.List(Me.cboColumnConfig.ListCount - 1)
        
        .Count = LstVwOutput.ColumnHeaders.Count + _
                    LstVwHide.ColumnHeaders.Count
        For i = 1 To LstVwOutput.ColumnHeaders.Count
            Index = LstVwOutput.ColumnHeaders(i).Tag
            .Column(Index).ColumnName = LstVwOutput.ColumnHeaders(i).Text
            .Column(Index).Order = i - 1
            .Column(Index).Visible = True
        Next
        For i = 1 To LstVwHide.ColumnHeaders.Count
            Index = LstVwHide.ColumnHeaders(i).Tag
            .Column(Index).ColumnName = LstVwHide.ColumnHeaders(i).Text
            .Column(Index).Order = i - 1
            .Column(Index).Visible = False
        Next
    End With
    SetActualColumnConfig PV_ActualColumnConfig
    
End Sub

Sub GetColumnConfigFromActualConfig(LstVwOutput As ListView, _
                                    LstVwHide As ListView, _
                                    LstVwPreview As ListView)
    
Dim i                   As Integer
Dim Lv_ListIndex()      As Integer
Dim Index               As Integer
Dim IndexMaxOut         As Integer
Dim IndexMaxHide        As Integer
Dim LV_ColumnHeader     As ColumnHeaders

    LstVwOutput.ColumnHeaders.Clear
    LstVw_AddColumnHeader LstVwOutput, PV_ActualColumnConfig.Count
    LstVwHide.ColumnHeaders.Clear
    LstVw_AddColumnHeader LstVwHide, PV_ActualColumnConfig.Count
    
    With PV_ActualColumnConfig
        For i = 0 To .Count - 1
            Index = .Column(i).Order + 1
            If .Column(i).Visible = True Then
                Set LV_ColumnHeader = LstVwOutput.ColumnHeaders
                If IndexMaxOut < Index Then
                    IndexMaxOut = Index
                End If
            Else
                Set LV_ColumnHeader = LstVwHide.ColumnHeaders
                If IndexMaxHide < Index Then
                    IndexMaxHide = Index
                End If
            End If
            LV_ColumnHeader.Item(Index).Text = .Column(i).ColumnName
            LV_ColumnHeader.Item(Index).Tag = i
        Next
    End With
    LstVw_RemoveColumnsStartingFrom IndexMaxOut + 1, LstVwOutput
    LstVw_RemoveColumnsStartingFrom IndexMaxHide + 1, LstVwHide
    
    LstVw_CopyColumnHeaders LstVwPreview, LstVwOutput   ', "Num"
    
    AutoAjusteColumnWidth LstVwOutput
    AutoAjusteColumnWidth LstVwHide
    AutoAjusteColumnWidth LstVwPreview

End Sub

Sub SetActualColumnConfigFromCbo(LV_Cbo As ComboBox)

    If LV_Cbo.ListIndex < LV_Cbo.ListCount - 1 Then
        Me.GetColumnConfigFromTemplate LV_Cbo.ListIndex
    End If
    
End Sub

Function GetNewSpreadConfigName() As String

Dim lvName          As String
Dim i               As Integer

    If PV_NewSpreadConfig = False Then
        Exit Function
    End If
    With Me
        For i = .OptionSpredSheetTypeConfig.LBound To .OptionSpredSheetTypeConfig.UBound
            If .OptionSpredSheetTypeConfig(i).Value = True Then
                Exit For
            End If
        Next
        Select Case i
            Case Is = 0
                GetNewSpreadConfigName = "A"
            Case Is = 1
                GetNewSpreadConfigName = "P" & CalcPulsesToConfig(.txtPulsesPerSheet.Text)
            Case Is = 2
                GetNewSpreadConfigName = "T" & CalcIntervalToConfig(.txtInterval.Text)
            Case Else
                GetNewSpreadConfigName = ""
                Exit Function
        End Select
        GetNewSpreadConfigName = GetNewSpreadConfigName & "_H" & Trim$(.txtSheetCountPerFile.Text)
        .txtNewConfigtype.Text = GetNewSpreadConfigName
    End With
    
End Function

Sub GetSpreadConfigFromCbo(LV_Cbo As ComboBox)

    If LV_Cbo.ListIndex < LV_Cbo.ListCount - 1 Then
        If PV_Index_WrkSpc <> LV_Cbo.ListIndex Then
            GetSpreadConfigFromTemplate LV_Cbo.ListIndex
            PV_Index_WrkSpc = GetIndexSpreadConfig(mProject.SpreadConfig, PV_ConfigSheetList)
            If m_Project.GetIntermediateDataReady = False Or _
                    PV_Index_WrkSpc <> LV_Cbo.ListIndex Then
                m_Project.ClearIntermediateDataReady
                PV_Index_WrkSpc = LV_Cbo.ListIndex
            End If
            GenerateIntermediaData
            m_Project.SetSheetConfigured
        End If
    End If
    
End Sub

Sub GetSpreadConfigFromTemplate(Index As Integer)

    With PV_ConfigSheetList.ConfigList(Index)
        If .byFiles = True Then
            Me.OptionSpredSheetTypeConfig(0).Value = True
        End If
        If .byInterval = True Then
            Me.OptionSpredSheetTypeConfig(2).Value = True
        End If
        If .byPulses = True Then
            Me.OptionSpredSheetTypeConfig(1).Value = True
        End If
        Me.txtSheetCountPerFile = .SheetsPerSpreadSheet
        Me.txtPulsesPerSheet = .PulsesPerSheet
        Me.txtInterval = .IntervalPerSheet
    End With
    
End Sub

Sub AddSpreadConfigToTemplate(Index As Integer)

    With PV_ConfigSheetList.ConfigList(Index)
        If Me.OptionSpredSheetTypeConfig(0).Value = True Then
            .byFiles = True
        End If
        If Me.OptionSpredSheetTypeConfig(1).Value = True Then
            .byPulses = True
        End If
        If Me.OptionSpredSheetTypeConfig(2).Value = True Then
            .byInterval = True
        End If
        .SheetsPerSpreadSheet = Me.txtSheetCountPerFile
        .PulsesPerSheet = Me.txtPulsesPerSheet
        .IntervalPerSheet = Me.txtInterval
        .SpreadConfigName = Me.cboConfigType.List(Me.cboConfigType.ListCount - 1)
    End With
    
End Sub

Sub GetColumnConfigFromTemplate(Index As Integer)

Dim i           As Integer

    With PV_ColumnConfigList.Config(Index)
        PV_ActualColumnConfig.Count = .Count
        ReDim PV_ActualColumnConfig.Column(.Count - 1)
        PV_ActualColumnConfig.ColumnConfigName = .ColumnConfigName
        For i = 0 To .Count - 1
            PV_ActualColumnConfig.Column(i).ColumnName = .Column(i).ColumnName
            PV_ActualColumnConfig.Column(i).Order = .Column(i).Order
            PV_ActualColumnConfig.Column(i).Visible = .Column(i).Visible
        Next
    End With
    SetActualColumnConfig PV_ActualColumnConfig
    
End Sub

Sub AddColumnConfigToTemplate(Index As Integer)

Dim i           As Integer

    With PV_ColumnConfigList.Config(Index)
        .Count = PV_ActualColumnConfig.Count
        .ColumnConfigName = PV_ActualColumnConfig.ColumnConfigName
        ReDim .Column(.Count - 1)
        For i = 0 To .Count - 1
            .Column(i).ColumnName = PV_ActualColumnConfig.Column(i).ColumnName
            .Column(i).Order = PV_ActualColumnConfig.Column(i).Order
            .Column(i).Visible = PV_ActualColumnConfig.Column(i).Visible
        Next
    End With

End Sub

Sub SetDefaultColumnConfig()

Dim lsStr       As String
Dim i           As Long

    lsStr = Space(261)
    
    With PV_ColumnConfigList
        ReDim Preserve .Config(.Count)
        With .Config(.Count)
            .ColumnConfigName = DEFAULT_OPTION
            .Count = Pulse_Field_Count
            ReDim .Column(.Count - 1)
            For i = 0 To .Count - 1
                Pulse_Field_Header i, lsStr
                .Column(i).ColumnName = lsStr
                .Column(i).Order = i
                .Column(i).Visible = True
            Next
        End With
        .Count = .Count + 1
    End With
    
End Sub

Sub SetSpreadConfigControlStatus(lvStatus As Boolean)

    With Me
        .pictureSpreadConfigControls.Enabled = lvStatus
    End With
    
End Sub

Sub SetColumnConfigControlStatus(lvStatus As Boolean)

    With Me
        .toolbarColumnConfig.Enabled = lvStatus
    End With

End Sub

Sub SetDefaultConfig()

    With Me
        .txtInterval = 0
        .txtPulsesPerSheet = 0
        .txtSheetCountPerFile = 1
        .OptionSpredSheetTypeConfig(0).Value = True
    End With
    
End Sub

'Sub SetConfigByFiles()
'
'    With Me
'        .txtInterval.Enabled = False
'        .txtPulsesPerSheet.Enabled = False
'    End With
'
'End Sub
            
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
            
'Sub SetConfigByPulses()
'
'Dim Interval         As Double
'Dim Pulses          As Long
'
'    With Me
'        .txtInterval.Enabled = False
'        .txtPulsesPerSheet.Enabled = True
'    End With
'
'End Sub
            
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
            
'Sub setConfigByInterval()
'
'    With Me
'        .txtInterval.Enabled = True
'        .txtPulsesPerSheet.Enabled = False
'    End With
'
'End Sub

Sub setConfigByInterval()

    With Me
        .txtInterval.Enabled = True
        .txtPulsesPerSheet.Enabled = False
        SetInterval .txtInterval
    End With
    
End Sub

Sub Add_ConfigColumnDefault(Cbo As ComboBox)

Dim Index           As Integer

    Index = Get_Index_CboBx(DEFAULT_OPTION, Cbo)
    If Index <> -1 Then
        Cbo.ListIndex = Index
    Else
        SetDefaultColumnConfig
        Me.GetColumnConfigFromTemplate 0
        Me.GetColumnConfigFromActualConfig Me.LstVwColumnConfig, _
                                            Me.LstVwHideColumns, _
                                            Me.LstVwPWD
        Cbo.AddItem DEFAULT_OPTION
        Save_Config_Column PV_ColumnConfigList
        Cbo.ListIndex = 0
    End If
    
End Sub

Sub Add_ConfigSpreadDefault(Cbo As ComboBox)

Dim Index           As Integer

    Index = Get_Index_CboBx(DEFAULT_OPTION, Cbo)
    If Index <> -1 Then
        Cbo.ListIndex = Index
    Else
        SetDefaultConfig
        Cbo.AddItem DEFAULT_OPTION
        Me.AddDefaultToCboSpreadConfig
        Save_ConfigSpreadSheet PV_ConfigSheetList
        Cbo.ListIndex = 0
    End If
    
End Sub

Sub AddDefaultToCboSpreadConfig()

    Me.Add_NewSpreadConfigToList
    
End Sub

Public Sub Init_ConfigSpreadSheet()

Dim IndexSpreadConfig, IndexColumnConfig      As Integer

    Set m_clsNewItem = New clsCboNew
    Set m_clsColumnConfig = New clsCboNew
    With Me
        .LstVwColumnConfig.ListItems.Clear
        .LstVwColumnConfig.ColumnHeaders.Clear
        .LstVwHideColumns.ListItems.Clear
        .LstVwHideColumns.ColumnHeaders.Clear
        
        .cboConfigType.Clear
        Load_ConfigSpreadSheet PV_ConfigSheetList, .cboConfigType
        IndexSpreadConfig = GetIndexSpreadConfig(mProject.SpreadConfig, PV_ConfigSheetList)
        If IndexSpreadConfig >= 0 Then
            .cboConfigType.ListIndex = IndexSpreadConfig
        Else
            Add_ConfigSpreadDefault .cboConfigType
        End If
        m_clsNewItem.SetControls .cboConfigType, _
                                    frmNewCbo, _
                                    .txtNewConfigtype, _
                                    .cmdAddNewConfigType, _
                                    .cmdCancelNewConfigType, _
                                    .lblWarning, True
        .GetSpreadConfigFromCbo Me.cboConfigType
        
        .cboColumnConfig.Clear
        Load_Config_Column PV_ColumnConfigList, .cboColumnConfig
        IndexColumnConfig = GetIndexColumnConfig(mProject.ColumnConfig, PV_ColumnConfigList)
        If IndexColumnConfig >= 0 Then
            .cboColumnConfig.ListIndex = IndexColumnConfig
        Else
            Add_ConfigColumnDefault .cboColumnConfig
        End If
        m_clsColumnConfig.SetControls .cboColumnConfig, _
                                    frmNewCbo, _
                                    .txtNewColumConfig, _
                                    .cmdAddNewColumnConfig, _
                                    .cmdCancelNewColumnConfig, _
                                    .lblWarningColumnConfig, True
        Me.SetActualColumnConfigFromCbo Me.cboColumnConfig
        Me.GetColumnConfigFromActualConfig Me.LstVwColumnConfig, _
                                            Me.LstVwHideColumns, _
                                            Me.LstVwPWD
    End With
    
End Sub

Private Sub cmdCancel_Click()

End Sub

Private Sub cboColumnConfig_Click()

    Me.SetActualColumnConfigFromCbo Me.cboColumnConfig
    Me.GetColumnConfigFromActualConfig Me.LstVwColumnConfig, _
                                        Me.LstVwHideColumns, _
                                        Me.LstVwPWD
    With Me.cboColumnConfig
        If PV_Init_Form = False Then
            SaveSetting App.Title, GC_CONFIGURATION_SECTION, .Name & ".ListIndex", .ListIndex
        End If
    End With
    
End Sub

Private Sub cboConfigType_Click()

    Me.GetSpreadConfigFromCbo Me.cboConfigType
    With Me.cboConfigType
        If PV_Init_Form = False Then
            SaveSetting App.Title, GC_CONFIGURATION_SECTION, .Name & ".ListIndex", .ListIndex
        End If
    End With
    
End Sub

Private Sub cmdAccept_Click()

    modProjectFunctions.SetColumnConfig PV_ActualColumnConfig
    modProjectFunctions.SetSpreadConfig PV_ConfigSheetList.ConfigList(Me.cboConfigType.ListIndex)
    m_Project.ClearSheetGenerated
    m_Project.ClearSheetGenerating
    m_Project.SaveProject
    GV_clsExportSpreadSheet.SetLoad
    
    Unload Me
    
End Sub

Private Sub cmdCancelar_Click()

    Unload Me
    
End Sub

Private Sub Form_Load()

Dim Index           As Integer

    PV_Init_Form = True
    GV_Mdi.Set_Status_MnuProject ConfigOutput, False
    GV_Mdi.Restore_Visible_All_Mnu_For_Open
    'GV_Mdi.ProjectMnu_DisableActiveForm ConfigOutput
    Me.InitForm
    Me.Init_ConfigSpreadSheet
    'Me.InitWorkSpace
    Me.ShowProjectInfo

    'Load_Pwd_ColumnHeader Me.LstVwPwd
    Me.cmdAccept.Enabled = Not m_Project.GetExportQueued
    
    If m_Project.GetIntermediateDataReady = False Then
        With Me.cboConfigType
            Index = GetSetting(App.Title, GC_CONFIGURATION_SECTION, .Name & ".ListIndex", .ListIndex)
            If .ListCount > Index Then
                .ListIndex = Index
            End If
        End With
        With Me.cboColumnConfig
            Index = GetSetting(App.Title, GC_CONFIGURATION_SECTION, .Name & ".ListIndex", .ListIndex)
            If .ListCount > Index Then
                .ListIndex = Index
            End If
        End With
    End If
    Set_MousePointer vbDefault
    
    PV_Init_Form = False
    
End Sub

Private Sub Form_Resize()

Dim lvVertGap               As Long
Dim lvCboWidth              As Long

    With Me
        If .Height < PV_Height_Min Then
            Exit Sub
        End If
        lvVertGap = 60
        ' Top = 6840, Height =
        '.cmdAccept.Top = .Height - .cmdAccept.Height - lvVertGap
        .cmdAccept.Top = .ScaleHeight - .cmdAccept.Height - lvVertGap
        .cmdCancelar.Top = .cmdAccept.Top
        ' Top = 2280, Height = 4455
        .FrameSheetDetails.Height = .cmdAccept.Top - .FrameSheetDetails.Top - _
                                    lvVertGap
        ' Height = 3295
        .TrVwSpreadFiles.Height = .FrameSheetDetails.Height - .TrVwSpreadFiles.Top - _
                                    lvVertGap
        ' Height =
        .LstVwConfigSpreadSheets.Height = .TrVwSpreadFiles.Height
        ' Height = 4095
        .FrameListView.Height = .ScaleHeight - .FrameListView.Top - lvVertGap
        ' Height = 3495
        .LstVwPWD.Height = .FrameListView.Height - .LstVwPWD.Top - lvVertGap
        
        ' Horizontal
        .frameConfigColums.Width = .ScaleWidth - .frameConfigColums.Left - lvVertGap
        .FrameListView.Width = .frameConfigColums.Width
        .LstVwColumnConfig.Width = .frameConfigColums.Width - 2 * .LstVwColumnConfig.Left
        .LstVwHideColumns.Width = .LstVwColumnConfig.Width
        .LstVwPWD.Width = .LstVwColumnConfig.Width
        
        lvCboWidth = .frameConfigColums.Width - .cboColumnConfig.Left - lvVertGap
        If lvCboWidth < PV_Cbo_Width Then
            .cboColumnConfig.Width = lvCboWidth
            .txtNewColumConfig = lvCboWidth
        Else
            If .cboColumnConfig.Width <> PV_Cbo_Width Then
                .cboColumnConfig.Width = PV_Cbo_Width
                .txtNewColumConfig = PV_Cbo_Width
            End If
        End If
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    m_Project.DiscardWorkSpace
    GV_clsTemplateConfigSpreadSheet.ClearLoaded
    GV_Mdi.ProjectMnuUpdate
    
End Sub

Private Sub LstVwColumnConfig_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    Set PV_ColHdrOut = ColumnHeader
    PV_LstVwOutput = True
    PV_ColumnOut_Index = ColumnHeader.Index
    
    
End Sub

Private Sub LstVwHideColumns_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    Set PV_ColHdrHide = ColumnHeader
    PV_LstVwOutput = False
    PV_ColumnHide_Index = ColumnHeader.Index
    
End Sub

Private Sub m_clsColumnConfig_AddNewItem()

    Me.SetColumnConfigControlStatus True
    
End Sub

Private Sub m_clsColumnConfig_NewItemAdded()

    Me.SetColumnConfigControlStatus False
    Me.GetColumnConfigFromControls Me.LstVwColumnConfig, Me.LstVwHideColumns
    Add_NewColumnConfigToList
    Save_Config_Column PV_ColumnConfigList
    'Me.cboColumnConfig.ListIndex = Me.cboColumnConfig.ListCount - 2
    Me.SetActualColumnConfigFromCbo Me.cboColumnConfig
    Me.GetColumnConfigFromActualConfig Me.LstVwColumnConfig, _
                                        Me.LstVwHideColumns, _
                                        Me.LstVwPWD
    
End Sub

Private Sub m_clsNewItem_AddNewItem()

    PV_NewSpreadConfig = True
    Me.SetSpreadConfigControlStatus True
    Me.GetNewSpreadConfigName
    
End Sub

Private Sub m_clsNewItem_NewItemAdded()

    PV_NewSpreadConfig = False
    Me.SetSpreadConfigControlStatus False
    Add_NewSpreadConfigToList
    Save_ConfigSpreadSheet PV_ConfigSheetList
    
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
    GetNewSpreadConfigName
    
End Sub

Private Sub toolbarColumnConfig_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
        Case Is = 1
            MoveToOutput
        Case Is = 2
            MoveToHide
        Case Is = 3
            MoveToLeft
        Case Is = 4
            MoveToRight
    End Select
    
End Sub

Private Sub txtInterval_Change()

Dim Interval             As Double

    With Me
        SetInterval .txtInterval
    End With
    GetNewSpreadConfigName
    
End Sub

Private Sub txtPulsesPerSheet_Change()

Dim lPulses             As Long

    With Me
        .SetPulsesPerSheet .txtPulsesPerSheet
    End With
    GetNewSpreadConfigName
    
End Sub

Private Sub txtSheetCountPerFile_Change()

    SetSheetPerFile Me.txtSheetCountPerFile
    GetNewSpreadConfigName
    
End Sub
