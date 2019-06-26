VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.MDIForm MDIMain 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   5160
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9570
   Icon            =   "MDIMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock sckExportExcel 
      Left            =   2760
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer TimerLoadForm 
      Interval        =   200
      Left            =   2040
      Top             =   480
   End
   Begin VB.PictureBox pictureLeftToolBar 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   4320
      Left            =   0
      ScaleHeight     =   4320
      ScaleWidth      =   2010
      TabIndex        =   2
      Top             =   465
      Width           =   2010
   End
   Begin MSComctlLib.Toolbar toolbarPulses 
      Align           =   1  'Align Top
      Height          =   465
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9570
      _ExtentX        =   16880
      _ExtentY        =   820
      ButtonWidth     =   609
      ButtonHeight    =   661
      Appearance      =   1
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBarMdiMain 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   4785
      Width           =   9570
      _ExtentX        =   16880
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuNewPjt 
         Caption         =   "&Nuevo Proyecto"
      End
      Begin VB.Menu mnuOpenMenu 
         Caption         =   "Abrir Proyecto"
         Begin VB.Menu mnuOpenPjtRecient 
            Caption         =   "&Reciente"
         End
         Begin VB.Menu mnuOpenFileFromDisk 
            Caption         =   "&desde Disco"
         End
      End
      Begin VB.Menu mnuSavePjt 
         Caption         =   "&Guardar Proyecto"
      End
      Begin VB.Menu mnuClosePjt 
         Caption         =   "&Cerrar Proyecto"
      End
      Begin VB.Menu mnuSepMenuPjt 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Cerrar"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "Sali&r"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&Vista"
      Begin VB.Menu mnuToolBarVisible 
         Caption         =   "Barra de Herramientas"
      End
      Begin VB.Menu mnuLeftToolBarVisible 
         Caption         =   "Ventana Navegación"
      End
      Begin VB.Menu mnuStatusBar 
         Caption         =   "Barra de Estado"
      End
      Begin VB.Menu mnuSepMnuVista 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPulsos 
         Caption         =   "&Pulsos"
      End
      Begin VB.Menu mnuSpreadSheet 
         Caption         =   "&SpreadSheet"
      End
   End
   Begin VB.Menu mnuProject 
      Caption         =   "&Proyecto"
      Begin VB.Menu mnuSelLocationPjt 
         Caption         =   "Seleccionar &Ubicacion"
      End
      Begin VB.Menu mnuSelPulsosPjt 
         Caption         =   "Seleccionar &Pulsos"
      End
      Begin VB.Menu mnuPjtConfigFormatOutput 
         Caption         =   "&Configuración Formato de Salida"
      End
      Begin VB.Menu mnuPjtPreview 
         Caption         =   "&Pre Visualizar Salida"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPjtExportSpreadSheet 
         Caption         =   "&Generar Archivos de Salida"
      End
      Begin VB.Menu mnuPjtErrorVerification 
         Caption         =   "&Verificación de Errores"
      End
   End
   Begin VB.Menu mnuWindows 
      Caption         =   "Ven&tanas"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Ayuda"
      Begin VB.Menu mnuAbout 
         Caption         =   "&Acerca de"
      End
   End
End
Attribute VB_Name = "MDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mMnuStatusBar           As clsMnuControlVisible
Private mSckData                As String
Private PV_Ini_Form             As Long
Private PV_WDT                  As Long
Private PV_Form_Export          As frmExportSpreadSheet
Private PV_Time_No_Comm         As Long
Private PV_Time_Send_Alive      As Long
Private PV_Alive_Counter        As Long

Sub SendAliveMsg()

Dim lsMsg                       As String

    Me.RefreshTimeToSendAlive
    PV_Alive_Counter = PV_Alive_Counter + 1
    lsMsg = GC_MSG_ALIVE & "," & PV_Alive_Counter
    Me.SendMessage lsMsg
    
End Sub

Sub SocketConnect()

Dim lvDefault           As Integer

    With Me.sckExportExcel
        lvDefault = 28127
        .RemoteHost = .LocalIP
        .RemotePort = GetSetting(App.Title, GC_CONFIGURATION_SECTION, .Name & ".LocalPort", lvDefault)
        .Connect
    End With
    
End Sub

Function IsTimeToSendAlive() As Boolean

    IsTimeToSendAlive = False
    If GetTickCount - PV_Time_Send_Alive > GV_TimeOutSendAlive Then
        IsTimeToSendAlive = True
    End If
    
End Function

Sub RefreshTimeToSendAlive()

    PV_Time_Send_Alive = GetTickCount
    
End Sub

Sub RefreshTimerComm()

    PV_Time_No_Comm = GetTickCount
    
End Sub

Function IsTimerCommCompleted() As Boolean

    IsTimerCommCompleted = False
    If GetTickCount - PV_Time_No_Comm > GV_TimeOutComm Then
        IsTimerCommCompleted = True
    End If
    
End Function

Sub KillSecondInstance()

    'KillApp
    
End Sub

Sub FinalizarApp()

Dim lvStatus            As Boolean

    If IsNothing(PV_Form_Export) = False Then
        lvStatus = Pulse_CancelXlsProcess
        If lvStatus = False Then
            Exit Sub
        End If
    End If
    Me.Close_Project
    Me.CloseAllInstances
    End
    
End Sub

Sub RefreshSckState()

Dim lvState                     As String

    lvState = ""
    Select Case Me.sckExportExcel.state
        Case Is = StateConstants.sckConnected
            lvState = "Connected"
        Case Is = StateConstants.sckConnecting
            lvState = "Connecting"
        Case Is = StateConstants.sckClosed
            lvState = "Closed"
        Case Is = StateConstants.sckClosing
            lvState = "Closing"
        Case Is = StateConstants.sckConnectionPending
            lvState = "ConnectionPending"
        Case Is = StateConstants.sckError
            lvState = "Error"
        Case Is = StateConstants.sckHostResolved
            lvState = "HostResolved"
        Case Is = StateConstants.sckListening
            lvState = "Listening"
        Case Is = StateConstants.sckOpen
            lvState = "Open"
        Case Is = StateConstants.sckResolvingHost
            lvState = "ResolvingHost"
    End Select
    Me.StatusBarMdiMain.Panels(1).Text = lvState
    
End Sub

Sub SendRunXls()

Dim i           As Integer

    If GV_PrevInstance = False Then
        With BackGroundProjectList
            If .ListIndex >= 0 Then
                i = BackGroundProjectList.ListIndex
            Else
                For i = 0 To .Count - 1
                    If .ProjectList(i).ProjectStarted = False Then
                        .ListIndex = i
                        Exit For
                    End If
                Next
            End If
            If .ListIndex >= 0 Then
                With .ProjectList(.ListIndex)
                    If .ProjectStarted = False Then
                        Me.SendMessage "START_MSG,RUNXLS,"
                        Me.SendMessage .ProjectName
                        Me.SendMessage ","
                        Me.SendMessage .ProjectPath
                        Me.SendMessage ","
                        Me.SendMessage .WorkSpacePath
                        Me.SendMessage ","
                        Me.SendMessage .OutputPath
                        Me.SendMessage ",END_MSG"
                        .ProjectStarted = True
                        .TickIni = GetTickCount
                    Else
                        If .GenerationDone = True Then
                            BackGroundProjectList.ListIndex = -1
                        End If
                    End If
                End With
            End If
        End With
    End If

End Sub

Sub RefreshWDT()

    PV_WDT = GetTickCount
    
End Sub
Sub RefreshTime()

    'If PV_WDT = True Then
        PV_Ini_Form = GetTickCount
    'End If
    
End Sub

Function IsTimeToKillSecondInstance() As Boolean

Dim lvEnlased           As Double

    lvEnlased = GetTickCount - PV_WDT
    IsTimeToKillSecondInstance = False
    If lvEnlased >= GV_TimeToKillSecond Then
        IsTimeToKillSecondInstance = True
    End If

End Function

Function IsTimeToKillMySelf() As Boolean

Dim lvEnlased           As Double

    lvEnlased = GetTickCount - PV_WDT
    IsTimeToKillMySelf = False
    If lvEnlased >= GV_TimeToKillMySelf Then
        IsTimeToKillMySelf = True
    End If

End Function

Function IsTimeToDiscardSckData() As Boolean

Dim lvEnlased           As Double

    lvEnlased = GetTickCount - PV_Ini_Form
    IsTimeToDiscardSckData = False
    If lvEnlased >= 1500 Then
        IsTimeToDiscardSckData = True
    End If

End Function

Sub DiscardSckData()

    If mSckData <> "" Then
        mSckData = ""
    End If
    
End Sub

Sub ProcessMsg(lsStr As String)

Dim lsProjectMsg()      As String
Dim i                   As Integer
Dim IndexMsg            As Message_Header_Const

    lsProjectMsg = Split(lsStr, ",")
    IndexMsg = Parse_Header_Message(lsProjectMsg)
    If Verify_Length_Msg(lsProjectMsg, IndexMsg) = False Then
        Exit Sub
    End If
    Select Case IndexMsg
        Case Is = Message_Header_Const.MSG_RUNXLS
            mProject.ProjectName = lsProjectMsg(1)
            mProject.ProjectPath = lsProjectMsg(2)
            mProject.WorkSpacePath = lsProjectMsg(3)
            mProject.OutputPath = lsProjectMsg(4)
            m_Project.LoadProject lsProjectMsg(1), lsProjectMsg(2)
            m_Project.LoadWorkSpace
            GV_clsExportSpreadSheet.SetLoad
        Case Is = Message_Header_Const.MSG_ERROR
            Exit Sub
        Case Is = Message_Header_Const.MSG_ALIVE
            Me.RefreshTimerComm
        Case Else
            If IsNothing(PV_Form_Export) = False Then
                PV_Form_Export.ParseMsg lsProjectMsg, IndexMsg
            End If
    End Select

End Sub

Sub ParseSckData(ByRef lsData As String)

Dim i               As Integer
Dim lsStr           As String

    Do
        i = InStr(lsData, "START_MSG,")
        If i Then
            lsData = Mid$(lsData, i + Len("START_MSG,"))
        Else
            lsData = ""
            Exit Do
        End If
        i = InStr(lsData, ",END_MSG")
        If i Then
            lsStr = Left$(lsData, i - 1)
            Me.ProcessMsg lsStr
            lsData = Mid$(lsData, i + 7)
            Me.RefreshTime
        Else
            Exit Do
        End If
    Loop While Len(lsData)
    
End Sub

Sub CallDebug()

Dim IndexFile   As Integer
Dim PulseQty    As Long
Dim PathName    As String
Dim Flag        As Boolean
Dim lsFlag      As String

    PathName = Obtener_Ultimo_Path_Form("frmOpenPulseDialog", App.Path)
    
    lsFlag = GetSetting(App.Title, GC_DEBUG, "Call Debug Routine", 0)
    If IsNumeric(lsFlag) = True Then
        If Val(lsFlag) Then
            Flag = True
        Else
            Flag = False
        End If
    Else
        Flag = False
    End If
    If Flag = True Then
        'SaveSettingDebug "Call Debug Routine", 1
        Pulse_Debug IndexFile, PulseQty, PathName
    Else
        SaveSettingDebug "Call Debug Routine", 0
    End If
    
End Sub

Sub CloseAllFormFrom(LV_MDI As MDIMain, ExceptForm As Form, Optional HideForm As Boolean = True)

Dim LV_Form         As Form
Dim lvFlag          As Boolean

    On Error Resume Next
    lvFlag = False
    For Each LV_Form In Forms
        lvFlag = LV_Form.MDIChild
        If IsNothing(ExceptForm) = False Then
            If LV_Form.Name = ExceptForm.Name Then
                lvFlag = False
                If HideForm = True Then
                    LV_Form.Hide
                End If
            End If
        End If
        If lvFlag = True Then
            lvFlag = False
            Unload LV_Form
        End If
    Next
    'On Error GoTo 0
    
End Sub

Sub CloseAllInstances(Optional FormExportHide As Boolean = True)

Dim LV_Form         As Form

    Me.CloseAllFormFrom Me, PV_Form_Export, FormExportHide
    Pulse_Destroy_All
    
End Sub

Sub Close_Project(Optional FormExportHide As Boolean = True)

    If m_Project.IsClosedProject = False And m_Project.IsEmptyProject = False Then
        If m_Project.IsNewProject = True Then
            m_Project.DiscardProject
        Else
            m_Project.CloseProject
        End If
    End If
    GV_Project_Closed = True
    Me.CloseAllInstances FormExportHide
    
End Sub

' Verificar los Menù habilitados segùn forms habilitados
'Public Sub ProjectMnu_UpdateAfterFormClosed(Form_Id As mnuProjectConstant)
'
'    SetMnuProject True
'    ProjectMnu_Enable_By_Status
'
'End Sub

Public Sub ProjectMnuUpdate() '(Form_Id As mnuProjectConstant)

    SetMnuProject True
    ProjectMnuDisableByStatus

End Sub

Public Sub ProjectMnu_DisableActiveForm(Form_Id As mnuProjectConstant)

    Set_Status_MnuProject Form_Id, False
    
End Sub

Public Sub ProjectMnuDisableByStatus()

    If m_Project.IsClosedProject = True Or m_Project.IsEmptyProject = True Then
        Set_Status_MnuProject ConfigOutput, False
        Set_Status_MnuProject GenOutput, False
        Set_Status_MnuProject SelPathMission, True
        Set_Status_MnuProject SelPathProject, False
    End If
    If m_Project.GetMissionSelected = False Then
        Set_Status_MnuProject ConfigOutput, False
        Set_Status_MnuProject GenOutput, False
        Set_Status_MnuProject SelPathProject, False
    End If
    If m_Project.GetPulsesAnalized = False Then
        Set_Status_MnuProject ConfigOutput, False
        Set_Status_MnuProject GenOutput, False
        'Set_Status_MnuProject SelPathMission, True
        'Set_Status_MnuProject SelPathProject, False
    End If
    If m_Project.GetIntermediateDataReady = False Then
        Set_Status_MnuProject GenOutput, False
        'Set_Status_MnuProject PreviewOutput, False
    End If
    If m_Project.GetSheetGenerated = False Then
        'Do nothing
    End If
    
End Sub

Function GetMdiCaption(lMdiForm As MDIForm) As String

    GetMdiCaption = "Conversor de Pulsos"      ' Leer_Ini(GV_File_Settings, lMdiForm.Name, "Caption", "Pulses Analize")
    
End Function

Sub Init_Form_Main()

    With Me
        Set mMnuStatusBar = New clsMnuControlVisible
        mMnuStatusBar.SetControl .mnuStatusBar, .StatusBarMdiMain
        .Caption = GetMdiCaption(Me)
        'Grabar_Ini GV_File_Settings, "MDI Main", "Caption", .Caption
        Load_Mdi_Params Me
        GetControlCheckedSetting .mnuToolBarVisible
        GetControlCheckedSetting .mnuLeftToolBarVisible
        If .mnuToolBarVisible.Checked = False Then
            .toolbarPulses.Visible = False
        End If
        If .mnuLeftToolBarVisible.Checked = False Then
            .pictureLeftToolBar.Visible = False
        End If
    End With
    
End Sub

Function LoadInstanceFormConfigSpreadSheet() As Integer

Dim frmTemp         As frmConfigSpreadSheets

    If GV_clsConfigSpreadSheet.GetLoaded = True Then
        Exit Function
    End If
    
    Set_MousePointer vbHourglass
    GV_clsConfigSpreadSheet.SetLoaded
    
    Set frmTemp = New frmConfigSpreadSheets
    
    Load frmTemp
    frmTemp.WindowState = vbMaximized
    frmTemp.Show
    

End Function

Function LoadInstanceFormExportSpreadSheet() As Integer

Dim frmTemp         As frmExportSpreadSheet

    If GV_clsExportSpreadSheet.GetLoaded = True Then
        Exit Function
    End If
    
    GV_clsExportSpreadSheet.SetLoaded
    
    If IsNothing(PV_Form_Export) = True Then
        Set frmTemp = New frmExportSpreadSheet
        
        Set_MousePointer vbHourglass
        Load frmTemp
        
        If GV_PrevInstance = False Then
            Set PV_Form_Export = frmTemp
            frmTemp.Show
        Else
    '        Set PV_Form_Export = frmTemp
            If GetSettingBooleanParameter(GC_VISIBLE_SND_INSTANCE, False) = False Then
                GV_Mdi.Hide
            Else
                frmTemp.Show
            End If
            With frmTemp
                .chkRunInBackGround.Value = 0
                .cmdAccept.Value = True
            End With
        End If
    Else
        PV_Form_Export.UpdateActualProject
        PV_Form_Export.Show
    End If
    
End Function

Function LoadInstanceFormProjectFromFile()

Dim frmTemp         As frmPjtFromFile

    If GV_clsPjtFromFile.GetLoaded = True Then
        Exit Function
    End If
    
    Set_MousePointer vbHourglass
    GV_clsPjtFromFile.SetLoaded
    
    Set frmTemp = New frmPjtFromFile
    
    Load frmTemp
    frmTemp.Show

End Function

Function LoadInstanceFormProjectFromList()

Dim frmTemp         As frmPjtFromList

    If GV_clsPjtFromList.GetLoaded = True Then
        Exit Function
    End If
    
    Set_MousePointer vbHourglass
    GV_clsPjtFromList.SetLoaded
    
    Set frmTemp = New frmPjtFromList
    
    Load frmTemp
    frmTemp.Show

End Function

Function LoadInstanceFormProjectSelLocation()

Dim frmTemp         As frmProjectSelLocation

    If GV_clsProjectSelLocation.GetLoaded = True Then
        Exit Function
    End If
    
    Set_MousePointer vbHourglass
    GV_clsProjectSelLocation.SetLoaded
    
    Set frmTemp = New frmProjectSelLocation
    
    Load frmTemp
    frmTemp.Show

End Function

Function LoadInstanceFormNavProject() As Integer

End Function

Function LoadInstanceSecondinstance() As Integer

    If GV_clsSecondInstance.GetLoaded = True Then
        Exit Function
    End If
    
    GV_clsSecondInstance.SetLoaded
    
    Me.StartUpApp
    
End Function

Function LoadInstanceFormErrorView() As Integer

Dim frmTemp         As frmErrorView

    If GV_clsErrorView.GetLoaded = True Then
        Exit Function
    End If
    
    Set_MousePointer vbHourglass
    GV_clsErrorView.SetLoaded
    
    Set frmTemp = New frmErrorView
    
    Load frmTemp
    frmTemp.Show
    frmTemp.WindowState = vbMaximized

End Function

Function LoadInstanceFormTemplateConfigSS() As Integer

Dim frmTemp         As frmTemplateConfigSpreadSheet

    If GV_clsTemplateConfigSpreadSheet.GetLoaded = True Then
        Exit Function
    End If
    
    Set_MousePointer vbHourglass
    GV_clsTemplateConfigSpreadSheet.SetLoaded
    
    Set frmTemp = New frmTemplateConfigSpreadSheet
    
    Load frmTemp
    frmTemp.Show
    frmTemp.WindowState = vbMaximized
    
End Function


Function LoadInstanceFormOpenPulseDialog() As Integer

Dim frmTemp         As frmOpenPulseDialog

    If GV_clsOpenPulseDialog.GetLoaded = True Then
        Exit Function
    End If
    
    Set_MousePointer vbHourglass
    GV_clsOpenPulseDialog.SetLoaded
    
    Set frmTemp = New frmOpenPulseDialog
    
    Load frmTemp
    frmTemp.Show
    frmTemp.WindowState = vbMaximized
    
End Function

Function LoadInstanceFormSpreadSheetView()

Dim frmTemp         As frmSpreadSheetView

    If GV_clsSpreadSheetView.GetLoaded = True Then
        Exit Function
    End If
    
    Set_MousePointer vbHourglass
    GV_clsSpreadSheetView.SetLoaded
    
    Set frmTemp = New frmSpreadSheetView
    
    Load frmTemp
    frmTemp.Show

End Function

Function LoadInstanceFormStart() As Integer

Dim LV_frmStart     As frmStart

    If GV_clsStart.GetLoaded = True Then
        Exit Function
    End If
    
    Set_MousePointer vbHourglass
    GV_clsStart.SetLoaded
    
    Set LV_frmStart = New frmStart
    Load LV_frmStart
    
    LV_frmStart.Show vbModal
    LV_frmStart.Caption = "Start"
    
End Function

Function LoadInstanceFormStartUp() As Integer

Dim LV_frmStartUp     As frmStartUp

    If GV_clsStartUp.GetLoaded = True Then
        Exit Function
    End If
    
    Set_MousePointer vbHourglass
    GV_clsStartUp.SetLoaded
    
    Set LV_frmStartUp = New frmStartUp
    Load LV_frmStartUp
    
    LV_frmStartUp.Show
    LV_frmStartUp.Caption = "Start"
    
End Function

Sub NewProject()

    m_Project.CloseProject
    'm_Project.ClearProject
    m_Project.NewProject
    GV_clsOpenPulseDialog.SetLoad
    GV_Project_Opened = True
    Me.Set_Visible_Mnu_For_Open True
    Me.mnuNewPjt.Enabled = False
    
End Sub

Sub ProjectFromFile()

    m_Project.CloseProject
    m_Project.DiscardProject
    GV_clsPjtFromFile.SetLoad
    Me.Set_Visible_Mnu_For_Open True
    Me.mnuOpenFileFromDisk.Enabled = False
    
End Sub

Sub ProjectFromListFile()

    m_Project.CloseProject
    GV_clsPjtFromList.SetLoad
    Me.Set_Visible_Mnu_For_Open True
    Me.mnuOpenPjtRecient.Enabled = False
    
End Sub

Sub MdiOpenProjectMode()

    SetMnuProject True
    Set_Visible_Mnu_For_Open True
    
End Sub

Public Sub SendMessage(lvMsg As String)

    With Me.sckExportExcel
        If .state = sckConnected Then
            .SendData lvMsg
        End If
    End With
    
End Sub

Private Sub Set_Mdi_For_ListProject()

    GV_Project_Closed = False
    SetMnuProject False
    Set_Visible_Mnu_For_Open True
    
End Sub

Sub Set_Mdi_For_Project()

    GV_Project_Opened = False
    With Me
        SetMnuProject True
    End With
    
End Sub

Sub Set_Status_MnuProject(Mnu_Id As mnuProjectConstant, lvStatus As Boolean)

    With Me
        Select Case Mnu_Id
            Case Is = mnuProjectConstant.MnuEmpty
                'Do Nothing
            Case Is = mnuProjectConstant.SelPathProject
                .mnuSelLocationPjt.Enabled = lvStatus
            Case Is = mnuProjectConstant.SelPathMission
                .mnuSelPulsosPjt.Enabled = lvStatus
            Case Is = mnuProjectConstant.ConfigOutput
                .mnuPjtConfigFormatOutput.Enabled = lvStatus
            Case Is = mnuProjectConstant.PreviewOutput
                .mnuPjtPreview.Enabled = lvStatus
            Case Is = mnuProjectConstant.GenOutput
                .mnuPjtExportSpreadSheet.Enabled = lvStatus     ' CalcMnuExportState(lvStatus)
            Case Is = mnuProjectConstant.VerifyErrors
                .mnuPjtErrorVerification.Enabled = lvStatus
        End Select
    End With
        
End Sub

Sub Set_Visible_MnuProject(Optional lvExcept As mnuProjectConstant = MnuEmpty)

    With Me
        SetMnuProject True
        Select Case lvExcept
            Case Is = mnuProjectConstant.MnuEmpty
                'Do Nothing
            Case Is = mnuProjectConstant.SelPathProject
                .mnuSelLocationPjt.Enabled = False
            Case Is = mnuProjectConstant.SelPathMission
                .mnuSelPulsosPjt.Enabled = False
            Case Is = mnuProjectConstant.ConfigOutput
                .mnuPjtConfigFormatOutput.Enabled = False
            Case Is = mnuProjectConstant.PreviewOutput
                .mnuPjtPreview.Enabled = False
            Case Is = mnuProjectConstant.GenOutput
                .mnuPjtExportSpreadSheet.Enabled = CalcMnuExportState(False)
            Case Is = mnuProjectConstant.VerifyErrors
                .mnuPjtErrorVerification.Enabled = False
        End Select
    End With
        
End Sub

Sub Restore_Visible_All_Mnu_For_Open()

    Me.Set_Visible_Mnu_For_Open True
    
End Sub

Sub Restore_Visible_MnuProject(lvMnu As mnuProjectConstant)

    If m_Project.IsClosedProject = True Then
        Exit Sub
    End If
    With Me
        Select Case lvMnu
            Case Is = mnuProjectConstant.MnuEmpty
                'Do Nothing
            Case Is = mnuProjectConstant.SelPathProject
                .mnuSelLocationPjt.Enabled = True
            Case Is = mnuProjectConstant.SelPathMission
                .mnuSelPulsosPjt.Enabled = True
            Case Is = mnuProjectConstant.ConfigOutput
                .mnuPjtConfigFormatOutput.Enabled = True
            Case Is = mnuProjectConstant.PreviewOutput
                .mnuPjtPreview.Enabled = True
            Case Is = mnuProjectConstant.GenOutput
                .mnuPjtExportSpreadSheet.Enabled = True
            Case Is = mnuProjectConstant.VerifyErrors
                .mnuPjtErrorVerification.Enabled = True
        End Select
    End With
        
End Sub

Sub StartSckServer()

Dim lvDefault           As Integer

    On Error Resume Next
    With Me.sckExportExcel
        lvDefault = 28127
        lvDefault = GetSetting(App.Title, GC_CONFIGURATION_SECTION, .Name & ".LocalPort", lvDefault) - 1
        Do
            lvDefault = lvDefault + 1
            .LocalPort = lvDefault
            SaveSetting App.Title, GC_CONFIGURATION_SECTION, .Name & ".LocalPort", lvDefault
        Loop Until .LocalPort = lvDefault
        .Listen
    End With
    'On Error GoTo 0
    
End Sub

Sub Set_Visible_Mnu_For_Open(lvState As Boolean)

    With Me
        .mnuNewPjt.Enabled = lvState
        .mnuOpenMenu.Enabled = lvState
        If m_Project.IsThereProjectList = False Then
            .mnuOpenPjtRecient.Enabled = False
        Else
            .mnuOpenPjtRecient.Enabled = lvState
        End If
        .mnuOpenFileFromDisk.Enabled = lvState
    End With
    
End Sub

Function CalcMnuExportState(ByVal lvState As Boolean) As Boolean

    If lvState = False Then
        If BackGroundProjectList.Count Then
            lvState = True
        End If
    End If
    CalcMnuExportState = lvState

End Function

Sub SetMnuProject(lvState As Boolean)

    With Me
        .mnuSelLocationPjt.Enabled = lvState
        .mnuSelPulsosPjt.Enabled = lvState
        If m_Project.IsNewProject = True Then
            .mnuPjtErrorVerification.Enabled = False
            .mnuPjtConfigFormatOutput.Enabled = False
            .mnuPjtExportSpreadSheet.Enabled = CalcMnuExportState(False)
            .mnuPjtPreview.Enabled = False
            .mnuClosePjt.Enabled = False
            .mnuSavePjt.Enabled = False
        Else
            .mnuPjtErrorVerification.Enabled = lvState
            .mnuPjtConfigFormatOutput.Enabled = lvState
            .mnuPjtExportSpreadSheet.Enabled = CalcMnuExportState(lvState)
            .mnuPjtPreview.Enabled = lvState
            .mnuClosePjt.Enabled = lvState
            .mnuSavePjt.Enabled = lvState
        End If
    End With
    
End Sub

'Verificar Que Form debe ser cerrado
Sub VerifyFormOpened()

End Sub


Sub StartUpApp()

Dim lvExe           As String

    'SaveSettingDebug "Shell_Tick_Ini", "Start"

    'On Error GoTo StartUpAppErr
    lvExe = Retroceder_Path(App.Path) & "\Exe\" & App.Title & ".exe"
    lvExe = GetSetting(App.Title, GC_CONFIGURATION_SECTION, "Exe_File", lvExe)
    SaveSetting App.Title, GC_CONFIGURATION_SECTION, "Exe_File", lvExe
    
    'SaveSettingDebug "Exe_File", lvExe
    'SaveSettingDebug "Shell_Tick_Start", lvExe
    Shell (lvExe)
    
    'On Error GoTo 0
    Exit Sub
    
StartUpAppErr:

    lvExe = Retroceder_Path(App.Path) & "\Exe\" & App.Title & ".exe"
    SaveSetting App.Title, GC_CONFIGURATION_SECTION, "Exe_File", lvExe
    'SaveSettingDebug "Shell_Tick_Error", lvExe
    
    Shell (lvExe)

End Sub

Private Sub MDIForm_Load()

Dim LV_frmStart     As frmStart

    CallDebug
    Init_Vars
    Init_Form_Main
    m_Project.Init_Pool_Memory
    m_Project.Init_Xls_Dll
    
    'GV_clsStart.SetLoad
    If GV_PrevInstance = False Then
        'SaveSettingDebug "GV_PrevInstance", "False"
        GV_clsStartUp.SetLoad
        Me.StartSckServer
        If GetSettingBooleanParameter(GC_START_UP_DISABLE, False) = False Then
            If GetSettingBooleanParameter(GC_INVERSE_PREV_INSTANCE, False) = False Then
                GV_clsSecondInstance.SetLoad
            End If
        End If
    Else
        'SaveSettingDebug "GV_PrevInstance", "False"
        'PV_WDT = True
        If GetSettingBooleanParameter(GC_INVERSE_PREV_INSTANCE, False) = True Then
            GV_clsSecondInstance.SetLoad
        End If
        Me.RefreshTime
        Me.RefreshWDT
        Me.SocketConnect
    End If
    'GV_clsOpenPulseDialog.SetLoad
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)

    CloseLogFile
    Pulse_Close_Log
    'Close_Log
    Save_Mdi_Params Me
    
End Sub

Private Sub mnuAbout_Click()

    Load frmAbout
    frmAbout.Show
    
End Sub

Private Sub mnuPjtConfigFormatOutput_Click()

    Me.CloseAllFormFrom Me, PV_Form_Export
    GV_clsTemplateConfigSpreadSheet.SetLoad
    
End Sub

Private Sub mnuPjtErrorVerification_Click()

    Me.CloseAllFormFrom Me, PV_Form_Export
    GV_clsErrorView.SetLoad

End Sub

Private Sub mnuPjtExportSpreadSheet_Click()

    Me.CloseAllFormFrom Me, PV_Form_Export
    GV_clsExportSpreadSheet.SetLoad
    
End Sub

'Private Sub mnuNew_Click()
'
'    CloseAllInstances
'    Pulse_Destroy_All
'
'End Sub

Sub mnuOpen_Click()

    CloseAllInstances
    Pulse_Destroy_All
    GV_clsOpenPulseDialog.SetLoad
    
End Sub

Private Sub mnuClosePjt_Click()

    Me.Close_Project
    
End Sub

Private Sub mnuLeftToolBarVisible_Click()

Dim lvState         As Integer

    With Me.mnuLeftToolBarVisible
        If .Checked = True Then
            .Checked = False
        Else
            .Checked = True
        End If
        lvState = .Checked
        Me.pictureLeftToolBar.Visible = .Checked
        SaveSetting App.Title, GC_CONFIGURATION_SECTION, .Name & ".Checked", lvState
    End With
    
End Sub

Private Sub mnuNewPjt_Click()

    Me.CloseAllInstances
    Me.NewProject
    
End Sub

Private Sub mnuOpenFileFromDisk_Click()

    Me.CloseAllInstances
    Me.ProjectFromFile
    
End Sub

Private Sub mnuOpenPjtRecient_Click()

    Me.CloseAllInstances
    Me.ProjectFromListFile
    
End Sub

Private Sub mnuPjtPreview_Click()

    Me.CloseAllFormFrom Me, PV_Form_Export
    GV_clsConfigSpreadSheet.SetLoad
    
End Sub

Private Sub mnuQuit_Click()

    End
    
End Sub

Private Sub mnuSavePjt_Click()

    m_Project.SaveProject
    
End Sub

Private Sub mnuSelLocationPjt_Click()

    Me.CloseAllFormFrom Me, PV_Form_Export
    GV_clsProjectSelLocation.SetLoad
    
End Sub

Private Sub mnuSelPulsosPjt_Click()

    Me.CloseAllFormFrom Me, PV_Form_Export
    GV_clsOpenPulseDialog.SetLoad
    
End Sub

Public Sub mnuToolBarVisible_Click()

Dim lvState         As Integer

    With Me.mnuToolBarVisible
        If .Checked = True Then
            .Checked = False
        Else
            .Checked = True
        End If
        lvState = .Checked
        Me.toolbarPulses.Visible = .Checked
        SaveSetting App.Title, GC_CONFIGURATION_SECTION, .Name & ".Checked", lvState
        
    End With
    
End Sub

Private Sub sckExportExcel_Close()

    With Me.sckExportExcel
        .Close
        If GV_PrevInstance = False Then
            .LocalPort = .LocalPort
            Me.StartSckServer
            DiscardBackGroundProcess
        Else
            Me.SocketConnect
        End If
    End With
    
End Sub


Private Sub sckExportExcel_Connect()

Dim i           As Long
    
    RefreshTime
    RefreshTimerComm
    RefreshTimeToSendAlive
    
End Sub

Private Sub sckExportExcel_ConnectionRequest(ByVal requestID As Long)

Dim i           As Integer

    Me.sckExportExcel.Close
'    For i = 1 To Me.sckExportExcel.UBound
'        If Me.sckExportExcel(i).state = sckClosed Then
'            Exit For
'        End If
'    Next
'    If i > Me.sckExportExcel.UBound Then
'        Load Me.sckExportExcel(i)
'    End If
    With Me.sckExportExcel
        .LocalPort = 0
        .Accept requestID
    End With
    'Me.sckExportExcel(Index).Listen
    
End Sub

Private Sub sckExportExcel_DataArrival(ByVal bytesTotal As Long)

Dim lvData          As String

    With Me.sckExportExcel
        .GetData lvData
        mSckData = mSckData & lvData
        If IsNothing(PV_Form_Export) = False Then
            PV_Form_Export.ShowSckData lvData
        End If
        Me.ParseSckData mSckData
        RefreshTimerComm
    End With
    
End Sub

Private Sub sckExportExcel_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

    With Me.sckExportExcel
        .Close
        If GV_PrevInstance = False Then
            Me.StartSckServer
        Else
            Me.SocketConnect
        End If
    End With
    
End Sub


Private Sub TimerLoadForm_Timer()

Dim lv_DeltaTime        As Date
Static LV_Tmr_Count     As Integer

    
    If GV_Restore_Project_Mnu = True Then
        SetMnuProject True
    End If
    If GV_Project_Opened = True Then
        Set_Mdi_For_Project
    End If
    If GV_Project_Closed = True Then
        Set_Mdi_For_ListProject
    End If
    
    If GV_clsErrorView.GetLoad = True Then
        Me.LoadInstanceFormErrorView
    End If
    If GV_clsTemplateConfigSpreadSheet.GetLoad = True Then
        Me.LoadInstanceFormTemplateConfigSS
    End If
    If GV_clsOpenPulseDialog.GetLoad = True Then
        Me.LoadInstanceFormOpenPulseDialog
    End If
    If GV_clsSpreadSheetView.GetLoad = True Then
        Me.LoadInstanceFormSpreadSheetView
    End If
    If GV_clsExportSpreadSheet.GetLoad = True Then
        Me.LoadInstanceFormExportSpreadSheet
    End If
    If GV_clsConfigSpreadSheet.GetLoad = True Then
        Me.LoadInstanceFormConfigSpreadSheet
    End If
    If GV_clsStart.GetLoad = True Then
        Me.LoadInstanceFormStart
    End If
    If GV_clsStartUp.GetLoad = True Then
        Me.LoadInstanceFormStartUp
    End If
    If GV_clsProjectSelLocation.GetLoad = True Then
        Me.LoadInstanceFormProjectSelLocation
    End If
    If GV_clsPjtFromFile.GetLoad = True Then
        Me.LoadInstanceFormProjectFromFile
    End If
    If GV_clsPjtFromList.GetLoad = True Then
        Me.LoadInstanceFormProjectFromList
    End If
    If GV_clsSecondInstance.GetLoad = True Then
        Me.LoadInstanceSecondinstance
    Else
        If Me.sckExportExcel.state <> sckConnected Then
        End If
    End If
    
'    If PV_WDT = True Then
'        lv_DeltaTime = Now() - PV_Ini_Form
'        If lv_DeltaTime > 0.0035 Then
'            End
'        End If
'    End If
    If Me.sckExportExcel.state = sckConnected Then
        Me.RefreshWDT
        If Me.IsTimeToDiscardSckData = True Then
            Me.DiscardSckData
        End If
        If GV_PrevInstance = False Then
            Me.SendRunXls
            If IsTimerCommCompleted = True Then
                KillSecondInstance
            End If
        Else
            If IsTimeToSendAlive = True Then
                SendAliveMsg
            End If
        End If
    Else
        If GV_PrevInstance = True Then
            If Find_PrevInstance(App.Title) = False Then
                m_Project.RevertParameters
                m_Project.SaveProject True
                If GetSettingBooleanParameter(GC_INVERSE_PREV_INSTANCE, False) = False Then
                    Me.FinalizarApp
                End If
            Else
                If Me.IsTimeToKillMySelf = True Then
                    Me.FinalizarApp
                End If
            End If
        Else
            LV_Tmr_Count = LV_Tmr_Count + 1
            'SaveSettingDebug "Tmr_Tick(" & Trim$(LV_Tmr_Count) & ")", GetTickCount
            If Find_PrevInstance(App.Title) = False Then
                If GV_clsSecondInstance.GetLoaded = True Then
                    GV_clsSecondInstance.ClearLoaded
                    GV_clsSecondInstance.SetLoad
                End If
            End If
        End If
    End If
    
    Me.RefreshSckState
    
End Sub

