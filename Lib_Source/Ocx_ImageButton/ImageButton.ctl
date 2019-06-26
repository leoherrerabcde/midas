VERSION 5.00
Begin VB.UserControl ImageButton 
   ClientHeight    =   645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1515
   ScaleHeight     =   645
   ScaleWidth      =   1515
   Begin VB.PictureBox PictureBotones 
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   1035
      TabIndex        =   0
      Top             =   0
      Width           =   1095
      Begin VB.Image ImageButton 
         Height          =   375
         Index           =   0
         Left            =   0
         Top             =   0
         Width           =   375
      End
   End
End
Attribute VB_Name = "ImageButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Module    : ImageButton
' Author    : Leo Herrera
' Date      : 23/10/2011
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

Private PV_pictureHeight                As Integer
Private PV_pictureWidth                 As Integer
Private PV_pictureInterval              As Integer
Private PV_LandscapeState               As Boolean
Private PV_PictureCount                 As Integer
Private PV_pictureLeft                  As Integer
Private PV_pictureTop                   As Integer
Private PV_ResizeState                  As Boolean
Private PV_ResizeControl                As ResizeStatus

Property Let PictureHeight(Height As Integer)

End Property

Property Let PictureWidth(Width As Integer)

End Property

Property Let PictureInterval(Interval As Integer)

End Property

Property Let Orientation(Orient As Integer)

End Property

Property Get PictureCount() As Integer

End Property

Property Let AddItem(PictureFile As String)

End Property

Property Let PictureImageButton(PictureFile As String, Index As Integer)

End Property

Property Let PictureLeft(Left As Integer)

End Property

Property Let PictureTop(Top As Integer)

End Property

Property Let ReSizeState(Value As Boolean)

End Property


Private Sub PictureBotones_Click()

End Sub

Private Sub PictureBotones_Resize()

    If PV_ResizeState = True Then
    End If
    
End Sub

Private Sub SetResizeStatus(Index As ResizeObjectConstant, lvNewVal As Integer)

    With PV_ResizeControl
        Select Case Index
            Case ResizeObjectConstant.ContainerObject
                .CointainerResizeCounter = lvNewVal
            Case ResizeObjectConstant.ControlObject
                .ControlResizeCounter = lvNewVal
            Case ResizeObjectConstant.PictureButtonsObject
                .PictureButtonsResizeCounter = lvNewVal
        End Select
    End With

End Sub

Private Function GetResizeStatus(Index As ResizeObjectConstant) As Integer

    With PV_ResizeControl
        Select Case Index
            Case ResizeObjectConstant.ContainerObject
                GetResizeStatus = .CointainerResizeCounter
            Case ResizeObjectConstant.ControlObject
                GetResizeStatus = .ControlResizeCounter
            Case ResizeObjectConstant.PictureButtonsObject
                GetResizeStatus = .PictureButtonsResizeCounter
        End Select
    End With

End Function

Private Sub ResetResizeStatus()

Dim i               As Integer

    For i = 0 To PV_ResizeControl.Count
        SetResizeStatus i, 0
    End With

End Sub

Private Function IncrementResizeStatu(Index As ResizeObjectConstant)

    SetResizeStatus i, GetResizeStatus(i) + 1
    
End Function

Private Sub UserControl_Initialize()

    PV_pictureHeight = 60
    PV_pictureWidth = 60
    PV_pictureInterval = 60
    PV_LandscapeState = True
    PV_PictureCount = 60
    PV_pictureLeft = 60
    PV_pictureTop = 60
    PV_ResizeState = False
    
    ResetResizeControl PV_ResizeControl
    
    apiSetSizePositionButtons PV_pictureWidth, _
            PV_pictureHeight, _
            PV_pictureLeft, _
            PV_pictureTop, _
            PV_pictureInterval, _
            PV_LandscapeState
    
End Sub

Public Function apiSetSizePositionButtons(lvNewWidth As Integer, _
        lvNewHeight As Integer, _
        lvNewLeft As Integer, _
        lvNewTop As Integer, _
        lvNewInterval As Integer, _
        lvNewOrientation As OrientationConstants, _
        Optional lvRiseRefresh As Boolean = False) As Boolean

    If apiSetSizeButtons(lvNewWidth, lvNewHeight) = False Then
        apiRiseError
    End If
    If apiSetPositionButtons(lvNewLeft, lvNewTop, lvNewInterval, lvNewOrientation) = False Then
        apiRiseError
    End If
    If lvRiseRefresh = True Then
    
    End If
    If apiSetSizeButtons(lvNewWidth, lvNewHeight) = False Then
        apiRiseError
    End If
    
End Function

Public Function apiSetSizeButtons(lvNewWidth As Integer, _
        lvNewHeight As Integer) As Boolean

Dim i       As Integer

    For i = ImageButton.LBound To ImageButton.UBound
        If apiSetSizeButton(lvNewWidth, lvNewHeight, i) = False Then
            apiRiseError
        End If
    Next
    
    apiSetSizeButtons = True
    
End Function

Public Function apiSetSizeButton(lvNewWidth As Integer, _
        lvNewHeight As Integer, Index As Integer) As Boolean

    With ImageButton(Index)
        If .Width <> lvNewWidth Then
            .Width = lvNewWidth
            apiRiseEventChangeImageSize
        End If
        If .Height <> lvNewHeight Then
            .Height = lvNewHeight
            apiRiseEventChangeImageSize
        End If
    End With
    apiSetSizeButton = True
    
End Function

Public Function apiSetPositionButtons(lvNewLeft As Integer, _
        lvNewTop As Integer, _
        lvNewInterval As Integer, _
        lvNewOrientation As OrientationConstants, _
        Optional lvRiseRefresh As Boolean = False) As Boolean

Dim i           As Integer

    If apiSetPositionFirstButton(lvNewLeft, lvNewTop, lvNewOrientation) = False Then
        apiRiseError
    End If
    For i = ImageButton.LBound + 1 To ImageButton.UBound
        If apiMoveButton(lvNewInterval, lvNewOrientation, i) = False Then
            apiRiseError
        End If
    Next

End Function

Public Function apiMoveButton(lvNewInterval As Integer, _
        lvNewOrientation As OrientationConstants, _
        Index As Integer) As Boolean

    If Index > ImageButton.LBound Then
        With ImageButton(Index)
            Select Case lvNewOrientation
                Case Is = OrientationConstants.orientPortrait
                    .Left = ImageButton(Index - 1).Left
                    .Top = ImageButton(Index - 1).Top + _
                            ImageButton(Index - 1).Height + _
                            lvNewInterval
                    
                Case Is = OrientationConstants.orientLandscape
                    .Top = ImageButton(Index - 1).Top
                    .Left = ImageButton(Index - 1).Left + _
                            ImageButton(Index - 1).Width + _
                            lvNewInterval
            End Select
        End With
    Else
        apiMoveButton = False
    End If

End Function

Public Function apiSetPositionFirstButton(lvNewLeft As Integer, _
        lvNewTop As Integer, _
        lvNewOrientation As OrientationConstants) As Boolean

    With ImageButton(ImageButton.LBound)
        Select Case lvNewOrientation
            Case Is = OrientationConstants.orientPortrait
                .Top = lvNewTop
                .Left = lvNewLeft
            Case Is = OrientationConstants.orientLandscape
                .Top = lvNewLeft
                .Left = lvNewTop
        End Select
    End With
    
End Function

Public Function SetSizeContainer _
        (lvNewWidth As Integer, _
        lvNewHeight As Integer) As Boolean

    With PictureBotones
        .Width = lvNewWidth
        .Height = lvNewHeight
    End With
    
End Function

Public Function SetSizeControl _
        (lvNewWidth As Integer, _
        lvNewHeight As Integer) As Boolean



End Function
