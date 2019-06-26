VERSION 5.00
Begin VB.UserControl UserControlExplorer 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.PictureBox pictureFrame 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   240
      ScaleHeight     =   1455
      ScaleWidth      =   1455
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "UserControlExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Module    : UserControlExplorer
' Author    : Leo Herrera
' Date      : 19/11/2013
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit


Private Sub UserControl_Initialize()

    dockForm Form1.hwnd, UserControl.pictureFrame
    
End Sub

Private Sub UserControl_Resize()

    With UserControl
        .pictureFrame.Left = 0
        .pictureFrame.Top = 0
        .pictureFrame.Width = .ScaleWidth
        .pictureFrame.Height = .ScaleHeight
    End With
    
End Sub
