VERSION 5.00
Begin VB.Form frmLinkNewProject 
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Line Line 
      X1              =   120
      X2              =   3120
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label lbl 
      Caption         =   "Proyecto Nuevo"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
End
Attribute VB_Name = "frmLinkNewProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmLinkNewProject
' Author    : Leo Herrera
' Date      : 16/07/2013
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit


Private Sub Form_Load()

    CopyBackGroundColor Me
    
End Sub

Private Sub Form_Resize()

    Form_Move_Controls_To_Center Me
    
End Sub
