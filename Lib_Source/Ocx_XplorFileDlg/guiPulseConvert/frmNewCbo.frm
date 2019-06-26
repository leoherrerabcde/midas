VERSION 5.00
Begin VB.Form frmNewCbo 
   Caption         =   "Ingresar Nombre Configuración"
   ClientHeight    =   1545
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1545
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAccept 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox txtNewConfigName 
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Text            =   "txtNewConfigName"
      Top             =   360
      Width           =   2775
   End
End
Attribute VB_Name = "frmNewCbo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmNewCbo
' Author    : Leo Herrera
' Date      : 16/12/2012
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

Public m_Text       As TextBox

Private Sub cmdAccept_Click()

    Unload Me
    
End Sub

Private Sub Form_Load()

    Set m_Text = Me.txtNewConfigName
    txtNewConfigName = ""
    
End Sub
