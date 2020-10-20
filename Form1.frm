VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2505
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11775
   LinkTopic       =   "Form1"
   ScaleHeight     =   2505
   ScaleWidth      =   11775
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnOk 
      Caption         =   "Confirmar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   2775
   End
   Begin VB.TextBox frmInputBox 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   "Test"
      ToolTipText     =   "Ayuda: Texto"
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Label lblResult 
      Caption         =   "Texto resultante"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4440
      TabIndex        =   2
      Top             =   1080
      Width           =   3975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "This is an example text"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnOk_Click_2()
Dim resp As String
resp = InputBox("Ingresa un texto cuaquiera", "Testing", 0)
If resp = "" Then
lblResult.Caption = "No se ha ingresado nada"
Else
lblResult.Caption = resp
End If
End Sub

Private Sub btnOk_Click()
    Dim resp As String
    resp = frmInputBox.Text
    If resp = "" Then
        lblResult.Caption = "Sin nada"
    Else
        lblResult.Caption = resp
    End If
End Sub
