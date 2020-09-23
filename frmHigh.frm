VERSION 5.00
Begin VB.Form frmHigh 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5250
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   2970
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmHigh.frx":0000
   ScaleHeight     =   5250
   ScaleWidth      =   2970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label lblScore 
      BackStyle       =   0  'Transparent
      Height          =   4515
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   2475
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "HighScores"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   0
      Width           =   2895
   End
   Begin VB.Shape Shape1 
      Height          =   4635
      Left            =   180
      Top             =   300
      Width           =   2595
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Click here to close"
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   4980
      Width           =   2865
   End
End
Attribute VB_Name = "frmHigh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    For A = 1 To 10
        temp = temp & A & ": " & HighS(A).PlName & vbNewLine _
         & "Score: " & HighS(A).plScore & "   Date: " & HighS(A).plDate & vbNewLine '& vbNewLine
    Next A
    lblScore = temp
End Sub

Private Sub Label1_Click()
    If DontClose Then SaveScore: End
    Me.Hide
    MainPause = False
End Sub

