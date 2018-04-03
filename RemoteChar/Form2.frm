VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Chat Type"
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   1440
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   162
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   1440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnOK 
      Caption         =   "O K"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
   End
   Begin VB.OptionButton op4 
      Caption         =   "Clan"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.OptionButton op3 
      Caption         =   "Party"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.OptionButton op2 
      Caption         =   "Shout"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.OptionButton op1 
      Caption         =   "Normal"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Value           =   -1  'True
      Width           =   1215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub btnOK_Click()
If op1.Value = True Then
Form1.lbchattype.Caption = "[ Normal ]"
Form1.lbchattype.ForeColor = vbBlack
ChatType = 1
ElseIf op2.Value = True Then
Form1.lbchattype.Caption = "[ Shout ]"
Form1.lbchattype.ForeColor = &H80FF&
ChatType = 2
ElseIf op3.Value = True Then
Form1.lbchattype.Caption = "[ Party ]"
Form1.lbchattype.ForeColor = &H808000
ChatType = 3
ElseIf op4.Value = True Then
Form1.lbchattype.Caption = "[ Clan ]"
Form1.lbchattype.ForeColor = vbGreen
ChatType = 4
End If
Me.Hide
Debug.Print "Selected Chat Type: " & ChatType
End Sub
