VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "FIND POINTER"
   ClientHeight    =   825
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   1815
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   162
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   825
   ScaleWidth      =   1815
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "O K"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   1575
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      Max             =   10
      TabIndex        =   0
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&&H423213"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   840
      TabIndex        =   2
      Top             =   15
      Width           =   870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "0"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   600
      TabIndex        =   1
      Top             =   15
      Width           =   105
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Hide
Form1.Show
End Sub

Private Sub Form_Load()
LoadSNDARRAY
End Sub

Private Sub HScroll1_Change()
Select Case HScroll1.Value
Case 0
KO_SND_FNC = SndFnc(1)
Label2.Caption = "&&H" & Hex(SndFnc(1))
Notice (Label2.Caption)
Case 1
KO_SND_FNC = SndFnc(2)
Label2.Caption = "&&H" & Hex(SndFnc(2))
Notice (Label2.Caption)
Case 2
KO_SND_FNC = SndFnc(3)
Label2.Caption = "&&H" & Hex(SndFnc(3))
Notice (Label2.Caption)
Case 3
KO_SND_FNC = SndFnc(4)
Label2.Caption = "&&H" & Hex(SndFnc(4))
Notice (Label2.Caption)
Case 4
KO_SND_FNC = SndFnc(5)
Label2.Caption = "&&H" & Hex(SndFnc(5))
Notice (Label2.Caption)
Case 5
KO_SND_FNC = SndFnc(6)
Label2.Caption = "&&H" & Hex(SndFnc(6))
Notice (Label2.Caption)
Case 6
KO_SND_FNC = SndFnc(7)
Label2.Caption = "&&H" & Hex(SndFnc(7))
Notice (Label2.Caption)
Case 7
KO_SND_FNC = SndFnc(8)
Label2.Caption = "&&H" & Hex(SndFnc(8))
Notice (Label2.Caption)
Case 8
KO_SND_FNC = SndFnc(9)
Label2.Caption = "&&H" & Hex(SndFnc(9))
Notice (Label2.Caption)
Case 9
KO_SND_FNC = SndFnc(10)
Label2.Caption = "&&H" & Hex(SndFnc(10))
Notice (Label2.Caption)
Case Else
End Select
Label1.Caption = HScroll1.Value
End Sub

Public Sub LoadSNDARRAY()
SndFnc(1) = &H474780
SndFnc(2) = &H474A20
SndFnc(3) = &H474CB0
SndFnc(4) = &H474F40
SndFnc(5) = &H4751D0
SndFnc(6) = &H475460
SndFnc(7) = &H4756F0
SndFnc(8) = &H475980
SndFnc(9) = &H475C10
SndFnc(10) = &H475EA0
End Sub
