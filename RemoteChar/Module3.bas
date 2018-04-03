Attribute VB_Name = "Module3"
Public ChatType As Integer
Public msgLength(0 To 11) As Integer
Public ReceivedMSG(0 To 11) As String
Public ReceivedX As Integer
Public ReceivedY As Integer
Public ReceivedChat As String
Public ReceivedADD As String

Public ADR_TYPING As Long
Public ADR_LASTMSG As Long

Function Chat(strType As Integer)
Dim pStr As String
Dim pBytes() As Byte

If strType = 0 Then
pStr = "1001FF00" & hexword
ElseIf strType = 1 Then
pStr = "1005FF00" & hexword
ElseIf strType = 2 Then
pStr = "1003FF00" & hexword
ElseIf strType = 3 Then
pStr = "1006FF00" & hexword
ElseIf strType = 4 Then
pStr = "1004FF00" & hexword
ElseIf strType = 5 Then
pStr = "100EFF00" & hexword
End If

ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes


End Function

Function Action(Index As Integer)

Select Case Index
Case 0
HexString "Ok"
Chat (0)
Form1.TOWN
Debug.Print "#### EXECUTED: TOWN ####"
Case 1
HexString "Ok"
Chat (0)
Form1.TOWN
Form1.OFFKO
Debug.Print "#### EXECUTED: TOWN+OFF KO ####"
Case 2
HexString "Ok"
Chat (0)
Form1.TOWN
Form1.OFFKO
Form1.OFFPC
Debug.Print "#### EXECUTED: TOWN+OFF PC ####"
Case 3
HexString "Ok"
Chat (0)
Form1.RESPAWN
Debug.Print "#### EXECUTED: RESPAWN ####"
Case 4
HexString "Ok"
Chat (0)
Tuþ yolla("C")
Case 5
HexString "I'm going to " & ReceivedX & "," & ReceivedY
Chat (0)
Form1.GOTOXY ReceivedX, ReceivedY
Case 7
HexString "Ok"
Chat (0)
Form1.DISBAND
Case 9
HexString "Ty"
Chat (0)
Form1.JOINPARTY
Case 10
HexString "Ok"
Chat (0)
Form1.ACCEPTTRADE
Sleep (1000)
Form1.APPLYTRADE
Case 11
Dim mystatus As String
Dim percEXP_, percEXP As String
percEXP_ = ReadLong(KO_ADR_CHR + KO_OFF_MAXEXP) / ReadLong(KO_ADR_CHR + KO_OFF_EXP)
percEXP = CInt(100 / percEXP_) & "%" '50%
mystatus = "HP: " & ReadLong(KO_ADR_CHR + KO_OFF_HP) & "\" & ReadLong(KO_ADR_CHR + KO_OFF_MAXHP) & ", MP: " & ReadLong(KO_ADR_CHR + KO_OFF_MP) & "\" & ReadLong(KO_ADR_CHR + KO_OFF_MAXMP) & ", Lvl: " & ReadLong(KO_ADR_CHR + KO_OFF_LVL) & ", Exp: " & percEXP
HexString mystatus
Chat (0)
Case Else
End Select

End Function
