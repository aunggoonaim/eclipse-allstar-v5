Attribute VB_Name = "Scripted"
Sub ScriptedClick(index, script)

' ʤ�Իʡ�� �������

Select Case script

Case 0
' �����ԡѴ�Ἱ���
Call PlayerWarp(index, GetPlayerMap(index), Map(GetPlayerMap(index)).MaxX, Map(GetPlayerMap(index)).MaxY)

Case 1
'
Call PlayerMsg(index, "ʤ�Ի 1 !", White)

Case 2
Call PlayerMsg(index, "ʤ�Ի 2 !", White)

Case Else
Call PlayerMsg(index, "�ѧ��������ҧʤ�Ի������.", BrightRed)
End Select

End Sub

Public Sub UseScript(ByVal index As Long, ByVal script As Long, Optional ByVal Target As Long = 0, Optional ByVal BuffTime As Byte = 0)

'scripting

Select Case script

Case 0

'
Call PlayerMsg(index, "����բ����Ź��.", BrightRed)
    
Case 1

' ��µ��
Player(index).BuffStatus(BUFF_INVISIBLE) = BUFF_INVISIBLE
Player(index).BuffTime(BUFF_INVISIBLE) = BuffTime
SendActionMsg GetPlayerMap(index), "��µ�� !", Yellow, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32) + 16
Call PlayerMsg(index, "�س��Դʶҹ���µ��", BrightGreen)

Case 2

Call PlayerMsg(index, "ʤ�Ի 2 !", White)

Case 3

' ��µ��
Player(index).BuffStatus(BUFF_INVISIBLE) = BUFF_INVISIBLE
Player(index).BuffTime(BUFF_INVISIBLE) = BuffTime
SendActionMsg GetPlayerMap(index), "��µ�� !", Yellow, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32) + 16
Call PlayerMsg(index, "�س��Դʶҹ���µ��", BrightGreen)

Case Else
Call PlayerMsg(index, "�ѧ��������ҧʤ�Ի������.", BrightRed)
End Select

End Sub

