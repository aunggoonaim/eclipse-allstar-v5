Attribute VB_Name = "Scripted"
Sub ScriptedClick(index, script)

' สคริปสกิล และไอเทม

Select Case script

Case 0
' สุ่มพิกัดในแผนที่
Call PlayerWarp(index, GetPlayerMap(index), Map(GetPlayerMap(index)).MaxX, Map(GetPlayerMap(index)).MaxY)

Case 1
'
Call PlayerMsg(index, "สคริป 1 !", White)

Case 2
Call PlayerMsg(index, "สคริป 2 !", White)

Case Else
Call PlayerMsg(index, "ยังไม่ได้สร้างสคริปนี้ไว้.", BrightRed)
End Select

End Sub

Public Sub UseScript(ByVal index As Long, ByVal script As Long, Optional ByVal Target As Long = 0, Optional ByVal BuffTime As Byte = 0)

'scripting

Select Case script

Case 0

'
Call PlayerMsg(index, "ไม่มีข้อมูลนี้.", BrightRed)
    
Case 1

' หายตัว
Player(index).BuffStatus(BUFF_INVISIBLE) = BUFF_INVISIBLE
Player(index).BuffTime(BUFF_INVISIBLE) = BuffTime
SendActionMsg GetPlayerMap(index), "หายตัว !", Yellow, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32) + 16
Call PlayerMsg(index, "คุณได้ติดสถานะหายตัว", BrightGreen)

Case 2

Call PlayerMsg(index, "สคริป 2 !", White)

Case 3

' หายตัว
Player(index).BuffStatus(BUFF_INVISIBLE) = BUFF_INVISIBLE
Player(index).BuffTime(BUFF_INVISIBLE) = BuffTime
SendActionMsg GetPlayerMap(index), "หายตัว !", Yellow, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32) + 16
Call PlayerMsg(index, "คุณได้ติดสถานะหายตัว", BrightGreen)

Case Else
Call PlayerMsg(index, "ยังไม่ได้สร้างสคริปนี้ไว้.", BrightRed)
End Select

End Sub

