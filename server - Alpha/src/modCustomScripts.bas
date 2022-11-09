Attribute VB_Name = "modCustomScripts"
' สคริป Event

Public Sub CustomScript(index As Long, caseID As Long)
Dim i As Long
    
    ' ทดสอบระบบใหม่ ๆ เร็ว ๆ นี้..
    Select Case caseID
        Case 1
        ' สุ่มพิกัด
        Call PlayerWarp(index, GetPlayerMap(index), Map(GetPlayerMap(index)).MaxX, Map(GetPlayerMap(index)).MaxY)
        
        Case 2 ' เปลี่ยนเป็น เบอเซิร์ก
        Call ClassChange(index, 4, Player(index).Sex)
    
        Case 3 ' เปลี่ยนเป็น พาลาดิน
        Call ClassChange(index, 5, Player(index).Sex)
        
        Case 4 ' เปลี่ยนเป็น วิซาร์ด
        Call ClassChange(index, 6, Player(index).Sex)
        
        Case 5 ' เปลี่ยนเป็น ซามูไร
        Call ClassChange(index, 7, Player(index).Sex)
        
        Case 6 ' เปลี่ยนเป็น ฮันเตอร์
        Call ClassChange(index, 8, Player(index).Sex)
        
        Case 7 ' เปลี่ยนเป็น สไนเปอร์
        Call ClassChange(index, 9, Player(index).Sex)
        
        Case 8 ' เปลี่ยนเป็น แอสแซสซิน
        Call ClassChange(index, 10, Player(index).Sex)
        
        Case 9 ' เปลี่ยนเป็น ดาร์คลอร์ด
        Call ClassChange(index, 11, Player(index).Sex)
        
        Case 10 ' Exp skill quest
        ' check spell level up ?
        For i = 1 To MAX_PLAYER_SPELLS
            If Player(index).Spell(i) > 0 Then
                If Player(index).skillLV(i) < MAX_SKILL_LEVEL Then
                    Player(index).skillEXP(i) = Player(index).skillEXP(i) + 100
                    Call CheckPlayerSkillUp(index, i)
                    SendPlayerData index
                Else
                    Player(index).skillLV(i) = MAX_SKILL_LEVEL - 1
                    SendPlayerData index
                End If
            End If
        Next
        
        SendActionMsg GetPlayerMap(index), "+ 100", BrightGreen, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
    
        Case Else
            PlayerMsg index, "คุณยังไม่ได้ทำการสร้างสคริปหมายเลข " & caseID & " ไว้. กรุณาตรวจสอบสคริปอีกครั้งค่ะ.", BrightRed
    End Select
End Sub
