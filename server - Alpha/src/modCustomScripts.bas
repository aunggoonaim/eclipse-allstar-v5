Attribute VB_Name = "modCustomScripts"
' ʤ�Ի Event

Public Sub CustomScript(index As Long, caseID As Long)
Dim i As Long
    
    ' ���ͺ�к����� � ���� � ���..
    Select Case caseID
        Case 1
        ' �����ԡѴ
        Call PlayerWarp(index, GetPlayerMap(index), Map(GetPlayerMap(index)).MaxX, Map(GetPlayerMap(index)).MaxY)
        
        Case 2 ' ����¹�� ������
        Call ClassChange(index, 4, Player(index).Sex)
    
        Case 3 ' ����¹�� ���ҴԹ
        Call ClassChange(index, 5, Player(index).Sex)
        
        Case 4 ' ����¹�� �ԫ���
        Call ClassChange(index, 6, Player(index).Sex)
        
        Case 5 ' ����¹�� ������
        Call ClassChange(index, 7, Player(index).Sex)
        
        Case 6 ' ����¹�� �ѹ����
        Call ClassChange(index, 8, Player(index).Sex)
        
        Case 7 ' ����¹�� ������
        Call ClassChange(index, 9, Player(index).Sex)
        
        Case 8 ' ����¹�� ����ʫԹ
        Call ClassChange(index, 10, Player(index).Sex)
        
        Case 9 ' ����¹�� ��������
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
            PlayerMsg index, "�س�ѧ�����ӡ�����ҧʤ�Ի�����Ţ " & caseID & " ���. ��سҵ�Ǩ�ͺʤ�Ի�ա���駤��.", BrightRed
    End Select
End Sub
