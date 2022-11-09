Attribute VB_Name = "modBosses"
Public Sub BossLogic(ByVal Target As Long, ByVal BossNum As Integer, ByVal npcNum As Long)
    Select Case BossNum
        Case 1
            'Your first boss call will be here.
            Call TestBossLogic(Target, 1)
        Case 2
            'Your first boss call will be here.

        Case 3
            'Your first boss call will be here.

        Case 4
            'Your first boss call will be here.

        Case 5
            'Your first boss call will be here.

        Case Else
            Call PlayerMsg(Target, "��ʹ���Դ��ͼԴ��Ҵ �ô�����ٻ������ѹ���ѧ����ҹ.", Red)
    End Select
End Sub

Private Sub TestBossLogic(ByVal Target As Long, ByVal npcNum As Long)
    'resume normally attacking the player
    Call TryNpcAttackPlayer(3, Target)
    Call Stun(Target)
End Sub

'Stuns the whole party for three seconds
Private Sub Stun(ByVal Target As Long)
Dim i As Long
    'Check if the player is in a party
    If TempPlayer(Target).inParty > 0 Then
        For i = 1 To TempPlayer(Target).inParty
            'Stun the player
            TempPlayer(i).StunDuration = 3 'seconds
            TempPlayer(i).StunTimer = GetTickCount
            'Show animation
            Call SendAnimation(mapNum, 3, GetPlayerX(i), GetPlayerY(i))
        Next
        
        Call PartyMsg(TempPlayer(Target).inParty, "������١ʵ�� 3 �Թҷ� !", Red)
    Else
        TempPlayer(Target).StunDuration = 3
        TempPlayer(Target).StunTimer = GetTickCount
        Call SendAnimation(mapNum, 3, GetPlayerX(Target), GetPlayerY(Target))
        Call PlayerMsg(Target, "�س�ءʵ�� 3 �Թҷ� !", Red)
    End If
End Sub
