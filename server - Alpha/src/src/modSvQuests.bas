Attribute VB_Name = "modSvQuests"
'/////////////////////////////////////////////////////////////////////
'///////////////// QUEST SYSTEM - Developed by Alatar ////////////////
'/////////////////////////////////////////////////////////////////////
Option Explicit
Private Declare Sub ZeroMemory Lib "Kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

'Constants
Public Const MAX_TASKS As Byte = 10
Public Const MAX_QUESTS As Byte = 70

Public Const QUEST_TYPE_GOSLAY As Byte = 1
Public Const QUEST_TYPE_GOGATHER As Byte = 2
Public Const QUEST_TYPE_GOTALK As Byte = 3
Public Const QUEST_TYPE_GOREACH As Byte = 4
Public Const QUEST_TYPE_GOGIVE As Byte = 5
Public Const QUEST_TYPE_GOKILL As Byte = 6
Public Const QUEST_TYPE_GOTRAIN As Byte = 7
Public Const QUEST_TYPE_GOGET As Byte = 8

Public Const QUEST_NOT_STARTED As Byte = 0
Public Const QUEST_STARTED As Byte = 1
Public Const QUEST_COMPLETED As Byte = 2
Public Const QUEST_COMPLETED_BUT As Byte = 3

'Types
Public Quest(1 To MAX_QUESTS) As QuestRec

Public Type PlayerQuestRec
    Status As Long '0=not started, 1=started, 2=completed, 3=completed but repeatable
    ActualTask As Long
    CurrentCount As Long 'Used to handle the Amount property
End Type

Public Type TaskRec
    Order As Long
    NPC As Long
    Item As Long
    Map As Long
    Resource As Long
    Amount As Long
    Speech As String * 200
    TaskLog As String * 100
    QuestEnd As Boolean
End Type

Public Type QuestRec
    Name As String * 30
    QuestLog As String * 100
    TasksCount As Long 'todo
    Repeat As Long
    
    Requirement(1 To 3) As Long '1=level, 2=item, 3=quest
    
    QuestGiveItem As Long
    QuestGiveItemValue As Long
    QuestRemoveItem As Long
    QuestRemoveItemValue As Long
    
    Chat(1 To 3) As String * 200
    
    RewardExp As Long
    RewardItem As Long
    RewardItemAmount As Long
    
    Task(1 To MAX_TASKS) As TaskRec
End Type

' //////////////
' // DATABASE //
' //////////////

Sub SaveQuests()
    Dim i As Long
    For i = 1 To MAX_QUESTS
        Call SaveQuest(i)
    Next
End Sub

Sub SaveQuest(ByVal QuestNum As Long)
    Dim filename As String
    Dim F As Long, i As Long
    filename = App.Path & "\data\quests\quest" & QuestNum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
        Put #F, , Quest(QuestNum).Name
        Put #F, , Quest(QuestNum).QuestLog
        Put #F, , Quest(QuestNum).TasksCount
        Put #F, , Quest(QuestNum).Repeat
        For i = 1 To 3
            Put #F, , Quest(QuestNum).Requirement(i)
        Next
        Put #F, , Quest(QuestNum).QuestGiveItem
        Put #F, , Quest(QuestNum).QuestGiveItemValue
        Put #F, , Quest(QuestNum).QuestRemoveItem
        Put #F, , Quest(QuestNum).QuestRemoveItemValue
        For i = 1 To 3
            Put #F, , Quest(QuestNum).Chat(i)
        Next
        Put #F, , Quest(QuestNum).RewardExp
        Put #F, , Quest(QuestNum).RewardItem
        Put #F, , Quest(QuestNum).RewardItemAmount
        For i = 1 To MAX_TASKS
            Put #F, , Quest(QuestNum).Task(i)
        Next
        
    Close #F
End Sub

Sub LoadQuests()
    Dim filename As String
    Dim i As Integer
    Dim F As Long, n As Long
    Dim sLen As Long
    
    Call CheckQuests

    For i = 1 To MAX_QUESTS
        filename = App.Path & "\data\quests\quest" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
            Get #F, , Quest(i).Name
            Get #F, , Quest(i).QuestLog
            Get #F, , Quest(i).TasksCount
            Get #F, , Quest(i).Repeat
            For n = 1 To 3
                Get #F, , Quest(i).Requirement(n)
            Next
            Get #F, , Quest(i).QuestGiveItem
            Get #F, , Quest(i).QuestGiveItemValue
            Get #F, , Quest(i).QuestRemoveItem
            Get #F, , Quest(i).QuestRemoveItemValue
            For n = 1 To 3
                Get #F, , Quest(i).Chat(n)
            Next
            Get #F, , Quest(i).RewardExp
            Get #F, , Quest(i).RewardItem
            Get #F, , Quest(i).RewardItemAmount
            For n = 1 To MAX_TASKS
                Get #F, , Quest(i).Task(n)
            Next
            
        Close #F
    Next
End Sub

Sub CheckQuests()
    Dim i As Long
    For i = 1 To MAX_QUESTS
        If Not FileExist("\Data\quests\quest" & i & ".dat") Then
            Call SaveQuest(i)
        End If
    Next
End Sub

Sub ClearQuest(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Quest(Index)), LenB(Quest(Index)))
    Quest(Index).Name = vbNullString
    Quest(Index).QuestLog = vbNullString
End Sub

Sub ClearQuests()
    Dim i As Long

    For i = 1 To MAX_QUESTS
        Call ClearQuest(i)
    Next
End Sub

' ////////////////////
' // C&S PROCEDURES //
' ////////////////////

Sub SendQuests(ByVal Index As Long)
    Dim i As Long

    For i = 1 To MAX_QUESTS
        If LenB(Trim$(Quest(i).Name)) > 0 Then
            Call SendUpdateQuestTo(Index, i)
        End If
    Next
End Sub

Sub SendUpdateQuestToAll(ByVal QuestNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim QuestSize As Long
    Dim QuestData() As Byte
    Set Buffer = New clsBuffer
    QuestSize = LenB(Quest(QuestNum))
    ReDim QuestData(QuestSize - 1)
    CopyMemory QuestData(0), ByVal VarPtr(Quest(QuestNum)), QuestSize
    Buffer.WriteLong SUpdateQuest
    Buffer.WriteLong QuestNum
    Buffer.WriteBytes QuestData
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateQuestTo(ByVal Index As Long, ByVal QuestNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim QuestSize As Long
    Dim QuestData() As Byte
    Set Buffer = New clsBuffer
    QuestSize = LenB(Quest(QuestNum))
    ReDim QuestData(QuestSize - 1)
    CopyMemory QuestData(0), ByVal VarPtr(Quest(QuestNum)), QuestSize
    Buffer.WriteLong SUpdateQuest
    Buffer.WriteLong QuestNum
    Buffer.WriteBytes QuestData
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendPlayerQuests(ByVal Index As Long)
    Dim i As Long
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerQuest
        For i = 1 To MAX_QUESTS
            Buffer.WriteLong Player(Index).PlayerQuest(i).Status
            Buffer.WriteLong Player(Index).PlayerQuest(i).ActualTask
            Buffer.WriteLong Player(Index).PlayerQuest(i).CurrentCount
        Next
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing

End Sub

Public Sub SendPlayerQuest(ByVal Index As Long, ByVal QuestNum As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerQuest
    Buffer.WriteLong Player(Index).PlayerQuest(QuestNum).Status
    Buffer.WriteLong Player(Index).PlayerQuest(QuestNum).ActualTask
    Buffer.WriteLong Player(Index).PlayerQuest(QuestNum).CurrentCount
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

'sends a message to the client that is shown on the screen
Public Sub QuestMessage(ByVal Index As Long, ByVal QuestNum As Long, ByVal message As String, ByVal QuestNumForStart As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SQuestMessage
    Buffer.WriteLong QuestNum
    Buffer.WriteString Trim$(message)
    Buffer.WriteLong QuestNumForStart
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
    
End Sub

' ///////////////
' // Functions //
' ///////////////

Public Function CanStartQuest(ByVal Index As Long, ByVal QuestNum As Long) As Boolean
    CanStartQuest = False
    If QuestNum < 1 Or QuestNum > MAX_QUESTS Then Exit Function
    If QuestInProgress(Index, QuestNum) Then Exit Function
    
    'check if now a completed quest can be repeated
    If Player(Index).PlayerQuest(QuestNum).Status = QUEST_COMPLETED Then
        If Quest(QuestNum).Repeat = YES Then
            Player(Index).PlayerQuest(QuestNum).Status = QUEST_COMPLETED_BUT
        End If
    End If
    
    'Check if player has the quest 0 (not started) or 3 (completed but it can be started again)
    If Player(Index).PlayerQuest(QuestNum).Status = QUEST_NOT_STARTED Or Player(Index).PlayerQuest(QuestNum).Status = QUEST_COMPLETED_BUT Then
        'Check if player's level is right
        If Quest(QuestNum).Requirement(1) <= Player(Index).Level Then
            'Check if item is needed
            If Quest(QuestNum).Requirement(2) > 0 And Quest(QuestNum).Requirement(2) <= MAX_ITEMS Then
                If HasItem(Index, Quest(QuestNum).Requirement(2)) = 0 Then
                    PlayerMsg Index, "You need " & Item(Quest(QuestNum).Requirement(2)).Name & " to take this quest!", BrightRed
                    Exit Function
                End If
            End If
            'Check if previous quest is needed
            If Quest(QuestNum).Requirement(3) > 0 And Quest(QuestNum).Requirement(3) <= MAX_QUESTS Then
                If Player(Index).PlayerQuest(Quest(QuestNum).Requirement(3)).Status = QUEST_NOT_STARTED Or Player(Index).PlayerQuest(Quest(QuestNum).Requirement(3)).Status = QUEST_STARTED Then
                    PlayerMsg Index, "You need to complete the " & Trim$(Quest(Quest(QuestNum).Requirement(3)).Name) & " quest in order to take this quest!", BrightRed
                    Exit Function
                End If
            End If
            'Go on :)
            CanStartQuest = True
        Else
            PlayerMsg Index, "You need to be a higher level to take this quest!", BrightRed
        End If
    Else
        PlayerMsg Index, "You can't start that quest again!", BrightRed
    End If
End Function

Public Function CanEndQuest(ByVal Index As Long, QuestNum As Long) As Boolean
    CanEndQuest = False
    If Quest(QuestNum).Task(Player(Index).PlayerQuest(QuestNum).ActualTask).QuestEnd = True Then
        CanEndQuest = True
    End If
End Function

'Tells if the quest is in progress or not
Public Function QuestInProgress(ByVal Index As Long, ByVal QuestNum As Long) As Boolean
    QuestInProgress = False
    If QuestNum < 1 Or QuestNum > MAX_QUESTS Then Exit Function
    
    If Player(Index).PlayerQuest(QuestNum).Status = QUEST_STARTED Then 'Status=1 means started
        QuestInProgress = True
    End If
End Function

Public Function QuestCompleted(ByVal Index As Long, ByVal QuestNum As Long) As Boolean
    QuestCompleted = False
    If QuestNum < 1 Or QuestNum > MAX_QUESTS Then Exit Function
    
    If Player(Index).PlayerQuest(QuestNum).Status = 2 Or Player(Index).PlayerQuest(QuestNum).Status = 3 Then
        QuestCompleted = True
    End If
End Function

'Gets the quest reference num (id) from the quest name (it shall be unique)
Public Function GetQuestNum(ByVal QuestName As String) As Long
    Dim i As Long
    GetQuestNum = 0
    
    For i = 1 To MAX_QUESTS
        If Trim$(Quest(i).Name) = Trim$(QuestName) Then
            GetQuestNum = i
            Exit For
        End If
    Next
End Function

Public Function GetItemNum(ByVal ItemName As String) As Long
    Dim i As Long
    GetItemNum = 0
    
    For i = 1 To MAX_ITEMS
        If Trim$(Item(i).Name) = Trim$(ItemName) Then
            GetItemNum = i
            Exit For
        End If
    Next
End Function

' /////////////////////
' // General Purpose //
' /////////////////////

Public Sub CheckTasks(ByVal Index As Long, ByVal TaskType As Long, ByVal TargetIndex As Long)
    Dim i As Long
    
    For i = 1 To MAX_QUESTS
        If QuestInProgress(Index, i) Then
            Call CheckTask(Index, i, TaskType, TargetIndex)
        End If
    Next
End Sub

Public Sub CheckTask(ByVal Index As Long, ByVal QuestNum As Long, ByVal TaskType As Long, ByVal TargetIndex As Long)
    Dim ActualTask As Long, i As Long
    ActualTask = Player(Index).PlayerQuest(QuestNum).ActualTask
    
    Select Case TaskType
        Case QUEST_TYPE_GOSLAY 'Kill X amount of X npc's.
        
            'is npc's defeated id is the same as the npc i have to kill?
            If TargetIndex = Quest(QuestNum).Task(ActualTask).NPC Then
                'Count +1
                Player(Index).PlayerQuest(QuestNum).CurrentCount = Player(Index).PlayerQuest(QuestNum).CurrentCount + 1
                'show msg
                PlayerMsg Index, "Quest: " + Trim$(Quest(QuestNum).Name) + " - " + Trim$(Player(Index).PlayerQuest(QuestNum).CurrentCount) + "/" + Trim$(Quest(QuestNum).Task(ActualTask).Amount) + " " + Trim$(NPC(TargetIndex).Name) + " killed.", Yellow
                'did i finish the work?
                If Player(Index).PlayerQuest(QuestNum).CurrentCount >= Quest(QuestNum).Task(ActualTask).Amount Then
                    QuestMessage Index, QuestNum, "Task completed", 0
                    'is the quest's end?
                    If CanEndQuest(Index, QuestNum) Then
                        EndQuest Index, QuestNum
                    Else
                        'otherwise continue to the next task
                        Player(Index).PlayerQuest(QuestNum).CurrentCount = 0
                        Player(Index).PlayerQuest(QuestNum).ActualTask = ActualTask + 1
                    End If
                End If
            End If
                        
        Case QUEST_TYPE_GOGATHER 'Gather X amount of X item.
            If TargetIndex = Quest(QuestNum).Task(ActualTask).Item Then
                
                'reset the count first
                Player(Index).PlayerQuest(QuestNum).CurrentCount = 0
                
                'Check inventory for the items
                For i = 1 To MAX_INV
                    If GetPlayerInvItemNum(Index, i) = TargetIndex Then
                        If Item(i).Type = ITEM_TYPE_CURRENCY Then
                            'Currency ToDo, something with: GetPlayerInvItemValue(Index, i)
                        Else
                            'If is the correct item add it to the count
                            Player(Index).PlayerQuest(QuestNum).CurrentCount = Player(Index).PlayerQuest(QuestNum).CurrentCount + 1
                        End If
                    End If
                Next
                
                PlayerMsg Index, "Quest: " + Trim$(Quest(QuestNum).Name) + " - You have " + Trim$(Player(Index).PlayerQuest(QuestNum).CurrentCount) + "/" + Trim$(Quest(QuestNum).Task(ActualTask).Amount) + " " + Trim$(Item(TargetIndex).Name), Yellow
                
                If Player(Index).PlayerQuest(QuestNum).CurrentCount >= Quest(QuestNum).Task(ActualTask).Amount Then
                    QuestMessage Index, QuestNum, "Task completed", 0
                    If CanEndQuest(Index, QuestNum) Then
                        EndQuest Index, QuestNum
                    Else
                        Player(Index).PlayerQuest(QuestNum).CurrentCount = 0
                        Player(Index).PlayerQuest(QuestNum).ActualTask = ActualTask + 1
                    End If
                End If
            End If
            
        Case QUEST_TYPE_GOTALK 'Interact with X npc.
            If TargetIndex = Quest(QuestNum).Task(ActualTask).NPC Then
                QuestMessage Index, QuestNum, Quest(QuestNum).Task(ActualTask).Speech, 0
                If CanEndQuest(Index, QuestNum) Then
                    EndQuest Index, QuestNum
                Else
                    Player(Index).PlayerQuest(QuestNum).ActualTask = ActualTask + 1
                End If
            End If
        
        Case QUEST_TYPE_GOREACH 'Reach X map.
            If TargetIndex = Quest(QuestNum).Task(ActualTask).Map Then
                QuestMessage Index, QuestNum, "Task completed", 0
                If CanEndQuest(Index, QuestNum) Then
                    EndQuest Index, QuestNum
                Else
                    Player(Index).PlayerQuest(QuestNum).ActualTask = ActualTask + 1
                End If
            End If
        
        Case QUEST_TYPE_GOGIVE 'Give X amount of X item to X npc.
            If TargetIndex = Quest(QuestNum).Task(ActualTask).NPC Then
                
                Player(Index).PlayerQuest(QuestNum).CurrentCount = 0
                
                For i = 1 To MAX_INV
                    If GetPlayerInvItemNum(Index, i) = Quest(QuestNum).Task(ActualTask).Item Then
                        If Item(i).Type = ITEM_TYPE_CURRENCY Then
                            'Currency ToDo, something with: GetPlayerInvItemValue(Index, i)
                        Else
                            'If is the correct item add it to the count
                            Player(Index).PlayerQuest(QuestNum).CurrentCount = Player(Index).PlayerQuest(QuestNum).CurrentCount + 1
                        End If
                    End If
                Next
                
                PlayerMsg Index, "Quest: " + Trim$(Quest(QuestNum).Name) + " - You have " + Trim$(Player(Index).PlayerQuest(QuestNum).CurrentCount) + "/" + Trim$(Quest(QuestNum).Task(ActualTask).Amount) + " " + Trim$(Item(TargetIndex).Name), Yellow
                
                If Player(Index).PlayerQuest(QuestNum).CurrentCount >= Quest(QuestNum).Task(ActualTask).Amount Then
                    'if we have enough items, then remove them and finish the task
                    For i = 1 To Quest(QuestNum).Task(ActualTask).Amount
                        TakeInvItem Index, Quest(QuestNum).Task(ActualTask).Item, 1
                        'ToDo stuff with currency
                    Next
                    
                    QuestMessage Index, QuestNum, Quest(QuestNum).Task(ActualTask).Speech, 0
                    
                    If CanEndQuest(Index, QuestNum) Then
                        EndQuest Index, QuestNum
                    Else
                        Player(Index).PlayerQuest(QuestNum).CurrentCount = 0
                        Player(Index).PlayerQuest(QuestNum).ActualTask = ActualTask + 1
                    End If
                End If
            End If
                    
        Case QUEST_TYPE_GOKILL 'Kill X amount of players.
            Player(Index).PlayerQuest(QuestNum).CurrentCount = Player(Index).PlayerQuest(QuestNum).CurrentCount + 1
            PlayerMsg Index, "Quest: " + Trim$(Quest(QuestNum).Name) + " - " + Trim$(Player(Index).PlayerQuest(QuestNum).CurrentCount) + "/" + Trim$(Quest(QuestNum).Task(ActualTask).Amount) + " players killed.", Yellow
            If Player(Index).PlayerQuest(QuestNum).CurrentCount >= Quest(QuestNum).Task(ActualTask).Amount Then
                QuestMessage Index, QuestNum, "Task completed", 0
                If CanEndQuest(Index, QuestNum) Then
                    EndQuest Index, QuestNum
                Else
                    Player(Index).PlayerQuest(QuestNum).CurrentCount = 0
                    Player(Index).PlayerQuest(QuestNum).ActualTask = ActualTask + 1
                End If
            End If
            
        Case QUEST_TYPE_GOTRAIN 'Hit X amount of times X resource.
            If TargetIndex = Quest(QuestNum).Task(ActualTask).Resource Then
                Player(Index).PlayerQuest(QuestNum).CurrentCount = Player(Index).PlayerQuest(QuestNum).CurrentCount + 1
                PlayerMsg Index, "Quest: " + Trim$(Quest(QuestNum).Name) + " - " + Trim$(Player(Index).PlayerQuest(QuestNum).CurrentCount) + "/" + Trim$(Quest(QuestNum).Task(ActualTask).Amount) + " hits.", Yellow
                If Player(Index).PlayerQuest(QuestNum).CurrentCount >= Quest(QuestNum).Task(ActualTask).Amount Then
                    QuestMessage Index, QuestNum, "Task completed", 0
                    If CanEndQuest(Index, QuestNum) Then
                        EndQuest Index, QuestNum
                    Else
                        Player(Index).PlayerQuest(QuestNum).CurrentCount = 0
                        Player(Index).PlayerQuest(QuestNum).ActualTask = ActualTask + 1
                    End If
                End If
            End If
                      
        Case QUEST_TYPE_GOGET 'Get X amount of X item from X npc.
            If TargetIndex = Quest(QuestNum).Task(ActualTask).NPC Then
                'ToDo, stuff with currency
                GiveInvItem Index, Quest(QuestNum).Task(ActualTask).Item, Quest(QuestNum).Task(ActualTask).Amount
                QuestMessage Index, QuestNum, Quest(QuestNum).Task(ActualTask).Speech, 0
                If CanEndQuest(Index, QuestNum) Then
                    EndQuest Index, QuestNum
                Else
                    Player(Index).PlayerQuest(QuestNum).CurrentCount = 0
                    Player(Index).PlayerQuest(QuestNum).ActualTask = ActualTask + 1
                End If
            End If
        
    End Select
    SavePlayer Index
    SendPlayerData Index
    SendPlayerQuests Index
End Sub

Public Sub EndQuest(ByVal Index As Long, ByVal QuestNum As Long)
    Dim i As Long
    
    'Check if quest is repeatable, set it as completed
    If Quest(QuestNum).Repeat = YES Then
        Player(Index).PlayerQuest(QuestNum).Status = QUEST_COMPLETED_BUT
        PlayerMsg Index, Trim$(Quest(QuestNum).Name) & ": " + Trim$(Quest(QuestNum).Chat(3)), Green
    Else
        Player(Index).PlayerQuest(QuestNum).Status = QUEST_COMPLETED
    End If
    
    'reset counters to 0
    Player(Index).PlayerQuest(QuestNum).ActualTask = 0
    Player(Index).PlayerQuest(QuestNum).CurrentCount = 0
    
    'give experience
    
    SetPlayerExp Index, GetPlayerExp(Index) + Quest(QuestNum).RewardExp
    
    'remove items on the end
    If Quest(QuestNum).QuestRemoveItem > 0 And Quest(QuestNum).QuestRemoveItem < MAX_ITEMS Then
        If Quest(QuestNum).QuestRemoveItemValue > 0 And Quest(QuestNum).QuestRemoveItemValue < MAX_INV Then 'ToDo: stuff with currency
            For i = 1 To MAX_INV
                If HasItem(Index, Quest(QuestNum).QuestRemoveItem) Then
                    TakeInvItem Index, Quest(QuestNum).QuestRemoveItem, Quest(QuestNum).QuestRemoveItemValue 'todo: currency stuff...
                End If
            Next
        End If
    End If
    
    'give rewards
    GiveInvItem Index, Quest(QuestNum).RewardItem, Quest(QuestNum).RewardItemAmount
    
    'show ending message
    QuestMessage Index, QuestNum, Trim$(Quest(QuestNum).Chat(3)), 0
    
    PlayerMsg Index, Trim$(Quest(QuestNum).Name) & ": quest completed", Green
    
    SavePlayer Index
    SendEXP Index
    Call SendStats(Index)
    SendPlayerData Index
    SendPlayerQuests Index
End Sub

