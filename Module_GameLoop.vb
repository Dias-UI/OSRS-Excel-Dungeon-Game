Option Explicit

' ============================================================================
' MODULE_GAMELOOP - Main game tick system and input handling
' ============================================================================

Public LastTickTime As Double
Public TickCount As Integer

' ============================================================================
' GAME START/STOP
' ============================================================================

Sub StartGame()
    ' Initialize everything for new game
    ' Note: Avoid ScreenUpdating = False to prevent flicker
    
    Call ResetGameState
    Call InitializePlayer
    
    ' Initialize visited rooms tracking (ensure it's ready)
    If VisitedDungeonRooms Is Nothing Then
        Set VisitedDungeonRooms = CreateObject("Scripting.Dictionary")
    End If
    
    ' Generate dungeon
    Call GenerateDungeon("large")
    
    ' Place player at start room
    CurrentDungeonRow = StartRoom(0)
    CurrentDungeonCol = StartRoom(1)
    PlayerRoomRow = 10  ' Center of room
    PlayerRoomCol = 10
    
    ' Load the starting room (this will mark it as visited)
    Call LoadCurrentRoom
    
    ' Ensure player sprite is drawn at start
    Call DrawPlayer
    
    ' Start the game tick
    GameRunning = True
    TickCount = 0
    LastTickTime = Timer
    
    ' Start tick loop
    Call GameTick
    
    Debug.Print "Game Started - Dungeon Size: " & DungeonSize & "x" & DungeonSize
    Debug.Print "Player at Dungeon Room: (" & CurrentDungeonRow & "," & CurrentDungeonCol & ")"
End Sub

Sub StopGame()
    ' Stop the game loop
    GameRunning = False
    Debug.Print "Game Stopped - Total Ticks: " & TickCount
End Sub

' ============================================================================
' MAIN GAME TICK (runs every 0.6 seconds)
' ============================================================================

Sub GameTick()
    If Not GameRunning Then Exit Sub
    
    ' Increment tick counter
    TickCount = TickCount + 1
    GameTicks = GameTicks + 1
    
    ' Note: Movement is now handled immediately on click, not in ticks
    ' This tick is just for combat, prayer drain, and other time-based events
    
    ' Process combat if active
    If InCombat Then
        Call ProcessCombatTick
    End If
    
    ' Process prayer drain
    Call ProcessPrayerDrain
    
    ' Update display
    Call UpdateStatsPanel
    Call UpdateMinimap
    
    ' Debug tick timing
    If TickCount Mod 10 = 0 Then
        Debug.Print "Tick " & TickCount & " - Time: " & Format(Timer - LastTickTime, "0.00") & "s"
        LastTickTime = Timer
    End If
    
    ' Schedule next tick (0.6 seconds)
    Application.OnTime Now + TimeSerial(0, 0, 0.6), "Module_GameLoop.GameTick"
End Sub

' ============================================================================
' MOVEMENT PROCESSING
' ============================================================================

Sub ProcessMovement()
    ' Move player 1 tile per half-tick with diagonal movement
    ' Movement priority: Diagonal > Horizontal > Vertical
    
    Dim nextRow As Integer, nextCol As Integer
    nextRow = PlayerRoomRow
    nextCol = PlayerRoomCol
    
    ' Calculate deltas
    Dim deltaRow As Integer, deltaCol As Integer
    deltaRow = TargetRow - PlayerRoomRow
    deltaCol = TargetCol - PlayerRoomCol
    
    ' Check if reached target
    If deltaRow = 0 And deltaCol = 0 Then
        IsMoving = False
        Exit Sub
    End If
    
    ' Move ONE tile with diagonal priority
    ' Try diagonal first, then horizontal, then vertical
    If deltaRow <> 0 And deltaCol <> 0 Then
        ' Try diagonal movement
        If deltaRow > 0 Then nextRow = nextRow + 1 Else nextRow = nextRow - 1
        If deltaCol > 0 Then nextCol = nextCol + 1 Else nextCol = nextCol - 1
    ElseIf deltaCol <> 0 Then
        ' Move horizontally only
        If deltaCol > 0 Then nextCol = nextCol + 1 Else nextCol = nextCol - 1
    ElseIf deltaRow <> 0 Then
        ' Move vertically only
        If deltaRow > 0 Then nextRow = nextRow + 1 Else nextRow = nextRow - 1
    End If
    
    ' Check if move is valid
    If CanMoveTo(nextRow, nextCol) Then
        PlayerRoomRow = nextRow
        PlayerRoomCol = nextCol
        
        ' Update display immediately after each tile movement for smooth animation
        Call DrawPlayer
        DoEvents  ' Allow screen to refresh
        
        ' Check for room transitions (doors) - this will stop movement
        Call CheckRoomTransition
        
        ' If we transitioned rooms, stop processing more moves
        If Not IsMoving Then Exit Sub
        
        ' Check for enemy encounters
        Call CheckEnemyEncounter
    Else
        ' Path blocked - try non-diagonal movement as fallback
        If deltaRow <> 0 And deltaCol <> 0 Then
            ' Diagonal was blocked, try horizontal or vertical
            nextRow = PlayerRoomRow
            nextCol = PlayerRoomCol
            
            ' Try horizontal first
            If deltaCol <> 0 Then
                If deltaCol > 0 Then nextCol = nextCol + 1 Else nextCol = nextCol - 1
                If CanMoveTo(nextRow, nextCol) Then
                    PlayerRoomRow = nextRow
                    PlayerRoomCol = nextCol
                    Call DrawPlayer
                    DoEvents
                    Call CheckRoomTransition
                    If IsMoving Then Call CheckEnemyEncounter
                    Exit Sub
                End If
            End If
            
            ' Try vertical
            nextRow = PlayerRoomRow
            nextCol = PlayerRoomCol
            If deltaRow <> 0 Then
                If deltaRow > 0 Then nextRow = nextRow + 1 Else nextRow = nextRow - 1
                If CanMoveTo(nextRow, nextCol) Then
                    PlayerRoomRow = nextRow
                    PlayerRoomCol = nextCol
                    Call DrawPlayer
                    DoEvents
                    Call CheckRoomTransition
                    If IsMoving Then Call CheckEnemyEncounter
                    Exit Sub
                End If
            End If
        End If
        
        ' All paths blocked
        IsMoving = False
        PathBlocked = True
        Debug.Print "Movement blocked - no valid path"
    End If
End Sub

' ============================================================================
' IMMEDIATE CONTINUOUS MOVEMENT
' ============================================================================

Sub MoveToDestination()
    ' Move player continuously until reaching destination
    ' This runs immediately when clicked, not waiting for ticks
    
    Dim maxMoves As Integer
    maxMoves = 200  ' Safety limit to prevent infinite loops
    Dim moveCount As Integer
    moveCount = 0
    
    Dim startTime As Double
    Dim elapsed As Double
    
    ' Keep moving until we reach destination or get blocked
    Do While IsMoving And moveCount < maxMoves
        Call ProcessMovement
        moveCount = moveCount + 1
        
        ' Small delay for visual smoothness using busy wait
        startTime = Timer
        Do
            DoEvents  ' Allow screen updates
            elapsed = Timer - startTime
            If elapsed < 0 Then elapsed = elapsed + 86400  ' Handle midnight rollover
        Loop While elapsed < 0.03  ' 30ms delay per tile for smooth movement
    Loop
    
    If moveCount >= maxMoves Then
        Debug.Print "Movement stopped - max moves reached"
        IsMoving = False
    End If
    
    Debug.Print "Moved " & moveCount & " tiles to destination"
End Sub

Function CanMoveTo(row As Integer, col As Integer) As Boolean
    ' Check if player can move to this cell
    CanMoveTo = False
    
    ' Check bounds
    If row < 1 Or row > 20 Or col < 1 Or col > 20 Then
        Exit Function
    End If
    
    ' Check for walls
    If CurrentRoomGrid(row, col) = 1 Then
        Exit Function
    End If
    
    ' Check for enemies (can't walk through them)
    If CurrentRoomGrid(row, col) = 2 Then
        Exit Function
    End If
    
    CanMoveTo = True
End Function

Sub CheckRoomTransition()
    ' Check if player is on a door and should transition to next room
    Dim cellType As Integer
    cellType = CurrentRoomGrid(PlayerRoomRow, PlayerRoomCol)
    
    Dim newDungeonRow As Integer, newDungeonCol As Integer
    newDungeonRow = CurrentDungeonRow
    newDungeonCol = CurrentDungeonCol
    
    Select Case cellType
        Case 4  ' North door
            newDungeonRow = newDungeonRow - 1
            PlayerRoomRow = 17  ' Enter from south side (not at edge)
        Case 5  ' East door
            newDungeonCol = newDungeonCol + 1
            PlayerRoomCol = 3   ' Enter from west side (not at edge)
        Case 6  ' South door
            newDungeonRow = newDungeonRow + 1
            PlayerRoomRow = 3   ' Enter from north side (not at edge)
        Case 7  ' West door
            newDungeonCol = newDungeonCol - 1
            PlayerRoomCol = 17  ' Enter from east side (not at edge)
        Case Else
            Exit Sub  ' Not on a door
    End Select
    
    ' Validate new room exists
    If RoomExists(newDungeonRow, newDungeonCol) Then
        CurrentDungeonRow = newDungeonRow
        CurrentDungeonCol = newDungeonCol
        
        ' STOP MOVEMENT when entering new room
        IsMoving = False
        
        Call LoadCurrentRoom
        
        ' CRITICAL: Redraw player sprite after loading new room
        Call DrawPlayer
        
        Debug.Print "Entered room (" & CurrentDungeonRow & "," & CurrentDungeonCol & ") - movement stopped"
    End If
End Sub

Sub CheckEnemyEncounter()
    ' Check if player walked onto enemy cell
    If CurrentRoomGrid(PlayerRoomRow, PlayerRoomCol) = 2 Then
        ' Start combat
        Debug.Print "Enemy encountered at (" & PlayerRoomRow & "," & PlayerRoomCol & ")"
        ' Combat will be implemented in Module_Combat
    End If
End Sub

' ============================================================================
' CLICK MOVEMENT HANDLER
' ============================================================================

Public Sub OnGameAreaClick(clickedCell As Range)
    ' Called when player clicks in gameplay area (K4:AK24)
    ' Maps clicked cell to room coordinates and moves player IMMEDIATELY
    
    If InCombat Then
        Debug.Print "Cannot move during combat"
        Exit Sub
    End If
    
    ' Calculate room coordinates from clicked cell
    ' K4 starts at (4, 11), room is offset by 4 columns to center
    ' Room cells are at columns 15-33 (K=11, +4 offset = 15)
    Dim clickRow As Integer, clickCol As Integer
    clickRow = clickedCell.row - 4  ' K4 starts at row 4, room row 1 is at row 5
    clickCol = clickedCell.Column - 14  ' Column 15 (O) = room col 1, so subtract 14
    
    ' Validate click is within room (19x19)
    If clickRow < 1 Or clickRow > 19 Or clickCol < 1 Or clickCol > 19 Then
        Debug.Print "Click outside room bounds: (" & clickRow & "," & clickCol & ")"
        Exit Sub
    End If
    
    ' Can't click on walls
    If CurrentRoomGrid(clickRow, clickCol) = 1 Then
        Debug.Print "Cannot move to wall at (" & clickRow & "," & clickCol & ")"
        Exit Sub
    End If
    
    ' Set movement target
    Module_GameState.TargetRow = clickRow
    Module_GameState.TargetCol = clickCol
    IsMoving = True
    PathBlocked = False
    
    Debug.Print "Moving to (" & clickRow & "," & clickCol & ") from (" & PlayerRoomRow & "," & PlayerRoomCol & ")"
    
    ' Move IMMEDIATELY and CONTINUOUSLY until destination reached
    Call MoveToDestination
End Sub

' ============================================================================
' PLACEHOLDER FUNCTIONS (will be implemented in other modules)
' ============================================================================

Sub ProcessCombatTick()
    ' Placeholder - will be implemented in Module_Combat
End Sub

Sub ProcessPrayerDrain()
    ' Drain prayer points based on active prayers
    Dim activePrayerCount As Integer
    activePrayerCount = 0
    
    Dim i As Integer
    For i = 1 To 30
        If PrayerActive(i) Then activePrayerCount = activePrayerCount + 1
    Next i
    
    If activePrayerCount > 0 Then
        PrayerDrainRate = activePrayerCount * 0.1  ' 0.1 per prayer per tick
        PlayerPrayer = PlayerPrayer - PrayerDrainRate
        
        If PlayerPrayer <= 0 Then
            PlayerPrayer = 0
            ' Turn off all prayers
            For i = 1 To 30
                PrayerActive(i) = False
            Next i
        End If
    End If
End Sub

' LoadCurrentRoom and UpdateDisplay are implemented in Module_Renderer
