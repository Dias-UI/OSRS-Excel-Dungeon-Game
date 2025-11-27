Option Explicit

' ============================================================================
' MODULE_RENDERER - Drawing rooms, UI updates, visual effects
' ============================================================================

' ============================================================================
' ROOM LOADING AND GENERATION
' ============================================================================

Public Sub LoadCurrentRoom()
    ' Load and display the current dungeon room
    
    On Error GoTo ErrorHandler
    
    Debug.Print "=== LoadCurrentRoom START ==="
    Debug.Print "Dungeon position: (" & CurrentDungeonRow & "," & CurrentDungeonCol & ")"
    
    ' Mark this room as visited
    Dim roomKey As String
    roomKey = CurrentDungeonRow & "," & CurrentDungeonCol
    If Not VisitedDungeonRooms.Exists(roomKey) Then
        VisitedDungeonRooms.Add roomKey, True
        Debug.Print "Room marked as visited: " & roomKey
    End If
    
    ' Get door configuration for this room
    Dim doors As Variant
    doors = GetRoomDoors(CurrentDungeonRow, CurrentDungeonCol)
    
    Debug.Print "Doors - N:" & doors(1) & " E:" & doors(2) & " S:" & doors(3) & " W:" & doors(4)
    
    ' Generate room content
    Call GenerateRoomContent(doors)
    
    ' Verify generation
    Debug.Print "Room grid populated - checking:"
    Debug.Print "  Corner (1,1): " & CurrentRoomGrid(1, 1)
    Debug.Print "  Center (10,10): " & CurrentRoomGrid(10, 10)
    Debug.Print "  Corner (19,19): " & CurrentRoomGrid(19, 19)
    
    ' Draw the room
    Debug.Print "Calling DrawRoom..."
    Call DrawRoom
    Debug.Print "DrawRoom complete"
    
    ' Update minimap
    Call UpdateMinimap
    
    Debug.Print "=== LoadCurrentRoom END ==="
    Exit Sub
    
ErrorHandler:
    Debug.Print "ERROR in LoadCurrentRoom: " & Err.Description
End Sub

' TEST FUNCTION - Run this to verify renderer works
Public Sub TestRenderer()
    ' Manual test of room rendering
    Debug.Print "=== TESTING RENDERER ==="
    
    ' Set up test data
    CurrentDungeonRow = 1
    CurrentDungeonCol = 1
    PlayerRoomRow = 10
    PlayerRoomCol = 10
    
    ' Create simple test room
    Dim i As Integer, j As Integer
    For i = 1 To 19
        For j = 1 To 19
            If i = 1 Or i = 19 Or j = 1 Or j = 19 Then
                CurrentRoomGrid(i, j) = 1  ' Walls
            Else
                CurrentRoomGrid(i, j) = 0  ' Floor
            End If
        Next j
    Next i
    
    ' Add test doors
    CurrentRoomGrid(1, 10) = 4   ' North door
    CurrentRoomGrid(10, 19) = 5  ' East door
    
    Debug.Print "Test room created, calling DrawRoom..."
    Call DrawRoom
    Debug.Print "=== TEST COMPLETE - Check O5:AG23 ==="
End Sub

Sub GenerateRoomContent(doors As Variant)
    ' Generate 19x19 room grid with walls, floor, doors, enemies, chests
    ' 0 = floor, 1 = wall, 2 = enemy, 3 = chest,
    ' 4 = door_north, 5 = door_east, 6 = door_south, 7 = door_west
    
    Dim row As Integer, col As Integer
    
    ' Fill entire room with floor
    For row = 1 To 19
        For col = 1 To 19
            CurrentRoomGrid(row, col) = 0
        Next col
    Next row
    
    ' Add walls around perimeter
    For col = 1 To 19
        CurrentRoomGrid(1, col) = 1  ' North wall
        CurrentRoomGrid(19, col) = 1  ' South wall
    Next col
    For row = 1 To 19
        CurrentRoomGrid(row, 1) = 1  ' West wall
        CurrentRoomGrid(row, 19) = 1  ' East wall
    Next row
    
    ' Add doors where specified (3-tile wide openings)
    If doors(1) Then  ' North door
        CurrentRoomGrid(1, 9) = 4
        CurrentRoomGrid(1, 10) = 4
        CurrentRoomGrid(1, 11) = 4
    End If
    If doors(2) Then  ' East door
        CurrentRoomGrid(9, 19) = 5
        CurrentRoomGrid(10, 19) = 5
        CurrentRoomGrid(11, 19) = 5
    End If
    If doors(3) Then  ' South door
        CurrentRoomGrid(19, 9) = 6
        CurrentRoomGrid(19, 10) = 6
        CurrentRoomGrid(19, 11) = 6
    End If
    If doors(4) Then  ' West door
        CurrentRoomGrid(9, 1) = 7
        CurrentRoomGrid(10, 1) = 7
        CurrentRoomGrid(11, 1) = 7
    End If
    
    ' Add random interior walls for variety (except in start/end rooms)
    If Not (CurrentDungeonRow = StartRoom(0) And CurrentDungeonCol = StartRoom(1)) And _
       Not (CurrentDungeonRow = EndRoom(0) And CurrentDungeonCol = EndRoom(1)) Then
        Call AddInteriorWalls
    End If
    
    ' Add enemies (except in start room)
    If Not (CurrentDungeonRow = StartRoom(0) And CurrentDungeonCol = StartRoom(1)) Then
        Call AddEnemies
    End If
    
    ' Add chests/loot
    If Rnd() < 0.3 Then  ' 30% chance of chest
        Call AddChest
    End If
End Sub

Sub AddInteriorWalls()
    ' Add some interior walls for interesting room layout
    Dim wallCount As Integer
    wallCount = Application.WorksheetFunction.RandBetween(2, 5)
    
    Dim i As Integer
    For i = 1 To wallCount
        Dim row As Integer, col As Integer
        Dim length As Integer
        Dim horizontal As Boolean
        
        horizontal = (Rnd() < 0.5)
        length = Application.WorksheetFunction.RandBetween(3, 6)
        
        If horizontal Then
            row = Application.WorksheetFunction.RandBetween(3, 16)
            col = Application.WorksheetFunction.RandBetween(3, 14)
            
            Dim c As Integer
            For c = col To Application.WorksheetFunction.Min(col + length, 17)
                If CurrentRoomGrid(row, c) = 0 Then  ' Only place on floor
                    CurrentRoomGrid(row, c) = 1
                End If
            Next c
        Else
            row = Application.WorksheetFunction.RandBetween(3, 14)
            col = Application.WorksheetFunction.RandBetween(3, 16)
            
            Dim r As Integer
            For r = row To Application.WorksheetFunction.Min(row + length, 17)
                If CurrentRoomGrid(r, col) = 0 Then
                    CurrentRoomGrid(r, col) = 1
                End If
            Next r
        End If
    Next i
End Sub

Sub AddEnemies()
    ' Add 1-3 enemies to the room
    Dim enemyCount As Integer
    enemyCount = Application.WorksheetFunction.RandBetween(1, 3)
    
    Dim i As Integer
    For i = 1 To enemyCount
        Dim placed As Boolean
        placed = False
        
        Dim attempts As Integer
        attempts = 0
        
        Do While Not placed And attempts < 20
            Dim row As Integer, col As Integer
            row = Application.WorksheetFunction.RandBetween(3, 16)
            col = Application.WorksheetFunction.RandBetween(3, 16)
            
            If CurrentRoomGrid(row, col) = 0 Then  ' Empty floor
                CurrentRoomGrid(row, col) = 2  ' Enemy
                placed = True
            End If
            
            attempts = attempts + 1
        Loop
    Next i
End Sub

Sub AddChest()
    ' Add a chest to the room
    Dim placed As Boolean
    placed = False
    
    Dim attempts As Integer
    attempts = 0
    
    Do While Not placed And attempts < 20
        Dim row As Integer, col As Integer
        row = Application.WorksheetFunction.RandBetween(3, 16)
        col = Application.WorksheetFunction.RandBetween(3, 16)
        
        If CurrentRoomGrid(row, col) = 0 Then
            CurrentRoomGrid(row, col) = 3  ' Chest
            placed = True
        End If
        
        attempts = attempts + 1
    Loop
End Sub

' ============================================================================
' ROOM DRAWING
' ============================================================================

Sub DrawRoom()
    ' Draw the current room in gameplay area (K4:AK24 = 21x27)
    ' OPTIMIZED: Use batch operations and disable events
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Game")
    
    ' Disable updates for faster rendering
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo CleanUp
    
    ' Clear gameplay area in one operation
    With ws.Range("K4:AK24")
        .Clear
        .Interior.Color = RGB(139, 69, 19)  ' Brown floor default
    End With
    
    ' Pre-define colors for faster access
    Dim floorColor As Long, wallColor As Long, doorColor As Long
    floorColor = RGB(139, 69, 19)
    wallColor = RGB(60, 60, 60)
    doorColor = RGB(100, 200, 255)
    
    ' Draw each cell - center 19x19 room in 21x27 area
    Dim row As Integer, col As Integer
    Dim cell As Range
    Dim cellType As Integer
    
    For row = 1 To 19
        For col = 1 To 19
            cellType = CurrentRoomGrid(row, col)
            If cellType > 0 Then  ' Only process non-floor cells
                Set cell = ws.Cells(4 + row, 10 + col + 4)
                
                Select Case cellType
                    Case 1  ' Wall
                        cell.Interior.Color = wallColor
                        With cell.Borders
                            .LineStyle = xlContinuous
                            .Color = RGB(40, 40, 40)
                            .Weight = xlThin
                        End With
                        
                    Case 2  ' Enemy
                        With cell
                            .Interior.Color = floorColor
                            .Value = "E"
                            .Font.Color = RGB(255, 0, 0)
                            .Font.Bold = True
                            .HorizontalAlignment = xlCenter
                            .VerticalAlignment = xlCenter
                        End With
                        
                    Case 3  ' Chest
                        With cell
                            .Interior.Color = floorColor
                            .Value = "C"
                            .Font.Color = RGB(255, 215, 0)
                            .Font.Bold = True
                            .HorizontalAlignment = xlCenter
                            .VerticalAlignment = xlCenter
                        End With
                        
                    Case 4, 5, 6, 7  ' Doors
                        With cell
                            .Interior.Color = doorColor
                            .Value = "D"
                            .Font.Color = RGB(255, 255, 255)
                            .Font.Bold = True
                            .Font.Size = 14
                            .HorizontalAlignment = xlCenter
                            .VerticalAlignment = xlCenter
                            .Borders.LineStyle = xlContinuous
                            .Borders.Color = RGB(255, 215, 0)
                            .Borders.Weight = xlThick
                        End With
                End Select
            End If
        Next col
    Next row
    
    ' Draw player
    Call DrawPlayer
    
CleanUp:
    ' Re-enable updates
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    Debug.Print "Room drawn at dungeon position (" & CurrentDungeonRow & "," & CurrentDungeonCol & ")"
End Sub

Sub DrawPlayer()
    ' Draw player sprite at current position - OPTIMIZED for smooth movement
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Game")
    
    ' Store old player position to clear it
    Static lastPlayerRow As Integer
    Static lastPlayerCol As Integer
    
    ' Disable screen updating for smoother animation
    Dim wasUpdating As Boolean
    wasUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    On Error GoTo CleanUp
    
    ' Clear old player position if it changed
    If lastPlayerRow > 0 And lastPlayerCol > 0 Then
        If lastPlayerRow <> PlayerRoomRow Or lastPlayerCol <> PlayerRoomCol Then
            Dim oldCell As Range
            Set oldCell = ws.Cells(4 + lastPlayerRow, 10 + lastPlayerCol + 4)
            oldCell.ClearContents
            ' Restore floor color based on what's in the grid
            Dim cellType As Integer
            cellType = CurrentRoomGrid(lastPlayerRow, lastPlayerCol)
            If cellType = 0 Then
                oldCell.Interior.Color = RGB(139, 69, 19)  ' Floor
            End If
        End If
    End If
    
    ' Delete old player sprite if it exists (faster than repositioning)
    On Error Resume Next
    ws.Shapes("PlayerSprite").Delete
    On Error GoTo CleanUp
    
    Dim playerCell As Range
    Set playerCell = ws.Cells(4 + PlayerRoomRow, 10 + PlayerRoomCol + 4)
    
    ' Try sprite first, fallback to text
    Dim assetsSheet As Worksheet
    On Error Resume Next
    Set assetsSheet = ThisWorkbook.Sheets("Assets")
    On Error GoTo CleanUp
    
    If Not assetsSheet Is Nothing Then
        Dim spriteSource As Shape
        On Error Resume Next
        Set spriteSource = assetsSheet.Shapes("sprite")
        On Error GoTo CleanUp
        
        If Not spriteSource Is Nothing Then
            ' Copy and paste sprite quickly
            Application.CutCopyMode = False
            spriteSource.Copy
            
            ' Paste without selecting - use PasteSpecial to avoid selection
            playerCell.Select
            ws.Paste
            
            ' Configure the pasted sprite
            Dim playerSprite As Shape
            Set playerSprite = ws.Shapes(ws.Shapes.Count)  ' Get last added shape
            playerSprite.Name = "PlayerSprite"
            
            With playerSprite
                .LockAspectRatio = msoTrue
                .Top = playerCell.Top + (playerCell.Height - .Height) / 2
                .Left = playerCell.Left + (playerCell.Width - .Width) / 2
                .Placement = xlFreeFloating
            End With
            
            Application.CutCopyMode = False
            
            ' CRITICAL: Deselect to remove selection box
            ws.Range("A1").Select
            
            GoTo UpdatePosition
        End If
    End If
    
    ' Fallback to text representation (faster than sprite)
    With playerCell
        .Interior.Color = RGB(139, 69, 19)
        .Value = "@"
        .Font.Color = RGB(0, 255, 0)
        .Font.Bold = True
        .Font.Size = 12
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Deselect to avoid selection box
    ws.Range("A1").Select
    
UpdatePosition:
    ' Update last position
    lastPlayerRow = PlayerRoomRow
    lastPlayerCol = PlayerRoomCol
    
CleanUp:
    ' Restore screen updating state
    Application.ScreenUpdating = wasUpdating
End Sub

' ============================================================================
' UI UPDATES
' ============================================================================

Public Sub UpdateDisplay()
    ' Update all dynamic UI elements - OPTIMIZED
    ' Only called when needed (not during continuous movement)
    
    Application.ScreenUpdating = False
    On Error Resume Next
    
    ' Redraw player position
    Call DrawPlayer
    
    ' Update stats panel
    Call UpdateStatsPanel
    
    ' Update minimap highlight
    Call UpdateMinimap
    
    Application.ScreenUpdating = True
End Sub

Sub UpdateStatsPanel()
    ' Update combat stats display (right panel)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Game")
    
    ' Main stats (Column BK-BN, rows 5-9)
    ws.Range("BK5").Value = PlayerCombatLevel
    ws.Range("BJ6").Value = PlayerAttackLevel
    ws.Range("BJ7").Value = PlayerStrengthLevel
    ws.Range("BJ8").Value = PlayerDefenceLevel
    ws.Range("BJ9").Value = PlayerHP & "/" & PlayerMaxHP
    ws.Range("BN6").Value = PlayerMagicLevel
    ws.Range("BN7").Value = PlayerRangedLevel
    ws.Range("BN8").Value = PlayerPrayer & "/" & PlayerMaxPrayer
    
    ' Equipment bonuses (rows 12-17)
    ws.Range("BJ12").Value = BonusStab
    ws.Range("BJ13").Value = BonusSlash
    ws.Range("BJ14").Value = BonusCrush
    ws.Range("BJ15").Value = BonusMagic
    ws.Range("BJ16").Value = BonusRanged
    ws.Range("BJ17").Value = BonusStrength
    
    ' Defense bonuses
    ws.Range("BN12").Value = DefStab
    ws.Range("BN13").Value = DefSlash
    ws.Range("BN14").Value = DefCrush
    ws.Range("BN15").Value = DefMagic
    ws.Range("BN16").Value = DefRanged
End Sub

Sub UpdateMinimap()
    ' Update minimap colors: visited rooms (cyan), current room (darker blue)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Game")
    
    Dim row As Integer, col As Integer
    Dim mapRow As Integer, mapCol As Integer
    Dim roomKey As String
    
    For row = 4 To 11
        For col = 2 To 9
            mapRow = row - 3
            mapCol = col - 1
            
            ' Only update rooms that exist in the dungeon
            If mapRow <= DungeonSize And mapCol <= DungeonSize Then
                roomKey = mapRow & "," & mapCol
                
                ' Check if this room exists in the dungeon
                If RoomExists(mapRow, mapCol) Then
                    ' Check if it's the current room
                    If mapRow = CurrentDungeonRow And mapCol = CurrentDungeonCol Then
                        ' Current room - darker blue
                        ws.Cells(row, col).Interior.Color = RGB(0, 100, 200)
                    ElseIf VisitedDungeonRooms.Exists(roomKey) Then
                        ' Visited room - cyan
                        ws.Cells(row, col).Interior.Color = RGB(0, 255, 255)
                    End If
                End If
            End If
        Next col
    Next row
End Sub

' ============================================================================
' CLICK HANDLER SETUP
' ============================================================================

Sub SetupClickHandlers()
    ' This should be called once at game start
    ' Sets up worksheet events for click handling
    ' Note: Requires worksheet change event in sheet module
    
    Debug.Print "Click handlers ready - Click gameplay area (K4:AK24) to move"
End Sub

