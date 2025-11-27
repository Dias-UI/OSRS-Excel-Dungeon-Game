Option Explicit

' ============================================================================
' DUNGEON GENERATION SYSTEM
' Generates small (4x4), medium (6x6), or large (8x8) dungeons
' ============================================================================

' Dungeon structure: rooms(row, col, direction)
' Direction: 1=North, 2=East, 3=South, 4=West (True = door exists)
Public DungeonRooms() As Boolean
Public DungeonSize As Integer
Public DungeonRoomCount As Integer
Public StartRoom As Variant ' Array(row, col)
Public EndRoom As Variant   ' Array(row, col)

' Room tracking
Private VisitedRooms As Object ' Dictionary for existence checking
Private RoomQueue As Object    ' Collection for BFS pathfinding

' ============================================================================
' MAIN GENERATION FUNCTION
' ============================================================================
Sub GenerateDungeon(dungeonType As String)
    ' dungeonType: "small", "medium", or "large"
    
    Dim targetSize As Integer
    Dim maxAttempts As Integer
    Dim attempt As Integer
    
    ' Set dungeon parameters
    Select Case LCase(dungeonType)
        Case "small"
            targetSize = 4
        Case "medium"
            targetSize = 6
        Case "large"
            targetSize = 8
        Case Else
            MsgBox "Invalid dungeon type. Use: small, medium, or large"
            Exit Sub
    End Select
    
    DungeonSize = targetSize
    maxAttempts = 100
    
    ' Keep generating until we get a good dungeon
    Dim targetFillRate As Double
    Dim minRoomCount As Integer
    
    If targetSize = 8 Then
        targetFillRate = 0.65  ' 65% for large dungeons (less dense, more interesting)
        minRoomCount = 42      ' At least 42 rooms
    ElseIf targetSize = 6 Then
        targetFillRate = 0.7   ' 70% for medium
        minRoomCount = 25      ' At least 25 rooms
    Else
        targetFillRate = 0.75  ' 75% for small
        minRoomCount = 12      ' At least 12 rooms
    End If
    
    For attempt = 1 To maxAttempts
        Call InitializeDungeon
        Call BuildDungeon
        
        ' Check if dungeon meets quality criteria
        ' Also ensure end room is a dead end
        If DungeonRoomCount >= minRoomCount And _
           IsArray(EndRoom) And CountRoomDoors(CInt(EndRoom(0)), CInt(EndRoom(1))) = 1 Then
            Exit For ' Good dungeon generated
        End If
    Next attempt
    
    ' Draw the preview map
    Call DrawDungeonPreview
    
    Debug.Print "Dungeon generated: " & DungeonSize & "x" & DungeonSize
    Debug.Print "Rooms created: " & DungeonRoomCount & " (" & _
                Format(DungeonRoomCount / (DungeonSize * DungeonSize), "0%") & ")"
    Debug.Print "Start: (" & StartRoom(0) & "," & StartRoom(1) & ")"
    Debug.Print "End: (" & EndRoom(0) & "," & EndRoom(1) & ")"
End Sub

' ============================================================================
' INITIALIZATION
' ============================================================================
Sub InitializeDungeon()
    ' Reset dungeon array
    ReDim DungeonRooms(1 To 8, 1 To 8, 1 To 4)
    
    ' Initialize tracking
    Set VisitedRooms = CreateObject("Scripting.Dictionary")
    DungeonRoomCount = 0
    
    ' Place start room randomly
    Dim startRow As Integer, startCol As Integer
    startRow = Application.WorksheetFunction.RandBetween(1, DungeonSize)
    startCol = Application.WorksheetFunction.RandBetween(1, DungeonSize)
    
    StartRoom = Array(startRow, startCol)
    Call AddRoom(startRow, startCol)
End Sub

' ============================================================================
' DUNGEON BUILDING ALGORITHM
' ============================================================================
Sub BuildDungeon()
    ' Build critical path first (start to end)
    Call BuildCriticalPath
    
    ' Check if critical path succeeded
    If Not IsArray(EndRoom) Then
        ' Failed to create valid dungeon, will retry
        DungeonRoomCount = 0
        Exit Sub
    End If
    
    ' Add branching paths - multiple passes for better coverage
    Dim pass As Integer
    Dim maxPasses As Integer
    
    ' More passes for larger dungeons
    If DungeonSize = 8 Then
        maxPasses = 3
    ElseIf DungeonSize = 6 Then
        maxPasses = 3  ' Increased for medium
    Else
        maxPasses = 3  ' Increased for small
    End If
    
    For pass = 1 To maxPasses
        Call AddBranchingPaths
    Next pass
    
    ' Fill in obvious gaps
    Call FillGaps
    
    ' IMPORTANT: Close off end room if it has multiple doors
    Call EnsureEndRoomIsDeadEnd
    
    ' Ensure end room is still reachable after modifications
    If Not IsRoomReachable(CInt(EndRoom(0)), CInt(EndRoom(1))) Then
        ' Failed to create valid dungeon, will retry
        DungeonRoomCount = 0
    End If
End Sub

Sub BuildCriticalPath()
    ' Create main path from start to end
    Dim currentRow As Integer, currentCol As Integer
    Dim pathLength As Integer
    Dim minPathLength As Integer
    Dim direction As Integer
    Dim newRow As Integer, newCol As Integer
    Dim attempts As Integer
    
    currentRow = StartRoom(0)
    currentCol = StartRoom(1)
    
    ' Minimum path length based on dungeon size
    minPathLength = DungeonSize + Application.WorksheetFunction.RandBetween(2, DungeonSize)
    pathLength = 0
    
    ' Build path
    Do While pathLength < minPathLength Or Not IsValidEndPosition(currentRow, currentCol)
        attempts = 0
        
        ' Try to extend path
        Do
            direction = Application.WorksheetFunction.RandBetween(1, 4)
            newRow = currentRow
            newCol = currentCol
            
            Select Case direction
                Case 1: newRow = newRow - 1 ' North
                Case 2: newCol = newCol + 1 ' East
                Case 3: newRow = newRow + 1 ' South
                Case 4: newCol = newCol - 1 ' West
            End Select
            
            attempts = attempts + 1
            If attempts > 20 Then Exit Do ' Prevent infinite loop
            
        Loop Until CanPlaceRoom(newRow, newCol)
        
        If attempts > 20 Then
            ' Backtrack or restart
            Exit Sub
        End If
        
        ' Place room and create door
        Call AddRoom(newRow, newCol)
        Call CreateDoor(currentRow, currentCol, newRow, newCol)
        
        currentRow = newRow
        currentCol = newCol
        pathLength = pathLength + 1
    Loop
    
    ' Set end room
    EndRoom = Array(currentRow, currentCol)
End Sub

Sub AddBranchingPaths()
    ' Add branches from existing rooms
    ' For large dungeons, be more aggressive with branching
    Dim roomKeys As Variant
    Dim i As Integer
    Dim roomKey As String
    Dim roomCoords As Variant
    Dim baseRow As Integer, baseCol As Integer
    Dim branchLength As Integer
    Dim direction As Integer
    Dim newRow As Integer, newCol As Integer
    Dim currentRow As Integer, currentCol As Integer
    Dim j As Integer
    Dim branchChance As Double
    Dim maxBranchLength As Integer
    
    ' Scale branching based on dungeon size
    ' Smaller dungeons need MORE branching relative to size
    If DungeonSize = 8 Then
        branchChance = 0.7  ' 70% chance for large dungeons
        maxBranchLength = 5
    ElseIf DungeonSize = 6 Then
        branchChance = 0.8  ' 80% chance for medium (increased)
        maxBranchLength = 4
    Else
        branchChance = 0.9  ' 90% chance for small (increased)
        maxBranchLength = 3
    End If
    
    ' Get current rooms (important: get fresh list each call)
    roomKeys = VisitedRooms.Keys
    
    ' Try to add branches from each existing room (except end room)
    For i = LBound(roomKeys) To UBound(roomKeys)
        roomKey = roomKeys(i)
        roomCoords = Split(roomKey, ",")
        baseRow = CInt(roomCoords(0))
        baseCol = CInt(roomCoords(1))
        
        ' Don't branch from end room
        If IsArray(EndRoom) Then
            If baseRow = EndRoom(0) And baseCol = EndRoom(1) Then
                GoTo NextRoom
            End If
        End If
        
        If Rnd() < branchChance Then
            ' Random branch length
            branchLength = Application.WorksheetFunction.RandBetween(2, maxBranchLength)
            
            currentRow = baseRow
            currentCol = baseCol
            
            ' Build branch
            For j = 1 To branchLength
                ' Try each direction randomly
                Dim directions As Variant
                directions = Array(1, 2, 3, 4)
                Call ShuffleArray(directions)
                
                Dim foundDirection As Boolean
                foundDirection = False
                
                Dim k As Integer
                For k = LBound(directions) To UBound(directions)
                    direction = directions(k)
                    newRow = currentRow
                    newCol = currentCol
                    
                    Select Case direction
                        Case 1: newRow = newRow - 1
                        Case 2: newCol = newCol + 1
                        Case 3: newRow = newRow + 1
                        Case 4: newCol = newCol - 1
                    End Select
                    
                    If CanPlaceRoom(newRow, newCol) Then
                        Call AddRoom(newRow, newCol)
                        Call CreateDoor(currentRow, currentCol, newRow, newCol)
                        currentRow = newRow
                        currentCol = newCol
                        foundDirection = True
                        Exit For
                    End If
                Next k
                
                If Not foundDirection Then Exit For ' Dead end
            Next j
        End If
NextRoom:
    Next i
End Sub

Sub FillGaps()
    ' Fill in isolated empty spaces, but maintain dungeon feel
    Dim row As Integer, col As Integer
    Dim adjacentRooms As Integer
    Dim addedRooms As Boolean
    Dim fillChance As Double
    Dim minAdjacent As Integer
    
    ' Control density - smaller dungeons need more filling
    If DungeonSize = 8 Then
        fillChance = 0.4  ' Only fill 40% of eligible gaps for large
        minAdjacent = 3   ' Need 3+ adjacent for large
    ElseIf DungeonSize = 6 Then
        fillChance = 0.6  ' 60% for medium
        minAdjacent = 2   ' Need 2+ adjacent for medium
    Else
        fillChance = 0.7  ' 70% for small
        minAdjacent = 2   ' Need 2+ adjacent for small
    End If
    
    ' Single pass to avoid over-filling
    For row = 1 To DungeonSize
        For col = 1 To DungeonSize
            ' Skip if room already exists
            If Not VisitedRooms.Exists(row & "," & col) Then
                adjacentRooms = 0
                
                ' Count adjacent rooms
                If row > 1 And VisitedRooms.Exists((row - 1) & "," & col) Then adjacentRooms = adjacentRooms + 1
                If row < DungeonSize And VisitedRooms.Exists((row + 1) & "," & col) Then adjacentRooms = adjacentRooms + 1
                If col > 1 And VisitedRooms.Exists(row & "," & (col - 1)) Then adjacentRooms = adjacentRooms + 1
                If col < DungeonSize And VisitedRooms.Exists(row & "," & (col + 1)) Then adjacentRooms = adjacentRooms + 1
                
                ' Fill based on size-specific rules
                If adjacentRooms >= minAdjacent And Rnd() < fillChance Then
                    Call AddRoom(row, col)
                    
                    ' Connect to adjacent rooms
                    If row > 1 And VisitedRooms.Exists((row - 1) & "," & col) Then
                        Call CreateDoor(row, col, row - 1, col)
                    End If
                    If row < DungeonSize And VisitedRooms.Exists((row + 1) & "," & col) Then
                        Call CreateDoor(row, col, row + 1, col)
                    End If
                    If col > 1 And VisitedRooms.Exists(row & "," & (col - 1)) Then
                        Call CreateDoor(row, col, row, col - 1)
                    End If
                    If col < DungeonSize And VisitedRooms.Exists(row & "," & (col + 1)) Then
                        Call CreateDoor(row, col, row, col + 1)
                    End If
                End If
            End If
        Next col
    Next row
End Sub

' ============================================================================
' ROOM MANAGEMENT
' ============================================================================
Function CanPlaceRoom(row As Integer, col As Integer) As Boolean
    ' Check if room can be placed at position
    CanPlaceRoom = False
    
    ' Check bounds
    If row < 1 Or row > DungeonSize Or col < 1 Or col > DungeonSize Then
        Exit Function
    End If
    
    ' Check if already exists
    If VisitedRooms.Exists(row & "," & col) Then
        Exit Function
    End If
    
    ' For large dungeons, be less strict about adjacent rooms
    If DungeonSize = 8 Then
        ' Allow up to 3 adjacent rooms for large dungeons
        Dim adjacentCount As Integer
        adjacentCount = 0
        
        If row > 1 And VisitedRooms.Exists((row - 1) & "," & col) Then adjacentCount = adjacentCount + 1
        If row < DungeonSize And VisitedRooms.Exists((row + 1) & "," & col) Then adjacentCount = adjacentCount + 1
        If col > 1 And VisitedRooms.Exists(row & "," & (col - 1)) Then adjacentCount = adjacentCount + 1
        If col < DungeonSize And VisitedRooms.Exists(row & "," & (col + 1)) Then adjacentCount = adjacentCount + 1
        
        If adjacentCount > 3 Then Exit Function
    Else
        ' Standard rules for small/medium
        Dim adjCount As Integer
        adjCount = 0
        
        If row > 1 And VisitedRooms.Exists((row - 1) & "," & col) Then adjCount = adjCount + 1
        If row < DungeonSize And VisitedRooms.Exists((row + 1) & "," & col) Then adjCount = adjCount + 1
        If col > 1 And VisitedRooms.Exists(row & "," & (col - 1)) Then adjCount = adjCount + 1
        If col < DungeonSize And VisitedRooms.Exists(row & "," & (col + 1)) Then adjCount = adjCount + 1
        
        If adjCount > 2 Then Exit Function
    End If
    
    CanPlaceRoom = True
End Function

Sub AddRoom(row As Integer, col As Integer)
    ' Add room to dungeon
    VisitedRooms.Add row & "," & col, True
    DungeonRoomCount = DungeonRoomCount + 1
End Sub

Sub CreateDoor(fromRow As Integer, fromCol As Integer, toRow As Integer, toCol As Integer)
    ' Create door between two adjacent rooms
    
    If fromRow < toRow Then ' Door south from first room
        DungeonRooms(fromRow, fromCol, 3) = True
        DungeonRooms(toRow, toCol, 1) = True
    ElseIf fromRow > toRow Then ' Door north from first room
        DungeonRooms(fromRow, fromCol, 1) = True
        DungeonRooms(toRow, toCol, 3) = True
    ElseIf fromCol < toCol Then ' Door east from first room
        DungeonRooms(fromRow, fromCol, 2) = True
        DungeonRooms(toRow, toCol, 4) = True
    ElseIf fromCol > toCol Then ' Door west from first room
        DungeonRooms(fromRow, fromCol, 4) = True
        DungeonRooms(toRow, toCol, 2) = True
    End If
End Sub

Function IsValidEndPosition(row As Integer, col As Integer) As Boolean
    ' Check if position is suitable for end room
    ' Should be far from start and not in a corner
    
    Dim distance As Double
    distance = Abs(row - StartRoom(0)) + Abs(col - StartRoom(1))
    
    ' Must be at least half dungeon size away
    IsValidEndPosition = (distance >= DungeonSize / 2)
End Function

Function IsRoomReachable(row As Long, col As Long) As Boolean
    ' BFS to check if room is reachable from start
    Dim queue As Object
    Dim visited As Object
    Dim currentKey As String
    Dim currentCoords As Variant
    Dim currentRow As Integer, currentCol As Integer
    
    Set queue = CreateObject("System.Collections.ArrayList")
    Set visited = CreateObject("Scripting.Dictionary")
    
    queue.Add StartRoom(0) & "," & StartRoom(1)
    visited.Add StartRoom(0) & "," & StartRoom(1), True
    
    Do While queue.Count > 0
        currentKey = queue(0)
        queue.RemoveAt 0
        
        currentCoords = Split(currentKey, ",")
        currentRow = CInt(currentCoords(0))
        currentCol = CInt(currentCoords(1))
        
        ' Check if we reached target
        If currentRow = row And currentCol = col Then
            IsRoomReachable = True
            Exit Function
        End If
        
        ' Add adjacent rooms with doors
        If DungeonRooms(currentRow, currentCol, 1) And Not visited.Exists((currentRow - 1) & "," & currentCol) Then
            queue.Add (currentRow - 1) & "," & currentCol
            visited.Add (currentRow - 1) & "," & currentCol, True
        End If
        If DungeonRooms(currentRow, currentCol, 2) And Not visited.Exists(currentRow & "," & (currentCol + 1)) Then
            queue.Add currentRow & "," & (currentCol + 1)
            visited.Add currentRow & "," & (currentCol + 1), True
        End If
        If DungeonRooms(currentRow, currentCol, 3) And Not visited.Exists((currentRow + 1) & "," & currentCol) Then
            queue.Add (currentRow + 1) & "," & currentCol
            visited.Add (currentRow + 1) & "," & currentCol, True
        End If
        If DungeonRooms(currentRow, currentCol, 4) And Not visited.Exists(currentRow & "," & (currentCol - 1)) Then
            queue.Add currentRow & "," & (currentCol - 1)
            visited.Add currentRow & "," & (currentCol - 1), True
        End If
    Loop
    
    IsRoomReachable = False
End Function

' ============================================================================
' PREVIEW DRAWING
' ============================================================================
Sub DrawDungeonPreview()
    ' Draw dungeon in B4:I11 (8x8 grid)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Game")
    
    Dim previewRange As Range
    Set previewRange = ws.Range("B4:I11")
    
    ' Clear and format preview area
    previewRange.Clear
    previewRange.Interior.Color = RGB(0, 0, 0) ' Black for unused
    previewRange.Borders.LineStyle = xlNone
    
    ' Draw each room
    Dim row As Integer, col As Integer
    Dim cell As Range
    
    For row = 1 To 8
        For col = 1 To 8
            Set cell = ws.Cells(3 + row, 1 + col) ' B4 starts at (4,2)
            
            If row <= DungeonSize And col <= DungeonSize And VisitedRooms.Exists(row & "," & col) Then
                ' Room exists - white fill
                cell.Interior.Color = RGB(255, 255, 255)
                
                ' Add borders for doors
                With cell.Borders
                    .LineStyle = xlContinuous
                    .Color = RGB(0, 0, 0)
                    .Weight = xlMedium
                End With
                
                ' Open borders where doors exist
                If DungeonRooms(row, col, 1) Then ' North door
                    cell.Borders(xlEdgeTop).LineStyle = xlNone
                End If
                If DungeonRooms(row, col, 2) Then ' East door
                    cell.Borders(xlEdgeRight).LineStyle = xlNone
                End If
                If DungeonRooms(row, col, 3) Then ' South door
                    cell.Borders(xlEdgeBottom).LineStyle = xlNone
                End If
                If DungeonRooms(row, col, 4) Then ' West door
                    cell.Borders(xlEdgeLeft).LineStyle = xlNone
                End If
                
                ' Mark start room (green)
                If row = StartRoom(0) And col = StartRoom(1) Then
                    cell.Interior.Color = RGB(100, 255, 100)
                    cell.Value = "S"
                    cell.Font.Bold = True
                    cell.HorizontalAlignment = xlCenter
                    cell.VerticalAlignment = xlCenter
                End If
                
                ' Mark end room (red)
                If row = EndRoom(0) And col = EndRoom(1) Then
                    cell.Interior.Color = RGB(255, 100, 100)
                    cell.Value = "E"
                    cell.Font.Bold = True
                    cell.HorizontalAlignment = xlCenter
                    cell.VerticalAlignment = xlCenter
                End If
            End If
        Next col
    Next row
End Sub

' ============================================================================
' HELPER FUNCTIONS
' ============================================================================
Sub ShuffleArray(arr As Variant)
    ' Fisher-Yates shuffle
    Dim i As Integer
    Dim j As Integer
    Dim temp As Variant
    
    For i = UBound(arr) To LBound(arr) + 1 Step -1
        j = Int((i - LBound(arr) + 1) * Rnd + LBound(arr))
        temp = arr(i)
        arr(i) = arr(j)
        arr(j) = temp
    Next i
End Sub

' ============================================================================
' PUBLIC HELPER: Get Room Door Info
' ============================================================================
Public Function GetRoomDoors(row As Integer, col As Integer) As Variant
    ' Returns array of booleans: (North, East, South, West)
    Dim doors(1 To 4) As Boolean
    
    If row >= 1 And row <= 8 And col >= 1 And col <= 8 Then
        doors(1) = DungeonRooms(row, col, 1)
        doors(2) = DungeonRooms(row, col, 2)
        doors(3) = DungeonRooms(row, col, 3)
        doors(4) = DungeonRooms(row, col, 4)
    End If
    
    GetRoomDoors = doors
End Function

Public Function RoomExists(row As Integer, col As Integer) As Boolean
    ' Check if room exists in dungeon
    RoomExists = VisitedRooms.Exists(row & "," & col)
End Function

Public Function CountRoomDoors(row As Integer, col As Integer) As Integer
    ' Count how many doors a room has
    Dim doorCount As Integer
    doorCount = 0
    
    If row >= 1 And row <= 8 And col >= 1 And col <= 8 Then
        If DungeonRooms(row, col, 1) Then doorCount = doorCount + 1
        If DungeonRooms(row, col, 2) Then doorCount = doorCount + 1
        If DungeonRooms(row, col, 3) Then doorCount = doorCount + 1
        If DungeonRooms(row, col, 4) Then doorCount = doorCount + 1
    End If
    
    CountRoomDoors = doorCount
End Function

Sub EnsureEndRoomIsDeadEnd()
    ' Make sure end room only has 1 door
    If Not IsArray(EndRoom) Then Exit Sub
    
    Dim endRow As Integer, endCol As Integer
    endRow = CInt(EndRoom(0))
    endCol = CInt(EndRoom(1))
    
    Dim doorCount As Integer
    doorCount = CountRoomDoors(endRow, endCol)
    
    ' If end room has more than 1 door, close extras
    If doorCount > 1 Then
        ' Keep the door we came from (first one found), close others
        Dim keptOne As Boolean
        keptOne = False
        
        If DungeonRooms(endRow, endCol, 1) Then ' North
            If Not keptOne Then
                keptOne = True
            Else
                DungeonRooms(endRow, endCol, 1) = False
                If endRow > 1 Then DungeonRooms(endRow - 1, endCol, 3) = False
            End If
        End If
        
        If DungeonRooms(endRow, endCol, 2) Then ' East
            If Not keptOne Then
                keptOne = True
            Else
                DungeonRooms(endRow, endCol, 2) = False
                If endCol < DungeonSize Then DungeonRooms(endRow, endCol + 1, 4) = False
            End If
        End If
        
        If DungeonRooms(endRow, endCol, 3) Then ' South
            If Not keptOne Then
                keptOne = True
            Else
                DungeonRooms(endRow, endCol, 3) = False
                If endRow < DungeonSize Then DungeonRooms(endRow + 1, endCol, 1) = False
            End If
        End If
        
        If DungeonRooms(endRow, endCol, 4) Then ' West
            If Not keptOne Then
                keptOne = True
            Else
                DungeonRooms(endRow, endCol, 4) = False
                If endCol > 1 Then DungeonRooms(endRow, endCol - 1, 2) = False
            End If
        End If
    End If
End Sub

