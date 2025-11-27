Option Explicit

' ============================================================================
' MODULE_GAMESTATE - Central data storage for all game variables
' ============================================================================

' === PLAYER STATS ===
Public PlayerHP As Integer
Public PlayerMaxHP As Integer
Public PlayerPrayer As Integer
Public PlayerMaxPrayer As Integer

' Combat stats
Public PlayerAttackLevel As Integer
Public PlayerStrengthLevel As Integer
Public PlayerDefenceLevel As Integer
Public PlayerRangedLevel As Integer
Public PlayerMagicLevel As Integer
Public PlayerCombatLevel As Integer

' === PLAYER POSITION ===
Public CurrentDungeonRow As Integer  ' Which dungeon room (1-8)
Public CurrentDungeonCol As Integer
Public PlayerRoomRow As Integer      ' Position within room (1-20)
Public PlayerRoomCol As Integer

' === MOVEMENT STATE ===
Public IsMoving As Boolean
Public TargetRow As Integer
Public TargetCol As Integer
Public PathBlocked As Boolean

' === GAME STATE ===
Public GameRunning As Boolean
Public GameTicks As Long
Public InCombat As Boolean

' === INVENTORY (28 slots like OSRS) ===
Public PlayerInventory(1 To 28) As Integer  ' ItemIDs, 0 = empty

' === EQUIPMENT ===
Public EquippedHelmet As Integer
Public EquippedAmulet As Integer
Public EquippedCape As Integer
Public EquippedWeapon As Integer
Public EquippedChest As Integer
Public EquippedShield As Integer
Public EquippedLegs As Integer
Public EquippedGloves As Integer
Public EquippedBoots As Integer
Public EquippedRing As Integer
Public EquippedAmmo As Integer

' === EQUIPMENT BONUSES (calculated from equipped items) ===
Public BonusStab As Integer
Public BonusSlash As Integer
Public BonusCrush As Integer
Public BonusMagic As Integer
Public BonusRanged As Integer
Public BonusStrength As Integer
Public DefStab As Integer
Public DefSlash As Integer
Public DefCrush As Integer
Public DefMagic As Integer
Public DefRanged As Integer

' === ACTIVE PRAYERS ===
Public PrayerActive(1 To 30) As Boolean  ' Track which prayers are on
Public PrayerDrainRate As Double

' === CURRENT ROOM GRID ===
' 19x19 grid representing current room
' 0 = floor, 1 = wall, 2 = enemy, 3 = chest, 4 = door_north, 5 = door_east, 6 = door_south, 7 = door_west
Public CurrentRoomGrid(1 To 19, 1 To 19) As Integer

' === VISITED ROOMS TRACKING ===
Public VisitedDungeonRooms As Object  ' Dictionary to track which dungeon rooms have been visited

' === COMBAT ===
Public CurrentEnemy As Object  ' Dictionary with enemy data
Public EnemyRoomRow As Integer
Public EnemyRoomCol As Integer

' ============================================================================
' INITIALIZATION FUNCTIONS
' ============================================================================

Public Sub InitializePlayer()
    ' Set starting stats
    PlayerHP = 100
    PlayerMaxHP = 100
    PlayerPrayer = 10
    PlayerMaxPrayer = 10
    
    ' Starting combat levels
    PlayerAttackLevel = 1
    PlayerStrengthLevel = 1
    PlayerDefenceLevel = 1
    PlayerRangedLevel = 1
    PlayerMagicLevel = 1
    PlayerCombatLevel = 3
    
    ' Clear inventory
    Dim i As Integer
    For i = 1 To 28
        PlayerInventory(i) = 0
    Next i
    
    ' Clear equipment
    EquippedHelmet = 0
    EquippedAmulet = 0
    EquippedCape = 0
    EquippedWeapon = 0
    EquippedChest = 0
    EquippedShield = 0
    EquippedLegs = 0
    EquippedGloves = 0
    EquippedBoots = 0
    EquippedRing = 0
    EquippedAmmo = 0
    
    ' Clear bonuses
    BonusStab = 0
    BonusSlash = 0
    BonusCrush = 0
    BonusMagic = 0
    BonusRanged = 0
    BonusStrength = 0
    DefStab = 0
    DefSlash = 0
    DefCrush = 0
    DefMagic = 0
    DefRanged = 0
    
    ' Clear prayers
    For i = 1 To 30
        PrayerActive(i) = False
    Next i
    
    ' Reset state
    GameTicks = 0
    InCombat = False
    IsMoving = False
    PathBlocked = False
End Sub

Public Sub ResetGameState()
    ' Full reset for new game
    Call InitializePlayer
    
    CurrentDungeonRow = 0
    CurrentDungeonCol = 0
    PlayerRoomRow = 10  ' START AT CENTER
    PlayerRoomCol = 10  ' START AT CENTER
    
    GameRunning = False
    Set CurrentEnemy = Nothing
    
    ' Initialize visited rooms tracking
    Set VisitedDungeonRooms = CreateObject("Scripting.Dictionary")
    
    Debug.Print "Game state reset - Player room position set to (10,10)"
End Sub

' ============================================================================
' DATA ACCESS FUNCTIONS
' ============================================================================

Function GetPlayerStats() As Object
    ' Returns dictionary of player stats for easy access
    Dim stats As Object
    Set stats = CreateObject("Scripting.Dictionary")
    
    stats.Add "HP", PlayerHP
    stats.Add "MaxHP", PlayerMaxHP
    stats.Add "Prayer", PlayerPrayer
    stats.Add "MaxPrayer", PlayerMaxPrayer
    stats.Add "Attack", PlayerAttackLevel
    stats.Add "Strength", PlayerStrengthLevel
    stats.Add "Defence", PlayerDefenceLevel
    stats.Add "Ranged", PlayerRangedLevel
    stats.Add "Magic", PlayerMagicLevel
    stats.Add "Combat", PlayerCombatLevel
    
    Set GetPlayerStats = stats
End Function

Function GetEquipmentBonuses() As Object
    ' Returns dictionary of equipment bonuses
    Dim bonuses As Object
    Set bonuses = CreateObject("Scripting.Dictionary")
    
    bonuses.Add "Stab", BonusStab
    bonuses.Add "Slash", BonusSlash
    bonuses.Add "Crush", BonusCrush
    bonuses.Add "Magic", BonusMagic
    bonuses.Add "Ranged", BonusRanged
    bonuses.Add "Strength", BonusStrength
    bonuses.Add "DefStab", DefStab
    bonuses.Add "DefSlash", DefSlash
    bonuses.Add "DefCrush", DefCrush
    bonuses.Add "DefMagic", DefMagic
    bonuses.Add "DefRanged", DefRanged
    
    Set GetEquipmentBonuses = bonuses
End Function

Public Sub SaveGameState()
    ' Save current game state to Data_PlayerSave sheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Data_PlayerSave")
    
    ' Update save data
    ws.Range("B2").Value = PlayerHP
    ws.Range("B3").Value = PlayerMaxHP
    ws.Range("B4").Value = PlayerRoomRow
    ws.Range("B5").Value = PlayerRoomCol
    ws.Range("B6").Value = CurrentDungeonRow
    ws.Range("B7").Value = CurrentDungeonCol
    
    ' Save inventory as comma-separated string
    Dim invString As String
    Dim i As Integer
    For i = 1 To 28
        invString = invString & PlayerInventory(i)
        If i < 28 Then invString = invString & ","
    Next i
    ws.Range("B8").Value = invString
End Sub

Public Sub LoadGameState()
    ' Load game state from Data_PlayerSave sheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Data_PlayerSave")
    
    PlayerHP = ws.Range("B2").Value
    PlayerMaxHP = ws.Range("B3").Value
    PlayerRoomRow = ws.Range("B4").Value
    PlayerRoomCol = ws.Range("B5").Value
    CurrentDungeonRow = ws.Range("B6").Value
    CurrentDungeonCol = ws.Range("B7").Value
    
    ' Load inventory
    Dim invString As String
    Dim invArray As Variant
    Dim i As Integer
    
    invString = ws.Range("B8").Value
    If Len(invString) > 0 Then
        invArray = Split(invString, ",")
        For i = 1 To 28
            PlayerInventory(i) = CInt(invArray(i - 1))
        Next i
    End If
End Sub
