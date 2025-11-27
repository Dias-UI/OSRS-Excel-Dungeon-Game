Option Explicit

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' Handle clicks in gameplay area (K4:AK24)
    
    If Not Intersect(Target, Me.Range("K4:AK24")) Is Nothing Then
        If Target.Cells.Count = 1 Then
            Call Module_GameLoop.OnGameAreaClick(Target)
        End If
    End If
End Sub
