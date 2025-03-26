'Moddule 1

'Create Row
Sub Button1_Click()
    'Disable Screen Update
    Application.ScreenUpdating = False
    
    'Variable Declaration
    Dim Rng As Range
    Dim shtRng As Range
    Dim WrkSht As Worksheet
    Dim i As Integer
    Dim from As String
    Dim tohere As String
    Dim val As Integer
    
    'Variable Initialization
    i = 1
    
    Set WrkSht = Application.ActiveSheet
    Set Rng = Selection
    val = Rng.Row
    from = "G" + Mid(Str(val), 2)
    tohere = "M" + Mid(Str(val), 2)
    Set shtRng = Range(from, tohere)
    
    For Each Rng In shtRng
        With WrkSht.CheckBoxes.Add(Left : = Rng.Left, Top : = Rng.Top, Width : = Rng.Width, Height : = Rng.Height).Select
            With Selection
                .Characters.Text = Rng.Value
                .Caption = ""
                .Caption = "Day " & i
                i = i + 1
            End With
        End With
    Next
    
    i = 1
    
    Set shtRng = Range("E" + Mid(Str(val), 2))
    
    For Each Rng In shtRng
        With WrkSht.CheckBoxes.Add(Left : = Rng.Left, Top : = Rng.Top, Width : = Rng.Width, Height : = Rng.Height).Select
            With Selection
                .Characters.Text = Rng.Value
                .Caption = ""
                .Caption = "Alarm " & i
                i = i + 1
            End With
        End With
    Next
    
    shtRng.ClearContents
    shtRng.Select
    
    'Enable Screen Update
    Application.ScreenUpdating = True
End Sub

'Delete Row
Sub Button2_Click()
    
    Dim from As String
    Dim tohere As String
    Dim val As Integer
    Dim ws As Worksheet
    Dim myRange As Range
    Dim check As CheckBox

    Set ws = Sheets("Sheet1")

    Set Rng = Selection
    val = Rng.Row
    from = "E" + Mid(Str(val), 2)
    tohere = "M" + Mid(Str(val), 2)

    'OD Checkboxes
    Set myRange = ws.Range(from, tohere)

    For Each check In ws.CheckBoxes
        If Not Intersect(check.TopLeftCell, myRange) Is Nothing Then
            check.Delete
        End If
    Next
    
    Range("A" + Mid(Str(val), 2)).Clear
    Range("A" + Mid(Str(val), 2)).EntireRow.Delete
End Sub