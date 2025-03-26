'Module 2

Dim CountDown As Date
Sub Timer()
    CountDown = Now + TimeValue("00:01:00")
    Application.OnTime CountDown, "Reset"
End Sub


Sub Reset()
    'On Error GoTo 10
    Dim count As Range
    Dim shtRng As Range
    Dim Day As String
    Dim txt As String
    Dim prog As String
    Dim addRng As String
    Dim WrkBook As Workbook
    Dim WrkSheet As Worksheet
    Dim chkBox As Shape
    Dim cellAddress As String
    Dim chkBoxName As String
    Dim isChecked As Boolean

    WrBook = "RPClock.xlsm"

    Set shtRng = Range("A5:A100")

    For Each Rng In shtRng

        Workbooks("RPClock.xlsm").Sheets("Sheet1").Range("A3").Value = Time()

        10 : If Rng.Value <> "" Then
            Day = Format(Now(), "dddd")
            If Format(Range("A" + Mid(Str(Rng.Row), 2)).Value, "hh:mm") = Format(Time(), "hh:mm") Then
                If CheckDay(Rng.Row, Day) Then
                    Beep
                    txt = Range("D" + Mid(Str(Rng.Row), 2)).Value

                    cellAddress = Range("E" + Mid(Str(Rng.Row), 2)).Address ' Get the cell address
                    chkBoxName = ""
                    
                    For Each chkBox In ActiveSheet.Shapes
                        If chkBox.Type = msoFormControl Then ' Check if the shape is a form control (like a checkbox)
                            If chkBox.TopLeftCell.Address = cellAddress Then ' Match checkbox's top-left cell to your range
                                chkBoxName = chkBox.Name
                                If chkBox.OLEFormat.Object.Value = 1 Then ' Check if the checkbox is ticked (Value = 1)
                                    isChecked = True
                                End If
                                Exit For
                            End If
                        End If
                    Next chkBox
                    
                    If isChecked Then ' Check if the check
                        prog = "c:/Users/rchrd/Documents/Python/CallerIDGUI/.venv/Scripts/python.exe C:\Users\rchrd\Documents\Python\Playsounds.py ""Chimes2"" """ & txt & """"
                    Else
                        prog = "c:/Users/rchrd/Documents/Python/CallerIDGUI/.venv/Scripts/python.exe C:\Users\rchrd\Documents\Python\Text2speech.py ""--lang=en"" """ & txt & """"
                    End If
                    Call Shell(prog, vbNormalFocus)
                    Exit For
                End If
            End If
        End If
    Next Rng
    
    Call Timer
End Sub

Sub DisableTimer()
    On Error Resume Next
    Application.OnTime EarliestTime : = CountDown, Procedure : = "Reset", Schedule : = False
End Sub

Function CheckDay(Rng As Integer, Day As String) As Boolean

    Dim shtRng As Range
    Dim addRng As String

    Set shtRng = Range("G" + Mid(Str(Rng), 2), "M" + Mid(Str(Rng), 2))

    On Error Resume Next

    For Each cb In ActiveSheet.CheckBoxes
        If cb.Value = 1 Then
            addRng = ActiveSheet.Shapes(cb.Name).TopLeftCell.Address

            Select Case addRng

                Case Range("G" + Mid(Str(Rng), 2)).Address
                    If Day = "Sunday" Then
                        CheckDay = True
                        Exit For
                    Else
                        CheckDay = False
                    End If

                Case Range("H" + Mid(Str(Rng), 2)).Address
                    If Day = "Monday" Then
                        CheckDay = True
                        Exit For
                    Else
                        CheckDay = False
                    End If

                Case Range("I" + Mid(Str(Rng), 2)).Address
                    If Day = "Tuesday" Then
                        CheckDay = True
                        Exit For
                    Else
                        CheckDay = False
                    End If

                Case Range("J" + Mid(Str(Rng), 2)).Address
                    If Day = "Wednesday" Then
                        CheckDay = True
                        Exit For
                    Else
                        CheckDay = False
                    End If

                Case Range("K" + Mid(Str(Rng), 2)).Address
                    If Day = "Thursday" Then
                        CheckDay = True
                        Exit For
                    Else
                        CheckDay = False
                    End If

                Case Range("L" + Mid(Str(Rng), 2)).Address
                    If Day = "Friday" Then
                        CheckDay = True
                        Exit For
                    Else
                        CheckDay = False
                    End If

                Case Range("M" + Mid(Str(Rng), 2)).Address
                    If Day = "Saturday" Then
                        CheckDay = True
                        Exit For
                    Else
                        CheckDay = False
                    End If

            End Select
        End If
    Next

End Function