Attribute VB_Name = "Module1"
Sub convert2longformat()
'This is a VBA procdeure to convert a wide format to a long format
Dim Rng As Range 'The name of the fixed range to be repeated in the loop
Dim strC, NextLetter, PreviousLetter As String
Dim LastLine, LastLine1, d, LastC, i As Long


LastLine = Range("A" & Rows.Count).End(xlUp).Row - 1
LastC = Cells("1", Columns.Count).End(xlToLeft).Column
Application.ScreenUpdating = False
'Ask User to give the starting column
strC = InputBox("Enter the name of the start column that you want to be a categorical variable e.g E,C etc")
NextLetter = Chr(Asc(strC) + 1)
PreviousLetter = Chr(Asc(strC) - 1)


Set Rng = Range(Range("A2"), Range(PreviousLetter & Rows.Count).End(xlUp))




Columns(strC).Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range(strC & 1).Value = "DoW"


For i = 0 To Asc(Chr(Asc("A") + LastC)) - Asc(strC) - 1
Rng.Select
Application.CutCopyMode = False
Selection.Copy
d = 2 + i * LastLine

Range("A" & d).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
:=False, Transpose:=False
Range(NextLetter & 1).Offset(0, i).Select
Selection.Copy
Range(NextLetter & 1).Offset(d - 1, -1).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
:=False, Transpose:=False
LastLine1 = Range("A" & Rows.Count).End(xlUp).Row
ActiveSheet.Paste
Range(strC & d & ":" & strC & LastLine1).Select
Application.CutCopyMode = False
Selection.FillDown


'============================================================
'Values in one column-Numeric Variable

Range(NextLetter & 2).Offset(0, i).Resize(LastLine).Select
Selection.Copy
Range(NextLetter & 2).Offset(d - 2, 0).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
:=False, Transpose:=False





Next i


'Give a name to the numeric variable
Range(NextLetter & 1).Value = "Value"

'Clear the rest of the staff of wide format
Range(Chr(Asc(strC) + 2) & ":" & Chr(Asc("A") + LastC)).ClearContents








Application.ScreenUpdating = True






End Sub

