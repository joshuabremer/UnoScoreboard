'This is the code for the UnoScore.xls file.
'This is just for reference. The code is already packed into the XLS file.

Sub setup()

Sheets("UNO").Select

intNumberOfPlayers = InputBox("Enter Number Of Players:")
If intNumberOfPlayers <= 1 Or intNumberOfPlayers > 100 Then
MsgBox ("Bad Value")
intNumberOfPlayers = 0
Call setup
End If



intNumberOfPlayers = CInt(intNumberOfPlayers)
Dim Players(100)

    
For i = 1 To intNumberOfPlayers
    strName = "Player " & i & " Name:"
    Players(i) = InputBox(strName)
Next i

intNumRounds = InputBox("Enter Number Of Rounds")

Range("D2").Select

For i = 1 To intNumberOfPlayers
    ActiveCell.Value = Players(i)
    ActiveCell.Offset(0, 1).Select
Next i


ActiveCell.Offset(0, -1).Select
intEndColumn = ActiveCell.Column
strEndColumnName = ActiveCell.Address
strEndColumnName = Application.WorksheetFunction.Substitute(strEndColumnName, "$", "")
strEndColumnName = Left(strEndColumnName, 1)

Range("B8").Select
For i = 1 To intNumRounds
    ActiveCell.Value = i
    ActiveCell.Offset(1).Select
Next i

intEndRound = 7 + intNumRounds




Range("D3").FormulaR1C1 = "=SUM(R[5]C:R[" & 5 + intNumRounds - 1 & "]C)"
strRange = "R3C4:R3C" & 4 + intNumberOfPlayers - 1
Range("D4").FormulaR1C1 = "=MIN(" & strRange & ")-R[-1]C"
Range("D5").FormulaR1C1 = "=RANK(R[-2]C," & strRange & ",1)"

strRange = "D3:" & strEndColumnName & "5"
Range("D3:D5").Select
Selection.AutoFill Destination:=Range(strRange), Type:=xlFillDefault

strRange = "D7:" & strEndColumnName & "7"
Range(strRange).Merge

Range("D8").Select



End Sub
