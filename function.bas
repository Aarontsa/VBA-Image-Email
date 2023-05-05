Function getLastRow(listobj As ListObject) As Integer
LastRow2 = listobj.ListColumns(1).Range.Rows.Count '<-- last row in Column A in your Table
lastrow1 = listobj.ListColumns(1).Range(LastRow2, 1).End(xlUp).Row
getLastRow = LastRow2 + lastrow1 - 1
'Debug.Print getLastRow
End Function
