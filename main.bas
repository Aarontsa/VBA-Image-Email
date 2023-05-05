Sub Automate(filenames As String)

Dim filename As String
filename = filenames
'Dim test As String
'test = Sheet5.Cells(17, 6).Value
'Debug.Print "222222222"
'Debug.Print filename & "222222222"
Windows("spiconnect_v0.2_alldatabase_3-6-V4-22112022 -table.xlsm").Activate
    Sheet1.Select
    Range("A2").Select
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    
Sheet3.Activate
Range("A2").Select

If Not ActiveCell.ListObject Is Nothing Then
On Error Resume Next
    ActiveCell.ListObject.DataBodyRange.Delete ' empty table content
End If

Call connectsql(filename)
DoEvents

If Sheet3.Cells(2, "A").Value = "" Then

Debug.Print "Email end not records"


Else

Call filtersql(filename)
DoEvents

Sheet4.Activate
Range("A28").Select

If Not ActiveCell.ListObject Is Nothing Then
On Error Resume Next
    ActiveCell.ListObject.DataBodyRange.Delete ' empty table content
End If

Range("A30:A55").ClearContents


Call Query_Sortlist
DoEvents

    
Sheet3.Activate
Application.Wait (Now + TimeValue("0:00:02"))
DoEvents

Call Email_text
DoEvents

Sheet4.Activate
Application.Wait (Now + TimeValue("0:00:02"))
DoEvents

'On Error Resume Next
'If Sheet3.Cells(2, "A").Value <> "" Then
Call ScheduleandSendmail
Debug.Print filename & "1111111111111111111"

'DoEvents
'End If
Debug.Print "Email end"
End If






End Sub
