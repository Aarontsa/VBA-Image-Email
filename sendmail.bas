Sub ScheduleandSendmail()
Dim Outlook As Object
Dim newEmail As Object
Dim MyCheck As Boolean

Dim peoplegroup As String

Set Outlook = CreateObject("Outlook.Application")
Set newEmail = Outlook.CreateItem(0)

'Call CopyTable
Sheet4.Activate
    ActiveSheet.Range("A1:K40").Copy

With newEmail
    .To = Sheet5.Cells(2, "M").Value '"tan@gmail.com.my" '
    .CC = "tan@gmail.com.my;low@gmail.com.my"
    .Subject = "[SPI Graph]- " & Sheet3.Cells(2, "A").Value & "'s Graph on " & Sheet3.Cells(2, "CF").Value
    .Display ' dan di display..
    Dim xInspect As Object
    Dim pageEditor As Object
    
    Debug.Print "send Email"
    Set xInspect = newEmail.GetInspector
    Set pageEditor = xInspect.WordEditor
    
    ActiveSheet.Range("A1:L55").Copy
    
    pageEditor.Application.Selection.Start = Len(.Body)
    pageEditor.Application.Selection.End = pageEditor.Application.Selection.Start
    pageEditor.Application.Selection.Paste
   ' pageEditor.Application.Selection.Paste
   .Send
    Set pageEditor = Nothing
    Set xInspect = Nothing
    
End With


DoEvents
Set newEmail = Nothing
Set Outlook = Nothing
'ResetList

Debug.Print "Done4"
End Sub



