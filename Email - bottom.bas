Sub Email_text() 'content add text to email send out


 Dim LastRow As Long


    Sheets("email_content").Select
     
    With ActiveSheet
    LastRow = Sheet5.Cells(7, 6).Value 'table
    Debug.Print LastRow & "fwedfawdd"
    'LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
    End With
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Dear Sir/Madam,"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "Here is the " & Sheet3.Cells(2, "A").Value & "'s SPI graph on " & Sheet3.Cells(2, "CF").Value & ". Please see the graph below for details."

    Debug.Print "Email_bottom"
    'Range("A" & LastRow + 4).Select
    'Range("A32").Select
    Range("A" & LastRow + 2).Select
    ActiveCell.FormulaR1C1 = "Note: If you do not receive this email every Monday. Please send an email to infotech@qdos.com.my for inquiries."
    'Range("A34").Select
    Range("A" & LastRow + 4).Select
    ActiveCell.FormulaR1C1 = "Thank You."
    'Range("A36").Select
    Range("A" & LastRow + 6).Select
    ActiveCell.FormulaR1C1 = "IT"
    'Range("A38").Select
    Range("A" & LastRow + 8).Select
    ActiveCell.FormulaR1C1 = _
        "*** This is an automatically generated email, please do not reply ***"
    'Range("A40").Select
    Range("A" & LastRow + 10).Select
    ActiveCell.FormulaR1C1 = _
        "© Copyright 2022. The information contained within this self generated email are proprietary and confidential. No Part of this self generated email may be disclosed in any manner to a third party without prior written consent."


End Sub







