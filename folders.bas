Sub getfoldersname()
Dim objFSO As Object
Dim objFolder As Object
Dim objSubFolder As Object
Dim i As Integer
Dim j As Integer

Sheet5.Activate
Range("A2").Select

If Not ActiveCell.ListObject Is Nothing Then
On Error Resume Next
    ActiveCell.ListObject.DataBodyRange.Delete ' empty table content
End If

'Create an instance of the FileSystemObject
Set objFSO = CreateObject("Scripting.FileSystemObject")
'Get the folder object
Set objFolder = objFSO.GetFolder("D:\mysqldata\data") ' "C:\Users\tan\Desktop\SPC\mysqldata\data") '
i = 1
j = 1


'loops through each file in the directory and prints their names and path
For Each objSubFolder In objFolder.subfolders
'Debug.Print Mid(objSubFolder.Name, 16, 8)
'Debug.Print Format(Now, "yyyymmdd")
    'If Mid(objSubFolder.Name, 16, 8) = Sheet5.Cells(2, "D").Value Then 'Format(Now, "yyyymmdd") Then 'mid of file name
    If Left(objSubFolder.Name, 15) = Sheet5.Cells(11, "D").Value Then 'Format(Now, "yyyymmdd") Then ' left of file name
    'print folder name
     Sheet5.Cells(j + 1, 1) = objSubFolder.Name
     Sheet5.Cells(j + 1, 2) = Mid(objSubFolder.Name, 16, 8)
    'print folder path
     'Sheet5.Cells(i + 1, 2) = objSubFolder.Path
     j = j + 1
     'End If
     End If
    
    i = i + 1
Next objSubFolder

Debug.Print "get folder name list"

End Sub
'https://software-solutions-online.com/list-files-and-folders-in-a-directory/
