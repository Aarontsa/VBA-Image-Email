Sub connectsql(filenameforconnect As String)

'Call getfoldersnametest
'Call getfoldersname


'Debug.Print filenameforconnect

Dim Result As Object
Dim recordCount As Integer
Dim databasematch As String

Set cnn = CreateObject("ADODB.Connection") 'connect to database
Set objMyCmd = CreateObject("ADODB.Command")
Set rs = CreateObject("ADODB.Recordset")
databasematch = filenameforconnect 'Sheet5.Cells(2, "F").Value '"spc_p7el20_041_20220319*"  '"spc_p7el20_041_" & Format(Now, "yyyymmdd") & "081751921" '"spc_p7el20_041_20220319081751921"
Debug.Print databasematch
cnn.Open "DRIVER={MySQL ODBC 5.3 ANSI Driver};UID=root;PWD=1234;DATABASE=" & databasematch & ";PORT=3306;COLUMN_SIZE_S32=1;" 'spc_p7el20_041_20220421180719709
'cnn.Open "DRIVER={MySQL ODBC 5.3 ANSI Driver};UID=root;PWD=1234;DATABASE=spc_p7el20_041_20220422094254119;PORT=3306;COLUMN_SIZE_S32=1;" '

'Set and Excecute SQL Command'
Dim u As String
Set objMyCmd.ActiveConnection = cnn
'objMyCmd.CommandText = "SELECT `ModelName`,`board_id`,`line_name`,`part_name`,`barcode`,`Insp_st_time`,`insp_tack_time`,`is_pass`,`result`,`operatorID`,`squeegee`,`set_print`,`warpage`,`shrink_X`,`shrink_Y`,`avg_volume`,`avg_height`,`avg_area`,`avg_offsetx`,`avg_offsety`,`min_volume`,`max_volume`,`min_height`,`max_height`,`min_area`,`max_area`,`min_offsetx`,`max_offsetx`,`min_offsety`,`max_offsety`,`sigma_volume`,`sigma_height`,`sigma_area`,`sigma_offsetx`,`sigma_offsety`,`cpk_volume`,`cpk_height`,`cpk_area`,`cpk_offsetx`,`cpk_offsety`,`cp_volume`,`cp_height`,`cp_area`,`cp_offsetx`,`cp_offsety`,`r_volume`,`r_height`,`r_area`,`r_offsetx`,`r_offsety`,`ucl_volume`,`lcl_volume`,`ucl_area`,`lcl_area`,`ucl_height`,`lcl_height`,`ucl_offset`,`pathname`,`defect_count`,`userok_count`,`warn_count`,`skip_count`,`good_count`,`dispense_count`,`surface_kind`,`lane_kind` FROM `" & databasematch & "`.`insp_board`  WHERE `avg_height` <> '0' LIMIT 3" 'set active and inactive value
objMyCmd.CommandText = "SELECT `ModelName`,`board_id`,`line_name`,`part_name`,`barcode`,`Insp_st_time`,`insp_tack_time`,`is_pass`,`result`,`operatorID`,`squeegee`,`set_print`,`warpage`,`shrink_X`,`shrink_Y`,`avg_volume`,`avg_height`,`avg_area`,`avg_offsetx`,`avg_offsety`,`min_volume`,`max_volume`,`min_height`,`max_height`,`min_area`,`max_area`,`min_offsetx`,`max_offsetx`,`min_offsety`,`max_offsety`,`sigma_volume`,`sigma_height`,`sigma_area`,`sigma_offsetx`,`sigma_offsety`,`cpk_volume`,`cpk_height`,`cpk_area`,`cpk_offsetx`,`cpk_offsety`,`cp_volume`,`cp_height`,`cp_area`,`cp_offsetx`,`cp_offsety`,`r_volume`,`r_height`,`r_area`,`r_offsetx`,`r_offsety`,`ucl_volume`,`lcl_volume`,`ucl_area`,`lcl_area`,`ucl_height`,`lcl_height`,`ucl_offset`,`pathname`,`defect_count`,`userok_count`,`warn_count`,`skip_count`,`good_count`,`dispense_count`,`surface_kind`,`lane_kind` FROM `" & databasematch & "`.`insp_board`" _
                        & "WHERE `avg_height` <> '0' and  `Insp_st_time` >DATE_SUB(NOW(), INTERVAL 70 MINUTE) and `Insp_st_time` <DATE_SUB(NOW(), INTERVAL 10 MINUTE)" 'and  `Insp_st_time` >DATE_SUB(CURDATE(), INTERVAL 1 DAY) and `Insp_st_time` <CURDATE()

Debug.Print objMyCmd.CommandText
objMyCmd.CommandType = adCmdText
objMyCmd.Execute
        
'Open Recordset'
Set rs.Source = objMyCmd
rs.Open
For iCols = 0 To rs.Fields.Count - 1
Sheet3.Cells(1, iCols + 1).Value = rs.Fields(iCols).Name
Next

'Copy Data to Excel'
Sheet3.Range("A2").CopyFromRecordset rs

rs.Close
cnn.Close

Debug.Print "Query_ExpiredList end"

End Sub

Sub filtersql(filenameforconnect As String)

'Call getfoldersnametest
'Call getfoldersname


'Debug.Print filenameforconnect

Dim Result As Object
Dim recordCount As Integer
Dim databasematch As String

Set cnn = CreateObject("ADODB.Connection") 'connect to database
Set objMyCmd = CreateObject("ADODB.Command")
Set rs = CreateObject("ADODB.Recordset")
databasematch = filenameforconnect 'Sheet5.Cells(2, "F").Value '"spc_p7el20_041_20220319*"  '"spc_p7el20_041_" & Format(Now, "yyyymmdd") & "081751921" '"spc_p7el20_041_20220319081751921"
Debug.Print databasematch
cnn.Open "DRIVER={MySQL ODBC 5.3 ANSI Driver};UID=root;PWD=1234;DATABASE=" & databasematch & ";PORT=3306;COLUMN_SIZE_S32=1;" 'spc_p7el20_041_20220421180719709
'cnn.Open "DRIVER={MySQL ODBC 5.3 ANSI Driver};UID=root;PWD=1234;DATABASE=spc_p7el20_041_20220422094254119;PORT=3306;COLUMN_SIZE_S32=1;" '

'Set and Excecute SQL Command'
Dim u As String
Set objMyCmd.ActiveConnection = cnn
'objMyCmd.CommandText = "SELECT `ModelName`,`board_id`,`line_name`,`part_name`,`barcode`,`Insp_st_time`,`insp_tack_time`,`is_pass`,`result`,`operatorID`,`squeegee`,`set_print`,`warpage`,`shrink_X`,`shrink_Y`,`avg_volume`,`avg_height`,`avg_area`,`avg_offsetx`,`avg_offsety`,`min_volume`,`max_volume`,`min_height`,`max_height`,`min_area`,`max_area`,`min_offsetx`,`max_offsetx`,`min_offsety`,`max_offsety`,`sigma_volume`,`sigma_height`,`sigma_area`,`sigma_offsetx`,`sigma_offsety`,`cpk_volume`,`cpk_height`,`cpk_area`,`cpk_offsetx`,`cpk_offsety`,`cp_volume`,`cp_height`,`cp_area`,`cp_offsetx`,`cp_offsety`,`r_volume`,`r_height`,`r_area`,`r_offsetx`,`r_offsety`,`ucl_volume`,`lcl_volume`,`ucl_area`,`lcl_area`,`ucl_height`,`lcl_height`,`ucl_offset`,`pathname`,`defect_count`,`userok_count`,`warn_count`,`skip_count`,`good_count`,`dispense_count`,`surface_kind`,`lane_kind` FROM `" & databasematch & "`.`insp_board`  WHERE `avg_height` <> '0' LIMIT 3" 'set active and inactive value
objMyCmd.CommandText = "SELECT `ModelName`,`board_id`,`line_name`,`part_name`,`barcode`,`Insp_st_time`,`insp_tack_time`,`is_pass`,`result`,`operatorID`,`squeegee`,`set_print`,`warpage`,`shrink_X`,`shrink_Y`,`avg_volume`,`avg_height`,`avg_area`,`avg_offsetx`,`avg_offsety`,`min_volume`,`max_volume`,`min_height`,`max_height`,`min_area`,`max_area`,`min_offsetx`,`max_offsetx`,`min_offsety`,`max_offsety`,`sigma_volume`,`sigma_height`,`sigma_area`,`sigma_offsetx`,`sigma_offsety`,`cpk_volume`,`cpk_height`,`cpk_area`,`cpk_offsetx`,`cpk_offsety`,`cp_volume`,`cp_height`,`cp_area`,`cp_offsetx`,`cp_offsety`,`r_volume`,`r_height`,`r_area`,`r_offsetx`,`r_offsety`,`ucl_volume`,`lcl_volume`,`ucl_area`,`lcl_area`,`ucl_height`,`lcl_height`,`ucl_offset`,`pathname`,`defect_count`,`userok_count`,`warn_count`,`skip_count`,`good_count`,`dispense_count`,`surface_kind`,`lane_kind` FROM `" & databasematch & "`.`insp_board`" _
                        & "WHERE `avg_height` <> '0' and  `Insp_st_time` > '2022-08-23' ORDER BY `Insp_st_time` desc  LIMIT 13" 'and  `Insp_st_time` >DATE_SUB(CURDATE(), INTERVAL 1 DAY) and `Insp_st_time` <CURDATE()

Debug.Print objMyCmd.CommandText
objMyCmd.CommandType = adCmdText
objMyCmd.Execute
        
'Open Recordset'
Set rs.Source = objMyCmd
rs.Open
For iCols = 0 To rs.Fields.Count - 1
Sheet3.Cells(1, iCols + 1).Value = rs.Fields(iCols).Name
Next

'Copy Data to Excel'
Sheet3.Range("A2").CopyFromRecordset rs

rs.Close
cnn.Close

Debug.Print "Query_ExpiredList end"

End Sub

Sub Query_Sortlist()

Dim sSQLQry As String
Dim ReturnArray

Dim conn As New ADODB.Connection
Dim recset As New ADODB.Recordset

Dim DBPath As String, sConnect As String


DBPath = ThisWorkbook.FullName

sConnect = "Provider=MSDASQL.1;DSN=Excel Files;DBQ=" & DBPath & ";HDR=Yes';"

conn.Open sConnect


 sSQLSting = "SELECT `data$`.`date to text`, `data$`.avg_volume_mil, `data$`.avg_height_mil, `data$`.avg_area_mil, `data$`.avg_offsetx_mil, `data$`.avg_offsety_mil, `data$`.cpk_height  FROM `data$` `data$` WHERE (`data$`.`date to text` Is Not Null)"
 
 Debug.Print sSQLSting
    recset.Open sSQLSting, conn
    
    
    Sheet4.Cells(28, 1).CopyFromRecordset recset
    recset.Close


conn.Close

End Sub

Sub alldatabase()

Dim i As Integer
Dim j As Integer
Dim filenames As String
'Dim test As String
j = Sheet5.Cells(7, 4).Value

For i = 2 To j
filenames = Sheet5.Cells(i, 1).Value
'test = Sheet5.Cells(17, 6).Value
Call Automate(filenames)

'Debug.Print filenames & "1111111111111111111"
Next
End Sub



