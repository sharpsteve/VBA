' If the a table back up exists (has the table name + "-mm-dd-yy")
' for today, delete existing table and copy from passed database file
Function GetNewTableWithData(strTable As String, strDatabase As String)
On Error GoTo CopyTable_Err
    Dim BackupTableName As String
    Dim TableExist As Boolean
    Dim tdf As TableDef
   
    BackupTableName = strTable & "-" & Format(Date, "mm-dd-yy")
   
    For Each tdf In CurrentDb.TableDefs
        If tdf.Name = BackupTableName Then
            TableExist = True
            Exit For
        End If
    Next tdf
    
    If TableExist Then
        DoCmd.DeleteObject acTable, strTable
        DoCmd.TransferDatabase acImport, "Microsoft Access", _
          strDatabase, acTable, strTable, strTable
    End If

CopyTable_Exit:
    Exit Function
CopyTable_Err:
    MsgBox Err.Number & "-" & Err.Description & " (" & TableName & ") "
    Resume CopyTable_Exit
End Function


