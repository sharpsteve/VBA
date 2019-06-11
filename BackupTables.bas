Option Compare Database

' Function to copy a table appending mm-dd-yy to the copy's name
Function CopyTable(strTable As String)
On Error GoTo CopyTable_Err
    
    Dim TableExist As Boolean
    Dim tdf As TableDef
   
    For Each tdf In CurrentDb.TableDefs
        If tdf.Name = strTable Then
            TableExist = True
            Exit For
        End If
    Next tdf
    If TableExist Then
        DoCmd.CopyObject "", strTable & "-" & Format(Date, "mm-dd-yy"), acTable, strTable
    Else
        MsgBox "Table does not exist"
    End If
    
CopyTable_Exit:
    Exit Function
CopyTable_Err:
    MsgBox Err.Number & "-" & Err.Description
    Resume CopyTable_Exit
End Function

' Function to opy the specified tables
Function BackupMainTables()
On Error GoTo CopyTable_Err
    
    Call CopyTable("") ' Add table name
    
CopyTable_Exit:
    Exit Function
CopyTable_Err:
    MsgBox Err.Number & "-" & Err.Description
    Resume CopyTable_Exit
End Function

