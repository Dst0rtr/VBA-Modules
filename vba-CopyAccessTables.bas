Option Compare Database
Option Explicit

' Copies records from a source table to a destination table.
' Parameters:
'   sourceTableName - The name of the source table (e.g., "tblA")
'   destTableName   - The name of the destination table (e.g., "tblB")
Public Sub CopyTableData(sourceTableName As String, destTableName As String)
    Dim db As DAO.Database
    Dim rsSource As DAO.Recordset, rsDest As DAO.Recordset
    Dim fldSource As DAO.Field
    
    Set db = CurrentDb
    Set rsSource = db.OpenRecordset(sourceTableName, dbOpenDynaset)
    Set rsDest = db.OpenRecordset(destTableName, dbOpenDynaset)
    
    Do While Not rsSource.EOF
        rsDest.AddNew
        
        For Each fldSource In rsSource.Fields
            ' Skip auto-number fields.
            If (fldSource.Attributes And dbAutoIncrField) = 0 Then
                If FieldExists(rsDest, fldSource.Name) Then
                    ' Handle attachment fields using SaveToFile and LoadFromFile.
                    If fldSource.Type = dbAttachment Then
                        If Not IsNull(fldSource.Value) Then
                            Dim rsAttachSource As DAO.Recordset2
                            Dim rsAttachDest As DAO.Recordset2
                            Dim strTempFile As String
                            
                            Set rsAttachSource = fldSource.Value
                            Set rsAttachDest = rsDest.Fields(fldSource.Name).Value
                            
                            Do While Not rsAttachSource.EOF
                                rsAttachDest.AddNew
                                ' Construct a temporary file path using the file name.
                                strTempFile = Environ("Temp") & "\" & rsAttachSource.Fields("FileName").Value
                                ' Save the attachment from the source record to a temp file.
                                rsAttachSource.Fields("FileData").SaveToFile strTempFile
                                ' Load the file into the destination attachment field.
                                rsAttachDest.Fields("FileData").LoadFromFile strTempFile
                                rsAttachDest.Update
                                ' Remove the temporary file.
                                Kill strTempFile
                                rsAttachSource.MoveNext
                            Loop
                            rsAttachSource.Close
                            rsAttachDest.Close
                        End If
                    Else
                        ' For non-attachment fields, copy the value directly.
                        rsDest.Fields(fldSource.Name).Value = fldSource.Value
                    End If
                End If
            End If
        Next fldSource
        
        rsDest.Update
        rsSource.MoveNext
    Loop
    
    rsDest.Close
    rsSource.Close
    Set db = Nothing
End Sub

' Helper function to check if a field exists in a recordset.
Public Function FieldExists(rs As DAO.Recordset, fldName As String) As Boolean
    On Error GoTo ErrHandler
    Dim dummy As DAO.Field
    Set dummy = rs.Fields(fldName)
    FieldExists = True
    Exit Function
ErrHandler:
    FieldExists = False
End Function
