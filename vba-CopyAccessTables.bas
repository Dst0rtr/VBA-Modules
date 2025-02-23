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
    
    ' Determine the temporary folder path (subfolder "temp" in the Access DB directory)
    Dim strFolderPath As String
    strFolderPath = CurrentProject.Path & "\temp"
    If Dir(strFolderPath, vbDirectory) = "" Then
        MkDir strFolderPath
    End If
    
    Do While Not rsSource.EOF
        rsDest.AddNew
        
        ' First pass: Copy non-attachment fields.
        Dim destFld As DAO.Field
        For Each fldSource In rsSource.Fields
            ' Skip auto-number fields.
            If (fldSource.Attributes And dbAutoIncrField) = 0 Then
                If FieldExists(rsDest, fldSource.Name) Then
                    ' Only copy non-attachment fields.
                    If fldSource.Type <> dbAttachment Then
                        Set destFld = rsDest.Fields(fldSource.Name)
                        ' If it is a Short Text field, check length.
                        If fldSource.Type = dbText Then
                            Dim sVal As String
                            sVal = Nz(fldSource.Value, "")
                            If Len(sVal) > destFld.Size Then
                                sVal = Left(sVal, destFld.Size)
                            End If
                            destFld.Value = sVal
                        Else
                            destFld.Value = fldSource.Value
                        End If
                    End If
                End If
            End If
        Next fldSource
        
        ' Commit the new record so attachments can be added.
        rsDest.Update
        
        ' Second pass: Process attachment fields.
        For Each fldSource In rsSource.Fields
            If (fldSource.Attributes And dbAutoIncrField) = 0 Then
                If FieldExists(rsDest, fldSource.Name) Then
                    If fldSource.Type = dbAttachment Then
                        If Not IsNull(fldSource.Value) Then
                            Dim rsAttachSource As DAO.Recordset2
                            Dim rsAttachDest As DAO.Recordset2
                            Dim strTempFile As String
                            Dim strFileName As String
                            
                            Set rsAttachSource = fldSource.Value
                            Set rsAttachDest = rsDest.Fields(fldSource.Name).Value
                            
                            Do While Not rsAttachSource.EOF
                                rsAttachDest.AddNew
                                ' Extract just the file name from the full SharePoint URL.
                                strFileName = ExtractFileName(rsAttachSource.Fields("FileName").Value)
                                ' Build the full temporary file path.
                                strTempFile = strFolderPath & "\" & strFileName
                                ' Save the attachment from the source to the temp file.
                                rsAttachSource.Fields("FileData").SaveToFile strTempFile
                                ' Load the file into the destination attachment field.
                                rsAttachDest.Fields("FileData").LoadFromFile strTempFile
                                rsAttachDest.Update
                                ' Delete the temporary file.
                                Kill strTempFile
                                rsAttachSource.MoveNext
                            Loop
                            
                            rsAttachSource.Close
                            rsAttachDest.Close
                        End If
                    End If
                End If
            End If
        Next fldSource
        
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

' Extracts just the file name from a full URL or path.
' For example, given:
'   "https://sharepoint-example.com/sites/site-example/list/examplelist/Attachments/1/examplefile.pdf"
' returns "examplefile.pdf".
Public Function ExtractFileName(ByVal fullUrl As String) As String
    Dim pos As Long
    pos = InStrRev(fullUrl, "/")
    If pos > 0 Then
        ExtractFileName = Mid(fullUrl, pos + 1)
    Else
        ExtractFileName = fullUrl
    End If
End Function
