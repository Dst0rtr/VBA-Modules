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
        Dim sVal As String
        For Each fldSource In rsSource.Fields
            ' Skip auto-number fields.
            If (fldSource.Attributes And dbAutoIncrField) = 0 Then
                If FieldExists(rsDest, fldSource.Name) Then
                    ' Only process non-attachment fields.
                    If fldSource.Type <> dbAttachment Then
                        Set destFld = rsDest.Fields(fldSource.Name)
                        Select Case fldSource.Type
                            Case dbText
                                ' For short text fields, truncate if needed.
                                sVal = Nz(fldSource.Value, "")
                                If Len(sVal) > destFld.Size Then
                                    sVal = Left(sVal, destFld.Size)
                                End If
                                destFld.Value = sVal
                            Case dbMemo
                                ' For long text fields, simply assign the full value.
                                destFld.Value = fldSource.Value
                            Case Else
                                destFld.Value = fldSource.Value
                        End Select
                    End If
                End If
            End If
        Next fldSource
        
        ' Commit the new record so attachments can be added.
        rsDest.Update
        rsDest.Bookmark = rsDest.LastModified
        
        ' Second pass: Process attachment fields.
        For Each fldSource In rsSource.Fields
            If (fldSource.Attributes And dbAutoIncrField) = 0 Then
                If FieldExists(rsDest, fldSource.Name) Then
                    If fldSource.Type = dbAttachment Then
                        If Not IsNull(fldSource.Value) Then
                            ' Put the parent record in edit mode for attachment updates.
                            rsDest.Edit
                            
                            Dim rsAttachSource As DAO.Recordset2
                            Dim rsAttachDest As DAO.Recordset2
                            Dim strTempFile As String
                            Dim strFileName As String
                            
                            Set rsAttachSource = fldSource.Value
                            Set rsAttachDest = rsDest.Fields(fldSource.Name).Value
                            
                            Do While Not rsAttachSource.EOF
                                rsAttachDest.AddNew
                                ' Extract the file name from the full SharePoint URL.
                                strFileName = ExtractFileName(rsAttachSource.Fields("FileName").Value)
                                ' Build the full temporary file path.
                                strTempFile = strFolderPath & "\" & strFileName
                                ' Save the source attachment to the temporary file.
                                rsAttachSource.Fields("FileData").SaveToFile strTempFile
                                ' Load the temporary file into the destination attachment.
                                rsAttachDest.Fields("FileData").LoadFromFile strTempFile
                                rsAttachDest.Update
                                ' Delete the temporary file.
                                Kill strTempFile
                                rsAttachSource.MoveNext
                            Loop
                            
                            rsAttachSource.Close
                            rsAttachDest.Close
                            
                            ' Commit the attachment updates.
                            rsDest.Update
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
