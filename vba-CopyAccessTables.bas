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
    
    ' Loop through each record in the source table.
    Do While Not rsSource.EOF
        rsDest.AddNew
        
        ' Loop through each field in the current record.
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
                            Dim strFileName As String
                            
                            Set rsAttachSource = fldSource.Value
                            Set rsAttachDest = rsDest.Fields(fldSource.Name).Value
                            
                            ' Loop through each attachment record.
                            Do While Not rsAttachSource.EOF
                                rsAttachDest.AddNew
                                ' Extract just the file name from the full URL.
                                strFileName = ExtractFileName(rsAttachSource.Fields("FileName").Value)
                                ' Build the full temporary file path.
                                strTempFile = strFolderPath & "\" & strFileName
                                ' Save the attachment from the source record to the temp file.
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

' Extracts just the file name from a full URL or path.
' For example: "https://sharepoint-example.com/sites/site/Attachments/1/examplefile.pdf"
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
