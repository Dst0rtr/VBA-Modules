Option Compare Database
Option Explicit

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
    If Dir(strFolderPath, vbDirectory) = "" Then MkDir strFolderPath
    
    Do While Not rsSource.EOF
        rsDest.AddNew
        Dim destFld As DAO.Field
        Dim sVal As String
        
        ' First pass: Copy simple fields (skip attachments and multi-valued fields)
        For Each fldSource In rsSource.Fields
            ' Skip auto-number fields.
            If (fldSource.Attributes And dbAutoIncrField) = 0 Then
                If FieldExists(rsDest, fldSource.Name) Then
                    ' Skip attachment fields and multi-valued fields (type 104)
                    If fldSource.Type <> dbAttachment And fldSource.Type <> 104 Then
                        Set destFld = rsDest.Fields(fldSource.Name)
                        Select Case fldSource.Type
                            Case dbText
                                sVal = Nz(fldSource.Value, "")
                                If Len(sVal) > destFld.Size Then
                                    sVal = Left(sVal, destFld.Size)
                                End If
                                destFld.Value = sVal
                            Case dbMemo
                                destFld.Value = fldSource.Value
                            Case Else
                                destFld.Value = fldSource.Value
                        End Select
                    End If
                End If
            End If
        Next fldSource
        
        ' Commit the main record so that attachments and multi-valued fields can be updated.
        rsDest.Update
        rsDest.Bookmark = rsDest.LastModified
        
        ' Second pass: Process attachment and multi-valued fields.
        For Each fldSource In rsSource.Fields
            If (fldSource.Attributes And dbAutoIncrField) = 0 Then
                If FieldExists(rsDest, fldSource.Name) Then
                    ' Process attachments first
                    If fldSource.Type = dbAttachment Then
                        If Not IsNull(fldSource.Value) Then
                            rsDest.Edit
                            Dim rsAttachSource As DAO.Recordset2
                            Dim rsAttachDest As DAO.Recordset2
                            Dim strTempFile As String, strFileName As String
                            
                            Set rsAttachSource = fldSource.Value
                            Set rsAttachDest = rsDest.Fields(fldSource.Name).Value
                            
                            Do While Not rsAttachSource.EOF
                                rsAttachDest.AddNew
                                strFileName = ExtractFileName(rsAttachSource.Fields("FileName").Value)
                                strTempFile = strFolderPath & "\" & strFileName
                                rsAttachSource.Fields("FileData").SaveToFile strTempFile
                                rsAttachDest.Fields("FileData").LoadFromFile strTempFile
                                rsAttachDest.Update
                                Kill strTempFile
                                rsAttachSource.MoveNext
                            Loop
                            
                            rsAttachSource.Close
                            rsAttachDest.Close
                            rsDest.Update
                        End If
                        
                    ' Process multi-valued fields (type 104)
                    ElseIf fldSource.Type = 104 Then
                        If Not IsNull(fldSource.Value) Then
                            rsDest.Edit
                            Dim rsMVSource As DAO.Recordset2
                            Dim rsMVDest As DAO.Recordset2
                            
                            Set rsMVSource = fldSource.Value
                            Set rsMVDest = rsDest.Fields(fldSource.Name).Value
                            
                            Do While Not rsMVSource.EOF
                                rsMVDest.AddNew
                                ' Multi-valued fields typically use "Value" as the field name in the subrecordset.
                                rsMVDest.Fields("Value").Value = rsMVSource.Fields("Value").Value
                                rsMVDest.Update
                                rsMVSource.MoveNext
                            Loop
                            
                            rsMVSource.Close
                            rsMVDest.Close
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

' Helper function to determine if a field exists in a recordset.
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
' "https://sharepoint-example.com/sites/site-example/list/examplelist/Attachments/1/examplefile.pdf"
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
