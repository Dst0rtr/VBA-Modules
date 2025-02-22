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
    Dim rsAttachmentSource As DAO.Recordset, rsAttachmentDest As DAO.Recordset
    
    Set db = CurrentDb
    ' Open the source and destination recordsets dynamically based on the passed table names.
    Set rsSource = db.OpenRecordset(sourceTableName, dbOpenDynaset)
    Set rsDest = db.OpenRecordset(destTableName, dbOpenDynaset)
    
    ' Loop through each record in the source table.
    Do While Not rsSource.EOF
        rsDest.AddNew
        
        ' Loop through each field in the current record.
        For Each fldSource In rsSource.Fields
            ' Skip auto-number fields (identified by the dbAutoIncrField attribute).
            If (fldSource.Attributes And dbAutoIncrField) = 0 Then
                ' Only process if a field with the same name exists in the destination.
                If FieldExists(rsDest, fldSource.Name) Then
                    ' Check if the field is of attachment type.
                    If fldSource.Type = dbAttachment Then
                        ' If the attachment field is not null, process each attachment.
                        If Not IsNull(fldSource.Value) Then
                            Set rsAttachmentSource = fldSource.Value
                            Set rsAttachmentDest = rsDest.Fields(fldSource.Name).Value
                            
                            ' Loop through each attachment record.
                            Do While Not rsAttachmentSource.EOF
                                rsAttachmentDest.AddNew
                                rsAttachmentDest("FileData") = rsAttachmentSource("FileData")
                                rsAttachmentDest("FileName") = rsAttachmentSource("FileName")
                                rsAttachmentDest("FileType") = rsAttachmentSource("FileType")
                                rsAttachmentDest.Update
                                rsAttachmentSource.MoveNext
                            Loop
                            
                            rsAttachmentSource.Close
                            rsAttachmentDest.Close
                        End If
                    Else
                        ' For non-attachment fields, simply copy the value.
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

' Helper function to determine if a field exists in a given recordset.
Public Function FieldExists(rs As DAO.Recordset, fldName As String) As Boolean
    On Error GoTo ErrHandler
    Dim dummy As DAO.Field
    Set dummy = rs.Fields(fldName)
    FieldExists = True
    Exit Function
ErrHandler:
    FieldExists = False
End Function
