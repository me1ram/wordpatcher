Sub ReplaceVariableInMultiFiles()
    Dim folderPath As String
    Dim fileName As String
    Dim variablesToReplace() As Variant
    Dim newValues() As Variant
    Dim doc As Document
    Dim findRange As Range
    Dim i As Integer

    ' Set folder path for documents
    folderPath = "C:\Users\Meiram\Documents\orders\"
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    ' Variables and their new values
    variablesToReplace = Array("{{FIO}}", "{{ADDRESS}}", "{{IIN}}")
    newValues = Array("Abai Abayev Abayevich", "Algabas Kokpar 20", "01920931020912")

    ' Loop through all .docx files in the folder
    fileName = Dir(folderPath & "*.docx")
    If fileName = "" Then
        MsgBox "No .docx files found in the specified folder.", vbInformation
        Exit Sub
    End If

    Do While fileName <> ""
        On Error Resume Next
        Set doc = Documents.Open(folderPath & fileName, ReadOnly:=False)
        On Error GoTo 0

        If Not doc Is Nothing Then
            ' Loop through each variable to replace
            For i = LBound(variablesToReplace) To UBound(variablesToReplace)
                Set findRange = doc.Content
                With findRange.Find
                    .Text = variablesToReplace(i)
                    .Replacement.Text = newValues(i)
                    .ClearFormatting
                    .Replacement.ClearFormatting
                    .Execute Replace:=wdReplaceAll
                End With
            Next i

            ' Save and close the document
            doc.Save
            doc.Close SaveChanges:=wdSaveChanges
        End If

        ' Get the next file
        fileName = Dir
    Loop

    ' Confirmation message
    MsgBox "Replacement completed successfully!", vbInformation
End Sub
