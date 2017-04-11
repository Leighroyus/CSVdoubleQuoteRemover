Attribute VB_Name = "mdlCSVdoubleQuoteRemover"
'Remove double quotes and replaces commas with pipes
'Leigh Sullivan
'2017-03-17
'v0.1
'
'Notes: Requires reference to 'Microsoft Scripting Runtime' (scrrun.dll)

Private Sub subCSVdoubleQuoteRemover()

    Dim oFileSystem As Object, oTextStream As Object, oFile As Object, oRepairFile As Object
    Dim lRow As Long
    Dim sFolderPath As String, sLine As String, sNewLine As String
    Dim cSingleChar As String
    Dim iQuotesCounter As Integer
    Dim bQuotesClosed As Boolean
    
    bQuotesClosed = True
    
    Set oFileSystem = CreateObject("Scripting.FileSystemObject")
    
    'get folder
    sFolderPath = GetFolder()
    
    ' ensure we are working with a real folder
    If oFileSystem.FolderExists(sFolderPath) Then
    
        'cycle through each file in folder
        For Each oFile In oFileSystem.GetFolder(sFolderPath).Files
    
            'make sure it is the correct file type (either *.CSV or *.csv")
            If VBA.Strings.LCase(VBA.Strings.Right(oFile.Name, 4)) = ".csv" Then
            
                'set file to write out repaired file
                sFilePathAndName = sFolderPath & "\" & VBA.Strings.Left(oFile.Name, Len(oFile.Name) - 4) & "_piped" & VBA.Strings.Right(oFile.Name, 4)
                Set oRepairFile = oFileSystem.CreateTextFile(sFilePathAndName)
            
                'Get first line
                Set oTextStream = oFile.OpenAsTextStream(1, -2)
                
                Do While oTextStream.AtEndOfStream <> True
                
                    sLine = oTextStream.ReadLine
                    
                    'Step 1 - search through line and replace commas with pipes if not inside double quotes
                    For i = 1 To Len(sLine)
                        
                        'get each character
                        cSingleChar = Mid(sLine, i, 1)
                        
                        'if character is a double quote character then set flag to don't replace
                        If (cSingleChar = Chr(34)) Then
                            
                            iQuotesCounter = iQuotesCounter + 1
                            
                            'test whether it is an opening or closing quote
                            
                            'if iQuotesCounter equals 2 then we know that this is the closing quotation mark
                            If iQuotesCounter = 2 Then
                                
                                bQuotesClosed = True
                                'reset iQuotesCounter for next pair
                                iQuotesCounter = 0
                            
                            Else
                                bQuotesClosed = False
                            End If
                            
                        End If
                        
                        'if character is a comma and bDoubleQuotesDetected equals False then replace
                        If (cSingleChar = Chr(44) And bQuotesClosed = True) Then
                        
                            sNewLine = sNewLine + "|"
                        
                        Else
                        
                             sNewLine = sNewLine + cSingleChar
                        
                        End If
    
                    Next
                    
                    'Step 2 - remove all double quotes
                    sNewLine = Replace(sNewLine, Chr(34), "")
                    'Debug.Print (sNewLine)
                    
                    'Step 3 - write out to file
                    oRepairFile.WriteLine (sNewLine)
                    
                    'clear new line for next line
                    sNewLine = ""
                    
                    lRow = lRow + 1
                
                Loop
                   
                'close file streams
                oTextStream.Close
                oRepairFile.Close
            End If
            
        Next oFile
        
    Else
        'no folder was found
        MsgBox "Folder was not found"
    End If
    
    Set oTextStream = Nothing
    Set oFileSystem = Nothing
End Sub

Function GetFolder() As String
    Dim fldr As FileDialog
    Dim sItem As String
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Select Folder"
        .AllowMultiSelect = False
        .InitialFileName = Application.DefaultFilePath
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With
NextCode:
    GetFolder = sItem
    Set fldr = Nothing
End Function

Function GetFolderAndFile() As String
    Dim fldr As FileDialog
    Dim sItem As String
    Set fldr = Application.FileDialog(msoFileDialogOpen)
    With fldr
        .Title = "Select File"
        .AllowMultiSelect = False
        .InitialFileName = Application.DefaultFilePath
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With
NextCode:
    GetFolderAndFile = sItem
    Set fldr = Nothing
End Function

