' Developed by Vinamra Mathur

Sub GetFileList()

'Variable Declaration

    Dim xFSO As Object
    Dim xFolder As Object
    Dim xFile As Object
    Dim objOL As Object
    Dim Msg As Object
    Dim xPath As String
    Dim thisFile As String
    Dim i As Integer
    Dim j As Integer
    Dim lastrow As Long
    
    'Enter the path
    
    xPath = Sheets("UI").Range("D7")
    
    Set xFSO = CreateObject("Scripting.FileSystemObject")
    
    Set xFolder = xFSO.GetFolder(xPath)
    

    i = 1

       
    For Each xFile In xFolder.Files
    'Check for empty row
        Do
            i = i + 1
        Loop While Worksheets("Info").Cells(i, 1).Value <> ""
        
        Worksheets("Info").Cells(i, 1) = xPath
        Worksheets("Info").Cells(i, 2) = Left(xFile.Name, InStrRev(xFile.Name, ".") - 1)
        Worksheets("Info").Cells(i, 3) = Mid(xFile.Name, InStrRev(xFile.Name, ".") + 1)
        Worksheets("Info").Cells(i, 4) = Left(FileDateTime(xFile), InStrRev(FileDateTime(xFile), " ") - 1)

        
    Next
    Set Msg = Nothing
    
    'Switch to the output Worksheet
    
    Worksheets("Info").Visible = True
    Worksheets("Info").Activate
    



End Sub

