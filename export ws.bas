Attribute VB_Name = "Module1"
Sub Export_FVs()


    ' Set the path for the export folder
    Dim exportPath As String
    exportPath = "C:\Users\keskilin\Documents\_Motor Fuel Tax\Return Practice\_Practice\_FVs\"
    
    ' Loop through each worksheet to export
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
    
        ' Check if the worksheet name should be exported
        If ws.Name = "FV60" Or ws.Name = "FV65" Then
       
            ' Create the file name for the exported worksheet
            Dim fileName As String
            fileName = ThisWorkbook.Name & " - " & ws.Name & ".xlsx"
            
            ' Check if the file already exists in the designated folder
            If Len(Dir(exportPath & fileName)) > 0 Then
                ' If the file already exists, delete it before saving the new file
                Kill exportPath & fileName
            End If
            
            ' Export the worksheet to the designated folder
            ws.Copy
            ActiveWorkbook.SaveAs fileName:=exportPath & fileName, FileFormat:=xlOpenXMLWorkbook
            ActiveWorkbook.Close SaveChanges:=False
            
             Application.ScreenUpdating = False
        End If
        
    Next ws
    


End Sub
