Attribute VB_Name = "Module1"
Sub Export_ws()


Dim wb As Workbook
    Dim ws As Worksheet
    Dim SaveToFolder As String
    Dim FileName As String
    Dim ExportedWorksheets As String
    
    'Set the folder to save the exported worksheets
    SaveToFolder = "C:\Users\keskilin\Path\"
    
    'Set the workbook to export worksheets from
    Set wb = ThisWorkbook
    
    'Loop through each worksheet in the workbook
    For Each ws In wb.Worksheets
        'Check if the worksheet is one of the worksheets to be exported
        If ws.Name = "wsname1" Or ws.Name = "wsname2" Then 'Replace "Sheet1" and "Sheet2" with the names of the worksheets to be exported
       
            'add more Or ws.name = " " ' as needed
        
        
            'Set the file name to be the original file name and the worksheet name
            FileName = wb.Name & " - " & ws.Name & ".xlsx"
            'Check if the file already exists in the save folder
            If Len(Dir(SaveToFolder & FileName)) > 0 Then
                'If the file exists, save on top of it
                ws.Copy
                Application.DisplayAlerts = False
                ActiveWorkbook.SaveAs SaveToFolder & FileName, FileFormat:=xlOpenXMLWorkbook
                Application.DisplayAlerts = True
                ActiveWorkbook.Close
            Else
                'If the file doesn't exist, save a new file
                ws.Copy
                Application.DisplayAlerts = False
                ActiveWorkbook.SaveAs SaveToFolder & FileName, FileFormat:=xlOpenXMLWorkbook
                Application.DisplayAlerts = True
                ActiveWorkbook.Close
            End If
            'Add the exported worksheet name to the ExportedWorksheets string
            ExportedWorksheets = ExportedWorksheets & ws.Name & ", "
        End If
    Next ws
    
    'Display a message box indicating which worksheets have been exported
    MsgBox "The following worksheets have been exported to " & SaveToFolder & ":" & vbNewLine & ExportedWorksheets & ".", vbInformation, "Worksheets Exported"

End Sub
