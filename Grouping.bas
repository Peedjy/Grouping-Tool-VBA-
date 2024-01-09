Attribute VB_Name = "Grouping"
Sub LoopAllExcelFilesInFolder()
'PURPOSE: To loop through all Excel files in a user specified folder and perform a set task on them
'SOURCE: www.TheSpreadsheetGuru.com

Dim wb, wb2 As Workbook
Dim myPath As String
Dim myFile As String
Dim myExtension As String
Dim FldrPicker As FileDialog
Dim myValue As Variant

'Optimize Macro Speed
  Application.ScreenUpdating = False
  Application.EnableEvents = False
  Application.Calculation = xlCalculationManual
  Application.DisplayAlerts = False

'Retrieve Target Folder Path From User
  Set FldrPicker = Application.FileDialog(msoFileDialogFolderPicker)

    With FldrPicker
      .Title = "Select A Target Folder"
      .AllowMultiSelect = False
        If .Show <> -1 Then GoTo NextCode
        myPath = .SelectedItems(1) & "\"
    End With

'In Case of Cancel
NextCode:
  myPath = myPath
  If myPath = "" Then GoTo ResetSettings

'Target File Extension (must include wildcard "*")
  myExtension = "*.xls*"

'Target Path with Ending Extention
  myFile = Dir(myPath & myExtension)
  
'Choose number of columns
  myValue = InputBox("Enter column letter to copy too")

'Loop through each Excel file in folder
  Do While myFile <> ""
    'Set variable equal to opened workbook
      Set wb2 = ActiveWorkbook
      Set wb = Workbooks.Open(Filename:=myPath & myFile)
    
    'Ensure Workbook has opened before moving on to next line of code
      DoEvents
    
    'Copy Data
      LastRow = wb.Worksheets(1).Range("A1").CurrentRegion.Rows.Count
      CopyRange = "A1:" & myValue & LastRow
      wb.Worksheets(1).Range(CopyRange).Copy
      LastRow = wb2.Sheets(1).Range("A1").CurrentRegion.Rows.Count
      PasteRange = "A" & LastRow + 1
      wb2.Sheets(1).Activate
      
      If LastRow = 1 Then
        PasteRange = "A" & LastRow
      Else
        PasteRange = "A" & LastRow + 1
      End If
        
      ActiveSheet.Paste Destination:=Worksheets("Sheet1").Range(PasteRange)
    
    'Save and Close Workbook
      wb.Close SaveChanges:=False
      
    'Ensure Workbook has closed before moving on to next line of code
      DoEvents

    'Get next file name
      myFile = Dir
  Loop

'Message Box when tasks are completed
  MsgBox "Task Complete!"

ResetSettings:
  'Reset Macro Optimization Settings
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub
