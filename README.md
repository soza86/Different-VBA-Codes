# Different-VBA-Codes
#This file contains explanatios about different codes implemented in VBA

This code tracks any changes made to the available spreadsheets of a workbook (besides “Changes” sheet, where all the changes are recorded).

`Code for tracking changes in the Overview sheet
Private Sub Worksheet_Change(ByVal Target As Range)

`This condition checks if the changes are made to the “Changes” sheet. If true, then this procedure ends.
If ActiveSheet.Name = "Changes" Then Exit Sub

Application.EnableEvents = False

UserName = Environ("USERNAME")
NewVal = Target.Value
Application.Undo
oldVal = Target.Value

`Finds the last row of the spreadsheet “Changes”
lr = Sheets("Changes").Range("A" &Rows.Count).End(xlUp).Row + 1

`The following block of statements stores the changes the last available row of the sheet “Changes” 
Sheets("Changes").Range("A" &lr) = Now
Sheets("Changes").Range("B" &lr) = UserName
Sheets("Changes").Range("C" &lr) = ActiveSheet.Name
Sheets("Changes").Range("D" &lr) = Target.Address
Sheets("Changes").Range("E" &lr) = oldVal
Sheets("Changes").Range("F" &lr) = NewVal


Target = NewVal
Application.EnableEvents = True

End Sub


This code lets the user select a folder path


'Code for selecting the folder where the files to be extracted are located
    Dim fldr As FileDialog, strPath As String, directory As String, sItem As String

Application.EnableEvents = False
Application.ScreenUpdating = False
Application.DisplayAlerts = False

    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Select a Folder"
        .AllowMultiSelect = False
        .InitialFileName = strPath
    If .Show <> -1 Then GoToNextCode
directory = .SelectedItems(1) & "\"
    End With
NextCode:
GetFolder = sItem
    Set fldr = Nothing

ActiveSheet.Range("G18").Value = directory

Sheets("HomePage").Select
Application.EnableEvents = True
Application.ScreenUpdating = True
Application.DisplayAlerts = True

This code clears the folder path, which was selected using the code for selecting folder path

'Code for deleting/clearing folder path
ActiveSheet.Range("G18").Select
Selection.ClearContents
ActiveSheet.Range("C19").Select


This code extracts sheets with the same name from different excel files. It can be used for creating summary workbooks.
Just change in the following line the name of the sheet you want to be extracted (source sheets) and the sheet name in the destination file where the source sheets will be placed after that (coloured in red)  
Sheets("Source_Sheet").Copy before:=Workbooks("FCET.xlsm").Sheets("Destination_Sheet")

'Code for extracting Test Sheets
    Dim directory As String, fileName As String, sheet As Worksheet, total As Integer

Application.EnableEvents = False
Application.ScreenUpdating = False
Application.DisplayAlerts = False

directory = ActiveSheet.Range("G18").Value
fileName = Dir(directory & "*.xl??")

    Do WhilefileName<> ""
Workbooks.Open (directory &fileName)
Sheets("Test").Copy before:=Workbooks("FCET.xlsm").Sheets("HomePage")
Sheets(Sheets.Count - 1).Name = Replace(fileName, ".xlsm", "")
Sheets(Sheets.Count - 1).Tab.ColorIndex = xlNone
Workbooks(fileName).Close
fileName = Dir()
    Loop

Sheets("HomePage").Select
Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.EnableEvents = True
This code deletes all sheets (besides the HomePage sheet) from the workbook 

'Code for deleting all sheets
    For Each wsIn Sheets
Application.DisplayAlerts = False
        If ws.Name<> "HomePage" Then ws.Delete
    Next

Application.DisplayAlerts = True


This code allows the user to delete a specific sheet from the workbook. If it is sound, it deletes it and shows a message to the user. Else, it shows a message that it was not found

'Code for deleting specific sheet
Dim sht As Variant
    Dim shstate As Boolean
sht = InputBox("Enter the sheet's name you want to delete:")

    For Each wsIn Sheets
Application.DisplayAlerts = False

        If ws.Name = sht Then
ws.Delete
shstate = True
        End If
    Next

    If shstate = True Then
MsgBox "Sheet deleted"
    Else
MsgBox "Sheet not found"
    End If

Application.DisplayAlerts = True

This code sorts all sheets in a given workbook

'Code for sorting sheets
Application.EnableEvents = False
Application.ScreenUpdating = False
Application.DisplayAlerts = False

    Dim sCount As Integer, i As Integer, j As Integer
sCount = Worksheets.Count
    If sCount = 1 Then Exit Sub
    For i = 1 TosCount - 1
    For j = i + 1 TosCount
        If Val(Worksheets(j).Name) < Val(Worksheets(i).Name) Then
Worksheets(j).Move before:=Worksheets(i)
    End If
    Next j
    Next i

Sheets("HomePage").Select
Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.EnableEvents = True

This code saves the workbook as new, macro-free, Excel file

'Code for saving the file as macro-free Excel file
fName = Application.GetSaveAsFilename(FileFilter:="Macro-Free Excel File (*.xlsx), *.xlsx")
    If fName = False Then
        Exit Sub
    Else
ThisWorkbook.SaveAsfileName:=fName, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    End If


