VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formStep1 
   Caption         =   "ai-one Excel-to-Textfile Utility"
   ClientHeight    =   7644
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   11028
   OleObjectBlob   =   "formStep1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formStep1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public titleRow As Integer
Public worksheetName As String
Public folderPath As String
Public fileNameColumn As Integer
Public validationResult As Boolean


Private Sub cancelButton_Click()

Unload Me

End Sub

Private Sub columnsBackButton_Click()

Me.MultiPage1.Pages(1).Enabled = False
Me.MultiPage1.Pages(0).Enabled = True

Me.MultiPage1.Value = 0

End Sub

Private Sub columnsCancelButton_Click()
Unload Me
End Sub

Private Sub columnsNextButton_Click()

fileNameColumn = Me.IDColumn.Column(1)

dupesResult = validateDuplicates(titleRow + 1, fileNameColumn, worksheetName)

charsResult = validateCharacters(titleRow + 1, fileNameColumn, worksheetName)

If dupesResult And charsResult Then
    Me.verifyTextBox.Value = Me.verifyTextBox.Value & Chr(13) & Chr(13) & "Press Start button to generate files."
    
    Me.startButton.Enabled = True
    
End If
    Me.MultiPage1.Pages(2).Enabled = True
    Me.MultiPage1.Pages(1).Enabled = False
    Me.MultiPage1.Pages(0).Enabled = False
    
    Me.MultiPage1.Value = 2

End Sub

Private Sub Label2_Click()

End Sub

Private Sub nextButton_Click()

worksheetName = ""

If Me.folderTextBox = "" Then
    MsgBox ("Please select a folder to save the text files.")
Else

    folderPath = Me.folderTextBox.Value
    
    If firstRowOptionButton.Value = True Then
        titleRow = 1
    Else
        Set RowRange = Application.InputBox(Prompt:="Select the row that contains the column titles", Type:=8)
        titleRow = RowRange.Rows(1).Row
    End If
    
    
    With Me.worksheetListBox
    
        For i = 0 To .ListCount - 1
            
            If .Selected(i) = True Then
                worksheetName = .List(i)
            End If
    
        Next i
        
    End With
    
    If worksheets(worksheetName).UsedRange.Address = "$A$1" Then
        
        MsgBox ("Worksheet " & worksheetName & " is empty. Please select a different worksheet")
        
    Else
        Call populateColumnForm(titleRow, worksheetName)
        
        Me.MultiPage1.Pages(1).Enabled = True
        Me.MultiPage1.Pages(0).Enabled = False
        
        Me.MultiPage1.Value = 1
    End If

End If





End Sub

Private Sub selectFolderButton_Click()
Dim newFolderPath As String
newFolderPath = GetFolder(Me.folderTextBox.Value)

Me.folderTextBox.Value = newFolderPath
End Sub

Private Sub startButton_Click()

' Create a columnArray that contains the columns selected to be part of the text file(s)
Dim columnString As String

With Me.TextColumns
    For i = 0 To .ListCount - 1
        If .Selected(i) Then
            If Len(columnString) > 0 Then
                columnString = columnString + "," + .Column(1, i)
            Else
                columnString = columnString + .Column(1, i)
            End If
        End If
    Next i
End With

' Create an array of columns.
' The Split command returns a String array, so we will need to convert to an integer array later
columnArrayString = Split(columnString, ",")

columnArrayStringLength = UBound(columnArrayString)

Dim columnArray() As Integer
ReDim columnArray(columnArrayStringLength)

For c = 0 To columnArrayStringLength
    columnArray(c) = columnArrayString(c)
Next


Call WriteTextFile(titleRow + 1, fileNameColumn, columnArray, folderPath)
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
    
    Dim worksheetCount As Integer
    
    worksheetCount = Sheets.Count
    
    Dim worksheets() As String
    ReDim worksheets(worksheetCount - 1) As String
    
    For j = 1 To worksheetCount
               
        worksheets(j - 1) = Sheets(j).Name
    
    Next j
    
    Me.worksheetListBox.Clear
    
    For i = 0 To UBound(worksheets)
    
        Me.worksheetListBox.AddItem
        Me.worksheetListBox.List(i) = worksheets(i)

    Next i
    
    Me.worksheetListBox.Selected(0) = True
    
    Me.firstRowOptionButton.Value = 1
    
    ' Set the default folder path
    Dim folderPath As String
    folderPath = Application.ActiveWorkbook.Path
    Me.folderTextBox.Value = folderPath
End Sub

Private Sub verifyBackButton_Click()

Me.MultiPage1.Pages(2).Enabled = False
Me.MultiPage1.Pages(1).Enabled = True
Me.MultiPage1.Pages(0).Enabled = False

Me.MultiPage1.Value = 1

End Sub

Private Sub verifyCancelButton_Click()

Unload Me

End Sub

Private Sub populateColumnForm(titleRow As Integer, worksheetName As String)

    ' Set the number of columns for the ComboBox and ListBox to 2
    Me.IDColumn.ColumnCount = 2
    Me.TextColumns.ColumnCount = 2

    ' Set the width columns for the ComboBox and ListBox.
    ' The second column, which contains the column number, is hidden from the user.
    Me.IDColumn.ColumnWidths = ("60pt; 0pt")
    Me.TextColumns.ColumnWidths = ("60pt; 0pt")
    
    Dim ColumnCount As Integer
    
    'ColumnCount = ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Column
    ColumnCount = worksheets(worksheetName).UsedRange.SpecialCells(xlCellTypeLastCell).Column
    
    ' Create blank array of column names
    Dim columns() As String
    ReDim columns(ColumnCount - 1, 2) As String
    
    For j = 1 To ColumnCount
        
        cellText = Trim(Cells(titleRow, j).Value)
        
        columns(j - 1, 1) = cellText
        columns(j - 1, 2) = j
    
    Next j
    
    Me.IDColumn.Clear
    Me.TextColumns.Clear
    
    For i = 0 To UBound(columns)
    
        Me.IDColumn.AddItem
        Me.IDColumn.List(i, 0) = columns(i, 1)
        Me.IDColumn.List(i, 1) = columns(i, 2)
        
        Me.TextColumns.AddItem
        Me.TextColumns.List(i, 0) = columns(i, 1)
        Me.TextColumns.List(i, 1) = columns(i, 2)
    Next i
    
    With Me.TextColumns
    
        For i = 0 To .ListCount - 1
            
            .Selected(i) = True

        Next i
        
    End With
   
    ' Default the IDColumn combobox to the first list entry
    Me.IDColumn.ListIndex = 0

End Sub

Sub WriteTextFile(startRow As Integer, fileNameColumn As Integer, ByRef columnArray() As Integer, folderPath As String)

    Dim filePath As String
    Dim CellData As String
    Dim LastCol As Long
    Dim LastRow As Long
    
    LastCol = ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Column
    LastRow = ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row
    
    'SelectedFolder = GetFolder(Application.ActiveWorkbook.Path)
    
    For i = startRow To LastRow
    
        Filename = Trim(Cells(i, fileNameColumn).Value)

        filePath = folderPath + "\" + Filename + ".txt"
        
        
        CellData = ""
        
        Open filePath For Output As #2
        
        'For j = 2 To LastCol
        Dim N As Integer
        For N = LBound(columnArray) To UBound(columnArray)
        
            'MsgBox ("i =>" & i & " N =>" & columnArray(N) & " Cell =>" & Trim(Cells(i, columnArray(N)).Value))
            CellData = CellData + Trim(Cells(i, columnArray(N)).Value) + vbCrLf + vbCrLf
            
        Next N
        
        'Next j
        
        Print #2, CellData
        
        Close #2
    
    Next i
    
    MsgBox ("File generation completed.")
    
    Unload Me
    
    
End Sub

Private Function validateDuplicates(startRow As Integer, columnID As Integer, worksheetName As String) As Boolean

    Dim dict As New Scripting.Dictionary
    Dim dupes As New Scripting.Dictionary
    
    Dim LastRow As Long
    
    Dim duplicatesFound As Boolean
    duplicatesFound = False
  
    LastRow = worksheets(worksheetName).UsedRange.SpecialCells(xlCellTypeLastCell).Row
    
    Dim tempRowNumber As String
    
    For i = startRow To LastRow
    
        cellValue = Trim(Cells(i, columnID).Value)
        
        tempRowNumber = ""

        If Not dict.Exists(cellValue) Then
            dict.Add cellValue, CStr(i)
        Else
            If Not dupes.Exists(cellValue) Then
                tempRowNumber = dict.Item(cellValue)
                tempRowNumber = tempRowNumber & "," & CStr(i)
                dupes.Add cellValue, tempRowNumber
            Else
                tempRowNumber = dupes.Item(cellValue)
                tempRowNumber = tempRowNumber & "," & CStr(i)
                dupes.Remove (cellValue)
                dupes.Add cellValue, tempRowNumber
            End If
        End If

    Next i
    
    Dim verifyOutput As String
    verifyOutput = "Verifying that there are no duplicate IDs (which would create duplicate filenames)..."
    
    If dupes.Count > 0 Then
        verifyOutput = verifyOutput & Chr(13) & Chr(13) & "Duplicate IDs found in selected ID column." & Chr(13) & Chr(13) & "Please cancel program and ensure IDs are unique." & Chr(13) & Chr(13)
        verifyOutput = verifyOutput & "Duplicated IDs listed below." & Chr(13) & Chr(13) & "ID --> Row Numbers"
        fileNames = dupes.Keys
        
        For i = 0 To dupes.Count - 1
            verifyOutput = verifyOutput & Chr(13) & fileNames(i) & " --> " & dupes.Item(fileNames(i))
        Next
        
        Me.verifyTextBox.Value = verifyOutput
        
        validateDuplicates = False
    Else
        Me.verifyTextBox.Value = verifyOutput & Chr(13) & Chr(13) & "No duplicates found.... Great!"
        validateDuplicates = True
    End If

End Function

Private Function validateCharacters(startRow As Integer, columnID As Integer, worksheetName As String) As Boolean

    Dim dict As New Scripting.Dictionary
    Dim badChars As New Scripting.Dictionary
    
    Dim LastRow As Long
    
    Dim badCharsFound As Boolean
    badCharsFound = False
    
    Dim badCharArray(9) As String
    badCharArray(0) = "<"
    badCharArray(1) = ">"
    badCharArray(2) = ":"
    badCharArray(3) = """"
    badCharArray(4) = "/"
    badCharArray(5) = "\"
    badCharArray(6) = "|"
    badCharArray(7) = "?"
    badCharArray(8) = "*"
    
    LastRow = worksheets(worksheetName).UsedRange.SpecialCells(xlCellTypeLastCell).Row
    
    Dim tempRowNumber As String
    
    For i = startRow To LastRow
    
        cellValue = Trim(Cells(i, columnID).Value)
        
        Dim pos As Integer
        pos = 0
        
        For bc = 0 To UBound(badCharArray) - 1
            
            pos = pos + InStr(cellValue, badCharArray(bc))
        
        Next bc
        
        If pos > 0 Then
            badChars.Add cellValue, i
        End If

    Next i
    
    Dim verifyOutput As String
    verifyOutput = "Verifying that there are no invalid characters in the IDs (which would cause the files to not be created)..."
    
    If badChars.Count > 0 Then
        verifyOutput = verifyOutput & Chr(13) & Chr(13) & "Invalid filename characters found in selected ID column." & Chr(13) & Chr(13) & "Please cancel program and ensure IDs do not contain invalid character." & Chr(13) & Chr(13)
        verifyOutput = verifyOutput & Chr(13) & Chr(13) & "The following characters are not allowed: < > : "" / \ | ? *" & Chr(13) & Chr(13)
        verifyOutput = verifyOutput & "IDs listed below." & Chr(13) & Chr(13) & "ID --> Row Numbers"
        fileNames = badChars.Keys
        
        For i = 0 To badChars.Count - 1
            verifyOutput = verifyOutput & Chr(13) & fileNames(i) & " --> " & badChars.Item(fileNames(i))
        Next
        
        Me.verifyTextBox.Value = Me.verifyTextBox.Value & Chr(13) & Chr(13) & verifyOutput
        
        validateCharacters = False
    Else
        Me.verifyTextBox.Value = Me.verifyTextBox.Value & Chr(13) & Chr(13) & verifyOutput & Chr(13) & Chr(13) & "No invalid characters found.... Great!"
        
        validateCharacters = True
    End If

End Function

Function GetFolder(strPath As String) As String
Dim fldr As FileDialog
Dim sItem As String
Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
With fldr
    .Title = "Select a Folder To Save Text Files"
    .AllowMultiSelect = False
    .InitialFileName = strPath
    If .Show <> -1 Then GoTo NextCode
    sItem = .SelectedItems(1)
End With
NextCode:
GetFolder = sItem
Set fldr = Nothing
End Function
