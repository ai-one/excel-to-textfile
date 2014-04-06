VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} mainUserForm 
   Caption         =   "ai-one Excel-to-Textfile Utility"
   ClientHeight    =   7644
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   11028
   OleObjectBlob   =   "mainUserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "mainUserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Declare global variables that are passed across form pages.

' titleRow = Row number containing column titles
Public titleRow As Integer

' worksheetName = Name of worksheet containing data to process
Public worksheetName As String

' folderPath = Path of folder to save text files.
Public folderPath As String

' fileNameColumn = Name of column that contains the unique IDs to be used as filenames.
Public fileNameColumn As Integer

' Page 2 Back Button
Private Sub backButton_Page2_Click()

    Me.MultiPage1.Pages(2).Enabled = False
    Me.MultiPage1.Pages(1).Enabled = False
    Me.MultiPage1.Pages(0).Enabled = True
    
    Me.MultiPage1.Value = 0

End Sub

' Page 3 Back Button
Private Sub backButton_Page3_Click()

    Me.MultiPage1.Pages(2).Enabled = False
    Me.MultiPage1.Pages(1).Enabled = True
    Me.MultiPage1.Pages(0).Enabled = False
    
    Me.MultiPage1.Value = 1

End Sub

' Close the program if user selects the Cancel Button
Private Sub cancelButton_Page1_Click()

    Unload Me

End Sub

' Page 2 Cancel Button
Private Sub cancelButton_Page2_Click()

    Unload Me

End Sub

' Page 3 Cancel Button
Private Sub cancelButton_Page3_Click()

    Unload Me

End Sub


' Page 1 Next Button
Private Sub nextButton_Page1_Click()

    worksheetName = ""
    
    ' Verify that a folder has been selected.
    If Me.folderTextBox = "" Then
        MsgBox ("Please select a folder to save the text files.")
    
    Else
    
        folderPath = Me.folderTextBox.Value
        
        ' Verify that user has selected a row that contains column titles.
        If firstRowOptionButton.Value = True Then
            titleRow = 1
        Else
            Set RowRange = Application.InputBox(Prompt:="Select the row that contains the column titles", Type:=8)
            titleRow = RowRange.Rows(1).Row
        End If
        
        ' Get the selected Worksheet Name
        With Me.worksheetListBox
        
            For i = 0 To .ListCount - 1
                
                If .Selected(i) = True Then
                    worksheetName = .List(i)
                End If
        
            Next i
            
        End With
        
        ' Verify that user selected a worksheet
        If worksheets(worksheetName).UsedRange.Address = "$A$1" Then
            
            MsgBox ("Worksheet " & worksheetName & " is empty. Please select a different worksheet")
            
        ' All good, now populate the column names and continue to the next page.
        Else
            Call populateColumnForm(titleRow, worksheetName)
            
            Me.MultiPage1.Pages(2).Enabled = False
            Me.MultiPage1.Pages(1).Enabled = True
            Me.MultiPage1.Pages(0).Enabled = False
            
            Me.MultiPage1.Value = 1
        End If
    
    End If

End Sub

' Page 2 Next Button
Private Sub nextButton_Page2_Click()

    fileNameColumn = Me.IDColumn.Column(1)
    
    ' Check if there are any duplicates. True = no duplicates found.
    dupesResult = validateDuplicates(titleRow + 1, fileNameColumn, worksheetName)
    
    ' Check if there are any invalid characters. True = no invalid characters.
    charsResult = validateCharacters(titleRow + 1, fileNameColumn, worksheetName)
    
    ' If no dupes or invalid characters, then enable the start menu
    ' Otherwise keep the start menu disabled.
    If dupesResult And charsResult Then
        Me.verifyTextBox.Value = Me.verifyTextBox.Value & Chr(13) & Chr(13) & "Press Start button to generate files."
        
        Me.startButton_Page3.Enabled = True
        
    End If
    
    ' Send user to the next page to see validation results and to start the process
    Me.MultiPage1.Pages(2).Enabled = True
    Me.MultiPage1.Pages(1).Enabled = False
    Me.MultiPage1.Pages(0).Enabled = False
    
    Me.MultiPage1.Value = 2
    
End Sub

' Show select folder dialog box.
Private Sub selectFolderButton_Click()
    
    Dim newFolderPath As String
    newFolderPath = GetFolder(Me.folderTextBox.Value)
    
    Me.folderTextBox.Value = newFolderPath
    
End Sub


' Page 3 Start Button
Private Sub startButton_Page3_Click()

    ' Create a columnArray that contains the columns selected to be part of the text file(s)
    Dim columnString As String
    
    ' Get the list of columns selected by user.
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
    
    ' Execute the write text file sub which performs the writing of files.
    Call WriteTextFile(titleRow + 1, fileNameColumn, columnArray, folderPath)
    
End Sub

' Initialization sub for the form. Called when the form is first showed.
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
    
    For i = startRow To LastRow
    
        Filename = Trim(Cells(i, fileNameColumn).Value)

        filePath = folderPath + "\" + Filename + ".txt"
        
        
        CellData = ""
        
        Open filePath For Output As #2

        Dim N As Integer
        For N = LBound(columnArray) To UBound(columnArray)
        
            CellData = CellData + Trim(Cells(i, columnArray(N)).Value) + vbCrLf + vbCrLf
            
        Next N
        
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
