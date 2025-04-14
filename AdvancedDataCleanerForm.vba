' AdvancedDataCleanerForm.vb
' This file contains the VBA UserForm code and layout for AdvancedDataCleanerForm.
' To use, create a UserForm in VBA Editor, add controls as described, and paste the code below.

' === UserForm Layout ===
' Name: AdvancedDataCleanerForm
' Controls:
' - CheckBox: chkBlankRows (Caption: "Remove Blank Rows")
' - CheckBox: chkDuplicates (Caption: "Remove Duplicates")
' - CheckBox: chkInvalidEmails (Caption: "Validate Emails")
' - CheckBox: chkTextStandardize (Caption: "Standardize Text")
' - ComboBox: cmbColumn (Caption above: "Numeric Column")
' - ComboBox: cmbRegexColumn (Caption above: "Regex Column")
' - ComboBox: cmbRefColumn (Caption above: "Reference Column")
' - ComboBox: cmbRegexType (Caption above: "Regex Type")
' - TextBox: txtMaxValue (Caption above: "Max Numeric Value")
' - TextBox: txtRegexPattern (Caption above: "Regex Pattern")
' - CommandButton: btnRun (Caption: "Run")
' - CommandButton: btnUndo (Caption: "Undo")
' - CommandButton: btnCancel (Caption: "Cancel")
' Suggested Layout:
' - Checkboxes aligned vertically on the left.
' - ComboBoxes and TextBoxes in two columns (Numeric and Regex) in the middle.
' - Buttons aligned horizontally at the bottom.

' === UserForm Code ===
Private Sub UserForm_Initialize()
    ' Populate column selectors
    Dim ws As Worksheet
    Dim lastCol As Long
    Set ws = ActiveSheet
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For i = 1 To lastCol
        cmbColumn.AddItem ws.Cells(1, i).Value
        cmbRegexColumn.AddItem ws.Cells(1, i).Value
        cmbRefColumn.AddItem ws.Cells(1, i).Value
    Next i
    cmbColumn.ListIndex = 0
    cmbRegexColumn.ListIndex = 0
    cmbRefColumn.ListIndex = 0
    
    ' Populate regex type selector
    With cmbRegexType
        .AddItem "Phone"
        .AddItem "Postal"
        .AddItem "Custom"
        .ListIndex = 0
    End With
    
    ' Default regex pattern
    txtRegexPattern.Value = "0#########"
End Sub

Private Sub cmbRegexType_Change()
    ' Update default pattern based on regex type
    Select Case cmbRegexType.Value
        Case "Phone"
            txtRegexPattern.Value = "0#########"
        Case "Postal"
            txtRegexPattern.Value = "######"
        Case "Custom"
            txtRegexPattern.Value = ""
    End Select
End Sub

Private Sub btnRun_Click()
    Dim ws As Worksheet, wsBackup As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim rng As Range
    Dim i As Long, j As Long
    Dim msg As String
    Dim errorsFound As Long, cellsChanged As Long
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
    
    ' Create backup for undo
    On Error Resume Next
    Set wsBackup = ThisWorkbook.Sheets("BackupData")
    If wsBackup Is Nothing Then
        Set wsBackup = ThisWorkbook.Sheets.Add
        wsBackup.Name = "BackupData"
        wsBackup.Visible = xlSheetHidden
    End If
    rng.Copy wsBackup.Range("A1")
    On Error GoTo 0
    
    ' 1. Remove blank rows
    If chkBlankRows.Value Then
        Dim rowsDeleted As Long
        For i = lastRow To 2 Step -1
            If Application.WorksheetFunction.CountA(ws.Rows(i)) = 0 Then
                ws.Rows(i).Delete
                rowsDeleted = rowsDeleted + 1
            End If
        Next i
        msg = msg & "- Removed " & rowsDeleted & " blank rows." & vbCrLf
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    End If
    
    ' 2. Remove duplicates
    If chkDuplicates.Value Then
        rng.RemoveDuplicates Columns:=Array(1, 2, 3), Header:=xlYes
        Dim dupRemoved As Long
        dupRemoved = lastRow - ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        msg = msg & "- Removed " & dupRemoved & " duplicate rows." & vbCrLf
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    End If
    
    ' 3. Standardize text
    If chkTextStandardize.Value Then
        For i = 2 To lastRow
            For j = 1 To lastCol
                If Len(ws.Cells(i, j).Value) > 0 Then
                    Dim origVal As String
                    origVal = ws.Cells(i, j).Value
                    ws.Cells(i, j).Value = Trim(UCase(ws.Cells(i, j).Value))
                    If ws.Cells(i, j).Value <> origVal Then cellsChanged = cellsChanged + 1
                End If
            Next j
        Next i
        msg = msg & "- Standardized " & cellsChanged & " cells (trimmed and uppercased)." & vbCrLf
    End If
    
    ' 4. Validate emails
    If chkInvalidEmails.Value Then
        errorsFound = 0
        For i = 2 To lastRow
            For j = 1 To lastCol
                If InStr(LCase(ws.Cells(1, j).Value), "email") > 0 Then
                    Dim email As String
                    email = ws.Cells(i, j).Value
                    If Len(email) > 0 And (InStr(email, "@") = 0 Or InStr(email, ".") = 0) Then
                        ws.Cells(i, j).Interior.Color = vbYellow
                        errorsFound = errorsFound + 1
                    End If
                End If
            Next j
        Next i
        msg = msg & "- Highlighted " & errorsFound & " invalid emails." & vbCrLf
    End If
    
    ' 5. Regex-like validation
    If Len(txtRegexPattern.Value) > 0 Then
        errorsFound = 0
        Dim regexCol As Long
        regexCol = cmbRegexColumn.ListIndex + 1
        For i = 2 To lastRow
            Dim val As String
            val = ws.Cells(i, regexCol).Value
            If Len(val) > 0 And Not val Like txtRegexPattern.Value Then
                ws.Cells(i, regexCol).Interior.Color = vbMagenta
                errorsFound = errorsFound + 1
            End If
        Next i
        msg = msg & "- Highlighted " & errorsFound & " invalid values in column " & cmbRegexColumn.Value & " (pattern: " & txtRegexPattern.Value & ")." & vbCrLf
    End If
    
    ' 6. Reference-based validation
    If cmbRefColumn.ListIndex >= 0 Then
        Dim refWs As Worksheet
        On Error Resume Next
        Set refWs = ThisWorkbook.Sheets("Reference")
        On Error GoTo 0
        If Not refWs Is Nothing Then
            Dim refLastRow As Long, refCol As Long
            refLastRow = refWs.Cells(refWs.Rows.Count, 1).End(xlUp).Row
            refCol = 1 ' Assume reference data in column A
            Dim refValues() As String
            ReDim refValues(1 To refLastRow - 1)
            For i = 2 To refLastRow
                refValues(i - 1) = UCase(refWs.Cells(i, refCol).Value)
            Next i
            
            errorsFound = 0
            Dim refColIndex As Long
            refColIndex = cmbRefColumn.ListIndex + 1
            For i = 2 To lastRow
                Dim inputVal As String
                inputVal = UCase(ws.Cells(i, refColIndex).Value)
                If Len(inputVal) > 0 Then
                    Dim found As Boolean
                    found = False
                    For Each refVal In refValues
                        If inputVal = refVal Then
                            found = True
                            Exit For
                        End If
                    Next refVal
                    If Not found Then
                        ws.Cells(i, refColIndex).Interior.Color = vbCyan
                        errorsFound = errorsFound + 1
                    End If
                End If
            Next i
            msg = msg & "- Highlighted " & errorsFound & " invalid entries in column " & cmbRefColumn.Value & " (not in Reference)." & vbCrLf
        End If
    End If
    
    ' 7. Numeric check
    If Len(txtMaxValue.Value) > 0 Then
        Dim maxVal As Double, colIndex As Long
        maxVal = CDbl(txtMaxValue.Value)
        colIndex = cmbColumn.ListIndex + 1
        errorsFound = 0
        For i = 2 To lastRow
            If IsNumeric(ws.Cells(i, colIndex).Value) Then
                If ws.Cells(i, colIndex).Value > maxVal Then
                    ws.Cells(i, colIndex).Interior.Color = vbRed
                    errorsFound = errorsFound + 1
                End If
            End If
        Next i
        msg = msg & "- Highlighted " & errorsFound & " values exceeding " & maxVal & " in column " & cmbColumn.Value & "." & vbCrLf
    End If
    
    ' 8. Summary
    If msg = "" Then msg = "No actions performed."
    MsgBox msg, vbInformation, "Advanced Cleaning Complete"
End Sub

Private Sub btnUndo_Click()
    Dim ws As Worksheet, wsBackup As Worksheet
    Set ws = ActiveSheet
    On Error Resume Next
    Set wsBackup = ThisWorkbook.Sheets("BackupData")
    On Error GoTo 0
    
    If Not wsBackup Is Nothing Then
        ws.Cells.Clear
        wsBackup.Cells.Copy ws.Range("A1")
        MsgBox "Data restored to previous state!", vbInformation
    Else
        MsgBox "No backup found!", vbExclamation
    End If
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub
