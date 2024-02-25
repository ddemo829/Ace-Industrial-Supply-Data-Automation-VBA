Attribute VB_Name = "VT_2022"
Sub Table_2022()

    Dim twentwo As Worksheet
    Set twentwo = Worksheets("2022")
    
    Dim FormattedVT As Worksheet
    Set FormattedVT = Worksheets("FormattedVT2022")
       
    Dim myrange As Range
    Set myrange = Selection

    Dim rsum As Integer
    rsum = myrange.Rows.Count
    
    For k = 2 To rsum
        For i = 2 To 10 Step 2
            Dim lra, lrc, ve, vo, vevosum, namesum As Long
            lra = FormattedVT.Range("A" & Rows.Count).End(xlUp).Offset(1, 0).Row
            lrc = FormattedVT.Range("C" & Rows.Count).End(xlUp).Offset(1, 0).Row
            ve = myrange.Cells(k, i).Value
            vo = myrange.Cells(k, i).Offset(0, 1).Value
            vevosum = ve + vo
            namesum = myrange.Cells(k, 14)
            'Dates
            If vevosum > 0 Then
            myrange.Cells(1, i).Copy
            FormattedVT.Range("A" & lra).PasteSpecial (xlPasteFormulasAndNumberFormats)
            FormattedVT.Range("A" & lra).NumberFormat = "m/d/yyyy"
            FormattedVT.Range("A" & lra).Value = Left(FormattedVT.Range("A" & lra).Value, _
            InStrRev(FormattedVT.Range("A" & lra).Value, "/", -1)) & "2022"
            FormattedVT.Range("A" & lra & ":" & "A" & lra + vevosum - 1).FillDown
            End If
            'Verified and Void
            If ve > 0 Then
                If ve = 1 Then
                FormattedVT.Range("C" & lrc).Value = "Verified"
                Else
                FormattedVT.Range("C" & lrc).Value = "Verified"
                FormattedVT.Range("C" & lrc & ":" & "C" & lrc + ve - 1).FillDown
                    End If
            End If
            If vo > 0 Then
            Dim lrc2 As Long
            lrc2 = FormattedVT.Range("C" & Rows.Count).End(xlUp).Offset(1, 0).Row
                If vo = 1 Then
                FormattedVT.Range("C" & lrc2).Value = "Void"
                Else
                FormattedVT.Range("C" & lrc2).Value = "Void"
                FormattedVT.Range("C" & lrc2 & ":" & "C" & lrc2 + vo - 1).FillDown
                End If
            End If
            'Name
            If vevosum > 0 Then
            Dim lrd As Long
            lrd = FormattedVT.Range("D" & Rows.Count).End(xlUp).Offset(1, 0).Row
                If vevosum = 1 Then
                myrange.Cells(k, 1).Copy Destination:=FormattedVT.Range("D" & lrd)
                FormattedVT.Range("D" & lrd).ClearFormats
                FormattedVT.Range("D" & lrd).Value = UCase(FormattedVT.Range("D" & lrd).Value)
                Else
                myrange.Cells(k, 1).Copy Destination:=FormattedVT.Range("D" & lrd)
                FormattedVT.Range("D" & lrd).ClearFormats
                FormattedVT.Range("D" & lrd).Value = UCase(FormattedVT.Range("D" & lrd).Value)
                FormattedVT.Range("D" & lrd & ":" & "D" & lrd + vevosum - 1).FillDown
                End If
            End If
        Next
    Next
        
End Sub
Sub Pend_2022()

    'ActiveCell must be on penddate
    
    Dim twentwo As Worksheet
    Set twentwo = Worksheets("2022")
    
    Dim FormattedVT As Worksheet
    Set FormattedVT = Worksheets("FormattedVT2022")
    
    Dim LRval As Long
    LRval = FormattedVT.Range("C" & Rows.Count).End(xlUp).Offset(1, 0).Row
    
    Dim pendttl As Integer
    pendttl = ActiveCell.Value
    
    Dim penddate As String
    penddate = ActiveCell.End(xlToRight).Value
    
    FormattedVT.Range("C" & LRval).Value = "Pend"
    FormattedVT.Range("C" & LRval & ":" & "C" & LRval + pendttl - 1).FillDown
    FormattedVT.Range("D" & LRval).Value = "PEND"
    FormattedVT.Range("D" & LRval & ":" & "D" & LRval + pendttl - 1).FillDown

    FormattedVT.Range("A" & LRval).Value = Right(penddate, 5)
    FormattedVT.Range("A" & LRval).NumberFormat = "m/d/yyyy"
    If Right(FormattedVT.Range("A" & LRval).Value, 4) = 2023 Or 1900 Or 2022 Then
    FormattedVT.Range("A" & LRval).Value = Left(FormattedVT.Range("A" & LRval).Value, _
    InStrRev(FormattedVT.Range("A" & LRval).Value, "/", -1)) & "2022"
    FormattedVT.Range("A" & LRval & ":" & "A" & LRval + pendttl - 1).FillDown
    Else
    FormattedVT.Range("A" & LRval).Value = FormattedVT.Range("A" & LRval).Value & "/2022"
    FormattedVT.Range("A" & LRval & ":" & "A" & LRval + pendttl - 1).FillDown
    End If
    
End Sub
Sub VeriVoid_2022()

    'ActiveCell must be on far left verified column B
    
    Dim twentwo As Worksheet
    Set twentwo = Worksheets("2022")
    
    Dim FormattedVT As Worksheet
    Set FormattedVT = Worksheets("FormattedVT2022")
    
    Dim ve1, vo1, ve2, vo2, ve3, vo3, ve4, vo4, ve5, v05 As Integer
    ve1 = ActiveCell.Value
    vo1 = ActiveCell.Offset(0, 1).Value
    ve2 = ActiveCell.Offset(0, 2).Value
    vo2 = ActiveCell.Offset(0, 3).Value
    ve3 = ActiveCell.Offset(0, 4).Value
    vo3 = ActiveCell.Offset(0, 5).Value
    ve4 = ActiveCell.Offset(0, 6).Value
    vo4 = ActiveCell.Offset(0, 7).Value
    ve5 = ActiveCell.Offset(0, 8).Value
    vo5 = ActiveCell.Offset(0, 9).Value
    
    Dim LRrange1 As Range
    Set LRrange1 = FormattedVT.Range("C" & Rows.Count).End(xlUp).Offset(1, 0)
    For i = 1 To ve1
        LRrange1.Offset(i - 1, 0).Value = "Verified"
    Next i
    
    Dim LRrange2 As Range
    Set LRrange2 = FormattedVT.Range("C" & Rows.Count).End(xlUp).Offset(1, 0)
    For i = 1 To vo1
        LRrange2.Offset(i - 1, 0).Value = "Void"
    Next i
    
    Dim LRrange3 As Range
    Set LRrange3 = FormattedVT.Range("C" & Rows.Count).End(xlUp).Offset(1, 0)
    For i = 1 To ve2
        LRrange3.Offset(i - 1, 0).Value = "Verified"
    Next i
    
    Dim LRrange4 As Range
    Set LRrange4 = FormattedVT.Range("C" & Rows.Count).End(xlUp).Offset(1, 0)
    For i = 1 To vo2
        LRrange4.Offset(i - 1, 0).Value = "Void"
    Next i
    
    Dim LRrange5 As Range
    Set LRrange5 = FormattedVT.Range("C" & Rows.Count).End(xlUp).Offset(1, 0)
    For i = 1 To ve3
        LRrange5.Offset(i - 1, 0).Value = "Verified"
    Next i
    
    Dim LRrange6 As Range
    Set LRrange6 = FormattedVT.Range("C" & Rows.Count).End(xlUp).Offset(1, 0)
    For i = 1 To vo3
        LRrange6.Offset(i - 1, 0).Value = "Void"
    Next i
    
    Dim LRrange7 As Range
    Set LRrange7 = FormattedVT.Range("C" & Rows.Count).End(xlUp).Offset(1, 0)
    For i = 1 To ve4
        LRrange7.Offset(i - 1, 0).Value = "Verified"
    Next i
    
    Dim LRrange8 As Range
    Set LRrange8 = FormattedVT.Range("C" & Rows.Count).End(xlUp).Offset(1, 0)
    For i = 1 To vo4
        LRrange8.Offset(i - 1, 0).Value = "Void"
    Next i
    
    Dim LRrange9 As Range
    Set LRrange9 = FormattedVT.Range("C" & Rows.Count).End(xlUp).Offset(1, 0)
    For i = 1 To ve5
        LRrange9.Offset(i - 1, 0).Value = "Verified"
    Next i
    
    Dim LRrange10 As Range
    Set LRrange10 = FormattedVT.Range("C" & Rows.Count).End(xlUp).Offset(1, 0)
    For i = 1 To vo5
        LRrange10.Offset(i - 1, 0).Value = "Void"
    Next i
    
End Sub
Sub NewDate_2022()
        
    'ActiveCell must be on date
    
    Dim twentwo As Worksheet
    Set twentwo = Worksheets("2022")
    
    Dim FormattedVT As Worksheet
    Set FormattedVT = Worksheets("FormattedVT2022")
    
    Dim LRval, LRvalC As Long
    LRval = FormattedVT.Range("A" & Rows.Count).End(xlUp).Offset(1, 0).Row
    LRvalC = FormattedVT.Range("C" & Rows.Count).End(xlUp).Offset(1, 0).Row
    
    ActiveCell.Copy
    FormattedVT.Range("A" & LRval).PasteSpecial (xlPasteFormulasAndNumberFormats)
    FormattedVT.Range("A" & LRval).NumberFormat = "m/d/yyyy"
    FormattedVT.Range("A" & LRval & ":" & "A" & LRvalC - 1).FillDown
    
End Sub
Sub Date_2022()

    'ActiveCell must be on date
    
    Dim twentwo As Worksheet
    Set twentwo = Worksheets("2022")
    
    Dim FormattedVT As Worksheet
    Set FormattedVT = Worksheets("FormattedVT2022")
    
    Dim LRval As Long
    LRval = FormattedVT.Range("A" & Rows.Count).End(xlUp).Offset(1, 0).Row
    
    Dim vvttldate As Variant
    vvttldate = InputBox("Verified + Void For Specified Date")
    
    ActiveCell.Copy
    FormattedVT.Range("A" & LRval).PasteSpecial (xlPasteFormulasAndNumberFormats)
    FormattedVT.Range("A" & LRval).NumberFormat = "m/d/yyyy"
    FormattedVT.Range("A" & LRval & ":" & "A" & LRval + vvttldate - 1).FillDown

End Sub
Sub Name_2022()
    
    'ActiveCell must be on name (column A)
    
    Dim twentwo As Worksheet
    Set twentwo = Worksheets("2022")
    
    Dim FormattedVT As Worksheet
    Set FormattedVT = Worksheets("FormattedVT2022")
    
    Dim verivoidttl As Integer
    verivoidttl = ActiveCell.Offset(0, 13).Value
    
    Dim nametxt As Range
    Set nametxt = ActiveCell
    
    Dim LRval As Long
    LRval = FormattedVT.Range("D" & Rows.Count).End(xlUp).Offset(1, 0).Row
    
    nametxt.Copy Destination:=FormattedVT.Range("D" & LRval)
    FormattedVT.Range("D" & LRval).ClearFormats
    FormattedVT.Range("D" & LRval).Value = UCase(FormattedVT.Range("D" & LRval).Value)
    FormattedVT.Range("D" & LRval & ":" & "D" & LRval + verivoidttl - 1).FillDown
    
End Sub

Sub TestMacros()

    Dim vtr As Workbook
    Set vtr = Workbooks("Verification Totals Remote")
    
    Dim Giu919 As Worksheet
    Set Giu919 = Worksheets("Giuliana 919 - 923")
        
    MsgBox vtr.Sheets.Count
    
End Sub





