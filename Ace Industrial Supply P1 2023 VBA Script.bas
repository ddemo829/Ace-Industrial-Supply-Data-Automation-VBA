Attribute VB_Name = "VT_2023"
Sub Table_2023()

    Dim twenthree As Worksheet
    Set twenthree = Worksheets("2023")
    
    Dim FormattedVT As Worksheet
    Set FormattedVT = Worksheets("FormattedVT2023")
       
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
            'Dates
            If vevosum > 0 Then
            myrange.Cells(1, i).Copy
            FormattedVT.Range("A" & lra).PasteSpecial (xlPasteFormulasAndNumberFormats)
            FormattedVT.Range("A" & lra).NumberFormat = "m/d/yyyy"
            FormattedVT.Range("A" & lra).Value = Left(FormattedVT.Range("A" & lra).Value, _
            InStrRev(FormattedVT.Range("A" & lra).Value, "/", -1)) & "2023"
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
Sub Pend_2023()

    'ActiveCell must be on pendtotal
    
    Dim twentwo As Worksheet
    Set twentwo = Worksheets("2022")
    
    Dim FormattedVT As Worksheet
    Set FormattedVT = Worksheets("FormattedVT2023")
    
    Dim LRval As Long
    LRval = FormattedVT.Range("C" & Rows.Count).End(xlUp).Offset(1, 0).Row
    
    Dim pendttl As Integer
    pendttl = ActiveCell.Value
    
    Dim penddate As String
    penddate = ActiveCell.Offset(0, 1).Value
    
    FormattedVT.Range("C" & LRval).Value = "Pend"
    FormattedVT.Range("C" & LRval & ":" & "C" & LRval + pendttl - 1).FillDown
    FormattedVT.Range("D" & LRval).Value = "PEND"
    FormattedVT.Range("D" & LRval & ":" & "D" & LRval + pendttl - 1).FillDown

    FormattedVT.Range("A" & LRval).Value = Right(penddate, Len(penddate) - InStrRev(penddate, " "))
    FormattedVT.Range("A" & LRval).NumberFormat = "m/d/yyyy"
    If Right(FormattedVT.Range("A" & LRval).Value, 4) = "2022" _
    Or Right(FormattedVT.Range("A" & LRval).Value, 4) = "1900" _
    Or Right(FormattedVT.Range("A" & LRval).Value, 4) = "2023" Then
    FormattedVT.Range("A" & LRval).Value = Left(FormattedVT.Range("A" & LRval).Value, _
    InStrRev(FormattedVT.Range("A" & LRval).Value, "/", -1)) & "2023"
    FormattedVT.Range("A" & LRval & ":" & "A" & LRval + pendttl - 1).FillDown
    Else
    FormattedVT.Range("A" & LRval).Value = FormattedVT.Range("A" & LRval).Value & "/2023"
    FormattedVT.Range("A" & LRval & ":" & "A" & LRval + pendttl - 1).FillDown
    End If
    If Left(FormattedVT.Range("A" & LRval).Value, 1) = "0" Then
    FormattedVT.Range("A" & LRval).Value = Right(FormattedVT.Range("A" & LRval).Value, _
    Len(FormattedVT.Range("A" & LRval).Value) - 1)
    FormattedVT.Range("A" & LRval & ":" & "A" & LRval + pendttl - 1).FillDown
    End If
    
End Sub

Sub SoNo_2023()

    Dim FormattedVT2023 As Worksheet
    Set FormattedVT2023 = Worksheets("FormattedVT2023")
    
    Dim LRval As Long
    LRval = FormattedVT2023.Range("A" & Rows.Count).End(xlUp).Row
    
    Dim lrb As Long
    lrb = FormattedVT2023.Range("B" & Rows.Count).End(xlUp).Offset(1, 0).Row
    
    FormattedVT2023.Range("B" & lrb).Value = "1111111"
    Range("B" & lrb & ":" & "B" & lrb + LRval - 2).FillDown
    
    
End Sub

Sub DateFormat_All()

    Dim update3 As Worksheet
    Set update3 = Worksheets(1)
    
    Dim myrange As Range
    Set myrange = Selection
    
    Dim lr As Long
    lr = update3.Cells(Rows.Count, 1).End(xlUp).Row

    For i = 1 To lr
        myrange.Cells(i, 1).Value = Left(myrange.Cells(i, 1).Value, _
        InStrRev(myrange.Cells(i, 1).Value, "/", -1)) & "2023"
        myrange.Cells(i, 1).NumberFormat = "mm/dd/yyyy"
    Next
    
End Sub






