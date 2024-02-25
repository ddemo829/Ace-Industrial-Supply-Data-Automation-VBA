Attribute VB_Name = "VT_P2"
Sub Sheet_P2()

    Dim vtr As Workbook
    Set vtr = Workbooks("Verification Totals Remote.xlsm")
    Dim mac As Worksheet
    Set mac = Worksheets("mac")
    Dim k, i, wssum As Long
    wssum = vtr.Worksheets.Count
    
    For k = 270 To wssum
    Dim wsk As Worksheet
    Set wsk = vtr.Worksheets(k)
    If wsk.Visible Then
        For i = 1 To 10
            Dim lri, lra, lrb, lrc, lrd, sonosum As Long
            lri = wsk.Cells(wsk.Rows.Count, i).End(xlUp).Offset(1, 0).Row
            lra = mac.Cells(mac.Rows.Count, 1).End(xlUp).Offset(1, 0).Row
            lrb = mac.Cells(mac.Rows.Count, 2).End(xlUp).Offset(1, 0).Row
            lrc = mac.Cells(mac.Rows.Count, 3).End(xlUp).Offset(1, 0).Row
            lrd = mac.Cells(mac.Rows.Count, 4).End(xlUp).Offset(1, 0).Row
            sonosum = wsk.Range(wsk.Cells(4, i), wsk.Cells(lri - 1, i)).Count
            'Sono
            wsk.Range(wsk.Cells(4, i), wsk.Cells(lri, i)).Copy Destination:=mac.Range("B" & lrb)
            If i = 1 Or i = 3 Or i = 5 Or i = 7 Or i = 9 Then
                If lri > 4 Then
                'Status
                mac.Range("C" & lrc).Value = "Verified"
                mac.Range("C" & lrc & ":" & "C" & lrc + sonosum - 1) = _
                mac.Range("C" & lrc).Value
                'Name
                mac.Range("D" & lrd).Value = Left(wsk.Name, InStr(wsk.Name, " ") - 1)
                mac.Range("D" & lrd).Value = Left(UCase(mac.Range("D" & lrd).Value), 1) & _
                Right(LCase(mac.Range("D" & lrd).Value), Len(mac.Range("D" & lrd).Value) - 1)
                mac.Range("D" & lrd & ":" & "D" & lrd + sonosum - 1) = _
                mac.Range("D" & lrd).Value
                'Date
                mac.Range("A" & lra).Value = _
                Right(wsk.Name, Len(wsk.Name) - InStr(wsk.Name, " "))
                mac.Range("A" & lra).Value = _
                Left(mac.Range("A" & lra).Value, InStr(mac.Range("A" & lra).Value, "-") - 1)
                mac.Range("A" & lra).NumberFormat = "General"
                    If InStr(mac.Range("A" & lra).Value, ",") > 0 Then 'For 2022
                    mac.Range("A" & lra).Value = _
                    Left(mac.Range("A" & lra).Value, InStr(mac.Range("A" & lra).Value, ",") - 1) & _
                    "/" & Right(mac.Range("A" & lra).Value, Len(mac.Range("A" & lra).Value) - InStr(mac.Range("A" & lra).Value, ","))
                    mac.Range("A" & lra).NumberFormat = "mm/dd/yyyy"
                    mac.Range("A" & lra).Value = _
                    Left(mac.Range("A" & lra).Value, InStrRev(mac.Range("A" & lra).Value, "/")) & "2022"
                    Else 'For 2023
                    mac.Range("A" & lra).Value = _
                    Left(mac.Range("A" & lra).Value, InStr(mac.Range("A" & lra).Value, ".") - 1) & _
                    "/" & Right(mac.Range("A" & lra).Value, Len(mac.Range("A" & lra).Value) - InStr(mac.Range("A" & lra).Value, "."))
                    mac.Range("A" & lra).NumberFormat = "mm/dd/yyyy"
                    mac.Range("A" & lra).Value = _
                    Left(mac.Range("A" & lra).Value, InStrRev(mac.Range("A" & lra).Value, "/")) & "2023"
                    End If
                        If i = 1 Then
                        mac.Range("A" & lra).Value = mac.Range("A" & lra).Value + (i - 1)
                        End If
                            If i = 3 Then
                            mac.Range("A" & lra).Value = mac.Range("A" & lra).Value + (i - 2)
                            End If
                                If i = 5 Then
                                mac.Range("A" & lra).Value = mac.Range("A" & lra).Value + (i - 3)
                                End If
                                    If i = 7 Then
                                    mac.Range("A" & lra).Value = mac.Range("A" & lra).Value + (i - 4)
                                    End If
                                        If i = 9 Then
                                        mac.Range("A" & lra).Value = mac.Range("A" & lra).Value + (i - 5)
                                        End If
                                        mac.Range("A" & lra & ":" & "A" & lra + sonosum - 1) = _
                                        mac.Range("A" & lra).Value
                                        mac.Range("A" & lra & ":" & "A" & lra + sonosum - 1).NumberFormat = "mm/dd/yyyy"
                End If
            Else
                If lri > 4 Then
                'Status
                mac.Range("C" & lrc).Value = "Void"
                mac.Range("C" & lrc & ":" & "C" & lrc + sonosum - 1) = _
                mac.Range("C" & lrc).Value
                'Name
                mac.Range("D" & lrd).Value = Left(wsk.Name, InStr(wsk.Name, " ") - 1)
                mac.Range("D" & lrd).Value = Left(UCase(mac.Range("D" & lrd).Value), 1) & _
                Right(LCase(mac.Range("D" & lrd).Value), Len(mac.Range("D" & lrd).Value) - 1)
                mac.Range("D" & lrd & ":" & "D" & lrd + sonosum - 1) = _
                mac.Range("D" & lrd).Value
                'Date
                mac.Range("A" & lra).Value = _
                Right(wsk.Name, Len(wsk.Name) - InStr(wsk.Name, " "))
                mac.Range("A" & lra).Value = _
                Left(mac.Range("A" & lra).Value, InStr(mac.Range("A" & lra).Value, "-") - 1)
                mac.Range("A" & lra).NumberFormat = "General"
                    If InStr(mac.Range("A" & lra).Value, ",") > 0 Then 'For 2022
                    mac.Range("A" & lra).Value = _
                    Left(mac.Range("A" & lra).Value, InStr(mac.Range("A" & lra).Value, ",") - 1) & _
                    "/" & Right(mac.Range("A" & lra).Value, Len(mac.Range("A" & lra).Value) - InStr(mac.Range("A" & lra).Value, ","))
                    mac.Range("A" & lra).NumberFormat = "mm/dd/yyyy"
                    mac.Range("A" & lra).Value = _
                    Left(mac.Range("A" & lra).Value, InStrRev(mac.Range("A" & lra).Value, "/")) & "2022"
                    Else 'For 2023
                    mac.Range("A" & lra).Value = _
                    Left(mac.Range("A" & lra).Value, InStr(mac.Range("A" & lra).Value, ".") - 1) & _
                    "/" & Right(mac.Range("A" & lra).Value, Len(mac.Range("A" & lra).Value) - InStr(mac.Range("A" & lra).Value, "."))
                    mac.Range("A" & lra).NumberFormat = "mm/dd/yyyy"
                    mac.Range("A" & lra).Value = _
                    Left(mac.Range("A" & lra).Value, InStrRev(mac.Range("A" & lra).Value, "/")) & "2023"
                    End If
                        If i = 2 Then
                        mac.Range("A" & lra).Value = mac.Range("A" & lra).Value + (i - 2)
                        End If
                            If i = 4 Then
                            mac.Range("A" & lra).Value = mac.Range("A" & lra).Value + (i - 3)
                            End If
                                If i = 6 Then
                                mac.Range("A" & lra).Value = mac.Range("A" & lra).Value + (i - 4)
                                End If
                                    If i = 8 Then
                                    mac.Range("A" & lra).Value = mac.Range("A" & lra).Value + (i - 5)
                                    End If
                                        If i = 10 Then
                                        mac.Range("A" & lra).Value = mac.Range("A" & lra).Value + (i - 6)
                                        End If
                                        mac.Range("A" & lra & ":" & "A" & lra + sonosum - 1) = _
                                        mac.Range("A" & lra).Value
                                        mac.Range("A" & lra & ":" & "A" & lra + sonosum - 1).NumberFormat = "mm/dd/yyyy"
                End If
            End If
        Next
    End If
    Next
    
End Sub
Sub TestMacro()

    Dim vtr As Workbook
    Set vtr = Workbooks("Verification Totals Remote.xlsm")
    
    Dim mac As Worksheet
    Set mac = Worksheets("mac")
    
    Dim ws As Worksheet
    For Each ws In vtr.Worksheets
        ws.Range("A3:J3").Value = "!"
        Next
        
End Sub





