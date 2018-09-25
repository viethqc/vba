Sub foo()
    Dim x As Workbook
    Dim y As Workbook
    
    '## Open both workbooks first:
    Dim first_row As Integer
    
    MAX_LENGTH_ROW = 754
    
    first_row = find_first_row("DATA")
    Sheets("DATA").Range("A" + CStr(first_row) + ":H" + CStr(first_row + MAX_LENGTH_ROW)).Copy Destination:=Sheets("test").Range("A1")
    
    col_start = Asc("I") - 64
    
    Dim sheet_month As String
    
    For i = 1 To 12
        sheet_month = "THANG" + CStr(i)
        first_row = find_first_row(sheet_month)
        cordinate_source = "I" + CStr(first_row) + ":L" + CStr(first_row + MAX_LENGTH_ROW)
        cordinate_dest = Split(Cells(1, col_start).Address(True, False), "$")(0)
        Sheets(sheet_month).Range(cordinate_source).Copy Destination:=Sheets("test").Range(cordinate_dest + CStr(1))
        col_start = col_start + 4
        
    Next i
    
    first_row = find_first_row("THANG1")

End Sub

Function find_first_row(sheet_name As String) As Integer
    Dim row_data As String
    Dim first_row As Integer

    For i = 1 To 100
        row_data = Sheets(sheet_name).Range("A" + CStr(i)).Value
        If row_data = "TT" Then
            first_row = i
            Exit For
        End If
    Next i
    
    find_first_row = first_row
End Function