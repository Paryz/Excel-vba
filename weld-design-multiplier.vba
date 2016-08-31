
Sub Macro2()

Dim a, b, c, d, e, f, g, h As String
Dim i, j, k, l, p, q, r, s As Integer
Dim testresult As Boolean

j = Cells(1, 12)

For i = 1 To j
    
    a = Sheets(1).Cells(i, 1)
    b = Sheets(1).Cells(i, 2)
    c = Sheets(1).Cells(i, 3)
    d = Sheets(1).Cells(i, 4)
    e = Sheets(1).Cells(i, 5)
    f = Sheets(1).Cells(i, 6)
    g = Sheets(1).Cells(i, 7)
    h = Sheets(1).Cells(i, 8)
    
    Sheets(2).Name = b
    
    Sheets(2).Cells(8, 1) = "Level " & a & " Column " & b
    Sheets(2).Cells(13, 7) = c
    Sheets(2).Cells(14, 7) = d
    Sheets(2).Cells(15, 7) = e
    Sheets(2).Cells(16, 7) = f
    Sheets(2).Cells(18, 7) = g
    Sheets(2).Cells(24, 3) = h

    l = Sheets(2).Cells(175, 2)
    p = Sheets(2).Cells(192, 8)
    q = Sheets(2).Cells(192, 4)
    r = Sheets(2).Cells(194, 8)
    s = Sheets(2).Cells(194, 4)

    For k = l To 0.7 * WorksheetFunction.Min(e, f)
        testresult = p > q And r > s
        If testresult = True Then
            Sheets(2).Cells(176, 2) = k
            Exit For
        Else
        End If
    Next k



    Application.DisplayAlerts = False
       Sheets(2).Copy
    With ActiveWorkbook
        .SaveAs Filename:="C:\Users\plkz00149\Desktop\excel spoiny\" & a & " " & b & ".xlsx", _
    FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
        .Close 0
    End With
   
      
Windows("duplo.xlsm").Activate
Next i

End Sub
