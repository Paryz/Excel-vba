Private Sub UpdateCalculations_Click()

'Variable declaration
'Input variables
Dim Mu As Integer
Dim Vu As Integer
Dim MserDead As Integer
Dim MserTotal As Integer
Dim b As Integer
Dim h As Integer
Dim bf As Integer
Dim hf As Integer
Dim l As Integer
Dim cover As Integer
Dim conc As Integer
Dim steel As Integer
Dim nRowsB As Integer
Dim dBarB As Integer
Dim nBarsB As Integer
Dim nRowsT As Integer
Dim dBarT As Integer
Dim nBarsT As Integer
Dim dStirrup As Integer
Dim nLegs As Integer
Dim sLegs As Integer
'Result variables
Dim Mcap As Integer
Dim Vstrut As Integer
Dim Vcap As Integer
Dim Vmoment As Integer
Dim Mcrack As Integer
Dim crackLong As Integer
Dim crackShort As Integer
'Other variables
Dim nBeams As Integer
Dim startRow As Integer 'first row of data (change if rows above are added or removed)
Dim startCol As Integer 'first column of data (change if columns on left side are added or removed)
Dim currentRow As Integer
Dim i As Integer 'counter
Dim dataT As ListObject 'Table variable


'Size and position of table with data
Set dataT = ActiveSheet.ListObjects(1)
nBeams = Range("nBeams").Value
'startRow = 4
'startColumn = 14

'Looping through every beam
For i = 1 To 2 'nBeams
    
    'Copying data on design sheet
    Worksheets(1).Range("C5").Value = dataT.DataBodyRange(i, 4)
    Worksheets(1).Range("C6").Value = dataT.DataBodyRange(i, 5)
    Worksheets(1).Range("I5").Value = dataT.DataBodyRange(i, 6)
    Worksheets(1).Range("I6").Value = dataT.DataBodyRange(i, 7)
    Worksheets(1).Range("C9").Value = dataT.DataBodyRange(i, 8)
    Worksheets(1).Range("C10").Value = dataT.DataBodyRange(i, 9)
    Worksheets(1).Range("C11").Value = dataT.DataBodyRange(i, 10)
    Worksheets(1).Range("C12").Value = dataT.DataBodyRange(i, 11)
    Worksheets(1).Range("C13").Value = dataT.DataBodyRange(i, 12)
    Worksheets(1).Range("C14").Value = dataT.DataBodyRange(i, 13)
    Worksheets(1).Range("C15").Value = dataT.DataBodyRange(i, 14)
    Worksheets(1).Range("C16").Value = dataT.DataBodyRange(i, 15)
    Worksheets(1).Range("C20").Value = dataT.DataBodyRange(i, 16)
    Worksheets(1).Range("C21").Value = dataT.DataBodyRange(i, 17)
    Worksheets(1).Range("C22").Value = dataT.DataBodyRange(i, 18)
    Worksheets(1).Range("C25").Value = dataT.DataBodyRange(i, 19)
    Worksheets(1).Range("C26").Value = dataT.DataBodyRange(i, 20)
    Worksheets(1).Range("C27").Value = dataT.DataBodyRange(i, 21)
    Worksheets(1).Range("C30").Value = dataT.DataBodyRange(i, 22)
    Worksheets(1).Range("C31").Value = dataT.DataBodyRange(i, 23)
    Worksheets(1).Range("C32").Value = dataT.DataBodyRange(i, 24)
    
    'Getting results
    dataT.DataBodyRange(i, 25) = Worksheets(1).Range("D36").Value
    dataT.DataBodyRange(i, 26) = Worksheets(1).Range("J32").Value
    dataT.DataBodyRange(i, 27) = Worksheets(1).Range("J34").Value
    dataT.DataBodyRange(i, 28) = Worksheets(1).Range("J36").Value
    dataT.DataBodyRange(i, 31) = Worksheets(1).Range("E40").Value
    dataT.DataBodyRange(i, 32) = Worksheets(1).Range("E41").Value
    dataT.DataBodyRange(i, 33) = Worksheets(1).Range("E42").Value
    
    'Checking warnings
    If Worksheets(1).Range("B38").Value <> "" Then
        dataT.DataBodyRange(i, 34) = "NO GOOD"
    Else
        dataT.DataBodyRange(i, 34) = "OK"
    End If
    
    If Worksheets(1).Range("E11").Value <> "" Or Worksheets(1).Range("E12").Value <> "" Then
        dataT.DataBodyRange(i, 35) = "NO GOOD"
    Else
        dataT.DataBodyRange(i, 35) = "OK"
    End If
        
Next i
    
End Sub
