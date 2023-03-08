Attribute VB_Name = "Module1"
Sub NextCells():
Attribute NextCells.VB_ProcData.VB_Invoke_Func = " \n14"
Dim i As Long
Dim tempsymbol As String
Dim tempvolume As Double
Dim tempprice As Double
Dim Summarytablerow As Integer


Const Tickercolumn As Integer = 9
Const Yearlychangecolumn As Double = 10
Const Volumecolumn As Integer = 11


Summary_Table_row = 1

tempvolume = 0
tempprice = 0
tempsymbol = ""

    For i = 2 To 753001
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
        
            tempprice = Cells(i, 6).Value
            tempvolume = Cells(i, 7).Value
            tempsymbol = Cells(i, 1).Value
            
            Summary_Table_row = Summary_Table_row + 1
            
            tempsymbol = Cells(i, 1).Value
            tempvolume = Cells(i, 7).Value
            tempprice = Cells(i, 6).Value
            
            Cells(Summary_Table_row, Tickercolumn).Value = Cells(i, 1)
            Cells(Summary_Table_row, Volumecolumn).Value = Cells(i, 7).Value
        Else
            
     tempvolume = tempvolume + Cells(i, 7).Value
    tempsymbol = Cells(i, 1).Value
   
     Cells(Summary_Table_row, Tickercolumn).Value = tempsymbol
    Cells(Summary_Table_row, Volumecolumn).Value = tempvolume
        
            End If
            Next i
        
        
    ActiveWindow.SmallScroll ToRight:=4
    ActiveWindow.SmallScroll Down:=0
    Columns("K:K").ColumnWidth = 21.83
    ActiveWindow.SmallScroll Down:=0
    Range("L1").Select
    Selection.ClearContents
    Columns("K:K").Select
End Sub
