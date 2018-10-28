Sub Solution()


Dim ws As Worksheet

'Start looping throughout the workbook sheets
For Each ws In Worksheets
Dim WorkSheetName As String
WorkSheetName = ws.Name

'MsgBox (WorkSheetName)to test the loop to activate the correct worksheet
ws.Activate


' Set an initial variable for holding the brand name
 Dim Brand_Name As String

' Set an initial variable for holding the total per Ticker brand
 Dim Brand_Total As Double
 Brand_Total = 0

' Keep track of the location for each Ticker brand in the summary table
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

'Id the last row in the worksheet
LastRow = Cells(Rows.Count, 1).End(xlUp).Row

 ' Loop through all ticker vol
 For I = 2 To LastRow


   ' Check if we are still within the same ticker name, if it is not...
   If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then


     ' Set the Brand name
     Brand_Name = Cells(I, 1).Value

     ' Add to the Brand Total
     Brand_Total = Brand_Total + Cells(I, 7).Value

     Range("K1").Value = "Ticker"
     Range("L1").Value = "Stock total value"

     ' Print the ticket name in the Summary Table
     Range("K" & Summary_Table_Row).Value = Brand_Name

     ' Print the Brand Amount to the Summary Table
     Range("L" & Summary_Table_Row).Value = Brand_Total

     ' Add one to the summary table row
     Summary_Table_Row = Summary_Table_Row + 1

     ' Reset the Brand Total
     Brand_Total = 0

   ' If the cell immediately following a row is the same brand...
   Else
     ' Add to the Brand Total
     Brand_Total = Brand_Total + Cells(I, 7).Value


   End If

 Next I

Next ws
End Sub
