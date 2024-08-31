Attribute VB_Name = "actions"

Public Sub copy(wsSource As Worksheet, wsDest As Worksheet, colName As String, destCol As String, isDate As Boolean)
  ' Set variables
  Dim lastRow As Long
  Dim colIndex As Integer
  Dim found As Range

  ' Get colName column index
  Set found = wsSource.Rows(1).Find(What:=colName, LookIn:=xlValues, LookAt:=xlWhole)

  IF Not found Is Nothing Then
    ' Copy the column
    colIndex = found.Column
    lastRow = wsSource.Cells(wsSource.Rows.Count, colIndex).End(xlUp).Row

    ' Paste the column in the destination sheet
    wsSource.Range(wsSource.Cells(1, colIndex), wsSource.Cells(lastRow, colIndex)).Copy
    wsDest.Range(destCol & "1").PasteSpecial Paste:=xlPasteValues

    ' Set dat format
    If isDate Then
      ' wsDest.Range(destCol & "1:" & destCol & lastRow).NumberFormat = "dd/mm/yyyy"
      wsDest.Range(destCol & "1:" & destCol & lastRow).NumberFormat = "dd/mm"
    End If

    ' Clear the clipboard
    Application.CutCopyMode = False
    Debug.Print "Datos copiados de " & colName & " a " & destCol
  Else
    Debug.Print "Columna " & colName & " no encontrada en la hoja de origen"
  End If
End Sub

Public sub fillData(ws As Worksheet, col As String, headerName As String, data As String)
  ' Set variables
  Dim lastRow As Long

  ' Disable screen updating and automatic calculation
  Application.ScreenUpdating = False
  Application.Calculation = xlCalculationManual

  ' Get last row
  lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

  ' Set Header Name
  ws.Range(col & "1").Value = headerName

  ' Fill the column
  If lastRow > 1 Then
    ws.Range(col & "2:" & col & lastRow).Value = data
  Else
    MsgBox "No data to fill"
  End If

  ' Enable screen updating and automatic calculation
  Application.ScreenUpdating = True
  Application.Calculation = xlCalculationAutomatic
End Sub

Public Sub updateChartData(ws As Worksheet)
  ' Set variables
  Dim lastRow As Long
  Dim lastCol As Long
  Dim found As Range
  Dim range As String

  ' Get last row and last column
  lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
  lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

  ' Set the new range
  range = "A1:" & columnNumberToLetter(lastCol) & lastRow

  For Each chartObj In ws.ChartObjects
    ' Update the chart data
    chartObj.Chart.SetSourceData Source:=ws.Range(range)
  Next chartObj
  

End Sub

Public Sub refreshData()
  ActiveWorkbook.RefreshAll
End Sub

Function columnNumberToLetter(colNum As Long) As String
  Dim dividend As Long
  Dim modulo As Integer
  Dim columnLetter As String
  
  dividend = colNum
  columnLetter = ""
  
  While dividend > 0
    modulo = (dividend - 1) Mod 26
    columnLetter = Chr(65 + modulo) & columnLetter
    dividend = Int((dividend - modulo) / 26)
  Wend
  
  columnNumberToLetter = columnLetter
End Function