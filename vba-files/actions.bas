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
      wsDest.Range(destCol & "1:" & destCol & lastRow).NumberFormat = "dd/mm/yyyy"
    End If

    ' Clear the clipboard
    Application.CutCopyMode = False
  Else
    MsgBox "Column " & colName & " not found in the source sheet"
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
