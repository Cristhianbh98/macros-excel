Attribute VB_Name = "actions"

Public Sub copy(wsSource As Worksheet, wsDest As Worksheet, colName As String, destCol As String, Optional isDate As Boolean = False, Optional isTime As Boolean = False)
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
      wsDest.Range(destCol & "1:" & destCol & lastRow).NumberFormat = "dd/mm/yyyy"
    ElseIF isTime Then
      wsDest.Range(destCol & "1:" & destCol & lastRow).NumberFormat = "hh:mm:ss"
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

  ' Get last row
  lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

  ' Set Header Name
  ws.Range(col & "1").Value = headerName

  ' Fill the column
  If lastRow > 1 Then
    ws.Range(col & "2:" & col & lastRow).Value = data
  Else
    Debug.Print "No data to fill"
  End If

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

Public Sub savePDF(ws As Worksheet, range As String)
  ' Set variables
  Dim path As String
  Dim fileName As String
  Dim imgWidth As Double
  Dim imgHeight As Double

  ' Set the image size
  imgWidth = 594
  imgHeight = 102

  ' Define the print area
  ws.PageSetup.PrintArea = ws.Range(range).Address

  ' Set header and footer for the pdf
  With ws.PageSetup
    .CenterHeader = "&G"
    .CenterHeaderPicture.Filename = ThisWorkbook.Path & "\img\header.png"
    .centerHeaderPicture.Width = imgWidth
    .centerHeaderPicture.Height = imgHeight

    .CenterFooter = "&G"
    .CenterFooterPicture.Filename = ThisWorkbook.Path & "\img\footer.png"
    .centerFooterPicture.Width = imgWidth
    .centerFooterPicture.Height = imgHeight

    .TopMargin = imgHeight + Application.InchesToPoints(0.2)
    .HeaderMargin = Application.InchesToPoints(0)
    .BottomMargin = imgHeight + Application.InchesToPoints(0.2)
    .FooterMargin = Application.InchesToPoints(0)
  End With

  ' Set the path and file name
  path = ThisWorkbook.Path & "\pdf\"

  ' Save the worksheet as PDF
  fileName = "Informe" & Format(Now, "yyyymmdd_hhmmss") & ".pdf"

  ' Save the PDF
  ws.ExportAsFixedFormat Type:=xlTypePDF, _
                        Filename:=path & fileName, _
                        Quality:=xlQualityStandard, _
                        IncludeDocProperties:=True, _
                        IgnorePrintAreas:=False, _
                        OpenAfterPublish:=True
End Sub

Public Sub changeColumnHeader(ws As Worksheet, col As String, newHeader As String)
  ws.Range(col & "1").Value = newHeadeR
End Sub

' Functions
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
