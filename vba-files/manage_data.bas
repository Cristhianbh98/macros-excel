Attribute VB_Name = "manage_data"

sub dataw1()
  ' Set variables
  Dim source As Worksheet, destination As Worksheet

  ' Set the source and destination sheets
  Set source = ThisWorkbook.Sheets("SOURCE")
  Set destination = ThisWorkbook.Sheets("C.I M.P")

  ' Call the copy function for each column
  Call actions.copy(source, destination, "FECHA", "A", True)
  Call actions.copy(source, destination, "HORA", "B")
  Call actions.fillData(destination, "C", "FECHA Y HORA", "=A2+B2")
  Call actions.copy(source, destination, "CONSUMO INSTA. MP BABOR", "D")
  Call actions.copy(source, destination, "CONSUMO INSTA. MP ESTRIBOR", "E")
  Call actions.updateChartData(destination)

  Debug.Print "Datos copiados a C.I M.P"
End Sub

sub dataw2()
  ' Set variables
  Dim source As Worksheet, destination As Worksheet

  ' Set the source and destination sheets
  Set source = ThisWorkbook.Sheets("SOURCE")
  Set destination = ThisWorkbook.Sheets("C.T.M.P")

  ' Set antother destination sheet
  Call actions.copy(source, destination, "FECHA", "A", True)
  Call actions.copy(source, destination, "HORA", "B", False, True)
  Call actions.fillData(destination, "C", "FECHA Y HORA", "=A2+B2")
  Call actions.copy(source, destination, "CONSUMO MP BABOR", "D")
  Call actions.copy(source, destination, "CONSUMO MP ESTRIBOR", "E")
  Call actions.updateChartData(destination)

  Debug.Print "Datos copiados a C.T.M.P"
End Sub

sub dataw3()
  ' Set variables
  Dim source As Worksheet, destination As Worksheet

  ' Set the source and destination sheets
  Set source = ThisWorkbook.Sheets("SOURCE")
  Set destination = ThisWorkbook.Sheets("C.I.M.A")

  ' Set antother destination sheet
  Call actions.copy(source, destination, "FECHA", "A", True)
  Call actions.copy(source, destination, "HORA", "B", False, True)
  Call actions.fillData(destination, "C", "FECHA Y HORA", "=A2+B2")
  Call actions.copy(source, destination, "CONSUMO INSTA. AUX 1", "D")
  Call actions.copy(source, destination, "CONSUMO INSTA. AUX 2", "E")
  Call actions.copy(source, destination, "CONSUMO INSTA. AUX 3", "F")
  Call actions.copy(source, destination, "CONSUMO INSTA. AUX 4", "G")
  Call actions.copy(source, destination, "CONSUMO INSTA. AUX 5", "H")

  Call actions.updateChartData(destination)

  Debug.Print "Datos copiados a C.I.M.A"
End Sub

sub dataw4()
  ' Set variables
  Dim source As Worksheet, destination As Worksheet

  ' Set the source and destination sheets
  Set source = ThisWorkbook.Sheets("SOURCE")
  Set destination = ThisWorkbook.Sheets("C.T.M.A")

  ' Set antother destination sheet
  Call actions.copy(source, destination, "FECHA", "A", True)
  Call actions.copy(source, destination, "HORA", "B", False, True)
  Call actions.fillData(destination, "C", "FECHA Y HORA", "=A2+B2")
  Call actions.copy(source, destination, "CONSUMO AUX 1", "D")
  Call actions.copy(source, destination, "CONSUMO AUX 2", "E")
  Call actions.copy(source, destination, "CONSUMO AUX 3", "F")
  Call actions.copy(source, destination, "CONSUMO AUX 4", "G")
  Call actions.copy(source, destination, "CONSUMO AUX 5", "H")
  Call actions.updateChartData(destination)

  Debug.Print "Datos copiados a C.T.M.A"
End Sub

Sub dataw5()
  ' Set variables
  Dim source As Worksheet, destination As Worksheet

  ' Set the source and destination sheets
  Set source = ThisWorkbook.Sheets("SOURCE")
  Set destination = ThisWorkbook.Sheets("C.I.M")

  ' Call actions
  Call actions.copy(source, destination, "FECHA", "A", True)
  Call actions.copy(source, destination, "HORA", "B", False, True)
  Call actions.fillData(destination, "C", "FECHA Y HORA", "=A2+B2")
  Call actions.copy(source, destination, "CONSUMO INSTA. MP BABOR", "D")
  Call actions.copy(source, destination, "CONSUMO INSTA. MP ESTRIBOR", "E")
  Call actions.copy(source, destination, "CONSUMO INSTA. AUX 1", "F")
  Call actions.copy(source, destination, "CONSUMO INSTA. AUX 2", "G")
  Call actions.copy(source, destination, "CONSUMO INSTA. AUX 3", "H")
  Call actions.copy(source, destination, "CONSUMO INSTA. AUX 4", "I")
  Call actions.copy(source, destination, "CONSUMO INSTA. AUX 5", "J")

  Debug.Print "Datos copiados a C.I.M"
End Sub

Sub dataw6()
  ' Set variables
  Dim source As Worksheet, destination As Worksheet

  ' Set the source and destination sheets
  Set source = ThisWorkbook.Sheets("SOURCE")
  Set destination = ThisWorkbook.Sheets("C.T.M")

  ' Call actions
  Call actions.copy(source, destination, "FECHA", "A", True)
  Call actions.copy(source, destination, "HORA", "B", False, True)
  Call actions.fillData(destination, "C", "FECHA Y HORA", "=A2+B2")
  Call actions.copy(source, destination, "CONSUMO MP BABOR", "D")
  Call actions.copy(source, destination, "CONSUMO MP ESTRIBOR", "E")
  Call actions.copy(source, destination, "CONSUMO AUX 1", "F")
  Call actions.copy(source, destination, "CONSUMO AUX 2", "G")
  Call actions.copy(source, destination, "CONSUMO AUX 3", "H")
  Call actions.copy(source, destination, "CONSUMO AUX 4", "I")
  Call actions.copy(source, destination, "CONSUMO AUX 5", "J")

  Debug.Print "Datos copiados a C.T.M"
End Sub

sub manageData()
  ' Disable screen updating and automatic calculation
  Application.ScreenUpdating = False
  Application.Calculation = xlCalculationManual

  ' Call subroutines
  Debug.Print "Llamando a dataw1"
  Call dataw1
  Debug.Print "Llamando a dataw2"
  Call dataw2
  Debug.Print "Llamando a dataw3"
  Call dataw3
  Debug.Print "Llamando a dataw4"
  Call dataw4
  
  Debug.Print "Finalizando manageData"

  ' Enable screen updating and automatic calculation
  Application.ScreenUpdating = True
  Application.Calculation = xlCalculationAutomatic

  MsgBox "Gr" & ChrW(225) & "ficos Actualizados"
End Sub

sub refreshData()
  ' Call subroutines
  Debug.Print "Llamando a refreshData"
  Call actions.refreshData
  MsgBox "Datos actualizados"
End sub

sub savePDF()
  ' Set variables
  Dim ws As Worksheet
  Dim range As String

  ' Set the worksheet
  Set ws = ThisWorkbook.Sheets("DASHBOARD")

  ' Set the range
  range = "A:G"

  ' Call subroutines
  Debug.Print "Llamando a savePDF"
  Call actions.savePDF(ws, range)
End sub

sub savePDF2()
  ' Set variables
  Dim ws As Worksheet
  Dim range As String

  ' Set the worksheet
  Set ws = ThisWorkbook.Sheets("DASHBOARD 2")

  ' Set the range
  range = "A:G"

  ' Call subroutines
  Debug.Print "Llamando a savePDF"
  Call actions.savePDF(ws, range)
End sub
