Attribute VB_Name = "manage_data"

sub dataw1()
  ' Set variables
  Dim source As Worksheet, destination As Worksheet

  ' Set the source and destination sheets
  Set source = ThisWorkbook.Sheets("SOURCE")
  Set destination = ThisWorkbook.Sheets("KW NOMINAL VS CARGA INDIVIDUAL")

  ' Call the copy function for each column
  Call actions.copy(source, destination, "TimeStr", "A", True)
  Call actions.copy(source, destination, "PotenciaG1", "B", False)
  Call actions.copy(source, destination, "PotenciaG2", "C", False)
  Call actions.fillData(destination, "D", "KW optimo G1", "510")
  Call actions.updateChartData(destination)
End Sub

sub dataw2()
  ' Set variables
  Dim source As Worksheet, destination As Worksheet

  ' Set the source and destination sheets
  Set source = ThisWorkbook.Sheets("SOURCE")
  Set destination = ThisWorkbook.Sheets("POTENCIA")

  ' Set antother destination sheet
  Call actions.copy(source, destination, "TimeStr", "A", True)
  Call actions.copy(source, destination, "PotenciaG1", "B", False)
  Call actions.copy(source, destination, "PotenciaG2", "C", False)
  Call actions.fillData(destination, "D", "Potencia G1+G2", "=B2+C2")
  Call actions.fillData(destination, "E", "Nominal G1", "510")
  Call actions.fillData(destination, "F", "Nominal G1+G2", "=E2*2")
  Call actions.fillData(destination, "G", "Nominal G1+G2+G3", "=E2*3")
  Call actions.updateChartData(destination)
End Sub

sub dataw3()
  ' Set variables
  Dim source As Worksheet, destination As Worksheet

  ' Set the source and destination sheets
  Set source = ThisWorkbook.Sheets("SOURCE")
  Set destination = ThisWorkbook.Sheets("COMPRESOR TORNILLO 1")

  ' Set antother destination sheet
  Call actions.copy(source, destination, "TimeStr", "A", True)
  Call actions.copy(source, destination, "AmperajeC1", "B", False)
  Call actions.copy(source, destination, "PotenciaCT1", "C", False)
  Call actions.copy(source, destination, "PresionAspiracionCT1", "D", False)
  Call actions.copy(source, destination, "PresionDescargaCT1", "E", False)
  Call actions.copy(source, destination, "TemperaturaAceiteCT1", "F", False)
  Call actions.copy(source, destination, "TemperaturaAspiracionCT1", "G", False)
  Call actions.copy(source, destination, "TemperaturaADescargaCT1", "H", False)
  Call actions.updateChartData(destination)
End Sub

sub dataw4()
  ' Set variables
  Dim source As Worksheet, destination As Worksheet

  ' Set the source and destination sheets
  Set source = ThisWorkbook.Sheets("SOURCE")
  Set destination = ThisWorkbook.Sheets("COMPRESOR PISTON 1")

  ' Set antother destination sheet
  Call actions.copy(source, destination, "TimeStr", "A", True)
  Call actions.copy(source, destination, "AmperajeC5", "B", False)
  Call actions.copy(source, destination, "PotenciaCP", "C", False)
  Call actions.copy(source, destination, "PresionAspiracionCP1", "D", False)
  Call actions.copy(source, destination, "PresionDescargaCP1", "E", False)
  Call actions.copy(source, destination, "TemperaturaAspiracionCP1", "F", False)
  Call actions.copy(source, destination, "TemperaturaADescargaCP1", "G", False)
  Call actions.updateChartData(destination)
End Sub

sub manageData()
  ' Call subroutines
  Call dataw1
  Call dataw2
  Call dataw3
  Call dataw4
End Sub
