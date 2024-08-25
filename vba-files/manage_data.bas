Attribute VB_Name = "manage_data"

sub manageData()
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

End Sub
