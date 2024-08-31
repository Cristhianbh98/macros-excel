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

  Debug.Print "Datos copiados a KW NOMINAL VS CARGA INDIVIDUAL"
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

  Debug.Print "Datos copiados a POTENCIA"
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

  Debug.Print "Datos copiados a COMPRESOR TORNILLO 1"
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

  Debug.Print "Datos copiados a COMPRESOR PISTON 1"
End Sub

sub manageData()
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
  Dim path As String
  Dim fileName As String
  Dim imgWidth As Double
  Dim imgHeight As Double

  ' Set the image size
  imgWidth = 594
  imgHeight = 102

  ' Set the worksheet
  Set ws = ThisWorkbook.Sheets("DASHBOARD")

  ' Define the print area
  ws.PageSetup.PrintArea = ws.Range("A:G").Address

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
End sub
