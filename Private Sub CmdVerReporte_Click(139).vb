Private Sub CmdVerReporte_Click()
'''If Me.DTPF_Ini.Value < Date Then
'''    MsgBox "Verifique las fechas, no puede sacar reportes para fechas pasadas.", vbExclamation, MSG
'''    Exit Sub
'''End If
'Cambio 15 may 2024
If Me.DTPF_Ini.Value > Me.DTPF_Fin.Value Then
    MsgBox "Verifique las fechas, la Inicial no puede ser mayor que la Final.", vbExclamation, MSG
    Exit Sub
End If

''''LLena hoja de excel
 'Dim F As ADODB.Field
'Dim hoja  As Variant

'''Dim xlApp As Excel.Application
'''Dim xlBook As Excel.Workbook
'''Dim xlSheet As Excel.Worksheet
'''Set xlApp = CreateObject("Excel.Application")
'''Set xlBook = xlApp.Workbooks.Add
'''Set xlSheet = xlBook.Worksheets("Sheet1")
'''xlSheet.Range(Cells(1, 1), Cells(10, 2)).Value = "Hello"
'''xlBook.Saved = True
'''Set xlSheet = Nothing
'''Set xlBook = Nothing
'''xlApp.Quit
'''Set xlApp = Nothing

''''Dim hoja  As Excel.Application
'''Set hoja = New Excel.Application

Dim vApp As Excel.Application
Dim libro As Excel.Workbook
Dim Hoja As Excel.Worksheet

Set vApp = CreateObject("Excel.Application")
Set libro = vApp.Workbooks.Add
Set Hoja = libro.Worksheets(1)





Renglon = 4
Columna = 1
'hoja.Sheets(1).Select
Hoja.Select
Resp = "Programa " & Format(Me.DTPF_Ini.Value, "d mmm yy") & " a " & Format(Me.DTPF_Fin.Value, "d mmm yy")
'hoja.Sheets(1).Name = Resp

Hoja.Name = Resp

''hoja.Sheets(Resp).Cells(1, 1) = "Gemtron de México S.A. de C.V."
'Hoja.Cells(1, 1) = "Gemtron de México S.A. de C.V."
'
'vApp.Visible = True
'
'Hoja.Range("A1:k1").Select
'Hoja.Range("A1:k1").HorizontalAlignment = xlCenter
'Hoja.Range("A1:k1").VerticalAlignment = xlBottom
'Hoja.Range("A1:k1").WrapText = False
'Hoja.Range("A1:k1").MergeCells = True
'
'Hoja.Range("A1:k1").Select
'Hoja.Range("A1:k1").Font.Name = "Times New Roman"
'Hoja.Range("A1:k1").Font.Size = 16
'Hoja.Range("A1:k1").Font.Bold = True
'
'
'Hoja.Cells(2, 1) = "Reporte Programa de Producción"
'Hoja.Range("A2:K2").Select
'Hoja.Range("A2:K2").HorizontalAlignment = xlCenter
'Hoja.Range("A2:K2").VerticalAlignment = xlBottom
'Hoja.Range("A2:K2").WrapText = False
'Hoja.Range("A2:K2").MergeCells = True
'
'Hoja.Range("A2:K2").Select
'Hoja.Range("A2:K2").Font.Name = "Times New Roman"
'Hoja.Range("A2:K2").Font.Size = 12
'Hoja.Range("A2:K2").Font.Bold = True
'Hoja.Range("A2:K2").Font.Italic = True


'vApp.Visible = True


'Encabezados
Hoja.Cells(4, 1) = "Fecha"

Hoja.Cells(4, 3) = "Linea 10"
Hoja.Cells(5, 2) = "Hora"
Hoja.Cells(5, 3) = "No.  JC"
Hoja.Cells(5, 4) = "No de Parte"
Hoja.Cells(5, 5) = "Cantidad"

Hoja.Cells(4, 6) = "Linea 20"
Hoja.Cells(5, 6) = "No.  JC"
Hoja.Cells(5, 7) = "No de Parte"
Hoja.Cells(5, 8) = "Cantidad"

Hoja.Cells(4, 9) = "Linea 60"
Hoja.Cells(5, 9) = "No.  JC"
Hoja.Cells(5, 10) = "No de Parte"
Hoja.Cells(5, 11) = "Cantidad"

Hoja.Cells(4, 12) = "Manual Izq."
Hoja.Cells(5, 12) = "No.  JC"
Hoja.Cells(5, 13) = "No de Parte"
Hoja.Cells(5, 14) = "Cantidad"

Hoja.Cells(4, 15) = "Manual Der."
Hoja.Cells(5, 15) = "No.  JC"
Hoja.Cells(5, 16) = "No de Parte"
Hoja.Cells(5, 17) = "Cantidad"


Hoja.Range("A4:A5").Select
Hoja.Range("A4:A5").HorizontalAlignment = xlCenter
Hoja.Range("A4:A5").VerticalAlignment = xlBottom
Hoja.Range("A4:A5").MergeCells = True

Hoja.Range("B4:B5").Select
Hoja.Range("B4:B5").HorizontalAlignment = xlCenter
Hoja.Range("B4:B5").VerticalAlignment = xlBottom
Hoja.Range("B4:B5").MergeCells = True



Hoja.Range("A5:Z5").Select
Hoja.Range("A5:Z5").Font.Name = "Times New Roman"
Hoja.Range("A5:Z5").Font.Size = 12
Hoja.Range("A5:Z5").Font.Bold = True

'Linea 10
Hoja.Range("C4:E4").Select
Hoja.Range("C4:E4").HorizontalAlignment = xlCenter
Hoja.Range("C4:E4").MergeCells = True

'Linea 20
Hoja.Range("F4:H4").Select
Hoja.Range("F4:H4").HorizontalAlignment = xlCenter
Hoja.Range("F4:H4").MergeCells = True

'Linea 60
Hoja.Range("I4:K4").Select
Hoja.Range("I4:K4").HorizontalAlignment = xlCenter
Hoja.Range("I4:K4").MergeCells = True

'Manual Izq.
Hoja.Range("L4:N4").Select
Hoja.Range("L4:N4").HorizontalAlignment = xlCenter
Hoja.Range("L4:N4").MergeCells = True

'Manual Der.
Hoja.Range("O4:Q4").Select
Hoja.Range("O4:Q4").HorizontalAlignment = xlCenter
Hoja.Range("O4:Q4").MergeCells = True


Hoja.Range("A4:Q4").Select
Hoja.Range("A4:Q4").Font.Name = "Times New Roman"
Hoja.Range("A4:Q4").Font.Size = 11
Hoja.Range("A4:Q4").Font.Bold = True


Hoja.Range("A4:Q5").Select
Hoja.Range("A4:Q5").Interior.Pattern = xlSolid
Hoja.Range("A4:Q5").Interior.PatternColorIndex = xlAutomatic
Hoja.Range("A4:Q5").Interior.PatternTintAndShade = 0
Hoja.Range("A4:Q5").Interior.ThemeColor = xlThemeColorLight2
Hoja.Range("A4:Q5").Interior.TintAndShade = 0.799981688894314


' para los Codigos de MP
Renglon = 6
'''Bordes
Hoja.Range("A4:Q17").Select
Hoja.Range("A4:Q17").VerticalAlignment = xlCenter
Hoja.Range("A4:Q17").HorizontalAlignment = xlCenter

Hoja.Range("A4:Q17").Select

Hoja.Range("A4:Q17").Borders(xlDiagonalDown).LineStyle = xlNone

Hoja.Range("A4:Q17").Borders(xlDiagonalDown).LineStyle = xlNone
Hoja.Range("A4:Q17").Borders(xlDiagonalUp).LineStyle = xlNone

Hoja.Range("A4:Q17").Borders(xlEdgeLeft).LineStyle = xlContinuous
Hoja.Range("A4:Q17").Borders(xlEdgeLeft).ColorIndex = 0
Hoja.Range("A4:Q17").Borders(xlEdgeLeft).TintAndShade = 0
Hoja.Range("A4:Q17").Borders(xlEdgeLeft).Weight = xlThin

Hoja.Range("A4:Q17").Borders(xlEdgeBottom).LineStyle = xlContinuous
Hoja.Range("A4:Q17").Borders(xlEdgeBottom).ColorIndex = 0
Hoja.Range("A4:Q17").Borders(xlEdgeBottom).TintAndShade = 0
Hoja.Range("A4:Q17").Borders(xlEdgeBottom).Weight = xlThin

Hoja.Range("A4:Q17").Borders(xlEdgeRight).LineStyle = xlContinuous
Hoja.Range("A4:Q17").Borders(xlEdgeRight).ColorIndex = 0
Hoja.Range("A4:Q17").Borders(xlEdgeRight).TintAndShade = 0
Hoja.Range("A4:Q17").Borders(xlEdgeRight).Weight = xlThin

Hoja.Range("A4:Q17").Borders(xlEdgeTop).LineStyle = xlContinuous
Hoja.Range("A4:Q17").Borders(xlEdgeTop).ColorIndex = 0
Hoja.Range("A4:Q17").Borders(xlEdgeTop).TintAndShade = 0
Hoja.Range("A4:Q17").Borders(xlEdgeTop).Weight = xlThin

Hoja.Range("A4:Q17").Borders(xlInsideVertical).LineStyle = xlContinuous
Hoja.Range("A4:Q17").Borders(xlInsideVertical).ColorIndex = 0
Hoja.Range("A4:Q17").Borders(xlInsideVertical).TintAndShade = 0
Hoja.Range("A4:Q17").Borders(xlInsideVertical).Weight = xlThin

Hoja.Range("A4:Q17").Borders(xlInsideHorizontal).LineStyle = xlContinuous
Hoja.Range("A4:Q17").Borders(xlInsideHorizontal).ColorIndex = 0
Hoja.Range("A4:Q17").Borders(xlInsideHorizontal).TintAndShade = 0
Hoja.Range("A4:Q17").Borders(xlInsideHorizontal).Weight = xlThin

    
Hoja.Range("A5:Z5").Select
Hoja.Range("A5:Z5").Font.Name = "Times New Roman"
Hoja.Range("A5:Z5").Font.Size = 11
Hoja.Range("A5:Z5").Font.Strikethrough = False
Hoja.Range("A5:Z5").Font.Superscript = False
Hoja.Range("A5:Z5").Font.Subscript = False
Hoja.Range("A5:Z5").Font.OutlineFont = False
Hoja.Range("A5:Z5").Font.Shadow = False
Hoja.Range("A5:Z5").Font.Underline = xlUnderlineStyleNone
Hoja.Range("A5:Z5").Font.ThemeColor = xlThemeColorLight1
Hoja.Range("A5:Z5").Font.TintAndShade = 0
Hoja.Range("A5:Z5").Font.ThemeFont = xlThemeFontNone
Hoja.Range("A5:Z5").Font.Bold = True


'Ahora las Fechas
Cantidad = (Me.DTPF_Fin.Value - Me.DTPF_Ini.Value) + 1
Dim Inicial As Integer
Dim Final As Integer
Fecha = Me.DTPF_Ini.Value

Inicial = 6
i = 1
j = 0
Do While i <= (Cantidad * 24)
    If j = 0 Then
        Hoja.Cells(i + 5, 1) = Format(Fecha, "dd mmm yyyy")
        Hora = CDate("12:00:00 a.m.")
    Else
        If j > 12 Then
            Hora = CDate(j - 12 & ":00:00 p.m.")
        Else
            If j = 12 Then
                Hora = CDate("12:00:00 p.m.")
            Else
                Hora = CDate(j & ":00:00 a.m.")
            End If
        End If
    End If
    Hoja.Cells(i + 5, 2) = Hora
'''    Hoja.Cells(i + 5, 6) = Hora
'''    Hoja.Cells(i + 5, 10) = Hora
    If (i Mod 24) = 0 Then
        Hoja.Range("A" & (i - 18) & ":A" & (i + 5)).Select
        Hoja.Range("A" & (i - 18) & ":A" & (i + 5)).HorizontalAlignment = xlCenter
        Hoja.Range("A" & (i - 18) & ":A" & (i + 5)).VerticalAlignment = xlCenter
        Hoja.Range("A" & (i - 18) & ":A" & (i + 5)).WrapText = False
        Hoja.Range("A" & (i - 18) & ":A" & (i + 5)).Orientation = 90
        Hoja.Range("A" & (i - 18) & ":A" & (i + 5)).AddIndent = False
        Hoja.Range("A" & (i - 18) & ":A" & (i + 5)).IndentLevel = 0
        Hoja.Range("A" & (i - 18) & ":A" & (i + 5)).ShrinkToFit = False
        Hoja.Range("A" & (i - 18) & ":A" & (i + 5)).ReadingOrder = xlContext
        Hoja.Range("A" & (i - 18) & ":A" & (i + 5)).MergeCells = True
        
        Fecha = Fecha + 1
        j = 0
    Else
        j = j + 1
    End If
    i = i + 1
Loop

MaxRenglon = i - 1

'para toda las lineas
Renglon = 6
k = 2
Do While k <= 6
    CNN.CmdRepLineas (Me.DTPF_Ini.Value), (Me.DTPF_Fin.Value), (k)
    If CNN.rsCmdRepLineas.EOF <> True Then
        CNN.rsCmdRepLineas.MoveFirst
        m = 1
        Do While CNN.rsCmdRepLineas.EOF <> True
            Fecha = CNN.rsCmdRepLineas!Fecha
            IdHora = CNN.rsCmdRepLineas!IdHora
            OF = CNN.rsCmdRepLineas!OF
            CNN.CmdBuscaEsp2 (OF)
            If CNN.rsCmdBuscaEsp2.EOF <> True Then
                NoParte = CNN.rsCmdBuscaEsp2!nodeparte
            Else
                CNN.CmdVidrio_MPS (OF)
                If CNN.rsCmdVidrio_MPS.EOF <> True Then
                    NoParte = CNN.rsCmdVidrio_MPS!Codigo
                Else
                    MsgBox "El PT o V MPS nO existe", vbInformation, MSG
                    NoParte = "Vacio"
                End If
                CNN.rsCmdVidrio_MPS.Close
            End If
            CNN.rsCmdBuscaEsp2.Close
            Renglon = (6 + (Fecha - Me.DTPF_Ini.Value) * 24) + IdHora
            
            Select Case k
                Case 4
                    Hoja.Cells(Renglon, 4) = NoParte
                    If Mid(OF, 1, 1) <> "X" Then
                        Hoja.Cells(Renglon, 3) = CNN.rsCmdRepLineas!jc
                        Hoja.Cells(Renglon, 5) = CNN.rsCmdRepLineas!pzspt
                    End If
                Case 5
                    Hoja.Cells(Renglon, 7) = NoParte
                    If Mid(OF, 1, 1) <> "X" Then
                        Hoja.Cells(Renglon, 6) = CNN.rsCmdRepLineas!jc
                        Hoja.Cells(Renglon, 8) = CNN.rsCmdRepLineas!pzspt
                    End If
               Case 6
                    Hoja.Cells(Renglon, 10) = NoParte
                    If Mid(OF, 1, 1) <> "X" Then
                        Hoja.Cells(Renglon, 9) = CNN.rsCmdRepLineas!jc
                        Hoja.Cells(Renglon, 11) = CNN.rsCmdRepLineas!pzspt
                    End If
                  
                Case 2
                    Hoja.Cells(Renglon, 13) = NoParte
                    If Mid(OF, 1, 1) <> "X" Then
                        Hoja.Cells(Renglon, 12) = CNN.rsCmdRepLineas!jc
                        Hoja.Cells(Renglon, 14) = CNN.rsCmdRepLineas!pzspt
                    End If
              
               Case 3
                    Hoja.Cells(Renglon, 16) = NoParte
                    If Mid(OF, 1, 1) <> "X" Then
                        Hoja.Cells(Renglon, 15) = CNN.rsCmdRepLineas!jc
                        Hoja.Cells(Renglon, 17) = CNN.rsCmdRepLineas!pzspt
                    End If
                  
            End Select
            CNN.rsCmdRepLineas.MoveNext
        Loop
    End If
    CNN.rsCmdRepLineas.Close
    k = k + 1
Loop


Hoja.Range("B6:Q" & MaxRenglon).Select
Hoja.Range("B6:Q" & MaxRenglon).HorizontalAlignment = xlCenter
Hoja.Range("B6:Q" & MaxRenglon).VerticalAlignment = xlCenter

Hoja.Rows("6:" & MaxRenglon + 5).RowHeight = 20.25
Hoja.Columns("A:Q").EntireColumn.AutoFit
Hoja.Rows("3:3").RowHeight = 27.25

Hoja.Range("A4:Q" & MaxRenglon + 5).Select
Hoja.Range("A4:Q" & MaxRenglon + 5).VerticalAlignment = xlCenter
Hoja.Range("A4:Q" & MaxRenglon + 5).HorizontalAlignment = xlCenter
Hoja.Range("A4:Q" & MaxRenglon + 5).WrapText = False
Hoja.Range("A4:Q" & MaxRenglon + 5).AddIndent = False
Hoja.Range("A4:Q" & MaxRenglon + 5).IndentLevel = 0
Hoja.Range("A4:Q" & MaxRenglon + 5).ShrinkToFit = False
Hoja.Range("A4:Q" & MaxRenglon + 5).ReadingOrder = xlContext
    
    
Hoja.Range("A4:Q" & MaxRenglon + 5).Borders(xlDiagonalDown).LineStyle = xlNone
Hoja.Range("A4:Q" & MaxRenglon + 5).Borders(xlDiagonalUp).LineStyle = xlNone

Hoja.Range("A4:Q" & MaxRenglon + 5).Borders(xlEdgeLeft).LineStyle = xlContinuous
Hoja.Range("A4:Q" & MaxRenglon + 5).Borders(xlEdgeLeft).ColorIndex = 0
Hoja.Range("A4:Q" & MaxRenglon + 5).Borders(xlEdgeLeft).TintAndShade = 0
Hoja.Range("A4:Q" & MaxRenglon + 5).Borders(xlEdgeLeft).Weight = xlThin

Hoja.Range("A4:Q" & MaxRenglon + 5).Borders(xlEdgeBottom).LineStyle = xlContinuous
Hoja.Range("A4:Q" & MaxRenglon + 5).Borders(xlEdgeBottom).ColorIndex = 0
Hoja.Range("A4:Q" & MaxRenglon + 5).Borders(xlEdgeBottom).TintAndShade = 0
Hoja.Range("A4:Q" & MaxRenglon + 5).Borders(xlEdgeBottom).Weight = xlThin

Hoja.Range("A4:Q" & MaxRenglon + 5).Borders(xlEdgeRight).LineStyle = xlContinuous
Hoja.Range("A4:Q" & MaxRenglon + 5).Borders(xlEdgeRight).ColorIndex = 0
Hoja.Range("A4:Q" & MaxRenglon + 5).Borders(xlEdgeRight).TintAndShade = 0
Hoja.Range("A4:Q" & MaxRenglon + 5).Borders(xlEdgeRight).Weight = xlThin

Hoja.Range("A4:Q" & MaxRenglon + 5).Borders(xlEdgeTop).LineStyle = xlContinuous
Hoja.Range("A4:Q" & MaxRenglon + 5).Borders(xlEdgeTop).ColorIndex = 0
Hoja.Range("A4:Q" & MaxRenglon + 5).Borders(xlEdgeTop).TintAndShade = 0
Hoja.Range("A4:Q" & MaxRenglon + 5).Borders(xlEdgeTop).Weight = xlThin

Hoja.Range("A4:Q" & MaxRenglon + 5).Borders(xlInsideVertical).LineStyle = xlContinuous
Hoja.Range("A4:Q" & MaxRenglon + 5).Borders(xlInsideVertical).ColorIndex = 0
Hoja.Range("A4:Q" & MaxRenglon + 5).Borders(xlInsideVertical).TintAndShade = 0
Hoja.Range("A4:Q" & MaxRenglon + 5).Borders(xlInsideVertical).Weight = xlThin

Hoja.Range("A4:Q" & MaxRenglon + 5).Borders(xlInsideHorizontal).LineStyle = xlContinuous
Hoja.Range("A4:Q" & MaxRenglon + 5).Borders(xlInsideHorizontal).ColorIndex = 0
Hoja.Range("A4:Q" & MaxRenglon + 5).Borders(xlInsideHorizontal).TintAndShade = 0
Hoja.Range("A4:Q" & MaxRenglon + 5).Borders(xlInsideHorizontal).Weight = xlThin


'hoja.Sheets(Resp).Cells(1, 1) = "Gemtron de México S.A. de C.V."
Hoja.Cells(1, 1) = "Gemtron de México S.A. de C.V."

Hoja.Range("A1:k1").Select
Hoja.Range("A1:k1").HorizontalAlignment = xlCenter
Hoja.Range("A1:k1").VerticalAlignment = xlBottom
Hoja.Range("A1:k1").WrapText = False
Hoja.Range("A1:k1").MergeCells = True

Hoja.Range("A1:k1").Select
Hoja.Range("A1:k1").Font.Name = "Times New Roman"
Hoja.Range("A1:k1").Font.Size = 16
Hoja.Range("A1:k1").Font.Bold = True


Hoja.Cells(2, 1) = "Reporte Programa de Templado"
Hoja.Range("A2:K2").Select
Hoja.Range("A2:K2").HorizontalAlignment = xlCenter
Hoja.Range("A2:K2").VerticalAlignment = xlBottom
Hoja.Range("A2:K2").WrapText = False
Hoja.Range("A2:K2").MergeCells = True

Hoja.Range("A2:K2").Select
Hoja.Range("A2:K2").Font.Name = "Times New Roman"
Hoja.Range("A2:K2").Font.Size = 12
Hoja.Range("A2:K2").Font.Bold = True
Hoja.Range("A2:K2").Font.Italic = True

libro.ActiveSheet.PageSetup.LeftHeaderPicture.FileName = "C:\legacy\schott.JPG"
libro.ActiveSheet.PageSetup.PrintTitleRows = "$1:$5"
libro.ActiveSheet.PageSetup.PrintTitleColumns = ""

libro.ActiveSheet.PageSetup.PrintArea = ""
libro.ActiveSheet.PageSetup.LeftHeader = "&G"
libro.ActiveSheet.PageSetup.CenterHeader = ""
libro.ActiveSheet.PageSetup.RightHeader = ""

NoFormaISO = "Forma #F-8.5.2-1-W (Jul.20'20)"

libro.ActiveSheet.PageSetup.LeftFooter = NoFormaISO & vbNewLine & "Imprimio: " & Nombre '& vbNewLine    '& "FORMA #F-TEM-020-A(08/17/2009)"
libro.ActiveSheet.PageSetup.CenterFooter = "&P"
libro.ActiveSheet.PageSetup.RightFooter = "&T" & Chr(10) & "&D"
libro.ActiveSheet.PageSetup.LeftMargin = libro.Application.InchesToPoints(0.236220472440945)
libro.ActiveSheet.PageSetup.RightMargin = libro.Application.InchesToPoints(0.275590551181102)
libro.ActiveSheet.PageSetup.TopMargin = libro.Application.InchesToPoints(0.275590551181102)
libro.ActiveSheet.PageSetup.BottomMargin = libro.Application.InchesToPoints(0.98551181102362)
libro.ActiveSheet.PageSetup.HeaderMargin = libro.Application.InchesToPoints(0.275590551181102)
libro.ActiveSheet.PageSetup.FooterMargin = libro.Application.InchesToPoints(0.31496062992126)
libro.ActiveSheet.PageSetup.PrintHeadings = False
libro.ActiveSheet.PageSetup.PrintGridlines = False
libro.ActiveSheet.PageSetup.PrintComments = xlPrintNoComments
libro.ActiveSheet.PageSetup.PrintQuality = 600
libro.ActiveSheet.PageSetup.CenterHorizontally = False
libro.ActiveSheet.PageSetup.CenterVertically = False
libro.ActiveSheet.PageSetup.Orientation = xlLandscape  ' xlPortrait
libro.ActiveSheet.PageSetup.Draft = False
libro.ActiveSheet.PageSetup.PaperSize = xlPaperLetter    'xlPaperLegal  '
libro.ActiveSheet.PageSetup.FirstPageNumber = xlAutomatic
libro.ActiveSheet.PageSetup.Order = xlDownThenOver
libro.ActiveSheet.PageSetup.BlackAndWhite = False
libro.ActiveSheet.PageSetup.Zoom = 77
libro.ActiveSheet.PageSetup.PrintErrors = xlPrintErrorsDisplayed
libro.ActiveSheet.PageSetup.OddAndEvenPagesHeaderFooter = False
libro.ActiveSheet.PageSetup.DifferentFirstPageHeaderFooter = False
libro.ActiveSheet.PageSetup.ScaleWithDocHeaderFooter = True
libro.ActiveSheet.PageSetup.AlignMarginsHeaderFooter = True
libro.ActiveSheet.PageSetup.EvenPage.LeftHeader.Text = ""
libro.ActiveSheet.PageSetup.EvenPage.CenterHeader.Text = ""
libro.ActiveSheet.PageSetup.EvenPage.RightHeader.Text = ""
libro.ActiveSheet.PageSetup.EvenPage.LeftFooter.Text = ""
libro.ActiveSheet.PageSetup.EvenPage.CenterFooter.Text = ""
libro.ActiveSheet.PageSetup.EvenPage.RightFooter.Text = ""
libro.ActiveSheet.PageSetup.FirstPage.LeftHeader.Text = ""
libro.ActiveSheet.PageSetup.FirstPage.CenterHeader.Text = ""
libro.ActiveSheet.PageSetup.FirstPage.RightHeader.Text = ""
libro.ActiveSheet.PageSetup.FirstPage.LeftFooter.Text = ""
libro.ActiveSheet.PageSetup.FirstPage.CenterFooter.Text = ""
libro.ActiveSheet.PageSetup.FirstPage.RightFooter.Text = ""

Hoja.Range("K6").Select
Randomize

Cadena = "PTemp" & Format(Day(Date), "00") & Format(Month(Date), "00") & Format(Year(Date), "00") & "_" & Int((1000 * Rnd) + 1)
Ruta = "C:\mis documentos\" & Cadena & ".xlsx"

Hoja.SaveAs (Ruta)
libro.Close
Set Hoja = Nothing
Set libro = Nothing
vApp.Quit
Set vApp = Nothing

Dim CnnExcel As New Excel.Application
Dim LibroPDF As Excel.Workbook

'Abrimos el documento
RutaPDF = "C:\mis documentos\" & Cadena & ".pdf"
If Dir(RutaPDF) <> "" Then
    Kill RutaPDF
End If

CnnExcel.Workbooks.Open (Ruta) 'Guarda Local
CnnExcel.ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=RutaPDF, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
CnnExcel.ActiveWorkbook.Close
'''CnnExcel.LibroPDF.Close
CnnExcel.Quit
Set CnnExcel = Nothing

RutaPDF2 = "\\10.18.164.207\Adjuntos$\" & Cadena & ".pdf"
If Dir(RutaPDF2) <> "" Then
    Kill RutaPDF2
End If

CnnExcel.Workbooks.Open (Ruta)
CnnExcel.ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:= _
    RutaPDF2, Quality:= _
    xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
    OpenAfterPublish:=False
CnnExcel.ActiveWorkbook.Close
'''CnnExcel.LibroPDF.Close
CnnExcel.Quit
Set CnnExcel = Nothing
'MsgBox "Tarea Terminada" & vbNewLine & "Su archivo esta en la sig ruta: " & Ruta & ".xlsx", vbInformation, MSG
Cadena = "Programa de Templado " & Me.DTPF_Ini.Value & " a " & Me.DTPF_Fin.Value
CadenaCorreo = "Se adjunta el Programa de Templado<br><br>"
'CadenaCorreo = CadenaCorreo & "<a href='" & RutaPDF & "'>" & RutaPDF & "</a> <br><br>"
CadenaCorreo = CadenaCorreo & "Gracias"

Randomize

'Cadena = "PFLAT " & Format(Day(Date), "00") & Format(Month(Date), "00") & Format(Year(Date), "00") & " " & Int((1000 * Rnd) + 1)

RutaPDF = "Q:\Production\Read\Reps\ProgTempFab" & Format(Day(Date), "00") & Format(Month(Date), "00") & Format(Year(Date), "00") & "" & Int((1000 * Rnd) + 1) & ".pdf"

CnnExcel.Workbooks.Open (Ruta)
'pause (2000)
CnnExcel.ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:= _
    RutaPDF, Quality:= _
    xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
    OpenAfterPublish:=False
CnnExcel.ActiveWorkbook.Close
'''CnnExcel.LibroPDF.Close
CnnExcel.Quit
Set CnnExcel = Nothing

CNN.CmdReportes ("RepTempGral")
If CNN.rsCmdReportes.EOF <> True Then
        CNN.rsCmdReportes!rutareporte = RutaPDF
    CNN.rsCmdReportes.Update
End If
CNN.rsCmdReportes.Close

''''Para = "cesar.reyes@schott.com, "
''''Para = Para & " ruben.salas@schott.com; manuel.almendares@schott.com, manuel.arriaga@schott.com, jorge.morales@schott.com, fermin-lael.villanueva-guerra@schott.com,  "
''''
''''CNN.rsCmdCorreo.Open
''''CNN.rsCmdCorreo.AddNew
''''    CNN.rsCmdCorreo!De = "cesar.reyes@schott.com" 'De
''''    CNN.rsCmdCorreo!Para = Para
''''    CNN.rsCmdCorreo!CC = "alejandro.cerda@schott.com, blas.reyes@schott.com, alejandro.portillo@schott.com, francisco.lucio@schott.com, "
''''    CNN.rsCmdCorreo!CCO = "helios.ireta@schott.com, blas.reyes@schott.com, regino.rivera@schott.com"
''''    CNN.rsCmdCorreo!Titulo = Cadena '& " Pruebas TI"
''''    CNN.rsCmdCorreo!Mensaje = CadenaCorreo
''''    CNN.rsCmdCorreo!aplicacion = "JC"
''''    CNN.rsCmdCorreo!Prioridad = 2
''''    CNN.rsCmdCorreo!Ruta = RutaPDF2
''''CNN.rsCmdCorreo.Update
''''CNN.rsCmdCorreo.Close
''''
''''MsgBox "Tarea Terminada " & vbNewLine & "Su archivo esta en la sig ruta: " & Ruta & " " & vbNewLine & vbNewLine & " El Reporte de Templado ya se publico. ", vbInformation, MSG


Para = "breyes@sswtechnologies.com; fvillanueva@sswtechnologies.com, creyes@sswtechnologies.com, "
Para = Para & " flucio@sswtechnologies.com, hhernandez@sswtechnologies.com; jibarra@sswtechnologies.com, jalmendarez@sswtechnologies.com,  jmruiz@sswtechnologies.com; "
Para = Para & " marriaga@sswtechnologies.com, rsalas@sswtechnologies.com, sneri@sswtechnologies.com, jarroyos@sswtechnologies.com, aportillo@sswtechnologies.com, " '

CNN.rsCmdCorreo.Open
CNN.rsCmdCorreo.AddNew
    CNN.rsCmdCorreo!De = "cesar.reyes@schott.com" 'De
    CNN.rsCmdCorreo!Para = Para
    CNN.rsCmdCorreo!CC = "hireta@sswtechnologies.com, vcastillo@sswtechnologies.com"
    CNN.rsCmdCorreo!CCO = ""
    CNN.rsCmdCorreo!Titulo = Cadena '& " Pruebas TI"
    CNN.rsCmdCorreo!Mensaje = CadenaCorreo
    CNN.rsCmdCorreo!aplicacion = "JC"
    CNN.rsCmdCorreo!Prioridad = 2
    CNN.rsCmdCorreo!Ruta = RutaPDF2
CNN.rsCmdCorreo.Update
CNN.rsCmdCorreo.Close

MsgBox "Tarea Terminada " & vbNewLine & "Su archivo esta en la sig ruta: " & Ruta & " " & vbNewLine & vbNewLine & " El Reporte de Produccion ya se publico. ", vbInformation, MSG


MsgBox "El correo del Programa tambien se envio a los siguientes usuarios: " & vbNewLine & Para & ", vbInformation, MSG"




Unload Me

End Sub

