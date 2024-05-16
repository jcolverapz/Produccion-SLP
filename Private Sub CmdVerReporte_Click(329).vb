Private Sub CmdVerReporte_Click()
'Reporte de Produccion Templado

'Valida fecha
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

Dim NombreProduccion As String

Renglon = 4
Columna = 1
'hoja.Sheets(1).Select
Hoja.Select
Resp = "Programa " & Format(Me.DTPF_Ini.Value, "d mmm yy") & " a " & Format(Me.DTPF_Fin.Value, "d mmm yy")
'hoja.Sheets(1).Name = Resp
NombreProduccion = Resp
Hoja.Name = Resp
'hoja.Sheets(Resp).Cells(1, 1) = "Gemtron de México S.A. de C.V."
Hoja.Cells(1, 1) = "Gemtron de México S.A. de C.V."

'vApp.Visible = True

'Titulo
Hoja.Range("A1:N1").Select
Hoja.Range("A1:N1").HorizontalAlignment = xlCenter
Hoja.Range("A1:N1").VerticalAlignment = xlBottom
Hoja.Range("A1:N1").WrapText = False
Hoja.Range("A1:N1").MergeCells = True

Hoja.Range("A1:N1").Select
Hoja.Range("A1:N1").Font.Name = "Times New Roman"
Hoja.Range("A1:N1").Font.Size = 16
Hoja.Range("A1:N1").Font.Bold = True


Hoja.Cells(2, 1) = "Reporte Programa de Producción"
Hoja.Range("A2:N2").Select
Hoja.Range("A2:N2").HorizontalAlignment = xlCenter
Hoja.Range("A2:N2").VerticalAlignment = xlBottom
Hoja.Range("A2:N2").WrapText = False
Hoja.Range("A2:N2").MergeCells = True

Hoja.Range("A2:N2").Select
Hoja.Range("A2:N2").Font.Name = "Times New Roman"
Hoja.Range("A2:N2").Font.Size = 12
Hoja.Range("A2:N2").Font.Bold = True
Hoja.Range("A2:N2").Font.Italic = True

'Encabezados


Hoja.Cells(4, 1) = "Fecha"
Hoja.Cells(5, 2) = "Hora"

Hoja.Cells(4, 3) = "Linea 1"
Hoja.Cells(5, 3) = "Tipo"
Hoja.Cells(5, 4) = "No.  JC"
Hoja.Cells(5, 5) = "No de Parte"
Hoja.Cells(5, 6) = "Cantidad"
Hoja.Cells(5, 7) = "No. OMP SAP"


Hoja.Cells(4, 8) = "Linea 2"
Hoja.Cells(5, 8) = "Tipo"
Hoja.Cells(5, 9) = "No.  JC"
Hoja.Cells(5, 10) = "No de Parte"
Hoja.Cells(5, 11) = "Cantidad"
Hoja.Cells(5, 12) = "No. OMP SAP"

Hoja.Cells(4, 13) = "Linea 3"
Hoja.Cells(5, 13) = "Tipo"
Hoja.Cells(5, 14) = "No.  JC"
Hoja.Cells(5, 15) = "No de Parte"
Hoja.Cells(5, 16) = "Cantidad"
Hoja.Cells(5, 17) = "No. OMP SAP"

Hoja.Cells(4, 18) = "Linea 4"
Hoja.Cells(5, 18) = "Tipo"
Hoja.Cells(5, 19) = "No.  JC"
Hoja.Cells(5, 20) = "No de Parte"
Hoja.Cells(5, 21) = "Cantidad"
Hoja.Cells(5, 22) = "No. OMP SAP"
 

Hoja.Range("A4:A5").Select
Hoja.Range("A4:A5").HorizontalAlignment = xlCenter
Hoja.Range("A4:A5").VerticalAlignment = xlBottom
Hoja.Range("A4:A5").MergeCells = True

Hoja.Range("B4:B5").Select
Hoja.Range("B4:B5").HorizontalAlignment = xlCenter
Hoja.Range("B4:B5").VerticalAlignment = xlBottom
Hoja.Range("B4:B5").MergeCells = True

Hoja.Range("A5:k5").Select
Hoja.Range("A5:k5").Font.Name = "Times New Roman"
Hoja.Range("A5:k5").Font.Size = 12
Hoja.Range("A5:k5").Font.Bold = True

'Linea 1
Hoja.Range("C4:G4").Select
Hoja.Range("C4:G4").HorizontalAlignment = xlCenter
Hoja.Range("C4:G4").MergeCells = True

'Linea 2
Hoja.Range("H4:L4").Select
Hoja.Range("H4:L4").HorizontalAlignment = xlCenter
Hoja.Range("H4:L4").MergeCells = True

'Linea 3
Hoja.Range("M4:Q4").Select
Hoja.Range("M4:Q4").HorizontalAlignment = xlCenter
Hoja.Range("M4:Q4").MergeCells = True



'Linea 4
Hoja.Range("R4:V4").Select
Hoja.Range("R4:V4").HorizontalAlignment = xlCenter
Hoja.Range("R4:V4").MergeCells = True


'Formato
Hoja.Range("A4:V4").Select
Hoja.Range("A4:V4").Font.Name = "Times New Roman"
Hoja.Range("A4:V4").Font.Size = 11
Hoja.Range("A4:V4").Font.Bold = True


Hoja.Range("A4:V5").Select
Hoja.Range("A4:V5").Interior.Pattern = xlSolid
Hoja.Range("A4:V5").Interior.PatternColorIndex = xlAutomatic
Hoja.Range("A4:V5").Interior.PatternTintAndShade = 0
Hoja.Range("A4:V5").Interior.ThemeColor = xlThemeColorLight2
Hoja.Range("A4:V5").Interior.TintAndShade = 0.799981688894314


' para los Codigos de MP
Renglon = 6
'''Bordes
Hoja.Range("A4:V5").Select
Hoja.Range("A4:V5").VerticalAlignment = xlCenter
Hoja.Range("A4:V5").HorizontalAlignment = xlCenter

Hoja.Range("A4:V5").Select

Hoja.Range("A4:V5").Borders(xlDiagonalDown).LineStyle = xlNone

Hoja.Range("A4:V5").Borders(xlDiagonalDown).LineStyle = xlNone
Hoja.Range("A4:V5").Borders(xlDiagonalUp).LineStyle = xlNone

Hoja.Range("A4:V5").Borders(xlEdgeLeft).LineStyle = xlContinuous
Hoja.Range("A4:V5").Borders(xlEdgeLeft).ColorIndex = 0
Hoja.Range("A4:V5").Borders(xlEdgeLeft).TintAndShade = 0
Hoja.Range("A4:V5").Borders(xlEdgeLeft).Weight = xlThin

Hoja.Range("A4:V5").Borders(xlEdgeBottom).LineStyle = xlContinuous
Hoja.Range("A4:V5").Borders(xlEdgeBottom).ColorIndex = 0
Hoja.Range("A4:V5").Borders(xlEdgeBottom).TintAndShade = 0
Hoja.Range("A4:V5").Borders(xlEdgeBottom).Weight = xlThin

Hoja.Range("A4:V5").Borders(xlEdgeRight).LineStyle = xlContinuous
Hoja.Range("A4:V5").Borders(xlEdgeRight).ColorIndex = 0
Hoja.Range("A4:V5").Borders(xlEdgeRight).TintAndShade = 0
Hoja.Range("A4:V5").Borders(xlEdgeRight).Weight = xlThin

Hoja.Range("A4:V5").Borders(xlEdgeTop).LineStyle = xlContinuous
Hoja.Range("A4:V5").Borders(xlEdgeTop).ColorIndex = 0
Hoja.Range("A4:V5").Borders(xlEdgeTop).TintAndShade = 0
Hoja.Range("A4:V5").Borders(xlEdgeTop).Weight = xlThin

Hoja.Range("A4:V5").Borders(xlInsideVertical).LineStyle = xlContinuous
Hoja.Range("A4:V5").Borders(xlInsideVertical).ColorIndex = 0
Hoja.Range("A4:V5").Borders(xlInsideVertical).TintAndShade = 0
Hoja.Range("A4:V5").Borders(xlInsideVertical).Weight = xlThin

Hoja.Range("A4:V5").Borders(xlInsideHorizontal).LineStyle = xlContinuous
Hoja.Range("A4:V5").Borders(xlInsideHorizontal).ColorIndex = 0
Hoja.Range("A4:V5").Borders(xlInsideHorizontal).TintAndShade = 0
Hoja.Range("A4:V5").Borders(xlInsideHorizontal).Weight = xlThin

    
Hoja.Range("A5:V5").Select
Hoja.Range("A5:V5").Font.Name = "Times New Roman"
Hoja.Range("A5:V5").Font.Size = 11
Hoja.Range("A5:V5").Font.Strikethrough = False
Hoja.Range("A5:V5").Font.Superscript = False
Hoja.Range("A5:V5").Font.Subscript = False
Hoja.Range("A5:V5").Font.OutlineFont = False
Hoja.Range("A5:V5").Font.Shadow = False
Hoja.Range("A5:V5").Font.Underline = xlUnderlineStyleNone
Hoja.Range("A5:V5").Font.ThemeColor = xlThemeColorLight1
Hoja.Range("A5:V5").Font.TintAndShade = 0
Hoja.Range("A5:V5").Font.ThemeFont = xlThemeFontNone
Hoja.Range("A5:V5").Font.Bold = True


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
        Hora = "12 a.m." & "-" & "1 a.m."
        'Hora = CDate("12:00:00 a.m.") & "-" & CDate("01:00:00 a.m.")
    Else
        If j > 12 Then
            
            Hora = j - 12 & " p.m." & "-" & j - 11 & " p.m."
            
            'Hora = CDate(j - 13 & ":00:00 p.m.") & "-" & CDate(j - 12 & ":00:00 p.m.")
            
            'If j = 13 Then
             '     Hora = j - 12 & " p.m." & "-" & j - 11 & " p.m."
            'End If
            
        Else
            If j = 12 Then
                Hora = "12 p.m." & "-" & "1 p.m."
               ' Hora = CDate("12:00:00 p.m.")
'               If j = 12 Then
'                  Hora = "12 p.m." & "-" & "1 p.m."
'               End If
            Else
                Hora = j & " a.m." & "-" & j + 1 & " a.m."
                If j = 11 Then
                  Hora = j & " a.m." & "-" & j + 1 & " p.m."
                End If
                
                'Hora = CDate(j & ":00:00 a.m.") & "-" & CDate(j + 1 & ":00:00 a.m.")
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

F_Ini = Me.DTPF_Ini.Value

'para toda las lineas
Renglon = 6
'k = 111
'Do While k <= 331
    
    '.rsCmdRepLineas.
    CNN.CmdRepLineasGral (Me.DTPF_Ini.Value), (Me.DTPF_Fin.Value) ', (k)
    If CNN.rsCmdRepLineasGral.EOF <> True Then
    ' CNN.rsCmdRepLineasGral.Close
        CNN.rsCmdRepLineasGral.MoveFirst
        m = 1
        Do While CNN.rsCmdRepLineasGral.EOF <> True
            Fecha = CNN.rsCmdRepLineasGral!Fecha
            IdHora = CNN.rsCmdRepLineasGral!IdHora
            OF = CNN.rsCmdRepLineasGral!OF
            
            CNN.CmdBuscaEsp2 (OF)
            If CNN.rsCmdBuscaEsp2.EOF <> True Then
                NoParte = CNN.rsCmdBuscaEsp2!nodeparte
            Else
                CNN.CmdVidrio_MPS (OF)
                If CNN.rsCmdVidrio_MPS.EOF <> True Then
                    NoParte = CNN.rsCmdVidrio_MPS!Codigo
                Else
                    MsgBox "El PT o V MPS NO existe", vbInformation, MSG
                    NoParte = "Vacio"
                End If
                CNN.rsCmdVidrio_MPS.Close
            End If
            CNN.rsCmdBuscaEsp2.Close
            Renglon = (6 + (Fecha - Me.DTPF_Ini.Value) * 24) + IdHora
            
            
            If InStr(1, CNN.rsCmdRepLineasGral!Descripcion, "1") Then
                k = 1
            End If
            If InStr(1, CNN.rsCmdRepLineasGral!Descripcion, "2") Then
                k = 2
            End If
            If InStr(1, CNN.rsCmdRepLineasGral!Descripcion, "3") Then
                k = 3
            End If
             If InStr(1, CNN.rsCmdRepLineasGral!Descripcion, "4") Then
                k = 4
            End If
            
            
            Select Case k
            
                Case 1 '111
                    Hoja.Cells(Renglon, 3) = CNN.rsCmdRepLineasGral!Descripcion
                    Hoja.Cells(Renglon, 5) = NoParte
                    
                    If Mid(OF, 1, 1) <> "X" Then
                    
                        Hoja.Cells(Renglon, 4) = CNN.rsCmdRepLineasGral!jc
                        
                        If IdHora = 22 Or IdHora = 16 Then
                            Hoja.Cells(Renglon, 6) = CNN.rsCmdRepLineasGral!pzspt / 2
                        Else
                        
                             If IdHora = 23 Then
                                 If CNN.rsCmdRepLineasGral!jc = Hoja.Cells((Renglon - 1), 4) Then
                                 
                                     Hoja.Cells(Renglon, 6) = Hoja.Cells(Renglon - 1, 6) + CNN.rsCmdRepLineasGral!pzspt
                                Else
                                     Hoja.Cells(Renglon, 6) = CNN.rsCmdRepLineasGral!pzspt
                                 End If
                             Else
                                 Hoja.Cells(Renglon, 6) = CNN.rsCmdRepLineasGral!pzspt
                            End If
                            
                        End If
                        
                        
                    End If
                    
                     Hoja.Cells(Renglon, 7) = CNN.rsCmdRepLineasGral!OrdenProd_SAP
                     
                Case 2 '221
                
                    Hoja.Cells(Renglon, 8) = CNN.rsCmdRepLineasGral!Descripcion
                    Hoja.Cells(Renglon, 10) = NoParte
                    
                    If Mid(OF, 1, 1) <> "X" Then
                    
                        Hoja.Cells(Renglon, 9) = CNN.rsCmdRepLineasGral!jc
                        
                        If IdHora = 22 Or IdHora = 16 Then
                            Hoja.Cells(Renglon, 10) = CNN.rsCmdRepLineasGral!pzspt / 2
                        Else
                             If IdHora = 23 Then
                             
                                 If CNN.rsCmdRepLineasGral!jc = Hoja.Cells((Renglon - 1), 9) Then
                                 
                                     Hoja.Cells(Renglon, 11) = Hoja.Cells((Renglon - 1), 11) + CNN.rsCmdRepLineasGral!pzspt
                                Else
                                     Hoja.Cells(Renglon, 11) = CNN.rsCmdRepLineasGral!pzspt
                                 End If
                             Else
                                 Hoja.Cells(Renglon, 11) = CNN.rsCmdRepLineasGral!pzspt
                            End If
                        End If
                        
                        
                        
                        Hoja.Cells(Renglon, 11) = CNN.rsCmdRepLineasGral!pzspt
                    End If
                    
                 Hoja.Cells(Renglon, 12) = CNN.rsCmdRepLineasGral!OrdenProd_SAP
                 
                Case 3 '331
                    Hoja.Cells(Renglon, 13) = CNN.rsCmdRepLineasGral!Descripcion
                    Hoja.Cells(Renglon, 15) = NoParte
                    
                    If Mid(OF, 1, 1) <> "X" Then
                    
                        Hoja.Cells(Renglon, 14) = CNN.rsCmdRepLineasGral!jc
                        
                        
                        If IdHora = 22 Or IdHora = 16 Then
                            Hoja.Cells(Renglon, 16) = CNN.rsCmdRepLineasGral!pzspt / 2
                        Else
                        
                             If IdHora = 23 Then
                             
                                 If CNN.rsCmdRepLineasGral!jc = Hoja.Cells((Renglon - 1), 14) Then
                                 
                                     Hoja.Cells(Renglon, 16) = Hoja.Cells(Renglon - 1, 16) + CNN.rsCmdRepLineasGral!pzspt
                                Else
                                     Hoja.Cells(Renglon, 16) = CNN.rsCmdRepLineasGral!pzspt
                                 End If
                             Else
                                 Hoja.Cells(Renglon, 16) = CNN.rsCmdRepLineasGral!pzspt
                            End If
                        End If
                        
                        
                    End If
                  Hoja.Cells(Renglon, 17) = CNN.rsCmdRepLineasGral!OrdenProd_SAP
                  
                Case 4 '441
                    Hoja.Cells(Renglon, 18) = CNN.rsCmdRepLineasGral!Descripcion
                    Hoja.Cells(Renglon, 20) = NoParte
                    
                    If Mid(OF, 1, 1) <> "X" Then
                    
                        Hoja.Cells(Renglon, 19) = CNN.rsCmdRepLineasGral!jc
                        
                        If IdHora = 22 Or IdHora = 16 Then
                            Hoja.Cells(Renglon, 21) = CNN.rsCmdRepLineasGral!pzspt / 2
                        Else
                             If IdHora = 23 Then
                             
                                 If CNN.rsCmdRepLineasGral!jc = Hoja.Cells((Renglon - 1), 19) Then
                                 
                                     Hoja.Cells(Renglon, 21) = Hoja.Cells(Renglon - 1, 21) + CNN.rsCmdRepLineasGral!pzspt
                                Else
                                     Hoja.Cells(Renglon, 21) = CNN.rsCmdRepLineasGral!pzspt
                                 End If
                             Else
                                 Hoja.Cells(Renglon, 21) = CNN.rsCmdRepLineasGral!pzspt
                            End If
                        End If
                        
                        
                        Hoja.Cells(Renglon, 21) = CNN.rsCmdRepLineasGral!pzspt
                    End If
                    
                    Hoja.Cells(Renglon, 22) = CNN.rsCmdRepLineasGral!OrdenProd_SAP
                    
            End Select
            CNN.rsCmdRepLineasGral.MoveNext
        Loop
    End If
    CNN.rsCmdRepLineasGral.Close

'    k = k + 110
'Loop



Hoja.Range("B6:V" & MaxRenglon).Select
Hoja.Range("B6:V" & MaxRenglon).HorizontalAlignment = xlCenter
Hoja.Range("B6:V" & MaxRenglon).VerticalAlignment = xlCenter

Hoja.Rows("6:" & MaxRenglon + 5).RowHeight = 20.25
Hoja.Columns("A:V").EntireColumn.AutoFit
Hoja.Rows("3:3").RowHeight = 27.25

Hoja.Range("A4:V" & MaxRenglon + 5).Select
Hoja.Range("A4:V" & MaxRenglon + 5).VerticalAlignment = xlCenter
Hoja.Range("A4:V" & MaxRenglon + 5).HorizontalAlignment = xlCenter
Hoja.Range("A4:V" & MaxRenglon + 5).WrapText = False
Hoja.Range("A4:V" & MaxRenglon + 5).AddIndent = False
Hoja.Range("A4:V" & MaxRenglon + 5).IndentLevel = 0
Hoja.Range("A4:V" & MaxRenglon + 5).ShrinkToFit = False
Hoja.Range("A4:V" & MaxRenglon + 5).ReadingOrder = xlContext
    
    
Hoja.Range("A4:V" & MaxRenglon + 5).Borders(xlDiagonalDown).LineStyle = xlNone
Hoja.Range("A4:V" & MaxRenglon + 5).Borders(xlDiagonalUp).LineStyle = xlNone

Hoja.Range("A4:V" & MaxRenglon + 5).Borders(xlEdgeLeft).LineStyle = xlContinuous
Hoja.Range("A4:V" & MaxRenglon + 5).Borders(xlEdgeLeft).ColorIndex = 0
Hoja.Range("A4:V" & MaxRenglon + 5).Borders(xlEdgeLeft).TintAndShade = 0
Hoja.Range("A4:V" & MaxRenglon + 5).Borders(xlEdgeLeft).Weight = xlThin

Hoja.Range("A4:V" & MaxRenglon + 5).Borders(xlEdgeBottom).LineStyle = xlContinuous
Hoja.Range("A4:V" & MaxRenglon + 5).Borders(xlEdgeBottom).ColorIndex = 0
Hoja.Range("A4:V" & MaxRenglon + 5).Borders(xlEdgeBottom).TintAndShade = 0
Hoja.Range("A4:V" & MaxRenglon + 5).Borders(xlEdgeBottom).Weight = xlThin

Hoja.Range("A4:V" & MaxRenglon + 5).Borders(xlEdgeRight).LineStyle = xlContinuous
Hoja.Range("A4:V" & MaxRenglon + 5).Borders(xlEdgeRight).ColorIndex = 0
Hoja.Range("A4:V" & MaxRenglon + 5).Borders(xlEdgeRight).TintAndShade = 0
Hoja.Range("A4:V" & MaxRenglon + 5).Borders(xlEdgeRight).Weight = xlThin

Hoja.Range("A4:V" & MaxRenglon + 5).Borders(xlEdgeTop).LineStyle = xlContinuous
Hoja.Range("A4:V" & MaxRenglon + 5).Borders(xlEdgeTop).ColorIndex = 0
Hoja.Range("A4:V" & MaxRenglon + 5).Borders(xlEdgeTop).TintAndShade = 0
Hoja.Range("A4:V" & MaxRenglon + 5).Borders(xlEdgeTop).Weight = xlThin

Hoja.Range("A4:V" & MaxRenglon + 5).Borders(xlInsideVertical).LineStyle = xlContinuous
Hoja.Range("A4:V" & MaxRenglon + 5).Borders(xlInsideVertical).ColorIndex = 0
Hoja.Range("A4:V" & MaxRenglon + 5).Borders(xlInsideVertical).TintAndShade = 0
Hoja.Range("A4:V" & MaxRenglon + 5).Borders(xlInsideVertical).Weight = xlThin

Hoja.Range("A4:V" & MaxRenglon + 5).Borders(xlInsideHorizontal).LineStyle = xlContinuous
Hoja.Range("A4:V" & MaxRenglon + 5).Borders(xlInsideHorizontal).ColorIndex = 0
Hoja.Range("A4:V" & MaxRenglon + 5).Borders(xlInsideHorizontal).TintAndShade = 0
Hoja.Range("A4:V" & MaxRenglon + 5).Borders(xlInsideHorizontal).Weight = xlThin



libro.ActiveSheet.PageSetup.LeftHeaderPicture.FileName = "C:\legacy\Gemtron.jpg"
libro.ActiveSheet.PageSetup.PrintTitleRows = "$1:$5"
libro.ActiveSheet.PageSetup.PrintTitleColumns = ""

libro.ActiveSheet.PageSetup.PrintArea = ""
libro.ActiveSheet.PageSetup.LeftHeader = "&G"
libro.ActiveSheet.PageSetup.CenterHeader = ""
libro.ActiveSheet.PageSetup.RightHeader = ""
libro.ActiveSheet.PageSetup.LeftFooter = "Imprimio: " & Nombre & vbNewLine & "Forma F-8.1.1-E (14 Sept 2020)" ' "FORMA #F-TEM-020-A(08/17/2009)"
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
'libro.ActiveSheet.PageSetup.Orientation = xlPortrait

libro.ActiveSheet.PageSetup.Orientation = xlLandscape

libro.ActiveSheet.PageSetup.Draft = False
libro.ActiveSheet.PageSetup.PaperSize = xlPaperLetter
libro.ActiveSheet.PageSetup.FirstPageNumber = xlAutomatic
libro.ActiveSheet.PageSetup.Order = xlDownThenOver
libro.ActiveSheet.PageSetup.BlackAndWhite = False
libro.ActiveSheet.PageSetup.Zoom = 57
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

'Agrega renglones

'MaxRenglon
For i = 12 To (MaxRenglon + 16)
    Band = False
    'If CDate(Hoja.Cells(i, 2)) = CDate("08:00:00 a.m.") Then    '"08:00:00 a. m."
    If Hoja.Cells(i, 2) = "6 a.m.-7 a.m." Then    '"6 a.m.-7 a.m."
        Band = True
            Hoja.Rows("" & i + 1 & ":" & i + 1 & "").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            Hoja.Cells("" & i + 1 & "", 6) = "=SUM(F" & (i - 8) + 1 & ":F" & (i - 1) + 1 & ")"
            Hoja.Cells("" & i + 1 & "", 11) = "=SUM(K" & (i - 8) + 1 & ":K" & (i - 1) + 1 & ")"
            Hoja.Cells("" & i + 1 & "", 16) = "=SUM(P" & (i - 8) + 1 & ":P" & (i - 1) + 1 & ")"
            Hoja.Cells("" & i + 1 & "", 21) = "=SUM(U" & (i - 8) + 1 & ":U" & (i - 1) + 1 & ")"
            Hoja.Rows("" & i + 1 & ":" & i + 1 & "").Interior.Pattern = xlSolid
            Hoja.Rows("" & i + 1 & ":" & i + 1 & "").Interior.PatternColorIndex = xlAutomatic
            Hoja.Rows("" & i + 1 & ":" & i + 1 & "").Interior.Color = 65535
            i = i + 1
    End If
    
    'If Hoja.Cells(i, 2) = "2 p.m.-3 p.m." Then  '2 p.m.-3 p.m.
    
    If Hoja.Cells(i, 2) = "2 p.m.-3 p.m." Then  '2 p.m.-3 p.m.
        Band = True
            Hoja.Rows("" & i + 1 & ":" & i + 1 & "").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            Hoja.Cells("" & i + 1 & "", 6) = "=SUM(F" & (i - 8) + 1 & ":F" & (i - 1) + 1 & ")"
            Hoja.Cells("" & i + 1 & "", 11) = "=SUM(K" & (i - 8) + 1 & ":K" & (i - 1) + 1 & ")"
            Hoja.Cells("" & i + 1 & "", 16) = "=SUM(P" & (i - 8) + 1 & ":P" & (i - 1) + 1 & ")"
            Hoja.Cells("" & i + 1 & "", 21) = "=SUM(U" & (i - 8) + 1 & ":U" & (i - 1) + 1 & ")"
            Hoja.Rows("" & i + 1 & ":" & i + 1 & "").Interior.Pattern = xlSolid
            Hoja.Rows("" & i + 1 & ":" & i + 1 & "").Interior.PatternColorIndex = xlAutomatic
            Hoja.Rows("" & i + 1 & ":" & i + 1 & "").Interior.Color = 65535
            i = i + 1
    End If
    If Hoja.Cells(i, 2) = "10 p.m.-11 p.m." Then  '10 p.m.-11 p.m.
        Band = True
            Hoja.Rows("" & i + 1 & ":" & i + 1 & "").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            Hoja.Cells("" & i + 1 & "", 6) = "=SUM(F" & (i - 7) & ":F" & (i - 1) + 1 & ")"
            Hoja.Cells("" & i + 1 & "", 11) = "=SUM(K" & (i - 7) & ":K" & (i - 1) + 1 & ")"
            Hoja.Cells("" & i + 1 & "", 16) = "=SUM(P" & (i - 7) & ":P" & (i - 1) + 1 & ")"
            Hoja.Cells("" & i + 1 & "", 21) = "=SUM(U" & (i - 7) & ":U" & (i - 1) + 1 & ")"
            Hoja.Rows("" & i + 1 & ":" & i + 1 & "").Interior.Pattern = xlSolid
            Hoja.Rows("" & i + 1 & ":" & i + 1 & "").Interior.PatternColorIndex = xlAutomatic
            Hoja.Rows("" & i + 1 & ":" & i + 1 & "").Interior.Color = 65535
            i = i + 1
    End If
Next i

Randomize

'Ruta = "C:\mis documentos\ProgFlat Fabricacion " & Format(Day(Date), "00") & Format(Month(Date), "00") & Format(Year(Date), "00") & " " & Int((1000 * Rnd) + 1) & ".xlsx"

Ruta = "C:\Legacy\ProgFlat Fabricacion " & Format(Day(Date), "00") & Format(Month(Date), "00") & Format(Year(Date), "00") & " " & Int((1000 * Rnd) + 1) & ".xlsx"

Hoja.SaveAs (Ruta)
libro.Close
Set Hoja = Nothing
Set libro = Nothing
vApp.Quit
Set vApp = Nothing

Dim CnnExcel As New Excel.Application
Dim LibroPDF As Excel.Workbook
'Abrimos el documento

Resp = Int((1000 * Rnd) + 1)

RutaPDF = "\\10.18.172.30\Departments\Materials\Read\Reporte Programa Produccion Flat\ProgFlat Fabricacion " & Format(Day(Date), "00") & Format(Month(Date), "00") & Format(Year(Date), "00") & " " & Resp & ".pdf"


RutaPDF = "Q:\Materials\Read\Reporte Programa Produccion FLAT\ProgFlat Fabricacion " & Format(Day(Date), "00") & Format(Month(Date), "00") & Format(Year(Date), "00") & " " & Resp & ".pdf"
'Q:\Materials\Read\Reporte Programa Produccion FLAT\


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
Cadena = "Programa de Fabricacion Flat " & Me.DTPF_Ini.Value & " a " & Me.DTPF_Fin.Value
RutaAdjunto = "\\10.18.172.30\Adjuntos\PFLAT Fabricacion " & Format(Day(Date), "00") & Format(Month(Date), "00") & Format(Year(Date), "00") & " " & Resp & ".pdf"


CNN.CmdReportes ("RepProdGral")
If CNN.rsCmdReportes.EOF <> True Then
        CNN.rsCmdReportes!rutareporte = RutaPDF
    CNN.rsCmdReportes.Update
End If
CNN.rsCmdReportes.Close



CnnExcel.Workbooks.Open (Ruta)   '''Prod Flat
'pause (2000)
CnnExcel.ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:= _
    RutaAdjunto, Quality:= _
    xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
    OpenAfterPublish:=False
CnnExcel.ActiveWorkbook.Close
'''CnnExcel.LibroPDF.Close
CnnExcel.Quit
Set CnnExcel = Nothing



'Para = "rfabela@sswtechnologies.com, jdiaz@sswtechnologies.com, agallegos@sswtechnologies.com"  ', abril.ramos@schott.com, "
'Para = "hireta@sswtechnologies.com; breyes@sswtechnologies.com; "
Para = "rfabela@sswtechnologies.com, jdiaz@sswtechnologies.com, agallegos@sswtechnologies.com, jolvera@sswtechnologies.com "

CC = "hireta@sswtechnologies.com; "

CNN.rsCmdCorreo.Open
CNN.rsCmdCorreo.AddNew
    CNN.rsCmdCorreo!De = "jdiaz@sswtechnologies.com" 'De
    CNN.rsCmdCorreo!Para = Para
    CNN.rsCmdCorreo!CC = CC
    CNN.rsCmdCorreo!CCO = "" '"hireta@sswtechnologies.com; breyes@sswtechnologies.com; regino.rivera@schott.com"
    CNN.rsCmdCorreo!Titulo = Cadena '& " Pruebas TI"
    CNN.rsCmdCorreo!Mensaje = CadenaCorreo
    CNN.rsCmdCorreo!aplicacion = "JC"
    CNN.rsCmdCorreo!Prioridad = 2
    CNN.rsCmdCorreo!Ruta = RutaAdjunto  'RutaAdjunto
CNN.rsCmdCorreo.Update
CNN.rsCmdCorreo.Close

 

''''  Templado   ********************************************************************************************  INICIA



Dim vAppTemplado As Excel.Application
Dim libroTemplado As Excel.Workbook
Dim HojaTemplado As Excel.Worksheet

Set vAppTemplado = CreateObject("Excel.Application")
Set libroTemplado = vAppTemplado.Workbooks.Add
Set HojaTemplado = libroTemplado.Worksheets(1)

'''Set HojaTemplado = libroTemplado.Worksheets(1)

Renglon = 4
Columna = 1
HojaTemplado.Select
Resp = "Prog Temp " & Format(Me.DTPF_Ini.Value, "d mmm yy") & " a " & Format(Me.DTPF_Fin.Value, "d mmm yy")


HojaTemplado.Name = Resp

'HojaTemplado.Sheets(Resp).Cells(1, 1) = "Gemtron de México S.A. de C.V."
HojaTemplado.Cells(1, 1) = "Gemtron de México S.A. de C.V."

'vAppTemplado.Visible = True

HojaTemplado.Range("A1:V1").Select
HojaTemplado.Range("A1:V1").HorizontalAlignment = xlCenter
HojaTemplado.Range("A1:V1").VerticalAlignment = xlBottom
HojaTemplado.Range("A1:V1").WrapText = False
HojaTemplado.Range("A1:V1").MergeCells = True

HojaTemplado.Range("A1:V1").Select
HojaTemplado.Range("A1:V1").Font.Name = "Times New Roman"
HojaTemplado.Range("A1:V1").Font.Size = 16
HojaTemplado.Range("A1:V1").Font.Bold = True


HojaTemplado.Cells(2, 1) = "Reporte Programa de Producción Templado"
HojaTemplado.Range("A2:V2").Select
HojaTemplado.Range("A2:V2").HorizontalAlignment = xlCenter
HojaTemplado.Range("A2:V2").VerticalAlignment = xlBottom
HojaTemplado.Range("A2:V2").WrapText = False
HojaTemplado.Range("A2:V2").MergeCells = True

HojaTemplado.Range("A2:V2").Select
HojaTemplado.Range("A2:V2").Font.Name = "Times New Roman"
HojaTemplado.Range("A2:V2").Font.Size = 12
HojaTemplado.Range("A2:V2").Font.Bold = True
HojaTemplado.Range("A2:V2").Font.Italic = True

'Encabezados
HojaTemplado.Cells(4, 1) = "Fecha"
HojaTemplado.Cells(5, 2) = "Hora"

HojaTemplado.Cells(4, 3) = "Linea 1"
HojaTemplado.Cells(5, 3) = "Tipo"
HojaTemplado.Cells(5, 4) = "No.  JC"
HojaTemplado.Cells(5, 5) = "No de Parte"
HojaTemplado.Cells(5, 6) = "Cantidad"
HojaTemplado.Cells(5, 7) = "No. OMP SAP"


HojaTemplado.Cells(4, 8) = "Linea 2"
HojaTemplado.Cells(5, 8) = "Tipo"
HojaTemplado.Cells(5, 9) = "No.  JC"
HojaTemplado.Cells(5, 10) = "No de Parte"
HojaTemplado.Cells(5, 11) = "Cantidad"
HojaTemplado.Cells(5, 12) = "No. OMP SAP"

HojaTemplado.Cells(4, 13) = "Linea 3"
HojaTemplado.Cells(5, 13) = "Tipo"
HojaTemplado.Cells(5, 14) = "No.  JC"
HojaTemplado.Cells(5, 15) = "No de Parte"
HojaTemplado.Cells(5, 16) = "Cantidad"
HojaTemplado.Cells(5, 17) = "No. OMP SAP"

HojaTemplado.Cells(4, 18) = "Linea 4"
HojaTemplado.Cells(5, 18) = "Tipo"
HojaTemplado.Cells(5, 19) = "No.  JC"
HojaTemplado.Cells(5, 20) = "No de Parte"
HojaTemplado.Cells(5, 21) = "Cantidad"
HojaTemplado.Cells(5, 22) = "No. OMP SAP"


HojaTemplado.Range("A4:A5").Select
HojaTemplado.Range("A4:A5").HorizontalAlignment = xlCenter
HojaTemplado.Range("A4:A5").VerticalAlignment = xlBottom
HojaTemplado.Range("A4:A5").MergeCells = True

HojaTemplado.Range("B4:B5").Select
HojaTemplado.Range("B4:B5").HorizontalAlignment = xlCenter
HojaTemplado.Range("B4:B5").VerticalAlignment = xlBottom
HojaTemplado.Range("B4:B5").MergeCells = True



HojaTemplado.Range("A5:k5").Select
HojaTemplado.Range("A5:k5").Font.Name = "Times New Roman"
HojaTemplado.Range("A5:k5").Font.Size = 12
HojaTemplado.Range("A5:k5").Font.Bold = True


'Linea 1
HojaTemplado.Range("C4:G4").Select
HojaTemplado.Range("C4:G4").HorizontalAlignment = xlCenter
HojaTemplado.Range("C4:G4").MergeCells = True

'Linea 2
HojaTemplado.Range("H4:L4").Select
HojaTemplado.Range("H4:L4").HorizontalAlignment = xlCenter
HojaTemplado.Range("H4:L4").MergeCells = True

'Linea 3
HojaTemplado.Range("M4:Q4").Select
HojaTemplado.Range("M4:Q4").HorizontalAlignment = xlCenter
HojaTemplado.Range("M4:Q4").MergeCells = True



'Linea 4
HojaTemplado.Range("R4:V4").Select
HojaTemplado.Range("R4:V4").HorizontalAlignment = xlCenter
HojaTemplado.Range("R4:V4").MergeCells = True




HojaTemplado.Range("A4:V4").Select
HojaTemplado.Range("A4:V4").Font.Name = "Times New Roman"
HojaTemplado.Range("A4:V4").Font.Size = 11
HojaTemplado.Range("A4:V4").Font.Bold = True


HojaTemplado.Range("A4:V5").Select
HojaTemplado.Range("A4:V5").Interior.Pattern = xlSolid
HojaTemplado.Range("A4:V5").Interior.PatternColorIndex = xlAutomatic
HojaTemplado.Range("A4:V5").Interior.PatternTintAndShade = 0
HojaTemplado.Range("A4:V5").Interior.ThemeColor = xlThemeColorLight2
HojaTemplado.Range("A4:V5").Interior.TintAndShade = 0.799981688894314


' para los Codigos de MP
Renglon = 6
'''Bordes
HojaTemplado.Range("A4:V5").Select
HojaTemplado.Range("A4:V5").VerticalAlignment = xlCenter
HojaTemplado.Range("A4:V5").HorizontalAlignment = xlCenter

HojaTemplado.Range("A4:V5").Select

HojaTemplado.Range("A4:V5").Borders(xlDiagonalDown).LineStyle = xlNone

HojaTemplado.Range("A4:V5").Borders(xlDiagonalDown).LineStyle = xlNone
HojaTemplado.Range("A4:V5").Borders(xlDiagonalUp).LineStyle = xlNone

HojaTemplado.Range("A4:V5").Borders(xlEdgeLeft).LineStyle = xlContinuous
HojaTemplado.Range("A4:V5").Borders(xlEdgeLeft).ColorIndex = 0
HojaTemplado.Range("A4:V5").Borders(xlEdgeLeft).TintAndShade = 0
HojaTemplado.Range("A4:V5").Borders(xlEdgeLeft).Weight = xlThin

HojaTemplado.Range("A4:V5").Borders(xlEdgeBottom).LineStyle = xlContinuous
HojaTemplado.Range("A4:V5").Borders(xlEdgeBottom).ColorIndex = 0
HojaTemplado.Range("A4:V5").Borders(xlEdgeBottom).TintAndShade = 0
HojaTemplado.Range("A4:V5").Borders(xlEdgeBottom).Weight = xlThin

HojaTemplado.Range("A4:V5").Borders(xlEdgeRight).LineStyle = xlContinuous
HojaTemplado.Range("A4:V5").Borders(xlEdgeRight).ColorIndex = 0
HojaTemplado.Range("A4:V5").Borders(xlEdgeRight).TintAndShade = 0
HojaTemplado.Range("A4:V5").Borders(xlEdgeRight).Weight = xlThin

HojaTemplado.Range("A4:V5").Borders(xlEdgeTop).LineStyle = xlContinuous
HojaTemplado.Range("A4:V5").Borders(xlEdgeTop).ColorIndex = 0
HojaTemplado.Range("A4:V5").Borders(xlEdgeTop).TintAndShade = 0
HojaTemplado.Range("A4:V5").Borders(xlEdgeTop).Weight = xlThin

HojaTemplado.Range("A4:V5").Borders(xlInsideVertical).LineStyle = xlContinuous
HojaTemplado.Range("A4:V5").Borders(xlInsideVertical).ColorIndex = 0
HojaTemplado.Range("A4:V5").Borders(xlInsideVertical).TintAndShade = 0
HojaTemplado.Range("A4:V5").Borders(xlInsideVertical).Weight = xlThin

HojaTemplado.Range("A4:V5").Borders(xlInsideHorizontal).LineStyle = xlContinuous
HojaTemplado.Range("A4:V5").Borders(xlInsideHorizontal).ColorIndex = 0
HojaTemplado.Range("A4:V5").Borders(xlInsideHorizontal).TintAndShade = 0
HojaTemplado.Range("A4:V5").Borders(xlInsideHorizontal).Weight = xlThin

    
HojaTemplado.Range("A5:V5").Select
HojaTemplado.Range("A5:V5").Font.Name = "Times New Roman"
HojaTemplado.Range("A5:V5").Font.Size = 11
HojaTemplado.Range("A5:V5").Font.Strikethrough = False
HojaTemplado.Range("A5:V5").Font.Superscript = False
HojaTemplado.Range("A5:V5").Font.Subscript = False
HojaTemplado.Range("A5:V5").Font.OutlineFont = False
HojaTemplado.Range("A5:V5").Font.Shadow = False
HojaTemplado.Range("A5:V5").Font.Underline = xlUnderlineStyleNone
HojaTemplado.Range("A5:V5").Font.ThemeColor = xlThemeColorLight1
HojaTemplado.Range("A5:V5").Font.TintAndShade = 0
HojaTemplado.Range("A5:V5").Font.ThemeFont = xlThemeFontNone
HojaTemplado.Range("A5:V5").Font.Bold = True


'Ahora las Fechas
Cantidad = (Me.DTPF_Fin.Value - Me.DTPF_Ini.Value) + 1
'Dim Inicial As Integer
'Dim Final As Integer
Fecha = Me.DTPF_Ini.Value

Inicial = 6
i = 1
j = 0

Do While i <= (Cantidad * 24)
    If j = 0 Then
        HojaTemplado.Cells(i + 5, 1) = Format(Fecha, "dd mmm yyyy")
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
    
    
    HojaTemplado.Cells(i + 5, 2) = Hora
'''    HojaTemplado.Cells(i + 5, 6) = Hora
'''    HojaTemplado.Cells(i + 5, 10) = Hora
    If (i Mod 24) = 0 Then
        HojaTemplado.Range("A" & (i - 18) & ":A" & (i + 5)).Select
        HojaTemplado.Range("A" & (i - 18) & ":A" & (i + 5)).HorizontalAlignment = xlCenter
        HojaTemplado.Range("A" & (i - 18) & ":A" & (i + 5)).VerticalAlignment = xlCenter
        HojaTemplado.Range("A" & (i - 18) & ":A" & (i + 5)).WrapText = False
        HojaTemplado.Range("A" & (i - 18) & ":A" & (i + 5)).Orientation = 90
        HojaTemplado.Range("A" & (i - 18) & ":A" & (i + 5)).AddIndent = False
        HojaTemplado.Range("A" & (i - 18) & ":A" & (i + 5)).IndentLevel = 0
        HojaTemplado.Range("A" & (i - 18) & ":A" & (i + 5)).ShrinkToFit = False
        HojaTemplado.Range("A" & (i - 18) & ":A" & (i + 5)).ReadingOrder = xlContext
        HojaTemplado.Range("A" & (i - 18) & ":A" & (i + 5)).MergeCells = True
        
        Fecha = Fecha + 1
        j = 0
    Else
        j = j + 1
    End If
    i = i + 1
Loop

MaxRenglon = i - 1

F_Ini = Me.DTPF_Ini.Value

'para toda las lineas
Renglon = 6
'k = 111
'Do While k <= 331
    
    '.rsCmdRepLineas.
    
    CNN.CmdRepGral_Templado (Me.DTPF_Ini.Value), (Me.DTPF_Fin.Value) ', (k)
    If CNN.rsCmdRepGral_Templado.EOF <> True Then
        CNN.rsCmdRepGral_Templado.MoveFirst
        m = 1
        Do While CNN.rsCmdRepGral_Templado.EOF <> True
            Fecha = CNN.rsCmdRepGral_Templado!Fecha
            IdHora = CNN.rsCmdRepGral_Templado!IdHora
            OF = CNN.rsCmdRepGral_Templado!OF
            CNN.CmdBuscaEsp2 (OF)
            If CNN.rsCmdBuscaEsp2.EOF <> True Then
                NoParte = CNN.rsCmdBuscaEsp2!nodeparte
            Else
                CNN.CmdVidrio_MPS (OF)
                If CNN.rsCmdVidrio_MPS.EOF <> True Then
                    NoParte = CNN.rsCmdVidrio_MPS!Codigo
                Else
                    MsgBox "El PT o V MPS NO existe", vbInformation, MSG
                    NoParte = "Vacio"
                End If
                CNN.rsCmdVidrio_MPS.Close
            End If
            CNN.rsCmdBuscaEsp2.Close
            Renglon = (6 + (Fecha - Me.DTPF_Ini.Value) * 24) + IdHora
            
            
            'Busca No. de Linea
            If InStr(1, CNN.rsCmdRepGral_Templado!Descripcion, "1") Then
                k = 1
            End If
            If InStr(1, CNN.rsCmdRepGral_Templado!Descripcion, "2") Then
                k = 2
            End If
            If InStr(1, CNN.rsCmdRepGral_Templado!Descripcion, "3") Then
                k = 3
            End If
            
             If InStr(1, CNN.rsCmdRepGral_Templado!Descripcion, "4") Then
                k = 4
            End If
            
            
            
            Select Case k
                Case 1 '111
                    HojaTemplado.Cells(Renglon, 3) = CNN.rsCmdRepGral_Templado!Descripcion
                    HojaTemplado.Cells(Renglon, 5) = NoParte
                    If Mid(OF, 1, 1) <> "X" Then
                        HojaTemplado.Cells(Renglon, 4) = CNN.rsCmdRepGral_Templado!jc
                        HojaTemplado.Cells(Renglon, 6) = CNN.rsCmdRepGral_Templado!pzspt
                        
                    End If
                    HojaTemplado.Cells(Renglon, 7) = CNN.rsCmdRepGral_Templado!OrdenProd_SAP
                    
                Case 2 '221
                    HojaTemplado.Cells(Renglon, 8) = CNN.rsCmdRepGral_Templado!Descripcion
                    HojaTemplado.Cells(Renglon, 10) = NoParte
                    If Mid(OF, 1, 1) <> "X" Then
                        HojaTemplado.Cells(Renglon, 9) = CNN.rsCmdRepGral_Templado!jc
                        HojaTemplado.Cells(Renglon, 11) = CNN.rsCmdRepGral_Templado!pzspt
                    End If
                     HojaTemplado.Cells(Renglon, 12) = CNN.rsCmdRepGral_Templado!OrdenProd_SAP
                
                Case 3 '331
                    HojaTemplado.Cells(Renglon, 13) = CNN.rsCmdRepGral_Templado!Descripcion
                    HojaTemplado.Cells(Renglon, 15) = NoParte
                    If Mid(OF, 1, 1) <> "X" Then
                        HojaTemplado.Cells(Renglon, 14) = CNN.rsCmdRepGral_Templado!jc
                        HojaTemplado.Cells(Renglon, 16) = CNN.rsCmdRepGral_Templado!pzspt
                    End If
                     HojaTemplado.Cells(Renglon, 17) = CNN.rsCmdRepGral_Templado!OrdenProd_SAP
                 
                 Case 4 '441
                    HojaTemplado.Cells(Renglon, 18) = CNN.rsCmdRepLineasGral!Descripcion
                    HojaTemplado.Cells(Renglon, 20) = NoParte
                    If Mid(OF, 1, 1) <> "X" Then
                        HojaTemplado.Cells(Renglon, 19) = CNN.rsCmdRepLineasGral!jc
                        HojaTemplado.Cells(Renglon, 21) = CNN.rsCmdRepLineasGral!pzspt
                    End If
                     HojaTemplado.Cells(Renglon, 22) = CNN.rsCmdRepGral_Templado!OrdenProd_SAP
                    
            End Select
            CNN.rsCmdRepGral_Templado.MoveNext
        Loop
    End If
    CNN.rsCmdRepGral_Templado.Close

'    k = k + 110
'Loop



HojaTemplado.Range("B6:V" & MaxRenglon).Select
HojaTemplado.Range("B6:V" & MaxRenglon).HorizontalAlignment = xlCenter
HojaTemplado.Range("B6:V" & MaxRenglon).VerticalAlignment = xlCenter

HojaTemplado.Rows("6:" & MaxRenglon + 5).RowHeight = 20.25
HojaTemplado.Columns("A:V").EntireColumn.AutoFit
HojaTemplado.Rows("3:3").RowHeight = 27.25

HojaTemplado.Range("A4:V" & MaxRenglon + 5).Select
HojaTemplado.Range("A4:V" & MaxRenglon + 5).VerticalAlignment = xlCenter
HojaTemplado.Range("A4:V" & MaxRenglon + 5).HorizontalAlignment = xlCenter
HojaTemplado.Range("A4:V" & MaxRenglon + 5).WrapText = False
HojaTemplado.Range("A4:V" & MaxRenglon + 5).AddIndent = False
HojaTemplado.Range("A4:V" & MaxRenglon + 5).IndentLevel = 0
HojaTemplado.Range("A4:V" & MaxRenglon + 5).ShrinkToFit = False
HojaTemplado.Range("A4:V" & MaxRenglon + 5).ReadingOrder = xlContext
    
    
HojaTemplado.Range("A4:V" & MaxRenglon + 5).Borders(xlDiagonalDown).LineStyle = xlNone
HojaTemplado.Range("A4:V" & MaxRenglon + 5).Borders(xlDiagonalUp).LineStyle = xlNone

HojaTemplado.Range("A4:V" & MaxRenglon + 5).Borders(xlEdgeLeft).LineStyle = xlContinuous
HojaTemplado.Range("A4:V" & MaxRenglon + 5).Borders(xlEdgeLeft).ColorIndex = 0
HojaTemplado.Range("A4:V" & MaxRenglon + 5).Borders(xlEdgeLeft).TintAndShade = 0
HojaTemplado.Range("A4:V" & MaxRenglon + 5).Borders(xlEdgeLeft).Weight = xlThin

HojaTemplado.Range("A4:V" & MaxRenglon + 5).Borders(xlEdgeBottom).LineStyle = xlContinuous
HojaTemplado.Range("A4:V" & MaxRenglon + 5).Borders(xlEdgeBottom).ColorIndex = 0
HojaTemplado.Range("A4:V" & MaxRenglon + 5).Borders(xlEdgeBottom).TintAndShade = 0
HojaTemplado.Range("A4:V" & MaxRenglon + 5).Borders(xlEdgeBottom).Weight = xlThin

HojaTemplado.Range("A4:V" & MaxRenglon + 5).Borders(xlEdgeRight).LineStyle = xlContinuous
HojaTemplado.Range("A4:V" & MaxRenglon + 5).Borders(xlEdgeRight).ColorIndex = 0
HojaTemplado.Range("A4:V" & MaxRenglon + 5).Borders(xlEdgeRight).TintAndShade = 0
HojaTemplado.Range("A4:V" & MaxRenglon + 5).Borders(xlEdgeRight).Weight = xlThin

HojaTemplado.Range("A4:V" & MaxRenglon + 5).Borders(xlEdgeTop).LineStyle = xlContinuous
HojaTemplado.Range("A4:V" & MaxRenglon + 5).Borders(xlEdgeTop).ColorIndex = 0
HojaTemplado.Range("A4:V" & MaxRenglon + 5).Borders(xlEdgeTop).TintAndShade = 0
HojaTemplado.Range("A4:V" & MaxRenglon + 5).Borders(xlEdgeTop).Weight = xlThin

HojaTemplado.Range("A4:V" & MaxRenglon + 5).Borders(xlInsideVertical).LineStyle = xlContinuous
HojaTemplado.Range("A4:V" & MaxRenglon + 5).Borders(xlInsideVertical).ColorIndex = 0
HojaTemplado.Range("A4:V" & MaxRenglon + 5).Borders(xlInsideVertical).TintAndShade = 0
HojaTemplado.Range("A4:V" & MaxRenglon + 5).Borders(xlInsideVertical).Weight = xlThin

HojaTemplado.Range("A4:V" & MaxRenglon + 5).Borders(xlInsideHorizontal).LineStyle = xlContinuous
HojaTemplado.Range("A4:V" & MaxRenglon + 5).Borders(xlInsideHorizontal).ColorIndex = 0
HojaTemplado.Range("A4:V" & MaxRenglon + 5).Borders(xlInsideHorizontal).TintAndShade = 0
HojaTemplado.Range("A4:V" & MaxRenglon + 5).Borders(xlInsideHorizontal).Weight = xlThin



libroTemplado.ActiveSheet.PageSetup.LeftHeaderPicture.FileName = "C:\legacy\Gemtron.jpg"
libroTemplado.ActiveSheet.PageSetup.PrintTitleRows = "$1:$5"
libroTemplado.ActiveSheet.PageSetup.PrintTitleColumns = ""

libroTemplado.ActiveSheet.PageSetup.PrintArea = ""
libroTemplado.ActiveSheet.PageSetup.LeftHeader = "&G"
libroTemplado.ActiveSheet.PageSetup.CenterHeader = ""
libroTemplado.ActiveSheet.PageSetup.RightHeader = ""
libroTemplado.ActiveSheet.PageSetup.LeftFooter = "Imprimio: " & Nombre & vbNewLine & "FORMA #F-TEM-020-A(08/17/2009)"
libroTemplado.ActiveSheet.PageSetup.CenterFooter = "&P"
libroTemplado.ActiveSheet.PageSetup.RightFooter = "&T" & Chr(10) & "&D"
libroTemplado.ActiveSheet.PageSetup.LeftMargin = libroTemplado.Application.InchesToPoints(0.236220472440945)
libroTemplado.ActiveSheet.PageSetup.RightMargin = libroTemplado.Application.InchesToPoints(0.275590551181102)
libroTemplado.ActiveSheet.PageSetup.TopMargin = libroTemplado.Application.InchesToPoints(0.275590551181102)
libroTemplado.ActiveSheet.PageSetup.BottomMargin = libroTemplado.Application.InchesToPoints(0.98551181102362)
libroTemplado.ActiveSheet.PageSetup.HeaderMargin = libroTemplado.Application.InchesToPoints(0.275590551181102)
libroTemplado.ActiveSheet.PageSetup.FooterMargin = libroTemplado.Application.InchesToPoints(0.31496062992126)
libroTemplado.ActiveSheet.PageSetup.PrintHeadings = False
libroTemplado.ActiveSheet.PageSetup.PrintGridlines = False
libroTemplado.ActiveSheet.PageSetup.PrintComments = xlPrintNoComments
libroTemplado.ActiveSheet.PageSetup.PrintQuality = 600
libroTemplado.ActiveSheet.PageSetup.CenterHorizontally = False
libroTemplado.ActiveSheet.PageSetup.CenterVertically = False
libroTemplado.ActiveSheet.PageSetup.Orientation = xlPortrait
libroTemplado.ActiveSheet.PageSetup.Draft = False
libroTemplado.ActiveSheet.PageSetup.PaperSize = xlPaperLetter
libroTemplado.ActiveSheet.PageSetup.FirstPageNumber = xlAutomatic
libroTemplado.ActiveSheet.PageSetup.Order = xlDownThenOver
libroTemplado.ActiveSheet.PageSetup.BlackAndWhite = False
libroTemplado.ActiveSheet.PageSetup.Zoom = 65
libroTemplado.ActiveSheet.PageSetup.PrintErrors = xlPrintErrorsDisplayed
libroTemplado.ActiveSheet.PageSetup.OddAndEvenPagesHeaderFooter = False
libroTemplado.ActiveSheet.PageSetup.DifferentFirstPageHeaderFooter = False
libroTemplado.ActiveSheet.PageSetup.ScaleWithDocHeaderFooter = True
libroTemplado.ActiveSheet.PageSetup.AlignMarginsHeaderFooter = True
libroTemplado.ActiveSheet.PageSetup.EvenPage.LeftHeader.Text = ""
libroTemplado.ActiveSheet.PageSetup.EvenPage.CenterHeader.Text = ""
libroTemplado.ActiveSheet.PageSetup.EvenPage.RightHeader.Text = ""
libroTemplado.ActiveSheet.PageSetup.EvenPage.LeftFooter.Text = ""
libroTemplado.ActiveSheet.PageSetup.EvenPage.CenterFooter.Text = ""
libroTemplado.ActiveSheet.PageSetup.EvenPage.RightFooter.Text = ""
libroTemplado.ActiveSheet.PageSetup.FirstPage.LeftHeader.Text = ""
libroTemplado.ActiveSheet.PageSetup.FirstPage.CenterHeader.Text = ""
libroTemplado.ActiveSheet.PageSetup.FirstPage.RightHeader.Text = ""
libroTemplado.ActiveSheet.PageSetup.FirstPage.LeftFooter.Text = ""
libroTemplado.ActiveSheet.PageSetup.FirstPage.CenterFooter.Text = ""
libroTemplado.ActiveSheet.PageSetup.FirstPage.RightFooter.Text = ""

HojaTemplado.Range("K6").Select


''''  Templado   ********************************************************************************************  Termina

'l ' 'ibro.Select (NombreProduccion)

'libroTemplado.Sheets(NombreProduccion).Activate
'l ibro.Worksheets(NombreProduccion).Activate
'HojaTemplado.Cells(1, 15) = "X"

'HojaTemplado.Select


Randomize

Ruta = "C:\mis documentos\ProgFlatTemplado" & Format(Day(Date), "00") & Format(Month(Date), "00") & Format(Year(Date), "00") & " " & Int((1000 * Rnd) + 1) & ".xlsx"

Ruta = "C:\Legacy\ProgFlatTemplado" & Format(Day(Date), "00") & Format(Month(Date), "00") & Format(Year(Date), "00") & " " & Int((1000 * Rnd) + 1) & ".xlsx"


HojaTemplado.SaveAs (Ruta)
libroTemplado.Close
Set HojaTemplado = Nothing
Set libroTemplado = Nothing
vAppTemplado.Quit
Set vAppTemplado = Nothing

Dim CnnExcel2 As New Excel.Application
Dim libroTempladoPDF2 As Excel.Workbook
'Abrimos el documento


Resp = Int((1000 * Rnd) + 1)
RutaPDF = "Q:\Materials\Read\Reporte Programa Produccion FLAT\ProgTemplado " & Format(Day(Date), "00") & Format(Month(Date), "00") & Format(Year(Date), "00") & " " & Resp & ".pdf"
RutaPDF = "Q:\Materials\Read\Reporte Programa Produccion Flat\ProgFlatTemplado" & Format(Day(Date), "00") & Format(Month(Date), "00") & Format(Year(Date), "00") & " " & Resp & ".pdf"
'''RutaAdjunto = "Q:\Materials\Read\Reporte Programa Produccion Flat\PFLAT " & Format(Day(Date), "00") & Format(Month(Date), "00") & Format(Year(Date), "00") & " " & Resp & ".pdf"

'\\10.18.164.207\e$\Adjuntos
'RutaPDF = "\\10.18.164.207\e$\Adjuntos\ProgAcumulado " & Format(Day(Date), "00") & Format(Month(Date), "00") & Format(Year(Date), "00") & " " & Resp & ".pdf"
''''RutaPDF = "C:\mis documentos\PFLAT " & Format(Day(Date), "00") & Format(Month(Date), "00") & Format(Year(Date), "00") & " " & Int((1000 * Rnd) + 1) & ".pdf"

'Guarda como PDF
CnnExcel.Workbooks.Open (Ruta)
CnnExcel.ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:= _
    RutaPDF, Quality:= _
    xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
    OpenAfterPublish:=False
CnnExcel.ActiveWorkbook.Close
'''CnnExcel.libroTempladoPDF.Close
CnnExcel.Quit
Set CnnExcel = Nothing
MsgBox "Tarea Terminada" & vbNewLine & "Su archivo esta en la sig ruta: " & RutaPDF & ".xlsx", vbInformation, MSG

Cadena = "Programa de Producción Flat Templado" & Me.DTPF_Ini.Value & " a " & Me.DTPF_Fin.Value

'''CadenaCorreo = "El Programa de Producción de Flat se localiza en la siguiente ruta:<br><br>"
''''CadenaCorreo = CadenaCorreo & "<a href='" & RutaPDF & "'>" & RutaPDF & "</a> <br><br>"
'''CadenaCorreo = CadenaCorreo & "Gracias"

'RutaAdjunto = "\\10.18.172.30\e$\Data\Departments\Materials\Read\Reporte Programa Produccion Flat\PFLAT " & Format(Day(Date), "00") & Format(Month(Date), "00") & Format(Year(Date), "00") & " " & Resp & ".pdf"


'''RutaAdjunto = "Q:\Materials\Read\Reporte Programa Produccion Flat\PFLAT " & Format(Day(Date), "00") & Format(Month(Date), "00") & Format(Year(Date), "00") & " " & Resp & ".pdf"

'RutaAdjunto = "\\10.18.172.30\Departments\Materials\Read\Reporte Programa Produccion Flat\PFLAT " & Format(Day(Date), "00") & Format(Month(Date), "00") & Format(Year(Date), "00") & " " & Resp & ".pdf"

'Publica en Q
CNN.CmdReportes ("RepProdGralTemp")
If CNN.rsCmdReportes.EOF <> True Then
        CNN.rsCmdReportes!rutareporte = RutaPDF
    CNN.rsCmdReportes.Update
End If
CNN.rsCmdReportes.Close


RutaAdjunto = "\\10.18.172.30\Adjuntos\ProgFlatTemplado" & Format(Day(Date), "00") & Format(Month(Date), "00") & Format(Year(Date), "00") & " " & Resp & ".pdf"
CnnExcel.Workbooks.Open (Ruta)
CnnExcel.ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:= _
    RutaAdjunto, Quality:= _
    xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
    OpenAfterPublish:=False
CnnExcel.ActiveWorkbook.Close
'''CnnExcel.libroTempladoPDF.Close
CnnExcel.Quit
Set CnnExcel = Nothing

'\\10.18.172.30\Adjuntos

Para = "rfabela@sswtechnologies.com, jdiaz@sswtechnologies.com, agallegos@sswtechnologies.com, jolvera@sswtechnologies.com "

''Para = "hireta@sswtechnologies.com; breyes@sswtechnologies.com; "


CNN.rsCmdCorreo.Open
CNN.rsCmdCorreo.AddNew
    CNN.rsCmdCorreo!De = "jdiaz@sswtechnologies.com" 'De
    CNN.rsCmdCorreo!Para = Para
    CNN.rsCmdCorreo!CC = ""
    CNN.rsCmdCorreo!CCO = "hireta@sswtechnologies.com "
    CNN.rsCmdCorreo!Titulo = Cadena '& " Pruebas TI"
    CNN.rsCmdCorreo!Mensaje = CadenaCorreo
    CNN.rsCmdCorreo!aplicacion = "JC"
    CNN.rsCmdCorreo!Prioridad = 2
    CNN.rsCmdCorreo!Ruta = RutaAdjunto ' RutaPDF  'RutaAdjunto
CNN.rsCmdCorreo.Update
CNN.rsCmdCorreo.Close




Unload Me
End Sub



'''Set vApp = Nothing
'''ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:= _
'''    "C:\Documents and Settings\heliosireta\Mis documentos\Cosolidado ProgPrd " & Format(Day(Date), "00") & Format(Month(Date), "00") & Format(Year(Date), "00") & ".pdf", Quality:= _
'''    xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
'''    OpenAfterPublish:=False
    
'''Randomize
'''ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:= _
'''    "\\10.18.164.226\Departments\Materials\Read\Reporte Programa Produccion Flat\PFLAT " & Format(Day(Date), "00") & Format(Month(Date), "00") & Format(Year(Date), "00") & " " & Int((1000 * Rnd) + 1) & ".pdf", Quality:= _
'''    xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
'''    OpenAfterPublish:=True

''vApp.Visible = True
'Libro.Close
'Set Hoja = Nothing
'Set Libro = Nothing
'Set vApp = Nothing
'vApp.Quit
'''Set vApp = Nothing

'''ChDir "C:\Mis documentos\Reps Programa Flat"
'''    ActiveWorkbook.SaveAs FileName:= _
'''        "C:\Mis documentos\Reps Programa Flat\RFlat 200410.xlsx" _
'''        , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False

'''vApp.Quit
'''Reset
'''Dim vApp As Excel.Application
'''Dim Libro As Excel.Workbook
'''Dim Hoja As Excel.Worksheet
'''Set vApp = CreateObject("Excel.Application")
'''Set Libro = vApp.Workbooks.Add
'''Set Hoja = Libro.Worksheets(1)






'''  ******************************************

'Option Explicit
'
'Private Sub CmdSalir_Click()
'Unload Me
'End Sub
'
'Private Sub Form_Load()
'Me.DTPF_Ini.Value = Date
'Me.DTPF_Fin.Value = Date + 2
'End Sub
'Private Sub CmdVerReporte_Click()
''''If Me.DTPF_Ini.Value < Date Then
''''    MsgBox "Verifique las fechas, no puede sacar reportes para fechas pasadas.", vbExclamation, MSG
''''    Exit Sub
''''End If
'If Me.DTPF_Ini.Value > Me.DTPF_Fin.Value Then
'    MsgBox "Verifique las fechas, la Inicial no puede ser mayor que la Final.", vbExclamation, MSG
'    Exit Sub
'End If
'
'''''LLena hoja de excel
' 'Dim F As ADODB.Field
''Dim hoja  As Variant
'
''''Dim xlApp As Excel.Application
''''Dim xlBook As Excel.Workbook
''''Dim xlSheet As Excel.Worksheet
''''Set xlApp = CreateObject("Excel.Application")
''''Set xlBook = xlApp.Workbooks.Add
''''Set xlSheet = xlBook.Worksheets("Sheet1")
''''xlSheet.Range(Cells(1, 1), Cells(10, 2)).Value = "Hello"
''''xlBook.Saved = True
''''Set xlSheet = Nothing
''''Set xlBook = Nothing
''''xlApp.Quit
''''Set xlApp = Nothing
'
'''''Dim hoja  As Excel.Application
''''Set hoja = New Excel.Application
'
'Dim vApp As Excel.Application
'Dim libro As Excel.Workbook
'Dim Hoja As Excel.Worksheet
'
'Set vApp = CreateObject("Excel.Application")
'Set libro = vApp.Workbooks.Add
'Set Hoja = libro.Worksheets(1)
'
'
'
'
'
'Renglon = 4
'Columna = 1
''hoja.Sheets(1).Select
'Hoja.Select
'Resp = "Programa " & Format(Me.DTPF_Ini.Value, "d mmm yy") & " a " & Format(Me.DTPF_Fin.Value, "d mmm yy")
''hoja.Sheets(1).Name = Resp
'
'Hoja.Name = Resp
'
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
'
''Encabezados
'Hoja.Cells(4, 1) = "Fecha"
'Hoja.Cells(4, 3) = "Linea 1"
'Hoja.Cells(5, 2) = "Hora"
'Hoja.Cells(5, 3) = "No.  JC"
'Hoja.Cells(5, 4) = "No de Parte"
'Hoja.Cells(5, 5) = "Cantidad"
'Hoja.Cells(4, 6) = "Linea 2"
''''Hoja.Cells(5, 6) = "Hora"
'Hoja.Cells(5, 6) = "No.  JC"
'Hoja.Cells(5, 7) = "No de Parte"
'Hoja.Cells(5, 8) = "Cantidad"
'Hoja.Cells(4, 9) = "Linea 3"
''''Hoja.Cells(5, 10) = "Hora"
'Hoja.Cells(5, 9) = "No.  JC"
'Hoja.Cells(5, 10) = "No de Parte"
'Hoja.Cells(5, 11) = "Cantidad"
'
'Hoja.Range("A4:A5").Select
'Hoja.Range("A4:A5").HorizontalAlignment = xlCenter
'Hoja.Range("A4:A5").VerticalAlignment = xlBottom
'Hoja.Range("A4:A5").MergeCells = True
'
'Hoja.Range("B4:B5").Select
'Hoja.Range("B4:B5").HorizontalAlignment = xlCenter
'Hoja.Range("B4:B5").VerticalAlignment = xlBottom
'Hoja.Range("B4:B5").MergeCells = True
'
'
'
'Hoja.Range("A5:k5").Select
'Hoja.Range("A5:k5").Font.Name = "Times New Roman"
'Hoja.Range("A5:k5").Font.Size = 12
'Hoja.Range("A5:k5").Font.Bold = True
'
''Linea 1
'Hoja.Range("C4:E4").Select
'Hoja.Range("C4:E4").HorizontalAlignment = xlCenter
'Hoja.Range("C4:E4").MergeCells = True
'
''Linea 2
'Hoja.Range("F4:H4").Select
'Hoja.Range("F4:H4").HorizontalAlignment = xlCenter
'Hoja.Range("F4:H4").MergeCells = True
'
''Linea 3
'Hoja.Range("I4:K4").Select
'Hoja.Range("I4:K4").HorizontalAlignment = xlCenter
'Hoja.Range("I4:K4").MergeCells = True
'
'
'Hoja.Range("A4:K4").Select
'Hoja.Range("A4:K4").Font.Name = "Times New Roman"
'Hoja.Range("A4:K4").Font.Size = 11
'Hoja.Range("A4:K4").Font.Bold = True
'
'
'Hoja.Range("A4:K5").Select
'Hoja.Range("A4:K5").Interior.Pattern = xlSolid
'Hoja.Range("A4:K5").Interior.PatternColorIndex = xlAutomatic
'Hoja.Range("A4:K5").Interior.PatternTintAndShade = 0
'Hoja.Range("A4:K5").Interior.ThemeColor = xlThemeColorLight2
'Hoja.Range("A4:K5").Interior.TintAndShade = 0.799981688894314
'
'
'' para los Codigos de MP
'Renglon = 6
''''Bordes
'Hoja.Range("A4:K5").Select
'Hoja.Range("A4:K5").VerticalAlignment = xlCenter
'Hoja.Range("A4:K5").HorizontalAlignment = xlCenter
'
'Hoja.Range("A4:K5").Select
'
'Hoja.Range("A4:K5").Borders(xlDiagonalDown).LineStyle = xlNone
'
'Hoja.Range("A4:K5").Borders(xlDiagonalDown).LineStyle = xlNone
'Hoja.Range("A4:K5").Borders(xlDiagonalUp).LineStyle = xlNone
'
'Hoja.Range("A4:K5").Borders(xlEdgeLeft).LineStyle = xlContinuous
'Hoja.Range("A4:K5").Borders(xlEdgeLeft).ColorIndex = 0
'Hoja.Range("A4:K5").Borders(xlEdgeLeft).TintAndShade = 0
'Hoja.Range("A4:K5").Borders(xlEdgeLeft).Weight = xlThin
'
'Hoja.Range("A4:K5").Borders(xlEdgeBottom).LineStyle = xlContinuous
'Hoja.Range("A4:K5").Borders(xlEdgeBottom).ColorIndex = 0
'Hoja.Range("A4:K5").Borders(xlEdgeBottom).TintAndShade = 0
'Hoja.Range("A4:K5").Borders(xlEdgeBottom).Weight = xlThin
'
'Hoja.Range("A4:K5").Borders(xlEdgeRight).LineStyle = xlContinuous
'Hoja.Range("A4:K5").Borders(xlEdgeRight).ColorIndex = 0
'Hoja.Range("A4:K5").Borders(xlEdgeRight).TintAndShade = 0
'Hoja.Range("A4:K5").Borders(xlEdgeRight).Weight = xlThin
'
'Hoja.Range("A4:K5").Borders(xlEdgeTop).LineStyle = xlContinuous
'Hoja.Range("A4:K5").Borders(xlEdgeTop).ColorIndex = 0
'Hoja.Range("A4:K5").Borders(xlEdgeTop).TintAndShade = 0
'Hoja.Range("A4:K5").Borders(xlEdgeTop).Weight = xlThin
'
'Hoja.Range("A4:K5").Borders(xlInsideVertical).LineStyle = xlContinuous
'Hoja.Range("A4:K5").Borders(xlInsideVertical).ColorIndex = 0
'Hoja.Range("A4:K5").Borders(xlInsideVertical).TintAndShade = 0
'Hoja.Range("A4:K5").Borders(xlInsideVertical).Weight = xlThin
'
'Hoja.Range("A4:K5").Borders(xlInsideHorizontal).LineStyle = xlContinuous
'Hoja.Range("A4:K5").Borders(xlInsideHorizontal).ColorIndex = 0
'Hoja.Range("A4:K5").Borders(xlInsideHorizontal).TintAndShade = 0
'Hoja.Range("A4:K5").Borders(xlInsideHorizontal).Weight = xlThin
'
'
'Hoja.Range("A5:K5").Select
'Hoja.Range("A5:K5").Font.Name = "Times New Roman"
'Hoja.Range("A5:K5").Font.Size = 11
'Hoja.Range("A5:K5").Font.Strikethrough = False
'Hoja.Range("A5:K5").Font.Superscript = False
'Hoja.Range("A5:K5").Font.Subscript = False
'Hoja.Range("A5:K5").Font.OutlineFont = False
'Hoja.Range("A5:K5").Font.Shadow = False
'Hoja.Range("A5:K5").Font.Underline = xlUnderlineStyleNone
'Hoja.Range("A5:K5").Font.ThemeColor = xlThemeColorLight1
'Hoja.Range("A5:K5").Font.TintAndShade = 0
'Hoja.Range("A5:K5").Font.ThemeFont = xlThemeFontNone
'Hoja.Range("A5:K5").Font.Bold = True
'
'
''Ahora las Fechas
'Cantidad = (Me.DTPF_Fin.Value - Me.DTPF_Ini.Value) + 1
'Dim Inicial As Integer
'Dim Final As Integer
'Fecha = Me.DTPF_Ini.Value
'
'Inicial = 6
'i = 1
'j = 0
'Do While i <= (Cantidad * 24)
'    If j = 0 Then
'        Hoja.Cells(i + 5, 1) = Format(Fecha, "dd mmm yyyy")
'        Hora = CDate("12:00:00 a.m.")
'    Else
'        If j > 12 Then
'            Hora = CDate(j - 12 & ":00:00 p.m.")
'        Else
'            If j = 12 Then
'                Hora = CDate("12:00:00 p.m.")
'            Else
'                Hora = CDate(j & ":00:00 a.m.")
'            End If
'        End If
'    End If
'    Hoja.Cells(i + 5, 2) = Hora
''''    Hoja.Cells(i + 5, 6) = Hora
''''    Hoja.Cells(i + 5, 10) = Hora
'    If (i Mod 24) = 0 Then
'        Hoja.Range("A" & (i - 18) & ":A" & (i + 5)).Select
'        Hoja.Range("A" & (i - 18) & ":A" & (i + 5)).HorizontalAlignment = xlCenter
'        Hoja.Range("A" & (i - 18) & ":A" & (i + 5)).VerticalAlignment = xlCenter
'        Hoja.Range("A" & (i - 18) & ":A" & (i + 5)).WrapText = False
'        Hoja.Range("A" & (i - 18) & ":A" & (i + 5)).Orientation = 90
'        Hoja.Range("A" & (i - 18) & ":A" & (i + 5)).AddIndent = False
'        Hoja.Range("A" & (i - 18) & ":A" & (i + 5)).IndentLevel = 0
'        Hoja.Range("A" & (i - 18) & ":A" & (i + 5)).ShrinkToFit = False
'        Hoja.Range("A" & (i - 18) & ":A" & (i + 5)).ReadingOrder = xlContext
'        Hoja.Range("A" & (i - 18) & ":A" & (i + 5)).MergeCells = True
'
'        Fecha = Fecha + 1
'        j = 0
'    Else
'        j = j + 1
'    End If
'    i = i + 1
'Loop
'
'MaxRenglon = i - 1
'
'F_Ini = Me.DTPF_Ini.Value
'
''para toda las lineas
'Renglon = 6
''k = 111
''Do While k <= 331
'
'    '.rsCmdRepLineas.
'
'    CNN.CmdRepLineasGral (Me.DTPF_Ini.Value), (Me.DTPF_Fin.Value) ', (k)
'    If CNN.rsCmdRepLineasGral.EOF <> True Then
'        CNN.rsCmdRepLineasGral.MoveFirst
'        m = 1
'        Do While CNN.rsCmdRepLineasGral.EOF <> True
'            Fecha = CNN.rsCmdRepLineasGral!Fecha
'            IdHora = CNN.rsCmdRepLineasGral!IdHora
'            OF = CNN.rsCmdRepLineasGral!OF
'            CNN.CmdBuscaEsp2 (OF)
'            If CNN.rsCmdBuscaEsp2.EOF <> True Then
'                NoParte = CNN.rsCmdBuscaEsp2!nodeparte
'            Else
'                CNN.CmdVidrio_MPS (OF)
'                If CNN.rsCmdVidrio_MPS.EOF <> True Then
'                    NoParte = CNN.rsCmdVidrio_MPS!Codigo
'                Else
'                    MsgBox "El PT o V MPS NO existe", vbInformation, MSG
'                    NoParte = "Vacio"
'                End If
'                CNN.rsCmdVidrio_MPS.Close
'            End If
'            CNN.rsCmdBuscaEsp2.Close
'            Renglon = (6 + (Fecha - Me.DTPF_Ini.Value) * 24) + IdHora
'
'
'            If InStr(1, CNN.rsCmdRepLineasGral!Descripcion, "1") Then
'                k = 1
'            End If
'            If InStr(1, CNN.rsCmdRepLineasGral!Descripcion, "2") Then
'                k = 2
'            End If
'            If InStr(1, CNN.rsCmdRepLineasGral!Descripcion, "3") Then
'                k = 3
'            End If
'
'            Select Case k
'                Case 1 '111
'                    Hoja.Cells(Renglon, 4) = NoParte
'                    If Mid(OF, 1, 1) <> "X" Then
'                        Hoja.Cells(Renglon, 3) = CNN.rsCmdRepLineasGral!jc
'                        Hoja.Cells(Renglon, 5) = CNN.rsCmdRepLineasGral!pzspt
'                    End If
'                Case 2 '221
'                    Hoja.Cells(Renglon, 7) = NoParte
'                    If Mid(OF, 1, 1) <> "X" Then
'                        Hoja.Cells(Renglon, 6) = CNN.rsCmdRepLineasGral!jc
'                        Hoja.Cells(Renglon, 8) = CNN.rsCmdRepLineasGral!pzspt
'                    End If
'                Case 3 '331
'                    Hoja.Cells(Renglon, 10) = NoParte
'                    If Mid(OF, 1, 1) <> "X" Then
'                        Hoja.Cells(Renglon, 9) = CNN.rsCmdRepLineasGral!jc
'                        Hoja.Cells(Renglon, 11) = CNN.rsCmdRepLineasGral!pzspt
'                    End If
'            End Select
'            CNN.rsCmdRepLineasGral.MoveNext
'        Loop
'    End If
'    CNN.rsCmdRepLineasGral.Close
'
''    k = k + 110
''Loop
'
'
'
'Hoja.Range("B6:K" & MaxRenglon).Select
'Hoja.Range("B6:K" & MaxRenglon).HorizontalAlignment = xlCenter
'Hoja.Range("B6:K" & MaxRenglon).VerticalAlignment = xlCenter
'
'Hoja.Rows("6:" & MaxRenglon + 5).RowHeight = 20.25
'Hoja.Columns("A:K").EntireColumn.AutoFit
'Hoja.Rows("3:3").RowHeight = 27.25
'
'Hoja.Range("A4:K" & MaxRenglon + 5).Select
'Hoja.Range("A4:K" & MaxRenglon + 5).VerticalAlignment = xlCenter
'Hoja.Range("A4:K" & MaxRenglon + 5).HorizontalAlignment = xlCenter
'Hoja.Range("A4:K" & MaxRenglon + 5).WrapText = False
'Hoja.Range("A4:K" & MaxRenglon + 5).AddIndent = False
'Hoja.Range("A4:K" & MaxRenglon + 5).IndentLevel = 0
'Hoja.Range("A4:K" & MaxRenglon + 5).ShrinkToFit = False
'Hoja.Range("A4:K" & MaxRenglon + 5).ReadingOrder = xlContext
'
'
'Hoja.Range("A4:K" & MaxRenglon + 5).Borders(xlDiagonalDown).LineStyle = xlNone
'Hoja.Range("A4:K" & MaxRenglon + 5).Borders(xlDiagonalUp).LineStyle = xlNone
'
'Hoja.Range("A4:K" & MaxRenglon + 5).Borders(xlEdgeLeft).LineStyle = xlContinuous
'Hoja.Range("A4:K" & MaxRenglon + 5).Borders(xlEdgeLeft).ColorIndex = 0
'Hoja.Range("A4:K" & MaxRenglon + 5).Borders(xlEdgeLeft).TintAndShade = 0
'Hoja.Range("A4:K" & MaxRenglon + 5).Borders(xlEdgeLeft).Weight = xlThin
'
'Hoja.Range("A4:K" & MaxRenglon + 5).Borders(xlEdgeBottom).LineStyle = xlContinuous
'Hoja.Range("A4:K" & MaxRenglon + 5).Borders(xlEdgeBottom).ColorIndex = 0
'Hoja.Range("A4:K" & MaxRenglon + 5).Borders(xlEdgeBottom).TintAndShade = 0
'Hoja.Range("A4:K" & MaxRenglon + 5).Borders(xlEdgeBottom).Weight = xlThin
'
'Hoja.Range("A4:K" & MaxRenglon + 5).Borders(xlEdgeRight).LineStyle = xlContinuous
'Hoja.Range("A4:K" & MaxRenglon + 5).Borders(xlEdgeRight).ColorIndex = 0
'Hoja.Range("A4:K" & MaxRenglon + 5).Borders(xlEdgeRight).TintAndShade = 0
'Hoja.Range("A4:K" & MaxRenglon + 5).Borders(xlEdgeRight).Weight = xlThin
'
'Hoja.Range("A4:K" & MaxRenglon + 5).Borders(xlEdgeTop).LineStyle = xlContinuous
'Hoja.Range("A4:K" & MaxRenglon + 5).Borders(xlEdgeTop).ColorIndex = 0
'Hoja.Range("A4:K" & MaxRenglon + 5).Borders(xlEdgeTop).TintAndShade = 0
'Hoja.Range("A4:K" & MaxRenglon + 5).Borders(xlEdgeTop).Weight = xlThin
'
'Hoja.Range("A4:K" & MaxRenglon + 5).Borders(xlInsideVertical).LineStyle = xlContinuous
'Hoja.Range("A4:K" & MaxRenglon + 5).Borders(xlInsideVertical).ColorIndex = 0
'Hoja.Range("A4:K" & MaxRenglon + 5).Borders(xlInsideVertical).TintAndShade = 0
'Hoja.Range("A4:K" & MaxRenglon + 5).Borders(xlInsideVertical).Weight = xlThin
'
'Hoja.Range("A4:K" & MaxRenglon + 5).Borders(xlInsideHorizontal).LineStyle = xlContinuous
'Hoja.Range("A4:K" & MaxRenglon + 5).Borders(xlInsideHorizontal).ColorIndex = 0
'Hoja.Range("A4:K" & MaxRenglon + 5).Borders(xlInsideHorizontal).TintAndShade = 0
'Hoja.Range("A4:K" & MaxRenglon + 5).Borders(xlInsideHorizontal).Weight = xlThin
'
'
'
'libro.ActiveSheet.PageSetup.LeftHeaderPicture.FileName = "C:\legacy\schott.JPG"
'libro.ActiveSheet.PageSetup.PrintTitleRows = "$1:$5"
'libro.ActiveSheet.PageSetup.PrintTitleColumns = ""
'
'libro.ActiveSheet.PageSetup.PrintArea = ""
'libro.ActiveSheet.PageSetup.LeftHeader = "&G"
'libro.ActiveSheet.PageSetup.CenterHeader = ""
'libro.ActiveSheet.PageSetup.RightHeader = ""
'libro.ActiveSheet.PageSetup.LeftFooter = "Imprimio: " & Nombre & vbNewLine & "FORMA #F-TEM-020-A(08/17/2009)"
'libro.ActiveSheet.PageSetup.CenterFooter = "&P"
'libro.ActiveSheet.PageSetup.RightFooter = "&T" & Chr(10) & "&D"
'libro.ActiveSheet.PageSetup.LeftMargin = libro.Application.InchesToPoints(0.236220472440945)
'libro.ActiveSheet.PageSetup.RightMargin = libro.Application.InchesToPoints(0.275590551181102)
'libro.ActiveSheet.PageSetup.TopMargin = libro.Application.InchesToPoints(0.275590551181102)
'libro.ActiveSheet.PageSetup.BottomMargin = libro.Application.InchesToPoints(0.98551181102362)
'libro.ActiveSheet.PageSetup.HeaderMargin = libro.Application.InchesToPoints(0.275590551181102)
'libro.ActiveSheet.PageSetup.FooterMargin = libro.Application.InchesToPoints(0.31496062992126)
'libro.ActiveSheet.PageSetup.PrintHeadings = False
'libro.ActiveSheet.PageSetup.PrintGridlines = False
'libro.ActiveSheet.PageSetup.PrintComments = xlPrintNoComments
'libro.ActiveSheet.PageSetup.PrintQuality = 600
'libro.ActiveSheet.PageSetup.CenterHorizontally = False
'libro.ActiveSheet.PageSetup.CenterVertically = False
'libro.ActiveSheet.PageSetup.Orientation = xlPortrait
'libro.ActiveSheet.PageSetup.Draft = False
'libro.ActiveSheet.PageSetup.PaperSize = xlPaperLetter
'libro.ActiveSheet.PageSetup.FirstPageNumber = xlAutomatic
'libro.ActiveSheet.PageSetup.Order = xlDownThenOver
'libro.ActiveSheet.PageSetup.BlackAndWhite = False
'libro.ActiveSheet.PageSetup.Zoom = 80
'libro.ActiveSheet.PageSetup.PrintErrors = xlPrintErrorsDisplayed
'libro.ActiveSheet.PageSetup.OddAndEvenPagesHeaderFooter = False
'libro.ActiveSheet.PageSetup.DifferentFirstPageHeaderFooter = False
'libro.ActiveSheet.PageSetup.ScaleWithDocHeaderFooter = True
'libro.ActiveSheet.PageSetup.AlignMarginsHeaderFooter = True
'libro.ActiveSheet.PageSetup.EvenPage.LeftHeader.Text = ""
'libro.ActiveSheet.PageSetup.EvenPage.CenterHeader.Text = ""
'libro.ActiveSheet.PageSetup.EvenPage.RightHeader.Text = ""
'libro.ActiveSheet.PageSetup.EvenPage.LeftFooter.Text = ""
'libro.ActiveSheet.PageSetup.EvenPage.CenterFooter.Text = ""
'libro.ActiveSheet.PageSetup.EvenPage.RightFooter.Text = ""
'libro.ActiveSheet.PageSetup.FirstPage.LeftHeader.Text = ""
'libro.ActiveSheet.PageSetup.FirstPage.CenterHeader.Text = ""
'libro.ActiveSheet.PageSetup.FirstPage.RightHeader.Text = ""
'libro.ActiveSheet.PageSetup.FirstPage.LeftFooter.Text = ""
'libro.ActiveSheet.PageSetup.FirstPage.CenterFooter.Text = ""
'libro.ActiveSheet.PageSetup.FirstPage.RightFooter.Text = ""
'
'Hoja.Range("K6").Select
'Randomize
'
'Ruta = "C:\mis documentos\ProgAcumulado " & Format(Day(Date), "00") & Format(Month(Date), "00") & Format(Year(Date), "00") & " " & Int((1000 * Rnd) + 1) & ".xlsx"
'
'Hoja.SaveAs (Ruta)
'libro.Close
'Set Hoja = Nothing
'Set libro = Nothing
'vApp.Quit
'Set vApp = Nothing
'
'Dim CnnExcel As New Excel.Application
'Dim LibroPDF As Excel.Workbook
''Abrimos el documento
'
'
'Resp = Int((1000 * Rnd) + 1)
'RutaPDF = "Q:\Materials\Read\Reporte Programa Produccion FLAT\ProgAcumulado " & Format(Day(Date), "00") & Format(Month(Date), "00") & Format(Year(Date), "00") & " " & Resp & ".pdf"
'
'
'RutaPDF = "\\10.18.172.30\Departments\Materials\Read\Reporte Programa Produccion Flat\ProgAcumulado " & Format(Day(Date), "00") & Format(Month(Date), "00") & Format(Year(Date), "00") & " " & Resp & ".pdf"
'
'
''\\10.18.164.207\e$\Adjuntos
'
''RutaPDF = "\\10.18.164.207\e$\Adjuntos\ProgAcumulado " & Format(Day(Date), "00") & Format(Month(Date), "00") & Format(Year(Date), "00") & " " & Resp & ".pdf"
'
'
'
'''''RutaPDF = "C:\mis documentos\PFLAT " & Format(Day(Date), "00") & Format(Month(Date), "00") & Format(Year(Date), "00") & " " & Int((1000 * Rnd) + 1) & ".pdf"
'
'CnnExcel.Workbooks.Open (Ruta)
'CnnExcel.ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:= _
'    RutaPDF, Quality:= _
'    xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
'    OpenAfterPublish:=False
'CnnExcel.ActiveWorkbook.Close
''''CnnExcel.LibroPDF.Close
'CnnExcel.Quit
'Set CnnExcel = Nothing
''''MsgBox "Tarea Terminada" & vbNewLine & "Su archivo esta en la sig ruta: " & Ruta & ".xlsx", vbInformation, MSG
'
'Cadena = "Programa de Producción Acumulado Flat " & Me.DTPF_Ini.Value & " a " & Me.DTPF_Fin.Value
'
''''CadenaCorreo = "El Programa de Producción de Flat se localiza en la siguiente ruta:<br><br>"
'''''CadenaCorreo = CadenaCorreo & "<a href='" & RutaPDF & "'>" & RutaPDF & "</a> <br><br>"
''''CadenaCorreo = CadenaCorreo & "Gracias"
'
'RutaAdjunto = "\\10.18.172.30\e$\Data\Departments\Materials\Read\Reporte Programa Produccion Flat\PFLAT " & Format(Day(Date), "00") & Format(Month(Date), "00") & Format(Year(Date), "00") & " " & Resp & ".pdf"
'
'RutaAdjunto = RutaPDF ' "\\10.18.164.207\e$\Adjuntos\ProgAcumulado " & Format(Day(Date), "00") & Format(Month(Date), "00") & Format(Year(Date), "00") & " " & Resp & ".pdf"
'
'
'
'Para = "hireta@sswtechnologies.com; agallegos@sswtechnologies.com; breyes@sswtechnologies.com; "
'
'
'CNN.rsCmdCorreo.Open
'CNN.rsCmdCorreo.AddNew
'    CNN.rsCmdCorreo!De = "jdiaz@sswtechnologies.com" 'De
'    CNN.rsCmdCorreo!Para = "rfabela@sswtechnologies.com, jdiaz@sswtechnologies.com, melissa.montoya@schott.com, agallegos@sswtechnologies.com,  benito.morales@schott.com, jesus.cruz@schott.com, abril.ramos@schott.com,  "
'    CNN.rsCmdCorreo!CC = "rfabela@sswtechnologies.com; "
'    CNN.rsCmdCorreo!CCO = "hireta@sswtechnologies.com; breyes@sswtechnologies.com; "
'    CNN.rsCmdCorreo!Titulo = Cadena '& " Pruebas TI"
'    CNN.rsCmdCorreo!Mensaje = CadenaCorreo
'    CNN.rsCmdCorreo!aplicacion = "JC"
'    CNN.rsCmdCorreo!Prioridad = 2
'    CNN.rsCmdCorreo!Ruta = RutaPDF  'RutaAdjunto
'CNN.rsCmdCorreo.Update
'CNN.rsCmdCorreo.Close
'
'Unload Me
'End Sub
'
'
'
''''Set vApp = Nothing
''''ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:= _
''''    "C:\Documents and Settings\heliosireta\Mis documentos\Cosolidado ProgPrd " & Format(Day(Date), "00") & Format(Month(Date), "00") & Format(Year(Date), "00") & ".pdf", Quality:= _
''''    xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
''''    OpenAfterPublish:=False
'
''''Randomize
''''ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:= _
''''    "\\10.18.164.226\Departments\Materials\Read\Reporte Programa Produccion Flat\PFLAT " & Format(Day(Date), "00") & Format(Month(Date), "00") & Format(Year(Date), "00") & " " & Int((1000 * Rnd) + 1) & ".pdf", Quality:= _
''''    xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
''''    OpenAfterPublish:=True
'
'''vApp.Visible = True
''Libro.Close
''Set Hoja = Nothing
''Set Libro = Nothing
''Set vApp = Nothing
''vApp.Quit
''''Set vApp = Nothing
'
''''ChDir "C:\Mis documentos\Reps Programa Flat"
''''    ActiveWorkbook.SaveAs FileName:= _
''''        "C:\Mis documentos\Reps Programa Flat\RFlat 200410.xlsx" _
''''        , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
'
''''vApp.Quit
''''Reset
''''Dim vApp As Excel.Application
''''Dim Libro As Excel.Workbook
''''Dim Hoja As Excel.Worksheet
''''Set vApp = CreateObject("Excel.Application")
''''Set Libro = vApp.Workbooks.Add
''''Set Hoja = Libro.Worksheets(1)
'
'
'
'
'






