Private Sub CmdVerReporte_Click()
'Reporte de produccion

'''If Me.DTPF_Ini.Value < Date Then
'''    MsgBox "Verifique las fechas, no puede sacar reportes para fechas pasadas.", vbExclamation, MSG
'''    Exit Sub
'''End If
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


Dim cLinea10 As Integer
Dim cLinea20 As Integer
Dim cLinea40CNC As Integer
Dim cLinea40CNC2 As Integer
Dim cLinea40w As Integer
Dim cLinea50 As Integer
Dim cLinea60 As Integer
Dim cLinea70 As Integer
Dim cLinea70i As Integer
 
   
vApp.Visible = True


'Encabezados
Hoja.Cells(4, 1) = "Fecha"
Hoja.Cells(5, 2) = "Hora"

cLinea10 = 3
Hoja.Cells(4, cLinea10) = "Linea 10"
Hoja.Cells(5, cLinea10) = "No.  JC"
Hoja.Cells(5, cLinea10 + 1) = "No de Parte"
Hoja.Cells(5, cLinea10 + 2) = "Cantidad"
Hoja.Cells(5, cLinea10 + 3) = "No. OMP SAP"

cLinea20 = 8
Hoja.Cells(4, cLinea20) = "Linea 20"
Hoja.Cells(5, cLinea20) = "No.  JC"
Hoja.Cells(5, cLinea20 + 1) = "No de Parte"
Hoja.Cells(5, cLinea20 + 2) = "Cantidad"
Hoja.Cells(5, cLinea20 + 3) = "No. OMP SAP"

cLinea40CNC = 13
Hoja.Cells(4, cLinea40CNC) = "Linea 40 CNC"
Hoja.Cells(5, cLinea40CNC) = "No.  JC"
Hoja.Cells(5, cLinea40CNC + 1) = "No de Parte"
Hoja.Cells(5, cLinea40CNC + 2) = "Cantidad"
Hoja.Cells(5, cLinea40CNC + 3) = "No. OMP SAP"

cLinea40CNC2 = 18
Hoja.Cells(4, cLinea40CNC2) = "Linea 40 CNC 2"
Hoja.Cells(5, cLinea40CNC2) = "No.  JC"
Hoja.Cells(5, cLinea40CNC2 + 1) = "No de Parte"
Hoja.Cells(5, cLinea40CNC2 + 2) = "Cantidad"
Hoja.Cells(5, cLinea40CNC2 + 3) = "No. OMP SAP"

cLinea40w = 23
Hoja.Cells(4, cLinea40w) = "Linea 40 WaterJet"
Hoja.Cells(5, cLinea40w) = "No.  JC"
Hoja.Cells(5, cLinea40w + 1) = "No de Parte"
Hoja.Cells(5, cLinea40w + 2) = "Cantidad"
Hoja.Cells(5, cLinea40w + 3) = "No. OMP SAP"

'Hoja.Cells(4, 21) = "Linea 40 CNC 2"
'Hoja.Cells(5, 21) = "No.  JC"
'Hoja.Cells(5, 22) = "No de Parte"
'Hoja.Cells(5, 23) = "Cantidad"

cLinea50 = 28
Hoja.Cells(4, cLinea50) = "Linea 50"
Hoja.Cells(5, cLinea50) = "No.  JC"
Hoja.Cells(5, cLinea50 + 1) = "No de Parte"
Hoja.Cells(5, cLinea50 + 2) = "Cantidad"
Hoja.Cells(5, cLinea50 + 3) = "No. OMP SAP"

cLinea60 = 33
Hoja.Cells(4, cLinea60) = "Linea 60"
Hoja.Cells(5, cLinea60) = "No.  JC"
Hoja.Cells(5, cLinea60 + 1) = "No de Parte"
Hoja.Cells(5, cLinea60 + 2) = "Cantidad"
Hoja.Cells(5, cLinea60 + 3) = "No. OMP SAP"

cLinea70 = 38
Hoja.Cells(4, cLinea70) = "Linea 70"
Hoja.Cells(5, cLinea70) = "No.  JC"
Hoja.Cells(5, cLinea70 + 1) = "No de Parte"
Hoja.Cells(5, cLinea70 + 2) = "Cantidad"
Hoja.Cells(5, cLinea70 + 3) = "No. OMP SAP"

cLinea70i = 43
Hoja.Cells(4, cLinea70i) = "Linea 70 Impresoras"
Hoja.Cells(5, cLinea70i) = "No.  JC"
Hoja.Cells(5, cLinea70i + 1) = "No de Parte"
Hoja.Cells(5, cLinea70i + 2) = "Cantidad"
Hoja.Cells(5, cLinea70i + 3) = "No. OMP SAP"
 


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


Hoja.Range("A4:AL5").Select
Hoja.Range("A4:AL5").Interior.Pattern = xlSolid
Hoja.Range("A4:AL5").Interior.PatternColorIndex = xlAutomatic
Hoja.Range("A4:AL5").Interior.PatternTintAndShade = 0
Hoja.Range("A4:AL5").Interior.ThemeColor = xlThemeColorLight2
Hoja.Range("A4:AL5").Interior.TintAndShade = 0.799981688894314


' para los Codigos de MP
Renglon = 6
'''Bordes
Hoja.Range("A4:AL5").Select
Hoja.Range("A4:AL5").VerticalAlignment = xlCenter
Hoja.Range("A4:AL5").HorizontalAlignment = xlCenter

Hoja.Range("A4:AL5").Select

Hoja.Range("A4:AL5").Borders(xlDiagonalDown).LineStyle = xlNone

Hoja.Range("A4:AL5").Borders(xlDiagonalDown).LineStyle = xlNone
Hoja.Range("A4:AL5").Borders(xlDiagonalUp).LineStyle = xlNone

Hoja.Range("A4:AL5").Borders(xlEdgeLeft).LineStyle = xlContinuous
Hoja.Range("A4:AL5").Borders(xlEdgeLeft).ColorIndex = 0
Hoja.Range("A4:AL5").Borders(xlEdgeLeft).TintAndShade = 0
Hoja.Range("A4:AL5").Borders(xlEdgeLeft).Weight = xlThin

Hoja.Range("A4:AL5").Borders(xlEdgeBottom).LineStyle = xlContinuous
Hoja.Range("A4:AL5").Borders(xlEdgeBottom).ColorIndex = 0
Hoja.Range("A4:AL5").Borders(xlEdgeBottom).TintAndShade = 0
Hoja.Range("A4:AL5").Borders(xlEdgeBottom).Weight = xlThin

Hoja.Range("A4:AL5").Borders(xlEdgeRight).LineStyle = xlContinuous
Hoja.Range("A4:AL5").Borders(xlEdgeRight).ColorIndex = 0
Hoja.Range("A4:AL5").Borders(xlEdgeRight).TintAndShade = 0
Hoja.Range("A4:AL5").Borders(xlEdgeRight).Weight = xlThin

Hoja.Range("A4:AL5").Borders(xlEdgeTop).LineStyle = xlContinuous
Hoja.Range("A4:AL5").Borders(xlEdgeTop).ColorIndex = 0
Hoja.Range("A4:AL5").Borders(xlEdgeTop).TintAndShade = 0
Hoja.Range("A4:AL5").Borders(xlEdgeTop).Weight = xlThin

Hoja.Range("A4:AL5").Borders(xlInsideVertical).LineStyle = xlContinuous
Hoja.Range("A4:AL5").Borders(xlInsideVertical).ColorIndex = 0
Hoja.Range("A4:AL5").Borders(xlInsideVertical).TintAndShade = 0
Hoja.Range("A4:AL5").Borders(xlInsideVertical).Weight = xlThin

Hoja.Range("A4:AL5").Borders(xlInsideHorizontal).LineStyle = xlContinuous
Hoja.Range("A4:AL5").Borders(xlInsideHorizontal).ColorIndex = 0
Hoja.Range("A4:AL5").Borders(xlInsideHorizontal).TintAndShade = 0
Hoja.Range("A4:AL5").Borders(xlInsideHorizontal).Weight = xlThin

    
Hoja.Range("A5:AL5").Select
Hoja.Range("A5:AL5").Font.Name = "Times New Roman"
Hoja.Range("A5:AL5").Font.Size = 11
Hoja.Range("A5:AL5").Font.Strikethrough = False
Hoja.Range("A5:AL5").Font.Superscript = False
Hoja.Range("A5:AL5").Font.Subscript = False
Hoja.Range("A5:AL5").Font.OutlineFont = False
Hoja.Range("A5:AL5").Font.Shadow = False
Hoja.Range("A5:AL5").Font.Underline = xlUnderlineStyleNone
Hoja.Range("A5:AL5").Font.ThemeColor = xlThemeColorLight1
Hoja.Range("A5:AL5").Font.TintAndShade = 0
Hoja.Range("A5:AL5").Font.ThemeFont = xlThemeFontNone
Hoja.Range("A5:AL5").Font.Bold = True


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


'para todaS las lineas
Renglon = 6
k = 10
Do While k <= 23
    
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
                Case 10
                    Hoja.Cells(Renglon, cLinea10 + 1) = NoParte
                    If Mid(OF, 1, 1) <> "X" Then
                        Hoja.Cells(Renglon, cLinea10) = CNN.rsCmdRepLineas!jc
                        Hoja.Cells(Renglon, cLinea10 + 2) = CNN.rsCmdRepLineas!pzspt
                       
                        Hoja.Cells(Renglon, cLinea10 + 3) = CNN.rsCmdRepLineas!OrdenProd_SAP
                    End If
                    
                    
                Case 20
                    Hoja.Cells(Renglon, cLinea20 + 1) = NoParte
                    If Mid(OF, 1, 1) <> "X" Then
                        Hoja.Cells(Renglon, cLinea20) = CNN.rsCmdRepLineas!jc
                        Hoja.Cells(Renglon, cLinea20 + 2) = CNN.rsCmdRepLineas!pzspt
                       '  Hoja.Cells(Renglon - 1, 31) = k
                        Hoja.Cells(Renglon, cLinea20 + 3) = CNN.rsCmdRepLineas!OrdenProd_SAP
                    End If
                    
'                Case 30
'                    Hoja.Cells(Renglon, 10) = NoParte
'                    If Mid(OF, 1, 1) <> "X" Then
'                        Hoja.Cells(Renglon, 9) = CNN.rsCmdRepLineas!jc
'                        Hoja.Cells(Renglon, 11) = CNN.rsCmdRepLineas!pzspt
'                        ' Hoja.Cells(Renglon - 1, 32) = k
'                        Hoja.Cells(Renglon, 32) = CNN.rsCmdRepLineas!OrdenProd_SAP
'                    End If
            End Select
            CNN.rsCmdRepLineas.MoveNext
        Loop
    End If
    CNN.rsCmdRepLineas.Close
    k = k + 1
Loop



'60:  Linea 40 CNC
Renglon = 6
k = 60
Do While k <= 60
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
                Case 60
                    Hoja.Cells(Renglon, cLinea60 + 1) = NoParte
                    If Mid(OF, 1, 1) <> "X" Then
                        Hoja.Cells(Renglon, cLinea60) = CNN.rsCmdRepLineas!jc
                        Hoja.Cells(Renglon, cLinea60 + 2) = CNN.rsCmdRepLineas!pzspt
                        Hoja.Cells(Renglon, cLinea60 + 3) = CNN.rsCmdRepLineas!OrdenProd_SAP
                        
                    End If
            End Select
            CNN.rsCmdRepLineas.MoveNext
        Loop
    End If
    CNN.rsCmdRepLineas.Close
    k = k + 1
Loop

'60  Linea 40 CNC   2
Renglon = 6
k = 32

Do While k <= 32
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
                Case 32
                    Hoja.Cells(Renglon, cLinea40CNC2 + 1) = NoParte
                    If Mid(OF, 1, 1) <> "X" Then
                        Hoja.Cells(Renglon, cLinea40CNC2) = CNN.rsCmdRepLineas!jc
                        Hoja.Cells(Renglon, cLinea40CNC2 + 2) = CNN.rsCmdRepLineas!pzspt
                        Hoja.Cells(Renglon, cLinea40CNC2 + 3) = CNN.rsCmdRepLineas!OrdenProd_SAP
                    
                    End If
            End Select
            CNN.rsCmdRepLineas.MoveNext
        Loop
    End If
    CNN.rsCmdRepLineas.Close
    k = k + 1
Loop

'240  Linea 40 WaterJet
Renglon = 6
k = 240
Do While k <= 240

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
                Case 240
                    Hoja.Cells(Renglon, cLinea40w + 1) = NoParte
                    If Mid(OF, 1, 1) <> "X" Then
                        Hoja.Cells(Renglon, cLinea40w) = CNN.rsCmdRepLineas!jc
                        Hoja.Cells(Renglon, cLinea40w + 2) = CNN.rsCmdRepLineas!pzspt
                        Hoja.Cells(Renglon, cLinea40w + 3) = CNN.rsCmdRepLineas!OrdenProd_SAP
                    End If
            End Select
            CNN.rsCmdRepLineas.MoveNext
        Loop
    End If
    CNN.rsCmdRepLineas.Close
    k = k + 1
Loop

vApp.Visible = True
'21  Linea 50
Renglon = 6
k = 21
Do While k <= 21
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
                Case 21
                    Hoja.Cells(Renglon, cLinea50 + 1) = NoParte
                    If Mid(OF, 1, 1) <> "X" Then
                        Hoja.Cells(Renglon, cLinea50) = CNN.rsCmdRepLineas!jc
                        Hoja.Cells(Renglon, cLinea50 + 2) = CNN.rsCmdRepLineas!pzspt
                        Hoja.Cells(Renglon, cLinea50 + 3) = CNN.rsCmdRepLineas!OrdenProd_SAP
                    End If
            End Select
            CNN.rsCmdRepLineas.MoveNext
        Loop
    End If
    CNN.rsCmdRepLineas.Close
    k = k + 1
Loop


'21  Linea 60
Renglon = 6
k = 70
Do While k <= 70
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
                Case 70
                    Hoja.Cells(Renglon, cLinea60 + 1) = NoParte
                    If Mid(OF, 1, 1) <> "X" Then
                        Hoja.Cells(Renglon, cLinea60) = CNN.rsCmdRepLineas!jc
                        Hoja.Cells(Renglon, cLinea60 + 2) = CNN.rsCmdRepLineas!pzspt
                        Hoja.Cells(Renglon, cLinea60 + 3) = CNN.rsCmdRepLineas!OrdenProd_SAP
                    End If
            End Select
            CNN.rsCmdRepLineas.MoveNext
        Loop
    End If
    CNN.rsCmdRepLineas.Close
    k = k + 1
Loop


'190   Linea 70 CNC
Renglon = 6
k = 190
Do While k <= 190
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
                Case 190
                    Hoja.Cells(Renglon, cLinea70 + 1) = NoParte
                    If Mid(OF, 1, 1) <> "X" Then
                        Hoja.Cells(Renglon, cLinea70) = CNN.rsCmdRepLineas!jc
                        Hoja.Cells(Renglon, cLinea70 + 2) = CNN.rsCmdRepLineas!pzspt
                         Hoja.Cells(Renglon, cLinea70 + 3) = CNN.rsCmdRepLineas!OrdenProd_SAP
                    End If
            End Select
            CNN.rsCmdRepLineas.MoveNext
        Loop
    End If
    CNN.rsCmdRepLineas.Close
    k = k + 1
Loop


'
''193 linea 70 impresoras
Renglon = 6
k = 193
Do While k <= 193
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
                Case 193
                    Hoja.Cells(Renglon, cLinea70i + 1) = NoParte
                    If Mid(OF, 1, 1) <> "X" Then
                        Hoja.Cells(Renglon, cLinea70i) = CNN.rsCmdRepLineas!jc
                        Hoja.Cells(Renglon, cLinea70i + 2) = CNN.rsCmdRepLineas!pzspt
                         Hoja.Cells(Renglon, cLinea70i + 3) = CNN.rsCmdRepLineas!OrdenProd_SAP
                    End If
            End Select
            CNN.rsCmdRepLineas.MoveNext
        Loop
    End If
    CNN.rsCmdRepLineas.Close
    k = k + 1
Loop


''Linea 10
'Hoja.Columns("AD:AD").Cut
'Hoja.Columns("F:F").Insert Shift:=xlToRight
'
''Linea 20
'Hoja.Columns("AE:AE").Cut
'Hoja.Columns("J:J").Insert Shift:=xlToRight
'
''Linea 40
'Hoja.Columns("AF:AF").Cut
'Hoja.Columns("N:N").Insert Shift:=xlToRight
'
''Linea 40 CNC2
'Hoja.Columns("AG:AG").Cut
'Hoja.Columns("R:R").Insert Shift:=xlToRight
'
''Linea 40 WATERJET
'Hoja.Columns("AH:AH").Cut
'Hoja.Columns("V:V").Insert Shift:=xlToRight
'
''Linea 50
'Hoja.Columns("AI:AI").Cut
'Hoja.Columns("Z:Z").Insert Shift:=xlToRight
'
''Linea 60
'Hoja.Columns("AJ:AJ").Cut
'Hoja.Columns("AD:AD").Insert Shift:=xlToRight
'
''Linea 70
'Hoja.Columns("AK:AK").Cut
'Hoja.Columns("AH:AH").Insert Shift:=xlToRight

'Linea 70 impresoras

'Corta y pega L70
'Hoja.Columns("I:K").Cut
'Hoja.Selection.Cut
'Hoja.Columns("U:U").Insert Shift:=xlToRight
'Hoja.Range("R4:T4").Select


'Corta y pega L50
'Hoja.Columns("I:K").Cut
'Hoja.Columns("AD:AD").Insert Shift:=xlToRight


Hoja.Range("C4:F4").Select
Hoja.Range("C4:F4").HorizontalAlignment = xlCenter
Hoja.Range("C4:F4").MergeCells = True

'Linea 20
Hoja.Range("G4:K4").Select
Hoja.Range("G4:K4").HorizontalAlignment = xlCenter
Hoja.Range("G4:K4").MergeCells = True

''Linea 40
Hoja.Range("L4:P4").Select
Hoja.Range("L4:P4").HorizontalAlignment = xlCenter
Hoja.Range("L4:P4").MergeCells = True

''Linea 40 CNC2
Hoja.Range("Q4:U4").Select
Hoja.Range("Q4:U4").HorizontalAlignment = xlCenter
Hoja.Range("Q4:U4").MergeCells = True

'Linea 40 WaterJet
Hoja.Range("V4:Z4").Select
Hoja.Range("V4:Z4").HorizontalAlignment = xlCenter
Hoja.Range("V4:Z4").MergeCells = True

'Linea 50
Hoja.Range("AA4:AE4").Select
Hoja.Range("AA4:AE4").HorizontalAlignment = xlCenter
Hoja.Range("AA4:AE4").MergeCells = True

''Linea 50
Hoja.Range("AA4:AD4").Select
Hoja.Range("AA4:AD4").HorizontalAlignment = xlCenter
Hoja.Range("AA4:AD4").MergeCells = True

'Linea 60
Hoja.Range("AE4:AH4").Select
Hoja.Range("AE4:AH4").HorizontalAlignment = xlCenter
Hoja.Range("AE4:AH4").MergeCells = True


'Linea 60
Hoja.Range("AI4:AL4").Select
Hoja.Range("AI4:AL4").HorizontalAlignment = xlCenter
Hoja.Range("AI4:AL4").MergeCells = True

'hoja.Sheets(Resp).Cells(1, 1) = "Gemtron de México S.A. de C.V."
Hoja.Cells(1, 1) = "Gemtron de México S.A. de C.V."


Hoja.Range("B6:AL" & MaxRenglon).Select
Hoja.Range("B6:AL" & MaxRenglon).HorizontalAlignment = xlCenter
Hoja.Range("B6:AL" & MaxRenglon).VerticalAlignment = xlCenter

Hoja.Rows("6:" & MaxRenglon + 5).RowHeight = 20.25
Hoja.Columns("A:AL").EntireColumn.AutoFit
Hoja.Rows("3:3").RowHeight = 27.25

Hoja.Range("A4:AL" & MaxRenglon + 5).Select
Hoja.Range("A4:AL" & MaxRenglon + 5).VerticalAlignment = xlCenter
Hoja.Range("A4:AL" & MaxRenglon + 5).HorizontalAlignment = xlCenter
Hoja.Range("A4:AL" & MaxRenglon + 5).WrapText = False
Hoja.Range("A4:AL" & MaxRenglon + 5).AddIndent = False
Hoja.Range("A4:AL" & MaxRenglon + 5).IndentLevel = 0
Hoja.Range("A4:AL" & MaxRenglon + 5).ShrinkToFit = False
Hoja.Range("A4:AL" & MaxRenglon + 5).ReadingOrder = xlContext
    
    
Hoja.Range("A4:AL" & MaxRenglon + 5).Borders(xlDiagonalDown).LineStyle = xlNone
Hoja.Range("A4:AL" & MaxRenglon + 5).Borders(xlDiagonalUp).LineStyle = xlNone

Hoja.Range("A4:AL" & MaxRenglon + 5).Borders(xlEdgeLeft).LineStyle = xlContinuous
Hoja.Range("A4:AL" & MaxRenglon + 5).Borders(xlEdgeLeft).ColorIndex = 0
Hoja.Range("A4:AL" & MaxRenglon + 5).Borders(xlEdgeLeft).TintAndShade = 0
Hoja.Range("A4:AL" & MaxRenglon + 5).Borders(xlEdgeLeft).Weight = xlThin

Hoja.Range("A4:AL" & MaxRenglon + 5).Borders(xlEdgeBottom).LineStyle = xlContinuous
Hoja.Range("A4:AL" & MaxRenglon + 5).Borders(xlEdgeBottom).ColorIndex = 0
Hoja.Range("A4:AL" & MaxRenglon + 5).Borders(xlEdgeBottom).TintAndShade = 0
Hoja.Range("A4:AL" & MaxRenglon + 5).Borders(xlEdgeBottom).Weight = xlThin

Hoja.Range("A4:AL" & MaxRenglon + 5).Borders(xlEdgeRight).LineStyle = xlContinuous
Hoja.Range("A4:AL" & MaxRenglon + 5).Borders(xlEdgeRight).ColorIndex = 0
Hoja.Range("A4:AL" & MaxRenglon + 5).Borders(xlEdgeRight).TintAndShade = 0
Hoja.Range("A4:AL" & MaxRenglon + 5).Borders(xlEdgeRight).Weight = xlThin

Hoja.Range("A4:AL" & MaxRenglon + 5).Borders(xlEdgeTop).LineStyle = xlContinuous
Hoja.Range("A4:AL" & MaxRenglon + 5).Borders(xlEdgeTop).ColorIndex = 0
Hoja.Range("A4:AL" & MaxRenglon + 5).Borders(xlEdgeTop).TintAndShade = 0
Hoja.Range("A4:AL" & MaxRenglon + 5).Borders(xlEdgeTop).Weight = xlThin

Hoja.Range("A4:AL" & MaxRenglon + 5).Borders(xlInsideVertical).LineStyle = xlContinuous
Hoja.Range("A4:AL" & MaxRenglon + 5).Borders(xlInsideVertical).ColorIndex = 0
Hoja.Range("A4:AL" & MaxRenglon + 5).Borders(xlInsideVertical).TintAndShade = 0
Hoja.Range("A4:AL" & MaxRenglon + 5).Borders(xlInsideVertical).Weight = xlThin

Hoja.Range("A4:AL" & MaxRenglon + 5).Borders(xlInsideHorizontal).LineStyle = xlContinuous
Hoja.Range("A4:AL" & MaxRenglon + 5).Borders(xlInsideHorizontal).ColorIndex = 0
Hoja.Range("A4:AL" & MaxRenglon + 5).Borders(xlInsideHorizontal).TintAndShade = 0
Hoja.Range("A4:AL" & MaxRenglon + 5).Borders(xlInsideHorizontal).Weight = xlThin

Hoja.Range("A4:AD4").Select
Hoja.Range("A4:AD4").Font.Name = "Times New Roman"
Hoja.Range("A4:AD4").Font.Size = 11
Hoja.Range("A4:AD4").Font.Bold = True

Hoja.Range("A1:k1").Select
Hoja.Range("A1:k1").HorizontalAlignment = xlCenter
Hoja.Range("A1:k1").VerticalAlignment = xlBottom
Hoja.Range("A1:k1").WrapText = False
Hoja.Range("A1:k1").MergeCells = True

Hoja.Range("A1:k1").Select
Hoja.Range("A1:k1").Font.Name = "Times New Roman"
Hoja.Range("A1:k1").Font.Size = 16
Hoja.Range("A1:k1").Font.Bold = True


Hoja.Cells(2, 1) = "Reporte Programa de Producción"
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



NoFormaISO = ""


'CNN.CmdNoForrma ("ProgProd")
'If CNN.rsCmdNoForrma.EOF Or CNN.rsCmdNoForrma.BOF Then
'    MsgBox "No hay Resultados par la busqueda", vbCritical, "Formato de Calidad"
'
'    NoFormaISO = "Forma #F-8.5.2-1-Q (Jul.20'20)"
'
'Else
'        NoFormaISO = CNN.rsCmdNoForrma!NOFORMA_2015
'
'End If
'CNN.rsCmdNoForrma.Close


'Report.ParameterFields(1).ClearCurrentValueAndRange
'Report.ParameterFields(1).AddCurrentValue (CStr(leyendaEtiqueta))

NoFormaISO = "Forma #F-8.5.2-1-Q (Jul.20'20)"


libro.ActiveSheet.PageSetup.LeftHeaderPicture.FileName = "C:\legacy\schott.JPG"
libro.ActiveSheet.PageSetup.PrintTitleRows = "$1:$5"
libro.ActiveSheet.PageSetup.PrintTitleColumns = ""

libro.ActiveSheet.PageSetup.PrintArea = ""
libro.ActiveSheet.PageSetup.LeftHeader = "&G"
libro.ActiveSheet.PageSetup.CenterHeader = ""
libro.ActiveSheet.PageSetup.RightHeader = ""
libro.ActiveSheet.PageSetup.LeftFooter = NoFormaISO & vbNewLine & "Imprimio: " & Nombre & vbNewLine '& "FORMA #F-TEM-020-A(08/17/2009)"
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
libro.ActiveSheet.PageSetup.PaperSize = xlPaperLegal  ' xlPaperLetter
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
Randomize

Cadena = "PFLAT" & Format(Day(Date), "00") & Format(Month(Date), "00") & Format(Year(Date), "00") & "" & Int((1000 * Rnd) + 1)
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
Cadena = "Programa de Producción Flat " & Me.DTPF_Ini.Value & " a " & Me.DTPF_Fin.Value
CadenaCorreo = "Se adjunta el Programa de Producción Flat<br><br>"
'CadenaCorreo = CadenaCorreo & "<a href='" & RutaPDF & "'>" & RutaPDF & "</a> <br><br>"
CadenaCorreo = CadenaCorreo & "Gracias"


Randomize

'Cadena = "PFLAT " & Format(Day(Date), "00") & Format(Month(Date), "00") & Format(Year(Date), "00") & " " & Int((1000 * Rnd) + 1)

RutaPDF = "Q:\Production\Read\Reps\ProgFlatFab" & Format(Day(Date), "00") & Format(Month(Date), "00") & Format(Year(Date), "00") & "" & Int((1000 * Rnd) + 1) & ".pdf"

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


CNN.CmdReportes ("RepProdGral")
If CNN.rsCmdReportes.EOF <> True Then
        CNN.rsCmdReportes!rutareporte = RutaPDF
    CNN.rsCmdReportes.Update
End If
CNN.rsCmdReportes.Close


Para = "breyes@sswtechnologies.com; fvillanueva@sswtechnologies.com, creyes@sswtechnologies.com, "
Para = Para & " flucio@sswtechnologies.com, hhernandez@sswtechnologies.com; jibarra@sswtechnologies.com, jalmendarez@sswtechnologies.com,  jmruiz@sswtechnologies.com; "
Para = Para & " marriaga@sswtechnologies.com, rsalas@sswtechnologies.com, sneri@sswtechnologies.com, jarroyos@sswtechnologies.com, gfernandez@sswtechnologies.com, " '

CNN.rsCmdCorreo.Open
CNN.rsCmdCorreo.AddNew
    CNN.rsCmdCorreo!De = "creyes@sswtechnologies.com" 'De
    CNN.rsCmdCorreo!Para = Para
    CNN.rsCmdCorreo!CC = "hireta@sswtechnologies.com, jlara@sswtechnologies.com"
    CNN.rsCmdCorreo!CCO = ""
    CNN.rsCmdCorreo!Titulo = Cadena '& " Pruebas TI"
    CNN.rsCmdCorreo!Mensaje = CadenaCorreo
    CNN.rsCmdCorreo!aplicacion = "JC"
    CNN.rsCmdCorreo!Prioridad = 2
    CNN.rsCmdCorreo!Ruta = RutaPDF2
CNN.rsCmdCorreo.Update
CNN.rsCmdCorreo.Close

MsgBox "Tarea Terminada " & vbNewLine & "Su archivo esta en la sig ruta: " & Ruta & " " & vbNewLine & vbNewLine & " El Reporte de Produccion ya se publico. ", vbInformation, MSG


MsgBox "El correo del Programa tambien se envio a los sigueintes usuarios: " & vbNewLine & Para, vbInformation, MSG


Unload Me



End Sub






'''Set vApp = Nothing
'''ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:= _
'''    "C:\Documents and Settings\heliosireta\Mis documentos\Cosolidado ProgPrd " & Format(Day(Date), "00") & Format(Month(Date), "00") & Format(Year(Date), "00") & ".pdf", Quality:= _
'''    xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
'''    OpenAfterPublish:=False
    
'''Randomize
'''ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:= _
'''    "\\10.18.172.5\Departments\Materials\Read\Reporte Programa Produccion Flat\PFLAT " & Format(Day(Date), "00") & Format(Month(Date), "00") & Format(Year(Date), "00") & " " & Int((1000 * Rnd) + 1) & ".pdf", Quality:= _
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

