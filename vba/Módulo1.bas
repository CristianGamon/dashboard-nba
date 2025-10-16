Attribute VB_Name = "Módulo1"
Option Explicit

Dim CFecha As Range
Dim FBD As Range
Dim celda As Range

Sub ActivarFormulario3()
    UserForm3.Show
    
End Sub

Sub ActivarFormulario1()

UserForm1.Show

End Sub

Sub DescargaResultados(FechaCompleta As Long)

    On Error Resume Next
    Dim ELocal As String, EVisitante As String, NombreConsulta As String, NombreWeb As String
    Dim i As Long
    
    FechaBD 'Llamada a macro para seleccionar columna Fecha de la BD
    
    
    For Each celda In Selection 'Comprobamos que el partido no haya sido ya descargado
        If celda.Value <> FechaCompleta Then
           'Nada
        Else
            MsgBox "Partidos correspondientes a esta fecha ya descargados", vbExclamation
            Hoja2.Activate
            Exit Sub
        End If
    Next celda
    
    ColumnaFecha 'Llamada a macro para seleccionar columna Fecha de NBACalendar23_24
    
    For Each celda In CFecha 'Descargamos los datos de cada partido desde la web
        If celda.Value = FechaCompleta Then
            celda.Offset(0, 4).Copy Destination:=Hoja20.Range("A1").End(xlDown).Offset(1, 1)
            celda.Offset(0, 2).Copy Destination:=Hoja20.Range("A1").End(xlDown).Offset(2, 1)
            celda.Offset(0, 5).Copy Destination:=Hoja20.Range("A1").End(xlDown).Offset(2, 12)
            celda.Offset(0, 6).Copy Destination:=Hoja20.Range("A1").End(xlDown).Offset(1, 12)
            With Hoja20.Range("A1").End(xlDown)
                .Offset(1, 0).Value = "Local"
                .Offset(1, 2).Value = FechaCompleta
                .Offset(2, 0).Value = "Visitor"
                .Offset(2, 2).Value = FechaCompleta
            End With
            
            ELocal = Hoja20.Range("A1").End(xlDown).Offset(-1, 1).Value
            EVisitante = Hoja20.Range("A1").End(xlDown).Offset(0, 1).Value
            NombreConsulta = FechaCompleta & ELocal & EVisitante
            NombreWeb = "https://www.hispanosnba.com/estadistica/partidos/" & FechaCompleta & "/" & ELocal & "-vs-" & EVisitante & ""
            
            ActiveWorkbook.Queries.Add Name:=NombreConsulta, Formula:= _
            "let" & Chr(13) & "" & Chr(10) & "    Origen = Web.Page(Web.Contents(""" & NombreWeb & """))," & Chr(13) & "" & Chr(10) & "    Data0 = Origen{0}[Data]," & Chr(13) & "" & Chr(10) & "    #""Tipo cambiado"" = Table.TransformColumnTypes(Data0,{{"""", type text}, {""2"", type text}, {""3"", Int64.Type}, {""1"", Int64.Type}, {""22"", Int64.Type}, {""32"", Int64.Type}, {""4"", Int64.Type}})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & " " & _
            "   #""Tipo cambiado"""
            ActiveWorkbook.Worksheets.Add
            With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
                "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & NombreConsulta & ";Extended Properties=""""" _
                , Destination:=Range("$A$1")).QueryTable
                .CommandType = xlCmdSql
                .CommandText = Array("SELECT * FROM [" & NombreConsulta & "]")
                .RowNumbers = False
                .FillAdjacentFormulas = False
                .PreserveFormatting = True
                .RefreshOnFileOpen = False
                .BackgroundQuery = True
                .RefreshStyle = xlInsertDeleteCells
                .SavePassword = False
                .SaveData = True
                .AdjustColumnWidth = True
                .RefreshPeriod = 0
                .PreserveColumnInfo = True
                .ListObject.DisplayName = NombreConsulta
                .Refresh BackgroundQuery:=False
            End With
            
            RellenarBD 'Llamada a macro para rellenar los datos en la BD
    
        End If
    Next celda
Hoja2.Activate
End Sub

Sub FechaBD() 'Selección columna Fecha de la BD

    Dim FBD As Range
    
    Hoja20.Activate
    Hoja20.Range("A1").Offset(0, 2).Select
    Set FBD = Range(Selection, Selection.End(xlDown))
    FBD.Select

End Sub

Sub ColumnaFecha() 'Selección columna Fecha de NBACalendar23_24

    Dim CFecha2 As Integer
    
    CFecha2 = Hoja3.Range("a1").CurrentRegion.Rows.Count - 1
    
    Set CFecha = Hoja3.Range("a1").CurrentRegion.Offset(1, 0).Resize(CFecha2, 1)

End Sub

Sub RellenarBD() 'Rellenamos la BD con los datos descargados

    Dim ConteoHojas, i, NumeroColumnas, NumeroFilas As Integer

    ConteoHojas = ThisWorkbook.Sheets.Count

    For i = 1 To ConteoHojas
    
    NumeroColumnas = Sheets(i).Range("A1").CurrentRegion.Columns.Count
    NumeroFilas = Sheets(i).Range("A1").CurrentRegion.Rows.Count
    
        If Sheets(i).Name <> "Teams23_24" And Sheets(i).Name <> "BD" And _
        Sheets(i).Name <> "NBACalendar23_24" And Sheets(i).Name <> "TablasFechas" And _
        Sheets(i).Name <> "Dashboard" And Sheets(i).Name <> "RegistroApuestas" And Sheets(i).Name <> "HistorialApuestas" Then
        Sheets(i).Activate
        Range("A1").CurrentRegion.Offset(1, 3).Resize(NumeroFilas - 1, NumeroColumnas - 3).Select
        Selection.Copy Destination:=Hoja20.Range("A1").End(xlDown).Offset(-1, 3)
        Application.DisplayAlerts = False
        Sheets(i).Delete
        Application.DisplayAlerts = True
            
        VictoriaDerrota 'Llamada a macro para determinar Victoria o Derrota
            
        DeterminarFavorito ' Macro para determinar el favorito
        
        Exit For
        Else
            'Nada
        End If
    Next i
GanadorDescanso
End Sub

Sub VictoriaDerrota()

    Dim PuntosVisitor, PuntosLocal, ResultadoVisitor, ResultadoLocal As Range
        
    Set PuntosVisitor = Hoja20.Range("A1").End(xlDown).Offset(0, 8)
    Set PuntosLocal = Hoja20.Range("A1").End(xlDown).Offset(-1, 8)
    Set ResultadoVisitor = Hoja20.Range("A1").End(xlDown).Offset(0, 11)
    Set ResultadoLocal = Hoja20.Range("A1").End(xlDown).Offset(-1, 11)
    
    If PuntosVisitor.Value > PuntosLocal.Value Then
        ResultadoLocal.Value = "D"
        ResultadoVisitor.Value = "V"
    Else
        ResultadoLocal.Value = "V"
        ResultadoVisitor.Value = "D"
    End If

End Sub

Sub DeterminarFavorito()

    Dim CuotaVisitor, CuotaLocal, FavoritoVisitor, FavoritoLocal As Range
    
    Set CuotaVisitor = Hoja20.Range("A1").End(xlDown).Offset(0, 12)
    Set CuotaLocal = Hoja20.Range("A1").End(xlDown).Offset(-1, 12)
    Set FavoritoVisitor = Hoja20.Range("A1").End(xlDown).Offset(0, 13)
    Set FavoritoLocal = Hoja20.Range("A1").End(xlDown).Offset(-1, 13)
    
    Select Case CuotaVisitor.Value
        Case Is < CuotaLocal.Value
            FavoritoVisitor.Value = "SI"
            FavoritoLocal.Value = "NO"
        Case Else
            FavoritoVisitor.Value = "NO"
            FavoritoLocal.Value = "SI"
    End Select

End Sub


Sub GanadorDescanso()

    Dim celdafinal, celda, PuntosVisitor, PuntosLocal, ResultadoVisitor, ResultadoLocal As Range
    Dim i As Integer

    Set celdafinal = Hoja20.Cells(Rows.Count, 1).End(xlUp)

    For i = 3 To celdafinal.Row Step 2

    Set PuntosVisitor = Hoja20.Cells(i, 1).Offset(0, 9)
    Set PuntosLocal = Hoja20.Cells(i, 1).Offset(-1, 9)
    Set ResultadoVisitor = Hoja20.Cells(i, 1).Offset(0, 10)
    Set ResultadoLocal = Hoja20.Cells(i, 1).Offset(-1, 10)


        If PuntosVisitor.Value > PuntosLocal.Value Then
            ResultadoLocal.Value = "No"
            ResultadoVisitor.Value = "SÍ"
        Else
            ResultadoLocal.Value = "Sí"
            ResultadoVisitor.Value = "No"
        End If

    Next i
End Sub
