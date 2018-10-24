Option Explicit
'Attribute VB_Name = "UtilesMacroUniversales"

'*******************************************************************
'*******************************************************************
'***UtilesMacroUniversales******************************************
'***Todas las macros son de autoria propia, las cuales pueden*******
'***ser copiadas mediante copyleft**********************************
'***Propietario: Omar André de la Sota******************************
'***correo: aozoro@gmail********************************************
'*******************************************************************

Public Function CrearLlave(ParamArray Args() As Variant) As String
    Dim Concat As String
    Dim i As Long
    
    Concat = Args(0)
    If UBound(Args) > 0 Then
        For i = LBound(Args) To UBound(Args) - 1
            Concat = Concat & "-" & Args(i + 1)
        Next i
    End If
    CrearLlave = Concat
End Function

Sub LimpiarCarpeta(ByVal NomCarpeta As String)
    Dim MyFolder As String, Myfile As String
    Dim Path As String
    
    Path = NomCarpeta
    Myfile = Dir(Path & "\*")
    Do While Myfile <> ""
        Myfile = Path & "\" & Myfile
        Kill Myfile
        Myfile = Dir
    Loop
End Sub

Sub CrearCarpeta(ByVal Ruta As String, ByVal NomCarpeta As String)
    Dim Path As String
    If Dir(Ruta, vbDirectory + vbHidden) <> "" Then
        If Dir(Ruta & "" & NomCarpeta, vbDirectory + vbHidden) = "" Then
            Path = Ruta & "" & NomCarpeta
            MkDir Path
        End If
    End If
End Sub

Sub Borrar_Adjcorreos(name_carpeta As String)
    Dim MyFolder As String, Myfile As String
    Dim Path As String
    
    MyFolder = Format(Date, "dd-mm-yyyy") & "_" & name_carpeta
    Path = "C:\" & "" & MyFolder
    Myfile = Dir(Path & "\*.msg")
    Do While Myfile <> ""
        Myfile = Path & "\" & Myfile
        Kill Myfile
        Myfile = Dir
    Loop
End Sub

Sub QuitarEspaciosDiferentes(ByVal Rng As Range)
    Dim celda As Range
    Dim largo As Integer
    Dim blank2 As String, blank3 As String, blank4 As String, blank5 As String, blank6 As String

    blank2 = " "
    blank3 = " "
    blank4 = vbCrLf
    blank5 = vbCr
    blank6 = vbLf

    For Each celda In Rng.Cells
        Do While Right(celda.Value, 1) = blank2 Or _
            Right(celda.Value, 1) = blank3 Or Right(celda.Value, 1) = blank4 _
            Or Right(celda.Value, 1) = blank5 Or Right(celda.Value, 1) = blank6
            
            largo = Len(celda.Value)
            celda.Value = Mid(celda.Value, 1, largo - 1)
        Loop
    Next
    
    For Each celda In Rng.Cells
        Do While Left(celda.Value, 1) = blank2 Or Left(celda.Value, 1) = blank3
            celda.Value = Mid(celda.Value, 2)
        Loop
    Next
End Sub

Sub QuitarApostrofes(ByVal Rng As Range)
    Dim celda As Range
    Dim Apostrofe1 As String, Apostrofe2 As String, _
        Apostrofe3 As String, Apostrofe4 As String, _
        Apostrofe5 As String

    Apostrofe1 = "'"
    Apostrofe2 = "´"
    Apostrofe3 = "`"
    Apostrofe4 = "¨"
    Apostrofe5 = """"

    For Each celda In Rng.Cells
        Do While Left(celda.Value, 1) = Apostrofe1 Or _
            Left(celda.Value, 1) = Apostrofe2 Or _
            Left(celda.Value, 1) = Apostrofe3 Or _
            Left(celda.Value, 1) = Apostrofe4 Or _
            Left(celda.Value, 1) = Apostrofe5
            
            celda.Value = Mid(celda.Value, 2)
        Loop
    Next celda
End Sub

Sub QuitarFilasEnBlanco(ByVal NumeroDeColumna As Integer, Optional ByVal Hoja As Worksheet = Nothing)
    If Hoja Is Nothing Then
        Hoja = ActiveSheet
    End If
    
    With Hoja
        If .Cells(1, NumeroDeColumna).End(xlDown).Address <> .Cells(.Rows.Count, NumeroDeColumna).End(xlUp).Address Then
            .Cells(1, NumeroDeColumna).EntireColumn.SpecialCells(xlCellTypeBlanks).EntireRow.Delete
        End If
    End With
End Sub

Sub QuitarFilasEnBlancoConjuntas(ByVal Hoja As Worksheet, ParamArray Args() As Variant)
    Dim i As Integer, j As Integer, N As Integer, T As Integer
    Dim VectorCantidadFilas() As Integer
    Dim sw As Boolean
    Dim limitehoja As Double
    
    With Hoja
        limitehoja = .Rows.Count
        T = UBound(Args)
        ReDim VectorCantidadFilas(T) As Integer
        For i = 0 To T
            VectorCantidadFilas(i) = .Cells(limitehoja, Args(i)).End(xlUp).Row
        Next i
        
        N = Application.WorksheetFunction.Max(VectorCantidadFilas)
            
        For i = 2 To N
            sw = False
            For j = 0 To T
                If .Cells(i, Args(j)).Value <> "" Then
                    sw = True
                    Exit For
                End If
            Next j
            
            If sw = False Then
                .Cells(i, 1).EntireRow.Delete
            End If
        Next i
    End With
End Sub

Sub DepuracionFilasBlanksDigitacion(Optional ByVal Hoja As Worksheet)
    With Hoja
        Do While .Range("E1").Value = ""
            .Range("E1").EntireRow.Delete
        Loop
    
        If .Range("E1").End(xlDown).Address <> .Range("E1048576").End(xlUp).Address Then
            Call QuitarFilasEnBlancoConjuntas(Hoja, 5)
        End If
    End With
End Sub

Public Sub AgregarHojas(ByVal cantidad As Integer, _
    Optional ByVal Nombre As String, _
    Optional ByVal Libro As Workbook)
    
    Dim i As Integer
    Dim j As Integer
    Dim Hoja As Worksheet
    
    If Libro Is Nothing Then Set Libro = ActiveWorkbook
    j = Libro.Worksheets.Count
    
    On Error Resume Next
    For i = 1 To cantidad
        With Libro
            Set Hoja = .Worksheets.Add(After:=.Worksheets(j))
            Hoja.Name = Nombre & i
            
            j = j + 1
        End With
    Next i
    
    If cantidad = 1 Then
        Hoja.Name = Left(Hoja.Name, Len(Hoja.Name) - 1)
    End If
    On Error GoTo 0
End Sub

Public Function SumaFilasEnLibro(ByVal Libro As Workbook, _
    Optional ByVal CriterioColumna As Integer = 1) As Double
    
    Dim Hoja As Worksheet
    Dim PrimeraFila As Double
    Dim Sumatoria As Double
    
    For Each Hoja In Libro.Worksheets
        Sumatoria = Sumatoria + EncontrarUltimaFila(Hoja, CriterioColumna) - 1
    Next Hoja
    
    SumaFilasEnLibro = Sumatoria
End Function

Public Function CocientePorExceso(ByVal numerador As Double, ByVal denominador As Double) As Integer
    If numerador Mod denominador = 0 Then
        CocientePorExceso = numerador / denominador
    Else
        CocientePorExceso = Fix(numerador / denominador) + 1
    End If
End Function

Sub CopiadoPegadoRangoToHoja(ByVal Rng As Range, _
    ByVal HojaPegado As Worksheet, _
    ByVal UltimaFila As Double, _
    Optional ByVal ColumnaPegado As Integer = 1)
    
    If Rng Is Nothing Then Exit Sub
    Rng.Copy
    With HojaPegado.Cells(UltimaFila + 1, ColumnaPegado)
        .PasteSpecial xlPasteValues
        .PasteSpecial xlPasteFormats
    End With
    Application.CutCopyMode = False
End Sub

Sub EliminarHojasPorIntervalo( _
    ByVal LimiteInferior As Integer, _
    ByVal LimiteSuperior As Integer, _
    Optional ByVal Libro As Workbook)
    
    Dim j As Integer
    
    If Libro Is Nothing Then Set Libro = ActiveWorkbook
    If LimiteInferior = 1 And LimiteSuperior = Libro.Worksheets.Count Then
        Call AgregarHojas(cantidad:=1, Libro:=Libro)
    End If
    
    j = LimiteInferior
    Do While LimiteInferior < LimiteSuperior
        Libro.Worksheets(j).Delete
        LimiteInferior = LimiteInferior + 1
    Loop
    Libro.Worksheets(j).Delete
End Sub

Sub ConsolidarHojas(Optional ByVal Libro As Workbook, _
    Optional ByVal NombreHojaConsolidado As String = "Consolidado", _
    Optional ByVal Encabezados As Boolean = True, _
    Optional ByVal LimiteFilasPorHoja As Double = 0, _
    Optional ByVal CriterioColumna As Integer = 1, _
    Optional ByVal TipoDigitacion As Boolean = False, _
    Optional ByVal BorrarHojas As Boolean = False, _
    Optional ByRef MensajeErrorConsolidado As String)
    
    Dim Cantidad_1 As Integer, Cantidad_2 As Integer, TotalHojas As Integer
    Dim N As Double, i As Double
    Dim fil As Double
    Dim col As Integer
    Dim Rng As Range
    Dim lim As Double
    Dim HojaPegado As Worksheet, Hoja As Worksheet
    Dim j As Integer
    Dim T As Double
    Dim fil_x As Double
    
    If Libro Is Nothing Then Set Libro = ActiveWorkbook
    If LimiteFilasPorHoja <= 0 Then LimiteFilasPorHoja = Libro.ActiveSheet.Rows.Count - 1
    Cantidad_1 = Libro.Worksheets.Count
    N = SumaFilasEnLibro(Libro)
    Cantidad_2 = CocientePorExceso(N, LimiteFilasPorHoja)
    
    If TipoDigitacion Then
        For Each Hoja In Libro.Sheets
            Call DepuracionFilasBlanks(10, Hoja)
        Next Hoja
    End If
    
    Call AgregarHojas(Cantidad_2, NombreHojaConsolidado, Libro)
    TotalHojas = Cantidad_1 + Cantidad_2
    
    '-----Falta:MensajeErrorEncabezados
    
    If Encabezados = True Then
        For i = 1 To Cantidad_2
            RangoAcotado(Libro.Worksheets(1), 1, 1, 1).Copy Destination:=Libro.Worksheets(Cantidad_1 + i).Cells(1, 1)
        Next i
    End If
    
    lim = 0
    j = Cantidad_1 + 1
    Set HojaPegado = Libro.Worksheets(j)
    For i = 1 To Cantidad_1
        Set Hoja = Libro.Worksheets(i)
        fil = EncontrarUltimaFila(Hoja, CriterioColumna)
        col = NumeroDeColumnas(Hoja)
        T = EncontrarUltimaFila(HojaPegado, CriterioColumna)
        
        lim = lim + fil - 1
        If lim < LimiteFilasPorHoja Then
            Set Rng = RangoAcotado(Hoja, 2, fil, 1, col)
            Call CopiadoPegadoRangoToHoja(Rng, HojaPegado, T)
        Else
            lim = lim - fil + 1
            fil_x = LimiteFilasPorHoja + 1 - lim
            Set Rng = RangoAcotado(Hoja, 2, fil_x, 1, col)
            Call CopiadoPegadoRangoToHoja(Rng, HojaPegado, T)
            j = j + 1
            Set HojaPegado = Libro.Worksheets(j)
            Set Rng = RangoAcotado(Hoja, fil_x + 1, fil, 1, col)
            Call CopiadoPegadoRangoToHoja(Rng, HojaPegado, 1)
            lim = fil
        End If
    Next i
    
    If BorrarHojas Then
        Call EliminarHojasPorIntervalo(1, Cantidad_1, Libro)
    End If
End Sub

Public Function VerificadorDeColumnas(ByVal NumeroDeColumnas As Integer) As Boolean
    Dim k As Integer
    Dim NumeroDeColumnasError As Integer

    VerificadorDeColumnas = False
    k = Cells(1, 16376).End(xlToLeft).Column
    If k = NumeroDeColumnas Then
        VerificadorDeColumnas = True
    Else
        If k > NumeroDeColumnas Then
            NumeroDeColumnasError = k - NumeroDeColumnas
            MsgBox "Hay más de " & NumeroDeColumnas & " Columnas. Por favor corregir la(s) " & NumeroDeColumnasError & " columnas de más manualmente de la hoja " & ActiveSheet.Name
        Else
            NumeroDeColumnasError = NumeroDeColumnas - k
            MsgBox "Hay menos " & NumeroDeColumnas & " Columnas. Por favor corregir la(s) " & NumeroDeColumnasError & " columnas faltantes manualmente de la hoja " & ActiveSheet.Name
        End If
    End If
End Function

Sub MontosQuitarComa(ByVal Rng As Range)
    Dim celda As Range

    For Each celda In Rng.Cells
        celda.Value = MontoDeComaAPunto(celda.Value)
    Next celda
End Sub

Public Function MontoDeComaAPunto(ByVal Importe As Variant) As Variant
    Dim n_1 As Integer
    Dim nchar As Integer

    If Left(Right(Importe, 3), 1) = "," Then
        n_1 = InStrRev(Importe, ".")
        Do Until n_1 = 0
            Importe = Mid(Importe, 1, n_1 - 1) & Mid(Importe, n_1 + 1)
            n_1 = InStrRev(Importe, ".")
        Loop

        nchar = Len(Importe)

        Importe = Left(Importe, nchar - 3) & "." & Right(Importe, 2)
    End If

    MontoDeComaAPunto = Importe
End Function

Public Function BuscadorDeEncabezado(ByVal Hoja As Worksheet, ParamArray NombreDeEncabezado() As Variant) As Integer
    Dim i As Integer, j As Integer, k As Integer, T As Integer
    Dim sw As Boolean

    sw = False
    T = UBound(NombreDeEncabezado)

    With Hoja
        k = .Cells(1, .Columns.Count).End(xlToLeft).Column
        For i = 1 To k
            For j = 0 To T
                On Error Resume Next
                If UCase(.Cells(1, i).Value) = UCase(NombreDeEncabezado(j)) Then
                    sw = True
                    If Err.Number <> 0 Then sw = False
                    Exit For
                End If

                On Error GoTo 0
            Next j

            If sw = True Then
                Exit For
            End If
        Next i
    End With

    If sw = True Then
        BuscadorDeEncabezado = i
    Else
        BuscadorDeEncabezado = 0
    End If
End Function

Public Function ContarSinRepetirOrdenando(ByVal NumeroDeColumna As Integer, Optional ByVal Hoja As Worksheet) As Double
    Dim i As Integer, j As Integer
    Dim N As Integer
    
    If Hoja Is Nothing Then Set Hoja = ActiveSheet
    
    N = EncontrarUltimaFila(Hoja, NumeroDeColumna)
    
    For i = 2 To N
        If Cells(i - 1, NumeroDeColumna) <> Cells(i, NumeroDeColumna).Value Then j = j + 1
    Next i
    
    ContarSinRepetirOrdenando = j
End Function

Public Function ElementoEnVector(ByVal Elemento, ByVal Vector) As Boolean
    Dim Matriz()
    Dim T As Integer
    Dim i As Integer
    Dim x
    
    T = UBound(Vector)
    ReDim Matriz(T, 0)
    For i = 0 To T
        Matriz(i, 0) = Vector(i)
    Next i
    
    On Error Resume Next
    x = Application.WorksheetFunction.VLookup(Elemento, Matriz, 1, 0)
    If Err.Number <> 0 Then
        On Error GoTo 0
        On Error Resume Next
        x = Application.WorksheetFunction.VLookup(Elemento * 1, Matriz, 1, 0)
        If Err.Number <> 0 Then
            ElementoEnVector = False
        Else
            ElementoEnVector = True
        End If
    Else
        ElementoEnVector = True
    End If
    On Error GoTo 0
End Function

Public Function UbicacionDelElementoEnVector(ByVal Elemento As Variant, ByVal Lista As Variant) As Integer
    Dim i As Integer, T As Integer
    Dim Numeracion() As Integer
    Dim Arreglo() As Variant
    
    T = UBound(Lista)
    ReDim Arreglo(T, 1) As Variant
    
    For i = 0 To T
        Arreglo(i, 0) = Lista(i)
        Arreglo(i, 1) = i
    Next i
    
    UbicacionDelElementoEnVector = -1
    On Error Resume Next
    UbicacionDelElementoEnVector = Application.WorksheetFunction.VLookup(Elemento, Arreglo, 2, 0)
    On Error GoTo 0
End Function

Sub BuscarEnLaColumna(ByVal HojaDeTrabajo As Worksheet, ByVal ColumnaABuscar As Integer, _
    ParamArray Criterio() As Variant)
    'Crea una columna dando valores 1=Encontrado, 0=No Encontrado
    
    Dim ColumnaCoincidencia As Integer
    Dim i As Double, N As Double
    
    ColumnaCoincidencia = ColumnaABuscar + 1
    
    With HojaDeTrabajo
        .Cells(1, ColumnaCoincidencia).EntireColumn.Insert Shift:=xlToRight
        .Cells(1, ColumnaCoincidencia).Value = "Coincidencia"
        N = .Cells(.Rows.Count, ColumnaABuscar).End(xlUp).Row
        For i = 2 To N
            If UbicacionDelElementoEnVector(.Cells(i, ColumnaABuscar).Value, Criterio) = -1 Then
                .Cells(i, ColumnaCoincidencia).Value = 0
            Else
                .Cells(i, ColumnaCoincidencia).Value = 1
            End If
        Next i
    End With
End Sub

Sub CopiarColumnasDeUnaHojaAOtra(ByVal HojaOutput As Worksheet, ByVal HojaInput As Worksheet, _
    ByVal ColumnaCondicion As Integer, ByVal Condicion As Boolean, ParamArray NumeroDeColumna() As Variant)
    
    Dim LimiteSuperior As Integer
    Dim CondicionNumero As Integer
    Dim i As Double, N As Double
    Dim j As Integer
    Dim k As Integer
    Dim Rng As Range
    
    If Condicion = True Then
        CondicionNumero = 1
    Else
        CondicionNumero = 0
    End If
        
    LimiteSuperior = UBound(NumeroDeColumna)
    
    N = HojaInput.Cells(Rows.Count, ColumnaCondicion).End(xlUp).Row
    For i = 2 To N
        If HojaInput.Cells(i, ColumnaCondicion).Value = CondicionNumero Then
            k = 0
            Set Rng = HojaOutput.Cells(Rows.Count, 1).End(xlUp).Offset(1, 0)
            For j = 0 To LimiteSuperior
                HojaInput.Cells(i, NumeroDeColumna(j)).Copy Destination:=Rng.Offset(0, k)
                k = k + 1
            Next j
        End If
    Next i
End Sub

Sub EncenderAcelerarMacro()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
End Sub

Sub ApagarAcelerarMacro()
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
End Sub

Public Function ElegirCarpeta()
    Dim diaFolder As FileDialog

    Set diaFolder = Application.FileDialog(msoFileDialogFolderPicker)
    diaFolder.AllowMultiSelect = False
    diaFolder.Show
    ElegirCarpeta = diaFolder.SelectedItems(1)
End Function

Public Function NuevoLibro(ParamArray NombreDeEncabezado() As Variant) As Workbook
    'Crea un nuevo libro con los encabezados imputados
    
    Dim UltimoIndice As Integer
    Dim i As Integer
    
    UltimoIndice = UBound(NombreDeEncabezado)
    Set NuevoLibro = Workbooks.Add
    
    For i = 0 To UltimoIndice
        NuevoLibro.ActiveSheet.Cells(1, i + 1).Value = NombreDeEncabezado(i)
    Next i
End Function

Sub QuitarFilasEnBlancoC(ByVal NumeroDeColuma As Integer, _
    ByVal UltimaFilaABuscar As Double, _
    Optional ByVal Hoja As Worksheet = Nothing)
        
    Dim FilaBuscada As Double
    Dim sw As Boolean
    Dim Inicial As Double, Final As Double
    Dim sw2 As Boolean
    
    sw2 = False
    If Hoja Is Nothing Then
        Set Hoja = ActiveSheet
    End If
    
    With Hoja
        If .Cells(UltimaFilaABuscar, NumeroDeColuma) = "" Then
            sw2 = True
        End If
        
        FilaBuscada = 1
        Do While FilaBuscada <= UltimaFilaABuscar
            sw = False
            Inicial = FilaBuscada
            Final = FilaBuscada
            Do While FilaBuscada <= UltimaFilaABuscar
                If .Cells(FilaBuscada, NumeroDeColuma).Value = "" Then
                    FilaBuscada = FilaBuscada + 1
                Else
                    sw = True
                    FilaBuscada = FilaBuscada + 1
                    Exit Do
                End If
            Loop
            
            Final = FilaBuscada - 1
            
            If sw = True And Inicial <> Final Then
                .Range(.Cells(Inicial, NumeroDeColuma), .Cells(Final - 1, NumeroDeColuma).EntireRow).Delete
                UltimaFilaABuscar = UltimaFilaABuscar - Final + Inicial
                FilaBuscada = Inicial + 1
            End If
            
            If sw2 = True And FilaBuscada > UltimaFilaABuscar Then
                .Range(.Cells(Inicial, NumeroDeColuma), .Cells(Final, NumeroDeColuma).EntireRow).Delete
            End If
        Loop
    End With
End Sub

Sub SepararHojas(ByVal Folder As String, ByVal File As String)
    Dim Hoja As Worksheet
    Dim Libro As Workbook
    Dim LibroNuevo As Workbook
    Dim NombreHoja As String
    
    Workbooks.Open FileName:=Folder & "\" & File
    Set Libro = ActiveWorkbook
     
    If Libro.Worksheets.Count <> 1 Then
        For Each Hoja In Libro.Worksheets
            NombreHoja = Hoja.Name
            Hoja.Copy
            Set LibroNuevo = ActiveWorkbook
            LibroNuevo.SaveAs FileName:=Folder & "\" & CrearLlave(NombreHoja, File)
            LibroNuevo.Close SaveChanges:=False
        Next Hoja
        Libro.Close SaveChanges:=False
        Kill Folder & "\" & File
    Else
        Libro.Close SaveChanges:=False
    End If
End Sub

Sub DepuracionFilasBlanks(ByVal N As Integer, Optional ByVal Hoja As Worksheet = Nothing)
    Dim j As Integer
    
    If Hoja Is Nothing Then
        Set Hoja = ActiveSheet
    End If
    
    With Hoja
        Do Until j = N
            If .Range("E1").Value = "" Then
                .Range("E1").EntireRow.Delete
            End If
            j = j + 1
        Loop
        
        If .Range("E1").End(xlDown).Address <> .Cells(.Rows.Count, 5).End(xlUp).Address Then
            Call QuitarFilasEnBlancoConjuntas(Hoja, 5)
        End If
    End With
End Sub

Sub ColocarEncabezados(ByVal Hoja As Worksheet, ParamArray Etiquetas() As Variant)
    Dim i As Integer, j As Integer
    Dim L As Integer, U As Integer
    L = LBound(Etiquetas)
    U = UBound(Etiquetas)
    
    For i = L To U
        j = j + 1
        With Hoja
            .Cells(1, j).Value = Etiquetas(i)
        End With
    Next i
End Sub

Sub MontosSinDecimales(Optional ByVal Hoja As Worksheet = Nothing, Optional ByVal ColumnaImporte As Integer = 0)
    Dim fechas() As Double
    Dim Importes() As Double
    Dim ubicacion As Integer
    Dim j As Integer, T As Integer
    Dim i As Double, N As Double
    Dim i_FechaProceso As Integer, i_Importe As Integer, i_Fecha As Integer
    Dim fechanumero As Double
    Dim FechasError() As Date
    
    If Hoja Is Nothing Then
        Set Hoja = ActiveSheet
    End If
    
    Call DepuracionFilasBlanks(10)
    
    i_FechaProceso = BuscadorDeEncabezado(Hoja, "FECHA PROCESO")
    
    If ColumnaImporte = 0 Then
        i_Importe = BuscadorDeEncabezado(Hoja, "IMPORTE")
    Else
        i_Importe = ColumnaImporte
    End If
    
    With Hoja
        N = .Cells(.Rows.Count, i_FechaProceso).End(xlUp).Row
 
        For i = 2 To N
            ReDim Preserve fechas(j) As Double
            ReDim Preserve Importes(j) As Double
            fechanumero = .Cells(i, i_FechaProceso).Value * 1
            If ContarSiVector(fechas, fechanumero) = 0 Then
                fechas(j) = fechanumero
                On Error Resume Next
                Importes(j) = .Cells(i, i_Importe).Value - Fix(.Cells(i, i_Importe).Value)
                On Error GoTo 0
                j = j + 1
            Else
                ubicacion = UbicacionDelElementoEnVector(fechanumero, fechas)
                On Error Resume Next
                Importes(ubicacion) = Importes(ubicacion) + (.Cells(i, i_Importe).Value - Fix(.Cells(i, i_Importe).Value))
                On Error GoTo 0
            End If
        Next i
    End With
    
    T = j - 1
    j = 0
    For i = 0 To T
        If Importes(i) = 0 Then
            ReDim FechasError(j) As Date
            FechasError(j) = fechas(i)
            j = j + 1
        End If
    Next i
    On Error Resume Next
    T = UBound(FechasError)
    If Err.Number = 0 Then
        With Hoja
            For i = 2 To N
                For j = 0 To T
                    If .Cells(i, i_FechaProceso).Value = FechasError(j) Then
                        On Error Resume Next
                        .Cells(i, i_Importe).Value = .Cells(i, i_Importe).Value / 100
                        On Error GoTo 0
                    End If
                Next j
            Next i
        End With
    End If
    On Error GoTo 0
End Sub

Public Function ContarSiVector(Vector() As Double, ByVal Criteria As Double) As Integer
    Dim i&
    Dim U As Integer, L As Integer
    
    U = UBound(Vector)
    L = LBound(Vector)
    
    For i = L To U
        ContarSiVector = ContarSiVector - (Vector(i) = Criteria)
    Next i
End Function

Sub CopiadoPegadoColumnaRangeToRange(ByVal HojaInput As Worksheet, ByVal ColumnaInput As Integer, _
    ByVal UltimaFilaInput As Double, ByVal HojaOutput As Worksheet, _
    ByVal ColumnaOutput As Integer, ByVal FilaInicioOutput As Double, _
    Optional ByVal TipoDeCopiado As Integer = 0)
    
    Dim RangoCopy As Range
    Dim RangoPaste As Range
    Dim ModoCopiado As String
    
    If ColumnaInput = 0 Then
        Exit Sub
    End If

    With HojaInput
        Set RangoCopy = .Range(.Cells(2, ColumnaInput), .Cells(UltimaFilaInput, ColumnaInput))
    End With
    Set RangoPaste = HojaOutput.Cells(FilaInicioOutput, ColumnaOutput)
    RangoCopy.Copy
    
    Select Case TipoDeCopiado
        Case 0
            RangoPaste.PasteSpecial Paste:=xlPasteValues, operation:=xlNone, Skipblanks:=False
        Case 1
            RangoPaste.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, operation:=xlNone, Skipblanks:=False
        Case Else
            RangoPaste.PasteSpecial Paste:=xlPasteValues, operation:=xlNone, Skipblanks:=False
    End Select
End Sub

Sub RutinaImporte(Optional ByVal Hoja As Worksheet = Nothing, _
    Optional ByVal ComaXPunto As Boolean = False, _
    Optional ByVal SinPunto As Boolean = False)
    
    Dim i_Importe As Integer
    Dim Rng As Range
    
    If Hoja Is Nothing Then
        Set Hoja = ActiveSheet
    End If
    
    i_Importe = BuscadorDeEncabezado(Hoja, "IMPORTE")
    
    With Hoja
        Set Rng = .Range(.Cells(2, i_Importe), .Cells(.Rows.Count, i_Importe).End(xlUp))
    End With
    
    If ComaXPunto = True Then
        Call MontosQuitarComa(Rng)
    End If
End Sub

Public Function VerificarLibroAbierto(ByVal NombreDeLibro As String) As String
    Dim Libro As Workbook

    VerificarLibroAbierto = ""

    For Each Libro In Workbooks
        If InStr(Libro.Name, NombreDeLibro) > 0 Then
            VerificarLibroAbierto = Libro.Name
            Exit For
        End If
    Next
End Function

Sub EliminadorDeColumnaConEncabezado(ByVal Hoja As Worksheet, _
    ParamArray NombreDeEncabezado() As Variant)
    
    Dim i As Integer, j As Integer, k As Integer, T As Integer
    Dim sw As Boolean
    Dim i_Criterio As Integer
    
    sw = False
    T = UBound(NombreDeEncabezado)
    
    With Hoja
        k = .Cells(1, .Columns.Count).End(xlToLeft).Column
        For i = 1 To k
            For j = 0 To T
                On Error Resume Next
                If .Cells(1, i).Value = NombreDeEncabezado(j) Then
                    sw = True
                    If Err.Number <> 0 Then sw = False
                    Exit For
                End If
                
                On Error GoTo 0
            Next j
            
            If sw = True Then
                Exit For
            End If
        Next i
    End With
    
    If sw = True Then
        i_Criterio = i
    Else
       i_Criterio = 0
    End If
    
    With Hoja
        If i_Criterio <> 0 Then
            .Cells(1, i_Criterio).EntireColumn.Delete
        End If
    End With
End Sub

Public Function BuscarFechaEnColumna(ByVal Columna As Integer, _
    Optional ByVal Hoja As Worksheet = Nothing) As Date

    Dim i As Double
    Dim Fecha As Date
    
    If Hoja Is Nothing Then
        Set Hoja = ActiveSheet
    End If
    
    With Hoja
        i = 1
        Do While Fecha = 0 And i < 100
            i = i + 1
            On Error Resume Next
            BuscarFechaEnColumna = .Cells(i, Columna).Value
            On Error GoTo 0
        Loop
    End With
End Function

Public Function ModaDeColumnas(Optional ByVal Libro As Workbook = Nothing) As Integer
    Dim Hoja As Worksheet
    Dim ContadorDeColumnas() As Integer
    Dim i As Integer, j As Integer
    
    If Libro Is Nothing Then
        Set Libro = ActiveWorkbook
    End If
    
    j = Libro.Worksheets.Count - 1
        
    If j = 0 Then
        With Libro.Worksheets(i + 1)
            ModaDeColumnas = .Cells(1, .Columns.Count).End(xlToLeft).Column
        End With
    Else
        ReDim ContadorDeColumnas(j) As Integer
        For i = 0 To j
            With Libro.Worksheets(i + 1)
                ContadorDeColumnas(i) = .Cells(1, .Columns.Count).End(xlToLeft).Column
            End With
        Next i
        
        ModaDeColumnas = Application.WorksheetFunction.Mode_Sngl(ContadorDeColumnas)
    End If
End Function

Public Function BuscadorDeEncabezadoAproximado( _
    ByVal Criteria As Variant, _
    Optional ByVal Hoja As Worksheet)
        
    Dim i As Integer, T As Integer
    Dim sw As Boolean
    
    sw = False
    
    With Hoja
        T = .Cells(1, .Columns.Count).End(xlToLeft).Column
        For i = 1 To T
            If InStr(1, .Cells(1, i).Value, Criteria, vbTextCompare) > 0 Then
                sw = True
                Exit For
            End If
        Next i
        If sw = True Then
            BuscadorDeEncabezadoAproximado = i
        Else
            BuscadorDeEncabezadoAproximado = 0
        End If
    End With
End Function

Public Function NumeroDeColumnas(Optional ByVal Hoja As Worksheet, _
    Optional ByVal CriterioFila As Double = 1) As Integer
    
    If Hoja Is Nothing Then Set Hoja = ActiveSheet

    With Hoja
        NumeroDeColumnas = .Cells(CriterioFila, .Columns.Count).End(xlToLeft).Column
    End With
End Function

Public Function EncontrarUltimaFila(Optional ByVal Hoja As Worksheet, _
    Optional ByVal CriterioColumna As Integer = 1) As Double
    
    If Hoja Is Nothing Then Set Hoja = ActiveSheet
    
    With Hoja
        If .Cells(.Rows.Count, CriterioColumna).Value <> "" Then
            EncontrarUltimaFila = .Rows.Count
        Else
            EncontrarUltimaFila = .Cells(.Rows.Count, CriterioColumna).End(xlUp).Row
        End If
    End With
End Function

Public Function RangoTotal(Optional ByVal Hoja As Worksheet = Nothing, _
    Optional ByVal CriterioColumna As Integer = 1, _
    Optional ByVal CriterioFila As Double = 1) As Range
    
    Dim N As Double
    Dim k As Integer
    
    If Hoja Is Nothing Then Set Hoja = ActiveSheet
    
    N = EncontrarUltimaFila(Hoja, CriterioColumna)
    k = NumeroDeColumnas(Hoja, CriterioFila)
    
    With Hoja
        Set RangoTotal = .Range(.Cells(1, 1), .Cells(N, k))
    End With
End Function
Sub FiltrarOrdenarHojaPorColumna(ByVal Columna As Integer, _
    Optional ByVal Hoja As Worksheet = Nothing, _
    Optional ByVal UltimaFila As Double, _
    Optional ByVal Ascending As Boolean = False)
    
    If Hoja Is Nothing Then
        Set Hoja = ActiveSheet
    End If
    
    If UltimaFila = 0 Then UltimaFila = EncontrarUltimaFila(Hoja, Columna)
    
    With Hoja
        If .AutoFilterMode = True Then
            .Range("A1").AutoFilter
        End If
        
        RangoTotal(Hoja, 1, 1).AutoFilter
        
        .AutoFilter.Sort.SortFields.Clear
        If Ascending = True Then
            .AutoFilter.Sort.SortFields.Add Key:=RangoTotalColumna(Columna, 1, UltimaFila, Hoja), Order:=xlAscending
        Else
            .AutoFilter.Sort.SortFields.Add Key:=RangoTotalColumna(Columna, 1, UltimaFila, Hoja), Order:=xlDescending
        End If
        .AutoFilter.Sort.Apply
    End With
End Sub

Public Function RangoTotalColumna(ByVal Columna As Integer, _
    Optional ByVal PrimeraFila As Double = 1, _
    Optional ByVal UltimaFila As Double, _
    Optional ByVal Hoja As Worksheet) As Range
    
    If Hoja Is Nothing Then Set Hoja = ActiveSheet
    
    If UltimaFila = 0 Then UltimaFila = EncontrarUltimaFila(Hoja, Columna)

    With Hoja
        Set RangoTotalColumna = .Range(.Cells(PrimeraFila, Columna), .Cells(UltimaFila, Columna))
    End With
End Function

Sub FormatoFecha(ByVal Hoja As Worksheet, ParamArray Columnas() As Variant)
    Dim U As Integer, L As Integer
    Dim i As Integer
    
    L = LBound(Columnas)
    U = UBound(Columnas)
    For i = L To U
        With Hoja
            .Cells(1, Columnas(i)).EntireColumn.NumberFormat = "dd/mm/yyyy"
        End With
    Next i
End Sub

Public Function RangoAcotado(Optional ByVal Hoja As Worksheet, _
    Optional ByVal FilaInicial As Double = 1, _
    Optional ByVal FilaFinal As Double, _
    Optional ByVal ColumnaInicial As Integer = 1, _
    Optional ByVal ColumnaFinal As Integer) As Range
    
    Dim i As Double
    
    If Hoja Is Nothing Then Set Hoja = ActiveSheet
    
    If ColumnaFinal = 0 Then ColumnaFinal = NumeroDeColumnas(Hoja, FilaInicial)
    
    If ColumnaFinal - ColumnaInicial < 0 Then
        Exit Function
    End If
    
    If FilaFinal = 0 Then
        With Hoja
            For i = ColumnaInicial To ColumnaFinal
                FilaFinal = Application.WorksheetFunction.Max(FilaFinal, .Cells(.Rows.Count, i).End(xlUp).Row)
            Next i
        End With
    End If
    
    If FilaFinal - FilaInicial < 0 Then
        Exit Function
    End If
        
    With Hoja
        Set RangoAcotado = .Range(.Cells(FilaInicial, ColumnaInicial), .Cells(FilaFinal, ColumnaFinal))
    End With

End Function

Public Function SumaVectorInteger(ByRef Vector() As Integer) As Integer
    Dim i As Double, o As Double, T As Double
    Dim SumaVector As Double
    
    o = LBound(Vector)
    T = UBound(Vector)
    For i = o To T
        SumaVector = SumaVector + Vector(i)
    Next i
    SumaVectorInteger = SumaVector
End Function

Sub BorrarFilasDistintasA(ByVal Hoja As Worksheet, _
    ByVal NumeroDeColuma As Integer, _
    ByVal UltimaFilaABuscar As Double, _
    ParamArray Criteria() As Variant)

    Dim FilaBuscada As Double
    Dim sw As Boolean
    Dim Inicial As Double, Final As Double
    Dim sw2 As Boolean

    With Hoja
        If UbicacionDelElementoEnVector(.Cells(UltimaFilaABuscar, NumeroDeColuma), Criteria) = -1 Then
            sw2 = True
        End If

        FilaBuscada = 1
        Do While FilaBuscada <= UltimaFilaABuscar
            sw = False
            Inicial = FilaBuscada
            Final = FilaBuscada
            Do While FilaBuscada <= UltimaFilaABuscar
                If UbicacionDelElementoEnVector(.Cells(FilaBuscada, NumeroDeColuma).Value, Criteria) = -1 Then
                    FilaBuscada = FilaBuscada + 1
                Else
                    sw = True
                    FilaBuscada = FilaBuscada + 1
                    Exit Do
                End If
            Loop
            Final = FilaBuscada - 1

            If sw = True And Inicial <> Final Then
                .Range(.Cells(Inicial, NumeroDeColuma), .Cells(Final - 1, NumeroDeColuma).EntireRow).Delete
                UltimaFilaABuscar = UltimaFilaABuscar - Final + Inicial
                FilaBuscada = Inicial + 1
            End If

            If sw2 = True And FilaBuscada > UltimaFilaABuscar Then
                .Range(.Cells(Inicial, NumeroDeColuma), .Cells(Final, NumeroDeColuma).EntireRow).Delete
            End If
        Loop
    End With
End Sub

Sub RutinaSeleccionadas(ByVal ColumnaInput As Integer, _
    ByVal HojaInput As Worksheet, _
    ByVal ColumnaOutput As Integer, _
    ByVal HojaOutput As Worksheet, _
    ByVal ColumnaBase As Integer, _
    ByVal HojaBase As Worksheet, _
    Optional ByVal ChancarValores As Boolean = False, _
    Optional ByVal ValorSi As Variant = "Seleccionada", _
    Optional ByVal ValorNo As Variant = "")
    
    Dim N_Input As Double
    Dim N_Base As Double
    Dim i As Double
    Dim Rng As Range
    Dim Var1 As Variant
    
    N_Input = EncontrarUltimaFila(HojaInput, ColumnaInput)
    N_Base = EncontrarUltimaFila(HojaBase, ColumnaBase)
    Set Rng = RangoTotalColumna(ColumnaBase, 2, N_Base, HojaBase)
    
    For i = 2 To N_Input
        On Error Resume Next
        Var1 = HojaInput.Cells(i, ColumnaInput).Value
        Var1 = Application.WorksheetFunction.VLookup(Var1, Rng, 1, 0)
        If Err.Number = 0 Then
            If ChancarValores = True Then
                HojaOutput.Cells(i, ColumnaOutput).Value = ValorSi
            Else
                If HojaOutput.Cells(i, ColumnaOutput).Value = "" Then HojaOutput.Cells(i, ColumnaOutput).Value = ValorSi
            End If
        Else
            HojaOutput.Cells(i, ColumnaOutput).Value = ValorNo
        End If
        On Error GoTo 0
    Next i
End Sub

Sub ContarSegunCriterio(ByVal ColumnaInput As Integer, _
    ByVal HojaInput As Worksheet, _
    ByVal ColumnaBase As Integer, _
    ByVal HojaBase As Worksheet, _
    ByVal ColumnaOutput As Integer, _
    ByVal HojaOutput As Worksheet)
    
    Dim N_Input As Double
    Dim N_Base As Double
    Dim Rng As Range
    Dim i As Double
    
    N_Input = EncontrarUltimaFila(HojaInput, ColumnaInput)
    N_Base = EncontrarUltimaFila(HojaBase, ColumnaBase)
    Set Rng = RangoTotalColumna(ColumnaBase, 2, N_Base, HojaBase)
    
    For i = 2 To N_Input
        HojaOutput.Cells(i, ColumnaOutput).Value = Application.WorksheetFunction.CountIf(Rng, HojaInput.Cells(i, ColumnaInput).Value)
    Next i
End Sub

Sub CambiarValoresCeldas(ByVal Columna As Integer, _
    ByVal MensajeInicial As Variant, _
    ByVal MensajeFinal As Variant, _
    Optional ByVal Inicio As Double = 2, _
    Optional ByVal Final As Double, _
    Optional ByVal Hoja As Worksheet)
    
    Dim i As Double
    
    If Hoja Is Nothing Then Set Hoja = ActiveSheet
    If Final = 0 Then Final = EncontrarUltimaFila(Hoja, Columna)
    
    For i = Inicio To Final
        If Cells(i, Columna).Value = MensajeInicial Then Cells(i, Columna).Value = MensajeFinal
    Next i
End Sub

Sub FormatoTablaQuince(Optional ByVal Hoja As Worksheet, _
    Optional ByVal Centrado As Boolean = False, _
    Optional ByVal Ajustado As Boolean = False)
    
    Dim Rng As Range
    Dim Nombre As String
    
    Nombre = Hoja.Name
    Set Rng = RangoCurrentRegion(Hoja)
    Rng.Columns.AutoFit
    
    With Hoja
        .ListObjects.Add(xlSrcRange, Rng, , xlYes).Name = Nombre
        '.ListObjects(Nombre).TableStyle = "TableStyleMedium15" 'con lineas de columnas
        .ListObjects(Nombre).TableStyle = "TableStyleMedium1" 'sin lineas de columnas
    End With
    
    If Centrado Then Rng.HorizontalAlignment = xlCenter
    If Ajustado Then Rng.EntireColumn.AutoFit
End Sub

Public Function RangoCurrentRegion(Optional ByVal Hoja As Worksheet, _
    Optional ByVal Fila As Double = 1, _
    Optional ByVal Columna As Integer = 1) As Range
    
    If Hoja Is Nothing Then Set Hoja = ActiveSheet
    Set RangoCurrentRegion = Hoja.Cells(1, 1).CurrentRegion
End Function

Sub BorrarFilasIgualesA(ByVal Hoja As Worksheet, _
    ByVal Columna As Integer, _
    ByVal PrimeraFila As Double, _
    ByVal UltimaFila As Double, _
    ParamArray Criteria() As Variant)
    
    Dim sw As Boolean, sw2 As Boolean
    Dim FilaBuscada As Double
    Dim Inicial As Double, Final As Double
    
    With Hoja
        If UbicacionDelElementoEnVector(.Cells(UltimaFila, Columna), Criteria) <> -1 Then
            sw2 = True 'el ultimo elemento es un criterio
        End If
    
        FilaBuscada = PrimeraFila
        
        Do While FilaBuscada <= UltimaFila
            sw = False
            Inicial = FilaBuscada
            Final = FilaBuscada
            Do While FilaBuscada <= UltimaFila
                If UbicacionDelElementoEnVector(.Cells(FilaBuscada, Columna).Value, Criteria) <> -1 Then
                    FilaBuscada = FilaBuscada + 1
                Else
                    sw = True
                    FilaBuscada = FilaBuscada + 1
                    Exit Do
                End If
            Loop
            Final = FilaBuscada - 1

            If sw = True And Inicial <> Final Then
                .Range(.Cells(Inicial, Columna), .Cells(Final - 1, Columna).EntireRow).Delete
                UltimaFila = UltimaFila - Final + Inicial
                FilaBuscada = Inicial + 1 '''
            End If

            If sw2 = True And FilaBuscada > UltimaFila Then
                .Range(.Cells(Inicial, Columna), .Cells(Final, Columna).EntireRow).Delete
            End If
        Loop
    End With
End Sub

Sub CrearColumnaLlave(ByVal Hoja As Worksheet, _
    ByVal NombreDeEncabezado As String, _
    ByVal UbicacionDeEncabezado As Integer, _
    ByVal Numerico As Boolean, _
    ByVal MaximoDeFilas As Boolean, _
    ParamArray Columnas() As Variant)
    
    Dim i As Double, N As Double
    Dim j As Integer, k As Integer, T As Integer
    Dim llave As String
    
    T = UBound(Columnas)
    
    If MaximoDeFilas = True Then
        For j = 0 To k
            N = Application.WorksheetFunction.Max(N, EncontrarUltimaFila(Hoja, Columnas(j)))
        Next j
    Else
        N = EncontrarUltimaFila(Hoja, Columnas(j))
        For j = 1 To k
            N = Application.WorksheetFunction.Min(N, EncontrarUltimaFila(Hoja, Columnas(j)))
        Next j
    End If
    
    k = NumeroDeColumnas(Hoja)
    k = k + 1
    With Hoja
        .Cells(1, k).Value = NombreDeEncabezado
        If Numerico = True Then
            For i = 2 To N
                For j = 0 To T
                    On Error Resume Next
                    llave = CrearLlave(llave, .Cells(i, Columnas(j)).Value * 1)
                    If Err.Number <> 0 Then llave = CrearLlave(llave, .Cells(i, Columnas(j)).Value)
                    On Error GoTo 0
                Next j
                 .Cells(i, k).Value = Mid(llave, 2)
                 llave = ""
            Next i
        Else
            For i = 2 To N
                For j = 0 To T
                    llave = CrearLlave(llave, .Cells(i, Columnas(j)).Value)
                Next j
                .Cells(i, k).Value = llave
                llave = ""
            Next i
        End If
    End With
    
    If UbicacionDeEncabezado > 0 Then
        With Hoja
            .Cells(1, k).EntireColumn.Cut
            .Cells(1, UbicacionDeEncabezado).Insert xlToRight
        End With
    End If
End Sub

Sub CrearColumnaConteo(ByVal ColumnaInput As Integer, _
    ByVal HojaInput As Worksheet, _
    ByVal ColumnaOutput As Integer, _
    ByVal HojaOutput As Worksheet, _
    ByVal NombreEncabezado As String)
    
    Dim Rng As Range
    Dim i As Double
    Dim N As Double
    Dim k As Integer
    
    Set Rng = RangoTotalColumna(ColumnaInput, _
        PrimeraFila:=2, _
        Hoja:=HojaInput)
    
    N = EncontrarUltimaFila(HojaInput, ColumnaInput)

    HojaOutput.Cells(1, ColumnaOutput).Value = NombreEncabezado
    For i = 2 To N
        HojaOutput.Cells(i, ColumnaOutput).Value = Application.WorksheetFunction.CountIf(Rng, HojaInput.Cells(i, ColumnaInput).Value)
    Next i
End Sub

Public Function CrearFiltroQuery(ByVal ColumnNumber As Integer, ByVal Filtros As Variant)
    Dim Column As String
    Dim i As Integer
    Dim T As Integer
    Dim o As Integer
    
    o = LBound(Filtros) + 1
    T = UBound(Filtros)
    Column = "[Column" & ColumnNumber & "] = "
    
    CrearFiltroQuery = Column & Filtros(LBound(Filtros))
    
    For i = o To T
        CrearFiltroQuery = CrearFiltroQuery & " or " & Column & Filtros(i)
    Next i
End Function

Public Function TextToMatrix(ByVal text As Variant, Optional ByVal Duplicados As Boolean = False) As Variant
    Dim Arreglo() As Variant
    Dim N As Integer
    Dim j As Integer
    
    If InStr(text, ",") = 0 Then
        If text <> "" Then
            ReDim Arreglo(0)
            On Error Resume Next
            Arreglo(0) = Trim(text)
            If Err.Number <> 0 Then Arreglo(0) = 0
            On Error GoTo 0
        Else
            ReDim Arreglo(0)
            Arreglo(0) = 0
        End If
    Else
        j = 0
        On Error Resume Next
        N = InStr(text, ",")
        ReDim Preserve Arreglo(j)
        Arreglo(j) = Trim(Left(text, N - 1))
        text = Trim(Mid(text, N + 1))
        j = j + 1
        Do While InStr(text, ",") <> 0
            N = InStr(text, ",")
            If Duplicados Or ElementoEnVector(Trim(Left(text, N - 1)), Arreglo) = False Then
                ReDim Preserve Arreglo(j)
                Arreglo(j) = Trim(Left(text, N - 1))
                j = j + 1
            End If
            text = Trim(Mid(text, N + 1))
        Loop
        
        If Duplicados Or ElementoEnVector(Trim(text), Arreglo) = False Then
            ReDim Preserve Arreglo(j)
            Arreglo(j) = Trim(text)
        End If
        
        If Err.Number <> 0 Then
            Exit Function
        End If

        On Error GoTo 0
    End If
    
    TextToMatrix = Arreglo
End Function

Public Function MatrixToText(ByVal Matrix As Variant, _
    Optional ByVal SaltoDeLinea As Boolean = False, _
    Optional ByVal sepLeft As String = "", _
    Optional ByVal sepRight As String = "") As String
    
    Dim txt As String
    Dim o As Integer, T As Integer, i As Integer
    Dim conector As String
    
    conector = ","
    
    If SaltoDeLinea Then
        conector = conector & Chr(13) + Chr(10)
    Else
        conector = conector & " "
    End If
    
    If IsArray(Matrix) Then
        T = UBound(Matrix)
        o = LBound(Matrix)
        txt = sepLeft & Matrix(o) & sepRight
        For i = o + 1 To T
            txt = txt & conector & sepLeft & Matrix(i) & sepRight
        Next i
    Else
        txt = Matrix
    End If
    If txt = "Falso" Then txt = ""
    MatrixToText = txt
End Function

Sub InsertarColumnasNuevas(ByVal ubicacion As Integer, ByVal Hoja As Worksheet, ParamArray Etiquetas() As Variant)
    Dim i, o, T As Integer
    Dim Indice As Integer
    
    o = LBound(Etiquetas)
    T = UBound(Etiquetas)
    
    For i = o To T
        Indice = T - i
        With Hoja
            .Cells(1, ubicacion).EntireColumn.Insert Shift:=xlToRight
            .Cells(1, ubicacion).Value = Etiquetas(Indice)
        End With
    Next i
End Sub

Public Function EncontrarMaximaFila(ByVal Hoja As Worksheet, ParamArray ColumnasCriterio() As Variant) As Double
    Dim i As Integer
    Dim o As Integer
    Dim T As Integer
    Dim Output As Double
    
    o = LBound(ColumnasCriterio)
    T = UBound(ColumnasCriterio)
    
    EncontrarMaximaFila = EncontrarUltimaFila(Hoja, ColumnasCriterio(o))
    o = o + 1
    For i = o To T
         EncontrarMaximaFila = Application.WorksheetFunction.Max(EncontrarMaximaFila, EncontrarUltimaFila(Hoja, ColumnasCriterio(i)))
    Next i
End Function

Public Function FechaFinalDefecto(ByVal Fecha As Date)
    If Fecha = 0 Then Fecha = 401769
    FechaFinalDefecto = Fecha
End Function

Public Function ValoresUnicos(ByVal Rng As Range) As Variant
    Dim celda As Range
    Dim Unicos()
    Dim RngAcotado As Range
    Dim j As Integer
    
    If Rng Is Nothing Then Exit Function
    j = 0
    For Each celda In Rng.Cells
        If j = 0 Then
            ReDim Unicos(j)
            Unicos(j) = celda.Value
            j = j + 1
        ElseIf j > 0 And celda.Value <> "" And UbicacionDelElementoEnVector(celda.Value, Unicos) = -1 Then
            ReDim Preserve Unicos(j)
            Unicos(j) = celda.Value
            j = j + 1
        End If
    Next celda
    
    ValoresUnicos = Unicos
End Function

Public Function CondicionFecha( _
    ByVal Fecha As Date, _
    ByVal FechaInferior As Date, _
    ByVal FechaSuperior As Date) As Boolean
    
    CondicionFecha = False
        
    If Format(Fecha, "dd/mm/yyyy") >= FechaInferior And _
        Format(Fecha, "dd/mm/yyyy") <= FechaSuperior _
        Then CondicionFecha = True
End Function

Sub FuncionSeleccion( _
    ByVal HojaInput_x As Worksheet, _
    ByVal ColumnaCriterio_x As Integer, _
    ByVal ColumnaTotales_x As Integer, _
    ByVal ColumnaSeleccion_x As Integer, _
    ByVal HojaOutput_y As Worksheet, _
    ByVal ColumnaCriterio_y As Integer, _
    ByVal ColumnaProbabilidad_y As Integer, _
    ByVal ColumnaSeleccion_y As Integer, _
    Optional ByVal QuitarNoSeleccion As Boolean = True, _
    Optional ByVal MensajeSeleccion As String = "Seleccionada", _
    Optional ByVal MensajeNoSeleccion As String = "")

    Dim i As Double, j As Double, n_x As Double, n_y As Double
    Dim Limites() As Integer
    Dim acuml As Double
    
    n_x = EncontrarMaximaFila(HojaInput_x, ColumnaTotales_x, ColumnaSeleccion_x)
    n_y = EncontrarUltimaFila(HojaOutput_y, ColumnaProbabilidad_y)

    Call FiltrarOrdenarHojaPorColumna(ColumnaProbabilidad_y, HojaOutput_y, n_y, False)
    Call FiltrarOrdenarHojaPorColumna(ColumnaCriterio_y, HojaOutput_y, n_y, False)
    Call FiltrarOrdenarHojaPorColumna(ColumnaCriterio_x, HojaInput_x, n_x, False)

    ReDim Limites(n_x - 1, 2) As Integer

    For i = 2 To n_x
        With HojaInput_x
            acuml = acuml + .Cells(i, ColumnaTotales_x).Value
            Limites(i - 1, 1) = acuml - .Cells(i, ColumnaTotales_x).Value + 1
            Limites(i - 1, 2) = Limites(i - 1, 1) + .Cells(i, ColumnaSeleccion_x).Value - 1
        End With
    Next i

    For i = 2 To n_y
        For j = 2 To n_x
            If i - 1 >= Limites(j - 1, 1) And i - 1 <= Limites(j - 1, 2) Then
                With HojaOutput_y.Cells(i, ColumnaSeleccion_y)
                    If .Value = "" Then
                        .Value = MensajeSeleccion
                        .NumberFormat = "General"
                    Else
                        .Value = MensajeNoSeleccion
                        .NumberFormat = "General"
                    End If
                End With
            End If
        Next j
    Next i

    If QuitarNoSeleccion Then Call QuitarFilasEnBlancoC(ColumnaSeleccion_y, n_y, HojaOutput_y)
End Sub

Sub EliminarColumnasIgualesA(ByVal Hoja As Worksheet, ParamArray Criteria() As Variant)
    Dim i As Integer, U As Integer
    Dim N_Columna As Integer
    
    U = UBound(Criteria)
    
    For i = 0 To U
        N_Columna = BuscadorDeEncabezado(Hoja, Criteria(i))
        If N_Columna <> 0 Then Hoja.Cells(1, N_Columna).EntireColumn.Delete
    Next i
End Sub

Sub EliminarColumnasDistintasA(ByVal Hoja As Worksheet, ParamArray Criteria() As Variant)
    Dim i As Integer, U As Integer
    Dim N_Columna As Integer
    
    U = NumeroDeColumnas(Hoja)
    i = 1
    Do While i <= U
        If ElementoEnVector(Hoja.Cells(1, i).Value, Criteria) = 0 Then
            Hoja.Cells(1, i).EntireColumn.Delete
            U = NumeroDeColumnas(Hoja)
        Else
            i = i + 1
        End If
    Loop
End Sub

Public Function ColumnToMatrix(ByVal Columna As Integer, _
    Optional ByVal FilaInicial As Double = 2, _
    Optional ByVal Hoja As Worksheet) As Variant
    
    Dim Rng As Range
    Dim T As Double
    Dim celda As Range
    Dim Matriz()
    Dim Dimension As Integer
    
    If Hoja Is Nothing Then Set Hoja = ActiveSheet
    T = EncontrarUltimaFila(Hoja, Columna)
    Set Rng = RangoTotalColumna(Columna, FilaInicial, T, Hoja)
    
    Dimension = 0
    For Each celda In Rng.Cells
        ReDim Preserve Matriz(Dimension)
        Matriz(Dimension) = celda.Value
        Dimension = Dimension + 1
    Next celda
    
    ColumnToMatrix = Matriz
End Function

Sub CrearHojaIterada(Optional HojaName As String = "Hoja", _
    Optional Before As Boolean = False, _
    Optional ByVal HojaRerencia As Worksheet)
    
    Dim Sufijo As String
    Dim Hoja As Worksheet
    Dim j As Integer
    
    If HojaRerencia Is Nothing Then Set HojaRerencia = ActiveSheet
    
    If Before Then
        Set Hoja = Worksheets.Add(Before:=HojaRerencia)
    Else
        Set Hoja = Worksheets.Add(After:=HojaRerencia)
    End If
    
    j = 1
    On Error Resume Next
    Hoja.Name = HojaName & j
    Do While Err.Number <> 0
        j = j + 1
        Hoja.Name = HojaName & j
        If Hoja.Name = HojaName & j Then
            On Error GoTo 0
        End If
    Loop
End Sub

Sub HojasDeLibroToCombo(ByVal Combo As Object, _
    ByVal PathArchive As String, _
    Optional ByVal Pass As String = "", _
    Optional ByVal LimpiarCombo As Boolean = True)
    
    Dim Libro As Workbook
    Dim Hoja As Worksheet

    Call EncenderAcelerarMacro
    If LimpiarCombo Then Combo.Clear
    On Error Resume Next
    Set Libro = Workbooks.Open(FileName:=PathArchive, UpdateLinks:=0, Password:=Pass)
    If Err.Number <> 0 Then Exit Sub
    On Error GoTo 0
    For Each Hoja In Libro.Worksheets
        Combo.AddItem Hoja.Name
    Next Hoja
    Libro.Close SaveChanges:=False
    Call ApagarAcelerarMacro
End Sub

Sub ComboToCombo(ByVal ComboBoxInput As Object, _
    ByVal ComboBoxOutput As Object, _
    Optional ByVal BorrarComboOutput As Boolean = True)
    
    Dim Item As Variant
    
    If BorrarComboOutput Then ComboBoxOutput.Clear
    For Each Item In ComboBoxInput.List
        If Item <> "" Then
            ComboBoxOutput.AddItem Item
        Else
            Exit For
        End If
    Next Item
End Sub

Sub ColocarFiltros(Optional ByVal Hoja As Worksheet, Optional Rng As Range)
    If Hoja Is Nothing Then Set Hoja = ActiveSheet
    If Rng Is Nothing Then Set Rng = Hoja.UsedRange
    
    If Hoja.AutoFilterMode Then Rng.AutoFilter
    Rng.AutoFilter
End Sub

Public Function RangoFiltrado(Optional ByVal ColumnaFiltro As Integer, _
    Optional ByVal Criteria As String, _
    Optional ByVal ColumnasFiltradas, _
    Optional ByVal Rng As Range, _
    Optional ByVal Hoja As Worksheet) As Range
    
    Dim i As Integer
    
    If Rng Is Nothing Then Set Rng = Hoja.UsedRange
    
    If ColumnaFiltro <> 0 And Criteria <> "" Then
        Call ColocarFiltros(Hoja, Rng)
        Rng.AutoFilter Field:=ColumnaFiltro, Criteria1:=Criteria
    End If
    
    If IsArray(ColumnasFiltradas) Then
        For i = LBound(ColumnasFiltradas) To UBound(ColumnasFiltradas)
            If IsNumeric(ColumnasFiltradas(i)) Then _
                Hoja.Columns(ColumnasFiltradas(i)).EntireColumn.Hidden = True
        Next i
    Else
        If IsNumeric(ColumnasFiltradas) Then _
            Hoja.Columns(ColumnasFiltradas).EntireColumn.Hidden = True
    End If
    
    Set RangoFiltrado = Hoja.UsedRange.SpecialCells(xlCellTypeVisible)
End Function

Public Function RangoNoBlanks(Optional ByVal Hoja As Worksheet) As Range
    Dim RngUsado As Range
    Dim celda As Range
    Dim RngNoBlanks As Range
    
    If Hoja Is Nothing Then Set Hoja = ActiveSheet
    Set RngUsado = Hoja.UsedRange
    For Each celda In RngUsado.Cells
        If celda.Value <> "" Then
            If RngNoBlanks Is Nothing Then
                Set RngNoBlanks = celda
            Else
                Set RngNoBlanks = Application.Union(RngNoBlanks, celda)
            End If
        End If
    Next celda
    
    Set RangoNoBlanks = RngNoBlanks
End Function

Public Function ColumnasDeEncabezado(ByVal Hoja As Worksheet, ParamArray Criterios() As Variant)
    Dim i As Integer, j As Integer
    Dim Arreglo()
    Dim Columna As Integer
    j = 0
    For i = LBound(Criterios) To UBound(Criterios)
        Columna = BuscadorDeEncabezado(Hoja, Criterios(i))
        If Columna >= 1 Then
            ReDim Preserve Arreglo(j)
                Arreglo(j) = Columna
            j = j + 1
        End If
    Next i
    
    ColumnasDeEncabezado = Arreglo
End Function

Sub ObjetoHabilitado(ByVal Objeto As Object, ByVal Habilitado As Boolean)
    Objeto.Enabled = Habilitado
    
    If Habilitado Then
        Objeto.BackColor = vbWhite
    Else
        Objeto.BackColor = RGB(220, 220, 220)
    End If
End Sub

Public Function FilaElementoRango(ByVal Elemento As Variant, ByVal Columna As Integer, Optional ByVal Hoja As Worksheet) As Double
    Dim Rng As Range
    
    On Error GoTo NoEncontrado
    If Hoja Is Nothing Then Set Hoja = ActiveSheet
    Set Rng = RangoTotalColumna(Columna, 2, Hoja:=Hoja)
    FilaElementoRango = Application.WorksheetFunction.Match(Elemento, Rng, 0) + 1
    Exit Function
NoEncontrado:
    FilaElementoRango = 0
End Function
