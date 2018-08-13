Option Explicit

'Attribute VB_Name = "UtilesMacroBF"

'*******************************************************************
'*******************************************************************
'***UtilesMacroUniversales******************************************
'***Todas las macros son de autoria propia, las cuales pueden*******
'***ser copiadas mediante copyleft**********************************
'***Propietario: Omar André de la Sota******************************
'***correo: aozoro@gmail********************************************
'***Módulos alternos: UtilesMacroUniversales************************
'*******************************************************************

Sub ModificarDNI(ByVal Rng As Range)
    Dim celda As Range
    For Each celda In Rng.Cells
        On Error Resume Next
        With celda
            .Value = celda.Value * 1
            .NumberFormat = "General"
        End With
        If Err.Number <> 0 Then celda.Value = "DNI incorrecto"
    Next celda
End Sub

Sub EliminarCE(ByVal Rng As Range)
    Dim celda As Range
    Dim largo As Integer
    Dim CE As String

    CE = "CE"
    For Each celda In Rng.Cells
        Do While Left(celda.Value, 2) = CE
            celda.Value = Mid(celda.Value, 3)
        Loop
    Next
End Sub

Sub QuitarColumnasDigitacion()
    Dim i As Integer, j As Integer, K As Integer
    Dim VectorBorrado(6) As String
    Dim DimensionVector As Integer
    
    K = Range("A1").End(xlToRight).column
    DimensionVector = 6
    VectorBorrado(0) = "PRODUCTO"
    VectorBorrado(1) = "SUBPRODUCTO"
    VectorBorrado(2) = "CENTRO ALTA"
    VectorBorrado(3) = "PAN"
    VectorBorrado(4) = "FECHA FACTURA"
    VectorBorrado(5) = "DATO INICIAL"
    VectorBorrado(6) = "DATO FINAL"
    
    Call QuitarEspaciosDiferentes(Range("A1", Range("A1").End(xlToRight)))
    
    For i = 1 To K
        For j = 0 To DimensionVector
            If Cells(1, i).Value = VectorBorrado(j) Then Cells(1, i).EntireColumn.Delete
        Next j
    Next i
    
End Sub

Public Sub EncabezadosDeNormalizacion(Optional ByVal Hoja As Worksheet)
    If Hoja Is Nothing Then Set Hoja = ActiveSheet

    Call ColocarEncabezados(Hoja, "ENTIDAD", "MARCA", "TIPO", "PRODUCTO", "SUBPRODUCTO", _
        "CENTRO ALTA", "NUM. CONTRATO", "PAN", "FECHA FACTURA", "FECHA PROCESO", "FACTURA SISTEMA", "FACTURA", _
        "CODIGO C.C.", "DESCRIPCION", "SIGNO", "IMPORTE", "IMPORTE LIQ", "CUOTAS", "DATO INICIAL", "DATO FINAL", _
        "N/C", "USUARIO", "ORIGEN")
End Sub

Sub InsertarNombresSegunDNI(Optional ByVal Hoja As Worksheet = Nothing, _
    Optional ByVal ElimacionDesconocidos As Boolean = False)
    
    Dim i As Double, N As Double, K As Integer
    Dim i_Usuario As Integer, i_Nombre As Integer
    Dim Rng As Range
    Dim DNI_Nombres As Worksheet
    
    If Hoja Is Nothing Then Set Hoja = ActiveSheet
        
    Set DNI_Nombres = ThisWorkbook.Worksheets("DNI-Nombre")
    Set Rng = DNI_Nombres.Range("A1").CurrentRegion
    
    i_Usuario = BuscadorDeEncabezado(Hoja, "USUARIO", "USUARIOUMO")
    
    If i_Usuario = 0 Then
        Exit Sub
    End If
    
    N = EncontrarUltimaFila(Hoja, i_Usuario)
    K = NumeroDeColumnas(Hoja)
        
    i_Nombre = i_Usuario + 1
    
    With Hoja
        If .Cells(1, i_Nombre).Value <> "NOMBRE" Then
            .Cells(1, i_Nombre).EntireColumn.Insert Shift:=xlToRight
        End If
        
        .Cells(1, i_Nombre).Value = "NOMBRE"
    
        For i = 2 To N
            On Error Resume Next
            .Cells(i, i_Nombre).Value = Application.WorksheetFunction.VLookup(.Cells(i, i_Usuario).Value * 1, Rng, 2, 0)
            If Err.Number <> 0 Then
                .Cells(i, i_Nombre).Value = ""
            End If
            On Error GoTo 0
        Next i
        
        If ElimacionDesconocidos = True Then
            Call QuitarFilasEnBlancoC(i_Nombre, N, Hoja)
        End If
    End With
End Sub

Sub TransformarCuenta(ByVal Rng As Range)
    Dim celda As Range
    Dim ValueCelda As Double

    For Each celda In Rng.Cells
        On Error Resume Next
        ValueCelda = celda.Value * 1
        If Err.Number <> 0 Then
            celda.Value = "columna incorrecta"
        Else
            If ValueCelda > 10 ^ 11 Then
                celda.Value = Fix(ValueCelda / 100)
            Else
                If ValueCelda = 0 Then celda.Value = "no se registró"
            End If
        End If
        celda.NumberFormat = "General"
    Next celda
End Sub

Public Function TransformarCC(Cuenta As Variant) As Variant
    Dim CuentaNumero As Double
    
    On Error Resume Next
    CuentaNumero = Cuenta * 1
    If Err.Number <> 0 Then
        TransformarCC = "No tiene Formato cuenta"
    Else
        If CuentaNumero > 10 ^ 11 Then
            TransformarCC = Fix(CuentaNumero / 100)
        Else
            TransformarCC = CuentaNumero
        End If
    End If
    On Error GoTo 0
End Function

Sub ColocarDNIyTransformarCuentas(ByVal OpcionDNI As Boolean, ByVal OpcionCuentas As Boolean, _
    Optional ByVal Hoja As Worksheet = Nothing, _
    Optional ByVal EliminacionNA As Boolean = False)
    
    Dim i_Cuenta As Integer
    Dim Rng As Range
    
    If Hoja Is Nothing Then Set Hoja = ActiveSheet
    
    If OpcionDNI = True Then
        Call InsertarNombresSegunDNI(Hoja, EliminacionNA)
    End If
    
    If OpcionCuentas = True Then
        i_Cuenta = BuscadorDeEncabezado_Cuenta
        If i_Cuenta = 0 Then
            Exit Sub
        End If
        Set Rng = RangoTotalColumna(Columna:=i_Cuenta, PrimeraFila:=2, Hoja:=Hoja)
        Call TransformarCuenta(Rng)
    End If
End Sub

Public Function CUENTADNI(ByVal Cuenta As Double, _
    Optional ByVal Libro As Workbook = Nothing)
    
    Dim NombreLibro As String
    Dim Limites As Variant
    Dim NumeroDeHoja As Integer
    
    If Libro Is Nothing Then
        NombreLibro = VerificarLibroAbierto("BuscarCuenta")
        If NombreLibro <> "" Then
            Set Libro = Workbooks(NombreLibro)
        Else
            Workbooks.Open FileName:=AddressBuscarCuenta()
            Set Libro = ActiveWorkbook
        End If
    End If
    
    On Error Resume Next
    NumeroDeHoja = Libro.Worksheets.Count
    NumeroDeHoja = Application.WorksheetFunction.VLookup(Cuenta, Libro.Worksheets(NumeroDeHoja).Range("A:B"), 2, 1)
    CUENTADNI = Application.WorksheetFunction.VLookup(Cuenta, Libro.Worksheets(NumeroDeHoja).Range("A:B"), 2, 0)
    If Err.Number <> 0 Then
        CUENTADNI = "N/A"
    End If
    On Error GoTo 0
    
End Function

Sub Etiquetas_Tabla12(Optional ByVal Hoja As Worksheet = Nothing)
    
    If Hoja Is Nothing Then Set Hoja = ActiveSheet

    Call ColocarEncabezados(Hoja, "CODENT", "CENTALTA", "CUENTA", _
       "NUMEXTCTA", "CLAMON", "NUMMOVEXT", "INDNORCOR", "TIPOFAC", _
        "FECFAC", "NUMREFFAC", "INDMOVEXT", "INDMOVANU", "INDRET", _
        "INDINCEST", "PAN", "CLAMONDIV", "IMPDIV", "IMPFAC", "CMBAPLI", _
        "NUMAUT", "CODCOM", "NOMCOMRED", "CODACT", "IMPLIQ", "CLAMONLIQ", _
        "IMPIMPTO", "FECPROCES", "CODPAIS", "NOMPOB", "FECCONTA", "ORIGENOPE", _
        "TIPFRAN", "SECOPE", "NUMSECREC", "SESIONRED", "SIAIDCD", "TIPDOCPAG", _
        "REFDOCPAG", "CALFRAUDE", "FECLIQ", "IMPAMORT", "LINREF", "FORPAGO", _
        "NUMOPECUO", "NUMFINAN", "NUMCUOTA", "TOTCUOTAS", "CODENTUMO", "CODOFIUMO", _
        "USUARIOUMO", "CODTERMUMO", "CONTCUR")
End Sub

Public Function BuscadorDeEncabezado_Cuenta(Optional ByVal Hoja As Worksheet) As Integer
    If Hoja Is Nothing Then Set Hoja = ActiveSheet
    BuscadorDeEncabezado_Cuenta = BuscadorDeEncabezado(Hoja, "CUENTA", "NUM CONTRATO", "NUM. CONTRATO", "NUM.CONTRATO")
End Function

Sub EliminarColumnasInnecesariasSAT(Optional ByVal Hoja As Worksheet)
    If Hoja Is Nothing Then Set Hoja = ActiveSheet
    Call EliminadorDeColumnaConEncabezado(Hoja, "ENTIDAD")
    Call EliminadorDeColumnaConEncabezado(Hoja, "MARCA")
    Call EliminadorDeColumnaConEncabezado(Hoja, "TIPO")
    Call EliminadorDeColumnaConEncabezado(Hoja, "PRODUCTO")
    Call EliminadorDeColumnaConEncabezado(Hoja, "SUBPRODUCTO")
    Call EliminadorDeColumnaConEncabezado(Hoja, "CENTRO ALTA")
    Call EliminadorDeColumnaConEncabezado(Hoja, "PAN")
    Call EliminadorDeColumnaConEncabezado(Hoja, "FECHA FACTURA")
    Call EliminadorDeColumnaConEncabezado(Hoja, "FACTURA SISTEMA")
    Call EliminadorDeColumnaConEncabezado(Hoja, "CUOTAS")
    Call EliminadorDeColumnaConEncabezado(Hoja, "IMPORTE LIQ")
    Call EliminadorDeColumnaConEncabezado(Hoja, "DATO INICIAL")
    Call EliminadorDeColumnaConEncabezado(Hoja, "DATO FINAL")
    Call EliminadorDeColumnaConEncabezado(Hoja, "N/C")
    Call EliminadorDeColumnaConEncabezado(Hoja, "ORIGEN")
    Call EliminadorDeColumnaConEncabezado(Hoja, "COINCIDENCIA")
End Sub

Sub BorrarImportesNA(Optional ByVal Hoja As Worksheet)
    If Hoja Is Nothing Then Set Hoja = ActiveSheet
    
    Dim i_Importe As Integer
    Dim N As Double
    
    i_Importe = BuscadorDeEncabezado(Hoja, "IMPORTE")
    N = EncontrarUltimaFila(Hoja, i_Importe)
    Call BorrarFilasIgualesA(Hoja, i_Importe, 2, N, "N/A", "NA")
End Sub

Public Sub DepurarCuentas(ByVal ColumnaCuenta As Integer, Optional ByVal Hoja As Worksheet)
    Dim Rng As Range
    
    Set Rng = RangoTotalColumna(Columna:=ColumnaCuenta, PrimeraFila:=2, Hoja:=Hoja)
    Call QuitarEspaciosDiferentes(Rng)
    Call QuitarApostrofes(Rng)
    Call EliminarCE(Rng)
End Sub

Public Sub DepurarDNI(ByVal ColumnaDNI As Integer, Optional ByVal Hoja As Worksheet)
    Dim Rng As Range
    
    Set Rng = RangoTotalColumna(Columna:=ColumnaDNI, PrimeraFila:=2, Hoja:=Hoja)
    Call QuitarEspaciosDiferentes(Rng)
    Call QuitarApostrofes(Rng)
End Sub