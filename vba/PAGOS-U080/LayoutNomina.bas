Attribute VB_Name = "Módulo3"
Sub LayoutNomina()
    MsgBox General() & Chr(10) & "Archivo : " & "NominaSantander.txt"
End Sub
Function General() As String
On Error GoTo err_GeneraL
    Dim Registro As String
    Dim Cont As Long
    Dim cont2 As Long
    Dim cont3 As Long
    Dim RegEnt As Long
    Dim SumaImporte As Double
    Dim ArchivoSal As String
    Dim Continua As Boolean
    Dim Fila As Long
    Dim FechaIng As Date
    Dim FechaIngS As String
    Dim RFC As String
    Dim Importes As Double
    cont3 = 1
    Worksheets(3).Range("E3:G200").ClearContents
    RUTA = ThisWorkbook.Path & "\"
    NOMBRE = "NominaSantander" & cont3 & ".txt"
    ArchivoSal = RUTA & NOMBRE
    Open ArchivoSal For Output As 1
   ' Print "PAGO A PROVEEDORES"
    Continua = True
    Fila = 12
    'REGISTRO TIPO  1 ENCABEZADO
    Registro = Fll(NZ(1, ""), 1) 'TIPO DE REGISTRO
    Registro = Registro & Format(Fll(NZ(Cont, 0), 5), "0000") 'CONSECUTIVO
    Registro = Registro & Fll(NZ(1, ""), 1) 'NUMERO DE EMPLEADO
    Registro = Registro & "E" 'LETRA E
    Registro = Registro & Format(Fll(NZ(Range("C9"), 0), 8), "mmddyyyy") 'FECHA
    Registro = Registro & Fll(NZ(Range("C6"), " "), 11) 'CUENTA DE CARGO
    Registro = Registro & Fll(NZ(Range("H" & Fila), " "), 5) 'ESPACIOS EN BLANCO
    Registro = Registro & Format(Range("C9"), "mmddyyyy") 'FECHA
    Print #1, UCase(Registro)
    'REGISTRO TIPO 2 PAGOS DE NOMINA
    Do While Continua
    If Len(Range("A" & Fila)) > 0 Then
    If Len(Range("B" & Fila)) > 0 And Range("B" & Fila) = "BANME" Then
        Registro = Fll(NZ(2, ""), 1) 'TIPO DE REGISTRO
        Registro = Registro & Format(Fll(NZ(Cont + 2, 0), 5), "00000") 'CONSECUTIVO
        Registro = Registro & Fll(NZ(Cont + 1, " "), 7) 'NUMERO DE EMPLEADO
        Registro = Registro & Fll(NZ(Range("C" & Fila), " "), 80) 'BENEFICIARIO
        Registro = Registro & Fll(NZ(Mid(Range("A" & Fila), 7, 11), " "), 16) 'CUENTA DE ABONO
        Registro = Registro & Format(Fll(Round(NZ(Range("F" & Fila), 0), 2) * 100, 18), "000000000000000000") 'IMPORTE DE LA TRANSFERENCIA
        Registro = Registro & "01" 'FIN DE LINEA
        Cont = Cont + 1
        cont2 = cont2 + 1
        Importes = Importes + Range("F" & Fila)
        Print #1, UCase(Registro)
        If cont2 = 25000 Then
            Worksheets(3).Range("E" & cont3 + 2) = NOMBRE
            Worksheets(3).Range("F" & cont3 + 2) = cont2
            Worksheets(3).Range("G" & cont3 + 2) = Importes
            Registro = Fll(NZ(3, ""), 1) 'TIPO DE REGISTRO
            Registro = Registro & Format(Fll(NZ(Cont + 2, 0), 5), "00000") 'CONSECUTIVO
            Registro = Registro & Format(Fll(NZ(Cont, 0), 5), "00000") 'NUMERO DE EMPLEADO
            Registro = Registro & Format(Fll(Round(NZ(Importes, 0), 2) * 100, 18), "000000000000000000") 'IMPORTE DE LA TRANSFERENCIA
            Cont = 0
            cont2 = 0
            cont3 = cont3 + 1
            Importes = 0
            Print #1, UCase(Registro)
            Close #1
            ArchivoSal = RUTA & "NominaSantander" & cont3 & ".txt"
            Open ArchivoSal For Output As 1
            Registro = Fll(NZ(1, ""), 1) 'TIPO DE REGISTRO
            Registro = Registro & Format(Fll(NZ(Cont, 0), 5), "0000") 'CONSECUTIVO
            Registro = Registro & Fll(NZ(1, ""), 1) 'NUMERO DE EMPLEADO
            Registro = Registro & "E" 'LETRA E
            Registro = Registro & Format(Fll(NZ(Range("C9"), 0), 8), "mmddyyyy") 'FECHA
            Registro = Registro & Fll(NZ(Range("C6"), " "), 11) 'CUENTA DE CARGO
            Registro = Registro & Fll(NZ(Range("H" & Fila), " "), 5) 'ESPACIOS EN BLANCO
            Registro = Registro & Format(Range("C9"), "mmddyyyy") 'FECHA
            Print #1, UCase(Registro)
        End If
    End If
    Fila = Fila + 1
    Else
        Continua = False
    End If
Loop
General = "REGISTROS PROCESADOS " & Cont

sal_GeneraL:
    Worksheets(3).Range("E" & cont3 + 2) = NOMBRE
    Worksheets(3).Range("F" & cont3 + 2) = cont2
    Worksheets(3).Range("G" & cont3 + 2) = Importes
    'REGISTRO TIPO 3
    Registro = Fll(NZ(3, ""), 1) 'TIPO DE REGISTRO
    Registro = Registro & Format(Fll(NZ(Cont + 2, 0), 5), "00000") 'CONSECUTIVO
    Registro = Registro & Format(Fll(NZ(Cont, 0), 5), "00000") 'NUMERO DE EMPLEADO
    Registro = Registro & Format(Fll(Round(NZ(Importes, 0), 2) * 100, 18), "000000000000000000") 'IMPORTE DE LA TRANSFERENCIA
    Print #1, UCase(Registro)
    Close #1
    Exit Function
    
err_GeneraL:
    General = Err.Description
    Resume sal_GeneraL
    Resume Next
End Function
Function Fll(Texto As String, Tamaño As Long) As String
On Error GoTo err_Fll
    Dim TextL As String
    TextL = Left(Texto, Tamaño)
    TextL = Replace(TextL, "(", " ")
    TextL = Replace(TextL, ")", " ")
    TextL = Replace(TextL, ".", " ")
    TextL = Replace(TextL, ",", " ")
    TextL = Replace(TextL, "°", " ")
    TextL = Replace(TextL, "'", " ")
    TextL = Replace(TextL, "!", " ")
    TextL = Replace(TextL, "#", " ")
    TextL = Replace(TextL, "%", " ")
    TextL = Replace(TextL, "=", " ")
    TextL = Replace(TextL, "?", " ")
    TextL = Replace(TextL, "¡", " ")
    TextL = Replace(TextL, "¿", " ")
    TextL = Replace(TextL, "*", " ")
    TextL = Replace(TextL, "{", " ")
    TextL = Replace(TextL, "}", " ")
    TextL = Replace(TextL, "[", " ")
    TextL = Replace(TextL, "]", " ")
    TextL = Replace(TextL, ">", " ")
    TextL = Replace(TextL, "<", " ")
    TextL = Replace(TextL, ";", " ")
    TextL = Replace(TextL, ":", " ")
    TextL = Replace(TextL, "_", " ")
    TextL = Replace(TextL, "-", " ")
    TextL = Replace(TextL, "+", " ")
    TextL = Replace(TextL, "-", " ")
    TextL = Replace(TextL, "&", " ")
    TextL = Replace(TextL, "|", " ")
    TextL = Replace(TextL, "/", " ")
    TextL = Replace(TextL, "á", "A")
    TextL = Replace(TextL, "é", "E")
    TextL = Replace(TextL, "í", "I")
    TextL = Replace(TextL, "ó", "O")
    TextL = Replace(TextL, "ú", "U")
    TextL = Replace(TextL, "Á", "A")
    TextL = Replace(TextL, "É", "E")
    TextL = Replace(TextL, "Í", "I")
    TextL = Replace(TextL, "Ó", "O")
    TextL = Replace(TextL, "Ú", "U")
    TextL = Replace(TextL, "ü", "U")
    TextL = Replace(TextL, "ö", "O")
    TextL = UCase(TextL)
    For i = 33 To 47
        TextL = Replace(TextL, Chr(i), " ")
    Next i
    For i = 123 To 255
        TextL = Replace(TextL, Chr(i), " ")
    Next i
    Fll = NZ(TextL, " ") & Space(Tamaño - Len(NZ(TextL, " ")))
sal_Fll:
    Exit Function
    
err_Fll:
    Fll = Space(Tamaño)
    Resume sal_Fll
   
End Function

Function NZ(Dato As Variant, Def As Variant) As Variant
    On Error GoTo err_NZ
    
    If IsNull(Dato) Then
        NZ = Def
    Else
        NZ = Dato
    End If
Sal_NZ:
    Exit Function
    
err_NZ:
    NZ = Def
    Exit Function
    
End Function
Function QuitaRuta(Archivo As String) As String
On Error GoTo err_QuitaRuta

    Dim UltDiag As Long
    Dim UltDot As Long
    Dim Tam As Long
    UltDiag = InStrRev(Archivo, "\")
    Tam = Len(Archivo)
    If UltDiag > 0 Then
        QuitaRuta = Right(Archivo, Tam - UltDiag)
    Else
        QuitaRuta = Archivo
    End If
    UltDot = InStrRev(QuitaRuta, ".")
    If UltDot > 0 Then
        QuitaRuta = Left(QuitaRuta, UltDot - 1)
    End If
    
sal_QuitaRuta:
    Exit Function
    
err_QuitaRuta:
    QuitaRuta = "Error: " & Err.Description
    Resume sal_QuitaRuta
    
End Function

