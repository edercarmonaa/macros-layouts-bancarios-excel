Attribute VB_Name = "Módulo1"
Sub LayoutInter()
    MsgBox General() & Chr(10) & "Archivo : " & "Interbancarias.txt"
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
    Worksheets(3).Range("A3:C200").ClearContents
    cont3 = 1
    RUTA = ThisWorkbook.Path & "\"
    NOMBRE = "Interbancarias" & cont3 & ".txt"
    ArchivoSal = RUTA & NOMBRE
    Open ArchivoSal For Output As 1
   ' Print "PAGO A PROVEEDORES"
    Continua = True
    Fila = 12
    Do While Continua
    If Len(Range("A" & Fila)) > 0 Then
    If Len(Range("B" & Fila)) > 0 And (Range("B" & Fila) <> "BACOM") Then
        Registro = Fll(NZ(Range("A" & Fila), ""), 18) 'CUENTA ABONO
        Registro = Registro & Format(Fll(Round(NZ(Range("C6"), 0), 2), 18), "000000000000000000") 'CUENTA CARGO
        Registro = Registro & "MXP" 'MONEDA
        Registro = Registro & Format(Range("F" & Fila), "0000000000000.00") 'IMPORTE DE LA TRANSFERENCIA
        Registro = Registro & Fll(NZ(Range("C" & Fila), " "), 30) 'BENEFICIARIO
        Registro = Registro & "40" 'TIPO
        Registro = Registro & Mid(Range("A" & Fila), 1, 3) 'CLAVE
        Registro = Registro & Fll(NZ(Range("C7"), " "), 30) 'CONCEPTO
        Registro = Registro & "0000001H" 'FIN
        Cont = Cont + 1
        cont2 = cont2 + 1
        Importes = Importes + Range("F" & Fila)
        Print #1, UCase(Registro)
    End If
    Fila = Fila + 1
    Else
        Continua = False
    End If
Loop
General = "REGISTROS PROCESADOS " & Cont

sal_GeneraL:
    Worksheets(3).Range("A" & cont3 + 2) = NOMBRE
    Worksheets(3).Range("B" & cont3 + 2) = cont2
    Worksheets(3).Range("C" & cont3 + 2) = Importes
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
