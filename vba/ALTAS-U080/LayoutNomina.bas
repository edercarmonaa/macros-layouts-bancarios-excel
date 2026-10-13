Attribute VB_Name = "Módulo3"
Sub LayoutNomina()
    MsgBox General() & Chr(10) & "Archivo : " & "LOE_Alta_Empleados_Cat_Nom_.txt"
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
    Nombre = "LOE_Alta_Empleados_Cat_Nom_" & cont3 & ".txt"
    ArchivoSal = RUTA & Nombre
    Open ArchivoSal For Output As 1
   ' Print "PAGO A PROVEEDORES"
    Continua = True
    Fila = 12
    'REGISTRO TIPO 2 PAGOS DE NOMINA
    Do While Continua
    If Len(Range("A" & Fila)) > 0 Then
    If Len(Range("B" & Fila)) > 0 And Range("B" & Fila) = "BANME" Then
        Registro = Fll(NZ(Mid(Range("A" & Fila), 7, 11), " "), 18) 'CUENTA DE ABONO
        arrSplitStrings = Split(Trim(Range("C" & Fila)), " ")
        ArrayLen = UBound(arrSplitStrings) - LBound(arrSplitStrings) + 1
        Registro = Registro & Fll(NZ(arrSplitStrings(0), " "), 30) 'BENEFICIARIO
        Registro = Registro & Fll(NZ(arrSplitStrings(1), " "), 20) 'BENEFICIARIO
        Select Case ArrayLen
            Case Is < 3
                Registro = Registro & Fll(NZ(" ", " "), 30) 'BENEFICIARIO
            Case 3
                Registro = Registro & Fll(NZ(arrSplitStrings(2), " "), 30) 'BENEFICIARIO
            Case Is > 3
                strSingleString = ""
                For i = 2 To ArrayLen - 1
                    strSingleString = strSingleString & arrSplitStrings(i) & " "
                Next i
                Registro = Registro & Fll(NZ(strSingleString, " "), 30) 'BENEFICIARIO
        End Select
        Erase arrSplitStrings
        Cont = Cont + 1
        cont2 = cont2 + 1
        Importes = Importes + Range("F" & Fila)
        Print #1, UCase(Registro)
        If cont2 = 10000 Then
            Worksheets(3).Range("E" & cont3 + 2) = Nombre
            Worksheets(3).Range("F" & cont3 + 2) = cont2
            Worksheets(3).Range("G" & cont3 + 2) = Importes
            Registro = Fll(NZ(3, ""), 1) 'TIPO DE REGISTRO
            cont2 = 0
            cont3 = cont3 + 1
            Importes = 0
            Close #1
            ArchivoSal = RUTA & "LOE_Alta_Empleados_Cat_Nom_" & cont3 & ".txt"
            Open ArchivoSal For Output As 1
        End If
    End If
    Fila = Fila + 1
    Else
        Continua = False
    End If
Loop
General = "REGISTROS PROCESADOS " & Cont

sal_GeneraL:
    Worksheets(3).Range("E" & cont3 + 2) = Nombre
    Worksheets(3).Range("F" & cont3 + 2) = cont2
    Worksheets(3).Range("G" & cont3 + 2) = Importes
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

