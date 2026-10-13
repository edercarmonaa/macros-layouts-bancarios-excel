Attribute VB_Name = "Módulo2"
Sub ValidaClabes()
i = 12
erronea = 0
nomina = 0
bancomer = 0
interbanc = 0
Do Until IsEmpty(Hoja1.Cells(i, 1)) And IsEmpty(Hoja1.Cells(i, 1).Offset(1, 0))
    clabe = Hoja1.Cells(i, 1)
    If Len(clabe) > 0 Then
        If Len(clabe) <> 18 Then
            Var = msgError("Longitud de CLABE Incorrecta", i, 7)
            erronea = erronea + Hoja1.Cells(i, 6)
        ElseIf InStr(clabe, " ") Then
            Var = msgError("CLABE contiene espacios", i, 7)
            erronea = erronea + Hoja1.Cells(i, 6)
        ElseIf IsNumeric(clabe) = False Then
            Var = msgError("CLABE contiene letras", i, 7)
            erronea = erronea + Hoja1.Cells(i, 6)
        ElseIf DigitoVerificador(CStr(clabe)) = False Then
            Var = msgError("CLABE Incorrecta", i, 7)
            erronea = erronea + Hoja1.Cells(i, 6)
        Else
             Hoja1.Cells(i, 2) = ClaveTransfer(Mid(Hoja1.Cells(i, 1), 1, 3))
             Hoja1.Cells(i, 4) = "01001"
             Hoja1.Cells(i, 5) = "01001"
             Var = msgOk("Correcto", i, 7)
             If Hoja1.Cells(i, 2) = "BANME" Then
                nomina = nomina + Hoja1.Cells(i, 6)
             Else
                interbanc = interbanc + Hoja1.Cells(i, 6)
             End If
        End If
    Else
        Hoja1.Cells(i, 7) = ""
        Hoja1.Cells(i, 7).Interior.Color = RGB(255, 255, 255)
        Hoja1.Cells(i, 7).Font.Color = RGB(0, 0, 0)
    End If
    Hoja1.Cells(5, 7) = erronea
    Hoja1.Cells(6, 7) = interbanc
    Hoja1.Cells(7, 7) = nomina
    i = i + 1
Loop
End Sub

Function msgError(mensaje As String, col, fil) As Boolean
    Hoja1.Cells(col, fil) = mensaje
    Hoja1.Cells(col, fil).Interior.Color = RGB(204, 51, 0)
    Hoja1.Cells(col, fil).Font.Color = RGB(255, 255, 255)
End Function

Function msgOk(mensaje As String, col, fil) As Boolean
    Hoja1.Cells(col, fil) = mensaje
    Hoja1.Cells(col, fil).Interior.Color = RGB(146, 208, 80)
    Hoja1.Cells(col, fil).Font.Color = RGB(0, 0, 0)
End Function

Function DigitoVerificador(celda As String) As Boolean
    peso = Array(3, 7, 1, 3, 7, 1, 3, 7, 1, 3, 7, 1, 3, 7, 1, 3, 7)
    i = 1
    suma = 0
    For a = 0 To UBound(peso)
       modulo = (peso(a) * Mid(celda, i, 1)) Mod 10
       suma = suma + modulo
       i = i + 1
    Next a
    digito = (10 - (suma Mod 10)) Mod 10
    If CInt(digito) <> CInt(Mid(celda, 18, 1)) Then
        DigitoVerificador = False
        Exit Function
    Else
        DigitoVerificador = True
        Exit Function
    End If
End Function

Function ClaveTransfer(clave As String) As String
    Dim lookupvalue As Variant, lookupRange As Range
    Set lookupRange = Worksheets(2).Range("A2:C48")
    lookupvalue = Application.VLookup(clave, lookupRange, 3, False)
    If IsError(lookupvalue) Then
        Exit Function
    Else
        ClaveTransfer = lookupvalue
    End If
End Function
