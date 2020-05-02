Attribute VB_Name = "Conversiones"
Function ConvertirNumeroATexto(Numero As Long) As String
Dim Letra As String
Const Maximo = 1999999999.99
'Validar que el Numero está dentro de los límites
If (Numero >= 0) And (Numero <= Maximo) Then

    Letra = NUMERORECURSIVO((Fix(Numero)))
    ConvertirNumeroATexto = Letra

Else
    'Si el Numero no está dentro de los límites, entivar un mensaje de error
    ConvertirNumeroATexto = "ERROR: El número excede los límites."


End If



End Function


Function NUMERORECURSIVO(Numero As Long) As String

Dim Unidades, Decenas, Centenas
Dim Resultado As String

'**************************************************
' Nombre de los números
'**************************************************
Unidades = Array("", "UN", "DOS", "TRES", "CUATRO", "CINCO", "SEIS", "SIETE", "OCHO", "NUEVE", "DIEZ", "ONCE", "DOCE", "TRECE", "CATORCE", "QUINCE", "DIECISÉIS", "DIECISIETE", "DIECIOCHO", "DIECINUEVE", "VEINTE", "VEINTIUNO", "VEINTIDOS", "VEINTITRES", "VEINTICUATRO", "VEINTICINCO", "VEINTISEIS", "VEINTISIETE", "VEINTIOCHO", "VEINTINUEVE")
Decenas = Array("", "DIEZ", "VEINTE", "TREINTA", "CUARENTA", "CINCUENTA", "SESENTA", "SETENTA", "OCHENTA", "NOVENTA", "CIEN")
Centenas = Array("", "CIENTO", "DOSCIENTOS", "TRESCIENTOS", "CUATROCIENTOS", "QUINIENTOS", "SEISCIENTOS", "SETECIENTOS", "OCHOCIENTOS", "NOVECIENTOS")
'**************************************************

Select Case Numero
    Case 0
        Resultado = "CERO"
    Case 1 To 29
        Resultado = Unidades(Numero)
    Case 30 To 100
        Resultado = Decenas(Numero \ 10) + IIf(Numero Mod 10 <> 0, " Y " + NUMERORECURSIVO(Numero Mod 10), "")
    Case 101 To 999
        Resultado = Centenas(Numero \ 100) + IIf(Numero Mod 100 <> 0, " " + NUMERORECURSIVO(Numero Mod 100), "")
    Case 1000 To 1999
        Resultado = "MIL" + IIf(Numero Mod 1000 <> 0, " " + NUMERORECURSIVO(Numero Mod 1000), "")
    Case 2000 To 999999
        Resultado = NUMERORECURSIVO(Numero \ 1000) + " MIL" + IIf(Numero Mod 1000 <> 0, " " + NUMERORECURSIVO(Numero Mod 1000), "")
    Case 1000000 To 1999999
        Resultado = "UN MILLÓN" + IIf(Numero Mod 1000000 <> 0, " " + NUMERORECURSIVO(Numero Mod 1000000), "")
    Case 2000000 To 1999999999
        Resultado = NUMERORECURSIVO(Numero \ 1000000) + " MILLONES" + IIf(Numero Mod 1000000 <> 0, " " + NUMERORECURSIVO(Numero Mod 1000000), "")
End Select

NUMERORECURSIVO = Resultado

End Function


