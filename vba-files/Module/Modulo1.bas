Attribute VB_Name = "Modulo1"
Option Explicit

Sub MemoLetraNumero()

    Hoja1.Range("E13").Value = LetrasMoneda(Hoja1.Range("E12").Value)
    Call LetrasMemo(Hoja1.Range("C10"), Hoja1.Range("E15"))
    
End Sub



' =============================================================
' Funcion principal: convierte numero a letras (Cordobas)
' Uso: =LetrasMoneda( E11 )   o desde VBA
' =============================================================
Public Function LetrasMoneda(ByVal Numero As Double) As String
    
    Dim ParteEntera   As Double
    Dim ParteDecimal  As Long
    Dim LetrasEntero  As String
    Dim LetrasDecimal As String
    Dim resultado     As String
    
    ' Separar parte entera y decimal
    ParteEntera = Int(Numero)
    ParteDecimal = Round((Numero - ParteEntera) * 100, 0)
    
    ' Convertir parte entera
    LetrasEntero = NumeroALetras(ParteEntera)
    
    ' Ajustes gramaticales para "Cordoba(s)"
    If ParteEntera = 1 Then
        LetrasEntero = Replace(LetrasEntero, "UN ", "UN ")
        resultado = LetrasEntero & " CCordoba"
    ElseIf ParteEntera = 0 Then
        resultado = "Cero Cordobas"
    Else
        resultado = LetrasEntero & " CCordobas"
    End If
    
    ' Parte decimal (centavos)
    If ParteDecimal = 0 Then
        LetrasDecimal = "Con Cero Centavos"
    Else
        LetrasDecimal = "Con " & NumeroALetras(ParteDecimal) & IIf(ParteDecimal = 1, " Centavo", " Centavos")
    End If
    
    ' Union y capitalizacion de palabras iniciales
    resultado = resultado & " " & LetrasDecimal
    resultado = CapitalizarCadaPalabra(resultado)
    
    LetrasMoneda = resultado
    
End Function


' =============================================================
' Nucleo: convierte numero entero (0 a 999.999.999) a letras
' =============================================================
Private Function NumeroALetras(ByVal n As Double) As String
    
    Dim Unidades() As Variant
    Dim Decenas()  As Variant
    Dim Centenas() As Variant
    Dim Especiales() As Variant
    
    Unidades = Array("", "Un", "Dos", "Tres", "Cuatro", "Cinco", "Seis", "Siete", "Ocho", "Nueve", _
                     "Diez", "Once", "Doce", "Trece", "Catorce", "Quince", "Diecis�is", "Diecisiete", _
                     "Dieciocho", "Diecinueve")
                     
    Decenas = Array("", "Diez", "Veinte", "Treinta", "Cuarenta", "Cincuenta", "Sesenta", "Setenta", _
                    "Ochenta", "Noventa")
                    
    Centenas = Array("", "Cien", "Doscientos", "Trescientos", "Cuatrocientos", "Quinientos", _
                     "Seiscientos", "Setecientos", "Ochocientos", "Novecientos")
                     
    Especiales = Array("Veintiun", "Veintiuna", "Veintidos", "Veintitres", "Veinticuatro", _
                       "Veinticinco", "Veintiseis", "Veintisiete", "Veintiocho", "Veintinueve")
    
    Dim s As String: s = ""
    Dim Millones As Long, Miles As Long, Cientos As Long
    
    If n = 0 Then
        NumeroALetras = "Cero"
        Exit Function
    End If
    
    ' Millones
    Millones = n \ 1000000
    If Millones > 0 Then
        If Millones = 1 Then
            s = "Un Mill�n"
        Else
            s = NumeroALetras(Millones) & " Millones"
        End If
        n = n - Millones * 1000000
    End If
    
    ' Miles
    Miles = n \ 1000
    If Miles > 0 Then
        If s <> "" Then s = s & " "
        If Miles = 1 Then
            s = s & "Mil"
        Else
            s = s & NumeroALetras(Miles) & " Mil"
        End If
        n = n - Miles * 1000
    End If
    
    ' Cientos / unidades
    Cientos = n
    If Cientos > 0 Then
        If s <> "" Then s = s & " "
        
        If Cientos >= 100 Then
            s = s & Centenas(Cientos \ 100)
            Cientos = Cientos Mod 100
            If Cientos > 0 Then s = s & " "
        End If
        
        If Cientos >= 20 And Cientos <= 29 Then
            s = s & Especiales(Cientos - 20)
        ElseIf Cientos >= 10 And Cientos <= 19 Then
            s = s & Unidades(Cientos)
        Else
            If Cientos >= 20 Then
                s = s & Decenas(Cientos \ 10)
                Cientos = Cientos Mod 10
                If Cientos > 0 Then s = s & " Y "
            End If
            If Cientos > 0 Then
                s = s & Unidades(Cientos)
            End If
        End If
    End If
    
    ' Correcciones gramaticales finales
    s = Replace(s, "Un Uno", "Uno")
    s = Replace(s, "Un Cien", "Cien")
    s = Replace(s, "Ciento Un", "Ciento Un")
    s = Replace(s, " Veintiun ", " Veintiun ")
    
    NumeroALetras = Trim(s)
    
End Function


' =============================================================
' Capitaliza la primera letra de cada palabra
' =============================================================
Private Function CapitalizarCadaPalabra(ByVal texto As String) As String
    
    Dim palabras() As String
    Dim i As Long
    Dim resultado As String
    
    palabras = Split(LCase(texto), " ")
    
    For i = LBound(palabras) To UBound(palabras)
        If Len(palabras(i)) > 0 Then
            palabras(i) = UCase(Left(palabras(i), 1)) & LCase(Mid(palabras(i), 2))
        End If
    Next i
    
    CapitalizarCadaPalabra = Join(palabras, " ")
    
End Function

'=============================================================================================
' Funcion para Extraer las Letras del Nombre, para el Memorandum
' Uso: Call LetrasMemo(inputRange, outputRange)
'=============================================================================================

Public Sub LetrasMemo(inputRange As Range, outputRange As Range)
    
    Dim nombre As String
    Dim iniciales As String
    
    nombre = Trim(inputRange.Value)
    
    ' Caso vacio o cadena vacia
    If nombre = "" Then
        outputRange.Value = "Sin Remitente"
    Else
        iniciales = ObtenerInicialesNombre(nombre)
        outputRange.Value = iniciales
    End If
    
End Sub


Private Function ObtenerInicialesNombre(ByVal texto As String) As String
    Dim palabras()      As String
    Dim palabra         As String
    Dim iniciales       As String
    Dim i               As Long
    
    ' Lista de conectores a omitir (puedes agregar mas si lo necesitas)
    Dim conectores As Variant
    conectores = Array("de", "del", "la", "las", "lo", "los", "el", "y", "e", "con", "en", "a", "al", "para")
    
    ' Separar por espacios y limpiar multiples espacios
    palabras = Split(Trim(Replace(texto, "  ", " ")), " ")
    
    iniciales = ""
    
    For i = LBound(palabras) To UBound(palabras)
        palabra = LCase(Trim(palabras(i)))
        
        ' Omitir si es conector o esta vacio
        If palabra <> "" And Not EstaEnArray(palabra, conectores) Then
            ' Tomamos la primera letra y la ponemos en mayuscula
            iniciales = iniciales & UCase(Left(palabra, 1))
        End If
    Next i
    
    ' Si no quedo ninguna inicial (solo conectores), devolvemos algo por defecto
    If iniciales = "" Then
        ObtenerInicialesNombre = "Sin Remitente"
    Else
        ObtenerInicialesNombre = iniciales
    End If
End Function


Private Function EstaEnArray(ByVal valor As String, arr As Variant) As Boolean
    Dim v As Variant
    EstaEnArray = False
    For Each v In arr
        If LCase(v) = LCase(valor) Then
            EstaEnArray = True
            Exit Function
        End If
    Next v
End Function
