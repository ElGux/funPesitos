Attribute VB_Name = "Módulo Pesitos"
Option Explicit
'Main Function

'''
''Funcion recursiva que genera cadena de un valor numerico
''Regresa: Cadena traducida de curValor ($20.50 -> "veinte pesos con cincuenta centavos"
Public Function funPesitos(ByVal curValor As Currency, Optional bolConCentavos As Boolean = True, _
    Optional bolTrunca2Decimales As Boolean = True, Optional bolDecimalesATexto As Boolean = False, _
    Optional bolConPesos As Boolean = True, Optional bolCiclo As Boolean) As String

Dim lstrValor As String, lcurSub As Currency, lstrSalida As String, i As Integer
Dim lbolContinua As Boolean, lcurCentavos As Currency, intPunto As Integer
    
    ''Tomo valor absoluto y lo convierto en string
    lstrValor = CStr(Abs(curValor))
    
    ''Busco si hay punto
    intPunto = InStr(1, lstrValor, ".")
    
    If intPunto > 0 Then
        If bolTrunca2Decimales Then
            ''Tomo solo los 2 primeros decimales
            lcurCentavos = CCur(Mid(lstrValor & "00", intPunto + 1, 2))
        Else
            ''Tomo todos los decimales a la derecha del punto
            lcurCentavos = CCur(Right(lstrValor, Len(lstrValor) - intPunto))
        End If
        ''Quito de la cadena de valor los numeros decimales y el punto
        lstrValor = Left(lstrValor, intPunto - 1)
    End If
    
    If lstrValor = "0" And Not bolCiclo Then
        lstrSalida = "CERO"
    Else
        ''Reviso los valores mayores si son billones
        If Len(lstrValor) > 12 Then
            lstrSalida = "overflow"
        
        ''Reviso si son millones
        ElseIf Len(lstrValor) > 6 Then
            lcurSub = CCur(Left(lstrValor, Len(lstrValor) - 6))
            
            If lcurSub = 1 Then
                lstrSalida = "UN MILLON"
            Else
                lstrSalida = funPesitos(lcurSub, , , , , True) & " MILLONES"
            End If
            
            lcurSub = CCur(Right(lstrValor, 6))
            If lcurSub > 0 Then
                lstrSalida = lstrSalida & " " & funPesitos(lcurSub, , , , , True)
            End If
        
        ''Reviso si son miles
        ElseIf Len(lstrValor) > 3 Then
            lcurSub = CCur(Left(lstrValor, Len(lstrValor) - 3))
            
            If lcurSub = 1 Then
                lstrSalida = "UN MIL"
            Else
                lstrSalida = funPesitos(lcurSub, , , , , True) & " MIL"
            End If
            
            lcurSub = CCur(Right(lstrValor, 3))
            If lcurSub > 0 Then
                lstrSalida = lstrSalida & " " & funPesitos(lcurSub, , , , , True)
            End If
        
        ''Reviso si son centenas
        ElseIf Len(lstrValor) > 2 Then
            lcurSub = CCur(Left(lstrValor, 1))
            
            lstrSalida = lstrSalida & "CIENTO"
            If lcurSub = 1 Then
                If lstrValor = "100" Then
                    lstrSalida = "CIEN"
                End If
            ElseIf lcurSub = 2 Then
                lstrSalida = "DOS" & lstrSalida & "S"
            ElseIf lcurSub = 3 Then
                lstrSalida = "TRES" & lstrSalida & "S"
            ElseIf lcurSub = 4 Then
                lstrSalida = "CUATRO" & lstrSalida & "S"
            ElseIf lcurSub = 5 Then
                lstrSalida = "QUINIENTOS"
            ElseIf lcurSub = 6 Then
                lstrSalida = "SEIS" & lstrSalida & "S"
            ElseIf lcurSub = 7 Then
                lstrSalida = "SETE" & lstrSalida & "S"
            ElseIf lcurSub = 8 Then
                lstrSalida = "OCHO" & lstrSalida & "S"
            ElseIf lcurSub = 9 Then
                lstrSalida = "NOVE" & lstrSalida & "S"
            End If
            
            lcurSub = CCur(Right(lstrValor, 2))
            lstrSalida = lstrSalida & " " & funPesitos(lcurSub, , , , , True)
            
        ''Reviso si son decenas
        ElseIf Len(lstrValor) > 1 Then
            lcurSub = CCur(Left(lstrValor, 1))
            
            If lcurSub = 1 Then
                lstrSalida = "DIEZ"
            ElseIf lcurSub = 2 Then
                lstrSalida = "VEINTE"
            ElseIf lcurSub = 3 Then
                lstrSalida = "TREINTA"
            ElseIf lcurSub = 4 Then
                lstrSalida = "CUARENTA"
            ElseIf lcurSub = 5 Then
                lstrSalida = "CINCUENTA"
            ElseIf lcurSub = 6 Then
                lstrSalida = "SESENTA"
            ElseIf lcurSub = 7 Then
                lstrSalida = "SETENTA"
            ElseIf lcurSub = 8 Then
                lstrSalida = "OCHENTA"
            ElseIf lcurSub = 9 Then
                lstrSalida = "NOVENTA"
            End If
            
            lcurSub = CCur(Right(lstrValor, 1))
            If lcurSub > 0 Then
                If lstrSalida = "DIEZ" Then
                    If lcurSub = 1 Then
                        lstrSalida = "ONCE"
                    ElseIf lcurSub = 2 Then
                        lstrSalida = "DOCE"
                    ElseIf lcurSub = 3 Then
                        lstrSalida = "TRECE"
                    ElseIf lcurSub = 4 Then
                        lstrSalida = "CATORCE"
                    ElseIf lcurSub = 5 Then
                        lstrSalida = "QUINCE"
                    ElseIf lcurSub >= 6 Then
                        lstrSalida = "DIECI" & funPesitos(lcurSub, , , , , True)
                    End If
                ElseIf lstrSalida = "VEINTE" Then
                    If lcurSub = 1 Then
                        lstrSalida = "VEINTIUN"
                    Else
                        lstrSalida = "VEINTI" & funPesitos(lcurSub, , , , , True)
                    End If
                Else
                    lstrSalida = lstrSalida & " Y " & funPesitos(lcurSub, , , , , True)
                End If
            End If
        
        ''Reviso si son unidades
        ElseIf Len(lstrValor) = 1 Then
            lcurSub = CCur(lstrValor)
            
            If lcurSub = 1 Then
                lstrSalida = "UNO"
            ElseIf lcurSub = 2 Then
                lstrSalida = "DOS"
            ElseIf lcurSub = 3 Then
                lstrSalida = "TRES"
            ElseIf lcurSub = 4 Then
                lstrSalida = "CUATRO"
            ElseIf lcurSub = 5 Then
                lstrSalida = "CINCO"
            ElseIf lcurSub = 6 Then
                lstrSalida = "SEIS"
            ElseIf lcurSub = 7 Then
                lstrSalida = "SIETE"
            ElseIf lcurSub = 8 Then
                lstrSalida = "OCHO"
            ElseIf lcurSub = 9 Then
                lstrSalida = "NUEVE"
            End If
        End If
    End If
    
    If Not bolCiclo Then
        If bolConPesos Then
            If lstrValor = "1" Then
                lstrSalida = "UN PESO"
            Else
                lstrSalida = lstrSalida & " PESOS"
            End If
        End If
        
        If bolConCentavos Then
            If bolDecimalesATexto Then
                lstrSalida = lstrSalida & " CON " & funPesitos(lcurCentavos, , , , , True) & " CENTAVOS"
            Else
                lstrSalida = lstrSalida & " " & Format(lcurCentavos, "00") & "/100"
            End If
        Else
            If lcurCentavos <> 0 Then
                If bolDecimalesATexto Then
                    lstrSalida = lstrSalida & " PUNTO " & funPesitos(lcurCentavos, , , , , True)
                Else
                    lstrSalida = lstrSalida & " PUNTO " & lcurCentavos
                End If
            End If
        End If
        
        If bolConPesos Then
            lstrSalida = lstrSalida & " M.N."
        End If
        
    End If
    
    funPesitos = lstrSalida

End Function

