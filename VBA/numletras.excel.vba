Function CONVERTIRNUMEROLETRA(tyCantidad As Currency) As String

    'A través de esta función de vba vamos a convertir números en texto’Dclaramos variables a utilizar en el código VBA"

    Dim lyCantidad As Currency, lyCentavos As Currency, lnDigito As Byte, lnPrimerDigito As Byte, lnSegundoDigito As Byte, lnTercerDigito As Byte, lcBloque As String, lnNumeroBloques As Byte, lnBloqueCero
    Dim laUnidades As Variant, laDecenas As Variant, laCentenas As Variant, I As Variant 'Si esta como Option Explicit

    tyCantidad = Round(tyCantidad, 2)
    lyCantidad = Int(tyCantidad)
    lyCentavos = (tyCantidad - lyCantidad) * 100

    laUnidades = Array("UN", "DOS", "TRES", "CUATRO", "CINCO", "SEIS", "SIETE", "OCHO", "NUEVE", "DIEZ", "ONCE", "DOCE", "TRECE", "CATORCE", "QUINCE", "DIECISEIS", "DIECISIETE", "DIECIOCHO", "DIECINUEVE", "VEINTE", "VEINTIUN", "VEINTIDOS", "VEINTITRES", "VEINTICUATRO", "VEINTICINCO", "VEINTISEIS", "VEINTISIETE", "VEINTIOCHO", "VEINTINUEVE")
    laDecenas = Array("DIEZ", "VEINTE", "TREINTA", "CUARENTA", "CINCUENTA", "SESENTA", "SETENTA", "OCHENTA", "NOVENTA")
    laCentenas = Array("CIENTO", "DOSCIENTOS", "TRESCIENTOS", "CUATROCIENTOS", "QUINIENTOS", "SEISCIENTOS", "SETECIENTOS", "OCHOCIENTOS", "NOVECIENTOS")

    lnNumeroBloques = 1

    Do
        lnPrimerDigito = 0
        lnSegundoDigito = 0
        lnTercerDigito = 0
        lcBloque = ""
        lnBloqueCero = 0

        For I = 1 To 3
            lnDigito = lyCantidad Mod 10

            If lnDigito <> 0 Then
                Select Case I
                    Case 1
                        lcBloque = " " & laUnidades(lnDigito - 1)
                        lnPrimerDigito = lnDigito
                    Case 2
                        If lnDigito <= 2 Then
                            lcBloque = " " & laUnidades((lnDigito * 10) + lnPrimerDigito - 1)
                        Else
                            lcBloque = " " & laDecenas(lnDigito - 1) & IIf(lnPrimerDigito <> 0, " Y", Null) & lcBloque
                        End If
                        lnSegundoDigito = lnDigito
                    Case 3

                    lcBloque = " " & IIf(lnDigito = 1 And lnPrimerDigito = 0 And lnSegundoDigito = 0, "CIEN", laCentenas(lnDigito - 1)) & lcBloque
                    lnTercerDigito = lnDigito
                End Select
            Else
                lnBloqueCero = lnBloqueCero + 1
            End If

            lyCantidad = Int(lyCantidad / 10)

            If lyCantidad = 0 Then
                Exit For
            End If
        Next I

        Select Case lnNumeroBloques
            Case 1
            CONVERTIRNUMEROLETRA = lcBloque
            Case 2
            CONVERTIRNUMEROLETRA = lcBloque & IIf(lnBloqueCero = 3, Null, " MIL") & CONVERTIRNUMEROLETRA
            Case 3 
                CONVERTIRNUMEROLETRA = lcBloque & IIf(lnPrimerDigito = 1 And lnSegundoDigito = 0 And lnTercerDigito = 0, " MILLON", " MILLONES") & CONVERTIRNUMEROLETRA
        End Select
        lnNumeroBloques = lnNumeroBloques + 1

    Loop Until lyCantidad = 0
    'Este es el resultado final en pantalla del texto convertido a número
    CONVERTIRNUMEROLETRA = "SON: (" & CONVERTIRNUMEROLETRA & IIf(tyCantidad > 1, " PESOS ", " PESO ") & Format(Str(lyCentavos), "00") & "/100 M.N.)"
End Function