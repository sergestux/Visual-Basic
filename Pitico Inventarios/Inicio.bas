Attribute VB_Name = "Inicio"
Option Explicit
Option Compare Text

Public CONTEO As Integer        'El num. de conteo actual 1,2,3
Public TIENDA As String         'Clave de la tienda Pitico 1= '001'
Public UBICACION As String      'Ubicacion del producto (Rack)
Public INVENTARIO As String     'Clave del Inventario actual
Public Conn As New ADODB.Connection    'La conexion de la base de datos

'/*******Para la funcion de Numeros a Letras***************/
Dim Unidades$(9), Decenas$(9), Oncenas$(9)
Dim Veintes$(9), centenas$(9)

'/*************Funcion Principal************************/
Sub Main()
Dim MsConexion As String
Dim TABLA As New ADODB.Recordset    'La tabla donde se haran los trabajos
    'Cadena de conexion
    MsConexion = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\INVENTARIOS.mdb;Mode=ReadWrite;Persist Security Info=False"
    DE1.Conn.ConnectionString = MsConexion
    'Paso el cursor del lado del cliente
    Conn.CursorLocation = adUseClient
    'Si la conexion esta abierta la cierro
    If Conn.State = adStateOpen Then
        Conn.Close
    Else
        Conn.Open (MsConexion)  'Abro la base de datos
    End If
    
    Set TABLA = Conn.Execute("SELECT ERROR FROM CONTROL")
    If TABLA.Fields(0).Value = False Then
        'MsgBox "SIN ERRORES"
        TABLA.Close     'PONGO LA VARIABLE DE ERROR A VERDADERO
        Set TABLA = Conn.Execute("UPDATE CONTROL SET ERROR= YES")
    Else
        'MsgBox "La aplicacion termino con errores", vbCritical
    End If
            
    FrmPrincipal.Show
    
End Sub
'Imprime el mensaje que se le pase en la barra de estado en el formulario Inicial
Sub IMPRIME(MENSAJE As String)
    Beep
    FrmPrincipal.Barra.Panels(1).Text = MENSAJE
End Sub

'Convierte el numero que se recibe a su correspondiente en letras (PESOS)
'Ejemplo: MsgBox NumLet$(Format(10, "#########0.00"))
'Imprime: DIEZ PESOS 00/100 M.N.
Public Function NumLet$(NUM#)
Dim dec$, MILM$, MILL$, MILE$, UNID$
    ReDim SALI$(11)
    Dim var$, i%, AUX$
    'NUM# = Round(NUM#, 2)
    var$ = Trim$(Str$(NUM#))
    If InStr(var$, ".") = 0 Then
            var$ = var$ + ".00"
    End If
    If InStr(var$, ".") = Len(var$) - 1 Then
         var$ = var$ + "0"
    End If
    var$ = String$(15 - Len(LTrim$(var$)), "0") + LTrim$(var$)
    dec$ = Mid$(var$, 14, 2)
    MILM$ = Mid$(var$, 1, 3)
    MILL$ = Mid$(var$, 4, 3)
    MILE$ = Mid$(var$, 7, 3)
    UNID$ = Mid$(var$, 10, 3)
    For i% = 1 To 11: SALI$(i%) = " ": Next i%
    i% = 0
    Unidades$(1) = "UN    "
    Unidades$(2) = "DOS    "
    Unidades$(3) = "TRES   "
    Unidades$(4) = "CUATRO "
    Unidades$(5) = "CINCO  "
    Unidades$(6) = "SEIS   "
    Unidades$(7) = "SIETE  "
    Unidades$(8) = "OCHO   "
    Unidades$(9) = "NUEVE  "
    Decenas$(1) = "DIEZ      "
    Decenas$(2) = "VEINTE    "
    Decenas$(3) = "TREINTA "
    Decenas$(4) = "CUARENTA "
    Decenas$(5) = "CINCUENTA "
    Decenas$(6) = "SESENTA "
    Decenas$(7) = "SETENTA "
    Decenas$(8) = "OCHENTA "
    Decenas$(9) = "NOVENTA "
    Oncenas$(1) = "ONCE       "
    Oncenas$(2) = "DOCE       "
    Oncenas$(3) = "TRECE      "
    Oncenas$(4) = "CATORCE    "
    Oncenas$(5) = "QUINCE     "
    Oncenas$(6) = "DIECISEIS  "
    Oncenas$(7) = "DIECISIETE "
    Oncenas$(8) = "DIECIOCHO  "
    Oncenas$(9) = "DIECINUEVE "
    Veintes$(1) = "VEINTIUNA    "
    Veintes$(2) = "VEINTIDOS    "
    Veintes$(3) = "VEINTITRES   "
    Veintes$(4) = "VEINTICUATRO "
    Veintes$(5) = "VEINTICINCO  "
    Veintes$(6) = "VEINTISEIS   "
    Veintes$(7) = "VEINTISIETE  "
    Veintes$(8) = "VEINTIOCHO   "
    Veintes$(9) = "VEINTINUEVE  "
    centenas$(1) = "       CIENTO "
    centenas$(2) = "   DOSCIENTOS "
    centenas$(3) = "  TRESCIENTOS "
    centenas$(4) = "CUATROCIENTOS "
    centenas$(5) = "   QUINIENTOS "
    centenas$(6) = "  SEISCIENTOS "
    centenas$(7) = "  SETECIENTOS "
    centenas$(8) = "  OCHOCIENTOS "
    centenas$(9) = "  NOVECIENTOS "
    If NUM# > 999999999999.99 Then NumLet$ = " ": Exit Function
    If Val(MILM$) >= 1 Then
       SALI$(2) = " MIL ":  '** MILES DE MILLONES
            SALI$(4) = " MILLONES "
            If Val(MILM$) <> 1 Then
                    Unidades$(1) = "UN     "
                    Veintes$(1) = "VEINTIUN     "
                    SALI$(1) = Descifrar$(Val(MILM$))
            End If
    End If
    If Val(MILL$) >= 1 Then
            If Val(MILL$) < 2 Then
                    SALI$(3) = "UN ": '*** UN MILLON
                    If Trim$(SALI$(4)) <> "MILLONES" Then
                            SALI$(4) = " MILLON "
                    End If
            Else
                    SALI$(4) = " MILLONES ": '*** VARIOS MILLONES
                    Unidades$(1) = "UN     "
                    Veintes$(1) = "VEINTIUN     "
                    SALI$(3) = Descifrar$(Val(MILL$))
            End If
    End If
    For i% = 2 To 9
            centenas$(i%) = Mid$(centenas(i%), 1, 11) + "OS"
    Next i%
    If Val(MILE$) > 0 Then
       SALI$(6) = " MIL ":   '*** MILES
            If Val(MILE$) <> 1 Then
                    SALI$(5) = Descifrar$(Val(MILE$))
            End If
    End If
    Unidades$(1) = "UN    "
    Veintes$(1) = "VEINTIUN"
    If Val(UNID$) >= 1 Then
            SALI$(7) = Descifrar$(Val(UNID$)):  '*** CIENTOS
            If Val(dec$) >= 10 Then
               ' SALI$(8) = " CON ": '*** DECIMALES
               ' SALI$(10) = Descifrar$(Val(DEC$))
               'MsgBox DEC$
            End If
    End If
    If Val(MILM$) = 0 And Val(MILL$) = 0 And Val(MILE$) = 0 And Val(UNID$) = 0 Then SALI$(7) = " CERO "
    AUX$ = ""
    For i% = 1 To 11
    AUX$ = AUX$ + SALI$(i%)
    Next i%
    NumLet$ = AUX$
    NumLet$ = Trim$(AUX$) & " PESOS " & dec$ & "/100 M.N."
    'NumLet$ = "(  " & Trim$(AUX$) & " PESOS " & dec$ & "/ 100 M.N.  )"
End Function
'Funcion Auxiliar de la funcion: NumLet$(NUM#)
Function Descifrar$(numero%)
    Static SAL$(4)
    Dim i%, CT As Double, DC As Double, DU As Double, UD  As Double
    Dim VARIABLE$
    For i% = 1 To 4: SAL$(i%) = " ": Next i%
    VARIABLE$ = String$(3 - Len(Trim$(Str$(numero%))), "0") + Trim$(Str$(numero%))
    CT = Val(Mid$(VARIABLE$, 1, 1)): '*** CENTENA
    DC = Val(Mid$(VARIABLE$, 2, 1)): '*** DECENA
    DU = Val(Mid$(VARIABLE$, 2, 2)): '*** DECENA + UNIDAD
    UD = Val(Mid$(VARIABLE$, 3, 1)): '*** UNIDAD
    If numero% = 100 Then
            SAL$(1) = "CIEN "
    Else
            If CT <> 0 Then SAL$(1) = centenas$(CT)
            If DC <> 0 Then
                    If DU <> 10 And DU <> 20 Then
                            If DC = 1 Then SAL$(2) = Oncenas$(UD): Descifrar$ = Trim$(SAL$(1) + " " + SAL$(2)): Exit Function
                            If DC = 2 Then SAL$(2) = Veintes$(UD): Descifrar$ = Trim$(SAL$(1) + " " + SAL$(2)): Exit Function
                    End If
                    SAL$(2) = " " + Decenas$(DC)
                    If UD <> 0 Then SAL$(3) = "Y "
            End If
            If UD <> 0 Then SAL$(4) = Unidades$(UD)
    End If
    Descifrar = Trim$(SAL$(1) + SAL$(2) + SAL$(3) + SAL$(4))
End Function

'Indica que solo se escriban números
Public Sub SoloNumeros(KeyAscii As Integer)
   If KeyAscii = 13 Then Exit Sub
   If Chr(KeyAscii) = "." Then
   ElseIf KeyAscii <> 8 Then
      If Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9" Then
          KeyAscii = 0
          IMPRIME "Solo Introduzca Numeros"
      End If
   End If
End Sub
