Attribute VB_Name = "ModGeneral"
'Tipos de datos y funciones necesarias para descomprimir
Private Type CBChar
  ch(4096) As Byte
End Type

Private Type UNZIPUSERFUNCTION
    UNZIPPrntFunction As Long
    UNZIPSndFunction As Long
    UNZIPReplaceFunction  As Long
    UNZIPPassword As Long
    UNZIPMessage  As Long
    UNZIPService  As Long
    TotalSizeComp As Long
    TotalSize As Long
    CompFactor As Long
    NumFiles As Long
    Comment As Integer
End Type

Private Type UNZIPOPTIONS
    ExtractOnlyNewer  As Long
    SpaceToUnderScore As Long
    PromptToOverwrite As Long
    fQuiet As Long
    ncflag As Long
    ntflag As Long
    nvflag As Long
    nUflag As Long
    nzflag As Long
    ndflag As Long
    noflag As Long
    naflag As Long
    nZIflag As Long
    C_flag As Long
    fPrivilege As Long
    Zip As String
    extractdir As String
End Type

Public Type ZIPnames
    s(0 To 99) As String
End Type

'Tipos de datos y funciones necesarias para comprimir
Public Type ZIPUSERFUNCTIONS
DLLPrnt As Long
DLLPassword As Long
DLLComment As Long
DLLService As Long
End Type

Public Type ZPOPT
fSuffix As Long
fEncrypt As Long
fSystem As Long
fVolume As Long
fExtra As Long
fNoDirEntries As Long
fExcludeDate As Long
fIncludeDate As Long
fVerbose As Long
fQuiet As Long
fCRLF_LF As Long
fLF_CRLF As Long
fJunkDir As Long
fRecurse As Long
fGrow As Long
fForce As Long
fMove As Long
fDeleteEntries As Long
fUpdate As Long
fFreshen As Long
fJunkSFX As Long
fLatestTime As Long
fComment As Long
fOffsets As Long
fPrivilege As Long
fEncryption As Long
fRepair As Long
flevel As Byte
date As String
szRootDir As String
End Type

'Funciónes para comprimir archivos ZIP
Public Declare Function ZpInit Lib "zip32.dll" (ByRef Zipfun As ZIPUSERFUNCTIONS) As Long
Public Declare Function ZpSetOptions Lib "zip32.dll" (ByRef Opts As ZPOPT) As Long
Public Declare Function ZpArchive Lib "zip32.dll" (ByVal argc As Long, ByVal funame As String, ByRef argv As ZIPnames) As Long

'Función para Descomprimir archivos ZIP
Public Declare Function Wiz_SingleEntryUnzip Lib "unzip32.dll" (ByVal ifnc As Long, ByRef ifnv As ZIPnames, ByVal xfnc As Long, ByRef xfnv As ZIPnames, dcll As UNZIPOPTIONS, Userf As UNZIPUSERFUNCTION) As Long

'Función que sirve para desplegar automaticamente un Combo
Public Declare Function SendMessageLong Lib "user32" Alias _
"SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
ByVal wParam As Long, ByVal lParam As Long) As Long

'Función que sustituye el sendkeys por ocasionar algunos problemas
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal _
bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

'Función que obtiene el nombre del ordenador dentro de la red
Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'Funcion para que la lista desplegable del combo sea mas grande que el combo
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long


Public tipotienda As Integer
Public lpprod As Boolean  'Que buscan productos
Public lpprov As Boolean
Public strcveprod As String
Public manual  As Boolean
Public CLAVEINVENTARIO As Integer ' LA CLAVE DE LA TIENDA EN DONDE SE TRABAJA
Public TIPOEXP As String
Public Const Ind = "D"  'Proved. Indirectos los que surten en bodega
Public Const Ins = "S"  'Proved. Instantaneo en el momento confirma y entrega el pedido en la tienda
Public Const Abi = "I"  'Proved. Abierto Confirma en la tienda y dias despues entrega el pedido

Public PORCOSTO As Boolean
Public cn As Connection    ' Conexion a sql Server
Public cCadConex As String ' Cadena de conexion
Public strconnect As String
Public cCadrpt As String
Public strcveprov As String
Public nOp As Integer      ' Número de opción a la que pasa Altas o modificaciones
Public paq As Integer
Public Sql As Boolean
Public cModo As String     'Modo en que se encuentra la forma ya que se utiliza para CAPTURA, CONFIRMACION y RECEPCION de pedidos.
Public cUsuario As String
Public cContraseña As String
Public cSucursal As String  'Sucursal en la que trabaja el sistema
Public Nivel As String
Public lDatAJu As String   'Para ajuste de inventario (frmModInv)
Public lProvAbto As String  'Tipo de proveedor Instantaneo, Abierto o Indirecto
Public ccaption As String   'Encabezado de la forma
Public cCveDesUsu As String 'Clave y Descripcion del usuario
Public MODIFICADO As Boolean
Public Forma As Integer     'Para sabe de que opcion se llama la forma frmCaptPed (Captura del pedido) sugerido, abastecimiento
Public TRASLADOSENV(1 To 30) As String
Public sucut As Integer
Public PORMERMAS As Boolean
Public PORAUTO As Boolean
Public SERVIDOR As String
Public lNvoPedprove As Boolean
Public puedegrabar As Boolean   'Función de precios
Public Caja  'Caja que realiza el cobro
Public ModVta 'Modulo realiza la venta
Public Preventa As Integer
'Public lDatAJu As Boolean   'Para ajuste de inventario (frmModInv)
Public SoloAct As Boolean   'Filtro para mostrar todos los productos o solo los activos
Public compInt  As String
Public RutPort As String
Public centro As Integer
Private consulta As String
Public ZONA As String


'PARA LA IMPRESION DE FACTURAS
Public nAlto As Double
Public NVENTA1 As Double
Public NVENTA2 As Double
Public NVENTA3 As Double
Public NVENTA4 As Double
Public NVENTA5 As Double
Public NVENTA6 As Double
Public NVENTA7 As Double
Public NVENTA8 As Double
Public NIVA1 As Double
Public NIVA2 As Double
Public NIVA3 As Double
Public NIVA4 As Double
Public NIVA5 As Double
Public NIVA6 As Double
Public NIVA7 As Double
Public NIVA8 As Double
Public NIEPS1 As Double
Public NIEPS2 As Double
Public NIEPS3 As Double
Public NIEPS4 As Double
Public NIEPS5 As Double
Public NIEPS6 As Double
Public NIEPS7 As Double
Public NIEPS8 As Double
Public rfct As String
Public nomt As String
Public dirt As String
Public telr As String
Public colt As String
Public ciut As String
Public PASSMASTER As String
Public noventar As Long
Public sert As String
Public SERIE As String
Public RFCFINAL As String
Public numfac As String
Public consecs(1 To 400) As String
Public descripcs(1 To 400) As String
Public medidas(1 To 400) As String
Public cantidads(1 To 400) As Double
Public cantidadps(1 To 400) As Double
Public costoss(1 To 400) As Double
Public costosp(1 To 400) As Double
Public precios(1 To 400) As Double
Public preciosp(1 To 400) As Double
Public tasas(1 To 400) As Double
Public ivas(1 To 400) As Double
Public iepss(1 To 400) As Double
Public importes(1 To 400) As Double
Public NOVENTAT(1 To 400) As Long
Public ncosto As Double
Public mes(11) As Variant

Public Function tasa(iva As Integer, ieps As Integer) As String
If iva = 0 And ieps = 0 Then
   tasa = "*E"
ElseIf iva = 15 And ieps = 0 Then
   tasa = "*I"
ElseIf iva = 15 And ieps = 15 Then
   tasa = "*IE"
ElseIf iva = 15 And ieps > 25 Then
   tasa = "*IE"
End If
End Function

'Procedimiento que pone informacion del detalle del pedido
'Numero de productos, cajas y piezas.
Sub poninfo()
On Error Resume Next
  Set rsttemp = New ADODB.Recordset
  rsttemp.Open "SELECT COUNT(*) As NumPro, SUM(df_cantsol) As TotCaj , SUM(df_cantsolp) As TotPza FROM DETALLEFACTURA WHERE df_pedido = '" & frmCaptPed.txtcampos(0).Text & "' AND df_sugerido = " & IIf(frmCaptPed.txtcampos(1).Text = "ABA", 0, 1), cn, adOpenKeyset, adLockOptimistic, adCmdText
  frmCaptPed.lblInfo.Caption = "Productos: "
  If IsNull(rsttemp!Numpro) Then
     frmCaptPed.lblInfo.Caption = frmCaptPed.lblInfo.Caption & "0"
  Else
     frmCaptPed.lblInfo.Caption = frmCaptPed.lblInfo.Caption & CStr(rsttemp!Numpro)
  End If
  frmCaptPed.lblInfo.Caption = frmCaptPed.lblInfo.Caption & Space(3) & "Cajas Sol: "
  If IsNull(rsttemp!totcaj) Then
     frmCaptPed.lblInfo.Caption = frmCaptPed.lblInfo.Caption & "0"
  Else
     frmCaptPed.lblInfo.Caption = frmCaptPed.lblInfo.Caption & CStr(rsttemp!totcaj)
  End If
  frmCaptPed.lblInfo.Caption = frmCaptPed.lblInfo.Caption & Space(3) & "Piezas Sol: "
  If IsNull(rsttemp!totpza) Then
     frmCaptPed.lblInfo.Caption = frmCaptPed.lblInfo.Caption & "0"
  Else
     frmCaptPed.lblInfo.Caption = frmCaptPed.lblInfo.Caption & CStr(rsttemp!totpza)
  End If
  frmCaptPed.lblInfo.Refresh
End Sub

'Funcion utilizada para determinar la clave de la sucursal en base al archivo especificado
'en la opcion de importar pedido
Public Function Pedsuc(Archivo As String) As String
Dim cvesuc As String

Select Case UCase(Mid(Archivo, 4, 3))
    Case "REF"
         cvesuc = "1"
    Case "BRE"
         cvesuc = "8"
    Case "CON"
         cvesuc = "78"
    Case "DOL"
         cvesuc = "22"
    Case "HID"
         cvesuc = "4"
    Case "IND"
         cvesuc = "5"
    Case "MCT"
         cvesuc = "12"
    Case "MCM"
         cvesuc = "16"
    Case "PER"
         cvesuc = "2"
    Case "ROS"
         cvesuc = "21"
    Case "TLA"
         cvesuc = "14"
    Case "SAN"
         cvesuc = "20"
    Case "CAR"
         cvesuc = "3"
    Case "ATZ"
         cvesuc = "13"
    Case "CIN"
         cvesuc = "7"
    Case "MAD"
         cvesuc = "9"
    Case "MIA"
         cvesuc = "26"
    Case "POR"
         cvesuc = "11"
    Case "TAP"
         cvesuc = "18"
    Case "TUX"
         cvesuc = "17"
    Case "TUT"
         cvesuc = "19"
    Case "OFI"
         cvesuc = "10"
    Case "ZIM"
       cvesuc = "27"
    Case "MMM"
       cvesuc = "30"
    Case "PDZ"
        cvesuc = "31"
    Case "CEN"
        cvesuc = "23"
    Case "MIN"
        cvesuc = "79"
    Case "PTO"
        cvesuc = "55"
    Case "COS"
        cvesuc = "24"
    Case "PTO"
        cvesuc = "55"
    Case "IST"
        cvesuc = "28"
    Case Else
         cvesuc = ""
End Select
If cvesuc = "" Then
    MsgBox "EXISTE UNA TIENDA QUE NO TIENE REGISTRADA SU CLAVE EN EL SISTEMA" & Archivo & _
    "FAVOR DE INFORMAR AL ADMINISTRADOR DEL SISTEMA, YA QUE EL PROCESO QUE ESTA REALIZANDO NO SE EJECUTARA CORRECTAMENTE", vbCritical
End If
Pedsuc = cvesuc
End Function

Public Function verImpresora() As Boolean
Dim lImp  As Boolean
lImp = False
For Each x In Printers
     ' MsgBox x.DeviceName
   If x.DeviceName Like "*ZEBRA*" Then
      lImp = True
      Set Printer = x
      Exit For
   End If
   If x.DeviceName Like "*zebra*" Then
      lImp = True
      Set Printer = x
      Exit For
   End If
   If x.DeviceName Like "*Zebra*" Then
      lImp = True
      Set Printer = x
      Exit For
   End If
Next x
If lImp = False Then
   MsgBox "NO ES POSIBLE IMPRIMIR PORQUE NO EXISTE LA IMPRESORA DE CODIGOS" & Chr(13) & Chr(13) & "PORQUE NO EXISTEN IMPRESORAS INSTALADAS" & Chr(13) & "PARA ESTE TIPO DE ETIQUETAS", vbCritical
End If
verImpresora = lImp
End Function

Public Function verImpticket() As Boolean
Dim lImp  As Boolean
lImp = False
For Each x In Printers
   If UCase(x.DeviceName) Like "*TICKET*" Then
      MsgBox "PREPARE LA IMPRESORA " & x.DeviceName, vbInformation
      lImp = True
      Set Printer = x
      Exit For
   End If
Next x
If lImp = False Then
   MsgBox "NO ES POSIBLE IMPRIMIR PORQUE NO EXISTE LA IMPRESORA DE tikets" & Chr(13) & Chr(13) & "PORQUE NO EXISTEN IMPRESORAS INSTALADAS" & Chr(13) & "PARA ESTE TIPO DE ETIQUETAS", vbCritical
End If
verImpticket = lImp
End Function

Public Sub agregarfactura(Factura As String, SERIE As String, CLIENTE)
pr = InStr(CLIENTE, "[")
tt = Mid(CLIENTE, pr + 1, Len(CLIENTE))
pr = InStr(tt, "]")
tt = Mid(tt, 1, pr - 1)
CAD = "insert into facventa(NOVENTA,faccliente,facfecha,total,iva,ieps,numfactura,serie) values (" & _
      12 & "," & tt & ",'" & date & "',0,0,0,'" & Trim(Factura) & "','" & Trim(SERIE) & "')"
'MsgBox CAD
cn.Execute CAD
End Sub


Public Sub UnZip(Zip As String, extractdir As String)
On Error GoTo err_Unzip

Dim Resultado As Long
Dim intContadorFicheros As Integer

Dim FuncionesUnZip As UNZIPUSERFUNCTION
Dim OpcionesUnZip As UNZIPOPTIONS

Dim NombresFicherosZip As ZIPnames, NombresFicheros2Zip As ZIPnames

NombresFicherosZip.s(0) = vbNullChar
NombresFicheros2Zip.s(0) = vbNullChar
FuncionesUnZip.UNZIPMessage = 0&
FuncionesUnZip.UNZIPPassword = 0&
FuncionesUnZip.UNZIPPrntFunction = DevolverDireccionMemoria(AddressOf UNFuncionParaProcesarMensajes)
FuncionesUnZip.UNZIPReplaceFunction = DevolverDireccionMemoria(AddressOf UNFuncionReplaceOptions)
FuncionesUnZip.UNZIPService = 0&
FuncionesUnZip.UNZIPSndFunction = 0&
OpcionesUnZip.C_flag = 1
OpcionesUnZip.fQuiet = 2
OpcionesUnZip.noflag = 1
OpcionesUnZip.Zip = Zip
OpcionesUnZip.extractdir = extractdir

Resultado = Wiz_SingleEntryUnzip(0, NombresFicherosZip, 0, NombresFicheros2Zip, OpcionesUnZip, FuncionesUnZip)

Exit Sub
err_Unzip:
    MsgBox "Unzip: " + Err.Description, vbExclamation
    Err.Clear
End Sub

Private Function UNFuncionParaProcesarMensajes(ByRef fname As CBChar, ByVal x As Long) As Long
On Error GoTo err_UNFuncionParaProcesarMensajes

    UNFuncionParaProcesarMensajes = 0

Exit Function
err_UNFuncionParaProcesarMensajes:
    MsgBox "UNFuncionParaProcesarMensajes: " + Err.Description, vbExclamation
    Err.Clear
End Function

Private Function UNFuncionReplaceOptions(ByRef p As CBChar, ByVal L As Long, ByRef m As CBChar, ByRef Name As CBChar) As Integer
On Error GoTo err_UNFuncionReplaceOptions

    UNFuncionParaProcesarPassword = 0

Exit Function
err_UNFuncionReplaceOptions:
    MsgBox "UNFuncionParaProcesarPassword: " + Err.Description, vbExclamation
    Err.Clear
End Function

Public Function DevolverDireccionMemoria(Direccion As Long) As Long
On Error GoTo err_DevolverDireccionMemoria

    DevolverDireccionMemoria = Direccion

Exit Function
err_DevolverDireccionMemoria:
    MsgBox "DevolverDireccionMemoria: " + Err.Description, vbExclamation
    Err.Clear
End Function

'Comprimir archivos
Function FuncionParaProcesarPassword(ByRef B1 As Byte, L As Long, ByRef B2 As Byte, ByRef B3 As Byte) As Long
    FuncionParaProcesarPassword = 0
End Function
Function FuncionParaProcesarServicios(ByRef fname As CBChar, ByVal x As Long) As Long
    FuncionParaProcesarServicios = 0
End Function
Function FuncionParaProcesarMensajes(ByRef fname As CBChar, ByVal x As Long) As Long
    FuncionParaProcesarMensajes = 0
End Function
Function FuncionParaProcesarComentarios(Comentario As CBChar) As CBChar
    Comentario.ch(0) = vbNullString
    FuncionParaProcesarComentarios = Comentario
End Function

'Función para habilitar opciones
Public Function autoriza(permisos As String, opcion As Integer) As Boolean
On Error Resume Next
autoriza = False
If Mid(permisos, opcion, 1) = "1" Then
   autoriza = True
Else
   Select Case opcion
       Case 1:   MsgBox "EL USUARIO NO ESTA AUTORIZADO PARA MODIFICAR EL INVENTARIO", vbExclamation
       Case 2:   MsgBox "EL USUARIO NO ESTA AUTORIZADO PARA CANCELAR FACTURAS", vbExclamation
       Case 3:   MsgBox "EL USUARIO NO ESTA AUTORIZADO PARA MODIFICAR PRECIOS DE VENTA DE MOSTRADOR O PREVENTA", vbExclamation
       Case 8:   MsgBox "EL USUARIO NO ESTA AUTORIZADO PARA CAMBIAR ESCALA DE PRECIOS A CLIENTES AUTORIZADOS", vbExclamation
       Case 9:   MsgBox "EL USUARIO NO ESTA AUTORIZADO PARA CAMBIAR PRECIOS MAS ALTO DE LO NORMAL", vbExclamation
       Case 10:  MsgBox "EL USUARIO NO ESTA AUTORIZADO PARA CAMBIAR ESCALA DE PRECIOS A CUALQUIER CLIENTE", vbExclamation
       Case 11: MsgBox "EL USUARIO NO ESTA AUTORIZADO PARA AUTORIZAR VENTAS A CREDITO", vbExclamation
   End Select
End If
End Function

Public Function computadora()
Dim nPC As String
Dim buffer As String
Dim estado As Long
buffer = String$(255, " ")
estado = GetComputerName(buffer, 255)
If estado <> 0 Then
   nPC = Left(buffer, 255)
End If
computadora = LCase(Mid(nPC, 1, Len(Trim(nPC)) - 1))
End Function

