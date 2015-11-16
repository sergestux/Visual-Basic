VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sistema Integral de Información  v.2"
   ClientHeight    =   1920
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   6360
   Icon            =   "frmLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1132.958
   ScaleMode       =   0  'User
   ScaleWidth      =   5969.473
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkPortatil 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   9
      Top             =   1661
      Width           =   255
   End
   Begin VB.CheckBox chkOtrasuc 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   8
      Top             =   1661
      Width           =   255
   End
   Begin VB.TextBox txtUsuario 
      Height          =   315
      Left            =   2640
      TabIndex        =   1
      Text            =   "SERGIO"
      Top             =   600
      Width           =   1605
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   600
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1080
      Width           =   1140
   End
   Begin VB.TextBox txtContraseña 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2640
      PasswordChar    =   "*"
      TabIndex        =   2
      Text            =   "SERGIO"
      Top             =   1080
      Width           =   1605
   End
   Begin VB.ComboBox cmbServer 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmLogin.frx":0442
      Left            =   2640
      List            =   "frmLogin.frx":0444
      TabIndex        =   0
      Text            =   "cmbServer"
      Top             =   120
      Width           =   3495
   End
   Begin MSComctlLib.StatusBar stbmen 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   10
      Top             =   1620
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   6862
            MinWidth        =   6862
            Picture         =   "frmLogin.frx":0446
            Text            =   "Conexión a otras sucursales"
            TextSave        =   "Conexión a otras sucursales"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   4419
            MinWidth        =   4419
            Picture         =   "frmLogin.frx":05D8
            Text            =   "       Versión portátil"
            TextSave        =   "       Versión portátil"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblserver 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Servidor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   1320
      TabIndex        =   7
      Top             =   165
      Width           =   975
   End
   Begin VB.Image Img1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Left            =   240
      Picture         =   "frmLogin.frx":076A
      Stretch         =   -1  'True
      Top             =   240
      Width           =   735
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Usuario:"
      Height          =   270
      Index           =   0
      Left            =   1320
      TabIndex        =   5
      Top             =   600
      Width           =   720
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "C&ontraseña:"
      Height          =   270
      Index           =   1
      Left            =   1320
      TabIndex        =   6
      Top             =   1080
      Width           =   960
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private prueba As Boolean

Private Sub chkOtrasuc_Click()
   lblserver.Caption = IIf(chkOtrasuc.Value = 1, "Sucursal", "Servidor")
   Me.stbmen.Panels(1).Text = IIf(chkOtrasuc.Value = 1, "Conexión a otras sucursales", "Conexión en la sucursal")
   'chkOtrasuc.Value = IIf(chkOtrasuc.Value = 0, 1, 0)
   cmbServer.Enabled = chkOtrasuc.Value = 1
End Sub

Private Sub chkPortatil_Click()
If chkPortatil.Value = 1 Then
   cmbServer.Enabled = True
   cmbServer.Clear
   cmbServer.AddItem "MIGUEL CABRERA  16"
   cmbServer.AddItem "PUERTO ESCONDIDO  55"
   cmbServer.AddItem "MIAHUATLAN  26"
   cmbServer.AddItem "ISTMO  28"
   cmbServer.AddItem "PITICO-13  13"
   cmbServer.Text = cmbServer.List(0)
Else
   cmbServer.Text = SERVIDOR
End If
End Sub


Private Sub cmdCancel_Click()
  'Open "LPT1" For Output As #1
  'Print #1, Chr$(27); Chr$(36); Chr$(27); Chr$(157); Chr$(27); Chr$(68); Chr$(9); Chr$(27); Chr$(66); Chr$(11); Chr$(27); Chr$(74)
  'Print #1, Chr$(27) + Chr$(36) + Chr$(27) + Chr$(157) + Chr$(27) + Chr$(68) + Chr$(9) + Chr$(27) + Chr$(66) + Chr$(11) + Chr$(27) + Chr$(74)
  'Print #1, Chr$(9) + Chr$(27) + Chr$(66) + Chr$(11)
  'Print #1, Chr(27) + Chr(64)
  'Print #1, Chr(27) + Chr(64)
  'Print #1, Chr(11)
  'Printer.PaperSize = vbPRPSStatement
  'Printer.NewPage
  'Printer.CurrentY = 0
  'Printer.Print "---------------------------"
  'Printer.EndDoc
  'Print #1, Chr(27) + Chr(66) + "0" + "0"   'Avanza lineas n
  'Print #1, Chr(27) + Chr(40) + Chr(86)
  'Close #1
  Unload Me
End Sub

Private Sub obtentienda()
On Error GoTo Error:
'DE UN ARCHIVO INI SE DEBERAN LEER
Open App.Path & "\BODEGA.INI" For Input As #1
Line Input #1, clavex           '[STORE=22]
Line Input #1, sucursal         '[MIGUEL CABRERA]
Line Input #1, Direccion        '[MIGUEL CABRERA 603, COL CENTRO]
Line Input #1, User             '[KARINA]
Line Input #1, maquina          '[MACHINE=A1]
Line Input #1, SERVIDOR         'SERGIO
Line Input #1, ZONA             'CHS
Line Input #1, RutPort
'ZONA = "CHS"
RETIRO = re
TOTALISIMO = 0
clavesucursal = Mid(clavex, 8, 2)
nomsucursal = Mid(sucursal, 2, Len(sucursal) - 2)
dirsucursal = Mid(Direccion, 2, Len(Direccion) - 2)
USUARIO = Mid(User, 2, Len(User) - 2)
Caja = Mid(maquina, 10, 2)
Close #1
Exit Sub
Error:
  MsgBox "No se pudo Detectar la Configuracion del Punto de Venta, Por favor Consulte al Administrador", vbCritical
End Sub

Private Sub cmdOK_Click()
Dim rstUsu As ADODB.Recordset
Dim cServer As String
On Error GoTo Error
 PORCOSTO = False
 txtusuario.Text = Trim(txtusuario.Text)
 txtusuario.Refresh
 txtContraseña.Text = Trim(txtContraseña.Text)
 txtContraseña.Refresh
 Caja = computadora()
 'SERVIDOR = "EMAGDIEL"
 cServer = SERVIDOR
 Sql = (chkOtrasuc.Value = 0 And chkPortatil.Value = 0)
 If Not Sql Then   'Se conecta a base de datos en Access
    cvesuc = Trim(Mid(Me.cmbServer.Text, Len(cmbServer.Text) - 2))
    If chkPortatil.Value = 1 Then
       cCadConex = "DSN=PITICOMDB;DBQ=" & App.Path & "\PITICO" & cvesuc & ".mdb;DefaultDir=" & App.Path & ";DriverId=25;FIL=MS Access;MaxBufferSize=2048;PageTimeout=5;PWD=PORTATIL;UID=admin;"
       ccaption = "SIO - PORTATIL" & Space(15) & "SUCURSAL: [" + cmbServer.Text + "]" & Space(15) & "COMPUTADORA: [" & Caja & "]"
       tipotienda = 4
    Else
       tipotienda = 1
       cCadConex = "DSN=PITICOMDB;DBQ=P:\PITICO\PITICO" & cvesuc & "\PITICO" & cvesuc & ".mdb;DefaultDir=P:\PITICO\PITICO" & cvesuc & ";DriverId=25;FIL=MS Access;MaxBufferSize=2048;PageTimeout=5;PWD=PORTATIL;UID=admin;"
       'cCadConex = "DSN=PITICOMDB;DBQ=c:\PITICO" & cvesuc & ".mdb;DefaultDir=c:\;DriverId=25;FIL=MS Access;MaxBufferSize=2048;PageTimeout=5;PWD=PORTATIL;UID=admin;"

       ccaption = "CONEXION A OTRA BASE DE DATOS      SUCURSAL: [" + cmbServer.Text + "]   " & "COMPUTADORA: " & Caja
       
    End If
 Else
    cCadrpt = "DSN=PITICO;SERVER=" & cServer & ";PWD=" & txtContraseña & ";UID=" & txtusuario & ";Initial Catalog=PITICO"
    cCadConex = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=PITICO;Data Source=" & cServer
 End If
 strconnect = cCadConex
 Set cn = New ADODB.Connection
 cn.ConnectionString = cCadConex
 cn.Open
 If Not Sql Then
    frmAreaRecibo.Show
    Exit Sub
 End If
 cUsuario = txtusuario.Text:  cContraseña = txtContraseña.Text
 Set rstUsu = New ADODB.Recordset
 rstUsu.ActiveConnection = cCadConex
 rstUsu.Open "SELECT * FROM [usuarios],[catTienda] WHERE usuarios.sucursal = catTienda.TiClave AND LOGIN = '" & cUsuario & "'", cCadConex, adOpenKeyset, adLockOptimistic, adCmdText
 If rstUsu.EOF Then
     MsgBox "El usuario " & cUsuario & " no esta autorizado en el sistema", vbCritical
     Me.txtusuario.SetFocus
     Exit Sub
 ElseIf Trim(rstUsu!pass) <> Trim(txtContraseña.Text) Or Trim(rstUsu!pass) = "" Then
     MsgBox "LA CONTRASEÑA DE ACCESO ES INCORRECTA", vbCritical
     txtContraseña.SetFocus
     Exit Sub
 End If
 compInt = IIf(Trim(rstUsu!LEVEL1) = "I", " AND Interno = 1", " AND Interno = 0")

 cSucursal = rstUsu!sucursal + Space(2) + rstUsu!tidescrip
 'cSucursal = "28" + Space(2) + rstUsu!tidescrip     'Facturar lo del DIF en la J2 porque la B tiene Tuxtepec
 cCveDesUsu = rstUsu!clave + Space(3) + rstUsu!Name
 
 If Trim(rstUsu!tizona) = "O" Then      'OFICINAS
    tipotienda = 1
 ElseIf Trim(rstUsu!tizona) = "B" Then  'BODEGA CARBONERA
    tipotienda = 2
 ElseIf Trim(rstUsu!tizona) = "M" Then  'BODEGA DE MAYOREO
     tipotienda = 4
 Else                                   'TIENDAS NORMALES
    tipotienda = 3
 End If
 tipotienda = 4
 ccaption = "COMPUTADORA: [" & Caja & "]" & Space(10) & "USUARIO: [" + LCase(rstUsu!Name) + "]" + Space(10) + "SUCURSAL: [" + LCase(rstUsu!tidescrip) + "]"
 If InStr(1, "CT", rstUsu!LEVEL1) = 0 Then
   ModVta = True
 Else
   ModVta = False
 End If
 Hide
 Nivel = rstUsu!LEVEL1
 frmAreaRecibo.Caption = ccaption
 frmAreaRecibo.Show
 If Sql Then Call corteinv
 If Trim(rstUsu!sucursal) = "16" Then VeOfertas  'Solo se hace en Mig. Cab.
 'frmDespensa.Show
 Exit Sub
Error:
  MsgBox Err.Description
End Sub

Private Sub VeOfertas()
Dim RST As ADODB.Recordset
Dim fecha As Date
On Error GoTo Error:
Set RST = New ADODB.Recordset
fecha = date
RST.Open "SELECT producto FROM preciostemp WHERE activado = 0 and fechaini = '" & fecha + 1 & "'", cn, adOpenForwardOnly, adLockOptimistic, adCmdText
If Not (RST.BOF And RST.EOF) Then MsgBox "EXISTEN OFERTAS QUE INICIAN EL DIA DE MAÑANA Y AUN NO SE HAN ACTIVADO", vbInformation, "Recordatorio"
RST.Close
RST.Open "SELECT producto FROM preciostemp WHERE fechafin = '" & date + 1 & "'", cn, adOpenForwardOnly, adLockOptimistic, adCmdText
If Not (RST.BOF And RST.EOF) Then MsgBox "EXISTEN OFERTAS QUE TERMINAN EL DIA DE MAÑANA Y AUN NO SE HAN DESACTIVADO", vbInformation, "Recordatorio"
RST.Close
Set RST = Nothing
Exit Sub
Error:
   MsgBox Err.Description
End Sub

Private Sub corteinv()
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dia = "dia" & Day(date)
rs.Open "SELECT SUM(" & Dia & ") AS totdia FROM INVCORTE WHERE MES = " & Month(date), cn, adOpenForwardOnly, adLockOptimistic, adCmdText
If rs!TOTDIA = 0 Or IsNull(rs!TOTDIA) Then
   rs.Close
   rs.Open "SELECT * FROM invcorte WHERE mes =" & Month(date), cn, adOpenForwardOnly, adLockOptimistic, adCmdText
   If rs.BOF And rs.EOF Then
      cn.Execute "INSERT INTO invcorte(producto,mes) SELECT inprod, Mes = " & Month(date) & " FROM inventario"
   End If
   cn.Execute "UPDATE invcorte SET " & Dia & "= incant," & Dia & "p = incantpza FROM inventario WHERE inprod = producto AND mes =" & Month(date)
   MsgBox "EL CORTE DE INVENTARIO SE REALIZO CORRECTAMENTE", vbInformation
End If
rs.Close
Set rs = Nothing
End Sub

Private Sub Form_Activate()
'prueba = True
If prueba Then
   txtusuario.Text = "DBA"
   txtContraseña.Text = "MAY7503"
   cmdOK_Click
End If
End Sub

Private Sub Form_Load()
Dim MSG As String
cmbServer.AddItem "OFICINAS CENTRALES 10"
cmbServer.AddItem "TUXTEPEC 17"
Call obtentienda
cmbServer.Text = UCase(SERVIDOR)
End Sub

Private Sub stbmen_PanelClick(ByVal Panel As MSComctlLib.Panel)
Select Case Panel.Index
Case 1
   lblserver.Caption = IIf(chkOtrasuc.Value = 0, "Sucursal", "Servidor")
   stbmen.Panels(1).Text = IIf(chkOtrasuc.Value = 0, "Conexión a otras sucursales", "Conexión en la sucursal")
   chkOtrasuc.Value = IIf(chkOtrasuc.Value = 0, 1, 0)
   cmbServer.Enabled = (chkOtrasuc.Value = 1)
   cmbServer.ListIndex = 0
Case 2
   chkPortatil.Value = IIf(chkPortatil.Value = 0, 1, 0)
End Select
End Sub

Private Sub txtContraseña_GotFocus()
  txtContraseña.SelStart = 0
  txtContraseña.SelLength = Len(txtContraseña.Text)
End Sub

Private Sub txtContraseña_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys vbTab
   KeyAscii = 0
End If
End Sub

Private Sub txtUsuario_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys vbTab
   KeyAscii = 0
End If
End Sub
