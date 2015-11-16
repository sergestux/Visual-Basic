VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCliente 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Catalogo de clientes"
   ClientHeight    =   6435
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   6300
   Icon            =   "frmCliente.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   6300
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Datos generales"
      ForeColor       =   &H80000002&
      Height          =   3855
      Left            =   240
      TabIndex        =   33
      Top             =   0
      Width           =   5895
      Begin VB.ComboBox cmbgiro 
         DataField       =   "giro"
         DataSource      =   "datPrimaryRS"
         Height          =   315
         ItemData        =   "frmCliente.frx":014A
         Left            =   3360
         List            =   "frmCliente.frx":0163
         TabIndex        =   58
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtfields 
         DataField       =   "apematerno"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   15
         Left            =   2040
         MaxLength       =   20
         TabIndex        =   3
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox txtfields 
         Alignment       =   2  'Center
         DataField       =   "nomnegocio"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   17
         Left            =   120
         MaxLength       =   30
         TabIndex        =   5
         Top             =   1680
         Width           =   5655
      End
      Begin VB.TextBox txtfields 
         DataField       =   "nombres"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   16
         Left            =   3960
         MaxLength       =   20
         TabIndex        =   4
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox txtfields 
         DataField       =   "apepaterno"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   14
         Left            =   120
         MaxLength       =   20
         TabIndex        =   2
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox txtfields 
         DataField       =   "ruta"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   13
         Left            =   1320
         MaxLength       =   5
         TabIndex        =   0
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtfields 
         DataField       =   "Cciudad"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   12
         Left            =   120
         MaxLength       =   50
         TabIndex        =   8
         Top             =   2880
         Width           =   1815
      End
      Begin VB.TextBox txtfields 
         DataField       =   "cColonia"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   11
         Left            =   3840
         MaxLength       =   50
         TabIndex        =   7
         Top             =   2280
         Width           =   1935
      End
      Begin VB.TextBox txtfields 
         Alignment       =   2  'Center
         DataField       =   "ctipo"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   10
         Left            =   2520
         TabIndex        =   1
         Text            =   "0"
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtfields 
         Alignment       =   2  'Center
         DataField       =   "cnombrefac"
         DataSource      =   "datPrimaryRS"
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   150
         TabIndex        =   11
         Top             =   3480
         Width           =   5655
      End
      Begin VB.TextBox txtfields 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         DataField       =   "cclave"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   0
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtfields 
         DataField       =   "cnombre"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   1
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   150
         TabIndex        =   13
         Top             =   3600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtfields 
         DataField       =   "cdireccion"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   2
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   150
         TabIndex        =   6
         Top             =   2280
         Width           =   3615
      End
      Begin VB.TextBox txtfields 
         DataField       =   "crfc"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   3
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   13
         TabIndex        =   9
         Top             =   2880
         Width           =   1815
      End
      Begin VB.TextBox txtfields 
         DataField       =   "ctelefono"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   4
         Left            =   4080
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   10
         Top             =   2880
         Width           =   1695
      End
      Begin VB.CheckBox chkModCre 
         Caption         =   "&Mod. credito"
         DataSource      =   "datPrimaryRS"
         Height          =   300
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "Giro"
         Height          =   255
         Index           =   17
         Left            =   3360
         TabIndex        =   57
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "Nombre del establecimiento"
         Height          =   255
         Index           =   16
         Left            =   120
         TabIndex        =   56
         Top             =   1440
         Width           =   5655
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "Nombres (s)"
         Height          =   255
         Index           =   15
         Left            =   3960
         TabIndex        =   55
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "Apellido Materno"
         Height          =   255
         Index           =   14
         Left            =   2040
         TabIndex        =   54
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "Apellido Paterno"
         Height          =   255
         Index           =   13
         Left            =   120
         TabIndex        =   53
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "Ruta"
         Height          =   255
         Index           =   12
         Left            =   1320
         TabIndex        =   52
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "Ciudad"
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   50
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "Localidad / Población"
         Height          =   255
         Index           =   10
         Left            =   3840
         TabIndex        =   49
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "Tipo"
         Height          =   255
         Index           =   9
         Left            =   2520
         TabIndex        =   48
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "Nombre  Facturación"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   39
         Top             =   3240
         Width           =   5655
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "Clave"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Caption         =   "Nombre Venta"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   37
         Top             =   3600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "Direc. fiscal"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   36
         Top             =   2040
         Width           =   3615
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "R.F.C."
         Height          =   255
         Index           =   3
         Left            =   2040
         TabIndex        =   35
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "Telefono:"
         Height          =   255
         Index           =   4
         Left            =   4080
         TabIndex        =   34
         Top             =   2640
         Width           =   1695
      End
   End
   Begin VB.Frame fraCon 
      BackColor       =   &H00C0C000&
      Caption         =   "Contraseña de acceso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1560
      TabIndex        =   40
      Top             =   3960
      Visible         =   0   'False
      Width           =   3615
      Begin VB.CommandButton cmdConCance 
         Caption         =   "&Cancelar"
         Height          =   255
         Left            =   2400
         TabIndex        =   43
         Top             =   840
         Width           =   930
      End
      Begin VB.CommandButton cmdConAceptar 
         Caption         =   "&Aceptar"
         Height          =   255
         Left            =   840
         TabIndex        =   42
         Top             =   840
         Width           =   1050
      End
      Begin VB.TextBox txtContra 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1440
         PasswordChar    =   "*"
         TabIndex        =   41
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblContra 
         BackColor       =   &H00C0C000&
         Caption         =   "Contraseña"
         Height          =   255
         Left            =   240
         TabIndex        =   44
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame fracredito 
      Caption         =   "Credito"
      ForeColor       =   &H80000002&
      Height          =   1335
      Left            =   240
      TabIndex        =   29
      Top             =   3960
      Visible         =   0   'False
      Width           =   5895
      Begin VB.CheckBox chkCambPre 
         Caption         =   "Cambiar escala de precios"
         DataField       =   "cambpre"
         DataSource      =   "datPrimaryRS"
         Height          =   375
         Left            =   3720
         TabIndex        =   51
         Top             =   120
         Width           =   1695
      End
      Begin VB.CheckBox chkPagocheque 
         Caption         =   "Paga con cheques"
         DataField       =   "cPagocheque"
         DataSource      =   "datPrimaryRS"
         Height          =   255
         Left            =   1800
         TabIndex        =   47
         Top             =   240
         Width           =   1695
      End
      Begin VB.CheckBox chkCredito 
         Caption         =   "Venta a credito"
         DataField       =   "cCredito"
         DataSource      =   "datPrimaryRS"
         Height          =   255
         Left            =   240
         TabIndex        =   46
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtfields 
         Alignment       =   2  'Center
         DataField       =   "cmontofianza"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   2
         EndProperty
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   9
         Left            =   1320
         TabIndex        =   16
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtfields 
         Alignment       =   2  'Center
         DataField       =   "cVigFianza"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/mm/yyYY "
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   8
         Left            =   3960
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   17
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtfields 
         Alignment       =   2  'Center
         DataField       =   "ctiempocredito"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   7
         Left            =   4440
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   15
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtfields 
         Alignment       =   2  'Center
         DataField       =   "climitecredito"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   2
         EndProperty
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   6
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   14
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label lbletiquetas 
         Caption         =   "Monto fianza"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   45
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblLabels 
         Caption         =   "Vig. Fianza"
         Height          =   255
         Index           =   8
         Left            =   2880
         TabIndex        =   32
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblLabels 
         Caption         =   "Plazo credito"
         Height          =   255
         Index           =   7
         Left            =   3360
         TabIndex        =   31
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblLabels 
         Caption         =   "Limite credito"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   30
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.PictureBox PicMod 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   6300
      TabIndex        =   28
      Top             =   5070
      Visible         =   0   'False
      Width           =   6300
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   300
         Left            =   3120
         TabIndex        =   19
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Grabar"
         Height          =   300
         Left            =   1680
         TabIndex        =   18
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   0
      ScaleHeight     =   540
      ScaleWidth      =   6300
      TabIndex        =   27
      Top             =   5565
      Width           =   6300
      Begin VB.CommandButton cmdBcaRfc 
         Caption         =   "&B. RFC"
         Height          =   300
         Left            =   3960
         TabIndex        =   25
         ToolTipText     =   "Busca clientes por RFC"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "&Buscar"
         Height          =   300
         Left            =   3120
         TabIndex        =   24
         ToolTipText     =   "Busca clientes por nombre"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdModificar 
         Caption         =   "&Modificar"
         Height          =   300
         Left            =   1200
         TabIndex        =   22
         ToolTipText     =   "Modifica los datos del cliente"
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Cerrar"
         Height          =   300
         Left            =   4800
         TabIndex        =   26
         ToolTipText     =   "Cerrar pantalla de clientes"
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Eliminar"
         Enabled         =   0   'False
         Height          =   300
         Left            =   2160
         TabIndex        =   23
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Ag&regar"
         Height          =   300
         Left            =   240
         TabIndex        =   21
         ToolTipText     =   "Agrega un nuevo cliente"
         Top             =   0
         Width           =   855
      End
   End
   Begin MSAdodcLib.Adodc datPrimaryRS 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   6105
      Width           =   6300
      _ExtentX        =   11113
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   " "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "frmCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private indice As Integer
Private Sub chkModCre_Click()
  fraCon.Visible = True
  txtContra.SetFocus
End Sub

Private Sub cmdBcaRfc_Click()
Dim cInd
CNOM = InputBox("TECLEE R.F.C. DEL CLIENTE A BUSCAR", "Buscar R.F.C.")
If Trim(CNOM) = "" Then Exit Sub
cInd = datPrimaryRS.Recordset.Bookmark
datPrimaryRS.Recordset.MoveFirst
datPrimaryRS.Recordset.Find "CRFC LIKE '" & Trim(CNOM) & "*'"
If datPrimaryRS.Recordset.EOF Then
   MsgBox "NO EXISTE EL RFC DEL CLIENTE ESPECIFICADO", vbExclamation
   datPrimaryRS.Recordset.Bookmark = cInd
End If

End Sub

Private Sub CmdBuscar_Click()
On Error GoTo eRROR:
Dim cInd
CNOM = InputBox("INTRODUZCA DATO A BUSCAR", "Buscar Cliente")
If Trim(CNOM) = "" Then Exit Sub
DATO = IIf(indice = 0, " CCLAVE = " & Trim(CNOM), " CNOMBRE LIKE '%" & Trim(CNOM) & "%'")
cInd = datPrimaryRS.Recordset.Bookmark
datPrimaryRS.Recordset.MoveFirst
datPrimaryRS.Recordset.Find DATO
If datPrimaryRS.Recordset.EOF Then
   MsgBox "NO EXISTE EL NOMBRE DEL CLIENTE ESPECIFICADO", vbExclamation
   datPrimaryRS.Recordset.Bookmark = cInd
End If
Exit Sub
eRROR:
    MsgBox Err.Description
End Sub

Private Sub cmdCancelar_Click()
Dim nMar
   On Error Resume Next
   nMar = datPrimaryRS.Recordset.Bookmark
   Me.datPrimaryRS.Refresh
   datPrimaryRS.Recordset.Bookmark = nMar
   Me.PicMod.Visible = False
   Me.picButtons.Visible = True
   For N = 1 To 17
     txtfields(N).Locked = True
   Next
End Sub

Private Sub cmdConAceptar_Click()

Dim RsCon As ADODB.Recordset
Set RsCon = New ADODB.Recordset
RsCon.Open "SELECT * FROM USUARIOS WHERE PASSMASTER = '" & Trim(txtContra.Text) & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
If Not autoriza(RsCon!permisos, 11) Then Exit Sub
   fraCon.Visible = False
   fracredito.Visible = True
   txtfields(6).SetFocus
'End If
End Sub

Private Sub cmdConCance_Click()
  fraCon.Visible = False
  chkCredito.Value = 0
End Sub

Private Sub cmdModificar_Click()
  Me.picButtons.Visible = False
  Me.PicMod.Visible = True
  For N = 1 To 17
     txtfields(N).Locked = False
  Next
End Sub

Private Sub Form_Load()
  datPrimaryRS.CursorType = adOpenKeyset
  datPrimaryRS.CommandType = adCmdText
  datPrimaryRS.ConnectionString = cCadConex
  If frmVentas.DbgrdPreventa.Visible = True Then
     datPrimaryRS.RecordSource = "SELECT * FROM CatCliente WHERE cclave = " & frmVentas.adopreventa.Recordset!clcliente & " Order By cNombre"
  ElseIf frmVentas.txtcampos(4).Enabled And Trim(frmVentas.txtcampos(4).Text) <> "" Then
     datPrimaryRS.RecordSource = "SELECT * FROM CatCliente WHERE cclave = " & frmVentas.txtcampos(4).Text & " Order By cNombre"
  Else
     datPrimaryRS.RecordSource = "SELECT * FROM CatCliente Order By cNombre"
  End If
  datPrimaryRS.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim RSTCLIE As ADODB.Recordset
' Set RSTCLIE = New ADODB.Recordset
' RSTCLIE.Open "CATCLIENTE", cn, adOpenDynamic, adLockOptimistic, adCmdTable
' frmVentas.cmbCliente.Clear
' While Not RSTCLIE.EOF
'     frmVentas.cmbCliente.AddItem RSTCLIE!cnombre
'     RSTCLIE.MoveNext
' Wend
datPrimaryRS.Recordset.Bookmark = datPrimaryRS.Recordset.Bookmark
frmVentas.cmbCliente.AddItem datPrimaryRS.Recordset!cNombre
'frmVentas.cmbCliente.SetFocus
End Sub

Private Sub datPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.eRROR, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error Resume Next
  'Esto mostrará la posición de registro actual para este Recordset
  datPrimaryRS.Caption = Space(25) & datPrimaryRS.Recordset!cNombre
  cmbgiro.Text = datPrimaryRS.Recordset!giro
End Sub

Private Sub datPrimaryRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Aquí se coloca el código de validación
  'Se llama a este evento cuando ocurre la siguiente acción
  Dim bCancel As Boolean

  Select Case adReason
  Case adRsnAddNew
  Case adRsnClose
  Case adRsnDelete
  Case adRsnFirstChange
  Case adRsnMove
  Case adRsnRequery
  Case adRsnResynch
  Case adRsnUndoAddNew
  Case adRsnUndoDelete
  Case adRsnUndoUpdate
  Case adRsnUpdate
  End Select

  If bCancel Then adStatus = adStatusCancel
End Sub

Private Sub cmdAdd_Click()
Dim rsttemp As ADODB.Recordset
  On Error GoTo AddErr
  If Not Sql Then
    Set rsttemp = New ADODB.Recordset
    rsttemp.Open "SELECT * FROM catcliente WHERE cclave < 0", cn, adOpenKeyset, adLockOptimistic, adCmdText
    'Obtengo la clave del cliente a dar de alta
    ccVeCli = IIf(rsttemp.BOF And rsttemp.EOF, -1, (rsttemp.RecordCount + 1) * -1)
  End If
  datPrimaryRS.Recordset.AddNew
  txtfields(0).Text = ccVeCli
  Me.picButtons.Visible = False
  Me.PicMod.Visible = True
  For N = 2 To 17
     txtfields(N).Locked = False
  Next
  txtfields(13).SetFocus
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
  On Error GoTo DeleteErr
  With datPrimaryRS.Recordset
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
  End With
  Exit Sub
DeleteErr:
  MsgBox Err.Description
End Sub

Private Sub cmdUpdate_Click()
'On Error GoTo UpdateErr
  If Trim(txtfields(1).Text) = "" Then
     MsgBox "ES NECESARIO ESPECIFICAR EL NOMBRE DEL CLIENTE", vbExclamation
     txtfields(1).SetFocus
     Exit Sub
  End If
  If Not IsNumeric(txtfields(4).Text) Then
     MsgBox "ES NECESARIO ESPECIFICAR UN NUMERO TELEFONICO", vbExclamation
     txtfields(4).SetFocus
     Exit Sub
  End If
  If Trim(RutPort) <> Mid(Trim(txtfields(13).Text), 1, 3) And Trim(RutPort) <> "" Then
     MsgBox "LA RUTA CAPTURADA ES DIFERENTE A " & RutPort, vbCritical, "Ruta"
     Exit Sub
  End If
  If chkCredito.Value = 1 Then
     If Trim(txtfields(5).Text) = "" Then
         MsgBox "ES NECESARIO ESPECIFICAR EL NOMBRE QUE APARECERA EN LA FACTURA", vbExclamation
         txtfields(1).SetFocus
         Exit Sub
     ElseIf Trim(txtfields(2).Text) = "" Then
         MsgBox "ES NECESARIO ESPECIFICAR UNA DIRECCION FISCAL", vbExclamation
         txtfields(2).SetFocus
         Exit Sub
     ElseIf Trim(txtfields(3).Text) = "" Then
         MsgBox "ES NECESARIO ESPECIFICAR EL R.F.C.", vbExclamation
         txtfields(3).SetFocus
         Exit Sub
     ElseIf Val(Format(txtfields(6).Text, "########0.00")) > Format(txtfields(9).Text, "########0.00") Then
         MsgBox "EL LIMITE DE CREDITO NO PUEDE EXCEDER EL MONTO DE LA FIANZA", vbExclamation
         txtfields(6).SetFocus
         Exit Sub
     End If
  End If
  If Not Sql Then
     datPrimaryRS.Recordset!cambpre = 1
     datPrimaryRS.Recordset!MODIFICADO = 1
  End If
  datPrimaryRS.Recordset.UpdateBatch adAffectAll
  Me.picButtons.Visible = True
  Me.PicMod.Visible = False
  For N = 1 To 17
     txtfields(N).Locked = True
  Next
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub txtContra_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cmdConAceptar_Click
ElseIf KeyAscii = 27 Then
   cmdConCance_Click
End If
End Sub

Private Sub txtFields_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   SendKeys vbTab
End If
End Sub

Private Sub txtfields_LostFocus(Index As Integer)
If Index = 16 Then
   If (Trim(txtfields(14)) <> "" Or Trim(txtfields(15)) <> "" Or Trim(txtfields(16)) <> "") Then
   txtfields(5).Text = Trim(txtfields(14).Text) & " " & Trim(txtfields(15).Text) & " " & Trim(txtfields(16).Text)
   txtfields(1).Text = Trim(txtfields(14).Text) & " " & Trim(txtfields(15).Text) & " " & Trim(txtfields(16).Text)
   End If
End If
indice = Index
End Sub
