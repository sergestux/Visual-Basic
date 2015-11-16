VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmTrasladaEnv 
   Caption         =   "Traslados"
   ClientHeight    =   8580
   ClientLeft      =   -15
   ClientTop       =   1305
   ClientWidth     =   11880
   Icon            =   "frmTraslados.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8580
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame fraCon 
      BackColor       =   &H00C0C000&
      Caption         =   "PROPORCIONE CONTRASEÑA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   3360
      TabIndex        =   40
      Top             =   4440
      Visible         =   0   'False
      Width           =   4215
      Begin VB.TextBox txtContra 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1920
         PasswordChar    =   "*"
         TabIndex        =   43
         Top             =   480
         Width           =   1815
      End
      Begin VB.CommandButton cmdConAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   840
         TabIndex        =   42
         Top             =   1200
         Width           =   1050
      End
      Begin VB.CommandButton cmdConCance 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   2520
         TabIndex        =   41
         Top             =   1200
         Width           =   1050
      End
      Begin VB.Label lblContra 
         BackColor       =   &H00C0C000&
         Caption         =   "Contraseña"
         Height          =   255
         Left            =   360
         TabIndex        =   44
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Frame fraajuste 
      BackColor       =   &H80000001&
      Caption         =   "Contraseña Para el Ajuste"
      Enabled         =   0   'False
      Height          =   1935
      Left            =   960
      TabIndex        =   74
      Top             =   3120
      Visible         =   0   'False
      Width           =   8535
      Begin VB.CommandButton btnajuste 
         Caption         =   "&Ajustar"
         Height          =   495
         Left            =   6720
         TabIndex        =   78
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txtobserva 
         Height          =   495
         Left            =   360
         TabIndex        =   76
         Top             =   360
         Width           =   8055
      End
      Begin VB.TextBox txtajuste 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   360
         PasswordChar    =   "*"
         TabIndex        =   75
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "Contraseña"
         Height          =   255
         Left            =   360
         TabIndex        =   77
         Top             =   1080
         Width           =   1455
      End
   End
   Begin VB.Frame fraprecios 
      BackColor       =   &H00808000&
      Caption         =   "Clave Para Cambio de Precios"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4200
      TabIndex        =   60
      Top             =   2880
      Visible         =   0   'False
      Width           =   4935
      Begin VB.TextBox tpre 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1080
         PasswordChar    =   "*"
         TabIndex        =   61
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackColor       =   &H00808000&
         Caption         =   "Escriba Contraseña para Modificar Precios por Escala"
         Height          =   375
         Left            =   360
         TabIndex        =   62
         Top             =   360
         Width           =   4335
      End
   End
   Begin VB.Frame frmgon 
      BackColor       =   &H00808080&
      Caption         =   "Nombre de la persona que realiza la salida"
      Enabled         =   0   'False
      Height          =   1455
      Left            =   840
      TabIndex        =   84
      Top             =   4680
      Visible         =   0   'False
      Width           =   10095
      Begin VB.CommandButton btngon 
         Caption         =   "&Aceptar"
         Height          =   495
         Left            =   7800
         Picture         =   "frmTraslados.frx":1272
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   480
         Width           =   1815
      End
      Begin VB.ComboBox cmbgon 
         Height          =   315
         Left            =   240
         TabIndex        =   85
         Top             =   600
         Width           =   6975
      End
   End
   Begin VB.Frame FraZONAS 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Seleccion de Zonas de Piso"
      Enabled         =   0   'False
      Height          =   1575
      Left            =   1200
      TabIndex        =   79
      Top             =   4200
      Visible         =   0   'False
      Width           =   5295
      Begin VB.CommandButton Command1 
         Caption         =   "&Continuar"
         Height          =   855
         Left            =   3720
         Picture         =   "frmTraslados.frx":1374
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   480
         Width           =   1455
      End
      Begin VB.ComboBox cmbzonas 
         Height          =   315
         Left            =   240
         TabIndex        =   80
         Text            =   "Combo1"
         Top             =   720
         Width           =   3015
      End
   End
   Begin VB.Frame Framermas 
      BackColor       =   &H00C0C000&
      Caption         =   "Clave para Mermas o Autoconsumo"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   1440
      TabIndex        =   70
      Top             =   3120
      Visible         =   0   'False
      Width           =   4935
      Begin VB.TextBox txtclamermas 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1080
         PasswordChar    =   "*"
         TabIndex        =   71
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         Caption         =   "Escriba Contraseña para Autorizar las mermas o autoconsumo"
         Height          =   375
         Left            =   240
         TabIndex        =   72
         Top             =   360
         Width           =   4695
      End
   End
   Begin VB.Frame fraAvance 
      Caption         =   "Cargando especificaciones de productos "
      Height          =   735
      Left            =   3840
      TabIndex        =   33
      Top             =   4680
      Visible         =   0   'False
      Width           =   4935
      Begin ComctlLib.ProgressBar pgb 
         Height          =   255
         Left            =   120
         TabIndex        =   65
         Top             =   360
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
   End
   Begin ComctlLib.StatusBar stbmensajes 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   64
      Top             =   8235
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   609
      Style           =   1
      SimpleText      =   "                                                                                              Para salir presione la tecla [Esc]"
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame FraCancelar 
      BackColor       =   &H00808000&
      Height          =   1935
      Left            =   2880
      TabIndex        =   34
      Top             =   1320
      Visible         =   0   'False
      Width           =   5535
      Begin VB.CommandButton cmdRegCan 
         Caption         =   "Regresar"
         Height          =   375
         Left            =   3000
         TabIndex        =   37
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar Trasl."
         Height          =   375
         Left            =   1080
         TabIndex        =   36
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtMotivo 
         Height          =   375
         Left            =   240
         MaxLength       =   50
         TabIndex        =   35
         Top             =   600
         Width           =   5175
      End
      Begin VB.Label lbletiquetas 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
         Caption         =   "Motivo"
         Height          =   255
         Index           =   7
         Left            =   360
         TabIndex        =   38
         Top             =   240
         Width           =   5175
      End
   End
   Begin VB.TextBox txtFoliotie 
      Alignment       =   2  'Center
      DataField       =   "t_foliotie"
      DataSource      =   "AdoTraslada"
      Enabled         =   0   'False
      Height          =   285
      Left            =   5760
      TabIndex        =   56
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox lstProd 
      Appearance      =   0  'Flat
      Height          =   1785
      Left            =   360
      Sorted          =   -1  'True
      TabIndex        =   30
      Top             =   3960
      Visible         =   0   'False
      Width           =   11115
   End
   Begin VB.PictureBox PicBot 
      Align           =   2  'Align Bottom
      BackColor       =   &H00808080&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   11820
      TabIndex        =   45
      Top             =   7500
      Visible         =   0   'False
      Width           =   11880
      Begin VB.CommandButton Command2 
         Height          =   400
         Left            =   5280
         Picture         =   "frmTraslados.frx":167E
         Style           =   1  'Graphical
         TabIndex        =   83
         ToolTipText     =   "Actualiza Precios del Envio"
         Top             =   80
         Width           =   500
      End
      Begin VB.CommandButton cmdprecios 
         Height          =   400
         Left            =   4680
         Picture         =   "frmTraslados.frx":1988
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "Actualiza Precios del Envio"
         Top             =   80
         Width           =   500
      End
      Begin VB.CommandButton cmdInven 
         Height          =   400
         Left            =   3120
         Picture         =   "frmTraslados.frx":1A8A
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   "Ver inventario"
         Top             =   80
         Width           =   500
      End
      Begin VB.CommandButton cmdGrabarTra 
         Caption         =   "&Dism.Inv."
         Height          =   400
         Left            =   3720
         Picture         =   "frmTraslados.frx":1B8C
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Grabar el envio y disminuir Inventario"
         Top             =   80
         Width           =   900
      End
      Begin VB.CommandButton cmdReporte 
         Height          =   400
         Left            =   5880
         Picture         =   "frmTraslados.frx":1C86
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Presentacion preeliminar del envio"
         Top             =   80
         Width           =   500
      End
      Begin VB.CommandButton cmdCodBarra 
         Height          =   400
         Left            =   2520
         Picture         =   "frmTraslados.frx":21B8
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Lectura de etiquetas por lector optico"
         Top             =   80
         Width           =   500
      End
      Begin VB.CommandButton cmdActualizar 
         Height          =   400
         Left            =   1920
         Picture         =   "frmTraslados.frx":22EE
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Actualizar envio para reflejar cambio srealizado por otros usuario"
         Top             =   80
         Width           =   500
      End
      Begin VB.CommandButton cmdTicket 
         Height          =   400
         Left            =   1320
         Picture         =   "frmTraslados.frx":23F0
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Imprimir ticket del envio"
         Top             =   80
         Width           =   500
      End
      Begin VB.CommandButton CMDCORTAR 
         Height          =   400
         Left            =   720
         Picture         =   "frmTraslados.frx":2562
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Cortar papel del ticket"
         Top             =   80
         Width           =   500
      End
      Begin VB.Label lblCajas 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PRODUCTOS:  XX   CAJAS: XX   PIEZAS:  XX"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7320
         TabIndex        =   52
         Top             =   195
         Visible         =   0   'False
         Width           =   4455
      End
   End
   Begin VB.PictureBox Rpt 
      Height          =   480
      Left            =   11040
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   87
      Top             =   7560
      Width           =   1200
   End
   Begin MSAdodcLib.Adodc AdoTraslada 
      Height          =   330
      Left            =   9600
      Top             =   480
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "AdoTraslada"
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
   Begin MSAdodcLib.Adodc AdoPedTra 
      Height          =   330
      Left            =   8400
      Top             =   0
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      ConnectStringType=   3
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
      Caption         =   "AdoPedEnTra"
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
   Begin MSAdodcLib.Adodc AdoCatUsu 
      Height          =   330
      Left            =   7320
      Top             =   0
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "AdoUsuarios"
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
   Begin MSAdodcLib.Adodc AdoDetPed 
      Height          =   330
      Left            =   7200
      Top             =   480
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
      LockType        =   3
      CommandType     =   1
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
      Caption         =   "AdoDetPed"
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
   Begin VB.TextBox txtCampos 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   0
      Left            =   2040
      MaxLength       =   10
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc AdoCatSuc 
      Height          =   330
      Left            =   4920
      Top             =   -120
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
      Caption         =   "AdoCatSuc"
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
   Begin VB.PictureBox PicBotones 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   8880
      ScaleHeight     =   495
      ScaleWidth      =   2655
      TabIndex        =   22
      Top             =   2640
      Width           =   2655
      Begin VB.CommandButton cmdRegresar 
         Caption         =   "&Regresar"
         Height          =   480
         Left            =   1440
         Picture         =   "frmTraslados.frx":2A94
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Regresar a la pantalla principal"
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         Height          =   480
         Left            =   120
         Picture         =   "frmTraslados.frx":2C06
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Grabar traslado y mostrar detalle"
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin MSAdodcLib.Adodc AdoProduct 
      Height          =   330
      Left            =   2520
      Top             =   0
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   1
      LockType        =   3
      CommandType     =   1
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
      Caption         =   "AdoProduct"
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
   Begin VB.Frame FraPedEnv 
      Caption         =   " Pedidos enviados en el traslado"
      ForeColor       =   &H8000000D&
      Height          =   1815
      Left            =   8880
      TabIndex        =   26
      Top             =   600
      Visible         =   0   'False
      Width           =   2655
      Begin VB.ComboBox cmbPedido 
         Height          =   315
         Left            =   1200
         TabIndex        =   32
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton cmdAgregar 
         Height          =   375
         Left            =   120
         Picture         =   "frmTraslados.frx":2D78
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Agregar pedido al traslado"
         Top             =   1200
         Width           =   975
      End
      Begin MSDataGridLib.DataGrid dbgrdPedEnTra 
         Bindings        =   "frmTraslados.frx":2EEA
         Height          =   735
         Left            =   600
         TabIndex        =   27
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1296
         _Version        =   393216
         AllowUpdate     =   0   'False
         BorderStyle     =   0
         HeadLines       =   1
         RowHeight       =   19
         RowDividerStyle =   4
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "Pedidos"
            Caption         =   "   FOLIO PED."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            RecordSelectors =   0   'False
            BeginProperty Column00 
               Alignment       =   2
            EndProperty
            BeginProperty Column01 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
   End
   Begin MSDataGridLib.DataGrid DbgrdDetTraAbi 
      Bindings        =   "frmTraslados.frx":2F02
      Height          =   2970
      Left            =   240
      TabIndex        =   16
      Top             =   3240
      Visible         =   0   'False
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   5239
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   17
      TabAction       =   2
      WrapCellPointer =   -1  'True
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "DESGLOSE DEL TRASLADO ABIERTO"
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "dt_producto"
         Caption         =   "      CLAVE"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "exicaj"
         Caption         =   "  EXI.CAJ."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "InCantPza"
         Caption         =   "  EXI.PZA."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "descripc"
         Caption         =   "                            DESCRIPCION DEL PRODUCTO"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "PRESENT"
         Caption         =   "PRESENTACION"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "dt_cantidad"
         Caption         =   "CAJ.ENV."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "dt_cantidadp"
         Caption         =   "PZA.ENV."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "dt_costo"
         Caption         =   "$COST.CAJ."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "dt_venta"
         Caption         =   "$VTA.CAJA"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Alignment       =   2
            Button          =   -1  'True
            ColumnWidth     =   1124.787
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
            ColumnWidth     =   4305.26
         EndProperty
         BeginProperty Column04 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   870.236
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            ColumnWidth     =   824.882
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
            Locked          =   -1  'True
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid dbgrdDetPed 
      Bindings        =   "frmTraslados.frx":2F1A
      Height          =   3945
      Left            =   240
      TabIndex        =   13
      Top             =   3240
      Visible         =   0   'False
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   6959
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      BackColor       =   16777215
      BorderStyle     =   0
      ColumnHeaders   =   -1  'True
      ForeColor       =   -2147483642
      HeadLines       =   1
      RowHeight       =   17
      TabAction       =   1
      WrapCellPointer =   -1  'True
      RowDividerStyle =   6
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "DESGLOSE DEL TRASLADO POR PEDIDO"
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "dt_cantidad"
         Caption         =   "CAJ.ENV."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "dt_cantidadP"
         Caption         =   "PZA.ENV."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "dt_producto"
         Caption         =   "    CLAVE"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "descripc"
         Caption         =   "                  DESCRIPCION DEL PRODUCTO"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Medida"
         Caption         =   "  PRESENTACION"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "TotCaj"
         Caption         =   "EXI. CAJAS"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "InCantPza"
         Caption         =   " EXI.PZA."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "dt_costo"
         Caption         =   "$COST.CAJ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "DT_VENTA"
         Caption         =   "$VTA.CAJA"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Alignment       =   1
            ColumnWidth     =   1349.858
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
            ColumnWidth     =   989.858
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
            ColumnWidth     =   3630.047
         EndProperty
         BeginProperty Column04 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   840.189
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
            Locked          =   -1  'True
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraGenerales 
      Caption         =   "Datos generales"
      ForeColor       =   &H8000000D&
      Height          =   2535
      Left            =   240
      TabIndex        =   14
      Top             =   600
      Visible         =   0   'False
      Width           =   8415
      Begin VB.CheckBox Chkvolumen 
         Caption         =   "Volumen"
         DataSource      =   "AdoTraslada"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3720
         TabIndex        =   82
         Top             =   1680
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Ajuste"
         DataField       =   "t_ajuste"
         DataSource      =   "AdoTraslada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3720
         TabIndex        =   73
         Top             =   2160
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CheckBox chkauto 
         Caption         =   "Autoconsumo"
         DataField       =   "t_auto"
         DataSource      =   "AdoTraslada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   69
         Top             =   2160
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CheckBox chkmerma 
         Caption         =   "Merma"
         DataField       =   "t_merma"
         DataSource      =   "AdoTraslada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   68
         Top             =   2260
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CheckBox chkpan 
         Caption         =   "Panaderia"
         DataField       =   "t_pan"
         DataSource      =   "AdoTraslada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3720
         TabIndex        =   67
         Top             =   1920
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CheckBox chkfrutas 
         Caption         =   "Frutas"
         DataField       =   "t_frutas"
         DataSource      =   "AdoTraslada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   66
         Top             =   1920
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CheckBox chkPapeleria 
         Caption         =   "Papeleria"
         DataField       =   "t_papeleria"
         DataSource      =   "AdoTraslada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   57
         Top             =   1950
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CheckBox chkRecibido 
         Caption         =   "&Recibido"
         DataField       =   "t_recibido"
         DataSource      =   "AdoTraslada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   54
         Top             =   1680
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CheckBox chkEnviado 
         Caption         =   "&Enviado"
         DataField       =   "t_enviado"
         DataSource      =   "AdoTraslada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   29
         Top             =   1680
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CheckBox chkTipoTrasl 
         Caption         =   "&Abierto"
         DataField       =   "t_tipo"
         DataSource      =   "AdoTraslada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1680
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox txtCampos 
         Alignment       =   2  'Center
         DataField       =   "T_FECHA"
         DataSource      =   "AdoTraslada"
         Height          =   285
         Index           =   6
         Left            =   6480
         TabIndex        =   12
         Top             =   1800
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.ComboBox cmbFlete 
         Height          =   315
         Left            =   2880
         TabIndex        =   8
         Top             =   1320
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtCampos 
         DataField       =   "t_PerFle"
         DataSource      =   "AdoTraslada"
         Height          =   285
         Index           =   5
         Left            =   2040
         TabIndex        =   7
         Top             =   1320
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtCampos 
         Alignment       =   2  'Center
         DataField       =   "t_costo"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   2
         EndProperty
         DataSource      =   "AdoTraslada"
         Height          =   285
         Index           =   4
         Left            =   6840
         TabIndex        =   11
         Top             =   1440
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.ComboBox cmbUsuEmi 
         Height          =   315
         Left            =   2880
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   600
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.ComboBox cmbSucRec 
         Height          =   315
         Left            =   2880
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   960
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.ComboBox cmbSucEmi 
         Height          =   315
         Left            =   2880
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   240
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.TextBox txtCampos 
         DataField       =   "t_sucursalreceptor"
         DataSource      =   "AdoTraslada"
         Height          =   285
         Index           =   3
         Left            =   2040
         MaxLength       =   3
         TabIndex        =   5
         Top             =   960
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtCampos 
         DataField       =   "t_perenv"
         DataSource      =   "AdoTraslada"
         Height          =   285
         Index           =   2
         Left            =   2040
         MaxLength       =   3
         TabIndex        =   3
         Top             =   600
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtCampos 
         DataField       =   "t_sucursalemisor"
         DataSource      =   "AdoTraslada"
         Height          =   285
         Index           =   1
         Left            =   2040
         MaxLength       =   5
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lbletiquetas 
         Caption         =   "Importe del traslado"
         Height          =   255
         Index           =   4
         Left            =   5280
         TabIndex        =   24
         Top             =   1440
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label lbletiquetas 
         Caption         =   "Clave del flete"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   23
         Top             =   1320
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lblDesprov 
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   2640
         TabIndex        =   21
         Top             =   360
         Width           =   3975
      End
      Begin VB.Label lbletiquetas 
         Caption         =   "Persona emisora"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   20
         Top             =   600
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lbletiquetas 
         Caption         =   "Fec.elab. de tras"
         Height          =   255
         Index           =   6
         Left            =   5280
         TabIndex        =   19
         Top             =   1800
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lbletiquetas 
         Caption         =   "Sucursal receptora"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   18
         Top             =   960
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lbletiquetas 
         Caption         =   "Sucursal emisora"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Visible         =   0   'False
         Width           =   1575
      End
   End
   Begin VB.Label lblinfo 
      Caption         =   "."
      Height          =   255
      Left            =   360
      TabIndex        =   63
      Top             =   7320
      Width           =   8775
   End
   Begin VB.Label LblProdAgr 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   360
      TabIndex        =   39
      Top             =   6960
      Width           =   11295
   End
   Begin VB.Label lbletiquetas 
      Caption         =   "Folio del traslado por tienda"
      Height          =   255
      Index           =   8
      Left            =   3720
      TabIndex        =   55
      Top             =   240
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label lblEntrada 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ENTRADA DE PRODUCTOS"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   315
      Left            =   4320
      TabIndex        =   53
      Top             =   120
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Label lblDescrip 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   6840
      TabIndex        =   31
      Top             =   5040
      Width           =   4935
   End
   Begin VB.Label lbletiquetas 
      Caption         =   "Folio unico del traslado"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   15
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "frmTrasladaEnv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Pedido As String
Private lCon As Boolean
Private clave
Private lModTra  As Integer
Private nFolAju As Integer
Private nAncho As Integer
Private lfranquicia As Boolean

Private Sub cmbJefeRec_Validate(Cancel As Boolean)
AdoCatUsu.Recordset.MoveFirst
AdoCatUsu.Recordset.Find "Name = '" & cmbJefeRec.Text & "'"
If AdoCatUsu.Recordset.EOF = True Then
   MsgBox "Debe seleccionar un usuario de la lista desplegable", vbExclamation
   cmbJefeRec.SetFocus
   Cancel = True
Else
   txtcampos(9).Text = AdoCatUsu.Recordset!clave
   txtcampos(9).SetFocus
End If
End Sub

Private Sub btnajuste_Click()
If txtajuste.Text = "P4567" Then
    cn.Execute "UPDATE TRASLADOS SET T_AJUSTE = 1 , T_OBSERVA = '" & Trim(txtobserva.Text) & "'  WHERE T_CLAVE = '" & Trim(txtcampos(0).Text) & "'"
    MsgBox "AHORA PUEDE REALIZAR EL AJUSTE A LOS PRODUCTOS ", vbInformation, "AJUSTES"
Else
    MsgBox "LA CONTRASEÑA NO COINCIDE PARA REALIZAR UN AJUSTE ", vbInformation, "AJUSTES"
    cn.Execute "UPDATE TRASLADOS SET T_AJUSTE = 0 T_OBSERVA = ''   WHERE T_CLAVE = '" & Trim(txtcampos(0).Text) & "'"
    Check1.Value = 0
End If
fraajuste.Enabled = False
fraajuste.Visible = False
End Sub

Private Sub btngon_Click()
On Error GoTo Error:
If cmbgon.Text <> "" Then
N = InStr(1, cmbgon, "[")
ty = Mid(cmbgon, N + 1, Len(cmbgon) - N - 1)
CAD = "UPDATE traslados SET T_GON = " & ty & "  WHERE t_clave = '" & Trim(txtcampos(0).Text) & "'"
cn.Execute CAD
frmgon.Enabled = False
frmgon.Visible = False
End If
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub chkauto_Click()
On Error GoTo Error:
If chkauto.Value = 1 Then
   If Not AdoTraslada.Recordset!t_enviado Then
        Framermas.Enabled = True
        Framermas.Visible = True
        txtclamermas.Text = ""
        Me.txtclamermas.SetFocus
        PORAUTO = True
        PORMERMAS = False
   End If
End If
Exit Sub
Error:
  Framermas.Enabled = False
  Framermas.Visible = False
  txtclamermas.Text = ""
End Sub

Private Sub chkmerma_Click()
If chkMerma.Value = 1 Then
   If Not AdoTraslada.Recordset!t_enviado Then
        Framermas.Enabled = True
        Framermas.Visible = True
        txtclamermas.Text = ""
        Me.txtclamermas.SetFocus
        PORMERMAS = True
        PORAUTO = False
   End If
End If
End Sub

Private Sub cmbFlete_GotFocus()
RESP = SendMessageLong(cmbFlete.hwnd, &H14F, True, 1)
End Sub

Private Sub cmbFlete_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   'SendKeys "{Tab}"
   keybd_event &H9, 0, 0, 0
End If
End Sub

Private Sub cmbFlete_LostFocus()
RESP = SendMessageLong(cmbFlete.hwnd, &H14F, False, 1)
End Sub

Private Sub cmbFlete_Validate(Cancel As Boolean)
 If InStr("12345", Mid(cmbFlete.Text, 1, 1)) = 0 Then
   MsgBox "Debe seleccionar un tipo de transporte la lista desplegable", vbExclamation
   cmbFlete.SetFocus
   Cancel = True
   Exit Sub
 End If
 txtcampos(5).Text = Mid(cmbFlete.Text, 1, 1)
 txtcampos(5).SetFocus
End Sub

Private Sub cmbPedido_GotFocus()
RESP = SendMessageLong(cmbPedido.hwnd, &H14F, True, 1)
End Sub

Private Sub cmbPedido_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Me.CmdAgregar.SetFocus
End If
End Sub

Private Sub cmbSucEmi_GotFocus()
RESP = SendMessageLong(cmbSucEmi.hwnd, &H14F, True, 1)
End Sub

Private Sub cmbSucEmi_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   'SendKeys "{Tab}"
   keybd_event &H9, 0, 0, 0
End If
End Sub

Private Sub cmbSucEmi_LostFocus()
RESP = SendMessageLong(cmbSucEmi.hwnd, &H14F, False, 1)
End Sub

Private Sub cmbSucEmi_Validate(Cancel As Boolean)
If cmbSucEmi.Text = "" Or IsNull(cmbSucEmi.Text) Then
   MsgBox "Debe seleccionar una sucursal de la lista desplegable", vbExclamation
   cmbSucEmi.SetFocus
   Cancel = True
Else
   AdoCatSuc.Recordset.MoveFirst
   AdoCatSuc.Recordset.Find "Tidescrip = '" & cmbSucEmi.Text & "'"
   If AdoCatSuc.Recordset.EOF = True Then
        MsgBox "Debe seleccionar una sucursal de la lista desplegable", vbExclamation
        cmbSucEmi.SetFocus
        Cancel = True
    Else
        txtcampos(1).Text = AdoCatSuc.Recordset!ticlave
        txtcampos(1).SetFocus
    End If
End If
End Sub

Private Sub cmbSucRec_GotFocus()
RESP = SendMessageLong(cmbSucRec.hwnd, &H14F, True, 1)
End Sub

Private Sub cmbSucRec_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   'SendKeys "{Tab}"
   keybd_event &H9, 0, 0, 0
End If
End Sub

Private Sub cmbSucRec_LostFocus()
RESP = SendMessageLong(cmbSucRec.hwnd, &H14F, False, 1)
End Sub

Private Sub cmbSucRec_Validate(Cancel As Boolean)
If cmbSucRec.Text = "" Or IsNull(cmbSucRec.Text) Then
   MsgBox "Debe seleccionar una sucursal de la lista desplegable", vbExclamation
   cmbSucRec.SetFocus
   Cancel = True
Else
    AdoCatSuc.Recordset.MoveFirst
    AdoCatSuc.Recordset.Find IIf(cModo = "DEVO", "NOMPROVE = '" & cmbSucRec.Text & "'", "Tidescrip = '" & cmbSucRec.Text & "'")
    If AdoCatSuc.Recordset.EOF = True Then
        MsgBox "Debe seleccionar una sucursal de la lista desplegable", vbExclamation
        cmbSucRec.SetFocus
        Cancel = True
    Else
        If cModo = "DEVO" Then
           txtcampos(3).Text = AdoCatSuc.Recordset!prove
        Else
           txtcampos(3).Text = AdoCatSuc.Recordset!ticlave
        End If
        txtcampos(3).SetFocus
    End If
End If

End Sub

Private Sub cmbUsuEmi_GotFocus()
'Resp = SendMessageLong(cmbUsuEmi.hwnd, &H14F, True, 1)
End Sub

Private Sub cmbUsuEmi_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   'SendKeys "{Tab}"
   keybd_event &H9, 0, 0, 0
End If
End Sub

Private Sub cmbUsuRec_Validate(Cancel As Boolean)
AdoCatUsu.Recordset.MoveFirst
AdoCatUsu.Recordset.Find "Name = '" & cmbUsuRec.Text & "'"
If AdoCatUsu.Recordset.EOF = True Then
   MsgBox "Debe seleccionar un usuario de la lista desplegable", vbExclamation
   cmbUsuRec.SetFocus
   Cancel = True
Else
   txtcampos(4).Text = AdoCatUsu.Recordset!clave
   txtcampos(4).SetFocus
End If
End Sub

Private Sub cmdCancelar_Click()
On Error GoTo Error:
  If Trim(txtMotivo.Text) = "" Or IsNull(txtMotivo) Then
     MsgBox "ES NECESARIO ESPECIFICAR UN MOTIVO DE LA CANCELACION", vbExclamation
     txtMotivo.SetFocus
     Exit Sub
  End If
  If AdoTraslada.Recordset!t_entrada Then
     cMensaje = "REALMENTE DESEAS CANCELAR LA ENTRADA" & Chr(13) & "Y DISMINUIR EL INVENTARIO?"
     COPERA = "-"
  Else
     cMensaje = "REALMENTE DESEAS CANCELAR LA SALIDA" & Chr(13) & "Y AUMENTAR EL INVENTARIO?"
     COPERA = "+"
  End If
  If MsgBox(cMensaje, vbInformation + vbYesNo) = vbYes Then
     AdoTraslada.Recordset!t_motivoCancela = txtMotivo.Text
     AdoTraslada.Recordset.Update
     AdoDetPed.Recordset.MoveFirst
     While Not AdoDetPed.Recordset.EOF
        cn.Execute "UPDATE INVENTARIO SET InCant = Incant " & COPERA & CStr(AdoDetPed.Recordset!dt_cantidad) & ", InCantPza = InCantPza " & COPERA & CStr(AdoDetPed.Recordset!dt_cantidadp) & " WHERE Inprod = '" & AdoDetPed.Recordset!Dt_producto & "'"
        'cn.Execute "UPDATE CODIGOS SET SalidaInv = '0', Fechasalida = NULL, Traslado = Null WHERE substring(CODIGO,1,10) = '" & AdoDetPed.Recordset!Dt_producto & "'"
        AdoDetPed.Recordset.MoveNext
     Wend
     cn.Execute " update traslados set t_costo = 0, t_monto = 0 where t_clave =  '" & Trim(txtcampos(0).Text) & "'"
     Unload Me
  End If
Exit Sub
Error:
    MsgBox Err.Description
End Sub

Private Sub CmdActualizar_Click()
On Error Resume Next
  AdoDetPed.Refresh
  LblProdAgr.Caption = ""
End Sub

Private Sub cmdAgregar_Click()
Dim RESP
'Dim rstped As ADODB.Recordset
Dim rstPedTra As ADODB.Recordset
'Dim rsttemp As ADODB.Recordset
'On Error GoTo ERROR:
  If Trim(cmbPedido.Text) = "" Or IsNull(cmbPedido.Text) Then
     MsgBox "No puedes dejar en blanco el folio del pedido a enviar", vbExclamation
     Exit Sub
  End If
  Set rstPedTra = New ADODB.Recordset
  rstPedTra.CursorType = adOpenKeyset
  rstPedTra.LockType = adLockReadOnly
'  cad = "SELECT * FROM DetalleTraslado d, Traslados t Where d.dt_clave = t.t_clave AND  d.dt_pedido = '" & Trim(cmbPedido.Text) & "' "
  CAD = "SELECT DT_CLAVE,DT_PEDIDO FROM DetalleTraslado d  Where d.dt_pedido = '" & Trim(cmbPedido.Text) & "' "
  rstPedTra.Source = CAD
  rstPedTra.ActiveConnection = cn
  rstPedTra.Open
  If Not (rstPedTra.BOF And rstPedTra.EOF) Then
     rstPedTra.MoveLast
     MsgBox "EL FOLIO DEL PEDIDO YA FUE ENVIADO EN EL TRASLADO " + rstPedTra!DT_CLAVE, vbQuestion
     Exit Sub
  End If
  Set rstPedTra = Nothing
  'rstPedTra.Close
  'Set rstped = New ADODB.Recordset
  'rstped.LockType = adLockOptimistic
  'rstped.CursorType = adOpenKeyset
  'rstped.ActiveConnection = cn
  'rstped.Source = "SELECT * FROM Pedidos,CatTienda Where P_pedido = '" & cmbPedido.Text & "' AND Pedidos.P_sucursal = CatTienda.TiClave"
  'rstped.Open
  'If rstped.BOF And rstped.EOF Then
  '   MsgBox "El numero de pedido no existe"
  '   txtPedido.SetFocus
  '   Exit Sub
  'Else
     'resp = MsgBox("EL PEDIDO CON FOLIO    " & cmbPedido.Text & " ES PARA " & rstped!tidescrip & " ELABORADO CON FECHA (M/D/A) " & rstped!p_fecped & Chr(13) & Space(30) & "DESEAS ENVIAR EL PEDIDO?", vbQuestion + vbYesNo)
     RESP = MsgBox("Esta seguro de realizar la integracion del pedido Abierto ", vbQuestion + vbYesNo)
     If RESP = vbYes Then
        'rstped.Close
        'rstped.Open "SELECT SUM(df_cantsolp) AS SOLPZA FROM Detallefactura WHERE df_pedido = '" & cmbPedido.Text & "'" & ""
        'If rstped!solpza > 0 Then
        '   cresp = MsgBox("EN EL PEDIDO SELECCIONADO EXISTEN CANTIDADES SOLICITADAS EN PIEZAS" & _
        '           " QUE ES LO QUE DESEA HACER ?" & Chr(13) & Chr(13) & _
        '           "[SI]   Pasar piezas a cajas en los que la presentacion es 1x1 y los demas se borran " & Chr(13) & _
        '           "[NO] Dejar las cantidades solicitadas en piezas", vbYesNoCancel + vbQuestion)
        '   cmensa = stbmensajes.SimpleText
        '   stbmensajes.SimpleText = Space(50) & "Espere un momento cargando el detalle del traslado......"
        '   stbmensajes.Refresh
        '   If cresp = vbCancel Then
        '      Exit Sub
        '   Else  'Si desean continuar con el proceso
        '      cn.Execute "INSERT INTO " & _
        '                 "[DetalleTraslado] (dt_clave,dt_producto,dt_cantidad,dt_cantidadp,dt_pedido) SELECT  dt_clave = '" & _
        '                 txtCampos(0).Text & "',df_prod,df_cantsol,df_cantsolp, pedido = '" & cmbPedido.Text & "' FROM detallefactura where df_pedido = '" & cmbPedido.Text & "' AND df_sugerido = 0"
              'If cresp = vbYes Then  'Paso piezas a cajas en prod con present. 1x1
              '   cn.Execute "UPDATE detalletraslado SET dt_cantidad = dt_cantidad + dt_cantidadp FROM tfproduc WHERE  Consec = dt_producto AND dt_clave = '" & txtCampos(0).Text & "' AND Paquetes = 1"
              '   cn.Execute "UPDATE detalletraslado SET dt_cantidadp = 0 WHERE dt_clave = '" & txtCampos(0).Text & "'"
              'End If
           'End If
       ' Else  'No solicitaron productos en piezas
        cn.Close
        cn.CommandTimeout = 0
        cn.ConnectionTimeout = 0
        cn.Open
        cn.Execute "INSERT INTO " & _
            "[DetalleTraslado] (dt_clave,dt_producto,dt_cantidad,dt_cantidadp,dt_pedido) SELECT  dt_clave = '" & _
            txtcampos(0).Text & "',df_prod,df_cantsol,df_cantsolp, pedido = '" & cmbPedido.Text & "' FROM detallefactura where df_pedido = '" & cmbPedido.Text & "' AND df_sugerido = 0"
        cn.Execute "UPDATE PEDIDOS SET p_traslado = '" & txtcampos(0).Text & "', p_fecent = '" & Trim(txtcampos(6).Text) & "' WHERE p_pedido = '" & Trim(cmbPedido.Text) & "' AND P_PROVEEDOR = 'ABA'"
        If lfranquicia Then
            cn.Execute "UPDATE DETALLETRASLADO SET DETALLETRASLADO.dt_costo = TFPRODUC.precosto, DETALLETRASLADO.dt_costoP = TFPRODUC.precosto / TFPRODUC.paquetes, DETALLETRASLADO.dt_venta = PREPROD.precio4, DETALLETRASLADO.dt_ventaP = PREPROD.precio4 / TFPRODUC.paquetes FROM PREPROD,TFPRODUC WHERE PREPROD.preclave = DETALLETRASLADO.dt_producto AND PREPROD.preclave = TFPRODUC.Consec AND TFPRODUC.Consec = DETALLETRASLADO.dt_producto AND DetalleTraslado.dt_clave = '" & txtcampos(0).Text & "' AND tfproduc.ACTIVO = 1"
        Else 'Precio a tiendas (PRECOSTO de tfproduc)
            cn.Execute "UPDATE DETALLETRASLADO SET dt_costo = TFPRODUC.Precosto, dt_costop = TFPRODUC.Precosto / TFPRODUC.Paquetes, dt_venta = PREPROD.Precio2, dt_ventap = PREPROD.Precio1 FROM TFPRODUC,PREPROD WHERE PREPROD.preclave = DETALLETRASLADO.dt_producto AND PREPROD.preclave = TFPRODUC.Consec AND TFPRODUC.Consec = DETALLETRASLADO.dt_producto AND DetalleTraslado.dt_clave = '" & txtcampos(0).Text & "' AND tfproduc.ACTIVO = 1"
        End If
        cn.Execute "update traslados set t_fecha = '" & date + Time & "' where t_clave = '" & Trim(txtcampos(0).Text) & "'"
        AdoPedTra.RecordSource = "SELECT MIN(dt_pedido) AS PEDIDOS FROM [DetalleTraslado]WHERE dt_clave = '" & txtcampos(0).Text & "' GROUP BY DT_PEDIDO"
        AdoPedTra.Refresh
        CmdAgregar.Enabled = False
        DesgTraPed  'Muestra el desglose del traslado cerrado o por pedido
        'Muestro etiqueta con el numero de productos, cajas y piezas
        PonPie
        Me.dbgrdDetPed.SetFocus
        StbMensajes.SimpleText = cmensa
        StbMensajes.Refresh
        MsgBox "Para mejor Rendimiento es necesario salirse del programa debido a que este proceso ha ocupado memoria virtual", vbExclamation, "SQL-SERVER"
    Else
       Exit Sub
    End If
    'Muestro los pedidos enviados en el traslado
    
  'End If
  'Set rstped = Nothing
Exit Sub
Error:
   MsgBox Err.Description
End Sub

Private Sub cmdCodBarra_Click()
If chkTipoTrasl.Value = 0 Then
  nOp = 2  'Para que la forma de lectura de codigos de barra sepa de donde se esta llamando
  frmCodBarrCap.Show 1
Else
  nOp = 3  'Para que la forma de lectura de codigos de barra sepa de donde se esta llamando
  frmCodBarrCap.Show 1
End If
End Sub

Private Sub cmdConAceptar_Click()
If txtContra.Text <> "P4567" Then
   MsgBox "LA CONTRASEÑA DE ACCESO ES INCORRECTA", vbExclamation
   txtContra.SetFocus
   SendKeys "+{HOME}"
   Exit Sub
Else
   lCon = True
   fraCon.Visible = False
   MsgBox "Despues de Hacer las Correcciones debe Actualizar Precios [ACT PRECIOS]", vbInformation
   If lModTra = 0 Then   'Modificar traslado por pedido cuando ya fue enviado
      dbgrdDetPed.AllowUpdate = True
      dbgrdDetPed.Columns(7).Locked = False 'Para tambien puedan modificar precio COSTO
      dbgrdDetPed.Columns(8).Locked = False 'Para tambien puedan modificar precio VENTA
      dbgrdDetPed.SetFocus
   Else
      DbgrdDetTraAbi.AllowUpdate = True
      DbgrdDetTraAbi.Columns(7).Locked = False  'Para tambien puedan modificar precio COSTO
      DbgrdDetTraAbi.Columns(8).Locked = False  'Para tambien puedan modificar precio VENTA
      DbgrdDetTraAbi.SetFocus
   End If
End If
End Sub

Private Sub cmdConCance_Click()
  fraCon.Visible = False
End Sub

Private Sub CmdCortar_Click()
On Error Resume Next
  'Handle = Shell("C:\WINDOWS\CORTA.EXE", 1)
  'Handle = Shell(App.Path & "\CORTA.EXE", 1)
End Sub

Private Sub cmdGrabar_Click()
Dim rsttemp As ADODB.Recordset
Dim rs As ADODB.Recordset
On Error GoTo Error:
  'Actualizo tabla de traslados
  cMenAnt = StbMensajes.SimpleText
  StbMensajes.SimpleText = Space(70) + "Grabando datos generales del traslado.."
  StbMensajes.Refresh
    
  If nOp = 1 Then 'Cuando es un nuevo traslado
     Set rsttemp = New ADODB.Recordset
     rsttemp.ActiveConnection = cCadConex
     rsttemp.CursorType = adOpenKeyset
                  
     rsttemp.Source = "SELECT MAX (CAST(SUBSTRING(t_clave,4,10) AS INT)) As FolTra FROM [Traslados] WHERE SUBSTRING(t_clave,1,3) = 'T" & Trim(Mid(cSucursal, 3, 5)) & "'"
     rsttemp.Open
     If IsNull(rsttemp!FolTra) Then
        Folio = "T" + Mid(Trim(Mid(cSucursal, 3)), 1, 2) + "1"
     Else
        Folio = "T" + Mid(Trim(Mid(cSucursal, 3)), 1, 2) + Trim(Str(rsttemp!FolTra + 1))
     End If
     txtcampos(0).Text = Folio
     AdoTraslada.Recordset!t_clave = txtcampos(0).Text
     AdoTraslada.Recordset!T_RECIBIDO = 0
     AdoTraslada.Recordset!t_enviado = 0
     AdoTraslada.Recordset!t_tipo = chkTipoTrasl.Value
     If Forma = 0 And cModo <> "DEVO" Then  'En salida de producto y que no sea devolucion
        Set rs = New ADODB.Recordset
        lbletiquetas(8).Visible = True  'Se muestra en salidas y en modificaciones
        txtFoliotie.Visible = True
        txtFoliotie.Enabled = False
     End If
  End If
  If AdoTraslada.Recordset!t_entrada Then
     chkEnviado.Visible = False
     chkEnviado.Enabled = False
     chkRecibido.Visible = True
     chkRecibido.Value = False
     cmdGrabarTra.Caption = "&Aum. Inv."
     cmdGrabarTra.ToolTipText = "Grabar recepcion de traslado y aumentar inventario"
  End If
  AdoTraslada.Recordset!T_devolucion = IIf(cModo = "DEVO", 1, 0)
  AdoTraslada.Recordset!t_entrada = Forma    'Si es entrada o salida de mercancia  1 = Entrada, 0 = Salida
  AdoTraslada.Recordset.Update
  chkEnviado.Visible = True
  txtcampos(3).Enabled = False:  txtcampos(5).Enabled = False
  Me.cmbSucRec.Enabled = False: Me.cmbFlete.Enabled = False
  'Aquellos que tienen como antecedente un pedido (Cerrado)
  If chkTipoTrasl.Value = 0 Then
        Set rsttemp = New ADODB.Recordset
        'rsttemp.Source = "SELECT * FRGOM PEDIDOS WHERE p_pedproveedor <> '' AND p_recibido = 0 AND p_sucursal = '" & txtCampos(3).Text & "'"
        rsttemp.Source = "SELECT * FROM PEDIDOS WHERE p_recibido = 0 AND p_sucursal = '" & txtcampos(3).Text & "' AND P_proveedor = 'ABA' AND p_traslado IS  NULL"
        rsttemp.ActiveConnection = cCadConex
        rsttemp.Open
        While Not rsttemp.EOF
            cmbPedido.AddItem rsttemp!p_Pedido
            rsttemp.MoveNext
        Wend
        
        'Cargo los pedidos que se incluyen en el traslado, se utiliza la misma tabla
        AdoPedTra.ConnectionString = cn 'cCadConex
        AdoPedTra.CommandType = adCmdText
        AdoPedTra.RecordSource = "SELECT MIN(dt_pedido) AS PEDIDOS FROM [DetalleTraslado]WHERE dt_clave = '" & txtcampos(0).Text & "' GROUP BY DT_PEDIDO"
        AdoPedTra.Refresh
        dbgrdPedEnTra.Visible = True
        FraPedEnv.Visible = True
        dbgrdDetPed.Visible = True
        dbgrdDetPed.Refresh
        dbgrdPedEnTra.SetFocus
  Else 'Traslados abiertos
     AdoDetPed.ConnectionString = cCadConex
     AdoDetPed.CommandType = adCmdText
     AdoDetPed.RecordSource = "SELECT consec, descripc,LTrim(str(Paquetes)) + ' X ' + LTRIM( str(contenid,10,3)) + ' ' + MEDIDA as PRESENT, dt_clave, dt_producto, dt_cantidad, dt_cantidadp, dt_clave,Inprod, CAST(InCant AS SMALLMONEY) as EXICAJ, Detalletraslado.dt_costo,Detalletraslado.dt_venta,Ubicacion, InCantPza, Paquetes FROM DetalleTraslado,TFPRODUC, Inventario WHERE DetalleTraslado.dt_clave = '" & txtcampos(0).Text & "' AND dt_producto = consec AND CONSEC *= INPROD AND INPROD =* DT_PRODUCTO ORDER BY descripc,contenid"
     AdoDetPed.Refresh
     AdoProduct.ConnectionString = cCadConex
     AdoProduct.CommandType = adCmdText
     CAD = "UPDATE traslados SET t_fecha = '" & date + Time & "' WHERE t_clave = '" & Trim(txtcampos(0).Text) & "'"
     cn.Execute CAD
     If AdoTraslada.Recordset!t_entrada Then
        StbMensajes.SimpleText = Space(50) + "Espere un momento seleccionando catalogo de productos"
        AdoProduct.RecordSource = "SELECT consec, descripc, LTrim(str(Paquetes)) + ' X ' + LTRIM( str(contenid,10,3)) + ' ' + MEDIDA as MEDIDA, Paquetes, Incant, inprod, InCantPza FROM Inventario,Tfproduc WHERE Inventario.Inprod = Tfproduc.Consec AND Paquetes > 0 AND tfproduc.activo = 1"
        chkEnviado.Visible = False
        chkRecibido.Visible = True
        chkRecibido.Value = False
        cmdGrabarTra.Caption = "&Aum. Inv."
        cmdGrabarTra.ToolTipText = "Grabar recepcion de traslado y aumentar inventario"
     Else
        StbMensajes.SimpleText = Space(50) + "Espere un momento seleccionando inventario con existencia mayor a cero"
        If cModo = "DEVO" Then
           'AdoProduct.RecordSource = "SELECT consec, descripc, LTrim(str(Paquetes)) + ' X ' + LTRIM( str(contenid,10,3)) + ' ' + MEDIDA as MEDIDA, Paquetes, Incant, InCantPza, inprod FROM Inventario,Tfproduc WHERE Inventario.Inprod = Tfproduc.Consec AND (Inventario.incant > 0 OR inCantPza > 0 ) AND Paquetes > 0 AND Claprove = '" & txtCampos(3).Text & "' AND TFPRODUC.activo = 1"
           AdoProduct.RecordSource = "SELECT consec, descripc, LTrim(str(Paquetes)) + ' X ' + LTRIM( str(contenid,10,3)) + ' ' + MEDIDA as MEDIDA, Paquetes, Incant, InCantPza, inprod FROM Inventario,Tfproduc WHERE Inventario.Inprod = Tfproduc.Consec AND (Inventario.incant > 0 OR inCantPza > 0 ) AND Paquetes > 0 AND Claprove = '" & txtcampos(3).Text & "'"
        Else
           'AdoProduct.RecordSource = "SELECT consec, descripc, LTrim(str(Paquetes)) + ' X ' + LTRIM( str(contenid,10,3)) + ' ' + MEDIDA as MEDIDA, Paquetes, Incant, InCantPza, inprod FROM Inventario,Tfproduc WHERE Inventario.Inprod = Tfproduc.Consec AND (Inventario.incant > 0 OR inCantPza > 0 ) AND Paquetes > 0 AND TFPRODUC.activo = 1"
           AdoProduct.RecordSource = "SELECT consec, descripc, LTrim(str(Paquetes)) + ' X ' + LTRIM( str(contenid,10,3)) + ' ' + MEDIDA as MEDIDA, Paquetes, Incant, InCantPza, inprod FROM Inventario,Tfproduc WHERE Inventario.Inprod = Tfproduc.Consec AND (Inventario.incant > 0 OR inCantPza > 0 ) AND Paquetes > 0"
        End If
     End If
     StbMensajes.Refresh
     AdoProduct.Refresh
     
     For N = 0 To 6
         txtcampos(N).Enabled = False
     Next
     cmbSucEmi.Enabled = FalseCant > 0  '?
     Me.cmbSucRec.Enabled = False
     Me.cmbFlete.Enabled = False
     Me.cmbUsuEmi.Enabled = False
     Me.fraAvance.Visible = True
     fraAvance.Refresh
     PGB.Min = 0: nreg = 0
     If AdoProduct.Recordset.RecordCount = 0 Then
        PGB.Max = 1
     Else
        PGB.Max = AdoProduct.Recordset.RecordCount
     End If
     Lstprod.Clear
     'PANADERIA
     If chkpan.Value Then
        Set rsttemppan = New ADODB.Recordset
        rsttemppan.Open "SELECT * FROM TFPRODUC WHERE CLAFAMIL = 'PAN' ", cn, adOpenKeyset, adLockOptimistic, adCmdText
        While Not rsttemppan.EOF
            If Not (IsNull(rsttemppan!medida)) And Not IsNull(rsttemppan!PAQUETES) Then
                 Lstprod.AddItem rsttemppan!descripc + " ( " + rsttemppan!medida + " )  ***PANANDERIA*** " _
                + " [ " + rsttemppan!CONSEC + " ]"
            End If
            rsttemppan.MoveNext
        Wend
        rsttemppan.Close
     End If
     'FRUTAS Y VERDURAS
     If chkfrutas.Value = 1 Then
        Set rsttempfru = New ADODB.Recordset
        rsttempfru.Open "SELECT * FROM TFPRODUC WHERE CLAFAMIL = 'FRU'", cn, adOpenKeyset, adLockOptimistic, adCmdText
        While Not rsttempfru.EOF
            If Not (IsNull(rsttempfru!medida)) And Not IsNull(rsttempfru!PAQUETES) Then
                 Lstprod.AddItem rsttempfru!descripc + " ( " + rsttempfru!medida + " ) *** FRUTAS Y VERDURAS ***  " _
                + " [ " + rsttempfru!CONSEC + " ]"
            End If
            rsttempfru.MoveNext
        Wend
        rsttempfru.Close
    End If
    ' SE AGREGAN GONDOLEROS
    Dim RSGON As ADODB.Recordset
    Set RSGON = New ADODB.Recordset
    cmbgon.Clear
    RSGON.Open "SELECT * FROM usuarios WHERE level1 = 'G'", cn, adOpenKeyset, adLockOptimistic, adCmdText
    While Not RSGON.EOF
           If Not (IsNull(RSGON!login)) Then
                 cmbgon.AddItem RSGON!Name + "[" + RSGON!clave + "]"
            End If
            RSGON.MoveNext
    Wend
    Set RSGON = Nothing
     'TODOS LOS DEMAS PRODUCTOS
     If chkpan.Value = 0 And chkfrutas.Value = 0 Then
        Do While Not AdoProduct.Recordset.EOF
              nreg = nreg + 1
              PGB.Value = nreg
              If Not (IsNull(AdoProduct.Recordset!medida)) And Not IsNull(AdoProduct.Recordset!PAQUETES) Then
                 Lstprod.AddItem AdoProduct.Recordset!descripc + " ( " + AdoProduct.Recordset!medida + " ) " _
                  + " CAJ: " + Str(AdoProduct.Recordset!InCant) + "  PZA: " + Str(AdoProduct.Recordset!InCantPza) + " [ " + AdoProduct.Recordset!CONSEC + " ]"
              End If
              AdoProduct.Recordset.MoveNext
              Loop
    End If
  fraAvance.Visible = False
  DbgrdDetTraAbi.Visible = True
  Me.LblProdAgr.Visible = True
  Me.StbMensajes.SimpleText = Space(12) & "F9 = Primer producto         F10 = Pagina Anterior           F11 = Pagina siguiente          F12 = Agregar Producto"
  End If
  chkTipoTrasl.Enabled = False

  PicBot.Visible = True
  cmdGrabar.Enabled = False
  cmdGrabarTra.Visible = True
  CmdActualizar.Visible = True
  cmdCodBarra.Visible = True
  cmdReporte.Visible = True
  cmdticket.Visible = True
  cmdReporte.Enabled = False
  cmdticket.Enabled = False

  CMDCORTAR.Visible = True
  If Me.chkTipoTrasl.Value = 1 Then
     DbgrdDetTraAbi.SetFocus
  ElseIf Not AdoTraslada.Recordset!t_entrada Then  'Traslado por pedido
        If AdoPedTra.Recordset.BOF And AdoPedTra.Recordset.EOF Then
           Me.cmbPedido.SetFocus
        Else
            CmdAgregar.Enabled = False
            cmbPedido.Enabled = False
            DesgTraPed  'Muestra el desglose del traslado cerrado o por pedido
        End If
    StbMensajes.SimpleText = cMenAnt
    StbMensajes.Refresh
  End If
  PonPie   'Pone Informacion de numero de product, cajas y piezas Exit Sub
'SE VALIDA SI ES UN AJUSTE AL INVENTARIO SEA SALIDA O ENTRADA
If Check1.Value = 1 Then
   If Not AdoTraslada.Recordset!t_enviado Then
      fraajuste.Enabled = True
      fraajuste.Visible = True
      txtobserva.SetFocus
   End If
End If

'SE PRETENDE ENVIAR A MINIINVENTARIOS
'SE    SELECCIONA LA ZONA A LA QUE SE V A ENVIAR
If txtcampos(3).Text = txtcampos(1).Text And AdoTraslada.Recordset!t_entrada = 0 Then
        RESP = MsgBox("El sistema a detectado un envio a Piso, " & vbCrLf & "Deseas Enviar a Una zona Especifica ? " & vbCrLf & "[NO=A PISO EN GENERAL] " & vbCrLf & vbCrLf & "[SI=SELECCIONA ZONA]" & vbCrLf, vbYesNo, "INVENTARIO DE PISO")
        If RESP = vbYes Then
            If OBTENZONAS Then
               'SE PRESENTA EL COMBO PARA SELECCIONAR
               FraZONAS.Enabled = True
               FraZONAS.Visible = True
               cmbzonas.SetFocus
            End If
        Else
            pisototal = 1 ' OSEA TODOS
        End If
End If
'SE PONE LA SEÑA DE QUE ES VOLUMEN

Call PONVOLUMEN
Exit Sub
Error:
   MsgBox Err.Description
End Sub

Private Sub PONVOLUMEN()
If Trim(AdoTraslada.Recordset!T_OBSERVA) = "VOLUMEN" Then
        Me.chkvolumen.Value = 1
Else
        chkvolumen.Value = 0
End If
End Sub

Private Sub UPLOAD_TRAS()
Set rsttemp = New ADODB.Recordset
rsttemp.Source = "SELECT * FROM PEDIDOS WHERE p_recibido = 0 AND p_sucursal = '" & txtcampos(3).Text & "' AND P_proveedor = 'ABA' AND p_traslado IS  NULL"
rsttemp.ActiveConnection = cCadConex
rsttemp.Open
While Not rsttemp.EOF
    cmbPedido.AddItem rsttemp!p_Pedido
    rsttemp.MoveNext
Wend
'Cargo los pedidos que se incluyen en el traslado, se utiliza la misma tabla
AdoPedTra.ConnectionString = cCadConex
AdoPedTra.CommandType = adCmdText
AdoPedTra.RecordSource = "SELECT MIN(dt_pedido) AS PEDIDOS FROM [DetalleTraslado]  WHERE dt_clave = '" & txtcampos(0).Text & "' GROUP BY DT_PEDIDO"
AdoPedTra.Refresh
dbgrdPedEnTra.Visible = True
FraPedEnv.Visible = True
dbgrdDetPed.Visible = True
dbgrdDetPed.Refresh
dbgrdPedEnTra.SetFocus
End Sub
Private Sub cmdGrabarTra_Click()
Dim rstInv As ADODB.Recordset
Dim rstTra As ADODB.Recordset
Dim rsttemp As ADODB.Recordset
Dim RstVeri As ADODB.Recordset
Dim RESP
Dim lTrans As Boolean
lTrans = False
On Error GoTo Error:
StbMensajes.SimpleText = "Ajustando Precios del Envio... "
Call cambiaprecios
StbMensajes.SimpleText = "Terminando de Ajustar Precios... "
cmen = StbMensajes.SimpleText
If Not AdoTraslada.Recordset!t_entrada Then    'Es salida de productos
   If chkEnviado.Value = False Then
        MsgBox "ES NECESARIO ACTIVAR LA CASILLA DE TRASLADO ENVIADO", vbExclamation
        Exit Sub
   End If
   cmensa = "Deseas grabar el detalle del traslado y DISMINUIR inventario"
End If
If AdoTraslada.Recordset!t_entrada Then  'Es entrada de productos
   If chkRecibido.Value = False And AdoTraslada.Recordset!t_entrada Then
        MsgBox "ES NECESARIO ACTIVAR LA CASILLA DE TRASLADO RECIBIDO", vbExclamation
        Exit Sub
   End If
   cmensa = "Deseas grabar el detalle del traslado y AUMENTAR Inventario"
End If
'Para que puedan grabar el traslado sin que exista pedido para imprimir el ticket en cero ya que no se puede
'reutilizar el folio para las tiendas
If Not AdoTraslada.Recordset!t_tipo Then
   If AdoPedTra.Recordset.BOF And AdoPedTra.Recordset.EOF Then
        'Obtengo datos informativos del detalle de traslado
        AdoDetPed.ConnectionString = cCadConex
        AdoDetPed.CommandType = adCmdText
        AdoDetPed.RecordSource = "SELECT detalleTraslado.dt_producto, tfproduc.descripc, LTrim(str(paquetes)) + ' X ' + LTRIM( str(contenid,10,3)) + ' ' + MEDIDA as MEDIDA, " _
        & " DetalleTraslado.dt_cantidad,DetalleTraslado.dt_cantidadP, DetalleTraslado.dt_pedido, Inventario.inprod, CAST(Inventario.inCant as SMALLMONEY) as TOTCAJ, dt_costo, dt_venta, paquetes, InCantPza FROM DetalleTraslado, tfproduc, Inventario WHERE DetalleTraslado.dt_CLAVE = '" & Trim(txtcampos(0).Text) & " '" _
        & "AND DetalleTraslado.dt_producto = tfproduc.consec AND detalleTraslado.dt_producto = Inventario.inProd ORDER BY Descripc,contenid"
        AdoDetPed.Refresh
   Else
        AdoDetPed.Refresh   'Para que no existan filas vacias
   End If
End If

RESP = MsgBox(cmensa, vbQuestion + vbYesNo)
If RESP = vbYes Then
   If (Not AdoTraslada.Recordset!t_frutas) And (Not AdoTraslada.Recordset!t_pan) Then
       Set RstVeri = New ADODB.Recordset
       'MsgBox "SELECT * FROM TFPRODUC,DETALLETRASLADO,INVENTARIO WHERE CONSEC = INPROD AND DT_PRODUCTO = CONSEC AND INPROD = DT_PRODUCTO AND (InCant < dt_cantidad OR (InCantPza < dt_cantidadp AND Incant = dt_cantidad ) AND DT_CLAVE = '" & txtCampos(0).Text & "' ORDER BY DESCRIPC,CONTENID"
       'verificando cajas
       CAD = "select descripc, contenid,medida,paquetes,incant,inprod,incantpza,dt_cantidad, dt_cantidadp from tfproduc,detalletraslado, inventario where  inprod = consec and inprod = dt_producto  and dt_clave = '" & Trim(txtcampos(0).Text) & "'"
       RstVeri.Open CAD, cn, adOpenKeyset, adLockOptimistic, adCmdText
       If Not AdoTraslada.Recordset!t_entrada Then
            While Not RstVeri.EOF
                prod = RstVeri!descripc & " " & RstVeri!PAQUETES & " x " & RstVeri!CONTENID & " " & RstVeri!medida
                If RstVeri!InCant < RstVeri!dt_cantidad Then
                        MsgBox "El producto " & prod & Chr(13) & " No tiene Existencia Suficiente Para surtir ", vbInformation, "INVENTARIO"
                        Exit Sub
                End If
                If (RstVeri!InCant * RstVeri!PAQUETES) + RstVeri!InCantPza < RstVeri!dt_cantidadp Then
                        MsgBox "El producto " & prod & " No tiene Existencia Suficiente Para surtir ", vbInformation, "INVENTARIO"
                        Exit Sub
                End If
                RstVeri.MoveNext
             Wend
       End If
       'RstVeri.Open "SELECT * FROM TFPRODUC,DETALLETRASLADO,INVENTARIO WHERE CONSEC = INPROD AND DT_PRODUCTO = CONSEC AND INPROD = DT_PRODUCTO AND (InCant < dt_cantidad OR (InCantPza < dt_cantidadp AND Incant <= dt_cantidad )) AND DT_CLAVE = '" & txtCampos(0).Text & "' ORDER BY DESCRIPC,CONTENID", cn, adOpenKeyset, adLockOptimistic, adCmdText
       'If RstVeri.RecordCount > 0 And Not AdoTraslada.Recordset!t_entrada Then
        '   MsgBox "EL INVENTARIO ES MENOR A LO SOLICTADO EN EL SIGUIENTE ARTICULO:" & Chr(13) _
        '        & RstVeri!descripc & " " & CStr(RstVeri!Paquetes) & " X" & CStr(RstVeri!Contenid) & " " & RstVeri!Medida, vbCritical
        '   Exit Sub
       'End If
   End If
   'En el caso que dejen vacias cantidades lo igualo a cero para que no marque error en procedimientos posteriores
   cn.Execute "UPDATE detalletraslado SET dt_cantidadp = 0 WHERE dt_cantidadP IS NULL AND dt_clave = '" & txtcampos(0).Text & "'"
   cn.Execute "UPDATE detalletraslado SET dt_cantidad = 0 WHERE dt_cantidad IS NULL AND dt_clave = '" & txtcampos(0).Text & "'"
   cn.BeginTrans: lTrans = True
   'Cargo el detalle de pedidos enviados en el traslado
   Set rstTra = New ADODB.Recordset
   rstTra.Open "SELECT * FROM DetalleTraslado WHERE dt_clave = '" & txtcampos(0).Text & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
   
   If Not (AdoDetPed.Recordset.BOF And AdoDetPed.Recordset.EOF) Then AdoDetPed.Recordset.MoveFirst
   While Not AdoDetPed.Recordset.EOF
      StbMensajes.SimpleText = Space(30) + "Disminuyendo del inventario el producto " & AdoDetPed.Recordset!Dt_producto & AdoDetPed.Recordset!descripc
      StbMensajes.Refresh
      If AdoTraslada.Recordset!t_entrada Then
         cCadena = "UPDATE Inventario SET Incant = InCant + " & AdoDetPed.Recordset!dt_cantidad & ", InCantpza = InCantPza + " & AdoDetPed.Recordset!dt_cantidadp & " WHERE Inprod = '" & AdoDetPed.Recordset!Dt_producto & "'"
      Else
         ' SE ENVIA Y SE DEBE DISMINUIR
         If AdoDetPed.Recordset!dt_cantidadp > 0 Then
            If AdoDetPed.Recordset!PAQUETES = 1 Then
               'Si la presentacion es 1x1 a la existencia en caja le disminuyo lo solicitado en caja  y pieza que es lo mismo
               cCadena = "UPDATE Inventario SET Incant = InCant - " & AdoDetPed.Recordset!dt_cantidad & ", Incantpza = InCantpza - " & AdoDetPed.Recordset!dt_cantidadp & " WHERE Inprod = '" & AdoDetPed.Recordset!Dt_producto & "'"
            Else
               If AdoDetPed.Recordset!InCantPza < AdoDetPed.Recordset!dt_cantidadp Then
                  'Disminuyo una caja del inventario y la convierto en piezas
                  cn.Execute "UPDATE inventario SET incantpza = incantpza + paquetes, incant = incant - 1 FROM tfproduc WHERE consec = inprod AND Inprod = '" & AdoDetPed.Recordset!Dt_producto & "'"
               End If
               cCadena = "UPDATE Inventario SET Incant = InCant - " & AdoDetPed.Recordset!dt_cantidad & ", IncantPza = InCantPza - " & AdoDetPed.Recordset!dt_cantidadp & " WHERE Inprod = '" & AdoDetPed.Recordset!Dt_producto & "'"
            End If
         Else
            cCadena = "UPDATE Inventario SET Incant = InCant - " & AdoDetPed.Recordset!dt_cantidad & " WHERE Inprod = '" & AdoDetPed.Recordset!Dt_producto & "'"
         End If
          
      End If
      'MsgBox ccadena
      cn.Execute cCadena
      AdoDetPed.Recordset.MoveNext
   Wend
   Set rsttemp = New ADODB.Recordset
   If lfranquicia Then
      rsttemp.Open "SELECT SUM(dt_cantidad * dt_venta) + SUM(dt_cantidadp * dt_ventap )  AS totTrasl,  SUM(dt_cantidad * dt_venta) + SUM(dt_cantidadp * dt_ventap )as montra FROM DetalleTraslado,TFPRODUC WHERE consec = dt_producto AND DetalleTraslado.dt_clave = '" & txtcampos(0).Text & "'", cCadConex, adOpenKeyset, adLockOptimistic, adCmdText
      cn.Execute "UPDATE DETALLETRASLADO SET dt_importe = (dt_cantidad * dt_venta) + (dt_cantidadp * dt_ventap) WHERE dt_clave = '" & Trim(txtcampos(0).Text) & "'"
   Else 'Si es tienda
      rsttemp.Open "SELECT SUM(dt_cantidad * dt_costo) + SUM(dt_cantidadp * dt_costop ) AS totTrasl,  SUM(dt_cantidad * dt_venta) + SUM(dt_cantidadp * dt_ventap )as montra FROM DetalleTraslado,TFPRODUC WHERE consec = dt_producto AND DetalleTraslado.dt_clave = '" & txtcampos(0).Text & "'", cCadConex, adOpenKeyset, adLockOptimistic, adCmdText
      cn.Execute "UPDATE DETALLETRASLADO SET dt_importe = (dt_cantidad * dt_costo) + (dt_cantidadp * dt_costop) WHERE dt_clave = '" & Trim(txtcampos(0).Text) & "'"
   End If
   Set rstemp1 = New ADODB.Recordset 'FOLIO DE TIENDA
   rstemp1.Open "select max(t_foliotie) as folio from traslados where t_sucursalreceptor =  '" & Trim(txtcampos(3).Text) & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
   nfoltie = 0
   If rstemp1.EOF Then
                nfoltie = 1
   Else
                nfoltie = rstemp1!Folio + 1
   End If
   If IsNull(nfoltie) Then
               nfoltie = 0
   End If
   rstemp1.Close
   CAD = "UPDATE traslados SET t_enviado = 1, t_foliotie = " & nfoltie & ",t_costo = " & IIf(IsNull(rsttemp!tottrasl), 0, rsttemp!tottrasl) & " , t_monto = " & IIf(IsNull(rsttemp!montra), 0, rsttemp!montra) & " WHERE t_clave = '" & Trim(txtcampos(0).Text) & "'"
   cn.Execute CAD
   'CONDICION PARA AUMENTAR EL INVENTARIO DE  PISO
   If txtcampos(3).Text = txtcampos(1).Text Then
      If Not AdoTraslada.Recordset!t_entrada Then
         If lpprov Then
            Call aumentazona
            CAD = "update traslados set  t_observa = 'PAQUETERIA' where  t_clave = '" & Trim(txtcampos(0).Text) & "'"
            cn.Execute CAD
         Else
            Call aumentapiso
            End If
        End If
      End If
   'nfoltie = txtFoliotie.Text
   'cn.Execute "UPDATE traslados SET t_enviado = 1, t_foliotie = " & nfoltie & ",t_costo = " & IIf(IsNull(rsttemp!tottrasl), 0, rsttemp!tottrasl) & " WHERE t_clave = '" & Trim(txtCampos(0).Text) & "'"
   txtcampos(4).Text = IIf(IsNull(rsttemp!tottrasl), 0, rsttemp!tottrasl)
   cn.CommitTrans
   StbMensajes.SimpleText = cmen
   StbMensajes.Refresh
   Me.cmdCodBarra.Enabled = True
   Me.cmdGrabarTra.Enabled = False
   chkEnviado.Enabled = False: chkRecibido.Enabled = False
   DbgrdDetTraAbi.AllowUpdate = False
   dbgrdDetPed.AllowUpdate = False
   DbgrdDetTraAbi.AllowDelete = False
   DbgrdDetTraAbi.Columns(0).Button = False
   cmdReporte.Enabled = True
   cmdticket.Enabled = True
   CMDCORTAR.Enabled = True
   cmdregresar.SetFocus
   txtFoliotie.Text = nfoltie
   txtFoliotie.Refresh
   'cad = "UPDATE traslados SET t_fecha = '" & Date + Time & "' WHERE t_clave = '" & Trim(txtcampos(0).Text) & "'"
   'cn.Execute cad
   manual = False
End If
'se pide que seleccione el gondolero
If txtcampos(1).Text = txtcampos(3).Text And AdoTraslada.Recordset!t_entrada = 0 Then
    frmgon.Enabled = True
    frmgon.Visible = True
    cmbgon.SetFocus
End If
CAD = "UPDATE traslados SET t_fecha = '" & date + Time & "' ,t_enviado = 1 WHERE t_clave = '" & Trim(txtcampos(0).Text) & "'"
cn.Execute CAD
Exit Sub
Error:
   MsgBox "OCURRIO EL SIGUIENTE ERROR: " + Chr(13) + UCase(Err.Description), vbCritical
   If lTrans Then
      MsgBox "A CONTINUACION SE DESHARAN LAS MODIFICACIONES REALIZADAS AL INVENTARIO", vbCritical
      cn.RollbackTrans
   End If
End Sub

Private Function OBTENZONAS() As Boolean
Dim rsttemp As ADODB.Recordset
Set rsttemp = New ADODB.Recordset
CAD = "select * from zonas "
rsttemp.Open CAD, cn, adOpenKeyset, adLockOptimistic, adCmdText
If rsttemp.EOF Then
   MsgBox "EL inventario No contiene el Organizacion de Zonas", vbInformation, "INVENTARIO PISO"
   OBTENZONAS = False
   Exit Function
Else
   cmbzonas.Clear
   While Not rsttemp.EOF
         cmbzonas.AddItem rsttemp!zdes & " [" & rsttemp!zclave & "]"
         rsttemp.MoveNext
   Wend
   OBTENZONAS = True
End If
End Function

Private Sub aumentapiso()
'On Error GoTo error:
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
AdoDetPed.Recordset.MoveFirst
While Not AdoDetPed.Recordset.EOF
    xproducto = AdoDetPed.Recordset!Dt_producto
    rs.Open "select consec,paquetes  from tfproduc where consec  = '" & Trim(xproducto) & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Not rs.EOF Then
       paq = rs!PAQUETES
    Else
       paq = 1
    End If
    cant = AdoDetPed.Recordset!dt_cantidad * paq + AdoDetPed.Recordset!dt_cantidadp
    rs.Close
    rs.Open "select * FROM INVENTARIOPISO  where INPROD  = '" & Trim(xproducto) & "'", cn, adOpenDynamic, adLockOptimistic, adCmdText
    If rs.EOF Then
        rs.AddNew
        rs!Inprod = xproducto
        rs!InCant = 0
        rs!InCantPza = cant
        rs.Update
    Else
        rs!InCantPza = InCantPza + cant
        rs.Update
    End If
    
    'CAD = "update inventariopiso set incantpza = incantpza + " & cant & " where inprod = '" & Trim(xproducto) & "'"
    'MsgBox CAD
    'cn.BeginTrans
    'cn.Execute CAD
    'cn.CommitTrans
    rs.Close
    AdoDetPed.Recordset.MoveNext
Wend
'MsgBox txtCampos(0).Text
'cad = "UPDATE INVENTARIOPISO SET incantpza =  incantpza + dt_cantidadp + (dt_cantidad * p.paquetes) from tfproduc p, detalletraslado as d where consec = d.dt_producto and d.dt_producto = INPROD and dt_clave =  '" & Trim(txtCampos(0).Text) & "'"
'MsgBox cad
'cn.Execute cad
Exit Sub
Error:
  MsgBox "Error al Actualizar Piso", vbCritical
End Sub

Private Sub aumentazona()
On Error GoTo Error:
' SE VA A HACER EN FORMA DE REGESITROS POR SI NO EXISTE ALGUN PRODUCTO
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
AdoDetPed.Recordset.MoveFirst
While Not AdoDetPed.Recordset.EOF
    rs.Open "select * from inventpaq where inprod = '" & Trim(AdoDetPed.Recordset!Dt_producto) & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
    If rs.EOF Then
        'se agrega
        CAD = "insert into inventpaq(inprod,incantpza,incant) values(" & _
              Trim(AdoDetPed.Recordset!Dt_producto) & "," & AdoDetPed.Recordset!dt_cantidad & "," & AdoDetPed.Recordset!dt_cantidadp & ")"
    Else
        'se actualiza
        CAD = "update inventpaq set incant = incant + " & AdoDetPed.Recordset!dt_cantidad & " , incantpza = incantpza +  " & AdoDetPed.Recordset!dt_cantidadp & _
              " where inprod = '" & Trim(AdoDetPed.Recordset!Dt_producto) & "'"
    End If
    cn.Execute CAD
    rs.Close
    AdoDetPed.Recordset.MoveNext
Wend
Set rs = Nothing
Exit Sub
Error:
  MsgBox "Error al Actualizar Piso", vbCritical
End Sub


Private Sub cmdInven_Click()
Dim cmensa As String
  nOp = 31  'se llama al inventario desde traslados
  cmensa = StbMensajes.SimpleText
  StbMensajes.SimpleText = Space(50) & "Espere un momento obteniendo existencia de productos........"
  StbMensajes.Refresh
  frmModInv.Show 1
  StbMensajes.SimpleText = cmensa
  StbMensajes.Refresh
End Sub

Private Sub cmdprecios_Click()
Call cambiaprecios
End Sub

Private Sub cambiaprecios()
Dim lTrans As Boolean
'On Error GoTo Error:
'If MsgBox("DESEAS ACTUALIZAR EL TRASLADO CON PRECIOS ACTUALES", vbQuestion + vbYesNo) = vbYes Then
    If lfranquicia Then
        i = 4
    Else
        i = 1
    End If
    Select Case i
    Case 1
        CADENA = "update detalletraslado set dt_venta = precio2  , dt_ventap =  precio2 / paquetes from tfproduc ,preprod where dt_producto = consec and dt_producto = preclave and dt_clave =  '" & Trim(txtcampos(0).Text) & "'"
    Case 2
        CADENA = "update detalletraslado set dt_venta = precio2  , dt_ventap =  precio2 / paquetes from tfproduc ,preprod where dt_producto = consec and dt_producto = preclave and dt_clave =  '" & Trim(txtcampos(0).Text) & "'"
    Case 3
        CADENA = "update detalletraslado set dt_venta = precio3   , dt_ventap =  precio3 / paquetes from tfproduc ,preprod where dt_producto = consec and   dt_producto = preclave and dt_clave =  '" & Trim(txtcampos(0).Text) & "'"
    Case 4
        CADENA = "update detalletraslado set dt_venta = precio4   , dt_ventap =  precio4 / paquetes from tfproduc, preprod where dt_producto = consec and dt_producto = preclave and dt_clave =  '" & Trim(txtcampos(0).Text) & "'"
    Case Else
        MsgBox "Debio Escribir un numero del 2 al 4 "
        Exit Sub
    End Select
    cn.BeginTrans
    lTrans = True
    cn.Execute CADENA
    cn.CommitTrans
    ' SE INCLUYE LA PARTE DE IVA Y IEPS EN EL DETALLE TRASLADO QUE VIENEN DEL TFPRODUC
    'CADENA1 = "update detalletraslado set dt_costo = precosto , dt_costop = precosto / paquetes from tfproduc where consec = dt_producto and dt_clave = '" & Trim(txtCampos(0).Text) & "'"
    CADENA1 = "update detalletraslado set dt_costo = precosto , dt_costop = precosto / paquetes, DT_IVA = IVA, DT_IEPS = IEPS, DT_TASAIEPS = TASAIEPS from tfproduc where consec = dt_producto and dt_clave = '" & Trim(txtcampos(0).Text) & "'"
    'CAD2 = "update traslados set t_costo = "
'End If
cn.BeginTrans
cn.Execute CADENA1
If tipotienda = 2 Then cn.Execute " ACT_TRASLADO '" & Trim(txtcampos(0).Text) & "'"
cn.CommitTrans
Exit Sub
Error:
  MsgBox "Se ha generado un error al momento de actualizar los precios, Es probable que Otro usuario este accesando Este Envio; salga del envio e intente de nuevo", vbInformation
  If lTrans Then cn.RollbackTrans
End Sub

Private Sub cmdRegCan_Click()
  Unload Me
End Sub

Private Sub cmdRegresar_Click()
   AdoTraslada.Refresh
   Unload Me
End Sub

Private Sub cmdReporte_Click()
On Error GoTo Error
 cMensaje = StbMensajes.SimpleText
 StbMensajes.SimpleText = Space(90) + "Espere un momento generando reporte..."
 StbMensajes.Refresh
 Rpt.Connect = cCadConex
 'Rpt.ReportFileName = App.Path & IIf(AdoTraslada.Recordset!t_tipo = 0, "\TraslRec.rpt", "\Trasabto.rpt")
 Rpt.ReportFileName = App.Path & "\Trasabto.rpt"
 Rpt.WindowTitle = "Traslado con folio " & txtcampos(0).Text
 Rpt.Formulas(0) = "FORMSELEC = '" & txtcampos(0).Text & "'"
 Rpt.Formulas(1) = "TRASLADO= 'ENVIO DEL TRASLADO CON FOLIO " & txtcampos(0).Text & "'"
 Rpt.Formulas(2) = "SUCEMI= 'SUCURSAL EMISORA:    " & txtcampos(1).Text & Space(3) & cmbSucEmi.Text & "'"
 Rpt.Formulas(3) = "SUCREC= 'SUCURSAL RECEPTORA:    " & txtcampos(3).Text & Space(3) & cmbSucRec.Text & "'"
 Rpt.Action = 1
 StbMensajes.SimpleText = cMensaje
 StbMensajes.Refresh
 Exit Sub
Error:
   MsgBox Err.Description
End Sub

Private Sub cmdTicket_Click()
Dim rsttemp As ADODB.Recordset
Dim nTotal
Dim nCajas
Dim ProEnTick
Dim nTicket As Integer
Dim nTotTick As Integer
Dim cCad
Dim lfran As Boolean
Dim NVENTA1 As Currency
Dim NVENTA2 As Currency
Dim NIVA2 As Currency
Dim NVENTA3 As Currency
Dim NIVA3 As Currency
Dim NIEPS3 As Currency
Dim NVENTA4 As Currency
Dim NIVA4 As Currency
Dim NIEPS4 As Currency
Dim lResumen As Boolean
lResumen = False  'Por default se pone la venta por depto por ticket y no por envio
'SAVEMC
If tipotienda = 4 Then
   If Not verImpticket Then Exit Sub
End If
'On Error GoTo ERROR:
nAncho = 250    'En puntos es el ancho del ticket de las Miniprinter
ProEnTick = 200  'Numero de productos que se imprimen en el ticket
Printer.ScaleMode = vbPoints
Set rsttemp = New ADODB.Recordset
rsttemp.Open "SELECT COUNT(*) AS TOTPRO FROM DETALLETRASLADO WHERE DetalleTraslado.dt_clave = '" & txtcampos(0).Text & "'", cn, adOpenDynamic, adLockOptimistic, adCmdText
nTotTick = IIf(Round(rsttemp!totpro / ProEnTick, 0) < rsttemp!totpro / ProEnTick, Round(rsttemp!totpro / ProEnTick, 0) + 1, Round(rsttemp!totpro / ProEnTick, 0))
rsttemp.Close
'SOLO SE TOMAN EN CUENTA LOS PRODUCTOS QUE TIENE CANTIDADES SOLICITADAS
CAD = "SELECT * FROM DetalleTraslado, Tfproduc WHERE DetalleTraslado.DT_producto = TFPRODUC.Consec AND DetalleTraslado.dt_clave = '" & txtcampos(0).Text & "' ORDER BY Descripc"
rsttemp.Open CAD, cn, adOpenKeyset, adLockReadOnly, adCmdText
'rsttemp.Open "SELECT * FROM DetalleTraslado, Tfproduc WHERE (dt_cantidad > 0 or dt_cantidadp > 0 ) and DetalleTraslado.DT_producto = TFPRODUC.Consec AND DetalleTraslado.dt_clave = '" & txtCampos(0).Text & "' ORDER BY Descripc", cn, adOpenKeyset, adLockOptimistic, adCmdText
nTotal = 0: nCajas = 0: nPiezas = 0: nProd = 0: nPeso = 0
lNoImp = True: nProd = ProEnTick: nTicket = 0
'Cuando esta vacio el detalle del traslado para que se imprima el encabezado
If rsttemp.BOF And rsttemp.EOF Then
   Encabezado nTicket, nTotTick  'Imprime encabezado
End If
'Exit Sub
Do While (Not rsttemp.EOF)
    If nProd = ProEnTick Then
       If Not lNoImp Then
          Printer.Print "-------------------------------------"
          Printer.Print ""
          Printer.Print ""
          Printer.CurrentX = 2
          Printer.Print "IMPTE:" & Format(nTotal, "$ ###,###,###.00")
          Printer.Print " "
          Printer.Print "CAJAS : " & Format(nCajas, "#,###,##0")
          Printer.Print ""
          Printer.Print "PIEZAS: " & Format(nPiezas, "#,###,##0")
          Printer.Print "KILOS : " & Format(nPeso, "#,###,###.00")
          Printer.Print " "
          Printer.Print " "
          Printer.Print " "
          Printer.Print " "
          Printer.Print " ---------------        ----------------"
          Printer.Print "      ENVIA                  RECIBE     "
          Printer.Print " "
          Printer.Print " "
          If lfranquicia Then Vtadepto NVENTA1, NVENTA2, NIVA2, NVENTA3, NIVA3, NIEPS3, NVENTA4, NIVA4, NIEPS4
          For N = 0 To 11
              Printer.Print " "
          Next
          Printer.EndDoc
          'MsgBox "CUANDO TERMINE DE IMPRIMIR EL TICKET " & CStr(nTicket) & Chr(13) & "PRESIONE ENTER PARA CORTAR EL PAPEL", vbInformation
          'CmdCortar_Click
       End If
       nTicket = nTicket + 1
       cresp = MsgBox("UNA VEZ CORTADO EL PAPEL, SELECCIONE LAS SIGUIENTES OPCIONES" & Chr(13) & Chr(13) & _
               "[SI  ]               Continuar con la impresion del ticket " & CStr(nTicket) & " de " & CStr(nTotTick) & Chr(13) & _
               "[NO]               Omitir la impresion del ticket: " & CStr(nTicket) & Chr(13) & _
               "[CANCELAR] Cancelar la impresion", vbInformation + vbYesNoCancel)
       If cresp = vbYes Then 'Imprimir el ticket
          Encabezado nTicket, nTotTick  'Encabezado
          lNoImp = False
       ElseIf cresp = 2 Then  'Cancelar la impresion
          Printer.EndDoc
          cmdregresar.SetFocus
          Exit Sub
       Else                   'Omitir la impresion del siguiente ticket
          lNoImp = True
       End If
       nTotal = 0: nCajas = 0: nPiezas = 0: nProd = 0: nPeso = 0
       nProd = 0
       If Not lResumen Then
          NVENTA1 = 0: NIVA1 = 0: NIEPS1 = 0 'Save
          NVENTA2 = 0: NIVA2 = 0: NIEPS2 = 0  'Save
          NVENTA3 = 0: NIVA3 = 0: NIEPS3 = 0  'Save
          NVENTA4 = 0: NIVA4 = 0: NIEPS4 = 0  'Save
       End If

    End If
    'Si tiene cantidad en cajas y se ha deseado imprimir el ticket
    If Not lNoImp And (rsttemp!dt_cantidad > 0 Or rsttemp!dt_cantidadp > 0) Then
        cCad = CStr(rsttemp!dt_cantidad) & " CAJ, " & CStr(rsttemp!dt_cantidadp) & " PZA " & Trim(rsttemp!descripc)
        'En caso de que sea muy grande la descripcion se imprime en dos lineas
        Printer.Print " "
        If Len(Trim(cCad)) > 37 Then
           Printer.Print Mid(cCad, 1, 36)
           'If nProd = ProEnTick - 1 Then Printer.Print " "  'Quiensabe porque se encimaba en el ultimo producto en algunas impresoras
           Printer.Print Mid(cCad, 37, 18)
        Else
           Printer.Print Mid(cCad, 1, 36)
        End If
        'If nProd = ProEnTick - 1 Then Printer.Print " "  'Quiensabe porque se encimaba en el ultimo producto en algunas impresoras
        'Printer.Print " "
        If IsNull(rsttemp!barraspza) Then
             Barras = 0
        Else
             Barras = rsttemp!barraspza
        End If
        Printer.Print "-" & Barras & "-  ";
        Printer.CurrentX = 165
        Printer.Print "  " & CStr(rsttemp!PAQUETES) & " X " & CStr(rsttemp!CONTENID) & " " & Mid(rsttemp!medida, 1, 4)
        'If nProd = ProEnTick - 1 Then Printer.Print " "  'Quiensabe porque se encimaba en el ultimo producto en algunas impresoras
        nprecio = IIf(lfranquicia, rsttemp!DT_venta, rsttemp!DT_costo)
        NPRECIOP = IIf(lfranquicia, rsttemp!DT_ventap, rsttemp!DT_costo / rsttemp!PAQUETES)
        'n = nAncho - (Printer.TextWidth(Format(rsttemp!dt_costo, "###,###,##0.00") & Space(3) & Format((rsttemp!dt_costo * rsttemp!dt_cantidad) + (rsttemp!dt_costo / rsttemp!PAQUETES * rsttemp!dt_cantidadp), "$###,###,##0.00")))
        N = nAncho - (Printer.TextWidth(Format(nprecio, "###,###,##0.00") & Space(3) & Format((nprecio * rsttemp!dt_cantidad) + (NPRECIOP * rsttemp!dt_cantidadp), "$###,###,##0.00")))
        Printer.CurrentX = N
        'Printer.Print Format(rsttemp!dt_costo, "###,###,##0.00") & Space(3) & Format(rsttemp!dt_costo * rsttemp!dt_cantidad + (rsttemp!dt_costo / rsttemp!PAQUETES * rsttemp!dt_cantidadp), "$###,###,##0.00")
        Printer.Print Format(nprecio, "###,###,##0.00") & Space(3) & Format(nprecio * rsttemp!dt_cantidad + (NPRECIOP * rsttemp!dt_cantidadp), "$###,###,##0.00")
        
        If Not IsNull(rsttemp!DT_costo) Then  'Verifico costo
           'nCosto = rsttemp!dt_cantidad * rsttemp!dt_costo + (rsttemp!dt_cantidadp * (rsttemp!dt_costo / rsttemp!PAQUETES))
           ncosto = rsttemp!dt_cantidad * nprecio + (rsttemp!dt_cantidadp * NPRECIOP)
           nTotal = nTotal + ncosto
           If lfranquicia And Val(txtcampos(3).Text) <> 5 And Val(txtcampos(3).Text) <> 12 Then  'Si es franquicia se desglosa la venta por departamentos
              If rsttemp!iva = 0 And rsttemp!ieps = 0 Then        'Depto 1
                 NVENTA1 = NVENTA1 + ncosto
                 NIVA1 = 0
                 NIEPS1 = 0
              ElseIf rsttemp!iva = 15 And rsttemp!ieps = 0 Then   'Depto 2
                 NVENTA2 = NVENTA2 + ncosto
                 NIVA2 = NIVA2 + (ncosto / 1.15 * (15 / 100))
                 NIEPS2 = 0
              ElseIf rsttemp!iva = 15 And rsttemp!ieps = 25 Then  'Depto 3
                 NVENTA3 = NVENTA3 + ncosto
                 NIVA3 = NIVA3 + (ncosto / 1.15 * (15 / 100))
                 NIEPS3 = NIEPS3 + (((ncosto / 1.15) / 1.25) * 25 / 100)
              ElseIf rsttemp!iva = 15 And rsttemp!ieps >= 25 Then 'Depto 4
                 NVENTA4 = NVENTA4 + ncosto
                 NIVA4 = NIVA4 + (ncosto / 1.15 * (15 / 100))
                 NIEPS4 = NIEPS4 + (((ncosto / 1.15) / 1.3) * 30 / 100)
              End If
           End If
        End If
        If Not IsNull(rsttemp!peso) Then  'Verifico peso
           nPeso = nPeso + rsttemp!peso
        End If
        
    End If
    nCajas = nCajas + rsttemp!dt_cantidad
    nPiezas = nPiezas + rsttemp!dt_cantidadp
    If lNoImp Then
       For N = 1 To ProEnTick
           rsttemp.MoveNext
           If rsttemp.EOF Then
              MsgBox "YA NO EXISTEN MAS TICKET'S PARA IMPRIMIR", vbInformation
              Exit Sub
           End If
           nProd = nProd + 1
       Next
    Else
       rsttemp.MoveNext
       nProd = nProd + 1
    End If
Loop
Printer.CurrentX = 2
Printer.Print "-------------------------------------"
Printer.Print ""
Printer.Print ""
Printer.Print "IMPTE:" & Format(nTotal, "$ ###,###,###.00")
Printer.Print " "
Printer.Print "CAJAS : " & Format(nCajas, "#,###,##0")
Printer.Print ""
Printer.Print "PIEZAS: " & Format(nPiezas, "#,###,##0")
Printer.Print "KILOS : " & Format(nPeso, "#,###,###.00")
Printer.Print " "
Printer.Print " "
Printer.Print " "
Printer.Print " "
Printer.Print " ---------------        ----------------"
Printer.Print "      ENVIA                  RECIBE     "
Printer.Print " "
Printer.Print " "
'If lfranquicia Then Vtadepto nVENTA1, nVENTA2, nIVA2, NVENTA3, NIVA3, NIEPS3, NVENTA4, NIVA4, NIEPS4
If lfranquicia Then Vtadepto NVENTA1, NVENTA2, NIVA2, NVENTA3, NIVA3, NIEPS3, NVENTA4, NIVA4, NIEPS4
For N = 0 To 11
    Printer.Print " "
Next
cmdregresar.SetFocus
'Printer.Print Chr(27) + Chr(64)
'Printer.Print Chr(27) + Chr(105)
Printer.EndDoc
'CmdCortar_Click
Exit Sub
Error:
  MsgBox Err.Description
End Sub

Private Sub Encabezado(nTick As Integer, nTTick As Integer)
  Printer.Print " "
  If ZONA = "OAX" Then
     Printer.Print "        VIVERES Y LICORES S.A DE C.V."
       Printer.Print "CARBONERA 1016 COL. TRINIDAD DE H."
       Printer.Print "   OAXACA, OAX. " & date & "  " & Format(Time, "HH:MM:SS")
  Else
     Printer.Print "HOLDING MEXICO CENTRO AMERICA SA DE CV"
     Printer.Print "4A. AV. SUR, PROLONGACION # 175   "
     Printer.Print "TAPACHULA, CHIS." & date & "  " & Format(Time, "HH:MM:SS")
  End If
  Printer.Print IIf(cModo = "DEVO", "DEV. A: ", "ENVIO A: "); Mid(Trim(Me.cmbSucRec.Text), 1, 18)
  Printer.Print "FOL.UNICO: " & txtcampos(0).Text & "   FOL.TIENDA: " & Trim(txtFoliotie.Text)
  Printer.Print "TICKET NUM: "; CStr(nTick) & " DE " & CStr(nTTick)
  Printer.Print "-------------------------------------"
End Sub

Private Sub Vtadepto(NVENTA1 As Currency, NVENTA2 As Currency, NIVA2 As Currency, NVENTA3 As Currency, NIVA3 As Currency, NIEPS3 As Currency, NVENTA4 As Currency, NIVA4 As Currency, NIEPS4 As Currency)
Printer.Print "======================================"
Printer.Print ""
Printer.Print "   TOTAL DE VENTAS POR DEPARTAMENTO   "
Printer.Print "TOTAL           IVA           IEPS"
Printer.Print "VENTA        TRASLADADO    TRASLADADO"
Printer.Print "--------------------------------------"
Printer.Print "DEPTO.  1"
N = nAncho - (150 + Printer.TextWidth(Format(NVENTA1, "$###,###,##0.00")))
Printer.CurrentX = N
Printer.Print Format(NVENTA1, "$###,###,##0.00")
Printer.Print " "

Printer.Print "DEPTO.  2 (IVA 15%)"
N = nAncho - (150 + Printer.TextWidth(Format(NVENTA2, "$###,###,##0.00")))
Printer.CurrentX = N
Printer.Print Format(NVENTA2, "$###,###,##0.00");
N = nAncho - (60 + Printer.TextWidth(Format(NIVA2, "$###,###,##0.00")))
Printer.CurrentX = N
Printer.Print Format(NIVA2, "$###,###,##0.00")
Printer.Print " "

Printer.Print "DEPTO.  3 (IVA 15% IEPS 25%)"
N = nAncho - (150 + Printer.TextWidth(Format(NVENTA3, "$###,###,##0.00")))
Printer.CurrentX = N
Printer.Print Format(NVENTA3, "$###,###,##0.00");
N = nAncho - (60 + Printer.TextWidth(Format(NIVA3, "$###,###,##0.00")))
Printer.CurrentX = N
Printer.Print Format(NIVA3, "$###,###,##0.00");
N = nAncho - (Printer.TextWidth(Format(NIEPS3, "$###,###,##0.00")))
Printer.CurrentX = N + 30
Printer.Print Format(NIEPS3, "$###,###,##0.00")
Printer.Print " "

Printer.Print "DEPTO.  4 (IVA 15% IEPS 30%)"
N = nAncho - (150 + Printer.TextWidth(Format(NVENTA4, "$###,###,##0.00")))
Printer.CurrentX = N
Printer.Print Format(NVENTA4, "$###,###,##0.00");
N = nAncho - (60 + Printer.TextWidth(Format(NIVA4, "$###,###,##0.00")))
Printer.CurrentX = N
Printer.Print Format(NIVA4, "$###,###,##0.00");
N = nAncho - (Printer.TextWidth(Format(NIEPS4, "$###,###,##0.00")))
Printer.CurrentX = N + 30
Printer.Print Format(NIEPS4, "$###,###,##0.00")
Printer.Print " "

End Sub


Private Sub Command1_Click()
FraZONAS.Enabled = False
FraZONAS.Visible = False
lpprov = True
End Sub

Private Sub Command2_Click()
Dim rs As ADODB.Recordset
If Not verImpresora Then Exit Sub
nAncho = 250    'En puntos
Printer.ScaleMode = vbPoints
Printer.CurrentX = 10
espacio = "   "
Set rs = New ADODB.Recordset
Me.StbMensajes.SimpleText = "Espere Imprimiendo Etiquetas..."
rs.Open "SELECT dt_cantidad,dt_cantidadp,paquetes, NOMCORTO,barraspza, ltrim(str(contenid,10,3)) + ' x ' + medida as medida FROM tfproduc,detalletraslado WHERE SUBSTRING(LTRIM(STR(barraspza,15)),1,3) = '777' and dt_clave = '" & Trim(txtcampos(0).Text) & "' and consec = dt_producto", cn, adOpenKeyset, adLockOptimistic, adCmdText
While Not rs.EOF
    For i = 1 To rs!dt_cantidad + rs!PAQUETES + rs!dt_cantidadp
        Printer.Font = "arial"
        Printer.FontSize = "8"
        Printer.Print espacio & Trim(rs!NOMCORTO)
        Printer.Print espacio & rs!medida & Space(5) & cmbSucRec.Text
        Printer.Font = "ZB 39* 10mil/2:1"
        Printer.FontSize = 40
        Printer.CurrentX = 10
        Printer.Print (rs!barraspza)
        Printer.EndDoc
    Next
    rs.MoveNext
Wend
Me.StbMensajes.SimpleText = "Fin de Impresion de Etiquetas..."
End Sub

Private Sub dbgrdDetPed_AfterUpdate()
Dim rsttemp As ADODB.Recordset
On Error Resume Next
If Me.chkTipoTrasl.Value = 0 Then
   AdoDetPed.Refresh
   AdoDetPed.Recordset.Bookmark = clave
   PonPie
End If
If AdoTraslada.Recordset!t_enviado Then
    Set rsttemp = New ADODB.Recordset
    rsttemp.Open "SELECT SUM(dt_cantidad * dt_costo) + SUM(dt_cantidadp * dt_costoP )  AS totTrasl FROM DetalleTraslado,TFPRODUC WHERE consec = dt_producto AND DetalleTraslado.dt_clave = '" & txtcampos(0).Text & "'", cCadConex, adOpenKeyset, adLockOptimistic, adCmdText
    cn.Execute "UPDATE traslados SET t_costo = " & IIf(IsNull(rsttemp!tottrasl), 0, rsttemp!tottrasl) & " WHERE t_clave = '" & Trim(txtcampos(0).Text) & "'"
    txtcampos(4).Text = IIf(IsNull(rsttemp!tottrasl), 0, rsttemp!tottrasl)
    Set rsttemp = Nothing
End If
End Sub

Private Sub dbgrdDetPed_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
Dim Dia As String
Dim rs As ADODB.Recordset
On Error GoTo Error:
Dim fila

If AdoTraslada.Recordset!t_enviado And lCon And UCase(dbgrdDetPed.Columns(ColIndex).DataField) <> "DT_COSTO" And UCase(dbgrdDetPed.Columns(ColIndex).DataField) <> "DT_VENTA" Then
    cresp = MsgBox("DESEAS PONER EN CERO EL INVENTARIO", vbYesNoCancel + vbQuestion + vbDefaultButton3)
    fila = clave  'Porque se pierde la fila al darle refresh
    If cresp = vbYes Then
       'Poner en cero la existencia en cajas
       If UCase(dbgrdDetPed.Columns(ColIndex).DataField) = "DT_CANTIDAD" Then
          If nFolAju = 0 Then 'Verifico si ya se metio el ajuste global
             cn.Execute "INSERT INTO Ajustes (a_responsable,a_fecha,a_motivo,a_observaciones) VALUES ('" & Trim(Mid(cCveDesUsu, 1, 3)) & "',' " & date + Time & "','SE AJUSTO EXISTENCIA DESDE TRASLADO " & txtcampos(0).Text & "','EN MAQUINA APARECIA EXISTENCIA Y FISICAMENTE YA NO HABIA' )"
             Set rs = New ADODB.Recordset
             rs.Open "SELECT MAX(a_clave) AS MaxClave FROM AJUSTES", cn, adOpenKeyset, adLockOptimistic, adCmdText
             nFolAju = rs!MaxClave
          End If
          'En caso de que el ajuste sea de un traslado de fecha anterior actualizo tabla de corte de inventarios
          If Day(AdoTraslada.Recordset!T_FECHA) <= Day(date) Then
             For N = Day(AdoTraslada.Recordset!T_FECHA) To Day(date)
                 Dia = "dia" & CStr(N)
                 cn.Execute "UPDATE invcorte SET " & Dia & " = " & Dia & " + " & AdoDetPed.Recordset!dt_cantidad - Val(dbgrdDetPed.Columns(ColIndex).Text) & " WHERE producto = '" & AdoDetPed.Recordset!Dt_producto & "' AND " & Dia & " > 0 "
             Next
          End If
          
          cn.Execute "INSERT INTO [DetalleAjustes] (da_clave,da_producto,da_cantidad,da_cantidadAnt) VALUES (" & nFolAju & ",'" & AdoDetPed.Recordset!Dt_producto & " ',-" & (AdoDetPed.Recordset!totcaj + (AdoDetPed.Recordset!dt_cantidad - Val(dbgrdDetPed.Columns(ColIndex).Text))) & "," & AdoDetPed.Recordset!dt_cantidad & ")"
          cn.Execute "UPDATE INVENTARIO SET INCANT = 0 WHERE inProd = '" & AdoDetPed.Recordset!Dt_producto & "'"
          cn.Execute "UPDATE DETALLETRASLADO SET DT_CANTIDAD = " & dbgrdDetPed.Columns(ColIndex).Text & " WHERE dt_producto = '" & AdoDetPed.Recordset!Dt_producto & "' AND DT_CLAVE = '" & txtcampos(0).Text & "'"
       'Poner en cero la existencia en piezas
       ElseIf UCase(dbgrdDetPed.Columns(ColIndex).DataField) = "DT_CANTIDADP" Then
            If nFolAju = 0 Then 'Verifico si ya se metio el ajuste global
               cn.Execute "INSERT INTO Ajustes (a_responsable,a_fecha,a_motivo,a_observaciones) VALUES ('" & Trim(Mid(cCveDesUsu, 1, 3)) & "',' " & date + Time & "','SE AJUSTO EXISTENCIA DESDE TRASLADO " & txtcampos(0).Text & "','EN MAQUINA APARECIA EXISTENCIA Y FISICAMENTE YA NO HABIA' )"
               Set rs = New ADODB.Recordset
               rs.Open "SELECT MAX(a_clave) AS MaxClave FROM AJUSTES", cn, adOpenKeyset, adLockOptimistic, adCmdText
               nFolAju = rs!MaxClave
            End If
            cn.Execute "INSERT INTO [DetalleAjustes] (da_clave,da_producto,da_cantidadP,da_cantidadAnt) VALUES (" & nFolAju & ",'" & AdoDetPed.Recordset!Dt_producto & "'," & (AdoDetPed.Recordset!InCantPza + (AdoDetPed.Recordset!dt_cantidadp - Val(dbgrdDetPed.Columns(ColIndex).Text))) & "," & AdoDetPed.Recordset!dt_cantidadp & ")"
            cn.Execute "UPDATE INVENTARIO SET InCantPza = 0 WHERE inProd = '" & AdoDetPed.Recordset!Dt_producto & "'"
            cn.Execute "UPDATE DETALLETRASLADO SET dt_cantidadP = " & dbgrdDetPed.Columns(ColIndex).Text & " WHERE dt_producto = '" & AdoDetPed.Recordset!Dt_producto & "' AND DT_CLAVE = '" & txtcampos(0).Text & "'"
       End If
       Cancel = True
    ElseIf cresp = vbNo Then  'Regresar al inventario
        'Regresar al inventario las CAJAS enviadas
        If UCase(dbgrdDetPed.Columns(ColIndex).DataField) = "DT_CANTIDAD" Then
            nCantidad = IIf(Val(OldValue) > Val(dbgrdDetPed.Columns(ColIndex).Text), Val(OldValue) - Val(dbgrdDetPed.Columns(ColIndex).Text), Val(dbgrdDetPed.Columns(ColIndex).Text) - Val(OldValue))
            If Val(OldValue) > Val(dbgrdDetPed.Columns(ColIndex).Text) Then
               cn.Execute "UPDATE INVENTARIO SET INCANT = Incant + " & nCantidad & " WHERE inProd = '" & AdoDetPed.Recordset!Dt_producto & "'"
               If Day(AdoTraslada.Recordset!T_FECHA) <= Day(date) Then
                  For N = Day(AdoTraslada.Recordset!T_FECHA) To Day(date)
                      Dia = "dia" & CStr(N)
                      cn.Execute "UPDATE invcorte SET " & Dia & " = " & Dia & " + " & nCantidad & " WHERE producto = '" & AdoDetPed.Recordset!Dt_producto & "'"
                  Next
               End If
            Else
                If nCantidad > AdoDetPed.Recordset!totcaj Then
                   Cancel = True
                   MsgBox "LA EXISTENCIA EN CAJAS: " & CStr(AdoDetPed.Recordset!totcaj) & " ES MENOR A LA CANTIDAD QUE SE DESEA AJUSTAR: " & CStr(nCantidad), vbExclamation
                   Exit Sub
                End If
                cn.Execute "UPDATE INVENTARIO SET INCANT = Incant - " & nCantidad & " WHERE inProd = '" & AdoDetPed.Recordset!Dt_producto & "'"
            End If
            cn.Execute "UPDATE DETALLETRASLADO SET DT_CANTIDAD = " & dbgrdDetPed.Columns(ColIndex).Text & " WHERE dt_producto = '" & AdoDetPed.Recordset!Dt_producto & "' AND DT_CLAVE = '" & txtcampos(0).Text & "'"
        'Regresar al inventario las PIEZAS enviadas
        ElseIf UCase(dbgrdDetPed.Columns(ColIndex).DataField) = "DT_CANTIDADP" Then
            If Val(dbgrdDetPed.Columns(ColIndex).Text) >= AdoDetPed.Recordset!PAQUETES Then
               MsgBox "LA CANTIDAD SOLICITADA EN PIEZAS ES MAYOR O IGUAL A LOS PAQUETES POR CAJA", vbExclamation
               Cancel = True
               Exit Sub
            End If
            nCantidad = IIf(Val(OldValue) > Val(dbgrdDetPed.Columns(ColIndex).Text), Val(OldValue) - Val(dbgrdDetPed.Columns(ColIndex).Text), Val(dbgrdDetPed.Columns(ColIndex).Text) - Val(OldValue))
            'Incrementar el inventario
            If Val(OldValue) > Val(dbgrdDetPed.Columns(ColIndex).Text) Then
                cn.Execute "UPDATE INVENTARIO SET InCantPza = InCantPza + " & nCantidad & " WHERE inProd = '" & AdoDetPed.Recordset!Dt_producto & "'"
            Else 'Disminuir el inventario
               'Si es mayor la cantidad enviada en piezas y es menor la existencia en piezas disminuyo una caja del inventario y la convierto en piezas
               If AdoDetPed.Recordset!InCantPza < Val(nCantidad) Then
                  If AdoDetPed.Recordset!totcaj = 0 Then
                     Cancel = True
                     MsgBox "NO ES POSIBLE SURTIR MAS PIEZAS PORQUE LA EXISTENCIA ES " & CStr(AdoDetPed.Recordset!InCantPza) & Chr(13) & "Y LA EXISTENCIA EN CAJAS ES: 0", vbExclamation
                     Exit Sub
                  End If
                  'Disminuyo una caja del ineventario y la convierto en piezas
                  cn.Execute "UPDATE inventario SET incantpza = incantpza + paquetes, incant = incant - 1 FROM tfproduc WHERE consec = inprod AND Inprod = '" & AdoDetPed.Recordset!Dt_producto & "'"
               End If
               cn.Execute "UPDATE INVENTARIO SET InCantPza = InCantPza - " & nCantidad & " WHERE inProd = '" & AdoDetPed.Recordset!Dt_producto & "'"
            End If
            cn.Execute "UPDATE DETALLETRASLADO SET Dt_cantidadP = " & dbgrdDetPed.Columns(ColIndex).Text & " WHERE dt_producto = '" & AdoDetPed.Recordset!Dt_producto & "' AND DT_CLAVE = '" & txtcampos(0).Text & "'"
        End If
        Cancel = True
    Else
        Cancel = True
        Exit Sub
    End If
    AdoDetPed.Refresh
    AdoDetPed.Recordset.Bookmark = fila
    SendKeys "{DOWN}"
ElseIf UCase(dbgrdDetPed.Columns(ColIndex).DataField) = "DT_CANTIDAD" Then
    cn.Execute "UPDATE DETALLETRASLADO SET Dt_cantidad = " & dbgrdDetPed.Columns(ColIndex).Text & " WHERE dt_producto = '" & AdoDetPed.Recordset!Dt_producto & "' AND DT_CLAVE = '" & txtcampos(0).Text & "'"
    marca = AdoDetPed.Recordset.Bookmark
    'AdoDetPed.Refresh
    'AdoDetPed.Recordset.Bookmark = marca
    Cancel = True
ElseIf UCase(dbgrdDetPed.Columns(ColIndex).DataField) = "DT_CANTIDADP" Then
    If Val(AdoDetPed.Recordset!PAQUETES) <= Val(dbgrdDetPed.Columns(ColIndex).Text) And AdoDetPed.Recordset!PAQUETES <> 1 Then
       MsgBox "EL NUMERO DE PIEZAS DEBE SER MENOR A LAS QUE TRAE LA CAJA", vbExclamation
       Cancel = True
    End If
    cn.Execute "UPDATE DETALLETRASLADO SET Dt_cantidadp = " & dbgrdDetPed.Columns(ColIndex).Text & " WHERE dt_producto = '" & AdoDetPed.Recordset!Dt_producto & "' AND DT_CLAVE = '" & txtcampos(0).Text & "'"
    marca = AdoDetPed.Recordset.Bookmark
    'AdoDetPed.Refresh
    'AdoDetPed.Recordset.Bookmark = marca
    Cancel = True
ElseIf UCase(dbgrdDetPed.Columns(ColIndex).DataField) = "DT_COSTO" Then
    cn.Execute "UPDATE DETALLETRASLADO SET Dt_costo = " & dbgrdDetPed.Columns(ColIndex).Text & " WHERE dt_producto = '" & AdoDetPed.Recordset!Dt_producto & "' AND DT_CLAVE = '" & txtcampos(0).Text & "'"
    'AdoDetPed.Refresh
    Cancel = True
    'SendKeys "{DOWN}"
ElseIf UCase(dbgrdDetPed.Columns(ColIndex).DataField) = "DT_VENTA" Then
    cn.Execute "UPDATE DETALLETRASLADO SET dt_venta = " & dbgrdDetPed.Columns(ColIndex).Text & " WHERE dt_producto = '" & AdoDetPed.Recordset!Dt_producto & "' AND DT_CLAVE = '" & txtcampos(0).Text & "'"
    'AdoDetPed.Refresh
    Cancel = True
    'SendKeys "{DOWN}"
End If
If Not AdoTraslada.Recordset!t_enviado Then
   'Precio a franquicias (PRECIO4  de PREPROD)
   If lfranquicia Then
       cn.Execute "UPDATE DETALLETRASLADO SET DETALLETRASLADO.dt_costo = TFPRODUC.PRECOSTO, DETALLETRASLADO.dt_costoP = TFPRODUC.PRECOSTO /TFPRODUC.paquetes, DETALLETRASLADO.dt_venta = PREPROD.Precio4, DETALLETRASLADO.dt_ventap = PREPROD.Precio4 / TFPRODUC.paquetes FROM PREPROD,TFPRODUC WHERE PREPROD.preclave = DETALLETRASLADO.dt_producto AND PREPROD.preclave = TFPRODUC.Consec AND TFPRODUC.Consec = DETALLETRASLADO.dt_producto AND DetalleTraslado.dt_clave = '" & txtcampos(0).Text & "' AND tfproduc.ACTIVO = 1"
   Else 'Precio a tiendas (PRECOSTO de TFPRODUC)
       cn.Execute "UPDATE DETALLETRASLADO SET dt_costo = TFPRODUC.Precosto, dt_costoP = TFPRODUC.Precosto / TFPRODUC.paquetes , dt_venta = PREPROD.precio2, dt_ventaP = PREPROD.precio1 FROM TFPRODUC,PREPROD WHERE PREPROD.preclave = DETALLETRASLADO.dt_producto AND PREPROD.preclave = TFPRODUC.Consec AND TFPRODUC.Consec = DETALLETRASLADO.dt_producto AND DetalleTraslado.dt_clave = '" & txtcampos(0).Text & "' and tfproduc.ACTIVO = 1"
   End If
   'marca = AdoDetPed.Recordset.Bookmark
   AdoDetPed.Refresh
   AdoDetPed.Recordset.Bookmark = marca
End If
Exit Sub
Error:
MsgBox Err.Description
End Sub


Private Sub dbgrdDetPed_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Select Case KeyCode
Case 116  'F5
     lModTra = 0
     fraCon.Visible = True
     Me.txtContra.SetFocus
End Select

End Sub

Private Sub dbgrdDetPed_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     KeyAscii = 0
     'SendKeys vbTab
     keybd_event &H9, 0, 0, 0
  End If
End Sub

Private Sub dbgrdDetPed_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
 clave = AdoDetPed.Recordset.Bookmark
End Sub

Private Sub DbgrdDetTraAbi_AfterColUpdate(ByVal ColIndex As Integer)
  AdoDetPed.Recordset!DT_CLAVE = txtcampos(0).Text
End Sub

Private Sub DbgrdDetTraAbi_AfterUpdate()
Dim rsttemp As ADODB.Recordset
Dim marca
On Error Resume Next
If AdoTraslada.Recordset!t_tipo Then
marca = AdoDetPed.Recordset.Bookmark
Set rsttemp = New ADODB.Recordset
PonPie   'Pone Informacion de numero de product, cajas y piezas
'Lo meti aqui y no en el AFTERINSERT porque pueden cambiar la clave del producto y ya no es el mismo articulo
  'Mientras no se haya enviado no se puede modificarr precios, para que los facturistas no puedan hacerlo, a traves de
  'contraseÑa si se puede.
  If Not AdoTraslada.Recordset!t_enviado Then
     'Precio a franquicias (PRECIO4  de PREPROD)
     'If Val(txtCampos(3).Text) = 5 Or Val(txtCampos(3).Text) = 12 Or Val(txtCampos(3).Text) = 13 Or Val(txtCampos(3).Text) = 15 Or Val(txtCampos(3).Text) = 14 Or Val(txtCampos(3).Text) = 27 Then
     If lfranquicia Then
         cn.Execute "UPDATE DETALLETRASLADO SET DETALLETRASLADO.dt_costo = TFPRODUC.PRECOSTO, DETALLETRASLADO.dt_costoP = TFPRODUC.PRECOSTO /TFPRODUC.paquetes, DETALLETRASLADO.dt_venta = PREPROD.Precio4, DETALLETRASLADO.dt_ventap = PREPROD.Precio4 / TFPRODUC.paquetes FROM PREPROD,TFPRODUC WHERE PREPROD.preclave = DETALLETRASLADO.dt_producto AND PREPROD.preclave = TFPRODUC.Consec AND TFPRODUC.Consec = DETALLETRASLADO.dt_producto AND DetalleTraslado.dt_clave = '" & txtcampos(0).Text & "' AND tfproduc.ACTIVO = 1"
     Else 'Precio a tiendas (PRECOSTO de TFPRODUC)
         cn.Execute "UPDATE DETALLETRASLADO SET dt_costo = TFPRODUC.Precosto, dt_costoP = TFPRODUC.Precosto / TFPRODUC.paquetes , dt_venta = PREPROD.precio2, dt_ventaP = PREPROD.precio1 FROM TFPRODUC,PREPROD WHERE PREPROD.preclave = DETALLETRASLADO.dt_producto AND PREPROD.preclave = TFPRODUC.Consec AND TFPRODUC.Consec = DETALLETRASLADO.dt_producto AND DetalleTraslado.dt_clave = '" & txtcampos(0).Text & "' and tfproduc.ACTIVO = 1"
     End If
     'ADEMAS SE DEBE GUARDAR EL IVA Y EL IEPS DEL PRODUCTO
     
  Else
     'If Val(txtCampos(3).Text) = 5 Or Val(txtCampos(3).Text) = 12 Or Val(txtCampos(3).Text) = 13 Or Val(txtCampos(3).Text) = 15 Or Val(txtCampos(3).Text) = 14 Then
     If lfranquicia Then
         rsttemp.Open "SELECT SUM(dt_cantidad * dt_venta) + SUM(dt_cantidadp * dt_ventap )  AS totTrasl FROM DetalleTraslado,TFPRODUC WHERE consec = dt_producto AND DetalleTraslado.dt_clave = '" & txtcampos(0).Text & "'", cCadConex, adOpenKeyset, adLockOptimistic, adCmdText
         cn.Execute "UPDATE DETALLETRASLADO SET dt_importe = (dt_cantidad * dt_venta) + (dt_cantidadp * dt_ventap) WHERE dt_clave = '" & Trim(txtcampos(0).Text) & "'"
      Else 'Si es tienda
         rsttemp.Open "SELECT SUM(dt_cantidad * dt_costo) + SUM(dt_cantidadp * dt_costop ) AS totTrasl FROM DetalleTraslado,TFPRODUC WHERE consec = dt_producto AND DetalleTraslado.dt_clave = '" & txtcampos(0).Text & "'", cCadConex, adOpenKeyset, adLockOptimistic, adCmdText
         cn.Execute "UPDATE DETALLETRASLADO SET dt_importe = (dt_cantidad * dt_costo) + (dt_cantidadp * dt_costop) WHERE dt_clave = '" & Trim(txtcampos(0).Text) & "'"
      End If
      cn.Execute "UPDATE traslados SET t_enviado = 1, t_costo = " & IIf(IsNull(rsttemp!tottrasl), 0, rsttemp!tottrasl) & " WHERE t_clave = '" & Trim(txtcampos(0).Text) & "'"
      txtcampos(4).Text = IIf(IsNull(rsttemp!tottrasl), 0, rsttemp!tottrasl)
  End If
'ESTA PARTE SE DESACTIVA PARA TODAS LAS TIENDAS
'EXCEPTO PARA CARBONERA
'Se cargan los nombres de los productos y se va al final del grid
If Trim(txtcampos(1).Text) = "3" Then
'   If Not AdoTraslada.Recordset!t_enviado Then
'         AdoDetPed.Refresh
'   End If
End If
LblProdAgr.Caption = ""
End If
End Sub

Private Sub DbgrdDetTraAbi_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
If ColIndex = 0 Then  'Obliga a que seleccionen un producto de la lista desplegable
   Cancel = True
   DbgrdDetTraAbi_ButtonClick (ColIndex)
End If
End Sub

Private Sub DbgrdDetTraAbi_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
Dim rs As ADODB.Recordset
Dim rsttemp As ADODB.Recordset
'On Error GoTo ERROR
Dim fila
'MsgBox AdoTraslada.Recordset!t_clave
'MsgBox AdoTraslada.Recordset!t_enviado
If AdoTraslada.Recordset!t_enviado And lCon And UCase(DbgrdDetTraAbi.Columns(ColIndex).DataField) <> "DT_COSTO" And UCase(DbgrdDetTraAbi.Columns(ColIndex).DataField) <> "DT_VENTA" Then
    nIndCant = 4  'Columna correspondiente a cantidad enviada en cajas
    cresp = MsgBox("DESEAS PONER EN CERO EL INVENTARIO", vbYesNoCancel + vbDefaultButton2 + vbQuestion)
    fila = clave  'Porque se pierde la fila al darle refresh
    'Poner en cero el inventario
    If cresp = vbYes Then
       'Si es columna de existencia en cajas
       If UCase(DbgrdDetTraAbi.Columns(ColIndex).DataField) = "DT_CANTIDAD" Then
          If nFolAju = 0 Then 'Verifico si ya se metio el ajuste global
             cn.Execute "INSERT INTO Ajustes (a_responsable,a_fecha,a_motivo,a_observaciones) VALUES ('" & Trim(Mid(cCveDesUsu, 1, 3)) & "',' " & date + Time & "','SE AJUSTO EXISTENCIA DESDE TRASLADO " & txtcampos(0).Text & "','EN MAQUINA APARECIA EXISTENCIA Y FISICAMENTE YA NO HABIA' )"
             Set rs = New ADODB.Recordset
             rs.Open "SELECT MAX(a_clave) AS MaxClave FROM AJUSTES", cn, adOpenKeyset, adLockOptimistic, adCmdText
             nFolAju = rs!MaxClave
          End If
          'En caso de que el ajuste sea de un traslado de fecha anterior actualizo tabla de corte de inventarios
          If Day(AdoTraslada.Recordset!T_FECHA) <= Day(date) Then
             For N = Day(AdoTraslada.Recordset!T_FECHA) To Day(date)
                 Dia = "dia" & CStr(N)
                 cn.Execute "UPDATE invcorte SET " & Dia & " = " & Dia & " + " & AdoDetPed.Recordset!dt_cantidad - Val(DbgrdDetTraAbi.Columns(ColIndex).Text) & " WHERE producto = '" & AdoDetPed.Recordset!Dt_producto & "' AND " & Dia & " > 0 "
             Next
          End If

          cn.Execute "INSERT INTO [DetalleAjustes] (da_clave,da_producto,da_cantidad,da_cantidadAnt) VALUES (" & nFolAju & ",'" & AdoDetPed.Recordset!Dt_producto & " ',-" & (AdoDetPed.Recordset!EXICAJ + (AdoDetPed.Recordset!dt_cantidad - Val(DbgrdDetTraAbi.Columns(ColIndex).Text))) & "," & AdoDetPed.Recordset!dt_cantidad & ")"
          cn.Execute "UPDATE INVENTARIO SET INCANT = 0 WHERE inProd = '" & AdoDetPed.Recordset!Dt_producto & "'"
          cn.Execute "UPDATE DETALLETRASLADO SET dt_cantidad = " & DbgrdDetTraAbi.Columns(ColIndex).Text & " WHERE dt_producto = '" & AdoDetPed.Recordset!Dt_producto & "' AND DT_CLAVE = '" & txtcampos(0).Text & "'"
       'Columna correspondiente a cantidades por pieza
       ElseIf UCase(DbgrdDetTraAbi.Columns(ColIndex).DataField) = "DT_CANTIDADP" Then
          If nFolAju = 0 Then 'Verifico si ya se metio el ajuste global
             cn.Execute "INSERT INTO Ajustes (a_responsable,a_fecha,a_motivo,a_observaciones) VALUES ('" & Trim(Mid(cCveDesUsu, 1, 3)) & "',' " & date + Time & "','SE AJUSTO EXISTENCIA DESDE TRASLADO " & txtcampos(0).Text & "','EN MAQUINA APARECIA EXISTENCIA Y FISICAMENTE YA NO HABIA' )"
             Set rs = New ADODB.Recordset
             rs.Open "SELECT MAX(a_clave) AS MaxClave FROM AJUSTES", cn, adOpenKeyset, adLockOptimistic, adCmdText
             nFolAju = rs!MaxClave
          End If
          cn.Execute "INSERT INTO [DetalleAjustes] (da_clave,da_producto,da_cantidadP,da_cantidadAnt) VALUES (" & nFolAju & ",'" & AdoDetPed.Recordset!Dt_producto & " ',-" & (AdoDetPed.Recordset!InCantPza + (AdoDetPed.Recordset!dt_cantidadp - Val(DbgrdDetTraAbi.Columns(ColIndex).Text))) & "," & AdoDetPed.Recordset!dt_cantidadp & ")"
          cn.Execute "UPDATE INVENTARIO SET InCantPza = 0 WHERE inProd = '" & AdoDetPed.Recordset!Dt_producto & "'"
          cn.Execute "UPDATE DETALLETRASLADO SET dt_cantidadP = " & DbgrdDetTraAbi.Columns(ColIndex).Text & " WHERE dt_producto = '" & AdoDetPed.Recordset!Dt_producto & "' AND DT_CLAVE = '" & txtcampos(0).Text & "'"
       End If
       Cancel = True
    'Regresar la cantidad al inventario
    ElseIf cresp = vbNo Then
        'Regresar al inventario las CAJAS enviadas
        If UCase(DbgrdDetTraAbi.Columns(ColIndex).DataField) = "DT_CANTIDAD" Then
            nCantidad = IIf(Val(OldValue) > Val(DbgrdDetTraAbi.Columns(ColIndex).Text), Val(OldValue) - Val(DbgrdDetTraAbi.Columns(ColIndex).Text), Val(DbgrdDetTraAbi.Columns(ColIndex).Text) - Val(OldValue))
            If Val(OldValue) > Val(DbgrdDetTraAbi.Columns(ColIndex).Text) Then
               cOper = IIf(Me.lblEntrada.Visible = True, "-", "+")
               cn.Execute "UPDATE INVENTARIO SET INCANT = Incant " & cOper & nCantidad & " WHERE inProd = '" & AdoDetPed.Recordset!Dt_producto & "'"
               'Si el ajustes es de traslado de fecha anterior actualizar tabla de cortes
               
               If Day(AdoTraslada.Recordset!T_FECHA) <= Day(date) Then
                  For N = Day(AdoTraslada.Recordset!T_FECHA) To Day(date)
                      Dia = "dia" & CStr(N)
                      cn.Execute "UPDATE invcorte SET " & Dia & " = " & Dia & " + " & nCantidad & " WHERE producto = '" & AdoDetPed.Recordset!Dt_producto & "'"
                  Next
               End If
            Else
                If nCantidad > AdoDetPed.Recordset!EXICAJ Then
                   Cancel = True
                   MsgBox "LA EXISTENCIA EN CAJAS: " & CStr(AdoDetPed.Recordset!EXICAJ) & " ES MENOR A LA CANTIDAD QUE SE DESEA AJUSTAR: " & CStr(nCantidad), vbExclamation
                   Exit Sub
                End If
                cn.Execute "UPDATE INVENTARIO SET INCANT = Incant - " & nCantidad & " WHERE inProd = '" & AdoDetPed.Recordset!Dt_producto & "'"
            End If
            cn.Execute "UPDATE DETALLETRASLADO SET DT_CANTIDAD = " & DbgrdDetTraAbi.Columns(ColIndex).Text & " WHERE dt_producto = '" & AdoDetPed.Recordset!Dt_producto & "' AND DT_CLAVE = '" & txtcampos(0).Text & "'"
        'Regresar al inventario las PIEZAS enviadas
        ElseIf UCase(DbgrdDetTraAbi.Columns(ColIndex).DataField) = "DT_CANTIDADP" Then
            If Val(DbgrdDetTraAbi.Columns(ColIndex).Text) >= AdoDetPed.Recordset!PAQUETES Then
               MsgBox "LA CANTIDAD SOLICITADA EN PIEZAS ES MAYOR O IGUAL A LOS PAQUETES POR CAJA", vbExclamation
               Cancel = True
               Exit Sub
            End If
            nCantidad = IIf(Val(OldValue) > Val(DbgrdDetTraAbi.Columns(ColIndex).Text), Val(OldValue) - Val(DbgrdDetTraAbi.Columns(ColIndex).Text), Val(DbgrdDetTraAbi.Columns(ColIndex).Text) - Val(OldValue))
            'Incrementar el inventario
            If Val(OldValue) > Val(DbgrdDetTraAbi.Columns(ColIndex).Text) Then
                          cOper = IIf(Me.lblEntrada.Visible = True, "-", "+")
               cn.Execute "UPDATE INVENTARIO SET InCantPza = InCantPza " & cOper & nCantidad & " WHERE inProd = '" & AdoDetPed.Recordset!Dt_producto & "'"
            Else 'Disminuir el inventario
               'Si es mayor la cantidad enviada en piezas y es menor la existencia en piezas disminuyo una caja del inventario y la convierto en piezas
               If AdoDetPed.Recordset!InCantPza < Val(nCantidad) Then
                  If AdoDetPed.Recordset!EXICAJ = 0 Then
                     Cancel = True
                     MsgBox "NO ES POSIBLE SURTIR MAS PIEZAS PORQUE LA EXISTENCIA ES " & CStr(AdoDetPed.Recordset!InCantPza) & Chr(13) & "Y LA EXISTENCIA EN CAJAS ES: 0", vbExclamation
                     Exit Sub
                  End If
                  'Disminuyo una caja del ineventario y la convierto en piezas
                  cn.Execute "UPDATE inventario SET incantpza = incantpza + paquetes, incant = incant - 1 FROM tfproduc WHERE consec = inprod AND Inprod = '" & AdoDetPed.Recordset!Dt_producto & "'"
               End If
               cn.Execute "UPDATE INVENTARIO SET InCantPza = InCantPza - " & nCantidad & " WHERE inProd = '" & AdoDetPed.Recordset!Dt_producto & "'"
            End If
            cn.Execute "UPDATE DETALLETRASLADO SET Dt_cantidadP = " & DbgrdDetTraAbi.Columns(ColIndex).Text & " WHERE dt_producto = '" & AdoDetPed.Recordset!Dt_producto & "' AND DT_CLAVE = '" & txtcampos(0).Text & "'"
        End If
        Cancel = True
    Else
        Cancel = True
        Exit Sub
    End If
    AdoDetPed.Refresh
    AdoDetPed.Recordset.Bookmark = fila
    SendKeys "{DOWN}"
ElseIf UCase(DbgrdDetTraAbi.Columns(ColIndex).DataField) = "DT_CANTIDAD" Then
    If Not IsNull(AdoDetPed.Recordset!descripc) Then
       cn.Execute "UPDATE DETALLETRASLADO SET Dt_CANTIDAD = " & DbgrdDetTraAbi.Columns(ColIndex).Text & " WHERE dt_producto = '" & AdoDetPed.Recordset!Dt_producto & "' AND DT_CLAVE = '" & txtcampos(0).Text & "'"
       marca1 = AdoDetPed.Recordset.Bookmark
       AdoDetPed.Refresh
       AdoDetPed.Recordset.Bookmark = marca1
       Cancel = True
    End If
ElseIf UCase(DbgrdDetTraAbi.Columns(ColIndex).DataField) = "DT_CANTIDADP" Then
    If AdoDetPed.Recordset!PAQUETES <= Val(DbgrdDetTraAbi.Columns(ColIndex).Text) And AdoDetPed.Recordset!PAQUETES <> 1 Then
       MsgBox "EL NUMERO DE PIEZAS DEBE SER MENOR A LAS QUE TRAE LA CAJA", vbExclamation
       Cancel = True
    Else
        If Not IsNull(AdoDetPed.Recordset!descripc) Then
           cn.Execute "UPDATE DETALLETRASLADO SET Dt_CANTIDADP = " & DbgrdDetTraAbi.Columns(ColIndex).Text & " WHERE dt_producto = '" & Trim(AdoDetPed.Recordset!Dt_producto) & "' AND DT_CLAVE = '" & Trim(txtcampos(0).Text) & "'"
           marca = AdoDetPed.Recordset.Bookmark
           AdoDetPed.Refresh
           AdoDetPed.Recordset.Bookmark = marca
           Cancel = True
        End If
    End If
ElseIf UCase(DbgrdDetTraAbi.Columns(ColIndex).DataField) = "DT_COSTO" Then
    'MsgBox "UPDATE DETALLETRASLADO SET Dt_costo = " & DbgrdDetTraAbi.Columns(ColIndex).Text & " WHERE dt_producto = '" & AdoDetPed.Recordset!Dt_producto & "' AND DT_CLAVE = '" & txtCampos(0).Text & "'"
    cn.Execute "UPDATE DETALLETRASLADO SET Dt_costo = " & DbgrdDetTraAbi.Columns(ColIndex).Text & " WHERE dt_producto = '" & AdoDetPed.Recordset!Dt_producto & "' AND DT_CLAVE = '" & txtcampos(0).Text & "'"
    'PIEZAS
    Set rsttemp = New ADODB.Recordset
    rsttemp.Open "SELECT paquetes FROM TFPRODUC WHERE consec = '" & Trim(AdoDetPed.Recordset!Dt_producto) & "'", cCadConex, adOpenKeyset, adLockOptimistic, adCmdText
    PAQUETES = rsttemp!PAQUETES
    Set rsttemp = Nothing
    costop = DbgrdDetTraAbi.Columns(ColIndex).Text / PAQUETES
    cn.Execute "UPDATE DETALLETRASLADO SET Dt_costop = " & costop & " WHERE dt_producto = '" & AdoDetPed.Recordset!Dt_producto & "' AND DT_CLAVE = '" & txtcampos(0).Text & "'"
    AdoDetPed.Refresh
    'SendKeys "{DOWN}"
    Cancel = True
ElseIf UCase(DbgrdDetTraAbi.Columns(ColIndex).DataField) = "DT_VENTA" Then
    'MsgBox "UPDATE DETALLETRASLADO SET Dt_costo = " & DbgrdDetTraAbi.Columns(ColIndex).Text & " WHERE dt_producto = '" & AdoDetPed.Recordset!Dt_producto & "' AND DT_CLAVE = '" & txtCampos(0).Text & "'"
    cn.Execute "UPDATE DETALLETRASLADO SET Dt_venta = " & DbgrdDetTraAbi.Columns(ColIndex).Text & " WHERE dt_producto = '" & AdoDetPed.Recordset!Dt_producto & "' AND DT_CLAVE = '" & txtcampos(0).Text & "'"
    Set rsttemp = New ADODB.Recordset
    rsttemp.Open "SELECT paquetes FROM TFPRODUC WHERE consec = '" & Trim(AdoDetPed.Recordset!Dt_producto) & "'", cCadConex, adOpenKeyset, adLockOptimistic, adCmdText
    PAQUETES = rsttemp!PAQUETES
    Set rsttemp = Nothing
    preciop = DbgrdDetTraAbi.Columns(ColIndex).Text / PAQUETES
    cn.Execute "UPDATE DETALLETRASLADO SET Dt_ventap = " & preciop & " WHERE dt_producto = '" & AdoDetPed.Recordset!Dt_producto & "' AND DT_CLAVE = '" & txtcampos(0).Text & "'"
    AdoDetPed.Refresh
    'SendKeys "{DOWN}"
    Cancel = True
End If
'MsgBox AdoTraslada.Recordset!t_clave
If AdoTraslada.Recordset!t_enviado Then
    Set rsttemp = New ADODB.Recordset
    rsttemp.Open "SELECT SUM(dt_cantidad * dt_costo) + SUM(dt_cantidadp * dt_costoP )  AS totTrasl FROM DetalleTraslado,TFPRODUC WHERE consec = dt_producto AND DetalleTraslado.dt_clave = '" & txtcampos(0).Text & "'", cCadConex, adOpenKeyset, adLockOptimistic, adCmdText
    cn.Execute "UPDATE traslados SET t_costo = " & IIf(IsNull(rsttemp!tottrasl), 0, rsttemp!tottrasl) & " WHERE t_clave = '" & Trim(txtcampos(0).Text) & "'"
    AdoTraslada.Refresh
    Set rsttemp = Nothing
End If
Exit Sub
Error:
  MsgBox Err.Description
End Sub

Private Sub DbgrdDetTraAbi_BeforeDelete(Cancel As Integer)
If Not lCon And chkEnviado.Value = 1 Then
   fraCon.Visible = True
   txtContra.SetFocus
   Cancel = True
   Exit Sub
End If
If MsgBox("REALMENTE DESEAS BORRAR EL PRODUCTO" & Chr(13) & AdoDetPed.Recordset!descripc & "  " & AdoDetPed.Recordset!Present, vbQuestion + vbYesNo) = vbYes Then
   If chkEnviado.Value = 1 Then
      cn.Execute "UPDATE Inventario SET InCant = InCant + " & AdoDetPed.Recordset!dt_cantidad & " WHERE Inprod = '" & AdoDetPed.Recordset!Dt_producto & "'"
      cn.Execute "DELETE FROM detalletraslado WHERE dt_clave = '" & Trim(txtcampos(0).Text) & "' AND dt_producto = '" & AdoDetPed.Recordset!Dt_producto & "'"
   Else
      cn.Execute "DELETE FROM detalletraslado WHERE dt_clave = '" & Trim(txtcampos(0).Text) & "' AND dt_producto = '" & AdoDetPed.Recordset!Dt_producto & "'"
   End If
   AdoDetPed.Refresh
End If
Cancel = True
DbgrdDetTraAbi.SetFocus
End Sub

Private Sub DbgrdDetTraAbi_ButtonClick(ByVal ColIndex As Integer)
'On Error GoTo Error:
Dim L As ListBox
    Select Case ColIndex
       Case ColIndex_lstProd
            Set L = Lstprod
    End Select
    If ColIndex = -1 Then Exit Sub
      With L
          'Abajo (3):
          .Left = DbgrdDetTraAbi.Left + DbgrdDetTraAbi.Columns(ColIndex).Left
          If Not (AdoDetPed.Recordset.BOF And AdoDetPed.Recordset.EOF) Then .Top = DbgrdDetTraAbi.Top + DbgrdDetTraAbi.RowTop(DbgrdDetTraAbi.Row) + DbgrdDetTraAbi.RowHeight
          '.Width = dbgrdDetPed.Columns(ColIndex).Width + 15
          '.ListIndex = 0
          .Visible = True
          .ZOrder 0
          .SetFocus
    End With
   Exit Sub
Error:
   MsgBox Err.Description
End Sub

Private Sub DbgrdDetTraAbi_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Select Case KeyCode
Case 116  'F5
     lModTra = 1
     fraCon.Visible = True
     Me.txtContra.SetFocus
Case 120  'F9
     AdoDetPed.Recordset.MoveFirst
Case 121  'F10
     SendKeys "{PGUP}"
Case 122  'F11
     SendKeys "{PGDN}"
Case 123  'F12
     AdoDetPed.Recordset.MoveLast
     SendKeys "{DOWN}"
End Select

End Sub

Private Sub DbgrdDetTraAbi_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     KeyAscii = 0
     keybd_event &H9, 0, 0, 0
  End If
End Sub

Private Sub DesgTraPed()
Dim cmensa As String
On Error Resume Next
If Not (AdoPedTra.Recordset.BOF And AdoPedTra.Recordset.EOF) Then
   cmensa = StbMensajes.SimpleText
   StbMensajes.SimpleText = Space(40) & "Espere un momento, buscando productos correspondientes al traslado....."
   StbMensajes.Refresh
   Pedido = dbgrdPedEnTra.Columns(0)
   'Obtengo datos informativos del detalle de traslado
   AdoDetPed.ConnectionString = cCadConex
   AdoDetPed.CommandType = adCmdText
   AdoDetPed.RecordSource = "SELECT detalleTraslado.dt_producto, tfproduc.descripc, LTrim(str(paquetes)) + ' X ' + LTRIM( str(contenid,10,3)) + ' ' + MEDIDA as MEDIDA, " _
         & " DetalleTraslado.dt_cantidad,DetalleTraslado.dt_cantidadP, DetalleTraslado.dt_pedido, Inventario.inprod, CAST(Inventario.inCant AS SMALLMONEY) as TOTCAJ, dt_costo, dt_venta, paquetes, inCantPza, Ubicacion FROM DetalleTraslado, tfproduc, Inventario WHERE DetalleTraslado.dt_CLAVE = '" & Trim(txtcampos(0).Text) & " '" _
         & "AND DetalleTraslado.dt_producto = tfproduc.consec AND detalleTraslado.dt_producto = Inventario.inProd ORDER BY Descripc,contenid"
   AdoDetPed.Refresh
   StbMensajes.SimpleText = cmensa
   StbMensajes.Refresh
 End If
Me.FraPedEnv.Visible = True
End Sub

Private Sub DbgrdDetTraAbi_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
 clave = AdoDetPed.Recordset.Bookmark
End Sub

Private Sub Form_Activate()
If Val(Mid(cSucursal, 1, 1)) = 3 And tipotienda = 2 Then
    fg = cuadraproductos(strcveprod)
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 119 Then
      frmCalc.Show  'F9
   ElseIf KeyCode = 118 Then 'f7
       MsgBox "Opcion Activada para no generar Folio  de Tienda en Automatico", vbInformation
       manual = True
       lblInfo.Caption = "Opcion Activada para no generar Folio  de Tienda en Automatico"
       lblInfo.Refresh
   End If
End Sub

Private Sub Form_Load()
Dim N As Integer
 lCon = False
 txtcampos(0).Visible = True
 lbletiquetas(0).Visible = True
 AdoTraslada.ConnectionString = cCadConex
 AdoTraslada.CommandType = adCmdText
 'AdoTraslada.RecordSource = "SELECT * FROM [Traslados]"
 'AdoTraslada.Refresh
 nFolAju = 0
 manual = False
End Sub

Private Function cuadraproductos(traslado As String)
'se borran si hay duplicados
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
clave = traslado
rs.Open "SELECT * from detalletraslado WHERE dt_clave = '" & Trim(clave) & "' order by dt_producto ", cn, adOpenDynamic, adLockOptimistic, adCmdText
prod = "moy"
While Not rs.EOF
    If Trim(prod) = Trim(rs!Dt_producto) Then
       rs!dt_borrado = 1
       rs.Update
    End If
    prod = Trim(rs!Dt_producto)
    rs.MoveNext
Wend
Set rs = Nothing
CAD = "delete detalletraslado where dt_borrado = 1 and dt_clave = '" & Trim(clave) & "'"
'MsgBox cad
cn.Execute CAD
End Function
Private Sub Form_Unload(Cancel As Integer)
  frmtraslados.Show
End Sub

Private Sub Text1_Change()

End Sub

Private Sub tpre_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Call cambiaprecios
   fraprecios.Enabled = False
   fraprecios.Visible = False
End If
End Sub

Private Sub txtCampos_GotFocus(Index As Integer)
Dim rsttemp As ADODB.Recordset
Dim Folio As String
'Campos de tipo fecha en los que se muestra en control calendario
   Select Case Index
   Case 0  'Clave del folio del pedido
        'Unload frmtraslados
        frmtraslados.Hide
        If Forma = 1 Then lblEntrada.Visible = True
        If nOp = 1 Then 'En caso de Altas
            Set rsttemp = New ADODB.Recordset
            rsttemp.ActiveConnection = cCadConex
            rsttemp.CursorType = adOpenKeyset
                 rsttemp.Source = "SELECT MAX (CAST(SUBSTRING(t_clave,4,10) AS INT)) As FolTra FROM [Traslados] WHERE SUBSTRING(t_clave,1,3) = 'T" & Trim(Mid(cSucursal, 3, 5)) & "'"
     rsttemp.Open
     If IsNull(rsttemp!FolTra) Then
        Folio = "T" + Mid(Trim(Mid(cSucursal, 3)), 1, 2) + "1"
     Else
        Folio = "T" + Mid(Trim(Mid(cSucursal, 3)), 1, 2) + Trim(Str(rsttemp!FolTra + 1))
     End If

            
            AdoTraslada.RecordSource = "SELECT * FROM traslados WHERE T_CLAVE = '" & Folio & "'"
            AdoTraslada.Refresh
            AdoTraslada.Recordset.AddNew
            txtcampos(0).Text = Folio
            chkTipoTrasl.Value = 1 'Pongo por default traslado abierto
            chkPapeleria.Value = 0 'Traslado normal
            chkfrutas.Value = 0
            chkpan.Value = 0
            chkMerma.Value = 0
            chkvolumen.Value = 0
            chkauto.Value = 0
            Check1.Value = 0
            'Si es devolucion a proveedores
            If cModo = "DEVO" Then
               lbletiquetas(3).Caption = "Clave del proveedor"
               cmbSucRec.Width = 5400
            End If
            'MsgBox "Presione Cualquier Tecla para continuar...", vbInformation, "ENVIOS"
            SendKeys "{TAB}": SendKeys "{TAB}"
            'keybd_event &H9, 0, 0, 0: keybd_event &H9, 0, &H2, 0:  keybd_event &H9, 0, 0, 0: keybd_event &H9, 0, &H2, 0
        End If
        lbletiquetas(8).Visible = (Forma = 0 And nOp = 0) 'Se muestra en salidas y en modificaciones
        txtFoliotie.Visible = (Forma = 0 And nOp = 0)
        txtFoliotie.Enabled = False
   Case 3
        cmbUsuEmi.Enabled = False
    End Select
End Sub

Private Sub txtCampos_LostFocus(Index As Integer)
Dim Tipos
Tipos = Array("1 COMBI", "2 CAMIONETA", "3 TORTON", "4 TRAYLER", "5 PARTICULAR")
'On Error GoTo ERROR:
Select Case Index
    Case 0   'Clave del traslado
         txtcampos(Index).Text = Trim(txtcampos(Index).Text)
         txtcampos(Index).Refresh
         If LTrim(RTrim(txtcampos(0).Text)) = "" Or IsNull(txtcampos(0).Text) Then
            MsgBox "No puede dejar en blanco la clave del traslado", vbCritical
            txtcampos(0).SetFocus
            Exit Sub
         End If
         
         'Realizó la busqueda para determinar si es Alta o Modificación
         If nOp <> 1 Then  ' Modificaciones
           AdoTraslada.CursorType = adOpenDynamic
           AdoTraslada.ConnectionString = cCadConex
           AdoTraslada.RecordSource = "SELECT * FROM traslados WHERE T_CLAVE = '" & txtcampos(0).Text & "'"
           AdoTraslada.Refresh
         End If
         
         If Not AdoTraslada.Recordset!t_enviado And nOp = 3 Then
            MsgBox "NO ES POSIBLE CANCELAR EL TRASLADO PORQUE AUN NO SE HA ENVIADO Y POR LO TANTO NO SE HA DISMINUIDO EL INVENTARIO", vbExclamation
            Unload Me
            Exit Sub
         End If
         
         'Cargo Tabla de Catalogo de Sucursales y lleno el combo de Suc.
         AdoCatSuc.ConnectionString = cCadConex
         AdoCatSuc.CommandType = adCmdText
         AdoCatSuc.RecordSource = IIf(cModo = "DEVO", "SELECT * FROM Catprov", "SELECT * FROM CatTienda")
         AdoCatSuc.Refresh
 
         cmbSucEmi.Clear
         Do While Not AdoCatSuc.Recordset.EOF
            If cModo = "DEVO" Then
               If Not IsNull(AdoCatSuc.Recordset!NOMPROVE) And Not IsNull(AdoCatSuc.Recordset!prove) Then
                  cmbSucEmi.AddItem AdoCatSuc.Recordset!prove
                  cmbSucRec.AddItem AdoCatSuc.Recordset!NOMPROVE
               End If
            Else
               If Not IsNull(AdoCatSuc.Recordset!tidescrip) Then
                  cmbSucEmi.AddItem AdoCatSuc.Recordset!tidescrip
                  cmbSucRec.AddItem AdoCatSuc.Recordset!tidescrip
               End If
            End If
             AdoCatSuc.Recordset.MoveNext
         Loop
         fraGenerales.Visible = True
         lbletiquetas(1).Visible = True
         
         txtcampos(1).Visible = True
         cmbSucEmi.Visible = True
      If Forma = 1 Then 'Entradas
         txtcampos(1).SetFocus
      Else
         txtcampos(1).Text = Trim(Mid(cSucursal, 1, 3))
         cmbSucEmi.Text = Mid(cSucursal, 3)
         cmbSucEmi.Enabled = False
         txtcampos(1).Enabled = False
         txtcampos(0).Enabled = False
          
         lbletiquetas(2).Visible = True
         txtcampos(2).Visible = True
         cmbUsuEmi.Visible = True
         txtcampos(2).Text = Mid(cCveDesUsu, 1, 3)
         cmbUsuEmi.Text = Trim(Mid(cCveDesUsu, 3))
         txtcampos(2).SetFocus
         'SendKeys vbTab
         keybd_event &H9, 0, 0, 0
         'txtCampos(2).Enabled = False
      End If
      'Si ya fue enviado
                  If AdoTraslada.Recordset!t_enviado Then
         Call PONVOLUMEN
         For N = 0 To 9
            'SendKeys "{ENTER}"
            keybd_event &H9, 0, 0, 0
            keybd_event &H9, 0, &H2, 0
         Next
      End If
    Case 1    'Clave de la sucursal emisora
          If Forma = 1 Then
             txtcampos(Index).Text = Trim(txtcampos(Index).Text)
             txtcampos(Index).Refresh
             If txtcampos(Index).Text = "" Or IsNull(txtcampos(Index).Text) Then
                cmbSucEmi.SetFocus
                Exit Sub
             Else
                AdoCatSuc.Recordset.MoveFirst
                AdoCatSuc.Recordset.Find "TiClave= '" & txtcampos(Index).Text & "'"
                If AdoCatSuc.Recordset.EOF = True Then
                    'MsgBox "No existe la clave de la sucursal especificada", vbExclamation
                    Me.cmbSucEmi.SetFocus
                    Exit Sub
                Else
                    Me.cmbSucEmi.Text = AdoCatSuc.Recordset!tidescrip
                    txtcampos(2).Visible = True
                    cmbUsuEmi.Visible = True
                    lbletiquetas(2).Visible = True
                    lbletiquetas(2).Caption = "Persona receptora"
                    txtcampos(2).Text = Mid(cCveDesUsu, 1, 3)
                    cmbUsuEmi.Text = Trim(Mid(cCveDesUsu, 3))
                    txtcampos(2).Enabled = False
                    'Sucursal receptora  '
                    lbletiquetas(3).Visible = True
                    txtcampos(3).Visible = True
                    cmbSucRec.Visible = True
                    cmbSucRec.Text = Trim(Mid(cSucursal, 3))
                    txtcampos(3).Text = Trim(Mid(cSucursal, 1, 3))
                    txtcampos(3).SetFocus
                    'SendKeys vbTab
                    keybd_event &H9, 0, 0, 0
                End If
             End If
          End If
     Case 2    'Clave de la persona emisora
         txtcampos(Index).Text = Trim(txtcampos(Index).Text)
         txtcampos(Index).Refresh
         If txtcampos(Index).Text = "" Or IsNull(txtcampos(Index).Text) Then
            Me.cmbUsuEmi.SetFocus
            Exit Sub
         End If
         cmbUsuEmi.Text = Trim(Mid(cCveDesUsu, 3))
         lbletiquetas(3).Visible = True
         txtcampos(3).Visible = True
         If chkEnviado.Value = 0 Then txtcampos(3).SetFocus
         cmbSucRec.Visible = True
    
    Case 3  'Clave de la sucursal receptora
         txtcampos(Index).Text = Trim(txtcampos(Index).Text)
         txtcampos(Index).Refresh
         If txtcampos(Index).Text = "" Or IsNull(txtcampos(Index).Text) Then
            Me.cmbSucRec.SetFocus
            Exit Sub
         Else
            AdoCatSuc.Recordset.MoveFirst
            AdoCatSuc.Recordset.Find IIf(cModo = "DEVO", "PROVE= '" & txtcampos(Index).Text & "'", "TiClave= '" & txtcampos(Index).Text & "'")
            If AdoCatSuc.Recordset.EOF = True Then
                Me.cmbSucRec.SetFocus
                Exit Sub
            End If
         End If
         If cModo = "DEVO" Then
            cmbSucRec.Text = Trim(AdoCatSuc.Recordset.Fields!NOMPROVE)
         Else
            cmbSucRec.Text = Trim(AdoCatSuc.Recordset.Fields!tidescrip)
         End If
         'If txtCampos(3).Text = txtCampos(1).Text Then MsgBox "LOS PRODUCTOS CAPTURADOS A CONTINUACION SE CONSIDERARAN COMO AUTOCONSUMO", vbInformation
         'On Error Resume Next
         lfranquicia = AdoCatSuc.Recordset!franquicia
         cmbFlete.Visible = True
         lbletiquetas(5).Visible = True
         txtcampos(5).Visible = True
         txtcampos(5).SetFocus
         
         cmbFlete.Clear
         For N = 0 To 4
             cmbFlete.AddItem Tipos(N)
         Next
         If txtcampos(3).Text = "50" Then  'Si es clientes por default el flete es particular
            txtcampos(5).Text = "5"
            'SendKeys vbTab
            keybd_event &H9, 0, 0, 0
         End If
    
    Case 5  'Clave del tipo de camion
         txtcampos(Index).Text = Trim(txtcampos(Index).Text)
         txtcampos(Index).Refresh
         nPos = InStr(1, "01234", Val(txtcampos(Index).Text) - 1)
         If nPos = 0 Then
            cmbFlete.SetFocus
            Exit Sub
         End If
         cmbFlete.Text = Tipos(Val(txtcampos(Index).Text) - 1)
         For N = 4 To 6
            lbletiquetas(N).Visible = True
            txtcampos(N).Visible = True
         Next
         txtcampos(4).Enabled = False
         chkTipoTrasl.Visible = True
         chkPapeleria.Visible = True
         chkfrutas.Visible = True
         chkpan.Visible = True
         Check1.Visible = True
         chkMerma.Visible = True
         chkvolumen.Visible = True
         chkauto.Visible = True
         chkTipoTrasl.Enabled = (nOp = 1 And cModo <> "DEVO")
         If AdoTraslada.Recordset!t_tipo = True Then
            If AdoTraslada.Recordset!t_enviado = 0 Then txtcampos(6).Text = date + Time
         End If
         txtcampos(6).Enabled = False
    
         cmdGrabar.Visible = True
         cmdregresar.Visible = True
         'Solo cuando es un nuevo traslado
         If nOp = 1 And cModo <> "DEVO" Then
            chkTipoTrasl.SetFocus
         Else
            cmdGrabar.SetFocus
         End If
         
         'Si ya fue enviado desactivo todas los campos y botones
         If AdoTraslada.Recordset!t_enviado Then
            For N = 0 To 6
              txtcampos(N).Enabled = False
            Next
            cmbSucEmi.Enabled = False
            Me.cmbSucRec.Enabled = False
            Me.cmbFlete.Enabled = False
            Me.cmbUsuEmi.Enabled = False
            Me.DbgrdDetTraAbi.Columns(0).Button = False
            chkTipoTrasl.Enabled = False
            chkEnviado.Visible = True
            chkEnviado.Enabled = False
            chkPapeleria.Enabled = False
            chkfrutas.Enabled = False
            chkpan.Enabled = False
            Check1.Enabled = False
            chkMerma.Enabled = False
            chkvolumen.Enabled = False
            chkauto.Enabled = False
            'moy
            AdoDetPed.ConnectionString = cCadConex
            AdoDetPed.CommandType = adCmdText
            'Si es traslado cerrado =  por pedido
            If AdoTraslada.Recordset!t_tipo = False Then
               dbgrdDetPed.AllowUpdate = False
               DbgrdDetTraAbi.AllowUpdate = False
               DesgTraPed 'Cargo el detalle del traslado
               dbgrdDetPed.Visible = True
               DbgrdDetTraAbi.Visible = False
               FraPedEnv.Visible = True
               'Cargo los pedidos que se incluyen en el traslado, se utiliza la misma tabla
               AdoPedTra.ConnectionString = cCadConex
               AdoPedTra.CommandType = adCmdText
               AdoPedTra.RecordSource = "SELECT DISTINCT dt_pedido AS PEDIDOS FROM [DetalleTraslado]WHERE dt_clave = '" & txtcampos(0).Text & "' AND NOT DT_PEDIDO IS NULL"
               AdoPedTra.Refresh
            'Si el traslado es abierto
            Else
               DbgrdDetTraAbi.AllowAddNew = False
               dbgrdDetPed.AllowUpdate = False
               DbgrdDetTraAbi.AllowUpdate = False
               'AdoDetPed.RecordSource = "SELECT consec, descripc,LTrim(str(Paquetes)) + ' X ' + LTRIM( str(contenid,10,3)) + ' ' + MEDIDA as PRESENT, dt_clave, dt_producto, dt_cantidad, dt_cantidadp, dt_clave,Inprod, InCant as EXICAJ, dt_costo,Paquetes,InCantPza FROM DetalleTraslado,TFPRODUC, Inventario WHERE DetalleTraslado.dt_clave = '" & txtCampos(0).Text & "' AND dt_producto = consec AND CONSEC *= INPROD AND INPROD =* DT_PRODUCTO ORDER BY descripc,contenid"
               AdoDetPed.RecordSource = "SELECT consec, descripc,LTrim(str(Paquetes)) + ' X ' + LTRIM( str(contenid,10,3)) + ' ' + MEDIDA as PRESENT, dt_clave, dt_producto, dt_cantidad, dt_cantidadp, dt_clave,Inprod, CAST(InCant as SmallMoney) as EXICAJ, dt_costo, dt_venta, Paquetes,InCantPza FROM DetalleTraslado,TFPRODUC, Inventario WHERE DetalleTraslado.dt_clave = '" & txtcampos(0).Text & "' AND dt_producto = consec AND CONSEC *= INPROD AND INPROD =* DT_PRODUCTO ORDER BY descripc,contenid"
               AdoDetPed.Refresh
               Me.dbgrdDetPed.Visible = False
               Me.DbgrdDetTraAbi.Visible = True
               DbgrdDetTraAbi.AllowAddNew = False
               'Que permita borrar cancelar productos en traslados enviados
               'DbgrdDetTraAbi.AllowDelete = False
               'SendKeys "{ENTER}"
            End If
            PicBot.Visible = True
            PonPie   'Pongo Inf. referente al traslado
            'SendKeys "{ENTER}"
            SendKeys "{ESC}"
            cmdGrabar.Enabled = False
            CmdAgregar.Enabled = False
            cmdGrabarTra.Enabled = False
            'CmdActualizar.Enabled = False
            cmdCodBarra.Enabled = True
            cmbPedido.Enabled = False
            chkEnviado.Enabled = False
            chkPapeleria.Enabled = False
            chkfrutas.Enabled = False
            chkpan.Enabled = False
            Check1.Enabled = False
            chkMerma.Enabled = False
            chkvolumen.Enabled = False
            chkauto.Enabled = False
        End If
        If nOp = 3 Then  'Si se cancela el traslado completo
           Me.FraCancelar.Visible = True
           Me.CMDCORTAR.Enabled = False
           Me.cmdReporte.Enabled = False
           Me.cmdticket.Enabled = False
           txtMotivo.SetFocus
        End If
End Select
Exit Sub
Error:
    MsgBox Err.Description
End Sub


Private Sub txtCampos_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
     KeyAscii = 0
     'SendKeys "{TAB}"
     keybd_event &H9, 0, 0, 0
End If

If KeyAscii = 27 Then Unload Me
End Sub

Private Sub LstProd_LostFocus()
    '//Oculta la lista si pierde el enfoque
    Lstprod.Visible = False
End Sub

Private Sub LstProd_DblClick()
    lstProd_KeyPress vbKeyReturn
End Sub

Private Sub lstProd_KeyPress(KeyAscii As Integer)
On Error Resume Next
Dim cveprod As String
Dim N As Integer
    'Habilita el uso de teclado para asignar a la celda o Escape...
    Select Case KeyAscii
        Case vbKeyReturn
             'Asigna la clave del producto seleccionado a la celda
             N = InStr(1, Lstprod.List(Lstprod.ListIndex), "[")
             cveprod = Mid(Lstprod.List(Lstprod.ListIndex), N + 1, Len(Lstprod.List(Lstprod.ListIndex)) - N - 1)
             'Nuevo Modificó:Eric
             If Not (AdoDetPed.Recordset.BOF And AdoDetPed.Recordset.EOF) Then
                AdoDetPed.Recordset.MoveFirst
                AdoDetPed.Recordset.Find "CONSEC = '" & Trim(cveprod) & "'"
                If Not AdoDetPed.Recordset.EOF Then
                   Lstprod.Visible = False
                   MsgBox "EL PRODUCTO YA FUE CAPTURADO EN EL TRASLADO", vbInformation
                   Me.DbgrdDetTraAbi.SetFocus
                   keybd_event &H9, 0, 0, 0: keybd_event &H9, 0, &H2, 0:  keybd_event &H9, 0, 0, 0: keybd_event &H9, 0, &H2, 0:  keybd_event &H9, 0, 0, 0: keybd_event &H9, 0, &H2, 0:  keybd_event &H9, 0, 0, 0: keybd_event &H9, 0, &H2, 0:  keybd_event &H9, 0, 0, 0: keybd_event &H9, 0, &H2, 0
                   Exit Sub
                End If
             End If
             AdoDetPed.Recordset.AddNew
             AdoProduct.Recordset.MoveFirst
             AdoProduct.Recordset.Find "Consec = '" & Trim(cveprod) & "'"
             DbgrdDetTraAbi.Columns(0).Text = Trim(cveprod)
             DbgrdDetTraAbi.Columns(5).Text = 0
             DbgrdDetTraAbi.Columns(6).Text = 0
             Me.LblProdAgr.Caption = "PRODUCTO A AGREGAR: " & Space(20) & AdoProduct.Recordset!descripc & "  " & AdoProduct.Recordset!medida
             LblProdAgr.Refresh
             'lblDescrip.Caption = AdoProduct.Recordset.Fields!descripc
             Lstprod.Visible = False
             DbgrdDetTraAbi.SetFocus   'Para que se posicione en la columna de cajas
             'SendKeys "{TAB}": SendKeys "{TAB}": SendKeys "{TAB}": SendKeys "{TAB}": SendKeys "{TAB}"
             keybd_event &H9, 0, 0, 0: keybd_event &H9, 0, &H2, 0:  keybd_event &H9, 0, 0, 0: keybd_event &H9, 0, &H2, 0:  keybd_event &H9, 0, 0, 0: keybd_event &H9, 0, &H2, 0:  keybd_event &H9, 0, 0, 0: keybd_event &H9, 0, &H2, 0:  keybd_event &H9, 0, 0, 0: keybd_event &H9, 0, &H2, 0
        Case vbKeyEscape
             LblProdAgr.Caption = ""
             Lstprod.Visible = False
             DbgrdDetTraAbi.SetFocus   'Para que se posicione en la columna de cajas
    End Select
    
End Sub

Private Sub PonPie()
On Error Resume Next
Dim rsttemp As ADODB.Recordset
    Set rsttemp = New ADODB.Recordset
    rsttemp.Open "SELECT COUNT(dt_producto) AS TOTPROD, SUM(dt_cantidad) AS TOTCAJAS, SUM(dt_cantidadP) AS TOTPIEZAS  FROM DetalleTraslado WHERE DT_clave = '" & txtcampos(0).Text & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
    lblCajas.Visible = True
    If IsNull(rsttemp!TOTCAJAS) Or IsNull(rsttemp!TOTPIEZAS) Then
         lblCajas.Caption = "TOTAL DE PRODUCTOS: 0   CAJAS: 0   PIEZAS: 0"
    Else
        lblCajas.Caption = "PRODUCTOS: " & CStr(rsttemp!TOTPROD) & Space(3) & "CAJAS: " & CStr(rsttemp!TOTCAJAS) & Space(3) & "PIEZAS: " & CStr(rsttemp!TOTPIEZAS)
    End If
    lblCajas.Refresh
End Sub

Private Sub txtclamermas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If PORMERMAS Then
      If Trim(txtclamermas.Text) = "P4567" Then
         chkMerma.Value = 1
         chkauto.Value = 0
      Else
         chkMerma.Value = 0
         MsgBox "Contraseña Incorrecta, No es posible generar envio tipo autoconsumo", vbInformation, "TRASLADOS"
      End If
   ElseIf PORAUTO Then
      If Trim(txtclamermas.Text) = "P4567" Then
         chkauto.Value = 1
         chkMerma.Value = 0
      Else
         chkauto.Value = 0
         MsgBox "Contraseña Incorrecta, No es posible generar envio tipo autoconsumo", vbInformation, "TRASLADOS"
      End If
   End If
   PORAUTO = False
   PORMERMAS = False
   Framermas.Enabled = False
   Framermas.Visible = False
End If
End Sub

Private Sub txtContra_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    cmdConAceptar_Click
 ElseIf KeyAscii = 27 Then
    cmdConCance_Click
 End If
 
End Sub

