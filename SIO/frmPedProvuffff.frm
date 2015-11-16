VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form frmPedProv 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pedidos por proveedor pendientes de confimar"
   ClientHeight    =   8340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "frmPedProv.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc AdoPedpro 
      Height          =   330
      Left            =   360
      Top             =   5640
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
   Begin VB.Frame fraAvance 
      Caption         =   "Cargando especificaciones de productos "
      Height          =   855
      Left            =   4200
      TabIndex        =   41
      Top             =   3240
      Visible         =   0   'False
      Width           =   4935
      Begin MSComctlLib.ProgressBar Pgb 
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   360
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin VB.Frame fraCon 
      BackColor       =   &H80000001&
      Caption         =   "PROPORCIONE CONTRASEÑA"
      Height          =   1695
      Left            =   4560
      TabIndex        =   30
      Top             =   3120
      Visible         =   0   'False
      Width           =   4215
      Begin VB.CommandButton cmdConCance 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   2520
         TabIndex        =   33
         Top             =   1200
         Width           =   1050
      End
      Begin VB.CommandButton cmdConAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   840
         TabIndex        =   32
         Top             =   1200
         Width           =   1050
      End
      Begin VB.TextBox txtContra 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1920
         PasswordChar    =   "*"
         TabIndex        =   31
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label lblContra 
         BackColor       =   &H80000001&
         Caption         =   "Contraseña"
         Height          =   255
         Left            =   360
         TabIndex        =   34
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Frame fraAgrega 
      BackColor       =   &H80000001&
      Height          =   1695
      Left            =   2160
      TabIndex        =   35
      Top             =   240
      Visible         =   0   'False
      Width           =   9135
      Begin VB.TextBox txtPzaSol 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   4560
         TabIndex        =   38
         Text            =   "0"
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox txtCajSol 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1560
         TabIndex        =   37
         Text            =   "0"
         Top             =   1080
         Width           =   735
      End
      Begin VB.CommandButton cmdCanpro 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   7560
         TabIndex        =   40
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmdGrapro 
         Caption         =   "&Grabar"
         Height          =   375
         Left            =   6120
         TabIndex        =   39
         Top             =   1080
         Width           =   1215
      End
      Begin VB.ComboBox cmbprod 
         Height          =   315
         Left            =   240
         TabIndex        =   36
         Top             =   480
         Width           =   8775
      End
      Begin VB.Label lblEtiquetas 
         BackColor       =   &H80000001&
         Caption         =   "Piezas solicitadas"
         Height          =   255
         Index           =   6
         Left            =   3120
         TabIndex        =   44
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label lblEtiquetas 
         BackColor       =   &H80000001&
         Caption         =   "Cajas solicitadas"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   43
         Top             =   1080
         Width           =   1215
      End
   End
   Begin VB.CommandButton CmdAgregar 
      Caption         =   "&Agregar prod."
      Height          =   495
      Left            =   8280
      Picture         =   "frmPedProv.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7200
      Width           =   1215
   End
   Begin VB.TextBox txtCampos 
      DataField       =   "pp_observa"
      DataSource      =   "AdoPedProve"
      Height          =   615
      Index           =   5
      Left            =   1920
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   29
      Top             =   1320
      Width           =   5055
   End
   Begin VB.CommandButton cmdCodBarra 
      Caption         =   "Cod. &Barras"
      Height          =   495
      Left            =   9600
      Picture         =   "frmPedProv.frx":0404
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7200
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc AdoInventario 
      Height          =   330
      Left            =   9720
      Top             =   6000
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "AdoInventario"
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
   Begin MSAdodcLib.Adodc AdoFacturas 
      Height          =   330
      Left            =   7200
      Top             =   6000
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
      Caption         =   "AdoFacturas"
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
   Begin VB.TextBox txtRecib 
      Alignment       =   2  'Center
      DataField       =   "pp_fecrecibe"
      DataSource      =   "AdoPedProve"
      Height          =   285
      Index           =   0
      Left            =   9480
      TabIndex        =   27
      Top             =   960
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox txtRecib 
      Alignment       =   2  'Center
      DataField       =   "pp_notent"
      DataSource      =   "AdoPedProve"
      Height          =   285
      Index           =   3
      Left            =   7080
      TabIndex        =   25
      Top             =   1560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtRecib 
      Alignment       =   2  'Center
      DataField       =   "f_factura"
      DataSource      =   "AdoFacturas"
      Height          =   285
      Index           =   2
      Left            =   8760
      TabIndex        =   21
      Top             =   1560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtRecib 
      Alignment       =   2  'Center
      DataField       =   "f_total"
      DataSource      =   "AdoFacturas"
      Height          =   285
      Index           =   1
      Left            =   10200
      TabIndex        =   22
      Top             =   1560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CheckBox chkCampos 
      Caption         =   "Recibido"
      DataField       =   "pp_recibe"
      DataSource      =   "AdoPedProve"
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
      Index           =   1
      Left            =   7320
      TabIndex        =   20
      Top             =   960
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdReporte 
      Caption         =   "&Nota entrada"
      Height          =   495
      Left            =   7680
      Picture         =   "frmPedProv.frx":053A
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6600
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc AdoPedProve 
      Height          =   330
      Left            =   7200
      Top             =   5640
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
      Caption         =   "AdoPedProVed"
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
   Begin MSAdodcLib.Adodc AdoProv 
      Height          =   330
      Left            =   4920
      Top             =   5640
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
      Caption         =   "AdoProv"
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
   Begin VB.ComboBox cmbPerCon 
      Height          =   315
      Left            =   2640
      TabIndex        =   13
      Top             =   960
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.ComboBox cmbProv 
      Height          =   315
      Left            =   2640
      TabIndex        =   12
      Top             =   600
      Width           =   4335
   End
   Begin VB.TextBox txtCampos 
      DataField       =   "pp_perconfirma"
      DataSource      =   "AdoPedProve"
      Height          =   285
      Index           =   2
      Left            =   1920
      TabIndex        =   11
      Top             =   960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtCampos 
      DataField       =   "pp_fechagen"
      DataSource      =   "AdoPedProve"
      Height          =   285
      Index           =   3
      Left            =   5160
      TabIndex        =   8
      Top             =   240
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtCampos 
      DataField       =   "pp_proveedor"
      DataSource      =   "AdoPedProve"
      Height          =   285
      Index           =   1
      Left            =   1920
      TabIndex        =   7
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox txtCampos 
      DataField       =   "pp_pedido"
      DataSource      =   "AdoPedProve"
      Height          =   285
      Index           =   0
      Left            =   1920
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   495
      Left            =   10320
      Picture         =   "frmPedProv.frx":0A6C
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   495
      Left            =   9000
      Picture         =   "frmPedProv.frx":0BDE
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComctlLib.StatusBar StbMensajes 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   3
      Top             =   7995
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   609
      Style           =   1
      SimpleText      =   "                                                                                     Para salir presione la tecla [Esc]"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtCampos 
      Alignment       =   2  'Center
      DataField       =   "pp_fecconfirma"
      DataSource      =   "AdoPedProve"
      Height          =   285
      Index           =   4
      Left            =   9480
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CheckBox chkCampos 
      Caption         =   "&Confirmado"
      DataField       =   "pp_confirma"
      DataSource      =   "AdoPedProve"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   7320
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc AdoPedSol 
      Height          =   330
      Left            =   2640
      Top             =   5640
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
      Caption         =   "AdoPedSol"
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
   Begin MSDataGridLib.DataGrid dbgrdPedsol 
      Bindings        =   "frmPedProv.frx":0D50
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   6240
      Visible         =   0   'False
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   2990
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1.5
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "PEDIDOS QUE FUNDAMENTAN EL PEDIO POR PROVEEDOR"
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "folio"
         Caption         =   "FOLIO PED."
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
         DataField       =   "sucursal"
         Caption         =   "                            SUCURSAL"
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
      BeginProperty Column02 
         DataField       =   "Fecha_sol"
         Caption         =   "         FECHA ELAB."
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1289.764
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3495.118
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2085.166
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   7560
      ScaleHeight     =   615
      ScaleWidth      =   4215
      TabIndex        =   14
      Top             =   7560
      Width           =   4215
      Begin Crystal.CrystalReport CR1 
         Left            =   3720
         Top             =   480
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
         WindowLeft      =   0
         WindowTop       =   0
         WindowState     =   2
      End
   End
   Begin MSDataGridLib.DataGrid dbgrdpedpro 
      Bindings        =   "frmPedProv.frx":0D68
      Height          =   3975
      Left            =   120
      TabIndex        =   45
      Top             =   2040
      Visible         =   0   'False
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   7011
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      BorderStyle     =   0
      ColumnHeaders   =   -1  'True
      ForeColor       =   -2147483641
      HeadLines       =   1.5
      RowHeight       =   17
      TabAction       =   2
      RowDividerStyle =   6
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
      Caption         =   "DESGLOSE DEL PEDIDO"
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "dg_producto"
         Caption         =   "    CLAVE"
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
         DataField       =   "dg_cantsol"
         Caption         =   "CAJAS SOL."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "dg_cantsolp"
         Caption         =   "PZAS. SOL"
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
      BeginProperty Column03 
         DataField       =   "dg_cantreal"
         Caption         =   "CAJAS REC."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "dg_cantrealp"
         Caption         =   "PZAS. REC"
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Object.Visible         =   -1  'True
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1184.882
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1080
         EndProperty
      EndProperty
   End
   Begin VB.Label lblEtiquetas 
      Caption         =   "Observaciones"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   28
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label lblRec 
      Alignment       =   2  'Center
      Caption         =   "Nota de entrada"
      Height          =   255
      Index           =   3
      Left            =   7080
      TabIndex        =   26
      Top             =   1320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblRec 
      Alignment       =   2  'Center
      Caption         =   "Monto"
      Height          =   255
      Index           =   1
      Left            =   10200
      TabIndex        =   24
      Top             =   1320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblRec 
      Alignment       =   2  'Center
      Caption         =   "Num. Factura"
      Height          =   255
      Index           =   2
      Left            =   8520
      TabIndex        =   23
      Top             =   1320
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblEtiquetas 
      Caption         =   "Persona que confirma"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   960
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblEtiquetas 
      Caption         =   "Fecha de elaboracion"
      Height          =   255
      Index           =   3
      Left            =   3240
      TabIndex        =   9
      Top             =   240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblEtiquetas 
      Caption         =   "Proveedor"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   855
   End
   Begin VB.Label lblEtiquetas 
      Caption         =   "Clave del pedido global"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   1695
   End
End
Attribute VB_Name = "frmPedProv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private lMove As Boolean

Private Sub chkCampos_Click(Index As Integer)
If Index = 0 Then
   'If Not AdoPedProve.Recordset.BOF And AdoPedProve.Recordset.EOF Then
   If chkCampos(Index).Visible = True Then
      'MsgBox AdoPedProve.Recordset!pp_confirma = 0
      If AdoPedProve.Recordset!pp_confirma = 0 Then
         txtCampos(4).Text = Date + Time
         txtCampos(4).Enabled = False
      End If
         txtCampos(4).Visible = chkCampos(Index).Value = 1
   End If
ElseIf Index = 1 Then
    If Not (AdoPedProve.Recordset.BOF And AdoPedProve.Recordset.EOF) Then
      If AdoPedProve.Recordset!pp_recibe = 0 And chkCampos(Index).Visible Then
      txtRecib(0).Text = Date + Time
      txtRecib(0).Enabled = False
      End If
      For n = 0 To 3
          If n > 0 Then lblRec(n).Visible = chkCampos(Index).Value = 1
          If n < 3 Then txtRecib(n).Visible = chkCampos(Index).Value = 1
      Next
   End If
End If
End Sub

Private Sub cmbProv_DblClick()
  SendKeys "{TAB}"
End Sub

Private Sub cmbProv_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
If KeyAscii = 13 Then
   KeyAscii = 0
   SendKeys "{Tab}"
End If

End Sub

Private Sub cmbProv_Validate(Cancel As Boolean)
On Error GoTo Error:
Dim n As Integer
If nOp = 1 Then
   If cmbProv.Text = "" Or IsNull(cmbProv.Text) Then
       MsgBox "Debe seleccionar un proveedor de la lista desplegable", vbExclamation
       cmbProv.SetFocus
       Cancel = True
   Else
        AdoProv.Recordset.MoveFirst
        AdoProv.Recordset.Find "NomProve = '" & cmbProv.Text & "'"
        If AdoProv.Recordset.EOF = True Then
           MsgBox "Debe seleccionar un proveedor de la lista desplegable", vbExclamation
           cmbProv.SetFocus
           Cancel = True
        Else
        txtCampos(1).Text = AdoProv.Recordset!Prove
        End If
   End If
End If
Exit Sub
Error:
End Sub

Private Sub CmdAgregar_Click()
  cmdReporte.Enabled = False
  cmdAgregar.Enabled = False
  CmdGrabar.Enabled = False
  cmdCodBarra.Enabled = False
  cmdRegresar.Enabled = False
  fraCon.Visible = True
  txtContra.Text = ""
  txtContra.SetFocus
End Sub

Private Sub cmdCanpro_Click()
  fraAgrega.Visible = False
  cmdReporte.Enabled = True
  cmdAgregar.Enabled = True
  CmdGrabar.Enabled = True
  cmdCodBarra.Enabled = True
  cmdRegresar.Enabled = True
End Sub

Private Sub cmdCodBarra_Click()
  nOp = 0  'Para que la froma de lectura de codigos de barra sepa de donde se esta llamando
  'frmCodBarrCap.lblEtiquetas(2).Caption = dbgrdPedpro.Columns(1).Text + " " + dbgrdPedpro.Columns(2).Text
  frmCodBarrCap.Show 1
End Sub

Private Sub cmdConAceptar_Click()
Dim rsttemp As ADODB.Recordset
Dim nReg As Integer
If txtContra.Text <> "PITICO00" Then
   MsgBox "LA CONTRASEÑA DE ACCESO ES INCORRECTA", vbExclamation
   txtContra.SetFocus
Else
   fraCon.Visible = False
   Set rsttemp = New ADODB.Recordset
   rsttemp.Open "SELECT Descripc, LTRIM(STR(TFPRODUC.paquetes)) + ' X ' +  lTrim(str(TFPRODUC.contenid,10,3)) + space(2) + TFPRODUC.medida  AS MEDIDA, consec FROM TFPRODUC WHERE Claprove = '" & txtCampos(1).Text & "' ORDER BY Descripc", cn, adOpenKeyset, adLockOptimistic, adCmdText
   Pgb.Min = 0: nReg = 0
   Pgb.Max = rsttemp.RecordCount
   cmbProd.Clear
   fraAvance.Visible = True
   Me.Refresh
   While Not rsttemp.EOF
       nReg = nReg + 1
       Pgb.Value = nReg
       cmbProd.AddItem rsttemp!Descripc + "  " + rsttemp!Medida + "  " + rsttemp!CONSEC
       rsttemp.MoveNext
   Wend
   fraAvance.Visible = False
   fraAgrega.Visible = True
End If
End Sub

Private Sub cmdConCance_Click()
  fraCon.Visible = False
  cmdReporte.Enabled = True
  cmdAgregar.Enabled = True
  CmdGrabar.Enabled = True
  cmdCodBarra.Enabled = True
  cmdRegresar.Enabled = True
End Sub

Private Sub cmdGrabar_Click()
On Error GoTo Error
Dim FecConf As Date
Dim lback As Boolean
Dim rsttemp As ADODB.Recordset
Dim rstProPed As ADODB.Recordset
Dim lTrans As Boolean
  lTrans = False  'Flag de la Transaccion para saber cuando empieza
  lMove = False: lback = True ' Flag de Backorder
  'Cuando se confirma
  If chkCampos(0).Value = 0 Then
     MsgBox "Es necesario activar la casilla de pedido confirmado", vbExclamation
     chkCampos(0).SetFocus
     Exit Sub
  'Cuando se recibe
  ElseIf chkCampos(0).Value = 1 And chkCampos(1).Value = 0 And chkCampos(1).Visible = True Then
     MsgBox "Es necesario activar la casilla de pedido recibido", vbExclamation
     chkCampos(1).SetFocus
     Exit Sub

  End If
  'si se confirma el pedido
  If chkCampos(0).Value = 1 And chkCampos(1).Value = 0 Then
     'Grabo los datos generales del pedido por proveedor
     AdoPedProve.Recordset!pp_recibe = 0
     AdoPedProve.Recordset.Update
     'Agrego el detalle del pedido por proveedor para prepararlo para recibirlo
     Set rsttemp = New ADODB.Recordset
     rsttemp.ActiveConnection = cCadConex
     rsttemp.CursorType = adOpenDynamic
     rsttemp.LockType = adLockOptimistic
     rsttemp.Source = "DETALLEGLOBAL"
     rsttemp.Open
     AdoPedpro.Recordset.MoveFirst
     While Not AdoPedpro.Recordset.EOF
         rsttemp.AddNew
         rsttemp!dg_pedido = txtCampos(0).Text
         rsttemp!Dg_producto = AdoPedpro.Recordset!Clave
         rsttemp!dg_cantidad = AdoPedpro.Recordset!TotSol
         rsttemp!dg_cantrec = 0
         rsttemp.Update
         AdoPedpro.Recordset.MoveNext
     Wend
     'Confirmo los pedidos por tienda que incluye el Ped por proveedor
     rsttemp.Close
     rsttemp.Source = "SELECT * FROM PEDIDOS WHERE P_proveedor = '" & txtCampos(1).Text & "'  AND p_situacion = 0"
     rsttemp.Open
     While Not rsttemp.EOF
        rsttemp!p_fecConfirma = txtCampos(4).Text
        rsttemp!p_situacion = 1
        rsttemp!p_pedproveedor = txtCampos(0).Text
        rsttemp.Update
        rsttemp.MoveNext
     Wend
  
  'Si reciben el pedido
  ElseIf chkCampos(0).Value = 1 And chkCampos(1).Value = 1 Then
     If txtRecib(2).Text = "" Or IsNull(txtRecib(2).Text) Then
        MsgBox "DEBES ESPECIFICAR UN NUMERO DE FACTURA", vbExclamation
        txtRecib(2).SetFocus
        Exit Sub
     End If
     If Not IsNumeric(txtRecib(1).Text) Then
        MsgBox "EL MONTO DE LA FACTURA DEBE SER NUMERICO", vbCritical
        txtRecib(1).SetFocus
        Exit Sub
     End If

     cn.BeginTrans: lTrans = True
     FolNot = "N" + txtCampos(0).Text
         
     'Afecto inventario
     Set rstInvent = New ADODB.Recordset
     rstInvent.Open "Inventario", cn, adOpenDynamic, adLockOptimistic, adCmdTable
     Set rstProPed = New ADODB.Recordset
     rstProPed.Open "SELECT * FROM Tfproduc,Detalleglobal WHERE Tfproduc.Consec = DetalleGlobal.dg_producto AND DetalleGlobal.dg_pedido = '" & txtCampos(0).Text & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
     
     If Not AdoDetGlo.Recordset.EOF Then AdoDetGlo.Recordset.MoveFirst
     While Not AdoDetGlo.Recordset.EOF
        'Primero hago la busqueda en el catalogo de articulos para saber el numero de piezas de la caja
        rstProPed.MoveFirst
        rstProPed.Find "dg_producto = '" & AdoDetGlo.Recordset!Dg_producto & "'"
        If rstProPed.EOF Then
           MsgBox "EL ARTICULO CON CODIGO " & AdoDetGlo.Recordset!Dg_producto & "NO EXISTE EN EL CATALOGO DE PRODUCTOS" & _
           "A CONTINUACION SE DESAHARAN LOS CAMBIOS REALIZADOS", vbCritical
           cn.RollbackTrans
           Exit Sub
        End If
       'Ahora busco en el inventario
        rstInvent.MoveFirst
        rstInvent.Find "InProd = '" & AdoDetGlo.Recordset!Dg_producto & "'"
        If rstInvent.EOF Then
           MsgBox "NO EXISTE EN EL INVENTARIO EL ARTICULO " & Chr(13) & _
           AdoDetGlo.Recordset!Dg_producto & "  " & rstProPed!Descripc & Chr(13) & CStr(rstProPed!Paquetes) & " X " & CStr(rstProPed!Contenid) & " " & rstProPed!Medida & Chr(13) & _
           "A CONTINUACION SE DARA DE ALTA EN EL INVENTARIO", vbInformation
           rstInvent.AddNew
           rstInvent!Inprod = AdoDetGlo.Recordset!Dg_producto
           rstInvent!Insucursal = "0"
           rstInvent!InObserva = " "
           rstInvent!InFecCaduProx = "1/1/1900"
           rstInvent!inInicial = 0
           rstInvent!instock = 0
        End If
        rstInvent!Incant = rstInvent!Incant + (AdoDetGlo.Recordset!dg_cantreal * rstProPed!Paquetes) + AdoDetGlo.Recordset!dg_cantrealP
        rstInvent.Update
        'Si no surten lo solicitado puede haber backOrder
        If AdoDetGlo.Recordset!dg_cantsol <> AdoDetGlo.Recordset!dg_cantreal Or AdoDetGlo.Recordset!dg_cantsolP <> AdoDetGlo.Recordset!dg_cantrealP Then
           If lback Then   'Primera vez ver si existe Backorder
              rsttemp.Open "SELECT Nomprove, Backorder FROM CATPROV WHERE Prove = '" & AdoPedProve.Recordset!pp_proveedor & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
              If rsttemp!BackOrder Then
                 cResp = MsgBox("EL PROVEEDOR " & rsttemp!nomprove & Chr(13) & "SE LE PERMITEN VARIAS ENTREGAS." & Chr(13) & Space(15) & "DESEA CREAR BACKORDER?", vbYesNo + vbInformation)
                 If cResp = vbYes Then
                    'Obtengo el numero de backorder consecutivo
                    rsttemp.Close
                    rsttemp.Source = "SELECT MAX(NoBack) AS NumBack FROM detalleback"
                    rsttemp.Open
                    cn.Execute "INSERT INTO [DetalleBack](NoBack,producto,cantAsurtir,CantRecibida,cantasurtirp,cantrecibidap,pedidog,fecha,situacion) VALUES " & _
                               "('" & IIf(IsNull(rsttemp!NumBack), 1, rsttemp!NumBack) + 1 & "','" & AdoDetGlo.Recordset!Dg_producto & "','" & AdoDetGlo.Recordset!dg_cantsol - AdoDetGlo.Recordset!dg_cantreal & "','0','" & AdoDetGlo.Recordset!dg_cantsolP - AdoDetGlo.Recordset!dg_cantrealP & "','0','" & txtCampos(0).Text & "','" & Date & "','1' )"
                 End If
              End If
              lback = False
           ElseIf cResp = vbYes Then  'Si se acepta backorder grabo el detalle de Back.
                  cn.Execute "INSERT INTO [DetalleBack](NoBack,producto,cantAsurtir,CantRecibida,cantasurtirp,cantrecibidap,pedidog,fecha,situacion) VALUES " & _
                               "('" & IIf(IsNull(rsttemp!NumBack), 1, rsttemp!NumBack) + 1 & "','" & AdoDetGlo.Recordset!Dg_producto & "','" & AdoDetGlo.Recordset!dg_cantsol - AdoDetGlo.Recordset!dg_cantreal & "','0','" & AdoDetGlo.Recordset!dg_cantsolP - AdoDetGlo.Recordset!dg_cantrealP & "','0','" & txtCampos(0).Text & "','" & Date & "','1' )"
           End If
        End If
        AdoDetGlo.Recordset.MoveNext
     Wend
     
     'Obtengo los importes de cantidades solicitada y recibida
     Set rsttemp = New ADODB.Recordset
     rsttemp.Open "SELECT SUM(DetalleGlobal.dg_cantsol * Tfproduc.Precosto) + SUM(DetalleGlobal.dg_cantsolP * (Tfproduc.Precosto / tfProduc.Paquetes )) AS ImptSol, SUM(DetalleGlobal.dg_cantreal * Tfproduc.Precosto) + SUM(DetalleGlobal.dg_cantrealp * (Tfproduc.Precosto / tfProduc.Paquetes  )) AS ImptRec FROM DetalleGlobal,Tfproduc WHERE DetalleGlobal.dg_producto = Tfproduc.Consec AND DetalleGlobal.dg_pedido ='" & txtCampos(0).Text & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
     'Agrego Nota de entrada
     cn.Execute "INSERT INTO [NotaEntrada](Pedido,ClaveNota,ImporteFac,ImporteSol,ImporteRec,DifeProd, Difprecio, Factura) VALUES ('" & txtCampos(0).Text & "','" & FolNot & "','" & txtRecib(1).Text & "','" & rsttemp!ImptSol & "','" & rsttemp!ImptRec & "','" & rsttemp!ImptSol - rsttemp!ImptRec & "','" & Val(txtRecib(1) - rsttemp!ImptRec) & "','" & txtRecib(2).Text & "')"
     cn.Execute "INSERT INTO [DetalleNota](ClaveNota,Producto,Cantidad,cantidadp) SELECT CveNota = '" & FolNot & "', dg_producto, dg_cantreal, dg_cantrealp FROM [DetalleGlobal]WHERE dg_pedido = '" & txtCampos(0).Text & "'"

     AdoPedProve.Recordset!pp_pedido = txtCampos(0).Text
     AdoPedProve.Recordset!pp_proveedor = txtCampos(1).Text
     AdoPedProve.Recordset!pp_fechaGen = txtCampos(3).Text
     AdoPedProve.Recordset!pp_recibe = 1
     AdoPedProve.Recordset!pp_fecrecibe = txtRecib(0).Text
     AdoPedProve.Recordset!pp_NotEnt = Trim(FolNot)

     'Actualizo la tabla de facturas
     AdoFacturas.Recordset!f_pedido = txtCampos(0).Text
     AdoFacturas.Recordset!f_proveedor = txtCampos(1).Text
     AdoFacturas.Recordset!f_monto = 0
     AdoFacturas.Recordset!f_status = "R"
     MsgBox "A la nota de entrada generada se le asigo el folio " + FolNot, vbExclamation
     AdoPedProve.Recordset.Update
     AdoFacturas.Recordset.Update
     cn.CommitTrans
  End If
  Unload Me
Exit Sub
Error:
  MsgBox "OCURRIO EL SIGUIENTE ERROR: " + Chr(13) + UCase(Err.Description), vbCritical
  If lTrans = True Then
      MsgBox "A CONTINUACION SE DESHARAN LAS MODIFICACIONES REALIZADAS EN EL INVENTARIO", vbCritical
      cn.RollbackTrans
  End If
  Unload Me
End Sub

Private Sub cmdGrapro_Click()
On Error GoTo Error:
Dim rstExis As ADODB.Recordset
  If Not IsNumeric(txtCajSol.Text) Or Not IsNumeric(txtPzaSol.Text) Then
     MsgBox "LA CANTIDAD EN CAJAS Y PIEZAS DEBE SER NUMERICA", vbExclamation
     Exit Sub
  End If
  cClave = Trim(Mid(cmbProd.Text, Len(Trim(cmbProd.Text)) - 8))
  Set rstExis = New ADODB.Recordset
  rstExis.Open "SELECT * FROM DETALLEGLOBAL WHERE dg_producto = '" & cClave & "' AND dg_pedido = '" & txtCampos(0).Text & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
  If rstExis.RecordCount > 0 Then
     MsgBox "EL ARTICULO SELECCIONADO YA EXISTE EN EL PEDIDO", vbExclamation
     Exit Sub
  End If
  If MsgBox("REALMENTE DESEAS AGREGAR EL PRODUCTO " & Chr(13) & cmbProd.Text, vbQuestion + vbYesNo) = vbYes Then
     cn.Execute "INSERT INTO DETALLEGLOBAL(dg_pedido,dg_producto,dg_cantsol,dg_cantsolp) VALUES('" & txtCampos(0).Text & "','" & cClave & "'," & txtCajSol.Text & "," & txtPzaSol.Text & ")"
     AdoPedpro.Refresh
     AdoDetGlo.Refresh
     dbgrdpedpro.Columns(1).Width = 5560
  End If
  Exit Sub
Error:
  MsgBox Err.Description
End Sub

Private Sub cmdRegresar_Click()
Unload Me
End Sub

Private Sub cmdReporte_Click()
On Error GoTo Error:
cMensaje = stbMensajes.SimpleText
stbMensajes.SimpleText = Space(90) + "Espere un momento generando reporte..."
stbMensajes.Refresh
CR1.Connect = cCadConex
If cModo = "VERCONF" Then
    CR1.ReportFileName = App.Path & "\Pedprove.rpt"
    CR1.WindowTitle = "Pedido por proveedor numero " & txtCampos(0).Text
    CR1.Formulas(0) = "FORMSELEC = '" & txtCampos(0).Text & "'"
    CR1.Formulas(1) = "FECELAB = 'FECHA DE ELABORACION:  " & txtCampos(3).Text & "'"
    CR1.Formulas(2) = "NUMPED = 'NUMERO DE PEDIDO:  " & txtCampos(0).Text & "'"
    CR1.Formulas(3) = "PROVED = 'PROVEEDOR " & txtCampos(1).Text & Space(3) & frmpedBod.lblProve.Caption & "' "
    CR1.Formulas(4) = "FECCONF = 'FECHA DE CONFIRM.:  " & txtCampos(4).Text & "' "
Else
    CR1.ReportFileName = App.Path & "\prNotEnt.rpt"
    CR1.WindowTitle = "Nota de entrada del pedido " & txtCampos(0).Text
    CR1.Formulas(0) = "FORMSELEC = '" & txtCampos(0).Text & "'"
    CR1.Formulas(1) = "FACNUM = 'IMPORTE DE LA FACTURA NUM " & txtRecib(2).Text & "'"
    CR1.Formulas(2) = "FACMONTO = '" & Trim(txtRecib(1).Text) & "'"
    CR1.Formulas(3) = "PROVED = 'PROVEEDOR [ " & txtCampos(1).Text & Space(3) & frmpedBod.lblProve.Caption & " ]'"
    CR1.Formulas(4) = "FOLNOTENT = 'FOLIO " & Trim(txtRecib(3).Text) & "'"
End If
CR1.Action = 1
stbMensajes.SimpleText = cMensaje
stbMensajes.Refresh
Exit Sub
Error:
  MsgBox Err.Description
End Sub

Private Sub Command2_Click()

End Sub

Private Sub dbgrdPedpro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   SendKeys vbTab
End If
End Sub

Private Sub dbgrdRec_AfterColUpdate(ByVal ColIndex As Integer)
If ColIndex = 1 And AdoDetGlo.Recordset!dg_cantreal > AdoDetGlo.Recordset!dg_cantsol Then
   'Temporalmente se desactiva mientras se define lo de las promociones SAVE
   'MsgBox "LA CANTIDAD RECIBIDA EN CAJAS NO PUEDE SER MAYOR A LA SOLICITADA", vbExclamation
   'AdoDetGlo.Recordset!dg_cantreal = 0
ElseIf ColIndex = 2 And AdoDetGlo.Recordset!dg_cantrealP > AdoDetGlo.Recordset!dg_cantsolP Then
   'MsgBox "LA CANTIDAD RECIBIDA EN PIEZAS NO PUEDE SER MAYOR A LA SOLICITADA", vbExclamation
   'AdoDetGlo.Recordset!dg_cantrealP = 0
End If

End Sub

Private Sub dbgrdRec_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   SendKeys vbTab
End If
End Sub

Private Sub dbgrdRec_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
If lMove Then
   AdoPedpro.Recordset.MoveFirst
   AdoPedpro.Recordset.Find "clave = '" & AdoDetGlo.Recordset!Dg_producto & "'"
End If
End Sub

Private Sub Form_Load()
AdoPedProve.ConnectionString = cCadConex
AdoPedProve.CommandType = adCmdText
If Not (frmpedBod.AdoPedidos.Recordset.BOF = True And frmpedBod.AdoPedidos.Recordset.EOF) Then
  AdoPedProve.RecordSource = "SELECT * FROM Pedprove WHERE pp_pedido = '" & frmpedBod.dbgrdPed.Columns(0).Text & "'"
Else
  AdoPedProve.RecordSource = "SELECT * FROM Pedprove"
End If
AdoPedProve.Refresh

If nOp = 1 Then  'Nuevo
     AdoPedProve.Recordset.AddNew
     AdoProv.ConnectionString = cCadConex
     AdoProv.CommandType = adCmdText
     AdoProv.RecordSource = "SELECT * FROM Catprov"
     AdoProv.Refresh
     chkCampos(0).Value = 0: chkCampos(1).Value = 0
     cmbProv.Clear
     AdoProv.Recordset.MoveFirst
     Do While Not AdoProv.Recordset.EOF
        If Not IsNull(AdoProv.Recordset!nomprove) Then cmbProv.AddItem AdoProv.Recordset!nomprove
        AdoProv.Recordset.MoveNext
     Loop
     LBLETIQUETAS(1).Visible = True
     txtCampos(1).Visible = True
     cmbProv.Visible = True
     Me.dbgrdpedpro.Visible = False
ElseIf cModo = "RECIBIR" Then
    lMove = True 'Bandera que no hace el scroll al grabar si no se cicla
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  frmpedBod.Show
End Sub

Private Sub txtCampos_GotFocus(Index As Integer)
Select Case Index
Case 1
     frmpedBod.Hide
End Select
End Sub

Private Sub txtCampos_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case KeyAscii
     Case 27
         Unload Me
     Case 13
         KeyAscii = 0
         SendKeys "{Tab}"
End Select
End Sub

Private Sub txtCampos_LostFocus(Index As Integer)
Dim rstTie As ADODB.Recordset
Dim rsttemp As ADODB.Recordset
Dim ccadena As String
'On Error GoTo Error:
Select Case Index
Case 1 'Clave del proveedor
     txtCampos(Index).Text = Trim(UCase(txtCampos(Index).Text))
     txtCampos(Index).Refresh

     'NUEVO = Ver si existen pedidos para generar un pedido por proveedor
     If nOp = 1 Then
        If txtCampos(Index).Text = "" Or IsNull(txtCampos(Index).Text) Then
            cmbProv.SetFocus
            Exit Sub
        Else
            AdoProv.Recordset.MoveFirst
            AdoProv.Recordset.Find "Prove= '" & Trim(txtCampos(Index).Text) & "'"
            If AdoProv.Recordset.EOF = True Then
                'MsgBox "No existe la clave del proveedor especificado", vbExclamation
                cmbProv.SetFocus
                Exit Sub
            End If
        End If
        chkCampos(0).Visible = True
        cmbProv.Text = AdoProv.Recordset!nomprove
        'Obtengo el folio del pedido en base al proveedor Especif.
        Set rsttemp = New ADODB.Recordset
        rsttemp.ActiveConnection = cCadConex
        rsttemp.CursorType = adOpenKeyset
        rsttemp.Source = "SELECT MAX (CAST(SUBSTRING(PP_PEDIDO,5,10) AS INT)) As FolMay FROM [Pedprove] WHERE PP_proveedor = '" & txtCampos(1).Text & "'"
        rsttemp.Open
        If IsNull(rsttemp!FolMay) Then
            txtCampos(0).Text = Mid(Trim(Mid(txtCampos(1).Text, 1, 3)), 1, 3) + "-1"
        Else
            txtCampos(0).Text = txtCampos(1).Text + "-" + Trim(Str(rsttemp!FolMay + 1))
        End If
        cCond = "p_situacion = 0 AND p_proveedor = '" & txtCampos(1).Text & "'"
        'Muestro datos por default
        chkCampos(0).Value = 0
        txtCampos(3).Text = Date + Time
        cmbPerCon.Visible = True
        txtCampos(2).Text = Mid(cCveDesUsu, 1, 3)
        cmbPerCon.Text = Trim(Mid(cCveDesUsu, 3))
        cmbPerCon.Enabled = False
        For n = 0 To 3
            LBLETIQUETAS(n).Visible = True
            txtCampos(n).Visible = True
            txtCampos(n).Enabled = False
        Next
        cmbProv.Enabled = False
        dbgrdpedpro.Visible = True

        dbgrdPedsol.Visible = True
        dbgrdRec.Visible = False
        'save
        CmdGrabar.Visible = True
     Else 'Si es consulta del pedido confirmado O recepcion
        cCond = "p_pedproveedor = '" & frmpedBod.dbgrdPed.Columns(0).Text & "'"
        For n = 0 To 4
           If n < 4 Then LBLETIQUETAS(n).Visible = True
           txtCampos(n).Visible = True
           txtCampos(n).Enabled = False
        Next
        cmbProv.Visible = False
        chkCampos(0).Visible = True
        chkCampos(1).Visible = True
        chkCampos(0).Enabled = False
        chkCampos(1).Enabled = cModo = "RECIBIR"
        dbgrdpedpro.Visible = True
        
        dbgrdpedpro.Width = ScaleWidth - 400
        dbgrdPedsol.Visible = True
        If cModo = "RECIBIR" Then
           'dbgrdPedpro.Width = 9735
           cmdReporte.Visible = AdoPedProve.Recordset!pp_recibe
        Else
           cmdReporte.Visible = AdoPedProve.Recordset!pp_confirma
        End If
        AdoFacturas.ConnectionString = cCadConex
        AdoFacturas.CommandType = adCmdText
        AdoFacturas.RecordSource = "SELECT * FROM [Facturas] WHERE [f_pedido] = '" & txtCampos(0).Text & "'"
        AdoFacturas.Refresh
        If cModo = "RECIBIR" Then
            'dbgrdPedpro.AllowUpdate = False
            If Not AdoPedProve.Recordset!pp_recibe Then
               'No actualizo para que no se borre lo capturado por codigo de barras
               'cn.Execute "UPDATE DetalleGlobal SET dg_cantReal = dg_cantsol"
               'cn.Execute "UPDATE DetalleGlobal SET dg_cantRealp = dg_cantsolP"
               If AdoFacturas.Recordset.EOF = True Then AdoFacturas.Recordset.AddNew 'Para que no se borren los datos al ecribir en los campos del control adofacturas
               CmdGrabar.Visible = True
            Else
               For n = 0 To 3
                  txtRecib(n).Visible = True
                  txtRecib(n).Enabled = False
               Next
               chkCampos(1).Enabled = False
               txtRecib(0).Text = AdoPedProve.Recordset!pp_fecrecibe
               dbgrdRec.AllowUpdate = False
               CmdGrabar.Visible = True
               CmdGrabar.Enabled = False
               cmdCodBarra.Enabled = False
               cmdAgregar.Enabled = False
            End If
          '  AdoDetGlo.ConnectionString = cCadConex
          '  AdoDetGlo.CommandType = adCmdText
          '  AdoDetGlo.RecordSource = "SELECT * FROM DetalleGlobal, TFPRODUC WHERE dg_pedido = '" & txtCampos(0).Text & "' AND TFPRODUC.CONSEC = DETALLEGLOBAL.DG_PRODUCTO ORDER BY descripc, consec"
          '  AdoDetGlo.Refresh
          '  dbgrdRec.Visible = True
            chkCampos(1).Visible = True
        'Si es opcion ver confirmado y ya es recibido
        ElseIf AdoPedProve.Recordset!pp_recibe Then
            For n = 0 To 3
                txtRecib(n).Visible = True
                txtRecib(n).Enabled = False
            Next
            CmdGrabar.Visible = True
            CmdGrabar.Enabled = False
            cmdCodBarra.Enabled = False
            cmdAgregar.Enabled = False
        'Si es opcion ver confirmado y no ha sido recibido
        Else
            CmdGrabar.Visible = True
            CmdGrabar.Enabled = False
            cmdCodBarra.Enabled = False
            cmdAgregar.Enabled = False
        End If
     End If
     If cModo <> "RECIBIR" Then
        'Para crear ref. cruzada Sql Server recorro todas las tiendas
        cmdReporte.Caption = "Ped. Conf."
        Set rstTie = New ADODB.Recordset
        rstTie.Source = "SELECT * FROM cattienda"
        rstTie.ActiveConnection = cCadConex
        rstTie.Open
        'Genera la cadena del origen de datos es una referencia cruzada y
        'utiliza una vista de Sql. (DetPedTie)
        ccadena = "SELECT df_prod AS CLAVE, descripc As DESCRIPCION, str(paquetes) + ' X ' + LTRIM( str(contenid)) + ' ' + MEDIDA as ESPECIF," + Chr(13) _
        & " SUM(DF_CANTSOL) As TOTSOL,"
        While Not rstTie.EOF
            ccadena = ccadena + Chr(13) + " SUM(CASE p_sucursal WHEN '" & Trim(rstTie!Ticlave) & "' THEN df_cantsol ELSE 0 END) AS " & Mid(rstTie!Tidescrip, 1, 5) & ","
            rstTie.MoveNext
        Wend
        ccadena = Mid(ccadena, 1, Len(ccadena) - 1) & Chr(13) _
        & " From PedDetTie WHERE " & cCond _
        & " GROUP BY df_prod, descripc, str(paquetes) + ' X ' + LTRIM( str(contenid)) + ' ' + MEDIDA ORDER BY Descripcion"
        'Obtengo cantidades solicitadas por tienda de un proveedor especificado
        AdoPedpro.ConnectionString = cCadConex
        AdoPedpro.CommandType = adCmdText
        AdoPedpro.RecordSource = ccadena
        AdoPedpro.Refresh
        'Obtengo los folios de pedidos que solicitaron las tiendas
        AdoPedSol.ConnectionString = cCadConex
        AdoPedSol.CommandType = adCmdText
        AdoPedSol.RecordSource = "SELECT DISTINCT p_pedido as FOLIO, tidescrip AS SUCURSAL, P_fecPed As FECHA_SOL FROM PedDetTie WHERE " & cCond
        AdoPedSol.Refresh
        If AdoPedSol.Recordset.BOF And AdoPedSol.Recordset.EOF Then
           MsgBox "NO EXISTEN PEDIDOS PARA CONFIRMAR DEL PROVEEDOR ESPECIFICADO", vbExclamation
           Unload Me
           Exit Sub
        End If
     Else 'Si es recibir
        AdoPedpro.ConnectionString = cn.ConnectionString
        AdoPedpro.CommandType = adCmdText
        AdoPedpro.CursorType = adOpenDynamic
        AdoPedpro.RecordSource = "SELECT * FROM DETALLEGLOBAL WHERE dg_pedido = '" & txtCampos(0).Text & "'"
        
        AdoPedpro.Refresh
        Me.dbgrdpedpro.Refresh
        Me.dbgrdPedsol.Visible = False
     End If
     'If cModo = "RECIBIR" Then dbgrdPedpro.Columns(3).Width = 2150
     cmdRegresar.Visible = True
     'dbgrdPedpro.Columns(1).Width = 5560
     txtCampos(5).Enabled = False
     
End Select
Exit Sub
Error:
    MsgBox Err.Description
End Sub

Private Sub txtContra_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cmdConAceptar_Click
End If
End Sub

Private Sub txtRecib_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   SendKeys "{TAB}"
End If
End Sub
