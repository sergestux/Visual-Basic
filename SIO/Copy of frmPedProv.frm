VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPedProv 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pedidos por proveedor pendientes de confimar"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "frmPedProv.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame fraCon 
      BackColor       =   &H80000001&
      Caption         =   "PROPORCIONE CONTRASEÑA"
      Height          =   1695
      Left            =   4560
      TabIndex        =   22
      Top             =   3120
      Visible         =   0   'False
      Width           =   4215
      Begin VB.CommandButton cmdConCance 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   2520
         TabIndex        =   25
         Top             =   1200
         Width           =   1050
      End
      Begin VB.CommandButton cmdConAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   840
         TabIndex        =   24
         Top             =   1200
         Width           =   1050
      End
      Begin VB.TextBox txtContra 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1920
         PasswordChar    =   "*"
         TabIndex        =   23
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label lblContra 
         BackColor       =   &H80000001&
         Caption         =   "Contraseña"
         Height          =   255
         Left            =   360
         TabIndex        =   26
         Top             =   480
         Width           =   975
      End
   End
   Begin MSDataGridLib.DataGrid dbgrdfactu 
      Bindings        =   "frmPedProv.frx":030A
      Height          =   735
      Left            =   120
      TabIndex        =   44
      Top             =   2040
      Visible         =   0   'False
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   1296
      _Version        =   393216
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   2
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
      ColumnCount     =   20
      BeginProperty Column00 
         DataField       =   "Factura1"
         Caption         =   "FACTURA 1"
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
      BeginProperty Column01 
         DataField       =   "impfac1"
         Caption         =   "IMP. FAC. 1"
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
      BeginProperty Column02 
         DataField       =   "Factura2"
         Caption         =   "FACTURA 2"
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
         DataField       =   "impfac2"
         Caption         =   "IMP. FAC. 2"
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
      BeginProperty Column04 
         DataField       =   "factura3"
         Caption         =   "FACTURA 3"
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
         DataField       =   "impfac3"
         Caption         =   "IMP. FAC. 3"
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
      BeginProperty Column06 
         DataField       =   "factura4"
         Caption         =   "FACTURA 4"
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
      BeginProperty Column07 
         DataField       =   "impfac4"
         Caption         =   "IMP. FAC. 4"
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
         DataField       =   "factura5"
         Caption         =   "FACTURA 5"
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
      BeginProperty Column09 
         DataField       =   "impfac5"
         Caption         =   "IMP. FAC. 5"
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
      BeginProperty Column10 
         DataField       =   "factura6"
         Caption         =   "FACTURA 6"
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
      BeginProperty Column11 
         DataField       =   "impfac6"
         Caption         =   "IMP. FAC. 6"
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
      BeginProperty Column12 
         DataField       =   "Factura7"
         Caption         =   "FACTURA 7"
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
      BeginProperty Column13 
         DataField       =   "impfac7"
         Caption         =   "IMP. FAC. 7"
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
      BeginProperty Column14 
         DataField       =   "factura8"
         Caption         =   "FACTURA 8"
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
      BeginProperty Column15 
         DataField       =   "impfac8"
         Caption         =   "IMP. FAC. 8"
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
      BeginProperty Column16 
         DataField       =   "factura9"
         Caption         =   "FACTURA 9"
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
      BeginProperty Column17 
         DataField       =   "impfac9"
         Caption         =   "IMP. FAC. 9"
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
      BeginProperty Column18 
         DataField       =   "factura10"
         Caption         =   "FACTURA10"
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
      BeginProperty Column19 
         DataField       =   "impfac10"
         Caption         =   "IMP.FAC.10"
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
            ColumnWidth     =   1110.047
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1124.787
         EndProperty
         BeginProperty Column07 
            Alignment       =   2
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1110.047
         EndProperty
         BeginProperty Column09 
            Alignment       =   2
            ColumnWidth     =   1124.787
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column11 
            Alignment       =   2
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   1035.213
         EndProperty
         BeginProperty Column13 
            Alignment       =   2
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column15 
            Alignment       =   2
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   1035.213
         EndProperty
         BeginProperty Column17 
            Alignment       =   2
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column18 
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column19 
            ColumnWidth     =   975.118
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox PicBoton 
      Align           =   2  'Align Bottom
      BackColor       =   &H8000000C&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   11850
      TabIndex        =   14
      Top             =   7545
      Width           =   11910
      Begin VB.CommandButton CmdRefresh 
         Caption         =   "&Actualizar"
         Height          =   495
         Left            =   4800
         Picture         =   "frmPedProv.frx":0324
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Actualizar pedido"
         Top             =   120
         Width           =   1000
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         Height          =   495
         Left            =   8400
         Picture         =   "frmPedProv.frx":0426
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Grabar pedido e incrementar Inventario"
         Top             =   120
         Visible         =   0   'False
         Width           =   1000
      End
      Begin VB.CommandButton cmdRegresar 
         Caption         =   "&Regresar"
         Height          =   495
         Left            =   9600
         Picture         =   "frmPedProv.frx":0598
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Regresar a la pantalla principal de pedidos por proveedor"
         Top             =   120
         Visible         =   0   'False
         Width           =   1000
      End
      Begin VB.CommandButton cmdReporte 
         Caption         =   "&Nota ent."
         Height          =   495
         Left            =   1200
         Picture         =   "frmPedProv.frx":070A
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "reporte de Nota de entrada"
         Top             =   120
         Width           =   1000
      End
      Begin VB.CommandButton cmdCodBarra 
         Caption         =   "Cod. &Barra"
         Height          =   495
         Left            =   6000
         Picture         =   "frmPedProv.frx":0C3C
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Verificar codigo de barras de productos"
         Top             =   120
         Width           =   1000
      End
      Begin VB.CommandButton CmdAgregar 
         Caption         =   "&Agr. prod."
         Height          =   495
         Left            =   7200
         Picture         =   "frmPedProv.frx":0D72
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Agregar producto al pedido"
         Top             =   120
         Width           =   1000
      End
      Begin VB.CommandButton cmdExporta 
         Caption         =   "&Exportar"
         Height          =   495
         Left            =   3600
         Picture         =   "frmPedProv.frx":0E6C
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Exportar pedido a formato DBF"
         Top             =   120
         Width           =   1000
      End
      Begin VB.CommandButton cmdRpteMix 
         Caption         =   "&Ped. Sug."
         Height          =   495
         Left            =   2400
         Picture         =   "frmPedProv.frx":0F6E
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Reporte de pedidos sugeridos a surtirse en bodega"
         Top             =   120
         Width           =   1000
      End
      Begin Crystal.CrystalReport CR1 
         Left            =   240
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
         WindowLeft      =   0
         WindowTop       =   0
         WindowState     =   2
      End
   End
   Begin MSComDlg.CommonDialog cmdlg 
      Left            =   3240
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AdoDbf 
      Height          =   330
      Left            =   720
      Top             =   5040
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
      Caption         =   "AdoDbf"
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
      TabIndex        =   33
      Top             =   3240
      Visible         =   0   'False
      Width           =   4935
      Begin MSComctlLib.ProgressBar Pgb 
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   360
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin VB.Frame fraAgrega 
      BackColor       =   &H80000001&
      Height          =   1695
      Left            =   2280
      TabIndex        =   27
      Top             =   3240
      Visible         =   0   'False
      Width           =   9135
      Begin VB.TextBox txtPzaSol 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   4560
         TabIndex        =   30
         Text            =   "0"
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox txtCajSol 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1560
         TabIndex        =   29
         Text            =   "0"
         Top             =   1080
         Width           =   735
      End
      Begin VB.CommandButton cmdCanpro 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   7560
         TabIndex        =   32
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmdGrapro 
         Caption         =   "&Grabar"
         Height          =   375
         Left            =   6120
         TabIndex        =   31
         Top             =   1080
         Width           =   1215
      End
      Begin VB.ComboBox cmbprod 
         Height          =   315
         Left            =   240
         TabIndex        =   28
         Top             =   480
         Width           =   8775
      End
      Begin VB.Label lblEtiquetas 
         BackColor       =   &H80000001&
         Caption         =   "Piezas solicitadas"
         Height          =   255
         Index           =   6
         Left            =   3120
         TabIndex        =   36
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label lblEtiquetas 
         BackColor       =   &H80000001&
         Caption         =   "Cajas solicitadas"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   35
         Top             =   1080
         Width           =   1215
      End
   End
   Begin VB.TextBox txtCampos 
      DataField       =   "pp_observa"
      DataSource      =   "AdoPedProve"
      Height          =   615
      Index           =   5
      Left            =   1920
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   21
      Top             =   1320
      Width           =   9615
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
      TabIndex        =   18
      Top             =   600
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox txtRecib 
      Alignment       =   2  'Center
      DataField       =   "pp_notent"
      DataSource      =   "AdoPedProve"
      Height          =   285
      Index           =   3
      Left            =   9480
      TabIndex        =   16
      Top             =   960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox chkCampos 
      Caption         =   "Re&cibido"
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
      TabIndex        =   15
      Top             =   600
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc AdoDetGlo 
      Height          =   330
      Left            =   9720
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
      CursorType      =   1
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
      Caption         =   "AdoDetGlo"
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
      Alignment       =   2  'Center
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
      Alignment       =   2  'Center
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
      Alignment       =   2  'Center
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
   Begin MSComctlLib.StatusBar StbMensajes 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   3
      Top             =   8280
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
      Top             =   120
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
      Height          =   255
      Index           =   0
      Left            =   7320
      TabIndex        =   1
      Top             =   120
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
   Begin MSAdodcLib.Adodc AdoPedpro 
      Height          =   330
      Left            =   240
      Top             =   5640
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   582
      ConnectMode     =   3
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
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
      Caption         =   "AdoPedPro"
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
      Bindings        =   "frmPedProv.frx":14A0
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
   Begin MSDataGridLib.DataGrid dbgrdPedpro 
      Bindings        =   "frmPedProv.frx":14B8
      Height          =   4455
      Left            =   120
      TabIndex        =   19
      Top             =   2880
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   7858
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      HeadLines       =   1.5
      RowHeight       =   15
      TabAction       =   2
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
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "dg_producto"
         Caption         =   "CLAVE"
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
         DataField       =   "descripc"
         Caption         =   "DESCRIPCION"
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
         DataField       =   "MEDIDA"
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
      BeginProperty Column03 
         DataField       =   "DG_CANTSOL"
         Caption         =   "CAJAS SOL"
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
      BeginProperty Column04 
         DataField       =   "DG_CANTSOLP"
         Caption         =   "PZAS.SOL"
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
      BeginProperty Column05 
         DataField       =   "dg_promocion"
         Caption         =   "PROM.SOL."
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
      BeginProperty Column06 
         DataField       =   "dg_cantreal"
         Caption         =   "CAJ.REC"
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
         DataField       =   "DG_CANTREALP"
         Caption         =   "PZAS.REC"
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
      BeginProperty Column08 
         DataField       =   "DG_PROMOCIONR"
         Caption         =   "PROM.REC"
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   2
         Size            =   311
         BeginProperty Column00 
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   870.236
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   4694.74
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   1484.787
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   929.764
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   884.976
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   959.811
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            ColumnWidth     =   824.882
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   884.976
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
            ColumnWidth     =   989.858
         EndProperty
      EndProperty
   End
   Begin VB.Label lblEtiquetas 
      Caption         =   "Observaciones"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   20
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label lblRec 
      Alignment       =   2  'Center
      Caption         =   "Nota de entrada"
      Height          =   255
      Index           =   3
      Left            =   7560
      TabIndex        =   17
      Top             =   960
      Visible         =   0   'False
      Width           =   1335
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
Private lModPed As Boolean
Private lMove As Boolean

Private Sub AdoPedpro_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.ERROR, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error Resume Next
If Not lMove Then Exit Sub
RSTDESPRO.MoveFirst
RSTDESPRO.Find "CONSEC = '" & AdoPedpro.Recordset!clave & "'"
If RSTDESPRO.EOF Then
   lblDespro.Caption = ""
Else
   lblDespro.Caption = Trim(AdoPedpro.Recordset!Descripcion) & Chr(13) & " " & Trim(AdoPedpro.Recordset!Present) & "  " & "PROMOCION: " & CStr(RSTDESPRO!cajas) & "/" & CStr(RSTDESPRO!encajas)
End If
End Sub

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
      If AdoPedProve.Recordset!PP_RECIBE = 0 And chkCampos(Index).Visible Then
         txtRecib(0).Text = Date + Time
         txtRecib(0).Enabled = False
      Else
        lblRec(3).Visible = True
      End If
      'For n = 0 To 3
      '    If n > 0 Then lblRec(n).Visible = chkCampos(Index).Value = 1
      '    If n < 3 Then txtRecib(n).Visible = chkCampos(Index).Value = 1
      'Next
      txtRecib(0).Visible = chkCampos(Index).Value = 1
      Me.dbgrdfactu.Visible = chkCampos(Index).Value = 1
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
On Error GoTo ERROR:
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
ERROR:
End Sub

Private Sub CmdAgregar_Click()
  cmdReporte.Enabled = False
  cmdAgregar.Enabled = False
  cmdGrabar.Enabled = False
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
  cmdGrabar.Enabled = True
  cmdCodBarra.Enabled = True
  cmdRegresar.Enabled = True
End Sub

Private Sub cmdCodBarra_Click()
  nOp = 0  'Para que la forma de lectura de codigos de barra sepa de donde se esta llamando
  'frmCodBarrCap.lblEtiquetas(2).Caption = dbgrdPedpro.Columns(1).Text + " " + dbgrdPedpro.Columns(2).Text
  frmCodBarrCap.Show 1
End Sub

Private Sub cmdConAceptar_Click()
Dim rsttemp As ADODB.Recordset
Dim nReg As Integer
If txtContra.Text <> "VYL2000" Then
   MsgBox "LA CONTRASEÑA DE ACCESO ES INCORRECTA", vbExclamation
   txtContra.SetFocus
Else
   If Not lModPed Then
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
            cmbProd.AddItem rsttemp!Descripc + "  " + rsttemp!MEDIDA + "  " + rsttemp!cONSEC
            rsttemp.MoveNext
        Wend
        fraAvance.Visible = False
        fraAgrega.Visible = True
   Else
       Me.fraCon.Visible = False
       dbgrdPedpro.AllowUpdate = True
   End If
End If
End Sub

Private Sub cmdConCance_Click()
  fraCon.Visible = False
  cmdReporte.Enabled = True
  cmdAgregar.Enabled = True
  cmdGrabar.Enabled = True
  cmdCodBarra.Enabled = True
  cmdRegresar.Enabled = True
End Sub

'Exporta pedidos DETALLE DE LA NOTA DE ENTRADA a dbf
'Con los pedidos recibidos en carbonera para enviarlos a Oficinas centrales
Private Sub CmdExporta_Click()
On Error GoTo ERROR:
Dim nDias As Integer
Dim FolPed As String
Dim rsttemp As ADODB.Recordset
Dim cnFoxPro As ADODB.Connection
   cMenAnt = stbMensajes.SimpleText
   Cmdlg.DialogTitle = "Grabar archivo para enviar pedidos a Oficinas centrales"
   Cmdlg.Filter = "Archivos Visual Fox pro (*.dbf) | *.dbf"
   Cmdlg.CancelError = True
   Cmdlg.ShowSave
   cRutArc = Cmdlg.FileName
   If cRutArc = "" Or IsNull(cRutArc) Then
      MsgBox "DEBE ESPECIFICAR UNA RUTA Y NOMBRE DE ARCHIVO", vbExclamation
      Exit Sub
   End If
   stbMensajes.SimpleText = Space(45) & "Grabando archivo " & cRutArc
   stbMensajes.Refresh
   
   For n = 1 To Len(cRutArc)
      If Mid(cRutArc, n, 1) = "\" Then nPos = n
   Next
   cRuta = Mid(cRutArc, 1, nPos)
   cArch = Mid(cRutArc, nPos + 1)
   cArch = Mid(cArch, 1, Len(cArch) - 4) 'Le quito la extension porque si no marca error en la consula SQL.

   If MsgBox("DESEAS LIMPIAR EL ARCHIVO A ENVIAR", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
        stbMensajes.SimpleText = Space(65) & "Limpiando archivo " & cArch
        stbMensajes.Refresh
        Set fs = CreateObject("Scripting.FileSystemObject")
        Set f = fs.GetFile("C:\PASO\ESPEDCAR.DBF")
        f.Copy cRutArc, True
   End If

   Set rsttemp = New ADODB.Recordset
   AdoDbf.CommandType = adCmdText
   AdoDbf.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=PITICODBF;SourceDB=" & cRuta
   AdoDbf.RecordSource = "SELECT * FROM " & cArch
   AdoDbf.Refresh
     
   rsttemp.Open "SELECT * FROM DetalleNota WHERE ClaveNota = '" & txtRecib(3).Text & "'", cn, adOpenStatic, adLockOptimistic, adCmdText
   While Not rsttemp.EOF
      stbMensajes.SimpleText = Space(75) & "Exportando producto con la clave: " & CStr(rsttemp!producto)
      stbMensajes.Refresh
      AdoDbf.Recordset.AddNew
      AdoDbf.Recordset!Clavenota = rsttemp!Clavenota
      AdoDbf.Recordset!producto = rsttemp!producto
      AdoDbf.Recordset!CantSolc = rsttemp!cantsol
      AdoDbf.Recordset!cantsolp = rsttemp!cantsolp
      AdoDbf.Recordset!CantRecC = rsttemp!cantrec
      AdoDbf.Recordset!cantrecp = rsttemp!cantrecp
      AdoDbf.Recordset!costo = rsttemp!costo
      AdoDbf.Recordset!ImpFac = Trim(txtRecib(1).Text)
      AdoDbf.Recordset!sucursal = Trim(Mid(cSucursal, 1, 3))
      AdoDbf.Recordset!FecRec = txtRecib(0).Text
      AdoDbf.Recordset!Importado = False
      AdoDbf.Recordset.Update
      rsttemp.MoveNext
   Wend
   AdoDbf.Recordset.Close
   stbMensajes.SimpleText = cMenAnt
  Exit Sub
ERROR:
   If Err.Number <> 32755 Then  'Numero de error al presionar el boton cancelar
      MsgBox Err.Description
   End If
   stbMensajes.SimpleText = cMenAnt

End Sub

Private Sub cmdGrabar_Click()
On Error GoTo ERROR
Dim FecConf As Date
Dim lback As Boolean
Dim rsttemp As ADODB.Recordset
Dim rstProPed As ADODB.Recordset
Dim rstInvent As ADODB.Recordset
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
     AdoPedProve.Recordset!PP_RECIBE = 0
     AdoPedProve.Recordset.Update
     'Agrego el detalle del pedido por proveedor para prepararlo para recibirlo
     Set rsttemp = New ADODB.Recordset
     rsttemp.ActiveConnection = cCadConex
     rsttemp.CursorType = adOpenKeyset
     rsttemp.LockType = adLockOptimistic
     rsttemp.Source = "DETALLEGLOBAL"
     rsttemp.Open
     AdoPedpro.Recordset.MoveFirst
     While Not AdoPedpro.Recordset.EOF
         rsttemp.AddNew
         rsttemp!dg_pedido = txtCampos(0).Text
         rsttemp!dg_producto = AdoPedpro.Recordset!clave
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
     
     Set rsttemp = New ADODB.Recordset
     cn.BeginTrans: lTrans = True
     FolNot = "N" + txtCampos(0).Text
     
     stbMensajes.SimpleText = Space(25) & "Cargando inventario para su actualizacion"
     stbMensajes.Refresh
     'Afecto inventario
     Set rstInvent = New ADODB.Recordset
     Set rstProPed = New ADODB.Recordset
     
     'If Not (AdoPedpro.Recordset.BOF And AdoPedpro.Recordset.EOF) Then
     AdoPedpro.Recordset.MoveFirst
     While Not AdoPedpro.Recordset.EOF
        
        'Primero hago la busqueda en el catalogo de articulos para saber el numero de piezas de la caja
        rstProPed.Open "SELECT * FROM TFPRODUC WHERE CONSEC = '" & AdoPedpro.Recordset!dg_producto & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
        If rstProPed.BOF And rstProPed.EOF Then
           MsgBox "EL ARTICULO CON CODIGO " & AdoPedpro.Recordset!dg_producto & "NO EXISTE EN EL CATALOGO DE PRODUCTOS" & _
           "A CONTINUACION SE DESAHARAN LOS CAMBIOS REALIZADOS", vbCritical
           cn.RollbackTrans
           Exit Sub
        End If
       
       'Ahora busco en el inventario
        rstInvent.Open "SELECT * FROM INVENTARIO WHERE Inprod = '" & AdoPedpro.Recordset!dg_producto & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
        If rstInvent.BOF And rstInvent.EOF Then
           MsgBox "NO EXISTE EN EL INVENTARIO EL ARTICULO " & Chr(13) & _
           AdoPedpro.Recordset!dg_producto & "  " & rstProPed!Descripc & Chr(13) & CStr(rstProPed!PAQUETES) & " X " & CStr(rstProPed!CONTENID) & " " & rstProPed!MEDIDA & Chr(13) & _
           "A CONTINUACION SE DARA DE ALTA EN EL INVENTARIO", vbInformation
           rstInvent.AddNew
           rstInvent!Inprod = AdoPedpro.Recordset!dg_producto
           rstInvent!Insucursal = "3"
           rstInvent!InObserva = " "
           rstInvent!InFecCaduProx = "1/1/1900"
           rstInvent!InInicial = 0
           rstInvent!instock = 0
        End If
        cn.Execute "UPDATE detalleGlobal SET dg_existencia = " & rstInvent!InCant & " WHERE dg_producto = '" & AdoPedpro.Recordset!dg_producto & "' AND dg_pedido = '" & txtCampos(0).Text & "'"
        'Sumo cantidad recibida en Cajas, piezas y promociones
        rstInvent!InCant = rstInvent!InCant + AdoPedpro.Recordset!dg_cantreal + AdoPedpro.Recordset!dg_promocionr
        rstInvent.Update
        'Si no surten lo solicitado puede haber backOrder
        If AdoPedpro.Recordset!dg_cantsol > AdoPedpro.Recordset!dg_cantreal Or AdoPedpro.Recordset!dg_cantsolp <> AdoPedpro.Recordset!dg_cantrealP Then
           If lback Then   'Primera vez ver si existe Backorder
              rsttemp.Open "SELECT Nomprove, Backorder FROM CATPROV WHERE Prove = '" & AdoPedProve.Recordset!PP_PROVEEDOR & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
              If rsttemp!BackOrder Then
                 cResp = MsgBox("AL PROVEEDOR " & rsttemp!NOMPROVE & Chr(13) & "SE LE PERMITEN VARIAS ENTREGAS." & Chr(13) & Space(15) & "DESEA CREAR BACKORDER?", vbYesNo + vbInformation)
                 If cResp = vbYes Then
                    'Obtengo el numero de backorder consecutivo
                    rsttemp.Close
                    rsttemp.Source = "SELECT pp_pedBack, MAX(pp_NumBack) AS NumBack FROM Pedprove WHERE pp_pedback = '" & IIf(Mid(txtCampos(0).Text, 1, 1) = "B", AdoPedProve.Recordset!PP_PEDBACK, txtCampos(0).Text) & "' GROUP BY pp_pedback"
                    rsttemp.Open
                    'If IsNull(rsttemp!pp_pedback) Then
                    If rsttemp.RecordCount = 0 Then
                       FOLBACK = 1
                       FolPed = "B" & CStr(FOLBACK) & txtCampos(0).Text
                    Else
                       FOLBACK = rsttemp!NumBack + 1
                       FolPed = "B" & CStr(FOLBACK) & rsttemp!PP_PEDBACK
                    End If

                    cn.Execute "INSERT INTO [PedProve](pp_proveedor,pp_pedido,pp_fechagen,pp_confirma,pp_recibe,pp_numback,pp_pedback ) VALUES " & _
                               "('" & txtCampos(1).Text & "','" & FolPed & "','" & Date + Time & "',1,0," & FOLBACK & ",'" & IIf(FOLBACK = 1, txtCampos(0).Text, rsttemp!PP_PEDBACK) & "')"
                    cn.Execute "INSERT INTO [DetalleGlobal](dg_pedido,dg_producto,dg_cantsol,dg_cantsolp,dg_cantreal,dg_costo) VALUES " & _
                               "('" & FolPed & "','" & AdoPedpro.Recordset!dg_producto & "','" & AdoPedpro.Recordset!dg_cantsol - AdoPedpro.Recordset!dg_cantreal & "',0,0," & AdoPedpro.Recordset!DG_COSTO & ")"
                 End If
              End If
              lback = False
           ElseIf cResp = vbYes Then  'Si se acepta backorder grabo el detalle de Back.
                  cn.Execute "INSERT INTO [DetalleGlobal](dg_pedido,dg_producto,dg_cantsol,dg_cantsolp,dg_cantreal,dg_costo) VALUES " & _
                               "('" & FolPed & "','" & AdoPedpro.Recordset!dg_producto & "','" & AdoPedpro.Recordset!dg_cantsol - AdoPedpro.Recordset!dg_cantreal & "',0,0," & AdoPedpro.Recordset!DG_COSTO & ")"
           End If
        End If
        AdoPedpro.Recordset.MoveNext
        rstInvent.Close: rstProPed.Close
     Wend
     stbMensajes.SimpleText = Space(25) & "Generando Nota de entrada"
     stbMensajes.Refresh
     'Obtengo los importes de cantidades solicitada y recibida
     Set rsttemp = New ADODB.Recordset
     rsttemp.Open "SELECT SUM(DetalleGlobal.dg_cantsol * detalleglobal.dg_costo) + SUM(DetalleGlobal.dg_cantsolP * (dg_costo / tfProduc.Paquetes )) AS ImptSol, SUM(DetalleGlobal.dg_cantreal * Tfproduc.Precosto) + SUM(DetalleGlobal.dg_cantrealp * (Tfproduc.Precosto / tfProduc.Paquetes  )) AS ImptRec FROM DetalleGlobal,Tfproduc WHERE DetalleGlobal.dg_producto = Tfproduc.Consec AND DetalleGlobal.dg_pedido ='" & txtCampos(0).Text & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
     nTot = 0
     'AdoFacturas.Refresh
     nTot = nTot + AdoFacturas.Recordset!Impfac1
     nTot = nTot + IIf(IsNull(AdoFacturas.Recordset!Impfac2), 0, AdoFacturas.Recordset!Impfac2)
     nTot = nTot + IIf(IsNull(AdoFacturas.Recordset!Impfac3), 0, AdoFacturas.Recordset!Impfac3)
     nTot = nTot + IIf(IsNull(AdoFacturas.Recordset!Impfac4), 0, AdoFacturas.Recordset!Impfac4)
     nTot = nTot + IIf(IsNull(AdoFacturas.Recordset!Impfac5), 0, AdoFacturas.Recordset!Impfac5)
     nTot = nTot + IIf(IsNull(AdoFacturas.Recordset!Impfac6), 0, AdoFacturas.Recordset!Impfac6)
     nTot = nTot + IIf(IsNull(AdoFacturas.Recordset!Impfac7), 0, AdoFacturas.Recordset!Impfac7)
     nTot = nTot + IIf(IsNull(AdoFacturas.Recordset!Impfac8), 0, AdoFacturas.Recordset!Impfac8)
     nTot = nTot + IIf(IsNull(AdoFacturas.Recordset!Impfac9), 0, AdoFacturas.Recordset!Impfac9)
     nTot = nTot + IIf(IsNull(AdoFacturas.Recordset!Impfac10), 0, AdoFacturas.Recordset!Impfac10)
     'Agrego Nota de entrada
     cn.Execute "UPDATE [NotaEntrada] SET ImporteSol = " & rsttemp!ImptSol & ", ImporteRec = " & rsttemp!ImptRec & ", DifeProd = " & rsttemp!ImptSol - rsttemp!ImptRec & ", DifPrecio = " & nTot - rsttemp!ImptRec & " WHERE pedido = '" & Trim(txtCampos(0).Text) & "'"
     cn.Execute "INSERT INTO [DetalleNota](ClaveNota,Producto,Cantsol,cantsolp,Cantrec,CantrecP,costo) SELECT CveNota = '" & FolNot & "', dg_producto, dg_cantsol, dg_cantsolp, dg_cantreal, dg_cantrealp, dg_costo FROM [DetalleGlobal] WHERE dg_pedido = '" & txtCampos(0).Text & "'"
            
     AdoPedProve.Recordset!pP_PEDIDO = txtCampos(0).Text
     AdoPedProve.Recordset!PP_PROVEEDOR = txtCampos(1).Text
     AdoPedProve.Recordset!PP_FECHAGEN = txtCampos(3).Text
     AdoPedProve.Recordset!PP_RECIBE = 1
     AdoPedProve.Recordset!pp_fecrecibe = txtRecib(0).Text
     AdoPedProve.Recordset!pp_NotEnt = Trim(FolNot)
     stbMensajes.SimpleText = Space(25) & "Espere un momento, realizando prorrateo de pedidos sugeridos"
     stbMensajes.Refresh

     ActualizaSug AdoPedProve.Recordset!pP_PEDIDO
     
     MsgBox "A LA NOTA DE ENTRADA GENERADA SE LE ASIGNO EL FOLIO " + FolNot, vbExclamation
     AdoPedProve.Recordset.Update
     cn.CommitTrans
  End If
  Unload Me
Exit Sub
ERROR:
  MsgBox "OCURRIO EL SIGUIENTE ERROR: " + Chr(13) + UCase(Err.Description), vbCritical
  If lTrans = True Then
      MsgBox "A CONTINUACION SE DESHARAN LAS MODIFICACIONES REALIZADAS EN EL INVENTARIO", vbCritical
      cn.RollbackTrans
      AdoPedProve.Recordset!PP_RECIBE = 0
      AdoPedProve.Recordset.Update
  End If
  Unload Me
End Sub

Private Sub cmdGrapro_Click()
On Error GoTo ERROR:
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
  End If
  Exit Sub
ERROR:
  MsgBox Err.Description
End Sub

Private Sub CmdRefresh_Click()
  AdoPedpro.Refresh
End Sub

Private Sub cmdRegresar_Click()
Unload Me
End Sub

Private Sub cmdReporte_Click()
On Error GoTo ERROR:
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
    CR1.Formulas(3) = "PROVED = 'PROVEEDOR " & txtCampos(1).Text & Space(3) & frmpedBod.cmbProved.Text & "' "
    CR1.Formulas(4) = "FECCONF = 'FECHA DE CONFIRM.:  " & txtCampos(4).Text & "' "
Else
    CR1.ReportFileName = App.Path & "\prNotEnt.rpt"
    CR1.WindowTitle = "Nota de entrada del pedido " & txtCampos(0).Text
    CR1.Formulas(0) = "FORMSELEC = '" & txtCampos(0).Text & "'"
    CR1.Formulas(3) = "PROVED = 'PROVEEDOR [ " & txtCampos(1).Text & Space(3) & frmpedBod.cmbProved.Text & " ]'"
    CR1.Formulas(4) = "FOLNOTENT = 'FOLIO " & Trim(txtRecib(3).Text) & "'"
End If
CR1.Action = 1
stbMensajes.SimpleText = cMensaje
stbMensajes.Refresh
Exit Sub
ERROR:
  MsgBox Err.Description
End Sub

Private Sub cmdRpteMix_Click()
On Error GoTo ERROR:
 cMensaje = stbMensajes.SimpleText
 stbMensajes.SimpleText = Space(90) + "Espere un momento generando reporte..."
 stbMensajes.Refresh
 CR1.Connect = cCadConex
 cARcRpt = App.Path & "\provesug.rpt"
 ccondrpt = "FORMSELEC = {PEDIDOS.p_pedproveedor} = '" & txtCampos(0).Text & "'"
 
 CR1.WindowTitle = "Sugeridos del pedido por proveedor con folio " & txtCampos(0).Text
 CR1.ReportFileName = cARcRpt
 CR1.SQLQuery = "SELECT PEDIDOS.p_pedido, PEDIDOS.p_proveedor, PEDIDOS.p_fecped, PEDIDOS.p_fecconfirma, PEDIDOS.p_pedproveedor, PEDIDOS.p_traslado, " & _
                        "DETALLEFACTURA.df_cantreal, DETALLEFACTURA.df_cantsol, DETALLEFACTURA.df_cantsolp, " & _
                        "CATPROV.NOMPROVE, CATTIENDA.tidescrip, TFPRODUC.CLAPROVE, TFPRODUC.DESCRIPC, TFPRODUC.CONTENID, TFPRODUC.MEDIDA, TFPRODUC.PAQUETES, TFPRODUC.CONSEC " & Chr(13) & _
                "FROM pitico.dbo.PEDIDOS PEDIDOS, " & _
                        "pitico.dbo.DETALLEFACTURA DETALLEFACTURA, " & _
                        "pitico.dbo.CATPROV CATPROV, " & _
                        "pitico.dbo.CATTIENDA CATTIENDA, " & _
                        "pitico.dbo.TFPRODUC TFPRODUC " & Chr(13) & _
                "WHERE PEDIDOS.p_pedido = DETALLEFACTURA.df_pedido AND " & _
                        "PEDIDOS.p_proveedor = CATPROV.PROVE AND " & _
                        "PEDIDOS.p_sucursal = CATTIENDA.ticlave AND " & _
                        "DETALLEFACTURA.df_prod = TFPRODUC.CONSEC AND " & _
                        "PEDIDOS.p_pedproveedor = '" & Trim(txtCampos(0).Text) & "' AND TFPRODUC.CLAPROVE = '" & Trim(txtCampos(1).Text) & "' " & Chr(13) & _
                "ORDER BY PEDIDOS.p_pedido ASC, TFPRODUC.DESCRIPC ASC, TFPRODUC.CONTENID ASC"
 'MsgBox CR1.SQLQuery
 CR1.Formulas(0) = ccondrpt
 CR1.Formulas(1) = "FECELAB = 'FECHA DE ELABORACION:  " & txtCampos(3).Text & "'"
 CR1.Formulas(2) = "NUMPED = 'NUMERO DE PEDIDO:  " & txtCampos(0).Text & "'"
 CR1.Formulas(3) = "FECCONF = 'FECHA DE CONFIRM.:  " & IIf(chkCampos(0).Value, txtCampos(4).Text, "") & "' "
 CR1.Formulas(4) = ""
 CR1.Action = 1
 stbMensajes.SimpleText = cMensaje
 stbMensajes.Refresh
 Exit Sub
ERROR:
 MsgBox Err.Description
End Sub


Private Sub dbgrdfactu_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   SendKeys vbTab
End If
End Sub

Private Sub dbgrdPedpro_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
Dim marca
marca = AdoPedpro.Recordset.Bookmark
'If Not AdoPedProve.Recordset!pp_recibe Then
    If UCase(dbgrdPedpro.Columns(ColIndex).DataField) = "DG_CANTREAL" Then
        cn.Execute "UPDATE DETALLEGLOBAL SET DG_CANTREAL = " & dbgrdPedpro.Columns(ColIndex).Text & " WHERE DG_PRODUCTO = '" & AdoPedpro.Recordset!dg_producto & "' AND DG_PEDIDO = '" & Trim(txtCampos(0).Text) & "'"
    ElseIf UCase(dbgrdPedpro.Columns(ColIndex).DataField) = "DG_CANTREALP" Then
        cn.Execute "UPDATE DETALLEGLOBAL SET DG_CANTREALP = " & dbgrdPedpro.Columns(ColIndex).Text & " WHERE DG_PRODUCTO = '" & AdoPedpro.Recordset!dg_producto & "' AND DG_PEDIDO = '" & Trim(txtCampos(0).Text) & "'"
    ElseIf UCase(dbgrdPedpro.Columns(ColIndex).DataField) = "DG_PROMOCIONR" Then
        cn.Execute "UPDATE DETALLEGLOBAL SET DG_PROMOCIONR = " & dbgrdPedpro.Columns(ColIndex).Text & " WHERE DG_PRODUCTO = '" & AdoPedpro.Recordset!dg_producto & "' AND DG_PEDIDO = '" & Trim(txtCampos(0).Text) & "'"
    End If
'Else 'Si ya es recibido
'    nCant = dbgrdPedpro.Columns(6).Text
'   If MsgBox("CONFIRMA SI SE ACTUALIZA EL PEDIDO, INVENTARIO Y NOTA DE ENTRADA", vbQuestion + vbYesNo) = vbYes Then
'    If UCase(dbgrdPedpro.Columns(ColIndex).DataField) = "DG_CANTREAL" Then
'       cn.Execute "UPDATE DETALLEGLOBAL SET DG_CANTREAL = " & dbgrdPedpro.Columns(ColIndex).Text & " WHERE DG_PRODUCTO = '" & AdoPedpro.Recordset!dg_producto & "' AND DG_PEDIDO = '" & Trim(txtCampos(0).Text) & "'"
'       cn.Execute "UPDATE inventario SET INCANT = INCANT + " & dbgrdPedpro.Columns(6).Text - AdoPedpro.Recordset!dg_cantreal & " WHERE inprod = '" & AdoPedpro.Recordset!dg_producto & "'"
'       cn.Execute "UPDATE Detallenota SET cantrec = " & nCant & " WHERE producto = '" & AdoPedpro.Recordset!dg_producto & "' AND clavenota = '" & "N" + Trim(txtCampos(0).Text) & "'"
'       MsgBox "ACTUALIZADO...", vbInformation
'    End If
'   End If
'End If
AdoPedpro.Refresh
dbgrdPedpro.Refresh
AdoPedpro.Recordset.Bookmark = marca
Cancel = True
SendKeys "{DOWN}"
End Sub

Private Sub dbgrdRec_AfterColUpdate(ByVal ColIndex As Integer)
'If ColIndex = 1 And AdoDetGlo.Recordset!dg_cantreal > AdoDetGlo.Recordset!dg_cantsol Then
   'Temporalmente se desactiva mientras se define lo de las promociones SAVE
   'MsgBox "LA CANTIDAD RECIBIDA EN CAJAS NO PUEDE SER MAYOR A LA SOLICITADA", vbExclamation
   'AdoDetGlo.Recordset!dg_cantreal = 0
'ElseIf ColIndex = 2 And AdoDetGlo.Recordset!dg_cantrealP > AdoDetGlo.Recordset!dg_cantsolp Then
   'MsgBox "LA CANTIDAD RECIBIDA EN PIEZAS NO PUEDE SER MAYOR A LA SOLICITADA", vbExclamation
   'AdoDetGlo.Recordset!dg_cantrealP = 0
'End If

End Sub

Private Sub dbgrdRec_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   SendKeys vbTab
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 119 Then frmCalc.Show
'If KeyCode = 116 Then
'   lModPed = True
'   fraCon.Visible = True
'   txtContra.SetFocus
'End If
End Sub

Private Sub Form_Load()
On Error GoTo ERROR:
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
        If Not IsNull(AdoProv.Recordset!NOMPROVE) Then cmbProv.AddItem AdoProv.Recordset!NOMPROVE
        AdoProv.Recordset.MoveNext
     Loop
     lbletiquetas(1).Visible = True
     txtCampos(1).Visible = True
     cmbProv.Visible = True
     Me.dbgrdPedpro.Visible = False
ElseIf cModo = "RECIBIR" Then
    lMove = True 'Bandera que no hace el scroll al grabar si no se cicla
End If
Exit Sub
ERROR:
  MsgBox Err.Description
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
On Error GoTo ERROR:
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
        cmbProv.Text = AdoProv.Recordset!NOMPROVE
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
            lbletiquetas(n).Visible = True
            txtCampos(n).Visible = True
            txtCampos(n).Enabled = False
        Next
        cmbProv.Enabled = False
        dbgrdPedpro.Visible = True

        dbgrdPedsol.Visible = True
        dbgrdRec.Visible = False
        'save
        cmdGrabar.Visible = True
     Else 'Si es consulta del pedido confirmado O recepcion
        cCond = "p_pedproveedor = '" & frmpedBod.dbgrdPed.Columns(0).Text & "'"
        For n = 0 To 4
           If n < 4 Then lbletiquetas(n).Visible = True
           txtCampos(n).Visible = True
           txtCampos(n).Enabled = False
        Next
        cmbProv.Visible = False
        chkCampos(0).Visible = True
        chkCampos(1).Visible = True
        chkCampos(0).Enabled = False
        chkCampos(1).Enabled = cModo = "RECIBIR"
        dbgrdPedpro.Visible = True
        
        'dbgrdPedpro.Width = ScaleWidth - 400
        dbgrdPedsol.Visible = True
        If cModo = "RECIBIR" Then
           'dbgrdPedpro.Width = 9735
           cmdReporte.Visible = AdoPedProve.Recordset!PP_RECIBE
        Else
           cmdReporte.Visible = AdoPedProve.Recordset!pp_confirma
        End If
        AdoFacturas.ConnectionString = cCadConex
        AdoFacturas.CommandType = adCmdText
        AdoFacturas.RecordSource = "SELECT * FROM [NOTAENTRADA] WHERE [pedido] = '" & txtCampos(0).Text & "'"
        AdoFacturas.Refresh
        If cModo = "RECIBIR" Then
            'dbgrdPedpro.AllowUpdate = False
            If Not AdoPedProve.Recordset!PP_RECIBE Then
               'No actualizo para que no se borre lo capturado por codigo de barras
               'cn.Execute "UPDATE DetalleGlobal SET dg_cantReal = dg_cantsol"
               'cn.Execute "UPDATE DetalleGlobal SET dg_cantRealp = dg_cantsolP"
               If AdoFacturas.Recordset.BOF And AdoFacturas.Recordset.EOF = True Then
                  AdoFacturas.Recordset.AddNew 'Para que no se borren los datos al ecribir en los campos del control adofacturas
                  AdoFacturas.Recordset!Pedido = txtCampos(0).Text
                  AdoFacturas.Recordset!Clavenota = "N" & Trim(txtCampos(0).Text)
                  AdoFacturas.Recordset.Update
                  AdoFacturas.Refresh
               End If
               cmdGrabar.Visible = True
            Else
               'For n = 0 To 3
                  txtRecib(3).Visible = True
                  lblRec(3).Visible = True
                  Me.dbgrdfactu.AllowUpdate = False
                '  txtRecib(n).Enabled = False
               'Next
               chkCampos(1).Enabled = False
               txtRecib(0).Text = AdoPedProve.Recordset!pp_fecrecibe
               'dbgrdRec.AllowUpdate = False
               cmdGrabar.Visible = True
               cmdGrabar.Enabled = False
               cmdCodBarra.Enabled = False
               cmdAgregar.Enabled = False
               dbgrdPedpro.AllowUpdate = False
            End If
            chkCampos(1).Visible = True
        'Si es opcion ver confirmado y ya es recibido
        ElseIf AdoPedProve.Recordset!PP_RECIBE Then
            For n = 0 To 3
                txtRecib(n).Visible = True
                txtRecib(n).Enabled = False
            Next
            cmdGrabar.Visible = True
            cmdGrabar.Enabled = False
            cmdCodBarra.Enabled = False
            cmdAgregar.Enabled = False
        'Si es opcion ver confirmado y no ha sido recibido
        Else
            cmdGrabar.Visible = True
            cmdGrabar.Enabled = False
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
            ccadena = ccadena + Chr(13) + " SUM(CASE p_sucursal WHEN '" & Trim(rstTie!Ticlave) & "' THEN df_cantsol ELSE 0 END) AS " & Mid(rstTie!TIDESCRIP, 1, 5) & ","
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
        AdoPedpro.ConnectionString = cCadConex
        AdoPedpro.CommandType = adCmdText
        'AdoPedpro.RecordSource = "SELECT dg_producto AS CLAVE, descripc AS DESCRIPCION, LTRIM(STR(TFPRODUC.paquetes)) + ' X ' +  lTrim(str(TFPRODUC.contenid,10,3)) + space(2) + TFPRODUC.medida AS MEDIDA, dg_cantsol AS SOLxCAJA, dg_cantsolp AS SOLxPZA, dg_promocion AS PROM, DETALLEGLOBAL.dg_cantreal AS CAJ_REC, dg_cantrealP AS PZA_REC, dg_promocionR AS PROM_REC FROM DetalleGlobal,tfproduc WHERE dg_pedido = '" & txtCampos(0).Text & "' AND DETALLEGLOBAL.dg_producto = TFPRODUC.consec ORDER BY DESCRIPC,MEDIDA"
        AdoPedpro.RecordSource = "SELECT dg_producto, descripc, LTRIM(STR(TFPRODUC.paquetes)) + ' X ' +  lTrim(str(TFPRODUC.contenid,10,3)) + space(2) + TFPRODUC.medida AS MEDIDA, dg_cantsol, dg_cantsolp, dg_promocion, DETALLEGLOBAL.dg_cantreal, dg_cantrealP, dg_promocionR, dg_costo, dg_existencia FROM DetalleGlobal,tfproduc WHERE dg_pedido = '" & txtCampos(0).Text & "' AND DETALLEGLOBAL.dg_producto = TFPRODUC.consec ORDER BY DESCRIPC,MEDIDA"
        AdoPedpro.Refresh
        Me.dbgrdPedpro.Refresh
        Me.dbgrdPedsol.Visible = False
     End If
     'If cModo = "RECIBIR" Then dbgrdPedpro.Columns(3).Width = 2150
     cmdRegresar.Visible = True
     cmdRpteMix.Visible = True
     'dbgrdPedpro.Columns(1).Width = 5560
     txtCampos(5).Locked = True
End Select
Exit Sub
ERROR:
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

'Actualiza los pedidos sugeridos que formaron el pedido por proveedor
'y se generan inmediatamente un envio para que lo facturen
Private Sub ActualizaSug(PedProve As String)
Dim nPor
Dim rs As ADODB.Recordset
Dim rstemp As ADODB.Recordset
Dim rstRepart As ADODB.Recordset
Dim lSurExi As Boolean
Set rs = New ADODB.Recordset
Set rstemp = New ADODB.Recordset
'Se recorre todo el detalle del pedido global y se van poniendo cantidades a los sugeridos en base al
'porcentaje que representa lo solicitado en el sug. a lo recibido en el ped. por proveedor

If Trim(txtCampos(1).Text) <> "JAR" Then
  AdoPedpro.Recordset.MoveFirst
  rs.Open "SELECT SUM(dg_cantsol) AS SurBod FROM detalleglobal WHERE dg_pedido = '" & txtCampos(0).Text & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
  lSurExi = (rs!SurBod = 0)
  rs.Close
  While Not AdoPedpro.Recordset.EOF
    If lSurExi Then  'Todo se va a surtir de las existencias de bodega
       rs.Open "SELECT df_pedido,df_prod, df_cantsol, df_cantreal FROM detallefactura,pedidos WHERE df_pedido = p_pedido AND P_PEDPROVEEDOR = '" & PedProve & "' AND DF_PROD = '" & AdoPedpro.Recordset!dg_producto & "' AND p_sucursal <> 10", cn, adOpenKeyset, adLockOptimistic, adCmdText
       nCanSug = 0
       nSolOfi = 0
       'SI existe un sugerido de oficinas centrales se reparte entre las demas tiendas
       Set rstRepart = New ADODB.Recordset
       rstRepart.Open "SELECT * FROM detallefactura,pedidos WHERE df_pedido = p_pedido AND P_PEDPROVEEDOR = '" & PedProve & "' AND DF_PROD = '" & AdoPedpro.Recordset!dg_producto & "' AND P_SUCURSAL = 10", cn, adOpenKeyset, adLockOptimistic, adCmdText
       If rstRepart.RecordCount > 0 Then
          nSolOfi = Round(rstRepart!df_cantsol / rs.RecordCount, 0)
       End If

       While Not rs.EOF
           rstemp.Open "SELECT SUM(df_cantsol) AS Cansol, InCant FROM pedidos,detallefactura,inventario WHERE p_pedproveedor = '" & txtCampos(0).Text & "' AND DF_PEDIDO = P_PEDIDO AND DF_PROD = INPROD AND df_prod = '" & AdoPedpro.Recordset!dg_producto & "' GROUP BY InCant", cn, adOpenKeyset, adLockOptimistic, adCmdText
           nPor = 100
           If rstemp!InCant < rstemp!cansol And rstemp!InCant > 0 Then 'Se surte equitativamente entre todas las tiendas lo que hay en el inventario
              nPor = (rs!df_cantsol) * 100 / rstemp!InCant
              ncan = Round(rstemp!cansol * nPor / 100, 0)
           Else
              ncan = Round((rs!df_cantsol + nSolOfi) * nPor / 100, 0)
           End If
           nCanSug = nCanSug + ncan
           cProAnt = rs!df_prod
           If nCanSug > rstemp!InCant Then  'Cuando ya no alcanza el inventario para surtir
              rs!df_cantreal = rstemp!InCant - (nCanSug - ncan)
              rs.Update
              rs.MoveLast 'se agotaron las existencias se mueve al final paraque ya no siga asignando
           ElseIf nCanSug > rstemp!cansol Then  'Cuando hay mas inventario de lo solicitado
              rs!df_cantreal = rstemp!cansol - (nCanSug - ncan)
              rs.Update
           Else
              rs!df_cantreal = ncan
              rs.Update
           End If
           rs.MoveNext
           rstemp.Close
       Wend
       rs.Close
    Else  'Se reparte de lo recibido en el pedprove
        If AdoPedpro.Recordset!dg_cantreal > 0 Then
           'Aunque se repite el codigo es mas rapido porque no procesa los que la cantidad recibida es mayor a cero
           'obtengo todas las tiendas que pidieron el producto
           rs.Open "SELECT df_pedido,df_prod, df_cantsol, df_cantreal FROM detallefactura,pedidos WHERE df_pedido = p_pedido AND P_PEDPROVEEDOR = '" & PedProve & "' AND DF_PROD = '" & AdoPedpro.Recordset!dg_producto & "' AND p_sucursal <> 10", cn, adOpenKeyset, adLockOptimistic, adCmdText
           nCanSug = 0
           nSolOfi = 0
           Set rstRepart = New ADODB.Recordset
           
           'Obtengo el total de la cantidad solicitada de los sugeridos ya que no coincide con lo solicitado del pedprove
           rstRepart.Open "SELECT sum(df_cantsol) as SolSug FROM detallefactura,pedidos WHERE df_pedido = p_pedido AND P_PEDPROVEEDOR = '" & PedProve & "' AND DF_PROD = '" & AdoPedpro.Recordset!dg_producto & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
           nCanSolSug = IIf(IsNull(rstRepart), 0, rstRepart!SolSug)
           rstRepart.Close
           
           'SI existe un sugerido de oficinas centrales se reparte entre las demas tiendas
           rstRepart.Open "SELECT * FROM detallefactura,pedidos WHERE df_pedido = p_pedido AND P_PEDPROVEEDOR = '" & PedProve & "' AND DF_PROD = '" & AdoPedpro.Recordset!dg_producto & "' AND P_SUCURSAL = 10", cn, adOpenKeyset, adLockOptimistic, adCmdText
           If rstRepart.RecordCount > 0 Then
              If rs.RecordCount > 0 Then nSolOfi = Round(rstRepart!df_cantsol / rs.RecordCount, 0)
           End If
           cProAnt = ""
           While Not rs.EOF
                stbMensajes.SimpleText = Space(25) & "Prorrateando producto del pedido sugerido " & rs!df_pedido
                stbMensajes.Refresh

                If cProAnt <> rs!df_prod Then
                    nCanSug = 0
                End If
                nCanRec = AdoPedpro.Recordset!dg_cantreal + AdoPedpro.Recordset!dg_promocionr
                'nPor = (rs!df_cantsol + nSolOfi) * 100 / AdoPedpro.Recordset!dg_cantsol
                nPor = (rs!df_cantsol + nSolOfi) * 100 / nCanSolSug
                ncan = Round(nCanRec * nPor / 100, 0)
                nCanSug = nCanSug + ncan
                cProAnt = rs!df_prod
                If nCanSug > nCanRec Then
                    rs!df_cantreal = nCanRec - (nCanSug - ncan)
                Else
                    rs!df_cantreal = ncan
                End If
                rs.Update
                rs.MoveNext
            Wend
            'En el caso que aun sobre producto por lo regular es uno
            If nCanSug < nCanRec And rs.RecordCount > 0 Then
               rs.MoveLast
               rs!df_cantreal = rs!df_cantreal + (nCanRec - nCanSug)
               rs.Update
            End If
            rs.Close
        End If
    End If
    AdoPedpro.Recordset.MoveNext
  Wend
End If
'Genero envios con los sugeridos y los marco como recibidos
rs.Open "SELECT * FROM pedidos WHERE p_pedproveedor = '" & PedProve & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
While Not rs.EOF
  If Trim(rs!p_sucursal) <> "14" And Trim(rs!p_sucursal) <> "15" And Trim(rs!p_sucursal) <> "12" And Trim(rs!p_sucursal) <> "5" And Trim(rs!p_sucursal) <> "4" And Trim(rs!p_sucursal) <> "10" Then
     stbMensajes.SimpleText = Space(25) & "Espere un momento, Generando envio del pedido sugerido " & rs!p_pedido
     stbMensajes.Refresh
    
     rstemp.Open "SELECT MAX (CAST(SUBSTRING(t_clave,4,10) AS INT)) As FolTra FROM Traslados WHERE SUBSTRING(t_clave,1,3) = 'T" & Trim(Mid(cSucursal, 3, 5)) & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
     CFOLTRA = IIf(IsNull(rstemp!FolTra), "T" + Mid(Trim(Mid(cSucursal, 3)), 1, 2) + "1", "T" + Mid(Trim(Mid(cSucursal, 3)), 1, 2) + Trim(Str(rstemp!FolTra + 1)))
     rstemp.Close
     rstemp.Open "SELECT * FROM Cattienda WHERE ticlave = '" & Trim(rs!p_sucursal) & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
     nFolTie = rstemp!FolioTie + 1
     'Agrego el traslado
     cn.Execute "INSERT INTO traslados(t_clave, t_fecha, t_tipo, t_sucursalemisor, t_sucursalreceptor, t_perfle, t_foliotie ) VALUES ('" & CFOLTRA & "','" & Date + Time & "',0,'" & Trim(Mid(cSucursal, 1, 3)) & "','" & rs!p_sucursal & "','3'," & nFolTie & ")"
     If Trim(txtCampos(1).Text = "JAR") Then
        cn.Execute "INSERT INTO DetalleTraslado (dt_clave,dt_producto,dt_cantidad,dt_cantidadp,dt_pedido) SELECT  dt_clave = '" & _
                      CFOLTRA & "',df_prod,df_cantreal,df_cantrealp, pedido = '" & rs!p_pedido & "' FROM detallefactura WHERE df_pedido = '" & rs!p_pedido & "'"
     Else  'En Jarcieria quien toma el control total para el surtimiento es el facturista asi es que se le pone lo solicitado
        cn.Execute "INSERT INTO DetalleTraslado (dt_clave,dt_producto,dt_cantidad,dt_cantidadp,dt_pedido) SELECT  dt_clave = '" & _
                      CFOLTRA & "',df_prod,df_cantsol,df_cantsolp, pedido = '" & rs!p_pedido & "' FROM detallefactura WHERE df_pedido = '" & rs!p_pedido & "'"
     End If
     'Actualizo el consecutivo de folios por tienda
     cn.Execute "UPDATE cattienda SET foliotie = foliotie + 1 WHERE ticlave = '" & Trim(rs!p_sucursal) & "'"
     'Cargo precios del traslado; Precio a franquicias (PRECIO4  de PREPROD)
     If Val(txtCampos(3).Text) = 5 Or Val(txtCampos(3).Text) = 12 Or Val(txtCampos(3).Text) = 13 Or Val(txtCampos(3).Text) = 15 Or Val(txtCampos(3).Text) = 14 Then
         cn.Execute "UPDATE DETALLETRASLADO SET DETALLETRASLADO.dt_costo = PREPROD.Precio4 FROM preprod WHERE PREPROD.preclave = DETALLETRASLADO.dt_producto AND DetalleTraslado.dt_clave = '" & CFOLTRA & "'"
     Else 'Precio a tiendas (Precosto de TFPRODUC)
         cn.Execute "UPDATE DETALLETRASLADO SET dt_costo = TFPRODUC.Precosto FROM TFPRODUC WHERE TFPRODUC.Consec = DETALLETRASLADO.dt_producto AND DetalleTraslado.dt_clave = '" & CFOLTRA & "'"
     End If
     rs!p_traslado = CFOLTRA
     MsgBox "SE GENERO EL TRASLADO PARA " & Trim(rstemp!TIDESCRIP) & " CON LA CLAVE " & CStr(CFOLTRA), vbInformation
     rstemp.Close
  End If
    rs!p_recibido = 1
    rs!p_fecentreal = txtRecib(0).Text
    rs.Update
    rs.MoveNext
Wend
End Sub
