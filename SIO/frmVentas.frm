VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVentas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ventas Mayoreo"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVentas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame fraCon 
      BackColor       =   &H00808000&
      Caption         =   "Contraseña de acceso"
      Height          =   2020
      Left            =   8040
      TabIndex        =   67
      Top             =   2880
      Visible         =   0   'False
      Width           =   3855
      Begin VB.CommandButton cmdConCance 
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   70
         Top             =   1440
         Width           =   1050
      End
      Begin VB.CommandButton cmdConAceptar 
         Caption         =   "&Aceptar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   69
         Top             =   1440
         Width           =   1050
      End
      Begin VB.TextBox txtContra 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1680
         PasswordChar    =   "*"
         TabIndex        =   68
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label lblContra 
         BackColor       =   &H00808000&
         Caption         =   "Contraseña"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   71
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.Frame frapreventa 
      Caption         =   "Clientes del folio de preventa"
      ForeColor       =   &H80000002&
      Height          =   5895
      Left            =   120
      TabIndex        =   61
      Top             =   2040
      Visible         =   0   'False
      Width           =   11775
      Begin MSDataGridLib.DataGrid DbgrdPreventa 
         Bindings        =   "frmVentas.frx":0442
         Height          =   4815
         Left            =   240
         TabIndex        =   62
         Top             =   240
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   8493
         _Version        =   393216
         AllowUpdate     =   -1  'True
         ForeColor       =   8388608
         HeadLines       =   1.5
         RowHeight       =   15
         RowDividerStyle =   0
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
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "noventa"
            Caption         =   "Fol. Uni."
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
            DataField       =   "CNOMBRE"
            Caption         =   "                  Cliente"
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
         BeginProperty Column02 
            DataField       =   "credito"
            Caption         =   "Crédito"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "SI"
               FalseValue      =   "NO"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "montototal"
            Caption         =   "Importe"
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
            DataField       =   "cl_terminal"
            Caption         =   "Equipo"
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
            DataField       =   "facrfc"
            Caption         =   "Factura"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "SI"
               FalseValue      =   "NO"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "Situacion"
            Caption         =   "Modo"
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
            MarqueeStyle    =   3
            RecordSelectors =   0   'False
            BeginProperty Column00 
               Alignment       =   2
               DividerStyle    =   3
               Locked          =   -1  'True
               ColumnWidth     =   884.976
            EndProperty
            BeginProperty Column01 
               DividerStyle    =   3
               Locked          =   -1  'True
               ColumnWidth     =   5444.788
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   645.165
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               DividerStyle    =   3
               Locked          =   -1  'True
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               DividerStyle    =   3
               Locked          =   -1  'True
               ColumnWidth     =   1349.858
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               DividerStyle    =   3
               ColumnWidth     =   780.095
            EndProperty
            BeginProperty Column06 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   629.858
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdExporta 
         BackColor       =   &H00808080&
         Caption         =   "&Preparar Información"
         Height          =   255
         Left            =   9120
         Style           =   1  'Graphical
         TabIndex        =   101
         Top             =   5100
         Width           =   2415
      End
      Begin VB.CommandButton cmdDevPvta 
         Caption         =   "&Dev.pvta."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         Picture         =   "frmVentas.frx":045C
         Style           =   1  'Graphical
         TabIndex        =   100
         ToolTipText     =   "Genera reporte devolucion de preventa"
         Top             =   5280
         Width           =   800
      End
      Begin VB.CheckBox chkLiquida 
         Caption         =   "&Liquidada"
         Height          =   195
         Left            =   120
         TabIndex        =   99
         Top             =   5640
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CheckBox chkPrevLiq 
         Caption         =   "Carga&da"
         Height          =   255
         Left            =   120
         TabIndex        =   98
         Top             =   5400
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CommandButton cmdRegVta 
         Caption         =   "&Regresar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8040
         Picture         =   "frmVentas.frx":098E
         Style           =   1  'Graphical
         TabIndex        =   80
         ToolTipText     =   "Regresar a la pantalla de ventas"
         Top             =   5280
         Width           =   800
      End
      Begin VB.CommandButton cmdAgrVta 
         Caption         =   "&Agr. Vta."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7200
         Picture         =   "frmVentas.frx":0B00
         Style           =   1  'Graphical
         TabIndex        =   79
         ToolTipText     =   "Agregar una venta al agente"
         Top             =   5280
         Width           =   800
      End
      Begin VB.CommandButton cmdReporte 
         Caption         =   "&Rpt pvta."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         Picture         =   "frmVentas.frx":0BFA
         Style           =   1  'Graphical
         TabIndex        =   78
         ToolTipText     =   "Genera reporte concentrado / detallado"
         Top             =   5280
         Width           =   800
      End
      Begin VB.CommandButton cmdFacCte 
         Caption         =   "Fac.C&te."
         Height          =   495
         Left            =   4560
         Picture         =   "frmVentas.frx":112C
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   5280
         Width           =   800
      End
      Begin VB.CommandButton cmdFacConFin 
         Caption         =   "Fac.&Con"
         Height          =   495
         Left            =   5400
         Picture         =   "frmVentas.frx":122E
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   5280
         Width           =   915
      End
      Begin VB.CommandButton cmdModificar 
         Caption         =   "&Modif."
         Height          =   495
         Left            =   6360
         Picture         =   "frmVentas.frx":1330
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   5280
         Width           =   800
      End
      Begin VB.CommandButton cmdrptvta 
         Caption         =   "T&ickets"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3720
         Picture         =   "frmVentas.frx":147A
         Style           =   1  'Graphical
         TabIndex        =   74
         ToolTipText     =   "Genera reporte  del folio único de venta proporcionado."
         Top             =   5280
         Width           =   800
      End
      Begin VB.TextBox TxtTotVta 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   10080
         Locked          =   -1  'True
         TabIndex        =   73
         Text            =   "0.00"
         Top             =   5400
         Width           =   1455
      End
      Begin VB.CheckBox chkConcent 
         Caption         =   "&Concentrado"
         Height          =   255
         Left            =   120
         TabIndex        =   63
         Top             =   5170
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.PictureBox cRpt 
         Height          =   480
         Left            =   360
         ScaleHeight     =   420
         ScaleWidth      =   1140
         TabIndex        =   102
         Top             =   3360
         Width           =   1200
      End
      Begin VB.Label lbletiquetas 
         Caption         =   "Total Venta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   15
         Left            =   9120
         TabIndex        =   72
         Top             =   5400
         Width           =   855
      End
   End
   Begin VB.Frame fraAvance 
      BackColor       =   &H80000018&
      Caption         =   "Obteniendo especificaciones de productos "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3360
      TabIndex        =   43
      Top             =   4080
      Visible         =   0   'False
      Width           =   4935
      Begin MSComctlLib.ProgressBar pgb 
         Height          =   285
         Left            =   120
         TabIndex        =   47
         Top             =   360
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   1
         Enabled         =   0   'False
      End
   End
   Begin MSAdodcLib.Adodc adopreventa 
      Height          =   330
      Left            =   4680
      Top             =   0
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
      Caption         =   "AdoPreventa"
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
   Begin MSAdodcLib.Adodc AdoAgtedet 
      Height          =   330
      Left            =   2280
      Top             =   0
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
      Caption         =   "AdoAgtedet"
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
   Begin MSAdodcLib.Adodc AdoAgentes 
      Height          =   330
      Left            =   2400
      Top             =   0
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
      Caption         =   "AdoAgentes"
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
   Begin MSAdodcLib.Adodc AdoDetVta 
      Height          =   330
      Left            =   6360
      Top             =   0
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   1
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
      Caption         =   "AdoDetVta"
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
   Begin MSAdodcLib.Adodc AdoVentas 
      Height          =   330
      Left            =   8640
      Top             =   0
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
      Caption         =   "Adoventas"
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
   Begin ComctlLib.StatusBar Stb1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   24
      Top             =   8220
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   9
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   5644
            MinWidth        =   5644
            Text            =   "[Esc]  para salir"
            TextSave        =   "[Esc]  para salir"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   2207
            MinWidth        =   2207
            Text            =   "F2=Mod. precio"
            TextSave        =   "F2=Mod. precio"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Modificar precio (Mas altos o mas Bajos)"
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   2205
            MinWidth        =   2205
            Text            =   "F3=Canc. prod."
            TextSave        =   "F3=Canc. prod."
            Object.Tag             =   ""
            Object.ToolTipText     =   "Cancelar producto del pedido"
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   1588
            MinWidth        =   1588
            Text            =   "F4=Cobrar"
            TextSave        =   "F4=Cobrar"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Cobrar la venta"
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   1695
            MinWidth        =   1695
            Text            =   "F5=Pedidos"
            TextSave        =   "F5=Pedidos"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Visualizar pedidos de la Preventa"
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   2083
            MinWidth        =   2083
            Text            =   "F6=Precio Alto"
            TextSave        =   "F6=Precio Alto"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Modifica precio de venta solamente a mas alto"
         EndProperty
         BeginProperty Panel7 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   2205
            MinWidth        =   2205
            Text            =   "F7=Mas Piezas"
            TextSave        =   "F7=Mas Piezas"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Permite capturar mas piezas de las que trae la caja"
         EndProperty
         BeginProperty Panel8 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "F9=Obt. Inv."
            TextSave        =   "F9=Obt. Inv."
            Object.Tag             =   ""
            Object.ToolTipText     =   "Obtiene Inventario de bodega"
         EndProperty
         BeginProperty Panel9 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   "F10=Imp. Pvta"
            TextSave        =   "F10=Imp. Pvta"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Importa Pedidos de Agentes"
         EndProperty
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
   End
   Begin VB.Frame fraGenerales 
      Caption         =   "Datos generales de la venta"
      ForeColor       =   &H80000002&
      Height          =   1815
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   11775
      Begin VB.CheckBox chkPrev 
         Caption         =   "&Pre&venta"
         DataField       =   "prevta"
         DataSource      =   "AdoVentas"
         Height          =   195
         Left            =   10320
         TabIndex        =   96
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdCamSer 
         Caption         =   "&Escala"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   7210
         Picture         =   "frmVentas.frx":15EC
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Cambia escala de precios a todos los productos de la venta"
         Top             =   1200
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton cmdRegresar 
         Cancel          =   -1  'True
         Caption         =   "&Cerrar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   9750
         Picture         =   "frmVentas.frx":172E
         Style           =   1  'Graphical
         TabIndex        =   60
         ToolTipText     =   "Regresar a la pantalla anterior"
         Top             =   1200
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "&Grabar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   9120
         Picture         =   "frmVentas.frx":18A0
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Grabar datos de la venta"
         Top             =   1200
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.ComboBox cmbChofer 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7320
         TabIndex        =   6
         Top             =   840
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.TextBox txtCte 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   7
         Top             =   1080
         Visible         =   0   'False
         Width           =   4935
      End
      Begin VB.TextBox txtcampos 
         Alignment       =   2  'Center
         DataField       =   "folpreventa"
         DataSource      =   "AdoVentas"
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   10560
         TabIndex        =   59
         Top             =   1440
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkConRfc 
         Caption         =   "&Factura"
         DataField       =   "facRfc"
         DataSource      =   "AdoVentas"
         Height          =   195
         Left            =   10320
         TabIndex        =   55
         Top             =   600
         Width           =   1095
      End
      Begin VB.ComboBox cmbAgente 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   700
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.TextBox txtcampos 
         Alignment       =   2  'Center
         DataField       =   "agente"
         DataSource      =   "AdoVentas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   9
         Left            =   720
         TabIndex        =   3
         Top             =   700
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CheckBox chkCredito 
         Caption         =   "C&rédito"
         DataField       =   "credito"
         DataSource      =   "AdoVentas"
         Height          =   255
         Left            =   10320
         TabIndex        =   9
         Top             =   120
         Width           =   1095
      End
      Begin VB.TextBox txtFolio 
         Alignment       =   2  'Center
         DataField       =   "FolioVenta"
         DataSource      =   "AdoVentas"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   0
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtcampos 
         Alignment       =   2  'Center
         DataField       =   "clcliente"
         DataSource      =   "AdoVentas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Index           =   4
         Left            =   720
         TabIndex        =   8
         Top             =   1450
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtcampos 
         Alignment       =   2  'Center
         DataField       =   "tipoventa"
         DataSource      =   "AdoVentas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   6480
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
      Begin VB.ComboBox cmbCliente 
         DataField       =   "cNombre"
         DataSource      =   "AdoCliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   1450
         Visible         =   0   'False
         Width           =   4935
      End
      Begin VB.TextBox txtcampos 
         Alignment       =   2  'Center
         DataField       =   "chofer"
         DataSource      =   "AdoVentas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Index           =   3
         Left            =   6480
         MaxLength       =   15
         TabIndex        =   5
         Top             =   840
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox cmbTipVta 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7320
         Sorted          =   -1  'True
         TabIndex        =   11
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox txtcampos 
         Alignment       =   2  'Center
         DataField       =   "fecha"
         DataSource      =   "AdoVentas"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Index           =   0
         Left            =   3360
         TabIndex        =   1
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdClientes 
         Caption         =   "Cli&entes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   8490
         Picture         =   "frmVentas.frx":199A
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Consultar catalogo de clientes"
         Top             =   1200
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton cmdTicBod 
         Caption         =   "ticket"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   7860
         Picture         =   "frmVentas.frx":1A9C
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Imprime ticket para surtido en bodega"
         Top             =   1200
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton cmdticket 
         Caption         =   "&Factura"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   6600
         Picture         =   "frmVentas.frx":1BE6
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Imprimir ticket de venta"
         Top             =   1200
         Visible         =   0   'False
         Width           =   650
      End
      Begin MSComDlg.CommonDialog cmdg 
         Left            =   0
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label lblCte 
         Caption         =   "Buscar Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   66
         Top             =   1080
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lbletiquetas 
         Alignment       =   2  'Center
         Caption         =   "Fol. Preventa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   10440
         TabIndex        =   58
         Top             =   1200
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label LblFolio 
         Caption         =   "Numero de venta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lbletiquetas 
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   2880
         TabIndex        =   23
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lbletiquetas 
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   22
         Top             =   1450
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lbletiquetas 
         Caption         =   "Chofer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   5880
         TabIndex        =   21
         Top             =   720
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lbletiquetas 
         Caption         =   "Agente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   20
         Top             =   700
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Label lbletiquetas 
         Caption         =   "Tipo de venta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   5400
         TabIndex        =   18
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSDataGridLib.DataGrid DbgDetVta 
      Bindings        =   "frmVentas.frx":1D58
      Height          =   3135
      Left            =   120
      TabIndex        =   42
      Top             =   5040
      Visible         =   0   'False
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   5530
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   14737632
      BorderStyle     =   0
      HeadLines       =   1.5
      RowHeight       =   19
      RowDividerStyle =   0
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
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "cl_producto"
         Caption         =   "clave"
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
         DataField       =   "cantidad"
         Caption         =   "Cajas"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "cantidadp"
         Caption         =   "Piezas"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "descripc"
         Caption         =   "                                        Descripcion"
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
         DataField       =   "MEDIDA"
         Caption         =   "Medida"
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
         DataField       =   "precio"
         Caption         =   "Pre. Unit."
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
         DataField       =   "importe"
         Caption         =   "    Importe"
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
         DataField       =   "factura"
         Caption         =   "Factura"
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
         MarqueeStyle    =   4
         BeginProperty Column00 
            DividerStyle    =   3
            Object.Visible         =   0   'False
            ColumnWidth     =   1514.835
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            DividerStyle    =   3
            ColumnWidth     =   675.213
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            DividerStyle    =   3
            ColumnWidth     =   689.953
         EndProperty
         BeginProperty Column03 
            DividerStyle    =   3
            ColumnWidth     =   5880.189
         EndProperty
         BeginProperty Column04 
            DividerStyle    =   3
            ColumnWidth     =   1695.118
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            DividerStyle    =   3
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            DividerStyle    =   3
            ColumnWidth     =   1379.906
         EndProperty
         BeginProperty Column07 
            Alignment       =   2
            ColumnWidth     =   1110.047
         EndProperty
      EndProperty
   End
   Begin VB.Frame FraTotal 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   8040
      TabIndex        =   44
      Top             =   2880
      Visible         =   0   'False
      Width           =   3855
      Begin VB.Label lblpromo 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PROMOCION"
         Height          =   255
         Left            =   120
         TabIndex        =   97
         Top             =   240
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.Label lblImpVta 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Haettenschweiler"
            Size            =   45.75
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   120
         TabIndex        =   54
         Top             =   480
         Width           =   3705
      End
   End
   Begin VB.Frame FraVtadet 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2055
      Left            =   120
      TabIndex        =   30
      Top             =   2880
      Visible         =   0   'False
      Width           =   7815
      Begin VB.TextBox txtcampos 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Index           =   22
         Left            =   4560
         TabIndex        =   93
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtcampos 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Index           =   21
         Left            =   3480
         TabIndex        =   92
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtcampos 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Index           =   20
         Left            =   2400
         TabIndex        =   91
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtcampos 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Index           =   19
         Left            =   5640
         TabIndex        =   90
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtcampos 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Index           =   18
         Left            =   1320
         TabIndex        =   88
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox txtcampos 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Index           =   17
         Left            =   3480
         TabIndex        =   86
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox txtcampos 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Index           =   16
         Left            =   2400
         TabIndex        =   84
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox txtcampos 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Index           =   15
         Left            =   5640
         TabIndex        =   82
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox txtcampos 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Index           =   14
         Left            =   5280
         TabIndex        =   34
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtcampos 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Index           =   11
         Left            =   6720
         TabIndex        =   65
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtcampos 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Index           =   10
         Left            =   1320
         TabIndex        =   56
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtcampos 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Index           =   8
         Left            =   4560
         TabIndex        =   37
         Top             =   1320
         Width           =   975
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "&Agregar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6840
         Picture         =   "frmVentas.frx":1D70
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Agregra producto a la venta"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtcampos 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Index           =   13
         Left            =   6000
         TabIndex        =   35
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtcampos 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   12
         Left            =   3720
         TabIndex        =   46
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtcampos 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Index           =   7
         Left            =   6720
         TabIndex        =   36
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox txtcampos 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Index           =   6
         Left            =   4560
         TabIndex        =   33
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtcampos 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Index           =   5
         Left            =   3000
         TabIndex        =   32
         Top             =   480
         Width           =   735
      End
      Begin VB.ComboBox cmbprod 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmVentas.frx":1E6A
         Left            =   120
         List            =   "frmVentas.frx":1E6C
         Sorted          =   -1  'True
         TabIndex        =   31
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label lbletiquetas 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   21
         Left            =   120
         TabIndex        =   95
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label lbletiquetas 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "P r e c i o s"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   20
         Left            =   120
         TabIndex        =   94
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lbletiquetas 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Pieza Auto."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   250
         Index           =   19
         Left            =   1320
         TabIndex        =   89
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lbletiquetas 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1/2 Caj. Env."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   250
         Index           =   18
         Left            =   2400
         TabIndex        =   87
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lbletiquetas 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1/2 Caj. Bod."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   250
         Index           =   17
         Left            =   3480
         TabIndex        =   85
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lbletiquetas 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Caja  Interm."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   250
         Index           =   5
         Left            =   5640
         TabIndex        =   83
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lbletiquetas 
         Alignment       =   2  'Center
         Caption         =   "1/2 Caja"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   16
         Left            =   5280
         TabIndex        =   81
         Top             =   165
         Width           =   735
      End
      Begin VB.Label lbletiquetas 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Precio Pieza"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   64
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label lbletiquetas 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Precio  Caja"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   57
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label lbletiquetas 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Caja  Envío"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   250
         Index           =   9
         Left            =   4560
         TabIndex        =   49
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lbletiquetas 
         Alignment       =   2  'Center
         Caption         =   "Piezas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   6000
         TabIndex        =   48
         Top             =   165
         Width           =   735
      End
      Begin VB.Label lbletiquetas 
         Alignment       =   2  'Center
         Caption         =   "Exi.Pza."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   3720
         TabIndex        =   45
         Top             =   165
         Width           =   735
      End
      Begin VB.Label lbletiquetas 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Caja  Bod."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   250
         Index           =   8
         Left            =   6720
         TabIndex        =   41
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lbletiquetas 
         Alignment       =   2  'Center
         Caption         =   "Cajas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   4560
         TabIndex        =   40
         Top             =   165
         Width           =   735
      End
      Begin VB.Label lbletiquetas 
         Alignment       =   2  'Center
         Caption         =   "Exi.Caj."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   3000
         TabIndex        =   39
         Top             =   165
         Width           =   735
      End
   End
   Begin VB.Frame FraCliente 
      Caption         =   "Datos del cliente"
      ForeColor       =   &H80000002&
      Height          =   855
      Left            =   120
      TabIndex        =   25
      Top             =   2040
      Visible         =   0   'False
      Width           =   11775
      Begin VB.Label lblcliente 
         Caption         =   "TEL.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   7920
         TabIndex        =   53
         Top             =   480
         Width           =   615
      End
      Begin VB.Label lblcliente 
         Caption         =   "R.F.C.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   7920
         TabIndex        =   52
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblcliente 
         Caption         =   "DIRECCION  :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   51
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label lblcliente 
         Caption         =   "NOM FISCAL:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   50
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblTelefono 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   8640
         TabIndex        =   29
         Top             =   480
         UseMnemonic     =   0   'False
         Width           =   2895
      End
      Begin VB.Label lblRfc 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   8640
         TabIndex        =   28
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   2895
      End
      Begin VB.Label lblDirec 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1560
         TabIndex        =   27
         Top             =   480
         UseMnemonic     =   0   'False
         Width           =   6015
      End
      Begin VB.Label lblNombre 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1560
         TabIndex        =   26
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   6015
      End
   End
End
Attribute VB_Name = "frmVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsttemp As ADODB.Recordset
Private rscli As New ADODB.Recordset    'Clientes
Private nImpVta As Double
Private TipVta As Variant
Private rs As ADODB.Recordset           'Productos en inventario
Private cprod As String                 'Clave del producto seleccionado
Private lVta As Boolean                 'Exclusivamente venta (Modulo de venta)
Private lCob As Boolean                 'Modificar y cobrar (Caja )
Private nDispo As Currency    'Importe disponible para ventas a credito
Dim nAncho As Integer
Dim cModifico As String
Private CarCte As Boolean
Private nValAnt As Currency
Private nOpcion As Integer
Private globconFin As Boolean
Private pzasmay As Integer
Private lMasPza As Boolean
Private cBorrar As String
Dim Unidades$(9), Decenas$(9), Oncenas$(9)
Dim Veintes$(9), centenas$(9)

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
    NumLet$ = "(  " & Trim$(AUX$) & " PESOS " & dec$ & "/ 100 M.N.  )"
End Function
    
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

Private Sub adopreventa_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error Resume Next
   cmdFacCte.Enabled = adopreventa.Recordset!facrfc And adopreventa.Recordset!situacion = "1" And Not ModVta
End Sub


Private Sub chkCredito_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   SendKeys vbTab
ElseIf KeyAscii = 27 Then
   Unload Me
End If

End Sub

Private Sub chkLiquida_Click()
If chkLiquida.Value = 1 And chkLiquida.Enabled Then
   If MsgBox("Desea Confirmar la preventa como liquidada?", vbYesNo + vbQuestion, "Status de la preventa") = vbYes Then
      cn.Execute "UPDATE preventas SET impliq = " & Format(TxtTotVta.Text, "#########.##") & " WHERE folio = " & Trim(txtcampos(2).Text)
      chkLiquida.Enabled = False
      chkLiquida.Visible = True
      crpt.WindowTitle = "HOJA DE DEVOLUCION DE LA PREVENTA " & txtcampos(2).Text
      crpt.ReportFileName = App.Path & "\PrevLiq.rpt"
      crpt.Connect = cCadConex
      crpt.Formulas(0) = "ENCABEZADO = 'HOJA DE DEVOLUCION DE LA PREVENTA " & Trim(txtcampos(2).Text) & "'"
      cadsql = "SELECT TFPRODUC.DESCRIPC, TFPRODUC.CONTENID, TFPRODUC.MEDIDA, TFPRODUC.PAQUETES, " & _
                       "PRODBORRA.clave, PRODBORRA.cajas, PRODBORRA.piezas, PRODBORRA.importe, PRODBORRA.prevliq, PRODBORRA.preventa, " & _
                       "PREVENTAS.folio, PREVENTAS.impcapt, PREVENTAS.impsurt, PREVENTAS.impliq " & _
               "FROM pitico.dbo.TFPRODUC TFPRODUC, " & _
                       "PITICO.dbo.PRODBORRA PRODBORRA, " & _
                       "PITICO.dbo.PREVENTAS PREVENTAS " & _
               "WHERE TFPRODUC.CONSEC = PRODBORRA.clave AND " & _
                      "PRODBORRA.preventa = PREVENTAS.folio AND " & _
                      "PRODBORRA.prevliq = 1 AND " & _
                      "PRODBORRA.preventa = " & txtcampos(2).Text & ". " & _
               "ORDER BY TFPRODUC.DESCRIPC ASC "
      crpt.SQLQuery = cadsql
      crpt.Action = 1
   Else
      chkLiquida.Value = 0
   End If
End If
End Sub

Private Sub chkPrevliq_Click()
If chkPrevLiq.Value = 1 Then
If MsgBox("Desea Confirmar la preventa como previa a liquidar?", vbYesNo + vbQuestion, "Status de la preventa") = vbYes Then
   cn.Execute "UPDATE preventas SET impsurt = " & Format(TxtTotVta.Text, "#########.##") & " WHERE folio = " & Trim(txtcampos(2).Text)
   chkPrevLiq.Enabled = False
   chkLiquida.Visible = True
   chkLiquida.Value = 0
   lImp = False
   For Each x In Printers
      If x.DeviceName Like "*PREVENTA*" Then
         lImp = True
         Set Printer = x
         Exit For
      End If
   Next x
   If lImp = False Then Exit Sub
  ncopias = InputBox("Número de copias imprimir", "Copias", 2)
  If Not IsNumeric(ncopias) Then Exit Sub
  Printer.ScaleMode = 7
  For x = 1 To ncopias
  If ZONA = "OAX" Then
    Printer.Print "     VIVERES Y LICORES S.A DE C.V.    "
  Else
    Printer.Print "HOLDING MEXICO CENTRO AMERICA SA DE CV"
  End If
   Printer.Print "PEDIDOS DE LA PREVENTA " & txtcampos(2).Text & "  " & date
   Printer.Print "AGENTE: " & cmbAgente.Text
   Printer.Print "CHOFER: " & cmbChofer.Text
   Printer.Print "----------------------------------------------------------------------"
   adopreventa.Refresh: N = 0: TotalPvta = 0
   While Not adopreventa.Recordset.EOF
       If adopreventa.Recordset!montototal > 0 Then
          N = N + 1
          If Len(Trim(adopreventa.Recordset!cNombre)) > 25 Then
             Printer.Print N & ")" & Mid(Trim(adopreventa.Recordset!cNombre), 1, 20)
             Printer.Print "     " & Mid(Trim(adopreventa.Recordset!cNombre), 21);
          Else
             Printer.Print N & ")" & Trim(adopreventa.Recordset!cNombre);
          End If
          p = 6.2 - (Printer.TextWidth(Format(adopreventa.Recordset!montototal, "###,###,##0.00")))
          Printer.CurrentX = p
          Printer.Print Format(adopreventa.Recordset!montototal, "$###,####,##0.00")
          TotalPvta = TotalPvta + adopreventa.Recordset!montototal
       End If
       adopreventa.Recordset.MoveNext
   Wend
   Printer.Print "----------------------------------------------------------------------"
   Printer.Print "NUMERO DE PEDIDOS: " & N
   Printer.Print "IMPORTE TOTAL: " & Format(TotalPvta, "$###,###,##0.00")
   For N = 1 To 18
      Printer.Print " "
   Next
   Printer.EndDoc
  Next
Else
   chkPrevLiq.Value = 0
End If
End If
End Sub

Private Sub cmbagente_GotFocus()
 RESP = SendMessageLong(cmbAgente.hwnd, &H14F, True, 1)
End Sub

Private Sub cmbAgente_LostFocus()
 RESP = SendMessageLong(cmbAgente.hwnd, &H14F, False, 1)
End Sub

Private Sub cmbAgente_Validate(Cancel As Boolean)
Dim rsclie As ADODB.Recordset
On Error GoTo Error:
If Trim(cmbAgente.Text) <> "" Then
   Set rsclie = New ADODB.Recordset
   ccVeCli = cmbAgente.Text
   rsclie.Open "SELECT * FROM Catcliente WHERE cnombre = '" & ccVeCli & "' AND cTipo = 1", cn, adOpenKeyset, adLockOptimistic, adCmdText
   If rsclie.RecordCount = 0 Then
      cmbAgente.SetFocus
      Exit Sub
   End If
   txtcampos(9).Text = rsclie!cclave
   txtcampos(9).SetFocus
Else
   txtcampos(9).Text = ""
End If
Exit Sub
Error:
   MsgBox "OCURRIO UN ERROR: " & Err.Description, vbExclamation
End Sub

Private Sub cmbChofer_GotFocus()
 RESP = SendMessageLong(cmbChofer.hwnd, &H14F, True, 1)
End Sub

Private Sub cmbChofer_LostFocus()
 RESP = SendMessageLong(cmbChofer.hwnd, &H14F, False, 1)
End Sub

Private Sub cmbChofer_Validate(Cancel As Boolean)
Dim rsclie As ADODB.Recordset
On Error GoTo Error:
If Trim(cmbAgente.Text) <> "" Then
   Set rsclie = New ADODB.Recordset
   ccVeCli = cmbChofer.Text
   rsclie.Open "SELECT * FROM Catcliente WHERE cnombre = '" & ccVeCli & "' AND cTipo = 2", cn, adOpenKeyset, adLockOptimistic, adCmdText
   If rsclie.RecordCount = 0 Then
      cmbChofer.SetFocus
      Exit Sub
   End If
   txtcampos(3).Text = rsclie!cclave
   txtcampos(3).SetFocus
Else
   txtcampos(3).Text = ""
End If
Exit Sub
Error:
   MsgBox "OCURRIO UN ERROR: " & Err.Description, vbExclamation
End Sub

Private Sub cmbCliente_GotFocus()
 RESP = SendMessageLong(cmbCliente.hwnd, &H14F, True, 1)
End Sub

Private Sub cmbCliente_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   SendKeys vbTab
ElseIf KeyAscii = 27 Then
   Unload Me
End If
End Sub

Private Sub cmbCliente_LostFocus()
RESP = SendMessageLong(cmbCliente.hwnd, &H14F, False, 1)
End Sub

Private Sub cmbCliente_Validate(Cancel As Boolean)
Dim rsclie As ADODB.Recordset
On Error GoTo Error:
If Trim(cmbCliente.Text) <> "" Then
   Set rsclie = New ADODB.Recordset
   ccVeCli = cmbCliente.Text
   rsclie.Open "SELECT * FROM Catcliente WHERE cnombre = '" & ccVeCli & "' AND ctipo = 0", cn, adOpenKeyset, adLockOptimistic, adCmdText
   If rsclie.RecordCount = 0 Then
      cmbCliente.SetFocus
      Exit Sub
   End If
   txtcampos(4).Text = rsclie!cclave
   txtcampos(4).SetFocus
Else
   txtcampos(4).Text = ""
End If
Exit Sub
Error:
   MsgBox "OCURRIO UN ERROR: " & Err.Description, vbExclamation
End Sub

Private Sub cmbprod_GotFocus()
'Para que se despliegue automaticamente el combo
RESP = SendMessageLong(cmbprod.hwnd, &H14F, True, 1)
End Sub

Private Sub cmbprod_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 27 Then
   KeyAscii = 0
   SendKeys vbTab
'ElseIf KeyAscii = 27 Then
'   sendkey
End If
End Sub

Private Sub cmbprod_LostFocus()
  RESP = SendMessageLong(cmbprod.hwnd, &H14F, False, 1)
End Sub

Private Sub cmbTipVta_GotFocus()
   RESP = SendMessageLong(cmbTipVta.hwnd, &H14F, True, 1)
End Sub

Private Sub cmbTipVta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   SendKeys vbTab
ElseIf KeyAscii = 27 Then
   Unload Me
End If
End Sub

Private Sub cmbTipVta_LostFocus()
   RESP = SendMessageLong(cmbTipVta.hwnd, &H14F, False, 1)
End Sub

Private Sub cmdAgregar_Click()
Dim rsttemp As ADODB.Recordset
Dim iva As Double
Dim ieps As Double
Dim lTrans As Boolean
Dim dfecha As Date
On Error GoTo Error:
   If (Not IsNumeric(txtcampos(6).Text) Or Not IsNumeric(txtcampos(13).Text) Or Not IsNumeric(txtcampos(14).Text)) Or txtcampos(6).Text = "0" And txtcampos(13).Text = "0" And txtcampos(14).Text = "0" Then
       MsgBox "LA CANTIDAD ESPECIFICADA DEBE SER MAYOR A CERO", vbExclamation
       txtcampos(6).SetFocus
       Exit Sub
   ElseIf AdoVentas.Recordset!situacion = "1" And cBorrar = "" Then
       nOpcion = 4
       fraCon.Visible = True
       txtContra.Text = ""
       txtContra.SetFocus
       Exit Sub
   ElseIf chkLiquida.Value = 1 And chkCredito = 1 And Sql Then
      MsgBox "YA NO ES POSIBLE CAPTURAR PRODUCTO PORQUE LA PREVENTA YA FUE LIQUIDADA", vbInformation, "Preventa liquidada"
      Exit Sub
   Else
       'Se permite un producto de menos o mas a la media caja por aquellas cajas que no son pares.
       If Val(txtcampos(14).Text) > 0 Then
          If Abs(pzasmay - Val(txtcampos(14).Text)) > 1 Then
             MsgBox "EN VENTAS DE MEDIO MAYOREO SOLO SE PERMITEN MEDIAS CAJAS" & Chr(13) & "AJUSTE LA CANTIDAD DE MEDIO MAYOREO A " & pzasmay, vbInformation
             txtcampos(14).SetFocus
             Exit Sub
          End If
       End If
   End If
   If txtcampos(6).Text <= 0 And txtcampos(13).Text + txtcampos(14).Text <= 0 Then
      MsgBox "LA CANTIDAD ESPECIFICADA DEBE SER MAYOR A CERO", vbInformation, "Mayor a cero"
      Exit Sub
   End If
   NPREVTA = IIf(chkCredito.Value = 0 And chkPrev.Value = 0, txtcampos(7).Text, txtcampos(8).Text)
   nprevtapm = IIf(chkCredito.Value = 0 And chkPrev.Value = 0, txtcampos(21).Text, txtcampos(20).Text)
   NPRECIOP = txtcampos(10).Text
   If nDispo > 0 Then
      If Val(Format(lblImpVta.Caption, "#########.00")) + (NPREVTA * txtcampos(6).Text) + (txtcampos(13).Text * (txtcampos(10).Text)) + (nprevtapm * txtcampos(14)) > nDispo Then
         rscli.Requery
         MsgBox "NO ES POSIBLE CAPTURAR ESTE PRODUCTO PORQUE LA VENTA" & Chr(13) & "ASCENDERIA A " & Format(Val(Format(lblImpVta.Caption, "#########.00")) + (NPREVTA * txtcampos(6).Text) + (txtcampos(13).Text * (txtcampos(10).Text)) + (nprevtapm * txtcampos(14)), "$ ###,###,##0.00") & " Y ESTARIA EXCEDIENDO EL DISPONIBLE" & _
                 Chr(13) & Chr(13) & "MONTO AUTORIZADO:             " & Format(rscli!CLIMITECREDITO, "$ ###,###,##0.00") & Chr(13) & "EJERCIDO :     " & Space(25) & Format(rscli!CLIMITECREDITO - nDispo, "$ ###,###,##0.00") & Chr(13) & Chr(13) & "DISPONIBLE :" & Space(26) & Format(nDispo, "$ ###,###,#00.00"), vbExclamation, "Ventas"
         Exit Sub
      End If
   End If
   
   If Not Sql Then
       ESCALA = InputBox("Proporcione escala de precios a uilizar" & Chr(13) & Chr(13) & "2 = Precio de Envío" & Chr(13) & "3 = Precio Intermedio" & Chr(13) & "4 = Precio de bodega", "Teclee escala", 2)
       If Not IsNumeric(ESCALA) Or ESCALA > 4 Or ESCALA < 2 Then
          MsgBox "ESCALA INCORRECTA", vbInformation
          Exit Sub
       End If
    End If
    'Se checa la existencia del producto seleccionado
    Set rsttemp = New ADODB.Recordset
    rsttemp.Open "SELECT * FROM TFPRODUC,INVENTARIO WHERE CONSEC = INPROD AND (InCant < " & Val(txtcampos(6).Text) & " OR (InCantPza < " & Val(txtcampos(13).Text) + Val(txtcampos(14).Text) & " AND Incant <= " & Val(txtcampos(6).Text) & " )) AND CONSEC ='" & cprod & "' ORDER BY DESCRIPC,CONTENID", cn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not (rsttemp.BOF And rsttemp.EOF) And Sql Then
       MsgBox "EL INVENTARIO ES MENOR A LO SOLICITADO EN EL SIGUIENTE ARTICULO:" & Chr(13) _
           & rsttemp!descripc & " " & CStr(rsttemp!PAQUETES) & " X " & CStr(rsttemp!CONTENID) & " " & rsttemp!medida, vbCritical
       txtcampos(6).SetFocus
       Exit Sub
    End If
    If Sql Then cn.BeginTrans
    CmdAgregar.Enabled = False

    'Se valida que no capturen dos veces el mismo producto
    rsttemp.Close
    rsttemp.Open "SELECT * FROM VENTAS_DET WHERE NOVENTA = " & AdoVentas.Recordset!noventa & " AND cl_producto = '" & rs!CONSEC & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Not (rsttemp.BOF And rsttemp.EOF) Then
       MsgBox "YA NO ES POSIBLE CAPTURAR ESTE PRODUCTO" & Chr(13) & "PORQUE YA SE CAPTURO EN LA VENTA", vbInformation
       If Sql Then cn.CommitTrans
       cmbprod.SetFocus
       Exit Sub
    End If
    rsttemp.Close
    'SELECCION DEL PRECIO DE ACUERDO AL TIPO DE VENTA, CREDITO O CONTADO
    nprecio = IIf(chkCredito.Value = 0 And Me.chkPrev.Value = 0, txtcampos(7).Text, txtcampos(8).Text)
    'Se verifican las piezas si es mayor la Cant. solicitada o se convierte una caja en piezas
    If rs!InCantPza < Val(txtcampos(13).Text) Then
       'Agentes especiales y Agentes
       'If txtcampos(1).Text <> "5" Or txtcampos(1).Text <> "2" Then
       '   MsgBox "LA CANTIDAD SOLICITADA EN PIEZAS ES MAYOR AL INVENTARIO EN PIEZAS", vbInformation
       '   txtcampos(13).Text = 0
       '   txtcampos(13).SetFocus
       '   If Sql Then cn.CommitTrans
       '   Exit Sub
       'Else
          If MsgBox("LA EXISTENCIA EN PIEZAS ES MENOR A LO SOLICITADO, DESEAS CONVERTIR UNA CAJA EN PIEZAS", vbQuestion + vbYesNo) = vbYes Then
             If Sql Then
                cn.Execute "UPDATE inventario SET incantpza = incantpza + paquetes, incant = incant - 1 FROM tfproduc WHERE consec = inprod AND Inprod = '" & rs!CONSEC & "'"
             Else
                cn.Execute "UPDATE inventario SET incantpza = incantpza + 50 WHERE Inprod = '" & rs!CONSEC & "'"
             End If
             rs.Requery
          Else
             cn.CommitTrans
             Exit Sub
          End If
       'End If
    End If
    'Se verifica la existencia de Medio Mayoreo
    If Val(txtcampos(13).Text) + Val(txtcampos(14).Text) > rs!InCantPza And Sql Then
       If MsgBox("LA EXISTENCIA EN PIEZAS ES MENOR A LO SOLICITADO EN MEDIO MAYOREO Y PIEZAS, DESEAS CONVERTIR UNA CAJA EN PIEZAS", vbQuestion + vbYesNo) = vbYes Then
          cn.Execute "UPDATE inventario SET incantpza = incantpza + paquetes, incant = incant - 1 FROM tfproduc WHERE consec = inprod AND Inprod = '" & rs!CONSEC & "'"
       Else
          cn.CommitTrans
          Exit Sub
       End If
    End If
    'Se buscan los cargos del producto para grabar en el detalle de la venta y se pueda facturar
    rsttemp.Open "SELECT * FROM TASAIEPS WHERE DEPTO = " & rs!tasaieps, cn, adOpenKeyset, adLockOptimistic, adCmdText
    If rsttemp.BOF And rsttemp.EOF Then
        MsgBox "NO SE PUEDE CAPTURAR ESTE PRODUCTO PORQUE ESTA CLASIFICADO EN EL DEPTO " & rs!tasaieps & "; INFORME AL DEPTO. DE COMPRAS PARA QUE PROCEDA A CLASIFICARLO"
        Exit Sub
    Else
        iva = rsttemp!iva: ieps = rsttemp!ieps
        rsttemp.Close
    End If
     
    If Val(txtcampos(14).Text) > 0 And Val(txtcampos(13)) > 0 Then    'Medias cajas y piezas
       'npreciop = (Val(txtcampos(10).Text) + nprevtapm) / 2
       NPRECIOP = ((txtcampos(13).Text * (txtcampos(10).Text)) + (nprevtapm * txtcampos(14))) / (Val(txtcampos(13).Text) + Val(txtcampos(14).Text))
    ElseIf Val(txtcampos(13).Text) > 0 Then  'Si solo son piezas
       NPRECIOP = Val(txtcampos(10).Text)
    ElseIf Val(txtcampos(14).Text) > 0 Then  'Si solo son medias cajas
       NPRECIOP = nprevtapm
    End If
    If Sql Then
      If cModifico = "" Then
         CAD = "INSERT INTO VENTAS_DET(NoVenta,cl_producto,cantidad,cantidadp,precio,TipoCantidad,Ieps,Iva,precosto,Importe,precostop,preciop,TASAIEPS) VALUES (" & AdoVentas.Recordset!noventa & ",'" & rs!CONSEC & "'," & txtcampos(6).Text & "," & Val(txtcampos(13).Text) + Val(txtcampos(14).Text) & "," & NPREVTA & ",1," & ieps & "," & iva & "," & rs!PRECOSTO & "," & (NPREVTA * txtcampos(6).Text) + (txtcampos(13).Text * (txtcampos(10).Text)) + (nprevtapm * txtcampos(14)) & "," & rs!PRECOSTO / rs!PAQUETES & "," & NPRECIOP & "," & rs!tasaieps & ")"
         cn.Execute CAD
      Else
        rsttemp.Open "SELECT * FROM preprod WHERE preclave = '" & rs!CONSEC & "'", cn, adOpenDynamic, adLockOptimistic, adCmdText
        nPrebajo = IIf(rsttemp!PRECIO2 <= Val(NPREVTA), 0, 1)
        CAD = "INSERT INTO VENTAS_DET(NoVenta,cl_producto,cantidad,cantidadp,precio,TipoCantidad,Ieps,Iva,precosto,Importe,Modifico,prebajo,precostop,preciop,TASAIEPS) VALUES (" & AdoVentas.Recordset!noventa & ",'" & rs!CONSEC & "'," & txtcampos(6).Text & "," & Val(txtcampos(13).Text) + Val(txtcampos(14).Text) & "," & NPREVTA & ",1," & ieps & "," & iva & "," & rs!PRECOSTO & "," & (NPREVTA * txtcampos(6).Text) + (txtcampos(13).Text * (txtcampos(10).Text)) + (nprevtapm * txtcampos(14).Text) & ",'" & cModifico & "'," & nPrebajo & "," & rs!PRECOSTO / rs!PAQUETES & "," & (txtcampos(13).Text * txtcampos(10)) + (txtcampos(14).Text * nprevtapm) / 2 & "," & rs!tasaieps & ")"
        cn.Execute CAD
        rsttemp.Close
     End If
    Else  'Para preventa
        If ESCALA = 2 Then
           NPREVTA = txtcampos(8).Text
        ElseIf ESCALA = 3 Then
           NPREVTA = txtcampos(15).Text
        ElseIf ESCALA = 4 Then
           NPREVTA = txtcampos(7).Text
        End If
        cModifico = ""
        CAD = "INSERT INTO VENTAS_DET(NoVenta,cl_producto,cantidad,cantidadp,precio,TipoCantidad,Ieps,Iva,precosto,Importe,Modifico,precostop,preciop,escala) VALUES (" & AdoVentas.Recordset!noventa & ",'" & rs!CONSEC & "'," & txtcampos(6).Text & "," & Val(txtcampos(13).Text) + Val(txtcampos(14).Text) & "," & NPREVTA & ",1," & ieps & "," & iva & "," & rs!PRECOSTO & "," & (NPREVTA * txtcampos(6).Text) + (txtcampos(13).Text * (txtcampos(10).Text)) + (nprevtapm * txtcampos(14).Text) & ",'" & cModifico & "'," & rs!PRECOSTO / rs!PAQUETES & "," & NPRECIOP & "," & ESCALA & ")"
        'MsgBox CAD
        cn.Execute CAD
    End If
    If Sql Then
       cn.Execute "UPDATE INVENTARIO SET inCant = Incant - " & (txtcampos(6).Text) & ", inCantPza = IncantPza - " & Val(txtcampos(13).Text) + Val(txtcampos(14).Text) & " WHERE INPROD = '" & rs!CONSEC & "'"
    End If
    'se actualiza el iventario inicial por dia
    If Format(AdoVentas.Recordset!fecha, "dd-mm-yyyy") < Format(date, "dd-mm-yyyy") And Sql Then
       dfecha = Format(DateAdd("d", 1, AdoVentas.Recordset!fecha), "dd-mm-yyyy")
       While dfecha <= Format(date, "dd-mm-yyyy")
          diaini = "dia" & CStr(Day(dfecha))
          cn.Execute "UPDATE invcorte SET " & diaini & " = " & diaini & " - " & (txtcampos(6).Text) & " WHERE producto = '" & rs!CONSEC & "' AND mes = " & Month(date)
          dfecha = DateAdd("d", 1, dfecha)
       Wend
    End If
    If Me.chkCredito = 1 Or Me.chkPrev.Value = 1 Then   'En preventas se imprimen solamente los pedidos editados
       AdoVentas.Recordset!situacion = 0
       If Sql Then AdoVentas.Recordset.Update
       If Sql And Me.chkPrevLiq.Value = 1 Then cn.Execute "INSERT INTO prodborra(clave,noventa,fecha,cajas,piezas,importe,usuario,prevliq,preventa) VALUES ('" & rs!CONSEC & "'," & AdoVentas.Recordset!noventa & ",'" & date + Time & "'," & txtcampos(6).Text & "," & Val(txtcampos(13).Text) + Val(txtcampos(14).Text) & "," & (NPREVTA * txtcampos(6).Text) + (txtcampos(13).Text * (txtcampos(10).Text)) + (nprevtapm * txtcampos(14).Text) & ",'" & cBorrar & "'," & chkPrevLiq.Value & "," & txtcampos(2).Text & ")"
    End If
    If Sql Then cn.CommitTrans

    If rsttemp.State = 1 Then rsttemp.Close
    rsttemp.Open "SELECT SUM(Importe) AS Subto FROM VENTAS_DET, TFPRODUC WHERE consec = cl_producto AND NOVENTA = " & AdoVentas.Recordset!noventa, cn, adOpenForwardOnly, adLockReadOnly, adCmdText
    'rsttemp.Open "SELECT SUM(Importe) AS Subto, SUM( CASE CLAPROVE WHEN 'C52' THEN IMPORTE END) AS PROMO FROM VENTAS_DET, TFPRODUC WHERE consec = cl_producto AND NOVENTA = " & AdoVentas.Recordset!noventa, cn, adOpenKeyset, adLockOptimistic, adCmdText
    If IsNull(rsttemp!subto) Then
       lblImpVta.Caption = "$ 0.00"
       lblpromo.Caption = "PROM:  $ 0.00      BOLETOS: 0"
       nImpVta = 0
    Else
       lblImpVta.Caption = Format(rsttemp!subto, "$ ###,###,##0.00")
       'lblpromo.Caption = IIf(IsNull(rsttemp!PROMO), "PROM:  $ 0.00    BOLETOS: 0", "PROM.:  " & Format(rsttemp!PROMO, "$ ###,###,##0.00") & "      BOLETOS: " & Int(rsttemp!PROMO / 50))
       nImpVta = rsttemp!subto
    End If
    rsttemp.Close
    'ACTUALIZANDO EL MONTO DE LA VENTA EN GENERAL
    cn.Execute "UPDATE ventas SET MontoTotal = " & nImpVta & " WHERE Noventa = " & AdoVentas.Recordset!noventa
    cmbprod.SetFocus
    AdoDetVta.Refresh
    lbletiquetas(21).Caption = ""
    If IsNumeric(cmbprod.Text) Then
       cmbprod.Text = ""
       cmbprod.SetFocus
    End If
Exit Sub
Error:
  MsgBox "OCURRIO EL SIGUIENTE ERROR: " & Err.Description & Chr(13) & "A CONTINUACION SE CANCELARAN LOS CAMBIOS REALIZADOS"
  'If Sql Then cn.RollbackTrans
End Sub

Private Sub cmdAgrVta_Click()
Dim tipoventa
Dim agente
Dim chofer
Dim RSTCLIE As ADODB.Recordset
  tipoventa = AdoVentas.Recordset!tipoventa
  agente = AdoVentas.Recordset!agente
  chofer = AdoVentas.Recordset!chofer
  Preventa = AdoVentas.Recordset!FOLPREVENTA
  If Not Sql Then AdoVentas.Refresh
  AdoVentas.Recordset.AddNew
  frmVentas.txtcampos(1).Text = tipoventa
  frmVentas.txtcampos(9).Text = agente
  frmVentas.txtcampos(3).Text = IIf(IsNull(chofer), "", chofer)
  txtcampos(0).Text = date & " " & Time
  chkCredito.Value = 0
  'chkVales.Value = 0
  'chkPrev.Value = 0
  AdoVentas.Recordset!Prevta = 1
  AdoVentas.Recordset!facrfc = 0
  
  nOp = 0
  frapreventa.Visible = False
  cmdGrabar.Enabled = True
  cmdClientes.Enabled = True
  txtCte.Enabled = True
  fraCliente.Visible = False
  FraVtadet.Visible = False
  FraTotal.Visible = False
  DbgDetVta.Visible = False
  txtcampos(4).Enabled = True
  cmbCliente.Enabled = True
  txtcampos(2).Text = Preventa
  chkConRfc.Enabled = True
  chkPrev.Enabled = True
  chkCredito.Enabled = True
  txtCte.Text = ""
  Me.txtCte.SetFocus
End Sub

Private Sub cmdCamSer_Click()
 nOpcion = 1
 txtContra.Text = ""
 fraCon.Visible = True
 txtContra.SetFocus
End Sub

Private Sub cmdClientes_Click()
  CarCte = False
  frmCliente.Show 1
End Sub

Private Sub cmdConAceptar_Click()
Dim RsCon As ADODB.Recordset
Dim RST As ADODB.Recordset
Set RsCon = New ADODB.Recordset
RsCon.Open "SELECT * FROM USUARIOS WHERE PASSMASTER = '" & Trim(txtContra.Text) & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
If RsCon.BOF And RsCon.EOF Then
   MsgBox "LA CONTRASEÑA ES INCORRECTA", vbExclamation
   txtContra.SetFocus
   SendKeys "+{HOME}"
   Exit Sub
Else
   If nOpcion = 0 Then  'Cambio de precios
      If Not autoriza(RsCon!permisos, 3) Then Exit Sub
      MsgBox RsCon!Name & Chr(13) & "A CONTINUACION SE REGISTRARA SU CLAVE PARA DETERMINAR LOS MOVIMIENTOS DIARIOS DE PRECIOS", vbInformation
      cModifico = RsCon!login
      Index = IIf(chkCredito.Value = 0 And chkPrev.Value = 0, 7, 8)
      txtcampos(Index).Enabled = True
      txtcampos(10).Enabled = True
      fraCon.Visible = False
      CmdAgregar.Enabled = False
   ElseIf nOpcion = 1 Then 'Cambio de escala de precios
      If autoriza(RsCon!permisos, 8) Or autoriza(RsCon!permisos, 10) Then
         fraCon.Visible = False
         Set RST = New ADODB.Recordset
         If Not autoriza(RsCon!permisos, 10) Then
            RST.Open "SELECT * FROM catcliente WHERE cambpre = 1 and cclave = " & txtcampos(4).Text, cn, adOpenForwardOnly, adLockOptimistic, adCmdText
            If RST.BOF And RST.EOF Then
               MsgBox "A " & Me.cmbCliente.Text & " NO SE LE PERMITE CAMBIAR ESCALA DE PRECIOS", vbCritical
               Exit Sub
            End If
            RST.Close
         End If
         RESP = InputBox("Proporcione la escala a la que se actualizaran precios?" & Chr(13) & Chr(13) & _
                 "2 = Escala crédito y/o envío" & Chr(13) & "3 = Escala intermedia" & Chr(13) & "4 = Escala en bodega", "Proporcione escala")
         If Trim(RESP) = "" Then Exit Sub
         If Val(RESP) = 2 Then
            cn.Execute "UPDATE ventas_det SET precio = precio2, precioP = precio1, importe = (cantidad * precio2) + (cantidadp * precio1), modifico = '" & RsCon!login & "' FROM preprod " & _
                       "WHERE preclave = cl_producto AND noventa = " & AdoVentas.Recordset!noventa
         ElseIf Val(RESP) = 3 Then
            cn.Execute "UPDATE ventas_det SET precio = precio3, importe = (cantidad * precio3) + (cantidadp * preciop), modifico = '" & RsCon!login & "', prebajo = 1 FROM preprod " & _
                       "WHERE preclave = cl_producto AND noventa = " & AdoVentas.Recordset!noventa
         ElseIf Val(RESP) = 4 Then
            cn.Execute "UPDATE ventas_det SET precio = precio4, importe = (cantidad * precio4) + (cantidadp * preciop), modifico = '" & RsCon!login & "', prebajo = 1 FROM preprod " & _
                       "WHERE preclave = cl_producto AND noventa = " & AdoVentas.Recordset!noventa
         Else
            MsgBox "NUMERO DE ESCALA INCORRECTA", vbCritical
         End If
         RST.Open "SELECT SUM(Importe) AS Subto FROM VENTAS_DET, TFPRODUC WHERE consec = cl_producto AND NOVENTA = " & AdoVentas.Recordset!noventa, cn, adOpenKeyset, adLockOptimistic, adCmdText
         lblImpVta.Caption = Format(RST!subto, "$ ###,###,##0.00")
         'ACTUALIZANDO EL MONTO DE LA VENTA EN GENERAL
         cn.Execute "UPDATE ventas SET MontoTotal = " & RST!subto & " WHERE Noventa = " & AdoVentas.Recordset!noventa
         RST.Close
         Set RST = Nothing
         AdoDetVta.Refresh
      End If
   ElseIf nOpcion = 2 Then 'Venta por pieza mayor a los paquetes por caja
       If Not autoriza(RsCon!permisos, 9) Then Exit Sub
       lMasPza = True
       Me.fraCon.Visible = False
       'Solo piezas
       cn.Execute "UPDATE ventas_det SET precio = ROUND(importe / cantidadp,2), precioP = ROUND(importe / cantidadp,2) WHERE noventa = " & AdoVentas.Recordset!noventa & " and cantidadp > 0 AND cantidad = 0"
       'Cajas y piezas
       cn.Execute "UPDATE ventas_det SET precioP = ROUND(precio / T.PAQUETES,2), precio = ROUND(precio,2) FROM tfproduc t WHERE consec = cl_producto AND noventa = " & AdoVentas.Recordset!noventa & " and cantidadp > 0 AND cantidad > 0"
       cn.Execute "UPDATE ventas_det SET Importe = ROUND( (cantidad * precio) + (cantidadp * preciop),2 ) WHERE noventa = " & AdoVentas.Recordset!noventa
       
       Set RST = New ADODB.Recordset
       RST.Open "SELECT SUM(Importe) AS Subto FROM VENTAS_DET, TFPRODUC WHERE consec = cl_producto AND NOVENTA = " & AdoVentas.Recordset!noventa, cn, adOpenKeyset, adLockOptimistic, adCmdText
       lblImpVta.Caption = Format(RST!subto, "$ ###,###,##0.00")
       'ACTUALIZANDO EL MONTO DE LA VENTA EN GENERAL
       cn.Execute "UPDATE ventas SET MontoTotal = " & IIf(IsNull(RST!subto), 0, RST!subto) & " WHERE Noventa = " & AdoVentas.Recordset!noventa
       RST.Close
       Set RST = Nothing
       AdoDetVta.Refresh
       MsgBox "SE HA ACTIVADO LA VENTA DE PIEZAS MAYOR A LAS QUE TRAE UNA CAJA PARA TODA ESTA VENTA, PARA AJUSTAR EL PRECIO UNITARIO A PIEZAS, AL FINALIZAR LA CAPTURA EJECUTE NUEVAMENTE ESTA OPCION", vbInformation, "Ventas"
   ElseIf nOpcion = 4 Then 'Cancelar productos de la venta
       If Not autoriza(RsCon!permisos, 12) Then Exit Sub
       cBorrar = Trim(RsCon!login)
       Me.fraCon.Visible = False
       MsgBox RsCon!Name & Chr(13) & "SE HA ACTIVADO LA CAPTURA Y CANCELACION DE PRODUCTOS" & Chr(13) & "AL TERMINAR SALGA DE LA VENTA ", vbInformation, "Ventas"
  ElseIf nOpcion = 5 Then 'Prevestistas de Philip Morris Costo + 1 %
       If Not autoriza(RsCon!permisos, 9) Then Exit Sub
       fraCon.Visible = False
       cBorrar = Trim(RsCon!login)
       If InStr(1, cmbCliente.Text, "PHILIP MOR") = 0 Then
          MsgBox "ESTE CLIENTE NO ESTA CONSIDERADO COMO PREVENTISTA DE PHILIP MORRIS", vbInformation, "Ventas"
          Exit Sub
       Else
          cn.Execute "UPDATE ventas_det SET precio = ROUND( T.PREcosto  + (T.PREcosto * 0.01) ,1,-2) + .10 FROM tfproduc t WHERE t.consec = cl_producto And CLAprove = 'C29' And noventa = " & AdoVentas.Recordset!noventa
          cn.Execute "UPDATE ventas_det SET Importe = (cantidad * precio) + (cantidadP * preciop)  WHERE noventa = " & AdoVentas.Recordset!noventa
          Set RST = New ADODB.Recordset
          RST.Open "SELECT SUM(Importe) AS Subto FROM VENTAS_DET, TFPRODUC WHERE consec = cl_producto AND NOVENTA = " & AdoVentas.Recordset!noventa, cn, adOpenForwardOnly, adLockReadOnly, adCmdText
          lblImpVta.Caption = Format(RST!subto, "$ ###,###,##0.00")
          'ACTUALIZANDO EL MONTO DE LA VENTA EN GENERAL
          cn.Execute "UPDATE ventas SET MontoTotal = " & RST!subto & " WHERE Noventa = " & AdoVentas.Recordset!noventa
          RST.Close
          Set RST = Nothing
          AdoDetVta.Refresh
          MsgBox "LA VENTA SE ACTUALIZO CORRECTAMENTE", vbInformation, "Ventas"
       End If
   ElseIf nOpcion = 6 Then
      If Not autoriza(RsCon!permisos, 1) Then Exit Sub
      MsgBox RsCon!Name & Chr(13) & "A CONTINUACION SE REGISTRARA SU CLAVE PARA DETERMINAR LOS CAMBIOS DE PRECIOS", vbInformation
      RESP = InputBox("Porcentaje a incrementar a precio de costo", "Porcentaje", 1)
      If Not IsNumeric(RESP) Then Exit Sub
      cn.Execute "UPDATE ventas_det SET PRECIO = t.PRECOSTO * 1.01 , PRECIOP = t.precosto * 1.01 / paquetes , importe = (t.precosto * 1.01 *  cantidad) + (t.precosto * 1.01 / paquetes  * cantidadp) FROM tfproduc t WHERE consec = cl_producto AND noventa = " & AdoVentas.Recordset!noventa
      cn.Execute "UPDATE VENTAS set montototal = ( SELECT SUM(importe) FROM ventas_det WHERE noventa = " & AdoVentas.Recordset!noventa & ") WHERE noventa = " & AdoVentas.Recordset!noventa
      fraCon.Visible = False
      CmdAgregar.Enabled = False
      MsgBox "LA VENTA SE ACTUALIZO CORRECTAMENTE", vbInformation, "Ventas"
      Unload Me
   Else  'Cambio de precios solo a mas alto del normal
      If Not autoriza(RsCon!permisos, 9) Then Exit Sub
      fraCon.Visible = False
      Index = IIf(chkCredito.Value = 0, 7, 8)
      txtcampos(Index).Enabled = True
      txtcampos(10).Enabled = True
      txtcampos(Index).SetFocus
      txtcampos(Index).SelStart = 0
      txtcampos(Index).SelLength = Len(txtcampos(Index).Text)
   End If
End If
End Sub

Private Sub cmdConCance_Click()
  fraCon.Visible = False
End Sub

Private Sub cmdDevPvta_Click()
 crpt.WindowTitle = "HOJA DE DEVOLUCION DE LA PREVENTA " & txtcampos(2).Text
 crpt.ReportFileName = App.Path & "\PrevLiq.rpt"
 crpt.Connect = cCadConex
 crpt.Formulas(1) = "ENCABEZADO = 'HOJA DE DEVOLUCION DE LA PREVENTA " & Trim(txtcampos(2).Text) & "'"
 crpt.Formulas(2) = ""
 crpt.Formulas(3) = ""
 crpt.Formulas(4) = ""
 cadsql = "SELECT TFPRODUC.DESCRIPC, TFPRODUC.CONTENID, TFPRODUC.MEDIDA, TFPRODUC.PAQUETES, " & _
                 "PRODBORRA.clave, PRODBORRA.cajas, PRODBORRA.piezas, PRODBORRA.importe, PRODBORRA.prevliq, PRODBORRA.preventa, " & _
                 "PREVENTAS.folio, PREVENTAS.impcapt, PREVENTAS.impsurt, PREVENTAS.impliq " & _
          "FROM pitico.dbo.TFPRODUC TFPRODUC, " & _
                 "PITICO.dbo.PRODBORRA PRODBORRA, " & _
                 "PITICO.dbo.PREVENTAS PREVENTAS " & _
          "WHERE TFPRODUC.CONSEC = PRODBORRA.clave AND " & _
                 "PRODBORRA.preventa = PREVENTAS.folio AND " & _
                 "PRODBORRA.prevliq = 1 AND " & _
                 "PRODBORRA.preventa = " & txtcampos(2).Text & ". " & _
          "ORDER BY TFPRODUC.DESCRIPC ASC "
 crpt.SQLQuery = cadsql
 crpt.Action = 1
End Sub

Private Sub CmdExporta_Click()
Dim CNMDB As ADODB.Connection
Dim tmp As ADODB.Recordset
Set tmp = New ADODB.Recordset
Set CNMDB = New ADODB.Connection
 adopreventa.Recordset.MoveFirst
 DbgrdPreventa.SetFocus
 CNMDB.Open "DSN=PITICOMDB;DBQ=" & App.Path & "\" & txtcampos(9).Text & ".mdb;DefaultDir=" & App.Path & ";DriverId=25;FIL=MS Access;MaxBufferSize=2048;PageTimeout=5;UID=admin"
 CNMDB.Execute "DELETE FROM pedidos"
 CNMDB.Execute "DELETE FROM pedidosdet"
 CNMDB.Execute "DELETE FROM clientes"
 N = 0
 horant = Format(adopreventa.Recordset!fecha, "HH:MM")
 While Not adopreventa.Recordset.EOF
   If adopreventa.Recordset!montototal > 0 Then
     N = N + 1
     tmp.Open "SELECT * FROM ventas_det WHERE noventa = " & adopreventa.Recordset!noventa, cn, adOpenForwardOnly, adLockReadOnly, adCmdText
     nutilidad = 0
     While Not tmp.EOF
        CNMDB.Execute "INSERT INTO pedidosdet(folio,clave,cajas,piezas,precio,preciop,importe,escala,costo,costop) VALUES (" & N & ",'" & tmp!cl_producto & "'," & tmp!cantidad & "," & tmp!cantidadp & "," & tmp!PRECIO & "," & tmp!preciop & "," & tmp!importe & "," & tmp!ESCALA & "," & tmp!PRECOSTO & "," & tmp!PREcostop & ")"
        nutilidad = nutilidad + (tmp!importe - ((tmp!cantidad * tmp!PRECOSTO) + (tmp!cantidadp * tmp!PREcostop)))
        tmp.MoveNext
     Wend
     CNMDB.Execute "INSERT INTO pedidos(folio,cliente,factura,importe,credito,hora,tiempo,utilidad) VALUES (" & N & "," & adopreventa.Recordset!cclave & "," & IIf(adopreventa.Recordset!facrfc, 1, 0) & "," & adopreventa.Recordset!montototal & "," & IIf(adopreventa.Recordset!ccredito, 1, 0) & ",'" & Format(adopreventa.Recordset!fecha, "HH.MM") & "'," & DateDiff("n", horant, Format(adopreventa.Recordset!fecha, "HH:MM")) & "," & nutilidad & ")"
     horant = Format(adopreventa.Recordset!fecha, "HH:MM")
     tmp.Close
   End If
   adopreventa.Recordset.MoveNext
 Wend
 Open App.Path & "\" & Format(txtcampos(0).Text, "YYYYMMDD") & "-" & txtcampos(9).Text & "-P.TXT" For Output As #1
 tmp.Open "SELECT * FROM pedidos ORDER BY folio", CNMDB
 While Not tmp.EOF
     Print #1, tmp!Folio & Chr(9) & tmp!CLIENTE & Chr(9) & IIf(tmp!Factura, 1, 0) & Chr(9) & Format(tmp!importe, "########0.00") & Chr(9); IIf(tmp!credito, 1, 0) & Chr(9) & Format(tmp!HORA, "HH:MM:SS") & Chr(9) & tmp!TIEMPO & Chr(9) & Format(tmp!UTILIDAD, "########0.00")
     tmp.MoveNext
 Wend
 Close #1
 tmp.Close
 tmp.Open "SELECT * FROM catcliente WHERE MODIFICADO ", cn, adOpenForwardOnly, adLockReadOnly, adCmdText
 While Not tmp.EOF
     CNMDB.Execute "INSERT INTO clientes(clave,nombre,calle,colonia,poblacion,rfc,telefono,nuevo,modificado,nombres,apepaterno,apematerno,nomnegocio,ruta) VALUES (" & _
                tmp!cclave & ",'" & tmp!cNombre & "','" & tmp!cdireccion & "','" & tmp!ccolonia & "','" & tmp!cciudad & "','" & tmp!crfc & "','" & tmp!ctelefono & "'," & IIf(tmp!cclave < 0, 1, 0) & "," & IIf(tmp!MODIFICADO, 1, 0) & ",'" & tmp!nombres & "','" & tmp!apepaterno & "','" & tmp!apematerno & "','" & tmp!nomnegocio & "','" & tmp!ruta & "')"
     tmp.MoveNext
 Wend
 MsgBox "LA EXPORTACION SE REALIZO CORRECTAMENTE", vbInformation, "Exportación"
End Sub

Private Sub cmdFacConFin_Click()
Dim rsttemp As New ADODB.Recordset
Set rsttemp = New ADODB.Recordset
  If AdoVentas.Recordset.State = 0 Then AdoVentas.Refresh
  rsttemp.Open "SELECT v.noventa FROM VENTAS V, VENTAS_DET D WHERE V.NOVENTA = D.NOVENTA AND d.facturado = 0 AND V.FACRFC = 0 AND v.Credito = 0 AND V.FOLPREVENTA = " & AdoVentas.Recordset!FOLPREVENTA, cn, adOpenStatic, adLockOptimistic, adCmdText
  If rsttemp.BOF And rsttemp.EOF Then
     MsgBox "NO EXISTEN VENTAS PARA IMPRIMIR A CONSUMIDOR FINAL" & Chr(13) & "CON SITUACION 1 Y FOLIO DE PREVENTA " & AdoVentas.Recordset!FOLPREVENTA & Chr(13) & "", vbExclamation
     Exit Sub
  End If
  AdoVentas.CommandType = adCmdText
  AdoVentas.CursorType = adOpenDynamic
  AdoVentas.RecordSource = "SELECT * FROM VENTAS WHERE noventa = " & adopreventa.Recordset!noventa
  AdoVentas.Refresh
  If AdoVentas.Recordset!credito Or AdoVentas.Recordset!Prevta Then
        TDA = Mid(cSucursal, 1, 2)
        If TDA = 16 Then 'MIGUEL CABRERA
            todobien = FacturaNva(True)
        ElseIf TDA = 24 Then 'CENTRAL MAYOREO
            todobien = FacturaNva(True)
        ElseIf TDA = 10 Then  ' OFICINAS FRANQUICIAS
            todobien = Factura
        ElseIf TDA = 55 Then  ' PUERTO ESCONDIDO
            todobien = FacturaNva(True)
        ElseIf TDA = 26 Then  ' MIAHUATLAN
            todobien = FacturaNva(True)
        ElseIf TDA = 28 Then  ' ISTMO
            todobien = FacturaNva(True)
        Else
            todobien = FacturaNva(True)
        End If
  End If
  If todobien Then
     cn.Execute "UPDATE VENTAS SET situacion = '3' WHERE folpreventa = " & adopreventa.Recordset!FOLPREVENTA & " AND FACRFC = 0 AND Credito = 0"
     adopreventa.Refresh
     cmbprod.Clear
  End If
End Sub
'Factura por cliente toma el cliente seleccionado
Private Sub cmdFacCte_Click()
If adopreventa.Recordset!facrfc Then
   If adopreventa.Recordset!situacion = 3 Then
      MsgBox "AL CLIENTE SELECCIONADO YA SE LE IMPRIMIO FACTURA", vbInformation
      Exit Sub
   End If
   AdoVentas.CommandType = adCmdText
   AdoVentas.CursorType = adOpenDynamic
   AdoVentas.RecordSource = "SELECT * FROM VENTAS WHERE noventa = " & adopreventa.Recordset!noventa
   AdoVentas.Refresh
   cmbCliente.Text = adopreventa.Recordset!cNombre
   If MsgBox("REALMENTE DESEAS GENERAR FACTURA DEL CLIENTE" & Chr(13) & adopreventa.Recordset!cNombre, vbQuestion + vbYesNo) = vbYes Then
      ccVeCli = txtcampos(4).Text
      If rscli.State = adStateOpen Then rscli.Close
      rscli.Open "SELECT * FROM Catcliente WHERE cClave = '" & ccVeCli & "' AND ctipo = 0", cn, adOpenKeyset, adLockOptimistic, adCmdText
      cmbCliente.Text = rscli!cNombre
      lblNombre.Caption = IIf(IsNull(rscli!cnombrefac), "", rscli!cnombrefac)
      lblDirec.Caption = IIf(IsNull(rscli!cdireccion), "", rscli!cdireccion)
      lblRfc.Caption = IIf(IsNull(rscli!crfc), "", rscli!crfc)
      lblTelefono.Caption = rscli!ctelefono

      If AdoVentas.Recordset!credito Or AdoVentas.Recordset!Prevta Then
      'TIENDAS
        TDA = Mid(cSucursal, 1, 2)
        If TDA = 16 Then 'MIGUEL CABRERA
            todobien = FacturaNva
        ElseIf TDA = 24 Then 'CENTRAL MAYOREO
            todobien = Factura
        ElseIf TDA = 10 Then  ' OFICINAS FRANQUICIAS
            todobien = Factura
        ElseIf TDA = 55 Then  ' PUERTO ESCONDIDO
            todobien = FacturaNva
        ElseIf TDA = 26 Then  ' MIAHUATLAN
            todobien = FacturaNva
        ElseIf TDA = 28 Then  ' ISTMO
            todobien = FacturaNva
        End If
      End If
      If todobien Then
         If chkCredito.Value = 0 And chkPrev.Value = 0 Then
            CAD = "update ventas set situacion =  2  where noventa = " & noventar
         Else
            CAD = "update ventas set situacion =  3  where noventa = " & noventar
         End If
         cn.Execute CAD
         MsgBox "Impresión de Factura Finalizada....", vbInformation, "Ventas"
      End If
      adopreventa.Refresh
   End If
Else
   MsgBox "AL CLIENTE SELECCIONADO SE CAPTURO COMO CONSUMIDOR FINAL" & Chr(13) & "SI DESEA IMPRIMIR CON RFC MODIFIQUE LA VENTA", vbExclamation
End If
End Sub

Private Sub cmdGrabar_Click()
Dim RSTFOL As ADODB.Recordset
Dim rsvencre As ADODB.Recordset
Dim VenCre
Dim nreg As Integer
Dim folVta As Integer
On Error GoTo Error:
'cn.BeginTrans
nDispo = 0
cmbprod.Clear
nPos = InStr(1, "152", Trim(txtcampos(1).Text))
If ((chkCredito.Value = 0 And Me.chkPrev.Value = 0) Or (chkCredito.Value = 1 And Me.chkPrev.Value = 1)) And Trim(txtcampos(1).Text) = 2 Then
   MsgBox "ES NECESARIO ESPECIFICAR SI ES A CREDITO O DE PREVENTAS", vbInformation, "Ventas"
   Exit Sub
ElseIf chkCredito.Value = 1 And nPos > 1 And Not rscli!ccredito Then
   MsgBox "AL CLIENTE " & cmbCliente.Text & Chr(13) & "NO SE LE PERMITEN  VENTAS A CREDITO", vbExclamation, "Ventas"
   chkCredito.SetFocus
 '  cn.CommitTrans
   Exit Sub
ElseIf chkConRfc.Value = 1 And (IsNull(rscli!crfc) Or Len(Trim(rscli!crfc)) < 8) Then
   MsgBox "AL CLIENTE " & cmbCliente.Text & Chr(13) & "NO SE LE PUEDE FACTURAR CON RFC" & Chr(13) & "PORQUE NO SE HA CAPTURADO ESTE DATO", vbExclamation
   chkConRfc.SetFocus
  ' cn.CommitTrans
   Exit Sub
ElseIf chkCredito.Value = 1 And rscli!ccredito And rscli!ctiempocredito > 0 Then   'Que no sea en opción Cobro
   If vercredito(rscli!cclave) = True Then
      MsgBox "YA NO ES POSIBLE OTORGAR MAS CREDITOS A ESTE CLIENTE", vbCritical, "Agotado el crédito"
      AdoVentas.Refresh
      Exit Sub
   End If
End If
If Not Sql Then frmCliente.Show 1

'If chkVales.Value = 1 Then MsgBox "VERIFIQUE QUE EL VALE SEA ORIGINAL, EL FOLIO EN COLOR ROJO CON SELLO Y FIRMAS DE AUTORIZACION", vbInformation, "Ventas"
chkCredito.Enabled = False
chkConRfc.Enabled = False
chkPrev.Enabled = False
'Deshabilito los datos generales
For N = 0 To 4
  txtcampos(N).Enabled = False
Next
txtcampos(9).Enabled = False
cmbTipVta.Enabled = False
cmbCliente.Enabled = False
txtCte.Enabled = False
cmdGrabar.Enabled = False
cmdClientes.Enabled = False
cmbAgente.Enabled = False
cmbChofer.Enabled = False
If Me.chkPrev.Value = 1 Then
  If Mid(cSucursal, 1, 3) = 16 Then
     MsgBox "NO ESTA PERMITIDO VENDER EN PREVENTA", vbInformation, "Error"
     Exit Sub
  End If
End If

If nOp = 0 Then     'Si es una nueva venta
   Set RSTFOL = New ADODB.Recordset
   RSTFOL.Open "SELECT * FROM FOLIOS WHERE Sucursal = '" & Mid(cSucursal, 1, 3) & "' AND CAJA = '" & Caja & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
   If RSTFOL.BOF And RSTFOL.EOF Then
      'Agrego a la tabla de folios porque es primera venta de la caja
      MsgBox "BIENVENIDO AL MODULO DE VENTA A MAYOREO   " & Caja & Chr(13) & "A CONTINUACION SE REGISTRARA SU PRIMERA VENTA", vbInformation, "Bienvenida"
      cn.Execute "INSERT INTO FOLIOS(Sucursal,FolioVenta,FolioInfinito,FechaActualiza,Caja) VALUES ('" & Mid(cSucursal, 1, 3) & "',0,0,'" & date + Time & "','" & Caja & "')"
      RSTFOL.Close
      RSTFOL.Open "SELECT * FROM FOLIOS WHERE Sucursal = '" & Mid(cSucursal, 1, 3) & "' AND CAJA = '" & Caja & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
   End If
   RSTFOL!folioventa = RSTFOL!folioventa + 1
   RSTFOL!FolioInfinito = RSTFOL!FolioInfinito + 1
   TV = RSTFOL!folioventa
   TI = RSTFOL!FolioInfinito
   RSTFOL.Update
   AdoVentas.Recordset!CL_TERMINAL = Caja
   AdoVentas.Recordset!CL_operador = Trim(cUsuario)
   AdoVentas.Recordset!tienda = Mid(cSucursal, 1, 3)
   AdoVentas.Recordset!FolioInfinito = TV
   AdoVentas.Recordset!folioventa = TI
   AdoVentas.Recordset!situacion = "0"    'Venta Hecha por Mayoreo en tramite
   If chkCredito.Value = 1 Then
      AdoVentas.Recordset!Modocredito = "A"  'A = Activo,  V = VENCIDO, P = Pagado
      AdoVentas.Recordset!fechaPagcre = date + rscli!ctiempocredito
      AdoVentas.Recordset!Plazodias = rscli!ctiempocredito
   End If
   RSTFOL.Close
   'Si es venta a preventa y es nueva venta
   If Trim(txtcampos(1).Text) = 2 And Preventa = 0 Then
      RSTFOL.Open "SELECT MAX(FOLPREVENTA) AS FOLPRE FROM VENTAS", cn, adOpenKeyset, adLockOptimistic, adCmdText
      Prev = IIf(IsNull(RSTFOL!FOLPRE), 1, RSTFOL!FOLPRE + 1)
      AdoVentas.Recordset!FOLPREVENTA = Prev
      txtcampos(2).Text = Prev
      RSTFOL.Close
   End If
   AdoVentas.Recordset.Update
   Call Obtenprod   'Obtiene productos con existencia
   lCob = (nOp = 0)
Else  'Modificar o cobrar venta
   If nOp = 1 Then
      AdoVentas.Recordset!CL_operador = Trim(cUsuario)
      Call Obtenprod   'Obtiene productos con existencia
      'AdoVentas.Recordset.Update
   Else
      lblImpVta.Visible = True
   End If
   AdoVentas.Recordset!chofer = txtcampos(3).Text
   AdoVentas.Recordset.Update
   If nOp = 1 Then 'Cobrar venta
      DbgDetVta.Visible = True
      If nOp = 1 Then DbgDetVta.SetFocus
   End If
End If
lbletiquetas(11).Visible = IIf(IsNull(AdoVentas.Recordset!FOLPREVENTA) Or AdoVentas.Recordset!FOLPREVENTA <> 0, True, False)
txtcampos(2).Visible = IIf(IsNull(AdoVentas.Recordset!FOLPREVENTA) Or AdoVentas.Recordset!FOLPREVENTA <> 0, True, False)

Set rsttemp = New ADODB.Recordset
AdoDetVta.CommandTimeout = 0
AdoDetVta.ConnectionTimeout = 0
AdoDetVta.CursorType = adOpenStatic
AdoDetVta.ConnectionString = cCadConex
If Sql Then
   AdoDetVta.RecordSource = "SELECT cl_producto, Cantidad, cantidadp, Descripc ,LTRIM(STR(T.paquetes)) + ' X ' + LTRIM(STR(CONTENID,10,3)) + space(2) + t.medida AS MEDIDA, Precio, importe, serie + ' ' + factura AS factura FROM VENTAS_DET,TFPRODUC T WHERE cl_producto = T.Consec AND NoVenta = " & AdoVentas.Recordset!noventa & " ORDER BY T.DESCRIPC"
   If Me.chkCredito = 1 Or chkPrev.Value = 1 Then
    rsttemp.Open "SELECT * FROM preventas WHERE folio = " & txtcampos(2).Text, cn, adOpenForwardOnly, adLockOptimistic, admcdtext
    If Not (rsttemp.BOF And rsttemp.EOF) Then
       chkPrevLiq.Value = IIf(rsttemp!impsurt > 0, 1, 0)
       chkPrevLiq.Enabled = IIf(rsttemp!impsurt > 0, 0, 1)
       'Preventa liquidada
       chkLiquida.Visible = rsttemp!impsurt > 0
       chkLiquida.Value = IIf(rsttemp!impliq > 0, 1, 0)
       chkLiquida.Enabled = IIf(rsttemp!impliq > 0, 0, 1)
    Else
       chkPrevLiq.Value = 0
       chkLiquida.Value = 0
    End If
    rsttemp.Close
   End If
Else
   AdoDetVta.RecordSource = "SELECT cl_producto, Cantidad, cantidadp, Descripc ,LTRIM(STR(T.paquetes)) + ' X ' + LTRIM(STR(CONTENID)) + space(2) + t.medida AS MEDIDA, Precio, importe  FROM VENTAS_DET,TFPRODUC T WHERE cl_producto = T.Consec AND NoVenta = " & AdoVentas.Recordset!noventa & " ORDER BY T.DESCRIPC"
End If
AdoDetVta.Refresh

fraCliente.Visible = True
FraVtadet.Visible = True
FraTotal.Visible = True
DbgDetVta.Visible = True

lblImpVta.Visible = True
rsttemp.Open "SELECT SUM(Importe) AS Subto FROM VENTAS_DET, TFPRODUC WHERE consec = cl_producto AND NOVENTA = " & AdoVentas.Recordset!noventa, cn, adOpenKeyset, adLockOptimistic, adCmdText
'rsttemp.Open "SELECT SUM(Importe) AS Subto, SUM( CASE CLAPROVE WHEN 'C52' THEN IMPORTE END) AS PROMO FROM VENTAS_DET, TFPRODUC WHERE consec = cl_producto AND NOVENTA = " & AdoVentas.Recordset!noventa, cn, adOpenKeyset, adLockOptimistic, adCmdText
If Not IsNull(rsttemp!subto) Then
   lblImpVta.Caption = Format(rsttemp!subto, "$ ###,###,##0.00")
   'lblpromo.Caption = IIf(IsNull(rsttemp!PROMO), "PROM:  $ 0.00    BOLETOS: 0", "PROM.:  " & Format(rsttemp!PROMO, "$ ###,###,##0.00") & "      BOLETOS: " & Int(rsttemp!PROMO / 50))
   nImpVta = rsttemp!subto
Else
   lblpromo.Caption = "PROM:  $ 0.00      BOLETOS: 0"
   lblImpVta.Caption = "$ 0.00"
   nImpVta = 0
End If
lblImpVta.Refresh
rsttemp.Close
'cn.CommitTrans
Exit Sub
Error:
  MsgBox "OCURRIO UN ERROR :" & Err.Description & Chr(13) & "A CONTINUACION SE CANCELARAN LOS CAMBIOS REALIZADOS"
 ' cn.RollbackTrans
End Sub

Private Sub Obtenprod()
Dim rstprod As ADODB.Recordset
   PGB.Min = 0
   fraAvance.Visible = True
   fraAvance.Refresh
   cmbprod.Clear
   'Cargo Catalogo de productos con existencia mayor a cero
   Set rstprod = New ADODB.Recordset
   If Sql Then
      rstprod.Open "SELECT t.descripc,t.paquetes, LTRIM(STR(T.paquetes)) + ' X ' + lTrim(str(T.contenid,10,3)) + space(2) + t.medida As MEDIDA, t.contenid,t.consec FROM TFPRODUC t,INVENTARIO WHERE T.CONSEC = INVENTARIO.INPROD AND (INCANT > 0 OR IncantPza > 0)", cn, adOpenStatic, adLockOptimistic, adCmdText
   Else
      rstprod.Open "SELECT t.descripc,t.paquetes, LTRIM(STR(T.paquetes)) + ' X ' + lTrim(str(T.contenid)) + space(2) + t.medida As MEDIDA, t.contenid,t.consec FROM TFPRODUC t,INVENTARIO WHERE T.CONSEC = INVENTARIO.INPROD AND (INCANT > 0 OR IncantPza > 0)", cn, adOpenStatic, adLockOptimistic, adCmdText
   End If
   PGB.Max = rstprod.RecordCount
   nreg = 0
   rstprod.Close
   If Sql Then
      rstprod.Open "SELECT t.descripc,t.paquetes, LTRIM(STR(T.paquetes)) + ' X ' + lTrim(str(T.contenid,10,3)) + space(2) + t.medida As MEDIDA, t.contenid,t.consec FROM TFPRODUC t,INVENTARIO WHERE T.CONSEC = INVENTARIO.INPROD AND (INCANT > 0 OR IncantPza > 0)", cn, adOpenForwardOnly, adLockReadOnly, adCmdText
   Else
      rstprod.Open "SELECT t.descripc,t.paquetes, LTRIM(STR(T.paquetes)) + ' X ' + lTrim(str(T.contenid)) + space(2) + t.medida As MEDIDA, t.contenid,t.consec FROM TFPRODUC t,INVENTARIO WHERE T.CONSEC = INVENTARIO.INPROD AND (INCANT > 0 OR IncantPza > 0)", cn, adOpenForwardOnly, adLockReadOnly, adCmdText
   End If
   While Not rstprod.EOF
      nreg = nreg + 1
      PGB.Value = nreg
      cmbprod.AddItem rstprod!descripc & " " & rstprod!medida & "  " & rstprod!CONSEC
      rstprod.MoveNext
   Wend
   rstprod.Close
   Set rstprod = Nothing
   fraAvance.Visible = False
   If cmbprod.Visible Then cmbprod.SetFocus
End Sub

Private Sub cmdprod_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   SendKeys vbTab
ElseIf KeyAscii = 27 Then
   Unload Me
End If

End Sub

Function producto() As String
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
rs.Open "select descripc,contenid,medida, consec,paquetes from tfproduc where barraspza =  " & cmbprod.Text, cn, adOpenKeyset, adLockOptimistic, adCmdText
If rs.EOF Then
   MsgBox "No existe el codigo de barras", vbInformation
   producto = "NADA"
Else
   producto = rs!CONSEC
End If
Set rs = Nothing
End Function

Private Sub cmbprod_Validate(Cancel As Boolean)
On Error GoTo Error:
If Trim(cmbprod.Text) <> "" Then
   If IsNumeric(cmbprod.Text) Then
      cprod = Trim(producto())
   Else
      cprod = Trim(Mid(cmbprod.Text, Len(cmbprod.Text) - 10))
   End If
   Set rs = New ADODB.Recordset
   rs.Open "SELECT * FROM TFPRODUC,INVENTARIO, PREPROD WHERE TFPRODUC.CONSEC = INVENTARIO.INPROD AND TFPRODUC.CONSEC = PREPROD.PRECLAVE AND INVENTARIO.INPROD = PREPROD.PRECLAVE AND Consec = '" & cprod & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
   If Not (rs.EOF And rs.BOF) Then
      'If tipotienda = 2 Or tipotienda = 4 Then 'BODEGA
            txtcampos(5).Text = rs!InCant
            txtcampos(12).Text = IIf(IsNull(rs!InCantPza), 0, rs!InCantPza)
            txtcampos(6).Text = 0
            txtcampos(13).Text = 0
            txtcampos(14).Text = 0
            txtcampos(18).Text = Format(IIf(IsNull(rs!precio1), 0, rs!precio1), "#########0.00")               'Precio Pza. Autoservicio
            txtcampos(8).Text = Format(IIf(IsNull(rs!PRECIO2), 0, rs!PRECIO2), "#########0.00")                'Precio Caja credito
            txtcampos(15).Text = Format(IIf(IsNull(rs!PRECIO3), 0, rs!PRECIO3), "#########0.00")                'Precio Caja credito
            txtcampos(7).Text = Format(rs!precio4, "#########0.00")                                                                    'Precio Caja Contado
            txtcampos(16).Text = Format(IIf(IsNull(rs!precio5), 0, rs!precio5), "#########0.00")                'Precio Caja credito
            txtcampos(17).Text = Format(IIf(IsNull(rs!precio6), 0, rs!precio6), "#########0.00")                'Precio Caja credito
            
            txtcampos(10).Text = Format(IIf(IsNull(rs!precio1), 0, rs!precio1), "#########0.00")               'Precio Pza. Autoservicio
            txtcampos(11).Text = Format(IIf(IsNull(rs!precio4), 0, rs!precio4 / rs!PAQUETES), "#########0.00") 'Precio pieza bodega contado
            txtcampos(19).Text = Format(IIf(IsNull(rs!PRECIO3), 0, rs!PRECIO3 / rs!PAQUETES), "#########0.00") 'Precio pieza intermedio
            txtcampos(22).Text = Format(IIf(IsNull(rs!PRECIO2), 0, rs!PRECIO2 / rs!PAQUETES), "#########0.00") 'Precio pieza Credito
            If rs!precio5 And rs!precio6 > 0 Then
               pzasmay = Round(rs!PAQUETES / 2 + 0.1)
               txtcampos(20).Text = Format(rs!precio5 / pzasmay, "#########0.000")    'Precio pieza Medio mayoreo Bodega
               txtcampos(21).Text = Format(rs!precio6 / pzasmay, "#########0.000")    'Precio pieza Medio mayoreo Crédito
               txtcampos(14).Enabled = True
            Else
               txtcampos(20).Text = "0.00"
               txtcampos(21).Text = "0.00"
               txtcampos(14).Enabled = False
            End If
            'If chkVales.Value = 1 Then
            '   txtcampos(14).Enabled = False
            '   'Si son productos cotizados se da el precio especif.
            '   'En caso contrario Costo + 7.5 %
            '   If Not cotizados(cprod) Then
            '      txtcampos(18).Text = Format(RS!PRECOSTO / RS!PAQUETES * 1.075, "#########0.00")             'Precio Pza. Autoservicio
            '      txtcampos(11).Text = Format(RS!PRECOSTO / RS!PAQUETES * 1.075, "#########0.00")             'Precio Pza. Autoservicio
            '      txtcampos(10).Text = Format(RS!PRECOSTO / RS!PAQUETES * 1.075, "#########0.00")             'Precio Pza. Autoservicio
            '      txtcampos(7).Text = Format(RS!PRECOSTO * 1.075, "#########0.00")                            'Precio Caja Contado
            '   End If
            'End If
            CmdAgregar.Enabled = True
            cModifico = ""
            lbletiquetas(21).Caption = Str(rs!PAQUETES) & " X " & Str(rs!CONTENID) & " " & rs!medida
     'ElseIf tipotienda = 3 Then 'TIENDA
     '       txtcampos(5).Text = rs!InCant
     '       txtcampos(12).Text = IIf(IsNull(rs!InCantPza), 0, rs!InCantPza)
     '       txtcampos(6).Text = 0
     '       txtcampos(13).Text = 0
     '       txtcampos(7).Text = Format(rs!PRECIO2, "#########0.00")    'Precio Caja Contado
     '       txtcampos(8).Text = Format(IIf(IsNull(rs!PRECIO3), 0, rs!PRECIO3), "#########0.00")  'Precio Caja credito
     '       txtcampos(10).Text = Format(IIf(IsNull(rs!precio1), 0, rs!precio1), "#########0.00") 'Precio Pza. Autoservicio
     '       txtcampos(11).Text = Format(IIf(IsNull(rs!precio4), 0, rs!precio4), "#########0.00") 'Precio pieza contado
     '       CmdAgregar.Enabled = True
     '       cModifico = ""
     '       NIVA1 = rs!iva
     '       NIEPS1 = rs!ieps
     'Else
     '     MsgBox "Punto de Venta de Mayoreo esta Configurado...", vbInformation
     '     Exit Sub
     'End If
   Else
      MsgBox "EL PRODUCTO CON LA CLAVE " & cprod & " NO TIENE ESPECIFICADO EL PRECIO", vbCritical
   End If
End If
Exit Sub
Error:
  MsgBox "OCURRIO UN ERROR :" & Err.Description & Chr(13) & "A CONTINUACION SE CANCELARAN LOS CAMBIOS REALIZADOS"
End Sub

Private Function cotizados(clave As String) As Boolean
Dim CAD As String
cotizados = False
If Mid(cmbprod.Text, 1, Len("PASTA PARA SOPA ITAL-PASTA")) = "PASTA PARA SOPA ITAL-PASTA" Then
   txtcampos(18).Text = "1.66"  'Autoservicios
   txtcampos(7).Text = Format(Val(1.66) * rs!PAQUETES, "#########.00")
   txtcampos(10).Text = Format(Val(1.66), "#########.00")
   txtcampos(11).Text = Format(Val(1.66), "#########.00")
   cotizados = True
Else
   Select Case clave
   Case "1019992"
       txtcampos(18).Text = "5.81"  'Autoservicios
       txtcampos(7).Text = Format(Val(5.81) * rs!PAQUETES, "#########.00")
       txtcampos(10).Text = Format(5.81, "#########.00")
       txtcampos(11).Text = Format(5.81, "#########.00")
       cotizados = True
   Case "1008519"
       txtcampos(18).Text = "5.45"  'Autoservicios
       txtcampos(7).Text = Format(Val(5.45) * rs!PAQUETES, "#########.00")
       txtcampos(10).Text = Format(Val(5.45), "#########.00")
       txtcampos(11).Text = Format(Val(5.45), "#########.00")
       cotizados = True
   Case "1010646"
       txtcampos(18).Text = "9.15"  'Autoservicios
       txtcampos(7).Text = Format(Val(9.15) * rs!PAQUETES, "#########.00")
       txtcampos(10).Text = Format(Val(9.15), "#########.00")
       txtcampos(11).Text = Format(Val(9.15), "#########.00")
       cotizados = True
   Case "3000085"
       txtcampos(18).Text = "2.75"  'Autoservicios
       txtcampos(7).Text = Format(Val(2.75) * rs!PAQUETES, "#########.00")
       txtcampos(10).Text = Format(Val(2.75), "#########.00")
       txtcampos(11).Text = Format(Val(2.75), "#########.00")
       cotizados = True
   Case "3000055"
       txtcampos(18).Text = "26.00"  'Autoservicios
       txtcampos(7).Text = Format(Val(26) * rs!PAQUETES, "#########.00")
       txtcampos(10).Text = Format(Val(26), "#########.00")
       txtcampos(11).Text = Format(Val(26), "#########.00")
       cotizados = True
   Case "1005437"
       txtcampos(18).Text = "1.4"  'Autoservicios
       txtcampos(7).Text = Format(Val(1.4) * rs!PAQUETES, "#########.00")
       txtcampos(10).Text = Format(Val(1.4), "#########.00")
       txtcampos(11).Text = Format(Val(1.4), "#########.00")
       cotizados = True
   Case "1004823"
       txtcampos(18).Text = "3.75"  'Autoservicios
       txtcampos(7).Text = Format(Val(3.75) * rs!PAQUETES, "#########.00")
       txtcampos(10).Text = Format(Val(3.75), "#########.00")
       txtcampos(11).Text = Format(Val(3.75), "#########.00")
       cotizados = True
   Case "3000509"
       txtcampos(18).Text = "9.20"  'Autoservicios
       txtcampos(7).Text = Format(Val(9.2) * rs!PAQUETES, "#########.00")
       txtcampos(10).Text = Format(Val(9.2), "#########.00")
       txtcampos(11).Text = Format(Val(9.2), "#########.00")
       cotizados = True
   Case "3500641"
       txtcampos(18).Text = "4.06"  'Autoservicios
       txtcampos(7).Text = Format(Val(4.06) * rs!PAQUETES, "#########.00")
       txtcampos(10).Text = Format(Val(4.06), "#########.00")
       txtcampos(11).Text = Format(Val(4.06), "#########.00")
       cotizados = True
   End Select
End If
End Function

Private Sub cmdModificar_Click()
Dim rsttemp As New ADODB.Recordset
'Cargo la venta
AdoVentas.RecordSource = "SELECT * FROM VENTAS WHERE noventa = " & adopreventa.Recordset!noventa
AdoVentas.Refresh
'Cargo el detalle de la venta
AdoDetVta.CursorType = adOpenStatic
AdoDetVta.ConnectionString = cCadConex
If Sql Then
   AdoDetVta.RecordSource = "SELECT cl_producto, Cantidad, cantidadp, Descripc , LTRIM(RTRIM(STR(PAQUETES))) + ' X ' + LTRIM(STR(CONTENID,10,3)) + ' ' + MEDIDA AS MEDIDA, Precio, importe, serie + ' ' + factura AS factura FROM VENTAS_DET,TFPRODUC WHERE cl_producto = Consec AND NoVenta = '" & adopreventa.Recordset!noventa & "' ORDER BY DESCRIPC"
Else
   AdoDetVta.RecordSource = "SELECT cl_producto, Cantidad, cantidadp, Descripc ,LTRIM(STR(T.paquetes)) + ' X ' + LTRIM(STR(CONTENID)) + space(2) + t.medida AS MEDIDA, Precio, importe  FROM VENTAS_DET,TFPRODUC T WHERE cl_producto = T.Consec AND NoVenta = " & AdoVentas.Recordset!noventa & " ORDER BY T.DESCRIPC"
End If
AdoDetVta.Refresh
    Set rsttemp = New ADODB.Recordset
    'rsttemp.Open "SELECT SUM(Importe) AS Subto, SUM( CASE CLAPROVE WHEN 'C52' THEN IMPORTE END) AS Promo FROM VENTAS_DET, TFPRODUC WHERE consec = cl_producto AND NOVENTA = " & AdoVentas.Recordset!noventa, cn, adOpenForwardOnly, adLockOptimistic, adCmdText
    rsttemp.Open "SELECT SUM(Importe) AS Subto FROM VENTAS_DET, TFPRODUC WHERE consec = cl_producto AND NOVENTA = " & AdoVentas.Recordset!noventa, cn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If IsNull(rsttemp!subto) Then
       lblImpVta.Caption = "$ 0.00"
       'lblpromo.Caption = "PROM:  $ 0.00    BOLETOS: 0"
       nImpVta = 0
    Else
       lblImpVta.Caption = Format(rsttemp!subto, "$ ###,###,##0.00")
       'lblpromo.Caption = IIf(IsNull(rsttemp!PROMO), "PROM:  $ 0.00    BOLETOS: 0", "PROM.:  " & Format(rsttemp!PROMO, "$ ###,###,##0.00") & "      BOLETOS: " & Int(rsttemp!PROMO / 50))
       nImpVta = rsttemp!subto
    End If
    cn.Execute "UPDATE ventas SET MontoTotal = " & IIf(IsNull(rsttemp!subto), 0, rsttemp!subto) & " WHERE Noventa = " & AdoVentas.Recordset!noventa

    rsttemp.Close
    Set rsttemp = Nothing
    cmdRegVta_Click
    ccVeCli = txtcampos(4).Text
    If rscli.State = adStateOpen Then rscli.Close
    rscli.Open "SELECT * FROM Catcliente WHERE cClave = " & ccVeCli & " AND ctipo = 0", cn, adOpenKeyset, adLockOptimistic, adCmdText
    'If Sql Then
       If rscli!ccredito And chkCredito.Value = 1 Then vercredito (ccVeCli)
    'End If
    cmbCliente.Text = rscli!cNombre
    lblNombre.Caption = IIf(IsNull(rscli!cnombrefac), "", rscli!cnombrefac)
    lblDirec.Caption = IIf(IsNull(rscli!cdireccion), "", rscli!cdireccion)
    lblRfc.Caption = IIf(IsNull(rscli!crfc), "", rscli!crfc)
    lblTelefono.Caption = IIf(IsNull(rscli!ctelefono), 0, rscli!ctelefono)
    fraGenerales.Caption = "Datos generales de la venta   [ " & adopreventa.Recordset!noventa & " ]"
    Me.DbgDetVta.SetFocus
End Sub

Private Function vercredito(CLIENTE As Integer) As Boolean
Dim rsvencre As ADODB.Recordset

vercredito = False
Set rsvencre = New ADODB.Recordset
'rsvencre.Open "SELECT SUM( CASE WHEN Facfecha <= '" & date - (rscli!ctiempocredito + 2) & "' THEN porpagar END) AS vencido, SUM(porpagar) AS ejercido FROM facventa WHERE facfecha >= '01/01/2002' And cobrado = 0 AND cancelado = 0 AND Faccliente = '" & CLIENTE & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
rsvencre.Open "SELECT serie,Numfactura,porpagar,facfecha FROM facventa WHERE YEAR(facfecha) >= 2002 And cobrado = 0 AND cancelado = 0 AND Faccliente = " & CLIENTE & " ORDER BY facfecha", cn, adOpenKeyset, adLockOptimistic, adCmdText
nDispo = 0: Ejercido = 0: VENCIDO = 0
cMensaje = "SITUACION ACTUAL DEL CREDITO" & Chr(13) & Chr(13) & "PLAZO EN DIAS: " & rscli!ctiempocredito & Chr(13) & "MONTO AUTORIZADO:             " & Format(rscli!CLIMITECREDITO, "$ ###,###,##0.00") & Chr(13)
While Not rsvencre.EOF
   If rsvencre!FACFECHA <= date - (rscli!ctiempocredito + 2) Then
      VENCIDO = VENCIDO + rsvencre!porpagar
   End If
   Ejercido = Ejercido + rsvencre!porpagar
   cMensaje = cMensaje & Chr(13) & Trim(rsvencre!SERIE) & "-" & rsvencre!numfactura & Space(5) & rsvencre!FACFECHA & Space(5) & Format(rsvencre!porpagar, "$ ###,###,##0.00")
   rsvencre.MoveNext
Wend
cMensaje = cMensaje & Chr(13) & Chr(13) & "EJERCIDO  :" & Space(5) & Format(Ejercido, "$ ###,###,##0.00") & Chr(13) & Chr(13) & "VENCIDO      :" & Space(5) & Format(VENCIDO, "$ ###,###,##0.00") & Chr(13) & Chr(13) & "DISPONIBLE: " & Space(5) & Format(rscli!CLIMITECREDITO - VENCIDO, "$ ###,###,##0.00")
MsgBox cMensaje, vbInformation, "Ventas"

If VENCIDO >= rscli!CLIMITECREDITO Then
   vercredito = True
   nDispo = 1
ElseIf Ejercido > 0 Then
   nDispo = rscli!CLIMITECREDITO - IIf(IsNull(VENCIDO), 0, VENCIDO)
End If
rsvencre.Close
Set rsvencre = Nothing
End Function

Public Sub cmdRegresar_Click()
On Error GoTo Error:
Unload Me
Exit Sub
Error:
  Unload Me
End Sub

Private Sub cmdRegVta_Click()
  frapreventa.Visible = False
End Sub

Private Sub cmdReporte_Click()
Dim rs As ADODB.Recordset
crpt.Connect = cCadConex
If chkConcent.Value = 1 Then
   LETRAS = NumLet$(Format(TxtTotVta.Text, "#########0.00"))
   'crpt.ReportFileName = App.Path & IIf(lbletiquetas(2).BorderStyle = 0, "\ConAgte.rpt", "\ConAgtee.rpt")
   crpt.ReportFileName = App.Path & "\ConAgte.rpt"
   crpt.WindowTitle = "Reporte concentrado del agente " & cmbAgente.Text
   crpt.Formulas(1) = "ENCAB = ' VENTAS DEL AGENTE " & Trim(cmbAgente.Text) & " EN EL FOLIO DE PREVENTA " & txtcampos(2).Text & "'"
   crpt.Formulas(2) = "CHOFER = '" & Trim(cmbChofer.Text) & "'"
   crpt.Formulas(3) = "IMPLET = '" & LETRAS & "'"
   If ZONA = "OAX" Then
      crpt.Formulas(4) = "FECHA = 'Oaxaca de Juarez, Oaxaca a " & Format(date, "LONG DATE") & "'"
   Else
      crpt.Formulas(4) = "FECHA = 'Tapachula, Chiapas a " & Format(date, "LONG DATE") & "'"
   End If
   OrigenDatos = "SELECT VENTAS.noventa, VENTAS.folpreventa, " & _
                         "VENTAS_DET.cantidad, VENTAS_DET.cantidadp, VENTAS_DET.precio, VENTAS_DET.importe, VENTAS_DET.cancelado, " & _
                         "TFPRODUC.DESCRIPC, TFPRODUC.CONTENID, TFPRODUC.MEDIDA, TFPRODUC.PAQUETES, TFPRODUC.CONSEC " & _
                 "FROM pitico.dbo.VENTAS VENTAS, " & _
                        "pitico.dbo.VENTAS_DET VENTAS_DET, " & _
                        "pitico.dbo.TFPRODUC TFPRODUC " & _
                 "WHERE VENTAS.noventa = VENTAS_DET.noventa AND " & _
                        "VENTAS_DET.cl_producto = TFPRODUC.CONSEC AND " & _
                        "VENTAS.folpreventa = " & AdoVentas.Recordset!FOLPREVENTA & " AND VENTAS_DET.cancelado = 0 "
   If Sql Then  'Se actualizan las preventas
      cn.Execute "UPDATE ventas SET situacion = '1' WHERE situacion = '0' AND folpreventa = " & AdoVentas.Recordset!FOLPREVENTA
      Set rs = New ADODB.Recordset
      rs.Open "SELECT * FROM preventas WHERE folio = " & txtcampos(2).Text, cn, adOpenForwardOnly, adLockOptimistic, adCmdText
'      MsgBox "INSERT INTO preventas(folio,status,impcapt) VALUES (" & txtcampos(2).Text & ",1," & Format(TxtTotVta.Text, "#########.##") & ")"
      If rs.BOF And rs.EOF Then cn.Execute "INSERT INTO preventas(folio,status,impcapt) VALUES (" & txtcampos(2).Text & ",1," & Format(TxtTotVta.Text, "#########.##") & ")"
   End If
Else
   crpt.ReportFileName = App.Path & "\ConAgteD.rpt"
   crpt.WindowTitle = "Reporte detallado del agente " & cmbAgente.Text
   crpt.Formulas(1) = "AGENTE = '" & Trim(cmbAgente.Text) & Space(10) & "'"
   crpt.Formulas(2) = "CHOFER = '" & Trim(cmbChofer.Text) & "'"
   crpt.Formulas(3) = ""
   crpt.Formulas(4) = ""
   OrigenDatos = "SELECT VENTAS.noventa, VENTAS.fecha, VENTAS.clcliente, VENTAS.facrfc, VENTAS.folpreventa, " & _
                        "VENTAS_DET.cantidad, VENTAS_DET.cantidadp, VENTAS_DET.precio, VENTAS_DET.importe, VENTAS_DET.cancelado, " & _
                        "CATCLIENTE.cnombre, CATCLIENTE.cdireccion, CATCLIENTE.CColonia, CATCLIENTE.CCiudad, " & _
                        "TFPRODUC.DESCRIPC, TFPRODUC.CONTENID, TFPRODUC.MEDIDA, TFPRODUC.PAQUETES, TFPRODUC.CONSEC " & _
                 "FROM pitico.dbo.VENTAS VENTAS, " & _
                       "pitico.dbo.VENTAS_DET VENTAS_DET, " & _
                       "pitico.dbo.CATCLIENTE CATCLIENTE, " & _
                       "pitico.dbo.TFPRODUC TFPRODUC " & _
                 "WHERE VENTAS.noventa = VENTAS_DET.noventa AND " & _
                        "VENTAS.clcliente = CATCLIENTE.cclave AND " & _
                        "VENTAS_DET.cl_producto = TFPRODUC.CONSEC AND " & _
                        "VENTAS.folpreventa = " & AdoVentas.Recordset!FOLPREVENTA & " AND VENTAS_DET.cancelado = 0 " & Chr(13) & _
                 "ORDER BY VENTAS.noventa ASC"
End If
'MsgBox ORIGENDATOS
crpt.SQLQuery = OrigenDatos
crpt.Action = 1
End Sub

Private Sub cmdTicBod_Click()
Dim rsttemp As ADODB.Recordset
Dim nTotal
Dim nCajas
Dim ProEnTick
Dim nTicket As Integer
Dim nTotTick As Integer
Dim cCad
Dim Promocion As Currency
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

'lResumen = False  'Por default se pone la venta por depto por ticket y no por envio
On Error GoTo Error:
lImp = False
For Each x In Printers
   If x.DeviceName Like "*TICKET*" Then
      lImp = True
      Set Printer = x
      Exit For
   End If
Next x
If lImp = False Then
   If MsgBox("NO ES POSIBLE IMPRIMIR TICKET'S PARA SURTIR" & Chr(13) & Chr(13) & "PORQUE NO EXISTEN IMPRESORAS INSTALADAS" & Chr(13) & "PARA ESTE TIPO DE TICKET'S, DESEAS ENVIARLO A LA IMPRESORA PREDETERMINADA", vbCritical + vbYesNo) = vbNo Then Exit Sub
End If

nAncho = 250    'En puntos es el ancho del ticket de las Miniprinter
ProEnTick = 20  'Numero de productos que se imprimen en el ticket
Printer.ScaleMode = vbPoints

Set rsttemp = New ADODB.Recordset
rsttemp.Open "SELECT COUNT(*) AS TOTPRO FROM Ventas_det WHERE noventa = '" & frmVentas.AdoVentas.Recordset!noventa & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText

nTotTick = IIf(Round(rsttemp!totpro / ProEnTick, 0) < rsttemp!totpro / ProEnTick, Round(rsttemp!totpro / ProEnTick, 0) + 1, Round(rsttemp!totpro / ProEnTick, 0))

rsttemp.Close
rsttemp.Open "SELECT * FROM Ventas_det, Tfproduc WHERE  cl_producto = TFPRODUC.Consec AND NoVenta = '" & frmVentas.AdoVentas.Recordset!noventa & "' ORDER BY Descripc,contenid", cn, adOpenKeyset, adLockOptimistic, adCmdText

nTotal = 0: nCajas = 0: nPiezas = 0: nProd = 0: nPeso = 0: Promocion = 0
lNoImp = True: nProd = ProEnTick: nTicket = 0

'Cuando esta vacio el detalle del traslado para que se imprima el encabezado
If rsttemp.BOF And rsttemp.EOF Then
   MsgBox "NO SE PUEDE IMPRIMIR EL TICKET PORQUE NO SE HA" & Chr(13) & "CAPTURADO NINGUN PRODUCTO EN ESTA VENTA", vbInformation
   Exit Sub
   'Encabezado nTicket, nTotTick  'Imprime encabezado
End If
Do While (Not rsttemp.EOF)
    If nProd = ProEnTick Then
       If Not lNoImp Then
          Printer.Print "-------------------------------------"
          Printer.Print ""
          Printer.Print "IMPORTE TOTAL :     " & Format(nTotal, "$###,###,##0.00")
          Printer.Print " "
          'Printer.Print "*P = PROMOCION NESTLE: " & Format(Promocion, "$###,###,##0.00")
          'Printer.Print "BOLETOS DE PROMOCION : "; Int(Promocion / 50)
          'Printer.Print " "
          Printer.Print "TOTAL CAJAS : " & Format(nCajas, "#,###,##0")
          Printer.Print "TOTAL PIEZAS: " & Format(nPiezas, "#,###,##0")
          Printer.Print "TOTAL KILOS : " & Format(nPeso, "#,###,##0.00")
          Printer.Print " "
          Printer.Print " "
          Printer.Print " "
          Printer.Print " "
          Printer.Print "            ------------------          "
          Printer.Print "                   SURTE                "
          Printer.Print " "
          Printer.Print "====================================="
          Printer.Print " "
          Printer.Print " "
          'If lFran Then Vtadepto nVENTA1, nVENTA2, nIVA2, NVENTA3, NIVA3, NIEPS3, NVENTA4, NIVA4, NIEPS4
          For N = 0 To 11
              Printer.Print " "
          Next
          Printer.EndDoc
          'Handle = Shell(App.Path & "\CORTA.EXE", 1)
          'MsgBox "CUANDO TERMINE DE IMPRIMIR EL TICKET " & CStr(nTicket) & Chr(13) & "PRESIONE ENTER PARA CORTAR EL PAPEL", vbInformation
          'CmdCortar_Click
       End If
       nTicket = nTicket + 1
       cresp = 6
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
       nTotal = 0: nCajas = 0: nPiezas = 0: nProd = 0: nPeso = 0: Promocion = 0
       nProd = 0
    End If
    'Si tiene cantidad en cajas y se ha deseado imprimir el ticket
    If Not lNoImp And (rsttemp!cantidad > 0 Or rsttemp!cantidadp > 0) Then
        cCad = " " & CStr(rsttemp!cantidad) & " CAJ, " & CStr(rsttemp!cantidadp) & " PZA " & Trim(rsttemp!descripc)
        'En caso de que sea muy grande la descripcion se imprime en dos lineas
        If Len(Trim(cCad)) > 37 Then
           Printer.Print Mid(cCad, 1, 37)
           If nProd = ProEnTick - 1 Then Printer.Print " "  'Quiensabe porque se encimaba en el ultimo producto en algunas impresoras
           Printer.Print Mid(cCad, 38, 24);
        Else
           Printer.Print cCad
        End If
        
        If nProd = ProEnTick - 1 Then Printer.Print " "  'Quiensabe porque se encimaba en el ultimo producto en algunas impresoras
        'If Trim(rsttemp!claprove) = "C52" Then
        '   Promocion = Promocion + rsttemp!importe
        '   Printer.Print Space(5) & "*P";
        'End If
        Printer.CurrentX = 172
        Printer.Print "  " & CStr(rsttemp!PAQUETES) & " X " & CStr(rsttemp!CONTENID) & " " & rsttemp!medida
        If nProd = ProEnTick - 1 Then Printer.Print " "  'Quiensabe porque se encimaba en el ultimo producto en algunas impresoras
        
        N = nAncho - (Printer.TextWidth(Format(rsttemp!PRECIO, "###,###,##0.00") & Space(6) & Format(rsttemp!importe, "$###,###,##0.00")))
        Printer.CurrentX = N
        Printer.Print Format(rsttemp!PRECIO, "###,###,##0.00") & Space(6) & Format(rsttemp!importe, "$###,###,##0.00")
        
        If Not IsNull(rsttemp!PRECIO) Then  'Verifico costo
           ncosto = rsttemp!importe
           nTotal = nTotal + ncosto
        End If
        If Not IsNull(rsttemp!peso) Then  'Verifico peso
           nPeso = nPeso + (rsttemp!peso * rsttemp!cantidad)
        End If
    End If
    nCajas = nCajas + rsttemp!cantidad
    nPiezas = nPiezas + rsttemp!cantidadp
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
Printer.Print "-------------------------------------"
'Printer.Print ""
Printer.Print "IMPORTE TOTAL :     " & Format(nTotal, "$###,###,##0.00")
Printer.Print " "
'Printer.Print "*P = PROMOCION NESTLE: " & Format(Promocion, "$###,###,##0.00")
'Printer.Print "BOLETOS DE PROMOCION : "; Int(Promocion / 50)
'Printer.Print " "
Printer.Print "TOTAL CAJAS : " & Format(nCajas, "#,###,##0")
Printer.Print "TOTAL PIEZAS: " & Format(nPiezas, "#,###,##0")
Printer.Print "TOTAL KILOS : " & Format(nPeso, "#,###,##0.00")
Printer.Print " "
Printer.Print " "
Printer.Print " "
Printer.Print " "
Printer.Print "            ------------------          "
Printer.Print "                   SURTE                "
Printer.Print " "
Printer.Print "====================================="
Printer.Print " "
Printer.Print " "
For N = 0 To 11
    Printer.Print " "
Next
cmdregresar.SetFocus
'Printer.Print Chr(27) + Chr(64)
'Printer.Print Chr(27) + Chr(105)
Printer.EndDoc
If AdoVentas.Recordset!situacion = "0" Then
   AdoVentas.Recordset!situacion = "1"
   AdoVentas.Recordset.Update
   AdoVentas.Refresh
End If
MsgBox "EL TICKET SE IMPRIMIO CORRECTAMENTE EN EL AREA DE BODEGA" & Chr(13) & "Y EN ESTOS MOMENTOS INICIA EL SURTIDO DEL PEDIDO", vbInformation, "Ventas"
cmdregresar.SetFocus
Exit Sub
Error:
  MsgBox Err.Description
End Sub

Private Sub Encabezado(nTick As Integer, nTTick As Integer)
Dim RST As ADODB.Recordset
Set RST = New ADODB.Recordset
  RST.Open "SELECT * FROM cattienda WHERE ticlave = '" & Trim(Mid(cSucursal, 1, 3)) & "'", cn, adOpenForwardOnly, adLockOptimistic, adCmdText
  Printer.Print " "
  If ZONA = "OAX" Then
    Printer.Print "     VIVERES Y LICORES S.A DE C.V.    "
  Else
    Printer.Print "HOLDING MEXICO CENTRO AMERICA SA DE CV"
  End If
  'Printer.Print "     MIGUEL CABRERA # 603  CENTRO     "
  Printer.Print Space((40 - Len(Mid(RST!Direccion, 1, 40))) / 2) & RST!Direccion & Space(5);
  Printer.Print date & "  " & Format(Time, "HH:MM:SS")
  Printer.Print "VENTA A: "; frmVentas.cmbCliente.Text
  Printer.Print "FOL.UNICO: "; frmVentas.AdoVentas.Recordset!noventa & "   " & "FOL.VTA.DIA "; ; frmVentas.AdoVentas.Recordset!folioventa
  Printer.Print "TICKET NUM: "; CStr(nTick) & " DE " & CStr(nTTick)
  Printer.Print "COMPUTADORA: " & UCase(AdoVentas.Recordset!CL_TERMINAL)
  Printer.Print "-------------------------------------"
  RST.Close
  Set RST = Nothing
End Sub

Static Sub cmdTicket_Click()
todobien = FacturaNva
If todobien Then
   If chkCredito.Value = 0 And chkPrev.Value = 0 Then
      CAD = "update ventas set situacion =  2  where noventa = " & noventar
   Else
      CAD = "update ventas set situacion =  3  where noventa = " & noventar
   End If
   cn.Execute CAD
   MsgBox "Impresión de Factura Finalizada....", vbInformation
End If
End Sub


Private Sub pgrabafactura(Optional globconFin As Boolean)
Dim cobrado
'se inserta la factura
CLIENTE = IIf(globconFin = True, txtcampos(9).Text, txtcampos(4).Text)
Confin = IIf(globconFin = True, 1, 0)
'FALTA INSERTAR EL MONTO DE LA FACTURA, SE HARA EN EL MOMENTO DE REALIZAR EL CORTE CORRESPONDIENTE
cobrado = IIf(chkCredito.Value = 1 Or chkPrev.Value = 1, "0", "1")
CAD = "INSERT INTO facventa(facCliente,Noventa,FacFecha,serie,NumFactura,Cobrado,globconfin,rfc) VALUES (" & CLIENTE & "," & noventar & ",'" & Format(date, "dd/mm/yyyy") & "','" & Trim(SERIE) & "','" & Trim(numfac) & "'," & cobrado & "," & Confin & ",'" & rfct & "')"
cn.Execute CAD
End Sub

Function Factura(Optional globconFin As Boolean) As Boolean
'Function Factura(Optional GlobConFin As Boolean) As Boolean
Dim rstDetVta As ADODB.Recordset
Dim rstDir As ADODB.Recordset
Dim Impresora As Printer
Dim nTotal
Dim nCajas
Dim lNvaFac As Boolean
Dim nAlto
Dim rsFac As ADODB.Recordset
Dim TotFac As Double
Dim LETRAS

On Error GoTo Error:
Factura = False
'lImp = False
cimpresora = IIf(AdoVentas.Recordset!credito, "*CREDITO*", "*CONTADO*")
'For Each x In Printers
   'If UCase(x.DeviceName) Like cimpresora Then
      'lImp = True
      'MsgBox "PREPARE LA IMPRESORA " & x.DeviceName, vbInformation
      'Set Printer = x
      'Exit For
   'End If
'Next x
'If lImp = False Then
   'MsgBox "NO ES POSIBLE IMPRIMIR FACTURAS A " & cimpresora & Chr(13) & Chr(13) & "PORQUE NO EXISTEN IMPRESORAS INSTALADAS" & Chr(13) & "PARA ESTE TIPO DE FACTURAS", vbCritical
   'Exit Function
'End If
lNvaFac = True: nProd = 0
nAlto = -1.1
ProdenFac = 14
SERIE = IIf(Trim(Mid(cSucursal, 1, 3)) = "16", "Y1", "D")

Printer.ScaleMode = vbCentimeters
Printer.FontName = "ARIAL NARROW"
Printer.FontSize = 8
Printer.Width = 12190
Printer.Height = 7938
Set rstDetVta = New ADODB.Recordset
If globconFin = True Then
   rstDetVta.Open "SELECT MIN(CONSEC) as consec, sum(cantidad) as cantidad, sum(cantidadp) AS cantidadp, descripc, str(paquetes) + ' X ' + LTRIM(STR(CONTENID,10,3)) + ' ' + MEDIDA AS MEDIDA, MAX(D.IEPS) AS IEPS, MAX(D.IVA) AS IVA, AVG(D.PRECIO) AS PRECIO, SUM(D.IMPORTE) AS IMPORTE FROM VENTAS_DET D, TFPRODUC,VENTAS V WHERE V.folpreventa = " & AdoVentas.Recordset!FOLPREVENTA & " AND D.Cl_producto = Consec AND d.Cancelado = 0 AND V.FACRFC = 0 AND D.NOVENTA = V.NOVENTA AND V.situacion = '1' GROUP BY str(paquetes) + ' X ' + LTRIM(STR(CONTENID,10,3)) + ' ' +  MEDIDA,DESCRIPC", cn, adOpenStatic, adLockOptimistic, adCmdText
Else
   If nOp = 0 Or nOp = 1 Then
      rstDetVta.Open "SELECT CONSEC = '', cantidad, cantidadp = 0, concepto as descripc, MEDIDA = '', Iva, Ieps = 0, Precio, Importe FROM FacSinAnt WHERE Noventa = " & AdoVentas.Recordset!noventa & " ORDER BY DESCRIPC", cn, adOpenStatic, adLockOptimistic, adCmdText
   Else
      rstDetVta.Open "SELECT P.CONSEC, D.cantidad, D.cantidadp, P.descripc, str(P.paquetes) + ' X ' + LTRIM(STR(P.CONTENID,10,3)) + ' ' + P.MEDIDA AS MEDIDA, D.Iva, D.Ieps, D.Precio, D.Importe FROM VENTAS_DET D, TFPRODUC P WHERE Noventa = " & AdoVentas.Recordset!noventa & " AND Cl_producto = Consec AND D.FACTURA IS NULL ORDER BY DESCRIPC,CONTENID ", cn, adOpenStatic, adLockOptimistic, adCmdText
   End If
End If

cresp = MsgBox("DESEAS GENERAR LA FACTURACON FECHA DE LA VENTA" & Chr(13) & "[SI] = Genera factura con fecha de venta" & Chr(13) & "[NO] = Genera factura con fecha actual del sistema", vbYesNoCancel + vbQuestion)
If cresp = vbCancel Then
   Exit Function
End If

While Not rstDetVta.EOF
   If lNvaFac Then
      Do
        numfac = InputBox("PREPARE LA IMPRESORA Y ACOMODE EL PAPEL" & Chr(13) & Chr(13) & "Proporcione numero de factura", "Proporcione factura")
        nPos = InStr(1, numfac, "-")
        If IsNull(numfac) Or Trim(numfac) = "" Then
           If MsgBox("ES NECESARIO ESPECIFICAR NUMERO DE FACTURA" & Chr(13) & "DESEAS CONTINUAR CON LA IMPRESION DE LA FACTURA", vbYesNo + vbQuestion) = vbNo Then
              Printer.KillDoc
              Exit Function
           End If
        Else
            Set rsFac = New ADODB.Recordset
            rsFac.Open "SELECT * FROM FACVENTA WHERE SERIE = '" & IIf(frmVentas.AdoVentas.Recordset!credito, "B", SERIE) & "' AND NumFactura = '" & Trim(numfac) & "'", cn, adOpenKeyset, adLockBatchOptimistic, adCmdText
            If rsFac.RecordCount > 0 Then
               MsgBox "LA SERIE Y EL NUMERO DE FACTURA YA SE IMPRIMIO " & Chr(13) & "EN LA VENTA CON FOLIO UNICO " & rsFac!noventa & " DE FECHA " & rsFac!FACFECHA, vbExclamation
            Else
           
               cn.BeginTrans: lTrans = True
               CLIENTE = IIf(globconFin = True, 2, txtcampos(4).Text)
               Confin = IIf(globconFin = True, 1, 0)
               If cresp = vbYes Then
                  cn.Execute "INSERT INTO facventa(facCliente,Noventa,FacFecha,serie,NumFactura,Cobrado,globconfin) VALUES ('" & CLIENTE & "'," & AdoVentas.Recordset!noventa & ",'" & AdoVentas.Recordset!fecha & "','" & IIf(frmVentas.AdoVentas.Recordset!credito, "B", SERIE) & "','" & Trim(numfac) & "'," & IIf(AdoVentas.Recordset!credito, 0, 1) & "," & Confin & ")"
               Else
                  cn.Execute "INSERT INTO facventa(facCliente,Noventa,FacFecha,serie,NumFactura,Cobrado,globconfin) VALUES ('" & CLIENTE & "'," & AdoVentas.Recordset!noventa & ",'" & date + Time & "','" & IIf(frmVentas.AdoVentas.Recordset!credito, "B", SERIE) & "','" & Trim(numfac) & "'," & IIf(AdoVentas.Recordset!credito, 0, 1) & "," & Confin & ")"
               End If
               nAlto = -1.1
               'nAlto = IIf(Trim(Mid(cSucursal, 1, 3)) = "16", -1.1, -0.6)
               Exit Do
            End If
        End If
      Loop
      Printer.CurrentY = 3.5 + nAlto
      Printer.CurrentX = 0.5
      rscli.Requery
      Printer.Print IIf(AdoVentas.Recordset!facrfc = False Or globconFin = True, "C O N S U M I D O R   F I N A L", rscli!cnombrefac);
      Printer.CurrentX = 10
      Printer.Print IIf(AdoVentas.Recordset!facrfc = False Or globconFin = True, "COOF970101111", rscli!crfc)
      Printer.CurrentX = 0.5
      Printer.Print IIf(AdoVentas.Recordset!facrfc = False Or globconFin = True, "C O N O C I D O", rscli!cdireccion);
      Printer.CurrentX = 10
      Printer.Print rscli!ctelefono
      Printer.CurrentX = 0.5
      Printer.Print IIf(IsNull(rscli!ccolonia) Or AdoVentas.Recordset!facrfc = False, "", rscli!ccolonia);
      Printer.CurrentX = 10
      Printer.Print IIf(IsNull(rscli!cciudad) Or AdoVentas.Recordset!facrfc = False, "", rscli!cciudad);
      Printer.CurrentX = 18
      Printer.Print IIf(frmVentas.AdoVentas.Recordset!credito, "B", SERIE) & " " & numfac
      
      Printer.CurrentY = 5 + nAlto
      
      If Trim(cmbAgente) <> "" Then
         Printer.CurrentX = 9
         Printer.Print "AGENTE: " & Mid(cmbAgente.Text, 1, InStr(1, cmbAgente.Text, " "));
      End If
      If Trim(cmbChofer) <> "" Then
         Printer.CurrentX = 13
         Printer.Print "CHOFER: " & Mid(cmbChofer.Text, 1, InStr(1, cmbChofer.Text, " "));
      End If
      
      Printer.CurrentX = 16.5
      If cresp = vbYes Then
         Printer.Print Format(AdoVentas.Recordset!fecha, "long date") & Space(3) & Format(Time, "HH:MM AM/PM")
      Else
         Printer.Print Format(date, "long date") & Space(3) & Format(Time, "HH:MM AM/PM")
      End If

      
      NVENTA1 = 0: NIVA1 = 0: NIEPS1 = 0
      NVENTA2 = 0: NIVA2 = 0: NIEPS2 = 0
      NVENTA3 = 0: NIVA3 = 0: NIEPS3 = 0
      NVENTA4 = 0: NIVA4 = 0: NIEPS4 = 0
      lNvaFac = False
      Printer.CurrentY = 6.5 + nAlto

   End If
   If nProd > ProdenFac Then  'Se incluyen 15 productos por factura
        Printer.CurrentY = 11.5 + nAlto
        Printer.CurrentX = 2
        Printer.Print "DEP1";
        Printer.CurrentX = 4
        Printer.Print "DEP2";
        Printer.CurrentX = 6
        Printer.Print "DEP3";
        Printer.CurrentX = 8
        Printer.Print "DEP4"

        'SUBTOTALES
        'Cuando es consumidor final no se desglosa la factura ni se imprime ieps e iva
        Printer.CurrentX = 2
        If NVENTA1 > 0 Then Printer.Print Format(NVENTA1, "########0.00");
        Printer.CurrentX = 4
        If NVENTA2 > 0 Then
           Printer.Print IIf(AdoVentas.Recordset!facrfc = False Or globconFin = True, Format(NVENTA2, "########0.00"), Format(NVENTA2 - NIVA2, "########0.00"));
        End If
        Printer.CurrentX = 6
        If NVENTA3 > 0 Then
           Printer.Print IIf(AdoVentas.Recordset!facrfc = False Or globconFin = True, Format(NVENTA3, "########0.00"), Format(NVENTA3 - NIVA3 - NIEPS3, "########0.00"));
        End If
        Printer.CurrentX = 8
        If NVENTA4 > 0 Then
           Printer.Print IIf(AdoVentas.Recordset!facrfc = False Or globconFin = True, Format(NVENTA4, "########0.00"), Format(NVENTA4 - NIVA4 - NIEPS4, "########0.00"));
        End If
        Printer.CurrentX = 16
        Printer.Print "SUBTOTAL";
        Printer.CurrentX = 18.5
        Printer.FontSize = 10
        
        If AdoVentas.Recordset!facrfc = False Or globconFin = True Then
           Printer.Print String(10 - Len(Trim(Format(NVENTA1 + NVENTA2 + (NVENTA3 + NVENTA4), "#########0.00"))), " ") & Format(NVENTA1 + NVENTA2 + NVENTA3 + NVENTA4, "#########0.00")
        Else
           Printer.Print String(10 - Len(Trim(Format(NVENTA1 + (NVENTA2 - NIVA2) + (NVENTA3 - NIVA3 - NIEPS3) + (NVENTA4 - NIVA4 - NIEPS4), "#########0.00"))), " ") & Format(NVENTA1 + (NVENTA2 - NIVA2) + (NVENTA3 - NIVA3 - NIEPS3) + (NVENTA4 - NIVA4 - NIEPS4), "#########0.00")
        End If
        Printer.FontSize = 8
        
        'IVA
        Printer.CurrentX = 2
        If NVENTA1 > 0 And Not (AdoVentas.Recordset!facrfc = False Or globconFin = True) Then Printer.Print Format(NIVA1, "########0.00");
        Printer.CurrentX = 4
        If NVENTA2 > 0 And Not (AdoVentas.Recordset!facrfc = False Or globconFin = True) Then Printer.Print Format(NIVA2, "########0.00");
        Printer.CurrentX = 6
        If NVENTA3 > 0 And Not (AdoVentas.Recordset!facrfc = False Or globconFin = True) Then Printer.Print Format(NIVA3, "########0.00");
        Printer.CurrentX = 8
        If NVENTA4 > 0 And Not (AdoVentas.Recordset!facrfc = False Or globconFin = True) Then Printer.Print Format(NIVA4, "########0.00");
        Printer.CurrentX = 16
        Printer.Print "IVA";
        Printer.CurrentX = 18.5
        Printer.FontSize = 10
        If AdoVentas.Recordset!facrfc = False Or globconFin = True Then
            Printer.Print String(10 - Len("00"), " ") & "0.00"
        Else
            Printer.Print String(10 - Len(Trim(Format(NIVA2 + NIVA3 + NIVA4, "#########0.00"))), " ") & Format(NIVA2 + NIVA3 + NIVA4, "#########0.00")
        End If
        Printer.FontSize = 8
        'IEPS
        Printer.CurrentX = 2
        If NVENTA1 > 0 And Not (AdoVentas.Recordset!facrfc = False Or globconFin = True) Then Printer.Print Format(NIEPS1, "########0.00");
        Printer.CurrentX = 4
        If NVENTA2 > 0 And Not (AdoVentas.Recordset!facrfc = False Or globconFin = True) Then Printer.Print Format(NIEPS2, "########0.00");
        Printer.CurrentX = 6
        If NVENTA3 > 0 And Not (AdoVentas.Recordset!facrfc = False Or globconFin = True) Then Printer.Print Format(NIEPS3, "########0.00");
        Printer.CurrentX = 8
        If NVENTA4 > 0 And Not (AdoVentas.Recordset!facrfc = False Or globconFin = True) Then Printer.Print Format(NIEPS4, "########0.00");
        Printer.CurrentX = 16
        Printer.Print "IEPS";
        Printer.CurrentX = 18.5
        Printer.FontSize = 10
        If AdoVentas.Recordset!facrfc = False Or globconFin = True Then
           Printer.Print String(10 - Len("00"), " ") & "0.00"
        Else
           Printer.Print String(10 - Len(Trim(Format(NIEPS3 + NIEPS4, "#########0.00"))), " ") & Format(NIEPS3 + NIEPS4, "#########0.00")
        End If
        Printer.FontSize = 8
        'TOTAL DE LA VENTA
        Printer.CurrentX = 2
        If NVENTA1 > 0 Then Printer.Print Format(NVENTA1, "########0.00");
        Printer.CurrentX = 4
        If NVENTA2 > 0 Then Printer.Print Format(NVENTA2, "########0.00");
        Printer.CurrentX = 6
        If NVENTA3 > 0 Then Printer.Print Format(NVENTA3, "########0.00");
        Printer.CurrentX = 8
        If NVENTA4 > 0 Then Printer.Print Format(NVENTA4, "########0.00");
        Printer.CurrentX = 16
        Printer.Print "TOTAL";
        Printer.CurrentX = 18.5
        Printer.FontSize = 10
        Printer.Print String(10 - Len(Trim(Format(NVENTA1 + NVENTA2 + NVENTA3 + NVENTA4, "#########0.00"))), " ") & Format(NVENTA1 + NVENTA2 + NVENTA3 + NVENTA4, "#########0.00")
        
        Printer.FontSize = 8
        Printer.CurrentX = 2
        TotFac = NVENTA1 + NVENTA2 + NVENTA3 + NVENTA4
        LETRAS = NumLet$(TotFac)
        Printer.Print LETRAS
        
        lNvaFac = True: nProd = 0
        cn.Execute "UPDATE facventa SET total = " & NVENTA1 + NVENTA2 + NVENTA3 + NVENTA4 & ", IVA = " & NIVA1 + NIVA2 + NIVA3 + NIVA4 & ", IEPS = " & NIEPS1 + NIEPS2 + NIEPS3 + NIEPS4 & " WHERE serie = '" & IIf(frmVentas.AdoVentas.Recordset!credito, "B", SERIE) & "' AND numfactura = '" & Trim(numfac) & "'"
        marca = rstDetVta.Bookmark
        cn.CommitTrans
        Printer.EndDoc
        rstDetVta.Requery  'Al cerrar la transaccion quiensabe porque decia error catastrofico
        'For n = 1 To marca - 1
        '   rstDetVta.MoveNext
        'Next
        NVENTA1 = 0: NIVA1 = 0: NIEPS1 = 0
        NVENTA2 = 0: NIVA2 = 0: NIEPS2 = 0
        NVENTA3 = 0: NIVA3 = 0: NIEPS3 = 0
        NVENTA4 = 0: NIVA4 = 0: NIEPS4 = 0
   End If
 If Not lNvaFac Then
   Printer.CurrentX = 0.5
   If rstDetVta!cantidad > 0 And nOp <> 0 And nOp <> 1 Then
       Printer.Print rstDetVta!cantidad & " CAJ " & IIf(rstDetVta!cantidadp > 0, rstDetVta!cantidadp & " PZA ", "");
   ElseIf nOp <> 0 And nOp <> 1 Then
       Printer.CurrentX = 1.5
       Printer.Print rstDetVta!cantidadp & " PZA ";
   End If
   Printer.CurrentX = 3
   
   If nOp = 0 Or nOp = 1 Then
    If Printer.TextWidth(rstDetVta!descripc) > 10 Then
      nPos = Int(Len(rstDetVta!descripc) / 2)
      nEsp = 1
      Do  'Se imprimen palabras completas
        nComp = InStr(nEsp, rstDetVta!descripc, " ")
        nEspacio = nComp
        nEsp = nComp + 1
      Loop Until nComp >= nPos
      nPos = nEspacio - 1
      Printer.Print Mid(rstDetVta!descripc, 1, nPos)
      Printer.CurrentX = 3
      Printer.Print Mid(rstDetVta!descripc, nPos + 1);
     Else
       Printer.Print rstDetVta!descripc;
     End If
   Else
     Printer.Print rstDetVta!descripc;
   End If
   
   
   Printer.CurrentX = 11
   Printer.Print rstDetVta!medida;
   Printer.CurrentX = 15
   Printer.Print IIf(AdoVentas.Recordset!facrfc = False Or globconFin = True, "", Format(rstDetVta!ieps, "00"));
   Printer.CurrentX = 16
   Printer.Print IIf(AdoVentas.Recordset!facrfc = False Or globconFin = True, "", Format(rstDetVta!iva, "00"));
    
   Printer.CurrentX = 17
   Printer.Print String(10 - Len(Trim(Format(rstDetVta!PRECIO, "########0.00"))), " ") & Format(rstDetVta!PRECIO, "########0.00");
   Printer.CurrentX = 18.5
   Printer.Print String(12 - Len(Trim(Format(rstDetVta!importe, "########0.00"))), " ") & Format(rstDetVta!importe, "########0.00")
   ncosto = rstDetVta!importe
   
   If globconFin = True Then
      cn.Execute "UPDATE ventas_det SET Serie = '" & IIf(frmVentas.AdoVentas.Recordset!credito, "B", SERIE) & "', Factura = '" & Trim(numfac) & "' FROM VENTAS WHERE VENTAS.folpreventa = " & AdoVentas.Recordset!FOLPREVENTA & " AND VENTAS_DET.Cl_producto = '" & rstDetVta!CONSEC & "' AND VENTAS_DET.Cancelado = 0 AND VENTAS.FACRFC = 0 AND VENTAS_DET.NOVENTA = VENTAS.NOVENTA AND VENTAS.situacion = '1'"
   Else
      If nOp = 0 Or nOp = 1 Then   'Factura sin antecedentes
         cn.Execute "UPDATE Facsinant SET Serie = '" & IIf(frmVentas.AdoVentas.Recordset!credito, "B", SERIE) & "', Factura = '" & Trim(numfac) & "' WHERE noventa = " & AdoVentas.Recordset!noventa
         cn.Execute "UPDATE ventas_det SET Serie = '" & IIf(frmVentas.AdoVentas.Recordset!credito, "B", SERIE) & "', Factura = '" & Trim(numfac) & "' WHERE noventa = " & AdoVentas.Recordset!noventa
      Else
         cn.Execute "UPDATE ventas_det SET Serie = '" & IIf(frmVentas.AdoVentas.Recordset!credito, "B", SERIE) & "', Factura = '" & Trim(numfac) & "' WHERE noventa = " & AdoVentas.Recordset!noventa & " AND cl_producto = '" & rstDetVta!CONSEC & "'"
      End If
   End If
   
   If rstDetVta!iva = 0 And rstDetVta!ieps = 0 Then        'Depto 1
      NVENTA1 = NVENTA1 + ncosto
      NIVA1 = 0
      NIEPS1 = 0
   ElseIf rstDetVta!iva = 15 And rstDetVta!ieps = 0 Then   'Depto 2
       NVENTA2 = NVENTA2 + ncosto
       NIVA2 = NIVA2 + (ncosto / 1.15 * (15 / 100))
       NIEPS2 = 0
   ElseIf rstDetVta!iva = 15 And rstDetVta!ieps = 25 Then  'Depto 3
       NVENTA3 = NVENTA3 + ncosto
       NIVA3 = NIVA3 + (ncosto / 1.15 * (15 / 100))
       NIEPS3 = NIEPS3 + (((ncosto / 1.15) / 1.25) * 25 / 100)
   ElseIf rstDetVta!iva = 15 And rstDetVta!ieps >= 25 Then 'Depto 4
       NVENTA4 = NVENTA4 + ncosto
       NIVA4 = NIVA4 + (ncosto / 1.15 * (15 / 100))
       NIEPS4 = NIEPS4 + (((ncosto / 1.15) / 1.3) * 30 / 100)
   End If
   nProd = nProd + 1
   rstDetVta.MoveNext
 End If
Wend

'Imprimo el pie de pagina de la factura
Printer.CurrentY = 11.5 + nAlto
Printer.CurrentX = 2
Printer.Print "DEP1";
Printer.CurrentX = 4
Printer.Print "DEP2";
Printer.CurrentX = 6
Printer.Print "DEP3";
Printer.CurrentX = 8
Printer.Print "DEP4"

        'SUBTOTALES
        'Cuando es consumidor final no se desglosa la factura ni se imprime ieps e iva
        Printer.CurrentX = 2
        If NVENTA1 > 0 Then Printer.Print Format(NVENTA1, "########0.00");
        Printer.CurrentX = 4
        If NVENTA2 > 0 Then
           Printer.Print IIf(AdoVentas.Recordset!facrfc = False Or globconFin = True, Format(NVENTA2, "########0.00"), Format(NVENTA2 - NIVA2, "########0.00"));
        End If
        Printer.CurrentX = 6
        If NVENTA3 > 0 Then
           Printer.Print IIf(AdoVentas.Recordset!facrfc = False Or globconFin = True, Format(NVENTA3, "########0.00"), Format(NVENTA3 - NIVA3 - NIEPS3, "########0.00"));
        End If
        Printer.CurrentX = 8
        If NVENTA4 > 0 Then
           Printer.Print IIf(AdoVentas.Recordset!facrfc = False Or globconFin = True, Format(NVENTA4, "########0.00"), Format(NVENTA4 - NIVA4 - NIEPS4, "########0.00"));
        End If
        Printer.CurrentX = 16
        Printer.Print "SUBTOTAL";
        Printer.CurrentX = 18.5
        Printer.FontSize = 10
        
        If AdoVentas.Recordset!facrfc = False Or globconFin = True Then
           Printer.Print String(10 - Len(Trim(Format(NVENTA1 + NVENTA2 + (NVENTA3 + NVENTA4), "#########0.00"))), " ") & Format(NVENTA1 + NVENTA2 + NVENTA3 + NVENTA4, "#########0.00")
        Else
           Printer.Print String(10 - Len(Trim(Format(NVENTA1 + (NVENTA2 - NIVA2) + (NVENTA3 - NIVA3 - NIEPS3) + (NVENTA4 - NIVA4 - NIEPS4), "#########0.00"))), " ") & Format(NVENTA1 + (NVENTA2 - NIVA2) + (NVENTA3 - NIVA3 - NIEPS3) + (NVENTA4 - NIVA4 - NIEPS4), "#########0.00")
        End If
        Printer.FontSize = 8
        
        'IVA
        Printer.CurrentX = 2
        If NVENTA1 > 0 And Not (AdoVentas.Recordset!facrfc = False Or globconFin = True) Then Printer.Print Format(NIVA1, "########0.00");
        Printer.CurrentX = 4
        If NVENTA2 > 0 And Not (AdoVentas.Recordset!facrfc = False Or globconFin = True) Then Printer.Print Format(NIVA2, "########0.00");
        Printer.CurrentX = 6
        If NVENTA3 > 0 And Not (AdoVentas.Recordset!facrfc = False Or globconFin = True) Then Printer.Print Format(NIVA3, "########0.00");
        Printer.CurrentX = 8
        If NVENTA4 > 0 And Not (AdoVentas.Recordset!facrfc = False Or globconFin = True) Then Printer.Print Format(NIVA4, "########0.00");
        Printer.CurrentX = 16
        Printer.Print "IVA";
        Printer.CurrentX = 18.5
        Printer.FontSize = 10
        If AdoVentas.Recordset!facrfc = False Or globconFin = True Then
            Printer.Print String(10 - Len("00"), " ") & "0.00"
        Else
            Printer.Print String(10 - Len(Trim(Format(NIVA2 + NIVA3 + NIVA4, "#########0.00"))), " ") & Format(NIVA2 + NIVA3 + NIVA4, "#########0.00")
        End If
        Printer.FontSize = 8
        'IEPS
        Printer.CurrentX = 2
        If NVENTA1 > 0 And Not (AdoVentas.Recordset!facrfc = False Or globconFin = True) Then Printer.Print Format(NIEPS1, "########0.00");
        Printer.CurrentX = 4
        If NVENTA2 > 0 And Not (AdoVentas.Recordset!facrfc = False Or globconFin = True) Then Printer.Print Format(NIEPS2, "########0.00");
        Printer.CurrentX = 6
        If NVENTA3 > 0 And Not (AdoVentas.Recordset!facrfc = False Or globconFin = True) Then Printer.Print Format(NIEPS3, "########0.00");
        Printer.CurrentX = 8
        If NVENTA4 > 0 And Not (AdoVentas.Recordset!facrfc = False Or globconFin = True) Then Printer.Print Format(NIEPS4, "########0.00");
        Printer.CurrentX = 16
        Printer.Print "IEPS";
        Printer.CurrentX = 18.5
        Printer.FontSize = 10
        If AdoVentas.Recordset!facrfc = False Or globconFin = True Then
           Printer.Print String(10 - Len("00"), " ") & "0.00"
        Else
           Printer.Print String(10 - Len(Trim(Format(NIEPS3 + NIEPS4, "#########0.00"))), " ") & Format(NIEPS3 + NIEPS4, "#########0.00")
        End If
        Printer.FontSize = 8
        'TOTAL DE LA VENTA
        Printer.CurrentX = 2
        If NVENTA1 > 0 Then Printer.Print Format(NVENTA1, "########0.00");
        Printer.CurrentX = 4
        If NVENTA2 > 0 Then Printer.Print Format(NVENTA2, "########0.00");
        Printer.CurrentX = 6
        If NVENTA3 > 0 Then Printer.Print Format(NVENTA3, "########0.00");
        Printer.CurrentX = 8
        If NVENTA4 > 0 Then Printer.Print Format(NVENTA4, "########0.00");
        Printer.CurrentX = 16
        Printer.Print "TOTAL";
        Printer.CurrentX = 18.5
        Printer.FontSize = 10
        Printer.Print String(10 - Len(Trim(Format(NVENTA1 + NVENTA2 + NVENTA3 + NVENTA4, "#########0.00"))), " ") & Format(NVENTA1 + NVENTA2 + NVENTA3 + NVENTA4, "#########0.00")

Printer.FontSize = 8
Printer.CurrentX = 2
TotFac = NVENTA1 + NVENTA2 + NVENTA3 + NVENTA4
LETRAS = NumLet$(TotFac)
Printer.Print LETRAS

'Printer.Print Chr(27) + Chr(64)
'Printer.Print Chr(27) + Chr(105)
cn.Execute "UPDATE facventa SET total = " & NVENTA1 + NVENTA2 + NVENTA3 + NVENTA4 + NVENTA5 + NVENTA6 + NVENTA7 & ", IVA = " & NIVA1 + NIVA2 + NIVA3 + NIVA4 + NIVA5 + NIVA6 + NIVA7 & ", IEPS = " & NIEPS1 + NIEPS2 + NIEPS3 + NIEPS4 + NIEPS5 + NIEPS6 + NIEPS7 & " WHERE serie = '" & IIf(frmVentas.AdoVentas.Recordset!credito, "B", SERIE) & "' AND numfactura = '" & Trim(numfac) & "'"
Printer.EndDoc
Factura = True
'cn.CommitTrans
Exit Function
Error:
   MsgBox "OCURRIO EL SIGUIENTE ERROR: " & Err.Description & Chr(13) & "APAGUE LA IMPRESORA Y CANCELE LA IMPRESION DESDE EL PANEL DE CONTROL PARA NO DESPERDICIAR FACTURAS", vbExclamation
   If lTrans Then cn.RollbackTrans
End Function

Private Sub SUMAPRODUCTO(p As Integer)
Dim iva  As Currency
iva = IIf(ZONA = "OAX", 15, 10)
iva = iva / 100
If tasas(p) = 1 Then        'Depto 1
    NVENTA1 = NVENTA1 + ncosto
    NIVA1 = 0
    NIEPS1 = 0
ElseIf tasas(p) = 2 Then   'Depto 2
    NVENTA2 = NVENTA2 + ncosto
    NIVA2 = NIVA2 + (ncosto / (1 + iva) * iva)
    NIEPS2 = 0
ElseIf tasas(p) = 3 Then  'Depto 3
    NVENTA3 = NVENTA3 + ncosto
    NIVA3 = NIVA3 + (ncosto / (1 + iva) * iva)
    NIEPS3 = NIEPS3 + (((ncosto / (1 + iva)) / 1.25) * 25 / 100)
ElseIf tasas(p) = 4 Then  'Depto 4
    NVENTA4 = NVENTA4 + ncosto
    NIVA4 = NIVA4 + (ncosto / (1 + iva) * iva)
    NIEPS4 = NIEPS4 + (((ncosto / (1 + iva)) / 1.3) * 30 / 100)
ElseIf tasas(p) = 5 Then  'Depto 5
    NVENTA5 = NVENTA5 + ncosto
    'NIVA5 = NIVA5 + (((ncosto / (1 + iva)) / 1.6) * 60 / 100)
    NIVA5 = NIVA5 + (ncosto / (1 + iva) * iva)
    NIEPS5 = 0
    'NIEPS5 = NIEPS5 + (((ncosto / 1.15) / 1.5) * 50 / 100)
ElseIf tasas(p) = 6 Then   'Depto 6
    NVENTA6 = NVENTA6 + ncosto
    NIVA6 = NIVA6 + (ncosto / (1 + iva) * iva)
    NIEPS6 = NIEPS6 + (((ncosto / (1 + iva)) / 1.5) * 50 / 100)
ElseIf tasas(p) = 7 Then  'Depto 7
    NVENTA7 = NVENTA7 + ncosto
    NIVA7 = NIVA7 + (ncosto / (1 + iva) * iva)
    NIEPS7 = NIEPS7 + (((ncosto / (1 + iva)) / 1.05) * 5 / 100)
ElseIf tasas(p) = 8 Then  'Depto 8
    NVENTA8 = NVENTA8 + ncosto
    NIVA8 = NIVA8 + (ncosto / (1 + iva) * iva)
    NIEPS8 = NIEPS8 + (((ncosto / (1 + iva)) / 1.2) * 20 / 100)
    'NVENTA8 = NVENTA8 + ncosto
    'NIVA8 = NIVA8 + (ncosto / 1.15 * (15 / 100))
    'NIEPS8 = NIEPS8 + (((ncosto / 1.15) / 1.2) * 20 / 100)
End If
End Sub

Private Function GRABASERIEYFAC(p As Integer, Optional globconFin As Boolean) As Boolean
On Error GoTo Error:
'AQUI SI SE DEBE IMPLEMENTAR UN RECORDSET PARA OBTENER EL FOLIO DE VENTA
If globconFin = False Then
   CADENA = "UPDATE ventas_det SET FACTURADO = 1, Serie = '" & sert & "', Factura = '" & Trim(numfac) & "' WHERE noventa = " & NOVENTAT(p) & " AND cl_producto = '" & Trim(consecs(p)) & "'"
   cn.Execute CADENA
   GRABASERIEYFAC = True
Else
   CADENA = "UPDATE ventas_det SET facturado = 1, Serie = '" & sert & "', Factura = '" & Trim(numfac) & "' FROM VENTAS WHERE VENTAS.folpreventa = " & adopreventa.Recordset!FOLPREVENTA & " AND VENTAS_DET.Cl_producto = '" & consecs(p) & "' AND VENTAS_DET.Cancelado = 0 AND VENTAS.FACRFC = 0 AND VENTAS_DET.NOVENTA = VENTAS.NOVENTA AND FACTURADO = 0"
   cn.Execute CADENA
   GRABASERIEYFAC = True
End If
Exit Function
Error:
   GRABASERIEYFAC = False
End Function

Public Sub IMPRIMEDETALLEOLD(Prodreal As Integer)
On Error GoTo Error:
p = Prodreal
Printer.ScaleMode = vbCentimeters
Printer.CurrentX = 0.3
Printer.FontName = "ARIAL"
Printer.FontSize = 7
If cantidads(p) > 0 Then
   Printer.Print cantidads(p) & "CJ" & IIf(cantidadps(p) > 0, "-" & cantidadps(p) & "PZ", "");
Else
   Printer.CurrentX = 0.3
   Printer.Print cantidadps(p) & "PZ";
End If
Printer.CurrentX = 1.5
If Printer.TextWidth(descripcs(p)) > 5.5 Then
   For N = 1 To Len(descripcs(p))
     If Printer.TextWidth(Mid(descripcs(p), 1, N)) > 5 Then Exit For
   Next
   Printer.Print Mid(descripcs(p), 1, N);
Else
   Printer.Print descripcs(p);
End If
Printer.CurrentX = 6.5
Printer.Print medidas(p);
Printer.CurrentX = 9.5
'En la factura si es consumidor final el iva y ieps van en ceros
Printer.Print IIf(rfct = RFCFINAL Or globconFin = True, "", Format(iepss(p), "00"));
Printer.CurrentX = 10
Printer.Print IIf(rfct = RFCFINAL Or globconFin = True, "", Format(ivas(p), "00"));
Printer.CurrentX = 10.3
Printer.Print String(10 - Len(Trim(Format(precios(p), "########0.00"))), " ") & Format(precios(p), "########0.00");
Printer.CurrentX = 11.8
Printer.Print String(12 - Len(Trim(Format(importes(p), "########0.00"))), " ") & Format(importes(p), "########0.00")
ncosto = importes(p)
Call grabadetallefactura(consecs(p), cantidads(p), cantidadps(p), precios(p), preciosp(p), costoss(p), costosp(p), importes(p), ivas(p), iepss(p), numfac, sert, NOVENTAT(p), tasas(p))
Exit Sub
Error:
  cn.Execute "UPADTE ventas_det SET facturado = 0, factura = null, serie = null WHERE factura = '" & Factura & "' and serie = '" & Trim(SERIE) & "'"
  cn.Execute "DELETE FROM facventa WHERE numfactura = '" & Factura & "' and serie = '" & SERIE & "'"
  cn.Execute "DELETE FROM facventa_det WHERE factura = '" & Factura & "' and serie = '" & SERIE & "'"
  MsgBox "EXISTE UN ERROR EN EL MOMENTO DE IMPRIMIR EL DETALLE DE LA FACTURA (Error: 5)", vbCritical, "FACTURAS"
End Sub

Public Sub IMPRIMEDETALLEOLDchica(Prodreal As Integer)
On Error GoTo Error:
Printer.ScaleMode = vbCentimeters
Printer.CurrentX = 0.3
Printer.FontName = "ARIAL"
Printer.FontSize = 7

p = Prodreal

If cantidads(p) > 0 Then
   Printer.Print cantidads(p) & "CJ" & IIf(cantidadps(p) > 0, "-" & cantidadps(p) & "PZ", "");
Else
   Printer.CurrentX = 0.3
   Printer.Print cantidadps(p) & "PZ";
End If
Printer.CurrentX = 3
If Printer.TextWidth(descripcs(p)) > 5.5 Then
   For N = 1 To Len(descripcs(p))
     If Printer.TextWidth(Mid(descripcs(p), 1, N)) > 5 Then Exit For
   Next
   Printer.Print Mid(descripcs(p), 1, N);
Else
   Printer.Print descripcs(p);
End If
Printer.CurrentX = 11
Printer.Print medidas(p);
Printer.CurrentX = 15
'en la factura si es consumidor final el iva y ieps van en ceros
Printer.Print IIf(rfct = RFCFINAL Or globconFin = True, "", Format(iepss(p), "00"));
Printer.CurrentX = 16
Printer.Print IIf(rfct = RFCFINAL Or globconFin = True, "", Format(ivas(p), "00"));
Printer.CurrentX = 17
Printer.Print String(10 - Len(Trim(Format(precios(p), "########0.00"))), " ") & Format(precios(p), "########0.00");
Printer.CurrentX = 18.5
Printer.Print String(12 - Len(Trim(Format(importes(p), "########0.00"))), " ") & Format(importes(p), "########0.00")
ncosto = importes(p)
Call grabadetallefactura(consecs(p), cantidads(p), cantidadps(p), precios(p), preciosp(p), costoss(p), costosp(p), importes(p), ivas(p), iepss(p), numfac, sert, NOVENTAT(p), tasas(p))
Exit Sub
Error:
  cn.Execute "UPADTE ventas_det SET facturado = 0, factura = null, serie = null WHERE factura = '" & Factura & "' and serie = '" & Trim(SERIE) & "'"
  cn.Execute "DELETE FROM facventa WHERE numfactura = '" & Factura & "' and serie = '" & SERIE & "'"
  cn.Execute "DELETE FROM facventa_det WHERE factura = '" & Factura & "' and serie = '" & SERIE & "'"
  MsgBox "EXISTE UN ERROR EN EL MOMENTO DE IMPRIMIR EL DETALLE", vbCritical, "FACTURAS" & Chr(13) & "Reintente cobrar la venta en la misma factura"
End Sub


Private Sub grabadetallefactura(producto, cantidad, cantidadp, PRECIO, preciop, costo, costop, importe, iva, ieps, Factura, SERIE, venta, tasaieps)
'para generar el detalle de la factura
CAD = "INSERT INTO facventa_det(producto,cantidad,cantidadp,precio,preciop,costo,costop,importe,iva,ieps,factura,serie,venta,tasaieps,fecha_det,rfc_det) " & _
      " values (" & Trim(producto) & "," & cantidad & "," & cantidadp & "," & PRECIO & "," & preciop & "," & costo & "," & costop & "," & importe & "," & iva & "," & ieps & "," & Factura & ",'" & Trim(SERIE) & "'," & venta & "," & tasaieps & ",'" & Format(date, "dd/mm/yyyy") & "','" & rfct & "')"
cn.Execute CAD
End Sub

Private Sub IMPRIMESUBOLD()
Printer.FontName = "ARIAL"
Printer.FontSize = 7

Printer.CurrentY = 15.7 + nAlto
Printer.CurrentX = 0.3
Printer.Print "DEP1";
Printer.CurrentX = 1.6
Printer.Print "DEP2";
Printer.CurrentX = 3
Printer.Print "DEP3";
Printer.CurrentX = 4.3
Printer.Print "DEP4";
Printer.CurrentX = 5.6
Printer.Print "DEP5";
Printer.CurrentX = 6.8
Printer.Print "DEP6";
Printer.CurrentX = 8.1
Printer.Print "DEP7";
Printer.CurrentX = 9.4
Printer.Print "DEP8"
'SUBTOTALES
'Cuando es consumidor final no se desglosa la factura ni se imprime ieps e iva
Printer.CurrentX = 0.3
If NVENTA1 > 0 Then Printer.Print Format(NVENTA1, "########0.00");
Printer.CurrentX = 1.6
If NVENTA2 > 0 Then
   Printer.Print IIf(rfct = RFCFINAL Or globconFin = True, Format(NVENTA2, "########0.00"), Format(NVENTA2 - NIVA2, "########0.00"));
End If
Printer.CurrentX = 3
If NVENTA3 > 0 Then
   Printer.Print IIf(rfct = RFCFINAL Or globconFin = True, Format(NVENTA3, "########0.00"), Format(NVENTA3 - NIVA3 - NIEPS3, "########0.00"));
End If
Printer.CurrentX = 4.3
If NVENTA4 > 0 Then
   Printer.Print IIf(rfct = RFCFINAL Or globconFin = True, Format(NVENTA4, "########0.00"), Format(NVENTA4 - NIVA4 - NIEPS4, "########0.00"));
End If
Printer.CurrentX = 5.6
If NVENTA5 > 0 Then
   Printer.Print IIf(rfct = RFCFINAL Or globconFin = True, Format(NVENTA5, "########0.00"), Format(NVENTA5 - NIVA5 - NIEPS5, "########0.00"));
End If
Printer.CurrentX = 6.8
If NVENTA6 > 0 Then
   Printer.Print IIf(rfct = RFCFINAL Or globconFin = True, Format(NVENTA6, "########0.00"), Format(NVENTA6 - NIVA6 - NIEPS6, "########0.00"));
End If
Printer.CurrentX = 8.1
If NVENTA7 > 0 Then
   Printer.Print IIf(rfct = RFCFINAL Or globconFin = True, Format(NVENTA7, "########0.00"), Format(NVENTA7 - NIVA7 - NIEPS7, "########0.00"));
End If
Printer.CurrentX = 9.4
If NVENTA8 > 0 Then
   Printer.Print IIf(rfct = RFCFINAL Or globconFin = True, Format(NVENTA8, "########0.00"), Format(NVENTA8 - NIVA8 - NIEPS8, "########0.00"));
End If


Printer.CurrentX = 10.5
Printer.Print "SUBTOTAL";
Printer.CurrentX = 12
Printer.FontSize = 9
If rfct = RFCFINAL Or globconFin = True Then
   Printer.Print String(11 - Len(Trim(Format(NVENTA1 + NVENTA2 + NVENTA3 + NVENTA4 + NVENTA5 + NVENTA6 + NVENTA7 + NVENTA8, "#########0.00"))), " ") & Format(NVENTA1 + NVENTA2 + NVENTA3 + NVENTA4 + NVENTA5 + NVENTA6 + NVENTA7 + NVENTA8, "#########0.00")
Else
   Printer.Print String(11 - Len(Trim(Format(NVENTA1 + (NVENTA2 - NIVA2) + (NVENTA3 - NIVA3 - NIEPS3) + (NVENTA4 - NIVA4 - NIEPS4) + (NVENTA5 - NIVA5 - NIEPS5) + (NVENTA6 - NIVA6 - NIEPS6) + (NVENTA7 - NIVA7 - NIEPS7) + (NVENTA8 - NIVA8 - NIEPS8), "#########0.00"))), " ") & Format(NVENTA1 + (NVENTA2 - NIVA2) + (NVENTA3 - NIVA3 - NIEPS3) + (NVENTA4 - NIVA4 - NIEPS4) + (NVENTA5 - NIVA5 - NIEPS5) + (NVENTA6 - NIVA6 - NIEPS6) + (NVENTA7 - NIVA7 - NIEPS7) + (NVENTA8 - NIVA8 - NIEPS8), "#########0.00")
End If
Printer.FontSize = 7
        
'IVA
Printer.CurrentX = 0.3
If NVENTA1 > 0 And Not (rfct = RFCFINAL Or globconFin = True) Then Printer.Print Format(NIVA1, "########0.00");
Printer.CurrentX = 1.6
If NVENTA2 > 0 And Not (rfct = RFCFINAL Or globconFin = True) Then Printer.Print Format(NIVA2, "########0.00");
Printer.CurrentX = 3
If NVENTA3 > 0 And Not (rfct = RFCFINAL Or globconFin = True) Then Printer.Print Format(NIVA3, "########0.00");
Printer.CurrentX = 4.3
If NVENTA4 > 0 And Not (rfct = RFCFINAL Or globconFin = True) Then Printer.Print Format(NIVA4, "########0.00");
Printer.CurrentX = 5.6
If NVENTA5 > 0 And Not (rfct = RFCFINAL Or globconFin = True) Then Printer.Print Format(NIVA5, "########0.00");
Printer.CurrentX = 6.8
If NVENTA6 > 0 And Not (rfct = RFCFINAL Or globconFin = True) Then Printer.Print Format(NIVA6, "########0.00");
Printer.CurrentX = 8.1
If NVENTA7 > 0 And Not (rfct = RFCFINAL Or globconFin = True) Then Printer.Print Format(NIVA7, "########0.00");
Printer.CurrentX = 9.4
If NVENTA8 > 0 And Not (rfct = RFCFINAL Or globconFin = True) Then Printer.Print Format(NIVA8, "########0.00");
Printer.CurrentX = 10.5
Printer.Print "IVA";
Printer.CurrentX = 12
Printer.FontSize = 9
If rfct = RFCFINAL Or globconFin = True Then
   Printer.Print String(11 - Len("00"), " ") & "0.00"
Else
   Printer.Print String(11 - Len(Trim(Format(NIVA2 + NIVA3 + NIVA4 + NIVA5 + NIVA6 + NIVA7 + NIVA8, "#########0.00"))), " ") & Format(NIVA2 + NIVA3 + NIVA4 + NIVA5 + NIVA6 + NIVA7 + NIVA8, "#########0.00")
End If
Printer.FontSize = 7

'IEPS
Printer.CurrentX = 0.3
If NVENTA1 > 0 And Not (rfct = RFCFINAL Or globconFin = True) Then Printer.Print Format(NIEPS1, "########0.00");
Printer.CurrentX = 1.6
If NVENTA2 > 0 And Not (rfct = RFCFINAL Or globconFin = True) Then Printer.Print Format(NIEPS2, "########0.00");
Printer.CurrentX = 3
If NVENTA3 > 0 And Not (rfct = RFCFINAL Or globconFin = True) Then Printer.Print Format(NIEPS3, "########0.00");
Printer.CurrentX = 4.3
If NVENTA4 > 0 And Not (rfct = RFCFINAL Or globconFin = True) Then Printer.Print Format(NIEPS4, "########0.00");
Printer.CurrentX = 5.6
If NVENTA5 > 0 And Not (rfct = RFCFINAL Or globconFin = True) Then Printer.Print Format(NIEPS5, "########0.00");
Printer.CurrentX = 6.8
If NVENTA6 > 0 And Not (rfct = RFCFINAL Or globconFin = True) Then Printer.Print Format(NIEPS6, "########0.00");
Printer.CurrentX = 8.1
If NVENTA7 > 0 And Not (rfct = RFCFINAL Or globconFin = True) Then Printer.Print Format(NIEPS7, "########0.00");
Printer.CurrentX = 9.4
If NVENTA8 > 0 And Not (rfct = RFCFINAL Or globconFin = True) Then Printer.Print Format(NIEPS8, "########0.00");

Printer.CurrentX = 10.5
Printer.Print "IEPS";
Printer.CurrentX = 12
Printer.FontSize = 9
If rfct = RFCFINAL Or globconFin = True Then
   Printer.Print String(11 - Len("00"), " ") & "0.00"
Else
   Printer.Print String(11 - Len(Trim(Format(NIEPS3 + NIEPS4 + NIEPS5 + NIEPS6 + NIEPS7 + NIEPS8, "#########0.00"))), " ") & Format(NIEPS3 + NIEPS4 + NIEPS5 + NIEPS6 + NIEPS7 + NIEPS8, "#########0.00")
End If
Printer.FontSize = 7
'TOTAL DE LA VENTA
Printer.CurrentX = 0.3
If NVENTA1 > 0 Then Printer.Print Format(NVENTA1, "########0.00");
Printer.CurrentX = 1.6
If NVENTA2 > 0 Then Printer.Print Format(NVENTA2, "########0.00");
Printer.CurrentX = 3
If NVENTA3 > 0 Then Printer.Print Format(NVENTA3, "########0.00");
Printer.CurrentX = 4.3
If NVENTA4 > 0 Then Printer.Print Format(NVENTA4, "########0.00");
Printer.CurrentX = 5.6
If NVENTA5 > 0 Then Printer.Print Format(NVENTA5, "########0.00");
Printer.CurrentX = 6.8
If NVENTA6 > 0 Then Printer.Print Format(NVENTA6, "########0.00");
Printer.CurrentX = 8.1
If NVENTA7 > 0 Then Printer.Print Format(NVENTA7, "########0.00");
Printer.CurrentX = 9.4
If NVENTA8 > 0 Then Printer.Print Format(NVENTA8, "########0.00");


Printer.CurrentX = 10.5
Printer.Print "TOTAL";
Printer.CurrentX = 12
Printer.FontSize = 9
Printer.Print String(11 - Len(Trim(Format(NVENTA1 + NVENTA2 + NVENTA3 + NVENTA4 + NVENTA5 + NVENTA6 + NVENTA7 + NVENTA8, "#########0.00"))), " ") & Format(NVENTA1 + NVENTA2 + NVENTA3 + NVENTA4 + NVENTA5 + NVENTA6 + NVENTA7 + NVENTA8, "#########0.00")

Printer.FontSize = 7
Printer.CurrentX = 1
Dim TotFac As Double
TotFac = NVENTA1 + NVENTA2 + NVENTA3 + NVENTA4 + NVENTA5 + NVENTA6 + NVENTA7 + NVENTA8
nIva = NIVA1 + NIVA2 + NIVA3 + NIVA4 + NIVA5 + NIVA6 + NIVA7 + NIVA8
nIeps = NIEPS1 + NIEPS2 + NIEPS3 + NIEPS4 + NIEPS5 + NIEPS6 + NIEPS7 + NIEPS8
'EN ESTE MOMENTO SE ACTUALIZA EL MONTO DE LA FACTURA
If chkCredito.Value = 1 Or chkPrev.Value = 1 Then
   cn.Execute "UPDATE facventa SET totfac = " & TotFac & ", TOTAL = " & TotFac & ", PorPagar = " & TotFac & ", IVA = " & nIva & ", IEPS = " & nIeps & ", depto1 = " & NVENTA1 & ", depto2 = " & NVENTA2 & ", depto3 = " & NVENTA3 & ", depto4 = " & NVENTA4 & ", depto5 = " & NVENTA5 & ", depto6 = " & NVENTA6 & ", depto7 = " & NVENTA7 & ", depto8 = " & NVENTA8 & " WHERE numfactura = '" & Trim(numfac) & "' and serie = '" & Trim(SERIE) & "'"
Else
   cn.Execute "UPDATE facventa SET totfac = " & TotFac & ", TOTAL = " & TotFac & ", IVA = " & nIva & ", IEPS = " & nIeps & ", depto1 = " & NVENTA1 & ", depto2 = " & NVENTA2 & ", depto3 = " & NVENTA3 & ", depto4 = " & NVENTA4 & ", depto5 = " & NVENTA5 & ", depto6 = " & NVENTA6 & ", depto7 = " & NVENTA7 & ", depto8 = " & NVENTA8 & " WHERE numfactura = '" & Trim(numfac) & "' and serie = '" & Trim(SERIE) & "'"
End If

'Printer.CurrentY = 12.5

'RESP = InputBox("Leyenda 1", "Programa")
'Printer.Print RESP
'RESP = InputBox("Leyenda 2", "Programa")
'Printer.Print RESP
'Printer.CurrentY = 14

LETRAS = NumLet$(Format(TotFac, "########0.00"))
Printer.Print LETRAS
Printer.CurrentY = 19.5
Printer.CurrentX = 7
Printer.Print "UNA SOLA EXHIBICION " & IIf(chkCredito.Value = 0 Or chkPrev.Value = 0, "X", "____") & Space(10) & "PARCIALIDADES ____"
Printer.CurrentX = 7
If chkCredito.Value = 1 Or chkPrev.Value = 1 Then Printer.Print "EFECTOS FISCALES AL PAGO"
Printer.Print " "
Printer.EndDoc
Call INICIALIZAVALS
End Sub

Private Sub IMPRIMESUBOLDchica()
Printer.FontName = "ARIAL"
Printer.FontSize = 7

Printer.CurrentY = 11.5 + nAlto
Printer.CurrentX = 1
Printer.Print "DEP1";
Printer.CurrentX = 3
Printer.Print "DEP2";
Printer.CurrentX = 5
Printer.Print "DEP3";
Printer.CurrentX = 7
Printer.Print "DEP4";
Printer.CurrentX = 9
Printer.Print "DEP5";
Printer.CurrentX = 11
Printer.Print "DEP6";
Printer.CurrentX = 13
Printer.Print "DEP7";
Printer.CurrentX = 15
Printer.Print "DEP8"
'SUBTOTALES
'Cuando es consumidor final no se desglosa la factura ni se imprime ieps e iva
Printer.CurrentX = 1
If NVENTA1 > 0 Then Printer.Print Format(NVENTA1, "########0.00");
Printer.CurrentX = 3
If NVENTA2 > 0 Then Printer.Print IIf(rfct = RFCFINAL Or globconFin = True, Format(NVENTA2, "########0.00"), Format(NVENTA2 - NIVA2, "########0.00"));
Printer.CurrentX = 5
If NVENTA3 > 0 Then Printer.Print IIf(rfct = RFCFINAL Or globconFin = True, Format(NVENTA3, "########0.00"), Format(NVENTA3 - NIVA3 - NIEPS3, "########0.00"));
Printer.CurrentX = 7
If NVENTA4 > 0 Then Printer.Print IIf(rfct = RFCFINAL Or globconFin = True, Format(NVENTA4, "########0.00"), Format(NVENTA4 - NIVA4 - NIEPS4, "########0.00"));
Printer.CurrentX = 9
If NVENTA5 > 0 Then Printer.Print IIf(rfct = RFCFINAL Or globconFin = True, Format(NVENTA5, "########0.00"), Format(NVENTA5 - NIVA5 - NIEPS5, "########0.00"));
Printer.CurrentX = 11
If NVENTA6 > 0 Then Printer.Print IIf(rfct = RFCFINAL Or globconFin = True, Format(NVENTA6, "########0.00"), Format(NVENTA6 - NIVA6 - NIEPS6, "########0.00"));
Printer.CurrentX = 13
If NVENTA7 > 0 Then Printer.Print IIf(rfct = RFCFINAL Or globconFin = True, Format(NVENTA7, "########0.00"), Format(NVENTA7 - NIVA7 - NIEPS7, "########0.00"));
Printer.CurrentX = 8
If NVENTA8 > 0 Then Printer.Print IIf(rfct = RFCFINAL Or globconFin = True, Format(NVENTA8, "########0.00"), Format(NVENTA8 - NIVA8 - NIEPS8, "########0.00"));

Printer.CurrentX = 16
Printer.Print "SUBTOTAL";
Printer.CurrentX = 18.5
Printer.FontSize = 10
        
If rfct = RFCFINAL Or globconFin = True Then
   Printer.Print String(10 - Len(Trim(Format(NVENTA1 + NVENTA2 + (NVENTA3 + NVENTA4 + NVENTA5 + NVENTA6 + NVENTA7 + NVENTA8), "#########0.00"))), " ") & Format(NVENTA1 + NVENTA2 + NVENTA3 + NVENTA4 + NVENTA5 + NVENTA6 + NVENTA7 + NVENTA8, "#########0.00")
Else
   Printer.Print String(10 - Len(Trim(Format(NVENTA1 + (NVENTA2 - NIVA2) + (NVENTA3 - NIVA3 - NIEPS3) + (NVENTA4 - NIVA4 - NIEPS4) + (NVENTA5 - NIVA5 - NIEPS5) + (NVENTA6 - NIVA6 - NIEPS6) + (NVENTA7 - NIVA7 - NIEPS7) + (NVENTA8 - NIVA8 - NIEPS8), "#########0.00"))), " ") & Format(NVENTA1 + (NVENTA2 - NIVA2) + (NVENTA3 - NIVA3 - NIEPS3) + (NVENTA4 - NIVA4 - NIEPS4) + (NVENTA5 - NIVA5 - NIEPS5) + (NVENTA6 - NIVA6 - NIEPS6) + (NVENTA7 - NIVA7 - NIEPS7) + (NVENTA8 - NIVA8 - NIEPS8), "#########0.00")
End If
Printer.FontSize = 8
        
'IVA
Printer.CurrentX = 1
If NVENTA1 > 0 And Not (rfct = RFCFINAL Or globconFin = True) Then Printer.Print Format(NIVA1, "########0.00");
Printer.CurrentX = 3
If NVENTA2 > 0 And Not (rfct = RFCFINAL Or globconFin = True) Then Printer.Print Format(NIVA2, "########0.00");
Printer.CurrentX = 5
If NVENTA3 > 0 And Not (rfct = RFCFINAL Or globconFin = True) Then Printer.Print Format(NIVA3, "########0.00");
Printer.CurrentX = 7
If NVENTA4 > 0 And Not (rfct = RFCFINAL Or globconFin = True) Then Printer.Print Format(NIVA4, "########0.00");
Printer.CurrentX = 9
If NVENTA5 > 0 And Not (rfct = RFCFINAL Or globconFin = True) Then Printer.Print Format(NIVA5, "########0.00");
Printer.CurrentX = 11
If NVENTA6 > 0 And Not (rfct = RFCFINAL Or globconFin = True) Then Printer.Print Format(NIVA6, "########0.00");
Printer.CurrentX = 13
If NVENTA7 > 0 And Not (rfct = RFCFINAL Or globconFin = True) Then Printer.Print Format(NIVA7, "########0.00");
Printer.CurrentX = 15
If NVENTA8 > 0 And Not (rfct = RFCFINAL Or globconFin = True) Then Printer.Print Format(NIVA8, "########0.00");


Printer.CurrentX = 16
Printer.Print "IVA";
Printer.CurrentX = 18.5
Printer.FontSize = 10
If rfct = RFCFINAL Or globconFin = True Then
   Printer.Print String(10 - Len("00"), " ") & "0.00"
Else
   Printer.Print String(10 - Len(Trim(Format(NIVA2 + NIVA3 + NIVA4 + NIVA5 + NIVA6 + NIVA7 + NIVA8, "#########0.00"))), " ") & Format(NIVA2 + NIVA3 + NIVA4 + NIVA5 + NIVA6 + NIVA7 + NIVA8, "#########0.00")
End If
Printer.FontSize = 7
'IEPS
Printer.CurrentX = 1
If NVENTA1 > 0 And Not (rfct = RFCFINAL Or globconFin = True) Then Printer.Print Format(NIEPS1, "########0.00");
Printer.CurrentX = 3
If NVENTA2 > 0 And Not (rfct = RFCFINAL Or globconFin = True) Then Printer.Print Format(NIEPS2, "########0.00");
Printer.CurrentX = 5
If NVENTA3 > 0 And Not (rfct = RFCFINAL Or globconFin = True) Then Printer.Print Format(NIEPS3, "########0.00");
Printer.CurrentX = 7
If NVENTA4 > 0 And Not (rfct = RFCFINAL Or globconFin = True) Then Printer.Print Format(NIEPS4, "########0.00");
Printer.CurrentX = 9
If NVENTA5 > 0 And Not (rfct = RFCFINAL Or globconFin = True) Then Printer.Print Format(NIEPS5, "########0.00");
Printer.CurrentX = 11
If NVENTA6 > 0 And Not (rfct = RFCFINAL Or globconFin = True) Then Printer.Print Format(NIEPS6, "########0.00");
Printer.CurrentX = 13
If NVENTA7 > 0 And Not (rfct = RFCFINAL Or globconFin = True) Then Printer.Print Format(NIEPS7, "########0.00");
Printer.CurrentX = 15
If NVENTA8 > 0 And Not (rfct = RFCFINAL Or globconFin = True) Then Printer.Print Format(NIEPS8, "########0.00");

Printer.CurrentX = 16
Printer.Print "IEPS";
Printer.CurrentX = 18.5
Printer.FontSize = 10
If rfct = RFCFINAL Or globconFin = True Then
   Printer.Print String(10 - Len("00"), " ") & "0.00"
Else
   Printer.Print String(10 - Len(Trim(Format(NIEPS3 + NIEPS4 + NIEPS5 + NIEPS6 + NIEPS7 + NIEPS8, "#########0.00"))), " ") & Format(NIEPS3 + NIEPS4 + NIEPS5 + NIEPS6 + NIEPS7 + NIEPS8, "#########0.00")
End If
Printer.FontSize = 8
'TOTAL DE LA VENTA
Printer.CurrentX = 1
If NVENTA1 > 0 Then Printer.Print Format(NVENTA1, "########0.00");
Printer.CurrentX = 3
If NVENTA2 > 0 Then Printer.Print Format(NVENTA2, "########0.00");
Printer.CurrentX = 5
If NVENTA3 > 0 Then Printer.Print Format(NVENTA3, "########0.00");
Printer.CurrentX = 7
If NVENTA4 > 0 Then Printer.Print Format(NVENTA4, "########0.00");
Printer.CurrentX = 9
If NVENTA5 > 0 Then Printer.Print Format(NVENTA5, "########0.00");
Printer.CurrentX = 11
If NVENTA6 > 0 Then Printer.Print Format(NVENTA6, "########0.00");
Printer.CurrentX = 13
If NVENTA7 > 0 Then Printer.Print Format(NVENTA7, "########0.00");
Printer.CurrentX = 15
If NVENTA8 > 0 Then Printer.Print Format(NVENTA8, "########0.00");

Printer.CurrentX = 16
Printer.Print "TOTAL";
Printer.CurrentX = 18.5
Printer.FontSize = 10
Printer.Print String(10 - Len(Trim(Format(NVENTA1 + NVENTA2 + NVENTA3 + NVENTA4 + NVENTA5 + NVENTA6 + NVENTA7 + NVENTA8, "#########0.00"))), " ") & Format(NVENTA1 + NVENTA2 + NVENTA3 + NVENTA4 + NVENTA5 + NVENTA6 + NVENTA7 + NVENTA8, "#########0.00")

Printer.FontSize = 7
Printer.CurrentX = 2
Dim TotFac As Double
TotFac = NVENTA1 + NVENTA2 + NVENTA3 + NVENTA4 + NVENTA5 + NVENTA6 + NVENTA7 + NVENTA8
'EN ESTE MOMENTO SE ACTUALIZA EL MONTO DE LA FACTURA
CADENA = "Update facventa set totfac = " & TotFac & ", depto1 = " & NVENTA1 & ", depto2 = " & NVENTA2 & ", depto3 = " & NVENTA3 & ", depto4 = " & NVENTA4 & " where numfactura = '" & Trim(numfac) & "' and serie = '" & Trim(SERIE) & "'"
cn.Execute CADENA
cn.Execute "UPDATE facventa SET total = " & TotFac & ", IVA = " & NIVA1 + NIVA2 + NIVA3 + NIVA4 + NIVA5 + NIVA6 + NIVA7 + NIVA8 & ", IEPS = " & NIEPS1 + NIEPS2 + NIEPS3 + NIEPS4 + NIEPS5 + NIEPS6 + NIEPS7 + NIEPS8 & " WHERE serie = '" & Trim(SERIE) & "' AND numfactura = '" & Trim(numfac) & "'"
LETRAS = NumLet$(TotFac)
Printer.Print LETRAS
'Printer.Print Chr(27) + Chr(64)
'Printer.Print Chr(27) + Chr(105)
Printer.EndDoc
Call INICIALIZAVALS
End Sub

Private Sub obtengenerales(Optional glob As Boolean)
On Error GoTo Error:
rscli.Requery
If glob Then
   nomt = "C O N S U M I D O R   F I N A L"
   rfct = "COOF970101111"
   dirt = "C O N O C I D O"
   telr = "": colt = "": ciut = ""
Else
   nomt = IIf(AdoVentas.Recordset!facrfc = False, "C O N S U M I D O R   F I N A L", rscli!cnombrefac)
   rfct = IIf(AdoVentas.Recordset!facrfc = False, "COOF970101111", rscli!crfc)
   dirt = IIf(AdoVentas.Recordset!facrfc = False, "C O N O C I D O", rscli!cdireccion)
   telr = IIf(AdoVentas.Recordset!facrfc = False, "", rscli!ctelefono)
   colt = IIf(IsNull(rscli!ccolonia) Or AdoVentas.Recordset!facrfc = False, "", rscli!ccolonia)
   ciut = IIf(IsNull(rscli!cciudad) Or AdoVentas.Recordset!facrfc = False, "", rscli!cciudad)
End If
sert = ObtenSerie(frmVentas.AdoVentas.Recordset!credito, frmVentas.AdoVentas.Recordset!Prevta, glob)
rscli.Close
Exit Sub
Error:
  MsgBox "ERROR AL OBTENER DATOS GENERALES" & Err.Description, vbCritical
End Sub

Private Sub INICIALIZAVALS()
NVENTA1 = 0: NIVA1 = 0: NIEPS1 = 0
NVENTA2 = 0: NIVA2 = 0: NIEPS2 = 0
NVENTA3 = 0: NIVA3 = 0: NIEPS3 = 0
NVENTA4 = 0: NIVA4 = 0: NIEPS4 = 0
NVENTA5 = 0: NIVA5 = 0: NIEPS5 = 0
NVENTA6 = 0: NIVA6 = 0: NIEPS6 = 0
NVENTA7 = 0: NIVA7 = 0: NIEPS7 = 0
NVENTA8 = 0: NIVA8 = 0: NIEPS8 = 0
End Sub

Private Sub encabezadofacold()
On Error GoTo Error:
Printer.ScaleMode = vbCentimeters
Printer.FontName = "ARIAL NARROW"
Printer.FontSize = 8
'Printer.FontName = "ARIAL"
'Printer.FontSize = 7
Printer.FontBold = False
If sert = "G2" Or sert = "H2" Then
   Printer.CurrentY = 1.8
   Printer.CurrentX = 5
   Printer.Print "AHORA UBICADO EN: CALLE PRIMERA ORIENTE # 132, SECTOR REFORMA."
End If
'Printer.FontName = "ARIAL NARROW"
'Printer.FontName = "ARIAL"
'Printer.FontSize = 8
Printer.CurrentY = 3.5 + nAlto
Printer.CurrentX = 0.2
Printer.Print (nomt);
'Printer.CurrentX = 10
Printer.FontSize = 10
Printer.CurrentX = 6.5
Printer.Print (rfct)
Printer.FontSize = 8
Printer.CurrentX = 0.2
Printer.Print (dirt);
'Printer.CurrentX = 10
Printer.CurrentX = 6.5
Printer.Print telr;
If Trim(cmbAgente) <> "" Then
   'Printer.CurrentY = 3.7 + nAlto
   Printer.CurrentX = 8.6
   Printer.Print "A: " & Mid(cmbAgente.Text, 1, InStr(1, cmbAgente.Text, " "))
Else
   Printer.Print " "
End If
Printer.CurrentX = 0.2
Printer.Print (colt);
Printer.CurrentX = 6.5
Printer.Print (ciut);
Printer.CurrentX = 8.6
Printer.Print "U: " & Trim(cUsuario)

'Printer.CurrentY = 3.5
Printer.CurrentX = 0.2
Printer.Print (sert) & " " & numfac;
'Printer.CurrentY = 4 + nAlto
'Printer.CurrentY = 4.4 + nAlto

If Trim(cmbChofer) <> "" Then
    'Printer.CurrentY = 5 + nAlto
    Printer.CurrentX = 8.6
    Printer.Print "C: " & Mid(cmbChofer.Text, 1, InStr(1, cmbChofer.Text, " "));
End If
Printer.CurrentY = 4
Printer.CurrentX = 10.8

Printer.FontBold = True  'Se imprime mas grande y en negrita la fecha
Printer.FontSize = 15
Printer.Print Format(date, "DD/MM/YY");
Printer.FontBold = False
Printer.FontSize = 8
Printer.Print Space(1) & Format(Time, "HH:MM AM/PM")
Exit Sub
Error:
   MsgBox "Ocurrio un error al imprmir el encabezado...", vbCritical
End Sub

Private Sub encabezadofacoldchica()
On Error GoTo Error:
Printer.FontName = "ARIAL"
Printer.FontSize = 7

Printer.CurrentY = 3.7 + nAlto
Printer.CurrentX = 0.5
Printer.Print (nomt);
Printer.CurrentX = 10
Printer.Print (rfct)
Printer.CurrentX = 0.5
Printer.Print (dirt);
Printer.CurrentX = 10
Printer.Print telr
Printer.CurrentX = 0.5
Printer.Print (colt);
Printer.CurrentX = 10
Printer.Print (ciut)
Printer.CurrentX = 0.5
Printer.Print (sert) & " " & numfac

Printer.CurrentY = 4.8 + nAlto
Printer.CurrentX = 13.5
Printer.Print "U: " & Trim(cUsuario)
'Printer.CurrentY = 5 + nAlto
If Trim(cmbAgente) <> "" Then
    Printer.CurrentX = 13.5
    Printer.Print "A: " & Mid(cmbAgente.Text, 1, InStr(1, cmbAgente.Text, " "))
End If
If Trim(cmbChofer) <> "" Then
    Printer.CurrentX = 13.5
    Printer.Print "C: " & Mid(cmbChofer.Text, 1, InStr(1, cmbChofer.Text, " "));
End If
Printer.CurrentX = 19
Printer.Print Format(date, "DD/MM/YY") & Space(3) & Format(Time, "HH:MM AM/PM")
Printer.CurrentY = 8.3 + nAlto
Exit Sub
Error:
   MsgBox "Ocurrio un error al imprmir el encabezado...", vbCritical
End Sub

Private Function VERIFICA_FACTURA() As Boolean
On Error GoTo Error:
Dim rsFac As ADODB.Recordset
Set rsFac = New ADODB.Recordset
rsFac.Open "SELECT * FROM FACVENTA WHERE SERIE = '" & Trim(SERIE) & "' AND NumFactura = '" & Trim(numfac) & "'", cn, adOpenKeyset, adLockBatchOptimistic, adCmdText
If rsFac.EOF Then
    VERIFICA_FACTURA = False
Else
    VERIFICA_FACTURA = True
    'MsgBox "LA SERIE Y EL NUMERO DE FACTURA YA SE IMPRIMIO " & Chr(13) & "EN LA VENTA CON FOLIO UNICO " & rsFac!NOVENTA & " DE FECHA " & rsFac!Facfecha, vbExclamation
End If
rsFac.Close
Exit Function
Error:
   VERIFICA_FACTURA = True
   MsgBox "Se genero un error al validar la factura....", vbCritical
End Function

Private Sub cmdrptvta_Click()
Dim RESP As String
On Error GoTo Error:
RESP = InputBox("Proporcione el folio único de la venta a imprimir" & Chr(13) & Chr(13) & " 0 => Imprime todos los clientes de la preventa" & Chr(13) & "-1 => Imprime los clientes modificados", "Teclee folio", adopreventa.Recordset!noventa)
If IsNumeric(RESP) Then
   imppvta Val(RESP)   'Imprime ticket's de preventa
End If
Exit Sub:
Error:
   MsgBox Err.Description
End Sub


Private Sub DbgDetVta_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lTrans As Boolean
Dim dfecha As Date
On Error GoTo Error:
lTrans = False
Select Case KeyCode
   Case 113 'Modificar precio de venta
        nOpcion = 0: nValAnt = 0
        txtContra.Text = ""
        fraCon.Visible = True
        txtContra.SetFocus
   Case 114 'Cancelar producto de la venta
        If Me.chkLiquida.Value = 1 And (chkCredito.Value = 1 Or chkPrev.Value = 1) And Sql Then
           MsgBox "YA NO ES POSIBLE CANCELAR PRODUCTO PORQUE LA PREVENTA YA FUE LIQUIDADA", vbInformation, "Preventa liquidada"
           Exit Sub
        End If
        If AdoVentas.Recordset!situacion = "2" Or AdoVentas.Recordset!situacion = "3" Then
           MsgBox "NO ES POSIBLE CANCELAR PRODUCTOS" & Chr(13) & "PORQUE LA VENTA YA FUE FACTURADA" & Chr(13) & "CANCELA LA FACTURA Y GENERE UNA NUEVA VENTA", vbInformation, "Ventas"
           Exit Sub
        ElseIf AdoVentas.Recordset!situacion = "1" And Trim(cBorrar) = "" Then
           nOpcion = 4
           fraCon.Visible = True: txtContra.Text = ""
           txtContra.SetFocus
           Exit Sub
        End If
          
          If MsgBox("REALMENTE DESEAS BORRAR " & AdoDetVta.Recordset!cantidad & " PRODUCTOS DE" & Chr(13) & _
             AdoDetVta.Recordset!descripc & Chr(13) & Space(15) & AdoDetVta.Recordset!medida, vbYesNo + vbInformation, "Ventas") = vbNo Then Exit Sub
             Dim varBmk As Variant
             varBmk = AdoDetVta.Recordset.Bookmark
             If Sql Then cn.BeginTrans
             lTrans = True
             'Se guarda en un historial de cancelaciones
             If Sql Then cn.Execute "INSERT INTO prodborra(clave,noventa,fecha,cajas,piezas,importe,usuario,prevliq,preventa) VALUES ('" & AdoDetVta.Recordset!cl_producto & "'," & AdoVentas.Recordset!noventa & ",'" & date + Time & "'," & AdoDetVta.Recordset!cantidad * -1 & "," & AdoDetVta.Recordset!cantidadp * -1 & "," & AdoDetVta.Recordset!importe * -1 & ",'" & cBorrar & "'," & chkPrevLiq.Value & "," & IIf(txtcampos(2).Text = "", "0", txtcampos(2).Text) & ")"
             cn.Execute "DELETE FROM Ventas_det WHERE cl_producto = '" & DbgDetVta.Columns(0).Text & "' AND noventa = " & AdoVentas.Recordset!noventa
             cn.Execute "UPDATE Inventario SET InCant = Incant + " & AdoDetVta.Recordset!cantidad & ", InCantPza = IncantPza + " & AdoDetVta.Recordset!cantidadp & " WHERE Inprod = '" & AdoDetVta.Recordset!cl_producto & "'"
             If Format(AdoVentas.Recordset!fecha, "dd-mm-yyyy") < Format(date, "dd-mm-yyyy") And Sql Then
                dfecha = Format(DateAdd("d", 1, AdoVentas.Recordset!fecha), "dd-mm-yyyy")
                While dfecha <= Format(date, "dd-mm-yyyy")
                   diaini = "dia" & CStr(Day(dfecha))
                   cn.Execute "UPDATE invcorte SET " & diaini & " = " & diaini & " + " & AdoDetVta.Recordset!cantidad & " WHERE producto = '" & DbgDetVta.Columns(0).Text & "' AND mes = " & Month(date)
                   dfecha = DateAdd("d", 1, dfecha)
                Wend
             End If
             If Sql Then cn.CommitTrans
             lTrans = False
             rsttemp.Open "SELECT SUM(Importe) AS Subto FROM VENTAS_DET WHERE NOVENTA = " & AdoVentas.Recordset!noventa, cn, adOpenKeyset, adLockOptimistic, adCmdText
             'rsttemp.Open "SELECT SUM(Importe) AS Subto, SUM( CASE CLAPROVE WHEN 'C52' THEN IMPORTE END) AS PROMO FROM VENTAS_DET, TFPRODUC WHERE consec = cl_producto AND NOVENTA = " & AdoVentas.Recordset!noventa, cn, adOpenKeyset, adLockOptimistic, adCmdText
             If IsNull(rsttemp!subto) Then
                lblImpVta.Caption = "$ 0.00"
                'lblpromo.Caption = "PROM:  $ 0.00      BOLETOS: 0"
                nImpVta = 0
             Else
                lblImpVta.Caption = Format(rsttemp!subto, "$ ###,###,##0.00")
                'lblpromo.Caption = IIf(IsNull(rsttemp!PROMO), "PROM:  $ 0.00    BOLETOS: 0", "PROM.:  " & Format(rsttemp!PROMO, "$ ###,###,##0.00") & "      BOLETOS: " & Int(rsttemp!PROMO / 50))
                nImpVta = rsttemp!subto
             End If
             cn.Execute "UPDATE ventas SET MontoTotal = " & nImpVta & " WHERE Noventa = " & AdoVentas.Recordset!noventa
             rsttemp.Close
             AdoDetVta.Refresh
             If AdoDetVta.Recordset.RecordCount >= varBmk Then AdoDetVta.Recordset.Bookmark = varBmk
             If (chkCredito.Value = 1 Or chkPrev.Value) And Sql Then      'En preventas se imprimen solamente los pedidos editados
                AdoVentas.Recordset!situacion = 0
                AdoVentas.Recordset.Update
             End If
   Case 117 'F6 Vender mas piezas de los que trae una caja
        nOpcion = 2: nValAnt = 0
        txtContra.Text = ""
        fraCon.Visible = True
        txtContra.SetFocus
End Select
Exit Sub
Error:
  MsgBox "OCURRIO EL SIGUIENTE ERROR: " & Err.Description & Chr(13) & "A CONTINUACION SE CANCELARAN LOS CAMBIOS REALIZADOS"
  If lTrans And Sql Then cn.RollbackTrans
End Sub


Private Sub DbgrdPreventa_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
If adopreventa.Recordset!situacion = "3" Then
   MsgBox "NO SE PUEDE MODIFICAR LA VENTA PORQUE YA FUE FACTURADA", vbInformation, "Ventas"
   Cancel = True
   Exit Sub
End If
If IsNull(adopreventa.Recordset!crfc) Or Len(adopreventa.Recordset!crfc) <= 8 Then
   MsgBox "ESTE CLIENTE NO TIENE CAPTURADO RFC", vbInformation, "Ventas"
   Cancel = True
End If
End Sub

Private Sub DbgrdPreventa_DblClick()
  frmCliente.Show
End Sub

Private Sub DbgrdPreventa_HeadClick(ByVal ColIndex As Integer)
adopreventa.RecordSource = "SELECT  * FROM VENTAS V,CATCLIENTE WHERE V.FOLPREVENTA = " & Trim(txtcampos(2).Text) & " AND V.CLCLIENTE = CCLAVE ORDER BY " & Me.DbgrdPreventa.Columns(ColIndex).DataField
adopreventa.Refresh
End Sub

Private Sub Form_Activate()
frmPrincipal.Hide
If tipotienda = 3 Then  'En tiendas normales
    Me.lbletiquetas(12).Caption = "BODEGA"
    lbletiquetas(8).Caption = "MAYOREO"
    lbletiquetas(9).Caption = "MEDIO M"
    cmdticket.Visible = True
    cmdticket.Enabled = True
Else
    cmdticket.Enabled = False
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Error:
Dim rs As ADODB.Recordset
If Shift = vbCtrlMask And KeyCode = vbKeyC Then
   If AdoVentas.Recordset!FOLPREVENTA > 0 And chkPrevLiq.Value = 0 Then
      RESP = InputBox("Teclea la clave del chofer", "Clave")
      cn.Execute "UPDATE ventas SET chofer = " & RESP & " FROM catcliente WHERE cclave = " & RESP & " and ctipo = 2 And folpreventa = " & txtcampos(2).Text
      MsgBox "Cambio completado", vbInformation
      Unload Me
   End If
ElseIf Shift = vbCtrlMask And KeyCode = vbKeyP Then
      nOpcion = 6: nValAnt = 0
      txtContra.Text = ""
      fraCon.Visible = True
      txtContra.SetFocus
End If
Select Case KeyCode
   Case 113 'Modificar precio de venta     [F2]
        'If Not lVta Then Exit Sub
        nOpcion = 0: txtContra.Text = "": nValAnt = 0
        fraCon.Visible = True
        txtContra.SetFocus
   Case 115 'Cobrar venta      [F4]
        If tipotienda = 3 Then
           nOp = 3
           lCob = True
           Me.stb1.Panels(4).Enabled = True
        End If
        
        If lCob Then
           If frmVentas.AdoVentas.Recordset.State = 0 Then
              AdoVentas.Refresh
           End If
           frmCobra.txtsubtotal = Format(lblImpVta.Caption, "########0.00")
           frmCobra.txtcheques = Format(IIf(IsNumeric(AdoVentas.Recordset!cheques), AdoVentas.Recordset!cheques, 0), "########0.00")
           frmCobra.txtefectivo = Format(IIf(IsNumeric(AdoVentas.Recordset!efectivo), AdoVentas.Recordset!efectivo, 0), "########0.00")
           frmCobra.txtvales = Format(IIf(IsNumeric(AdoVentas.Recordset!VALES), AdoVentas.Recordset!VALES, 0), "########0.00")
           impte = AdoVentas.Recordset!VALES + AdoVentas.Recordset!cheques + AdoVentas.Recordset!efectivo + AdoVentas.Recordset!tdc
           frmCobra.txtImporte = Format(IIf(IsNumeric(impte), impte, 0), "########0.00")
           If AdoVentas.Recordset!situacion = 2 Or AdoVentas.Recordset!situacion = 3 Then
              frmCobra.txtCambio = Format(AdoVentas.Recordset!total - Format(lblImpVta.Caption, "########0.00"), "###,###,##0.00")
              frmCobra.cmdGrabar.Enabled = False
           'ElseIf AdoVentas.Recordset!CREDITO Then
           '   MsgBox "NO ES POSIBLE COBRAR EN ESTA OPCION VENTAS A CREDITO" & Chr(13) & "HAGALO MEDIANTE LA OPCION HISTORIAL Y ABONOS PARCIALES", vbInformation
           '   Exit Sub
           Else
              frmCobra.txtCambio.Text = "0.00"
           End If
           If rscli.State = 1 Then
              rscli.Requery
           Else
              rscli.Open "SELECT * FROM Catcliente WHERE cClave = '" & txtcampos(4).Text & "' AND ctipo = 0", cn, adOpenKeyset, adLockOptimistic, adCmdText
           End If
           'frmCobra.chkOficial.Visible = IIf(IsNull(rscli!CPAGOCHEQUE), False, rscli!CPAGOCHEQUE)
           frmCobra.stb1.Panels(3).Enabled = IIf(IsNull(rscli!cpagocheque), False, rscli!cpagocheque)
           frmCobra.Show
        End If
        
   Case 116    'Ver pedidos de preventa     [F5]
        If stb1.Panels(5).Enabled Then
            cmdExporta.Visible = Not Sql
            frapreventa.Visible = True
            adopreventa.ConnectionString = cCadConex
            adopreventa.CursorType = adOpenKeyset
            adopreventa.RecordSource = "SELECT  * FROM VENTAS V,CATCLIENTE WHERE V.FOLPREVENTA = " & Trim(txtcampos(2).Text) & " AND V.CLCLIENTE = CCLAVE ORDER BY noventa"
            adopreventa.Refresh
            Set rs = New ADODB.Recordset
            rs.Open "SELECT sum(d.importe) as  TOTVTA FROM ventas_det d, ventas v WHERE v.noventa = d.noventa AND d.cancelado = 0 AND v.folpreventa = " & txtcampos(2).Text, cn, adOpenDynamic, adLockOptimistic, adCmdText
            TxtTotVta.Text = IIf(IsNull(rs!TotVta), "$ 0.00", Format(rs!TotVta, "$ ###,###,##0.00"))
            rs.Close
            Set rs = Nothing
            Me.cmdFacCte.Enabled = Not ModVta
            Me.cmdFacConFin.Enabled = Not ModVta
        End If
        cBorrar = ""
   Case 117    'Cambio de precios a mas altos solamente  [F6]
        fraCon.Visible = True
        txtContra.Text = ""
        txtContra.SetFocus
        nOpcion = 3: nValAnt = 0
   Case 118    'Vender mas piezas de los que trae una caja    [F7]
        nOpcion = 2: nValAnt = 0
        txtContra.Text = ""
        fraCon.Visible = True
        txtContra.SetFocus
   Case 119    'Calculadora     [F8]
        frmCalc.Show 1
   Case 120    'Obtener nuevamente productos de Inventario    [F9]
        If AdoVentas.Recordset!situacion = "2" Or AdoVentas.Recordset!situacion = "3" Then
           MsgBox "NO ES POSIBLE AGREGAR PRODUCTOS" & Chr(13) & "PORQUE LA VENTA YA FUE FACTURADA" & Chr(13) & "CANCELA LA FACTURA Y GENERE UNA NUEVA VENTA", vbInformation, "Ventas"
        Else
           Obtenprod  'Obtiene productos con existencia
        End If
   Case 121    'Importa pedidos de agentes de venta   [F10]
        ImpPvtaH
   Case 122    'Ajusta precios de caja basandose en precios de autoservicio [F11]
        MsgBox "OPCION NO DISPONIBLE, CONSULTE CON EL ADMINISTRADOR DEL SISTEMA", vbInformation, "No disponible"
        'If MsgBox("CONFIRMA SI DESEAS AJUSTAR PRECIOS BASANDOSE EN PRECIOS DE AUTOSERVICIO", vbQuestion + vbYesNo, "Ajustar precios") = vbYes Then
        '   cn.Execute "UPDATE ventas_det SET PRECIO = precio1 * paquetes, IMPORTE = cantidad * (precio1 * paquetes) + (cantidadp * preciop) " & _
        '              "FROM tfproduc, preprod WHERE consec = preclave AND consec = cl_producto AND cl_producto = preclave AND " & _
        '              "NOVENTA = " & AdoVentas.Recordset!noventa
        '   cn.Execute "UPDATE VENTAS set montototal = ( SELECT SUM(importe) FROM ventas_det WHERE noventa = " & AdoVentas.Recordset!noventa & ") WHERE noventa = " & AdoVentas.Recordset!noventa
        '   MsgBox "El ajuste de precios se realizo correctamente", vbInformation, "Ventas"
        'End If
   Case 123   'Para preventistas de PHILIP MORRIS MEXICO Costo + 1 %  [F12]
        nOpcion = 5: nValAnt = 0
        txtContra.Text = ""
        fraCon.Visible = True
        txtContra.SetFocus
End Select
Exit Sub
Error:
  MsgBox "Probablemente Otro Usuario ha utilizado esta Venta, Salga y vuelva a entrar a la Venta"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     SendKeys vbTab
     KeyAscii = 0
  End If
End Sub

Private Sub Form_Load()
On Error GoTo Error:
 TipVta = Array("1   MOSTRADOR", "2   AGENTE DIRECTO", "3   VIA TELEFONICA BODEGA", "4   VIA TELEFONICA ENVIO", "5   CLIENTES ESPECIALES", "6   PREVENTA", "7   OTROS")
 'cn.BeginTrans
 Set rs = New ADODB.Recordset
 For N = 0 To 6
     cmbTipVta.AddItem TipVta(N)
 Next
 
 CarCte = True
 AdoVentas.CursorType = adOpenKeyset
 AdoVentas.ConnectionString = cCadConex

 If nOp = 0 Then
    AdoVentas.RecordSource = "SELECT * FROM ventas WHERE noventa = 1"
    AdoVentas.Refresh
    AdoVentas.Recordset.AddNew
    
    txtcampos(0).Text = date & " " & Time
    chkCredito.Value = 0
    chkConRfc.Value = 0
    'chkVales.Value = 0
    chkPrev.Value = 0
 Else 'Opcion Modificar Y Cobrar Venta
    AdoVentas.RecordSource = "SELECT * FROM VENTAS WHERE noventa = " & frmPrincipal.dbgrdVta.Columns(0).Text
    AdoVentas.Refresh
    'chkCredito.Enabled = False
    fraGenerales.Caption = "Datos generales de la venta   [ " & frmPrincipal.dbgrdVta.Columns(0).Text & " ]"
    If nOp = 3 Then SendKeys vbTab: SendKeys vbTab: SendKeys vbTab: SendKeys vbTab: SendKeys vbTab: SendKeys vbTab: SendKeys vbTab
    If nOp = 1 Then
       SendKeys vbTab: SendKeys vbTab: SendKeys vbTab
       txtcampos(1).Locked = True: txtcampos(9).Locked = True
       txtcampos(3).Locked = True: cmbAgente.Locked = True
       cmbChofer.Locked = True: cmbTipVta.Locked = True
    End If
 End If
 lVta = ModVta
 lCob = Not ModVta
 stb1.Panels(4).Enabled = Not lVta   'No se permite cobrar en los Modulos de venta
 stb1.Panels(2).Enabled = (nOp <> 3) 'En la opcion cobro no se puede modificar precio ni cancelar producto
 stb1.Panels(3).Enabled = (nOp <> 3)
 stb1.Panels(5).Enabled = IIf(IsNull(AdoVentas.Recordset!FOLPREVENTA) Or AdoVentas.Recordset!FOLPREVENTA <> 0, True, False)
 lblImpVta.Visible = ModVta
 'cn.CommitTrans
 SendMessage cmbprod.hwnd, &H160, 510, 0

Exit Sub
Error:
  MsgBox "OCURRIO EL SIGUIENTE ERROR: " & Err.Description & Chr(13) & "A CONTINUACION SE CANCELARAN LOS CAMBIOS REALIZADOS"
  'cn.RollbackTrans
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
cn.CommitTrans
frmPrincipal.Show
End Sub


Private Sub txtCampos_GotFocus(Index As Integer)
Select Case Index
Case 6, 13, 14
  txtcampos(Index).SelStart = 0
  txtcampos(Index).SelLength = Len(txtcampos(Index).Text)
  nValAnt = Val(txtcampos(Index).Text)
Case 7, 8
     If nValAnt = 0 Then nValAnt = Val(txtcampos(Index).Text)
End Select
End Sub

Private Sub txtCampos_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 27 Then
   Unload Me
ElseIf KeyAscii = 13 Then
   KeyAscii = 0
   SendKeys vbTab
End If
End Sub

Private Sub txtCampos_LostFocus(Index As Integer)
Dim nTipVta
Dim rspre As ADODB.Recordset
Dim rsttemp As ADODB.Recordset
On Error GoTo Error:
Select Case Index
   Case 1  'Tipo de Venta
        txtcampos(Index).Text = Trim(txtcampos(Index).Text)
        txtcampos(Index).Refresh
        nPos = InStr(1, "123456", txtcampos(Index).Text)
        If nPos = 0 Or txtcampos(Index).Text = "" Then
            cmbTipVta.SetFocus
            Exit Sub
        End If
        cmbTipVta.Text = TipVta(nPos - 1)
        nTipVta = Val(Mid(cmbTipVta.Text, 1, 3))
        'Desactivado en ventas telefonicas
        'chkCredito.Enabled = (nTipVta <> 3 And nOp <> 1)
        
        'Visible en ventas por Agente
        If nTipVta = 2 Then
           Set rsttemp = New ADODB.Recordset
           rsttemp.Open "SELECT * FROM CATCLIENTE WHERE ctipo = 1", cn, adOpenKeyset, adLockOptimistic, adCmdText
           While Not rsttemp.EOF
              cmbAgente.AddItem rsttemp!cNombre
              rsttemp.MoveNext
           Wend
           If nOp = 0 Then chkPrev.Value = 1
           'chkCredito.Enabled = False
        ElseIf nTipVta = 1 Then
           chkCredito.Enabled = False
           chkPrev.Enabled = False
        End If
        lbletiquetas(2).Visible = (nTipVta = 2)
        txtcampos(9).Visible = (nTipVta = 2)
        cmbAgente.Visible = (nTipVta = 2)
        
        'Visible en ventas Venta X telefono envio y Preventa
        lbletiquetas(3).Visible = (nTipVta = 4 Or nTipVta = 5 Or nTipVta = 2)
        txtcampos(3).Visible = (nTipVta = 4 Or nTipVta = 5 Or nTipVta = 2)
        cmbChofer.Visible = (nTipVta = 4 Or nTipVta = 5 Or nTipVta = 2)
        
        lbletiquetas(4).Visible = True
        txtcampos(4).Visible = True
        cmbCliente.Visible = True
        lblCte.Visible = True
        txtCte.Visible = True
        If nTipVta = 2 Then
           txtcampos(9).SetFocus
        ElseIf (nTipVta = 4 Or nTipVta = 5) Then
           txtcampos(3).SetFocus
        Else
           'txtCampos(4).SetFocus
           txtCte.SetFocus
        End If
        
   Case 4 'Clave del cliente
        Set rscli = New ADODB.Recordset
        ccVeCli = IIf(Trim(txtcampos(4).Text) = "", "0", txtcampos(4).Text)
        If rscli.State = adStateOpen Then rscli.Close
        rscli.Open "SELECT * FROM Catcliente WHERE cClave = " & ccVeCli & " AND ctipo = 0", cn, adOpenStatic, adLockOptimistic, adCmdText
        If rscli.BOF And rscli.EOF Then
           cmbCliente.SetFocus
           Exit Sub
        End If
        Me.cmbCliente = rscli!cNombre
        If Trim(cmbCliente.Text) = "" Or Trim(cmbCliente.Text) = "CONSUMIDOR FINAL" Or Trim(cmbCliente.Text) = "C O N S U M I D O R     F I N A L" Then
           MsgBox "ES NECESARIO ESPECIFICAR EL NOMBRE DEL CLIENTE", vbInformation
           Exit Sub
        End If
        txtcampos(4).Text = rscli!cclave
        cmbCliente.Text = rscli!cNombre
        lblNombre.Caption = IIf(IsNull(rscli!cnombrefac), "", rscli!cnombrefac)
        lblDirec.Caption = IIf(IsNull(rscli!cdireccion), "", rscli!cdireccion)
        lblRfc.Caption = IIf(IsNull(rscli!crfc), "", rscli!crfc)
        lblTelefono.Caption = rscli!ctelefono

        cmdGrabar.Visible = True
        cmdregresar.Visible = True
        cmdClientes.Visible = True
        cmdTicBod.Visible = True
        Me.cmdCamSer.Visible = True
        cmdGrabar.SetFocus
        If nOp = 3 Then
           cmdGrabar_Click
           DbgDetVta.SetFocus
        End If
   Case 7, 8 'Precio unitario del producto, cuando a traves de contraseña se permite modificar
        If nOpcion = 0 Then   'Cambio de precios mas alto o mas bajo
            Set rspre = New ADODB.Recordset
            rspre.Open "SELECT Precosto FROM TFPRODUC WHERE Consec = '" & cprod & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
            If rspre.RecordCount > 0 Then
                nPreBaj = IIf(AdoVentas.Recordset!credito Or AdoVentas.Recordset!Prevta, txtcampos(8), txtcampos(7).Text)
                'nPorcMin = IIf(Trim(Mid(cSucursal, 1, 3)) = "28", 0.01, 0.005)
                If Not Sql Then   'Version Portatil Preventistas Laptop
                   nPorcMin = 0.015
                ElseIf Trim(Mid(cSucursal, 1, 3)) = "28" Then   'Istmo
                   nPorcMin = 0.015
                Else
                   nPorcMin = 0.005                             'Todas las demas Bodegas
                End If
                If Round(rspre!PRECOSTO + (rspre!PRECOSTO * nPorcMin), 2) > Val(nPreBaj) And Val(txtcampos(7).Text) <> Val(txtcampos(8).Text) Then
                   MsgBox "EL PRECIO MINIMO PARA VENTA ES: " & Round(rspre!PRECOSTO + (rspre!PRECOSTO * nPorcMin), 2), vbInformation
                   txtcampos(Index).SetFocus
                   Exit Sub
                ElseIf Val(nPreBaj) < Val(txtcampos(7).Text) Then
                    MsgBox "ADVERTENCIA: ESTA ASIGNANDO PRECIO A CREDITO" & Chr(13) & "MENOR QUE AL PUESTO EN BODEGA: " & txtcampos(7).Text, vbInformation, "Confirmación de precio"
                End If
                'Se graba el usuario solamente cuando el precio es menor al sugerido
                If nValAnt <= Val(txtcampos(Index).Text) Then cModifico = ""
            End If
        ElseIf nOpcion = 3 Then 'solo precios mas altos
            If nValAnt > Val(txtcampos(Index).Text) Then
               MsgBox "NO ESTA AUTORIZADO PARA MODIFICAR A PRECIOS MAS BAJOS DE: " & nValAnt, vbInformation
               txtcampos(Index).SetFocus
               Exit Sub
            End If
            cModifico = ""
        End If
        txtcampos(7).Enabled = False
        txtcampos(8).Enabled = False
        txtcampos(10).Enabled = False
        CmdAgregar.Enabled = True
        'txtcampos(6).SetFocus
   Case 10
        Dim RSTPRE As ADODB.Recordset
        Set RSTPRE = New ADODB.Recordset
        RSTPRE.Open "SELECT Precosto,PAQUETES FROM TFPRODUC WHERE Consec = '" & cprod & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
        If (RSTPRE!PRECOSTO / RSTPRE!PAQUETES) * 1.015 >= Val(txtcampos(Index).Text) Then
           MsgBox "EL PRECIO MINIMO PARA VENTA ES: " & Round(RSTPRE!PRECOSTO / RSTPRE!PAQUETES + (RSTPRE!PRECOSTO / RSTPRE!PAQUETES * 0.015), 2), vbInformation
           txtcampos(Index).SetFocus
           Exit Sub
        End If
        Set RSTPRE = Nothing
        txtcampos(Index).Enabled = False
        txtcampos(7).Enabled = False
        txtcampos(8).Enabled = False
        txtcampos(10).Enabled = False
        CmdAgregar.Enabled = True
        
   Case 13  'Piezas vendidas
      If CmdAgregar.Enabled = True Then
         rs.Requery
         If Val(txtcampos(Index).Text) >= rs!PAQUETES And rs!PAQUETES <> 1 And Not lMasPza Then
            MsgBox "LA CANTIDAD EN PIEZAS NO PUEDE SER MAYOR O IGUAL A" & Chr(13) & "LOS PAQUETES QUE TRAE LA CAJA, AUMENTE LA CANTIDAD VENDIDA EN CAJAS", vbInformation, "Ventas"
            txtcampos(Index).Text = 0
            txtcampos(6).SetFocus
            Exit Sub
         End If
      End If
  Case 9  'Clave del agente
        Set rsttemp = New ADODB.Recordset
        rsttemp.Open "SELECT * FROM Catcliente WHERE cClave = " & IIf(Trim(txtcampos(Index).Text) = "", 0, Trim(txtcampos(Index).Text)) & " AND ctipo = 1", cn, adOpenKeyset, adLockOptimistic, adCmdText
        If rsttemp.EOF And rsttemp.BOF Then
           cmbAgente.SetFocus
           Exit Sub
        End If
        txtcampos(Index).Text = rsttemp!cclave
        cmbAgente.Text = rsttemp!cNombre
        nPos = InStr(1, "25", txtcampos(Index).Text)
        If nPos > 0 Then
           'chkCredito.Value = 1
           'chkCredito.Enabled = False
        End If
        'lbletiquetas(2).BorderStyle = IIf(rsttemp!franquicia, 1, 0)
        Set rsttemp = New ADODB.Recordset
        rsttemp.Open "SELECT * FROM CATCLIENTE WHERE ctipo = 2", cn, adOpenKeyset, adLockOptimistic, adCmdText
        cmbChofer.Clear
        While Not rsttemp.EOF
            cmbChofer.AddItem rsttemp!cNombre
            rsttemp.MoveNext
        Wend
        txtcampos(3).SetFocus
  Case 3   'Clave del chofer
        Set rsttemp = New ADODB.Recordset
        rsttemp.Open "SELECT * FROM Catcliente WHERE cClave = " & IIf(Trim(txtcampos(Index).Text) = "", 0, Trim(txtcampos(Index).Text)) & " AND ctipo = 2", cn, adOpenKeyset, adLockOptimistic, adCmdText
        If rsttemp.BOF And rsttemp.EOF Then
           cmbChofer.SetFocus
           Exit Sub
        End If
        txtcampos(Index).Text = rsttemp!cclave
        cmbChofer.Text = rsttemp!cNombre
        txtCte.SetFocus
End Select
Exit Sub
Error:
   MsgBox Err.Description
End Sub

Private Sub cmbTipVta_Validate(Cancel As Boolean)
 If InStr("123456", Mid(cmbTipVta.Text, 1, 1)) = 0 Then
   MsgBox "Debe seleccionar un tipo de venta de la lista desplegable", vbExclamation
   cmbTipVta.SetFocus
   Cancel = True
   Exit Sub
 End If
 txtcampos(1).Text = Mid(cmbTipVta.Text, 1, 1)
 txtcampos(1).SetFocus
End Sub

Private Sub txtcheques_GotFocus()
   txtcheques.SelLength = 8
End Sub

Private Sub txtcheques_LostFocus()
txtcheques.Text = Format(txtcheques.Text, "###,###,##0.00")
txtImporte.Text = Format(Val(txtvales.Text) + Val(txtcheques.Text) + Val(txttdc.Text) + Val(txtefectivo.Text), "###,###,###.00")
End Sub

Private Sub txtContra_GotFocus()
txtContra.SelLength = Len(txtContra.Text)
End Sub

Private Sub txtContra_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then cmdConCance_Click
If KeyAscii = 13 Then
   cmdConAceptar_Click
End If
End Sub


Private Sub txtefectivo_GotFocus()
   txtefectivo.SelLength = 8
End Sub

Private Sub txtefectivo_LostFocus()
  txtefectivo.Text = Format(txtefectivo.Text, "###,###,##0.00")
  txtImporte.Text = Format(Val(txtvales.Text) + Val(txtcheques.Text) + Val(txttdc.Text) + Val(txtefectivo.Text), "###,###,###.00")
  If chkCobrado.Value = 0 Then cmdCobrado.SetFocus
End Sub

Private Sub txttdc_GotFocus()
   txttdc.SelLength = 8
End Sub

Private Sub txttdc_LostFocus()
  txttdc.Text = Format(txttdc.Text, "###,###,##0.00")
  txtImporte.Text = Format(Val(txtvales.Text) + Val(txtcheques.Text) + Val(txttdc.Text) + Val(txtefectivo.Text), "###,###,###.00")
End Sub

Private Sub txtvales_GotFocus()
   txtvales.SelLength = 8
End Sub

Private Sub txtvales_LostFocus()
 txtvales.Text = Format(txtvales.Text, "###,###,##0.00")
 txtImporte.Text = Format(Val(txtvales.Text) + Val(txtcheques.Text) + Val(txttdc.Text) + Val(txtefectivo.Text), "###,###,###.00")
End Sub



Private Sub txtCte_LostFocus()
Dim rstcte As ADODB.Recordset
If Trim(txtCte.Text) <> "" Then
  Set rstcte = New ADODB.Recordset
  rstcte.Open "SELECT * FROM Catcliente WHERE cNombre like '%" & txtCte.Text & "%' AND ctipo = 0", cn, adOpenDynamic, adLockOptimistic, adCmdText
  cmbCliente.Clear
  While Not rstcte.EOF
     Me.cmbCliente.AddItem rstcte!cNombre
     rstcte.MoveNext
  Wend
  Set rstcte = Nothing
  cmbCliente.SetFocus
End If
End Sub

Function FacturaNva(Optional globconFin As Boolean) As Boolean
Dim rstDetVta As ADODB.Recordset
Dim rstDir As ADODB.Recordset
Dim Impresora As Printer
Dim nTotal
Dim nCajas
Dim lNvaFac As Boolean
Dim TotFac As Double
Dim LETRAS
Dim cimpresora As String
Dim lgraboprod As Boolean
Dim nTotVale As Currency
Dim nIniVta As Integer
Dim nNumVal As Integer
Dim lFormAnt As Boolean

On Error GoTo Error:
'compara  'En caso de ser Franquicia se hace el prorrateo

RFCFINAL = "COOF970101111"
FacturaNva = False
lImp = False
cimpresora = IIf(AdoVentas.Recordset!credito Or AdoVentas.Recordset!Prevta, "*CREDITO*", "*CONTADO*")
If SELECCIONA_IMPRESORA(cimpresora) Then
   MsgBox "PREPARE LA IMPRESORA " & cimpresora, vbInformation, "Ventas"
   lImp = True
Else
   lImp = False
   MsgBox "NO ES POSIBLE IMPRIMIR FACTURAS A " & cimpresora & Chr(13) & Chr(13) & "PORQUE NO EXISTEN IMPRESORAS INSTALADAS" & Chr(13) & "PARA ESTE TIPO DE FACTURAS", vbCritical
   Exit Function
End If

lNvaFac = True: nProd = 0
'nAlto = IIf(Trim(Mid(cSucursal, 1, 3)) <> "28" And Trim(Mid(cSucursal, 1, 3)) <> "16", -0.3, -0.6) 'Miguel cabrera
nAlto = IIf(Trim(Mid(cSucursal, 1, 3)) <> "28" And Trim(Mid(cSucursal, 1, 3)) <> "16", -0.3, -1.1) 'Miguel cabrera
nAlto = -1.3
'lFormAnt = Trim(Mid(cSucursal, 1, 3)) = "28"
lFormAnt = False 'Verdadero = Formato Horizontal, Falso = Vertical
ProdenFac = IIf(lFormAnt, 15, 35)
'Se Obtiene la serie correspondiente
SERIE = ObtenSerie(frmVentas.AdoVentas.Recordset!credito, frmVentas.AdoVentas.Recordset!Prevta, globconFin)

Set rstDetVta = New ADODB.Recordset
Call obtengenerales(globconFin)    'Obtiene generales de la factura (nombre,telefono,direciion,etc)
'SE CIERRA LA VENTA UNA VEZ OBTENIDO TODOS LOS DATOS NECESARIOS
noventar = AdoVentas.Recordset!noventa
If globconFin = True Then
   'Si es empacadora Gomez se factura a precio de Bodega y liquida a $ envío
   'If lbletiquetas(2).BorderStyle = 1 Then
   '   MsgBox "EL AGENTE DE LA PREVENTA ES EXTERNO, SE FACTURARA A PRECIO DE BODEGA", vbInformation, "Agente externo"
   '   cn.Execute "UPDATE ventas_det SET precosto = precio4, precostop = precio4 / t.paquetes  FROM ventas v, ventas_det d, preprod p, tfproduc T WHERE v.noventa = d.noventa and cl_producto = preclave and cl_producto = consec and preclave = consec and v.folpreventa = " & AdoVentas.Recordset!folpreventa
   '   rstDetVta.Open "SELECT MIN(CONSEC) as consec, sum(cantidad) as cantidad, sum(cantidadp) AS cantidadp, descripc, str(paquetes) + ' X ' + LTRIM(STR(CONTENID,10,3)) + ' ' + MEDIDA AS MEDIDA, MAX(D.IEPS) AS IEPS, MAX(D.IVA) AS IVA, AVG(D.Precosto) AS PRECIO, AVG(D.precosto / t.paquetes ) AS preciop , SUM((D.cantidad * d.precosto) + (d.cantidadp * d.precostop )) AS IMPORTE, MIN(FOLPREVENTA) AS PREVENTA, MAX(D.tasaieps) AS TasaIeps, AVG(t.precosto) AS costo, AVG(t.precosto / t.paquetes) AS costop  FROM VENTAS_DET D, TFPRODUC t,VENTAS V WHERE V.folpreventa = " & AdoVentas.Recordset!folpreventa & " AND D.Cl_producto = Consec AND d.Cancelado = 0 AND V.FACRFC = 0 AND D.NOVENTA = V.NOVENTA and d.facturado = 0 GROUP BY str(paquetes) + ' X ' + LTRIM(STR(CONTENID,10,3)) + ' ' +  MEDIDA,DESCRIPC", cn, adOpenStatic, adLockOptimistic, adCmdText
   'Else
      rstDetVta.Open "SELECT MIN(CONSEC) as consec, sum(cantidad) as cantidad, sum(cantidadp) AS cantidadp, descripc, str(paquetes) + ' X ' + LTRIM(STR(CONTENID,10,3)) + ' ' + MEDIDA AS MEDIDA, MAX(D.IEPS) AS IEPS, MAX(D.IVA) AS IVA, AVG(D.PRECIO) AS PRECIO, AVG(D.PRECIOP) AS preciop , SUM(D.IMPORTE) AS IMPORTE, MIN(FOLPREVENTA) AS PREVENTA, MAX(D.tasaieps) AS TasaIeps, AVG(d.precosto) AS costo, AVG(d.precostop) AS costop  FROM VENTAS_DET D, TFPRODUC,VENTAS V WHERE V.folpreventa = " & AdoVentas.Recordset!FOLPREVENTA & " AND D.Cl_producto = Consec AND d.Cancelado = 0 AND V.FACRFC = 0 AND v.prevta = 1 AND D.NOVENTA = V.NOVENTA and d.facturado = 0 GROUP BY str(paquetes) + ' X ' + LTRIM(STR(CONTENID,10,3)) + ' ' +  MEDIDA,DESCRIPC", cn, adOpenStatic, adLockOptimistic, adCmdText
   'End If
Else
   'If chkVales.Value = 1 Then
   '   rstDetVta.Open "SELECT * FROM valcheq WHERE empresa = '8  VALES MUNICIPIO' and NOVENTA = " & AdoVentas.Recordset!noventa, cn, adOpenForwardOnly, adLockOptimistic, adCmdText
   '   nNumVal = 0
   '   While Not rstDetVta.EOF
   '      nNumVal = nNumVal + 1
   '      rstDetVta.MoveNext
   '   Wend
   '   If nNumVal = 0 Then
   '      MsgBox "ES NECESARIO QUE EL CLIENTE PAGUE CON UN VALE DEL MUNICIPIO"
   '   End If
   '   rstDetVta.Close
   '   CADENA = "SELECT d.tasaieps, d.precosto as costo, d.precostop as costop, d.noventa, P.CONSEC, D.cantidad, D.cantidadp, P.descripc, str(P.paquetes) + ' X ' + LTRIM(STR(P.CONTENID,10,3)) + ' ' + P.MEDIDA AS MEDIDA, D.Iva, D.Ieps, D.Precio, D.preciop, D.Importe  FROM VENTAS_DET D, TFPRODUC P WHERE Noventa = " & noventar & " AND Cl_producto = Consec and  facturado = 0  ORDER BY IMPORTE "
   'Else
      CADENA = "SELECT d.tasaieps, d.precosto as costo, d.precostop as costop, d.noventa, P.CONSEC, D.cantidad, D.cantidadp, P.descripc, str(P.paquetes) + ' X ' + LTRIM(STR(P.CONTENID,10,3)) + ' ' + P.MEDIDA AS MEDIDA, D.Iva, D.Ieps, D.Precio, D.preciop, D.Importe  FROM VENTAS_DET D, TFPRODUC P WHERE Noventa = " & noventar & " AND Cl_producto = Consec and  facturado = 0  ORDER BY DESCRIPC,CONTENID "
   'End If
   rstDetVta.Open CADENA, cn, adOpenStatic, adLockOptimistic, adCmdText
End If
If rstDetVta.EOF Then
   MsgBox "Los Productos  ya fueron Facturados", vbInformation, "Ventas"
   Exit Function
End If
If AdoDetVta.Recordset.State = 1 Then
   AdoDetVta.Recordset.Close
   AdoVentas.Recordset.Close
End If
For i = 1 To 400
    NOVENTAT(i) = 0
    consecs(i) = " "
    descripcs(i) = " "
    medidas(i) = " "
    cantidads(i) = 0
    cantidadps(i) = 0
    costoss(i) = 0
    costosp(i) = 0
    precios(i) = 0
    preciosp(i) = 0
    importes(i) = 0
    tasas(i) = 0
    ivas(i) = 0
    iepss(i) = 0
Next
'Se buscan productos por la totalidad de los vales del Mun.
nTotVale = 0
i = 0: nIniVta = 0
While Not rstDetVta.EOF
    'If chkVales.Value = 1 Then
    '   nTotVale = nTotVale + rstDetVta!importe
    '   'temporalmente quedaria solo una venta por vale
    '   If nTotVale > (190 * nNumVal) And nIniVta = 0 Then
    '      i = i + 1
    '      consecs(i) = "1008833"
    '      descripcs(i) = "ABARROTES 1"
    '      medidas(i) = "1 X 1 PESOS"
    '      cantidads(i) = 1
    '      cantidadps(i) = 0
    '      costoss(i) = nTotVale - (190 * nNumVal)
    '      costosp(i) = 0
    '      precios(i) = nTotVale - (190 * nNumVal)
    '      preciosp(i) = 0
    '      importes(i) = nTotVale - (190 * nNumVal)
    '      tasas(i) = 1
    '      ivas(i) = 0
    '      iepss(i) = 0
    '      nIniVta = i
    '   ElseIf nIniVta > 0 Then
    '        i = i + 1
    '        NOVENTAT(i) = rstDetVta!noventa
    '        consecs(i) = rstDetVta!CONSEC
    '        descripcs(i) = rstDetVta!DESCRIPC
    '        medidas(i) = rstDetVta!medida
    '        cantidads(i) = rstDetVta!cantidad
    '        cantidadps(i) = rstDetVta!cantidadp
    '        costoss(i) = rstDetVta!costo
    '        costosp(i) = rstDetVta!costop
    '        precios(i) = rstDetVta!PRECIO
    '        preciosp(i) = rstDetVta!preciop
    '        importes(i) = rstDetVta!importe
    '        tasas(i) = rstDetVta!tasaieps
    '        ivas(i) = rstDetVta!iva
    '        iepss(i) = rstDetVta!ieps
    '   End If
    'Else
        i = i + 1
        If globconFin = False Then
            NOVENTAT(i) = rstDetVta!noventa
        End If
        If rstDetVta!tasaieps <= 0 Or rstDetVta!tasaieps >= 9 Or IsNull(rstDetVta!tasaieps) Then
           MsgBox "EL PRODUCTO " & rstDetVta!descripc & " " & rstDetVta!medida & "NO TIENE ESPECIFICADO EL DEPARTAMENTO, INFORME AL ADMINISTRADOR DEL SISTEMA", vbCritical
           Exit Function
        End If
        If rstDetVta!importe = 0 And (rstDetVta!cantidad > 0 Or rstDetVta!cantidadp > 0) And InStr(1, rstDetVta!descripc, "GRATIS") = 0 Then
           MsgBox "EL PRODUCTO " & rstDetVta!CONSEC & rstDetVta!descripc & " " & rstDetVta!medida & "NO TIENE PRECIO, INFORME AL AREA DE COSTOS", vbCritical
           Exit Function
        End If
        consecs(i) = rstDetVta!CONSEC
        descripcs(i) = rstDetVta!descripc
        medidas(i) = rstDetVta!medida
        cantidads(i) = rstDetVta!cantidad
        cantidadps(i) = rstDetVta!cantidadp
        costoss(i) = rstDetVta!costo
        costosp(i) = rstDetVta!costop
        precios(i) = rstDetVta!PRECIO
        preciosp(i) = rstDetVta!preciop
        importes(i) = rstDetVta!importe
        tasas(i) = rstDetVta!tasaieps
        ivas(i) = rstDetVta!iva
        iepss(i) = rstDetVta!ieps
    'End If
    rstDetVta.MoveNext
Wend

numfacturast = (i / ProdenFac)
numfacturas = Int(i / ProdenFac)
If numfacturast > numfacturas Then
   numfacturas = numfacturas + 1
End If
'se empieza a imprimir
Dim Prodreal As Integer
For t = 1 To numfacturas
   Dim RST As ADODB.Recordset
   Set RST = New ADODB.Recordset
   RST.Open "SELECT * FROM cattienda WHERE ticlave = '" & Trim(Mid(cSucursal, 1, 3)) & "'", cn, adOpenForwardOnly, adLockReadOnly, adCmdText
   
   'folfact = IIf(chkCredito.Value = 1, RST!folcredito + 1, RST!folcontado + 1)
   If chkPrev.Value = 1 Or globconFin Then
      folfact = RST!folprevta + 1
      campo = "folprevta"
   ElseIf chkCredito.Value = 1 Then
      folfact = RST!folcredito + 1
      campo = "Folcredito"
   Else
      folfact = RST!folcontado + 1
      campo = "folcontado"
   End If
   numfac = InputBox("PREPARE LA IMPRESORA Y ACOMODE EL PAPEL" & Chr(13) & Chr(13) & "Proporcione numero de factura", "Proporcione factura", folfact)
   RST.Close
   Set RST = Nothing
   If Len(numfac) < 1 Then
        MsgBox "ES NECESARIO ESPECIFICAR NUMERO DE FACTURA ", vbCritical
        Exit Function
   End If
   nPos = InStr(1, numfac, "-")
   If IsNull(numfac) Or Trim(numfac) = "" Or Not IsNumeric(numfac) Then
       If MsgBox("ES NECESARIO ESPECIFICAR NUMERO DE FACTURA" & Chr(13) & "DESEAS CONTINUAR CON LA IMPRESION DE LA FACTURA", vbYesNo + vbQuestion) = vbNo Then
          Printer.KillDoc
          Exit Function
       End If
   Else
       If Not VERIFICA_FACTURA Then
            grabafactura = True
            'nAlto = -1.1
            'nAlto = IIf(Trim(Mid(cSucursal, 1, 3)) <> "28" And Trim(Mid(cSucursal, 1, 3)) <> "16", -1.1, -0.7)   'Miguel cabrera
            nAlto = IIf(Trim(Mid(cSucursal, 1, 3)) <> "28" And Trim(Mid(cSucursal, 1, 3)) <> "16", -0.3, -1.1) 'Miguel cabrera
            nAlto = -1.3
        Else
            MsgBox "Factura Ya existe en la Base de Datos " & Chr(13) & " Los Productos Anteriores Ya fueron marcados con su Factura ", vbInformation
            Exit Function
        End If
   End If
   If lFormAnt Then
      Call encabezadofacoldchica
      Printer.CurrentY = (6.7 + nAlto)
   Else
      Call encabezadofacold    'Imprime encabezado de la factura
      Printer.CurrentY = (6.4 + nAlto)
   End If
   Call INICIALIZAVALS      'Inicializa variables de totales de la factura
   For j = 1 To ProdenFac
        If t > 1 Then
           If t > 2 Then
              Prodreal = (ProdenFac * (t - 1) + j)
           Else
              Prodreal = (ProdenFac + j)
           End If
        Else
           Prodreal = j
        End If
        Me.stb1.Panels(1).Text = "Clave " & consecs(Prodreal) & "   Número " & Prodreal & " de " & i: stb1.Refresh
        'en el caso de que se pase del total de producto que deben imprimirse
        If Prodreal > i Then
            Exit For
        End If
       If lFormAnt Then
          IMPRIMEDETALLEOLDchica (Prodreal)
       Else
          IMPRIMEDETALLEOLD (Prodreal) 'Imprime una linea de la factura (productos)
       End If
       Call SUMAPRODUCTO(Prodreal)  'Suma el importe del producto para obtener totales por depto.
           While Not GRABASERIEYFAC(Prodreal, globconFin)
               If MsgBox("Actualmente existe una transacción en proceso, Desea reintentar la actualización de la factura", vbRetryCancel + vbExclamation) = vbCancel Then
                  MsgBox "No se pudo actualizar el detalle de la factura, cancele la impresion en el panel de control y vuelva a imprimir la factura", vbCritical
                  cn.Execute "UPDATE ventas_det SET facturado = 0, factura = null, serie = null WHERE factura = '" & Factura & "' and serie = '" & Trim(SERIE) & "'"
                  cn.Execute "DELETE FROM facventa WHERE numfactura = '" & Factura & "' and serie = '" & SERIE & "'"
                  cn.Execute "DELETE FROM facventa_det WHERE factura = '" & Factura & "' and serie = '" & SERIE & "'"
                  Printer.KillDoc
                  Exit Function
               End If
           Wend
   Next
   pgrabafactura globconFin 'graba la factura en la tabla facventa
   If lFormAnt Then
      Call IMPRIMESUBOLDchica    'Imprime subtotales por depto de la factura
   Else
      Call IMPRIMESUBOLD
   End If
   cn.Execute "UPDATE cattienda SET " & campo & " = " & numfac & "  WHERE ticlave = '" & Trim(Mid(cSucursal, 1, 3)) & "'"
Next
stb1.Panels(1).Text = "[Esc] para salir": stb1.Refresh
FacturaNva = True
Exit Function
Error:
   MsgBox "OCURRIO EL SIGUIENTE ERROR: " & Err.Description & Chr(13) & "APAGUE LA IMPRESORA Y CANCELE LA IMPRESION DESDE EL PANEL DE CONTROL PARA NO DESPERDICIAR FACTURAS", vbExclamation
End Function

Sub imppvta(pvta As Double)
Dim rsttemp As ADODB.Recordset
Dim RST As ADODB.Recordset
Dim tmp As ADODB.Recordset
Dim nTotal
Dim nCajas
Dim cCad As String
On Error GoTo Error:
lImp = False
For Each x In Printers
   If x.DeviceName Like "*PREVENTA*" Then
      lImp = True
      Set Printer = x
      Exit For
   End If
Next x
If lImp = False Then
   If MsgBox("NO ES POSIBLE IMPRIMIR TICKET'S DE PREVENTA" & Chr(13) & Chr(13) & "PORQUE NO EXISTEN IMPRESORAS INSTALADAS" & Chr(13) & "PARA ESTE TIPO DE TICKET'S, DESEAS ENVIARLO A LA IMPRESORA PREDETERMINADA", vbCritical + vbYesNo) = vbNo Then Exit Sub
End If

nAncho = 250    'En puntos es el ancho del ticket de las Miniprinter

'Printer.PaperSize = vbPRPSUser

Printer.ScaleMode = vbPoints
Printer.Height = 1000000
Printer.Width = 20000
'Printer.PaperSize = vbPRPSUser

Set tmp = New ADODB.Recordset
Set rsttemp = New ADODB.Recordset
Set rs = New ADODB.Recordset
rsttemp.LockType = adLockReadOnly
rsttemp.CursorType = adOpenForwardOnly
rs.Open "SELECT * FROM cattienda WHERE ticlave = '" & Mid(cSucursal, 1, 3) & "'", cn, adOpenForwardOnly, adLockOptimistic, adCmdText
Direc = rs!Direccion: tel = rs!telefonos
rs.Close
If pvta > 0 Then
   rs.Open "SELECT v.agente, v.chofer ,v.noventa, cnombre, cdireccion,ctelefono,ccolonia,cciudad,V.cl_terminal,V.facrfc, ccredito, climitecredito, ctiempocredito, cclave, montototal, nomnegocio FROM ventas v, ventas_det d, catcliente C WHERE c.cclave = v.clcliente and c.ctipo = 0 and v.noventa = d.noventa AND v.noventa = " & pvta & " ORDER BY v.noventa", cn, adOpenKeyset, adLockOptimistic, adCmdText
   NTOTTICKET = 1
   pvta = Val(txtcampos(2).Text)
ElseIf pvta < 0 Then 'Se imprimen los pedidos modificados
   pvta = Val(txtcampos(2).Text)
   rs.Open "SELECT v.agente, v.chofer ,v.noventa, cnombre, cdireccion,ctelefono,ccolonia,cciudad,V.cl_terminal, facrfc, ccredito, climitecredito, ctiempocredito,cclave, montototal, nomnegocio FROM ventas v, ventas_det d, catcliente C WHERE c.cclave = v.clcliente and c.ctipo = 0 and v.noventa = d.noventa AND v.folpreventa = " & pvta & " and situacion = 0 ORDER BY v.noventa", cn, adOpenForwardOnly, adLockReadOnly, adCmdText
   rsttemp.Open "SELECT count(DISTINCT v.noventa) as Tot FROM Ventas v, Ventas_det d WHERE v.noventa = d.noventa AND SITUACION = 0 AND v.FOLPREVENTA = " & pvta, cn, adOpenDynamic, adLockOptimistic, adCmdText
   NTOTTICKET = rsttemp!tot
   rsttemp.Close
Else 'Se imprime toda la preventa
   pvta = Val(txtcampos(2).Text)
   rs.Open "SELECT v.agente, v.chofer ,v.noventa, cnombre, cdireccion,ctelefono,ccolonia,cciudad,V.cl_terminal, facrfc, ccredito, climitecredito, ctiempocredito,cclave, montototal, nomnegocio FROM ventas v, ventas_det d, catcliente C WHERE c.cclave = v.clcliente and c.ctipo = 0 and v.noventa = d.noventa AND v.folpreventa = " & pvta & " ORDER BY v.noventa", cn, adOpenForwardOnly, adLockReadOnly, adCmdText
   rsttemp.Open "SELECT count(DISTINCT v.noventa) as Tot FROM Ventas v, Ventas_det d WHERE v.noventa = d.noventa AND v.FOLPREVENTA = " & pvta, cn, adOpenForwardOnly, adLockReadOnly, adCmdText
   NTOTTICKET = rsttemp!tot
   rsttemp.Close
End If
nVtaAnt = 0: nTicket = 0: nGranTot = 0
stb1.Style = sbrSimple
While Not rs.EOF
   If nVtaAnt <> rs!noventa Then
      rsttemp.Open "SELECT * FROM Ventas v, Ventas_det d, Tfproduc WHERE v.noventa = d.noventa AND cl_producto = TFPRODUC.Consec AND v.NoVenta = '" & rs!noventa & "' ORDER BY Descripc,contenid", cn, adOpenForwardOnly, adLockReadOnly, adCmdText
      If Not (rsttemp.BOF And rsttemp.EOF) Then
          nTicket = nTicket + 1
          stb1.SimpleText = "            Imprimiendo ticket " & nTicket & " correspondiente al cliente " & Trim(rs!cNombre)
          stb1.Refresh
          Printer.FontName = "Arial"
          Printer.FontSize = 8
          If rs!ccredito Then
             tmp.Open "SELECT SUM(Total) debe FROM facventa WHERE facfecha < '" & date - (rs!ctiempocredito + 2) & "' and cobrado = 0 AND cancelado = 0 AND faccliente = " & rs!cclave, cn, adOpenForwardOnly, adLockReadOnly, adCmdText
             Printer.Print "=========================================================================="
             Printer.Print "   ***  CLIENTE AUTORIZADO CON CREDITO  ***"
             Printer.Print "=========================================================================="
             Printer.Print "LIMITE CREDITO:  " & Format(rs!CLIMITECREDITO, "$##,###,###.00")
             Printer.Print "P L A Z O     :    " & rs!ctiempocredito & " DIAS"
             Printer.Print ""
             Printer.Print "VENCIDO   : " & IIf(IsNull(tmp!debe), "$0.00", Format(tmp!debe, "##,###,###.00"))
             Printer.Print "DISPONIBLE  : " & Format(rs!CLIMITECREDITO - IIf(IsNull(tmp!debe), 0, tmp!debe + rs!montototal), "##,###,###.00")
             Printer.Print "--------------------------------------------------------------------------"
             Printer.Print ""
             tmp.Close
          End If
          Printer.FontBold = True
          If rs!facrfc Then
             Printer.Print ""
             Printer.Print "--------------------------------------------------------------------------"
             Printer.Print "    ***  ESTE CLIENTE SOLICITA FACTURA  ***"
             Printer.Print "--------------------------------------------------------------------------"
             Printer.Print ""
          End If
          tmp.Open "SELECT * FROM catcliente WHERE ctipo = 1 AND cclave = " & Trim(txtcampos(9).Text), cn, adOpenForwardOnly, adLockReadOnly, adCmdText
          'If Not tmp!franquicia Then
            If ZONA = "OAX" Then
               Printer.Print "        VIVERES Y LICORES S.A DE C.V."
            Else
               Printer.Print "HOLDING MEXICO CENTRO AMERICA SA DE CV"
            End If
            Printer.Print Direc; " "
            Printer.Print "    TEL:" & tel
          'Else
          '  Printer.Print "EMPACADORA DE GRANOS Y SEM. NAL. SA CV"
          '  Printer.Print "        AV. SIMBOLOS PATRIOS Nº 912 "
          '  Printer.Print "            COL. REFORMA AGRARIA"
          'End If
          Printer.Print ""
          Printer.Print "PREVENTA: ";
          Printer.FontBold = False
          Printer.Print pvta & "  -  " & date & "  -  " & Format(Time, "HH:MM:SS")
          Printer.FontBold = True
          Printer.Print "AGENTE: ";
          Printer.FontBold = False
          AGTE = tmp!cNombre
          Printer.Print AGTE
          tmp.Close
          tmp.Open "SELECT * FROM catcliente WHERE ctipo = 2 AND cclave = " & Trim(txtcampos(3).Text), cn, adOpenForwardOnly, adLockReadOnly, adCmdText
          Printer.FontBold = True
          Printer.Print "CHOFER: ";
          Printer.FontBold = False
          Printer.Print tmp!cNombre
          tmp.Close
          Printer.Print ""
          Printer.FontBold = True
          Printer.Print "COMPUTADORA: ";
          Printer.FontBold = False
          Printer.Print UCase(rs!CL_TERMINAL)
          Printer.FontBold = True
          Printer.Print "FOL.UNICO VTA: ";
          Printer.FontBold = False
          Printer.Print rs!noventa
          Printer.FontBold = True
          Printer.Print "TICKET   : ";
          Printer.FontBold = False
          Printer.Print nTicket & " DE " & NTOTTICKET
          
          Printer.Print ""
          Printer.FontBold = True
          Printer.Print "VENTA A: ";
          Printer.FontBold = False
          If Len(Trim(rs!cNombre)) > 36 Then
             Printer.Print Mid(rs!cNombre, 1, 30)
             Printer.Print Mid(rs!cNombre, 30)
          Else
             Printer.Print rs!cNombre
          End If
          If Trim(rs!nomnegocio) <> "" Then
             Printer.FontBold = True
             Printer.Print "NEGOCIO  : ";
             Printer.FontBold = False
             Printer.Print rs!nomnegocio
          End If
          Printer.FontBold = True
          Printer.Print "DIRECCION: ";
          Printer.FontBold = False
          Printer.Print rs!cdireccion
          If rs!ctelefono <> 0 Then
             Printer.FontBold = True
             Printer.Print "TELEFONO : ";
             Printer.FontBold = False
             Printer.Print rs!ctelefono
          End If
          If Trim(rs!ccolonia) <> "" Then
             Printer.FontBold = True
             Printer.Print "COLONIA : ";
             Printer.FontBold = False
             Printer.Print rs!ccolonia
          End If
          If rs!cciudad <> 0 Then
             Printer.FontBold = True
             Printer.Print "CIUDAD  : ";
             Printer.FontBold = False
             Printer.Print rs!cciudad
          End If
          Printer.Print "--------------------------------------------------------------------------"
          nTotal = 0: nCajas = 0: nPiezas = 0: nProd = 0: nPeso = 0: NPROMO = 0
          Printer.FontBold = False
          Printer.FontName = "Arial Narrow"
          'Cuando esta vacio el detalle del traslado para que se imprima el encabezado
          Do While (Not rsttemp.EOF)
             If (rsttemp!cantidad > 0 Or rsttemp!cantidadp > 0) Then
                If rsttemp!cantidad > 0 And rsttemp!cantidadp > 0 Then
                   cCad = CStr(rsttemp!cantidad) & " CAJ, " & CStr(rsttemp!cantidadp) & " PZA " & Trim(rsttemp!descripc)
                ElseIf rsttemp!cantidad > 0 And rsttemp!cantidadp = 0 Then
                   cCad = CStr(rsttemp!cantidad) & " CAJ " & Trim(rsttemp!descripc)
                ElseIf rsttemp!cantidadp > 0 And rsttemp!cantidad = 0 Then
                   cCad = CStr(rsttemp!cantidadp) & " PZA " & Trim(rsttemp!descripc)
                End If
                 'En caso de que sea muy grande la descripcion se imprime en dos lineas
                 'cprom = IIf(Trim(rsttemp!claprove) = "G52", "**P", "")
                 cprom = ""
                 If Len(Trim(cCad)) > 36 Then
                    Printer.Print Mid(cCad, 1, 30)
                    Printer.Print Mid(cCad, 31) & "  " & cprom
                 Else
                    Printer.Print cCad & "  " & cprom
                 End If
                 Printer.CurrentX = 30
                 Printer.Print CStr(rsttemp!PAQUETES) & " X " & CStr(rsttemp!CONTENID) & " " & rsttemp!medida;
        
                 N = nAncho - 120 - (Printer.TextWidth(Format(rsttemp!PRECIO, "###,###,##0.00") & Space(6) & Format(rsttemp!importe, "$###,###,##0.00")))
                 Printer.CurrentX = N + 20
                 Printer.Print Format(rsttemp!PRECIO, "$###,###,##0.00") & Space(15) & Format(rsttemp!importe, "$###,###,##0.00")
                 'If Trim(rsttemp!claprove) = "C52" Then NPROMO = NPROMO + rsttemp!importe
                 nTotal = nTotal + rsttemp!importe
                 nProd = nProd + 1
                 nCajas = nCajas + rsttemp!cantidad
                 nPiezas = nPiezas + rsttemp!cantidadp
                 If Not IsNull(rsttemp!peso) Then nPeso = nPeso + (rsttemp!peso * rsttemp!cantidad)
             End If
             rsttemp.MoveNext
          Loop
          Printer.FontName = "Arial"
          Printer.Print "--------------------------------------------------------------------------"
          Printer.FontSize = 10
          Printer.FontBold = True
          Printer.Print "IMPORTE TOTAL :     " & Format(nTotal, "$###,###,##0.00")
          Printer.Print "--------------------------------------------------------------------------"
          Printer.Print " "
          Printer.FontSize = 8: Printer.FontBold = False
          'Printer.Print "**P = PROMOCION NESTLE: " & NPROMO
          'Printer.Print "BOLETOS DE PROMOCION: "; Int(NPROMO / 50)
          'Printer.Print " "
          Printer.Print "TOTAL PRODUCTOS: " & nProd
          Printer.Print "TOTAL CAJAS    : " & Format(nCajas, "#,###,##0")
          Printer.Print "TOTAL PIEZAS   : " & Format(nPiezas, "#,###,##0")
          Printer.Print "TOTAL PESO     : " & nPeso
          nGranTot = nGranTot + nTotal
          Printer.EndDoc
      End If
      rsttemp.Close
   End If
   nVtaAnt = rs!noventa
   rs.MoveNext
Wend
If NTOTTICKET > 1 Then  'Solo cuando se imprime toda la preventa
   Printer.FontSize = 10
   Printer.FontBold = True
    Printer.Print "--------------------------------------------------------------------------"
    Printer.Print "FECHA : " & date
    Printer.Print "AGENTE : " & AGTE
    Printer.Print "FOLIO DE PREVENTA : " & pvta
    Printer.Print "TOTAL PREVENTA    : " & Format(nGranTot, "$###,###,##0.00")
    Printer.Print "TOTAL DE PEDIDOS  : " & NTOTTICKET
    Printer.Print "--------------------------------------------------------------------------"
    Printer.EndDoc
'Else
'   MsgBox "NO EXISTEN TICKET'S PARA IMPRIMIR", vbInformation, "Ticket's"
End If
stb1.Style = sbrNormal
Exit Sub
Error:
  MsgBox Err.Description
End Sub
  
Private Function ObtenSerie(credito As Boolean, Preventa As Boolean, Confin As Boolean) As String
ObtenSerie = ""
If Confin Or Preventa Then
   If Trim(Mid(cSucursal, 1, 3)) = "16" Then        'Miguel Cabrera
      ObtenSerie = "HHH"
   ElseIf Trim(Mid(cSucursal, 1, 3)) = "55" Then    'Puerto Escondido
      ObtenSerie = "DDD"
   ElseIf Trim(Mid(cSucursal, 1, 3)) = "26" Then    'Miahuatlan
      ObtenSerie = "HHH"
   ElseIf Trim(Mid(cSucursal, 1, 3)) = "28" Then    'Istmo
      ObtenSerie = "JJJ"
   ElseIf Trim(Mid(cSucursal, 1, 3)) = "13" Then    'Tapachula
      ObtenSerie = "ABX"
   End If
   
ElseIf credito Then
   If Trim(Mid(cSucursal, 1, 3)) = "16" Then        'Miguel Cabrera
      ObtenSerie = "HHH"
   ElseIf Trim(Mid(cSucursal, 1, 3)) = "24" Then    'Cosijopi Central
      ObtenSerie = "J2"
   ElseIf Trim(Mid(cSucursal, 1, 3)) = "28" Then    'Istmo
      ObtenSerie = "JJJ"
   ElseIf Trim(Mid(cSucursal, 1, 3)) = "55" Then    '
      ObtenSerie = "H2"                             'Puerto Escondido
   ElseIf Trim(Mid(cSucursal, 1, 3)) = "13" Then    'Tapachula
      ObtenSerie = "AB"
   End If
Else   ' CONTADO
   If Trim(Mid(cSucursal, 1, 3)) = "16" Then        'Miguel cabrera
      ObtenSerie = "Y1"
   ElseIf Trim(Mid(cSucursal, 1, 3)) = "24" Then    'Cosijopi Central
      ObtenSerie = "I2"
   ElseIf Trim(Mid(cSucursal, 1, 3)) = "23" Then    'Central de Abastos
      ObtenSerie = "D2"
   ElseIf Trim(Mid(cSucursal, 1, 3)) = "26" Then    'Miahuatlan
      ObtenSerie = "GGG"
   ElseIf Trim(Mid(cSucursal, 1, 3)) = "28" Then    'Istmo
      ObtenSerie = "JJJ"
   ElseIf Trim(Mid(cSucursal, 1, 3)) = "55" Then    'Puerto Escondido
      'ObtenSerie = "G2"
      ObtenSerie = "DDD"
   ElseIf Trim(Mid(cSucursal, 1, 3)) = "13" Then    'Tapachula
      ObtenSerie = "AB"
   End If
End If
If ObtenSerie = "" Then
   MsgBox "NO EXISTE SERIE ASOCIADA A ESTA SUCURSAL, DEBERA CANCELAR LA IMPRESION DE LA FACTURA E INFORMAR AL ADMINISTRADOR DEL SISTEMA", vbCritical
End If
End Function

Private Sub ImpPvtaH()
Dim nSurCaj
Dim nSurPza
    On Error GoTo Error:
    cmdg.DialogTitle = "Importar pedido de Agentes de venta"
    cmdg.Filter = "Pedidos de Agentes ( *.mdb ) | *.mdb"
    cmdg.CancelError = True
    cmdg.ShowOpen
    cRutArc = cmdg.FileName
    If cRutArc = "" Or IsNull(cRutArc) Then
       MsgBox "DEBE ESPECIFICAR UNA RUTA Y NOMBRE DE ARCHIVO", vbExclamation
       Exit Sub
    End If
    For N = 1 To Len(cRutArc)
       If Mid(cRutArc, N, 1) = "\" Then nPos = N
    Next
    cruta = Mid(cRutArc, 1, nPos)
    cArch = Mid(cRutArc, nPos + 1)
    cArch = Mid(cArch, 1, Len(cArch) - 4) 'Le quito la extension porque si no marca error en la consula SQL.
    If cArch <> Trim(txtcampos(9).Text) Then
       MsgBox "EL ARCHIVO SELECCIONADO CORRESPONDE A OTRO AGENTE", vbCritical, "Archivo incorrecto"
       Exit Sub
    End If
    Dim RSTEMP As ADODB.Recordset
    Set RSTEMP = New ADODB.Recordset
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    AdoAgentes.CommandType = adCmdText
    AdoAgentes.ConnectionString = "DSN=PITICOMDB;DBQ=" & cruta & cArch & ".mdb;DefaultDir=" & cruta & ";DriverId=25;FIL=MS Access;MaxBufferSize=2048;PageTimeout=5;UID=admin;"
    AdoAgentes.RecordSource = "SELECT * FROM pedidos ORDER BY folio"
    AdoAgentes.Refresh
    If AdoAgentes.Recordset.RecordCount > 0 Then
       MsgBox "A CONTINUACION SE PROCESARAN " & AdoAgentes.Recordset.RecordCount & " PEDIDOS ", vbInformation, "Importa pedidos de agentes"
    Else
       MsgBox "EN EL PEDIDO DE MAYOREO NO EXISTEN PRODUCTOS PARA DAR DE ALTA", vbInformation, "Franquicias"
       Exit Sub
    End If
    'Se agrega el detalle de la venta
    rs.Open "SELECT MAX(FOLPREVENTA) AS FOLPRE FROM VENTAS", cn, adOpenKeyset, adLockOptimistic, adCmdText
    Prev = IIf(IsNull(rs!FOLPRE), 1, rs!FOLPRE + 1)
    rs.Close
    
    cmen = stb1.Panels(1).Text
    lTrans = True
    'cn.BeginTrans
    folant = 0
    While Not AdoAgentes.Recordset.EOF
        stb1.Panels(1).Text = "Pedido: " & Trim(AdoAgentes.Recordset!Folio) & "        Cliente: " & AdoAgentes.Recordset!CLIENTE
        stb1.Refresh
        AdoAgtedet.CommandType = adCmdText
        AdoAgtedet.ConnectionString = "DSN=PITICOMDB;DBQ=" & cruta & cArch & ".mdb;DefaultDir=" & cruta & ";DriverId=25;FIL=MS Access;MaxBufferSize=2048;PageTimeout=5;UID=admin;"
        AdoAgtedet.RecordSource = "SELECT * FROM pedidosdet WHERE ( (ISNULL(cajasv) And ISNULL(piezasv)) OR (cajasv = 0 And piezasv = 0 ) ) and folio = " & AdoAgentes.Recordset!Folio
        AdoAgtedet.Refresh
        If Not (AdoAgtedet.Recordset.BOF And AdoAgtedet.Recordset.EOF) Then
           agente = txtcampos(9).Text
           chofer = txtcampos(3).Text
           Preventa = Prev
           nDispo = 0
           'Si es un nuevo cliente primero se da de alta (BUSCA EN ACCESS)
           rs.Open "SELECT * FROM clientes WHERE clave =" & AdoAgentes.Recordset!CLIENTE, AdoAgentes.ConnectionString, adOpenKeyset, adLockOptimistic, adCmdText
           adopreventa.CommandType = adCmdText
           adopreventa.ConnectionString = cCadConex
           adopreventa.RecordSource = "SELECT * FROM catcliente WHERE cclave = " & AdoAgentes.Recordset!CLIENTE
           adopreventa.Refresh
           If Not (rs.BOF And rs.EOF) Then
              If rs!NUEVO Then adopreventa.Recordset.AddNew
              adopreventa.Recordset!cNombre = UCase(rs!NOMBRE)
              adopreventa.Recordset!cnombrefac = UCase(rs!NOMBRE)
              adopreventa.Recordset!cdireccion = IIf(IsNull(rs!calle), ".", UCase(rs!calle))
              adopreventa.Recordset!crfc = IIf(IsNull(rs!rfc), "", rs!rfc)
              adopreventa.Recordset!ctelefono = IIf(IsNull(rs!telefono), ".", rs!telefono)
              adopreventa.Recordset!ccolonia = IIf(IsNull(rs!colonia), ".", UCase(rs!colonia))
              adopreventa.Recordset!cciudad = IIf(IsNull(rs!poblacion), ".", UCase(rs!poblacion))
              adopreventa.Recordset!apepaterno = IIf(IsNull(rs!apepaterno), " ", UCase(rs!apepaterno))
              adopreventa.Recordset!apematerno = IIf(IsNull(rs!apematerno), " ", UCase(rs!apematerno))
              adopreventa.Recordset!nombres = IIf(IsNull(rs!nombres), " ", UCase(rs!nombres))
              adopreventa.Recordset!nomnegocio = IIf(IsNull(rs!nomnegocio), " ", UCase(rs!nomnegocio))
              adopreventa.Recordset!ruta = rs!ruta
              adopreventa.Recordset.Update
              CLIENTE = adopreventa.Recordset!cclave
           Else
              CLIENTE = AdoAgentes.Recordset!CLIENTE
              If adopreventa.Recordset!ccredito = True Then
                 plazo = adopreventa.Recordset!ctiempocredito
                 limcred = adopreventa.Recordset!CLIMITECREDITO
                 rs.Close
                 'rs.Open "SELECT SUM(TOTAL) as totdebe FROM facventa WHERE Cobrado = 0 AND cancelado = 0 AND Faccliente = '" & CLIENTE & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
                 rs.Open "SELECT SUM(porpagar) as totdebe FROM facventa,catcliente WHERE facfecha <= '" & date - (ctiempocredito + 2) & "' and cobrado = 0 AND cancelado = 0 AND Faccliente = cclave AND FACCLIENTE = " & CLIENTE, cn, adOpenKeyset, adLockOptimistic, adCmdText
                 If rs!totdebe >= limcred Then
                    nDispo = 1
                 ElseIf rs!totdebe > 0 Then
                   nDispo = limcred - rs!totdebe
                 End If
              End If
           End If
           rs.Close
           AdoVentas.Recordset.AddNew
           AdoVentas.Recordset!fecha = date & " " & Time
           AdoVentas.Recordset!tipoventa = 2
           AdoVentas.Recordset!CL_TERMINAL = Caja
           AdoVentas.Recordset!CL_operador = Trim(cUsuario)
           AdoVentas.Recordset!clcliente = CLIENTE
           AdoVentas.Recordset!tienda = Mid(cSucursal, 1, 3)
           AdoVentas.Recordset!chofer = chofer
           AdoVentas.Recordset!agente = agente
           If AdoAgentes.Recordset!credito Then
              AdoVentas.Recordset!credito = 1
           Else
              AdoVentas.Recordset!Prevta = 1
           End If
           AdoVentas.Recordset!facrfc = IIf(AdoAgentes.Recordset!Factura, 1, 0)
           AdoVentas.Recordset!situacion = "0"    'Venta Hecha por Mayoreo en tramite
           AdoVentas.Recordset!FOLPREVENTA = Preventa
           AdoVentas.Recordset.Update
        End If
        nventa = 0
        While Not AdoAgtedet.Recordset.EOF
          If AdoAgtedet.Recordset!cajas > 0 Or AdoAgtedet.Recordset!piezas > 0 Then
             rs.Open "SELECT * FROM inventario, preprod,tfproduc WHERE consec = inprod and consec = preclave and inprod = preclave And inprod = '" & Trim(AdoAgtedet.Recordset!clave) & "'", cn, adOpenStatic, adLockOptimistic, adCmdText
             If rs!InCant > 0 Then
                nSurPza = 0
                nSurCaj = 0
                'Cuando piden cajas
                If rs!InCant >= AdoAgtedet.Recordset!cajas Then
                   nSurCaj = AdoAgtedet.Recordset!cajas
                ElseIf rs!InCant < AdoAgtedet.Recordset!cajas Then
                   nSurCaj = rs!InCant
                End If
                'Cuando piden Piezas y hay en inventario
                If AdoAgtedet.Recordset!piezas > 0 And rs!InCantPza > 0 Then
                   If rs!InCantPza >= AdoAgtedet.Recordset!piezas Then
                      nSurPza = AdoAgtedet.Recordset!piezas
                   ElseIf rs!InCantPza < AdoAgtedet.Recordset!piezas Then
                      nSurPza = rs!InCantPza
                   End If
                End If
                'PRECIO = "Precio" & Trim(Str(AdoAgtedet.Recordset!ESCALA))
                PRECIO = AdoAgtedet.Recordset!PRECIO
                Prebajo = IIf(AdoAgtedet.Recordset!ESCALA <> 2, 1, 0)
                PREPZA = IIf(rs!precio4 / rs!PAQUETES > AdoAgtedet.Recordset!preciop, rs!precio4 / rs!PAQUETES, AdoAgtedet.Recordset!preciop)
                AdoAgtedet.Recordset!cajasv = nSurCaj
                AdoAgtedet.Recordset!piezasv = nSurPza
                AdoAgtedet.Recordset.Update
                If nSurCaj > 0 Or nSurPza >= 0 Then
                   'Validar si el limite de credito nDispo
                   'NVENTA = NVENTA + (nSurCaj * RS.Fields(PRECIO).Value) + (nSurPza * PREPZA)
                   nventa = nventa + (nSurCaj * PRECIO) + (nSurPza * PREPZA)
                   If nDispo <> 0 Then
                      If nDispo < nventa Then
                         AdoAgtedet.Recordset!cajasv = 0
                         AdoAgtedet.Recordset!piezasv = 0
                         AdoAgtedet.Recordset.Update
                         AdoAgtedet.Recordset.MoveLast
                         nventa = nventa - ((nSurCaj * PRECIO) + (nSurPza * PREPZA))
                         'AdoAgtedet.Recordset.MoveNext
                      Else
                         If Not (nSurCaj = 0 And nSurPza = 0) Then
                            cn.Execute "INSERT INTO ventas_det(noventa,cl_producto,cantidad,cantidadp,precio,preciop,prebajo,importe,precosto,precostop,iva,ieps,tasaieps) VALUES (" & AdoVentas.Recordset!noventa & ",'" & AdoAgtedet.Recordset!clave & "'," & nSurCaj & "," & nSurPza & "," & PRECIO & "," & PREPZA & "," & Prebajo & "," & (nSurCaj * PRECIO) + (nSurPza * PREPZA) & "," & rs!PRECOSTO & "," & rs!PRECOSTO / rs!PAQUETES & "," & rs!iva & "," & rs!ieps & "," & rs!tasaieps & ")"
                            cn.Execute "UPDATE inventario SET incant = incant - " & nSurCaj & ", incantpza = incantpza - " & nSurPza & " WHERE inprod = '" & AdoAgtedet.Recordset!clave & "'"
                         End If
                      End If
                   Else
                      If Not (nSurCaj = 0 And nSurPza = 0) Then
                         cn.Execute "INSERT INTO ventas_det(noventa,cl_producto,cantidad,cantidadp,precio,preciop,prebajo,importe,precosto,precostop,iva,ieps,tasaieps) VALUES (" & AdoVentas.Recordset!noventa & ",'" & AdoAgtedet.Recordset!clave & "'," & nSurCaj & "," & nSurPza & "," & PRECIO & "," & PREPZA & "," & Prebajo & "," & (nSurCaj * PRECIO) + (nSurPza * PREPZA) & "," & rs!PRECOSTO & "," & rs!PRECOSTO / rs!PAQUETES & "," & rs!iva & "," & rs!ieps & "," & rs!tasaieps & ")"
                         cn.Execute "UPDATE inventario SET incant = incant - " & nSurCaj & ", incantpza = incantpza - " & nSurPza & " WHERE inprod = '" & AdoAgtedet.Recordset!clave & "'"
                      End If
                   End If
                End If
             Else
                AdoAgtedet.Recordset!cajasv = 0
                AdoAgtedet.Recordset!piezasv = 0
                AdoAgtedet.Recordset.Update
             End If
             rs.Close
          End If
          AdoAgtedet.Recordset.MoveNext
        Wend
        cn.Execute "UPDATE ventas SET montototal = " & nventa & " WHERE noventa = " & AdoVentas.Recordset!noventa
        'Hasta que se calcule el total del pedido
        'While Not Actualped(AdoVentas.Recordset!noventa)
        '   MsgBox "Presione Enter para continuar importando la preventa", vbInformation, "Ventas"
        'Wend
        AdoAgentes.Recordset.MoveNext
    Wend
    MsgBox "Se genero la preventa número " & Prev & Chr(13) & Chr(13) & "A continuacion se generará reporte de faltantes", vbInformation, "Folio de preventa generado"
    
    AdoAgentes.ConnectionString = "DSN=PITICOMDB;DBQ=" & cruta & cArch & ".mdb;DefaultDir=" & cruta & ";DriverId=25;FIL=MS Access;MaxBufferSize=2048;PageTimeout=5;UID=admin;"
    AdoAgentes.RecordSource = "SELECT * FROM pedidosdet D, pedidos P WHERE ((d.piezas > d.piezasv OR d.cajas > d.cajasv) OR ( isnull(d.cajasv) OR isnull(d.piezasv)) ) AND p.Folio = d.Folio ORDER BY p.folio"
    AdoAgentes.Refresh
    'MsgBox AdoAgentes.Recordset.RecordCount
    Open App.Path & "\F" & Prev & ".TXT" For Output As #1 ' Abre el archivo para operaciones de salida.
    If ZONA = "OAX" Then
       Print #1, Tab(25); "VIVERES Y LICORES S.A DE C.V."
    Else
       Print #1, Tab(25); "HOLDING MEXICO CENTRO AMERICA SA DE CV"
    End If
    Print #1, Tab(20); "REPORTE DE FALTANTES EN LA PREVENTA "; Prev
    Print #1, Tab(15); "AGENTE: " & cmbAgente.Text & "        Fecha: " & date
    Print #1, String(107, "-")
    Print #1, "    PRODUCTO                                            PRESENTACION   CAJ   CAJV  $CAJ    PZA  PZAV $PZA."
    Print #1, String(107, "-")
    cteAnt = ""
    While Not AdoAgentes.Recordset.EOF
       If cteAnt <> AdoAgentes.Recordset!CLIENTE Then
          rs.Open "SELECT * FROM catcliente WHERE cclave =  " & Trim(AdoAgentes.Recordset!CLIENTE), cn, adOpenStatic, adLockOptimistic, adCmdText
          Print #1, ""
          If rs.BOF And rs.EOF Then
             Print #1, "** CLIENTE NUEVO PEDIDO: " & AdoAgentes.Recordset!Folio
          Else
             Print #1, "** Cliente => " & rs!cNombre
          End If
          rs.Close
       End If
       rs.Open "SELECT descripc, LTRIM(STR(paquetes)) + ' x ' + LTRIM(STR(CONTENID,10,3)) + space(1) + RTRIM(medida) AS MEDIDA FROM tfproduc WHERE consec = " & AdoAgentes.Recordset!clave, cn, adOpenForwardOnly, adLockOptimistic, adCmdText
       Print #1, Tab(4); rs!descripc; String(50 - Len(rs!descripc), " "); rs!medida; String(20 - Len(rs!medida), " ") & Space(2) & Format(AdoAgentes.Recordset!cajas, "#####0") & Space(3) & Format(AdoAgentes.Recordset!cajasv, "#####0") & Space(3) & Format(AdoAgentes.Recordset!PRECIO, "#####0.00") & Space(5) & Format(AdoAgentes.Recordset!piezas, "#####0") & Space(3) & Format(AdoAgentes.Recordset!piezasv, "###,##0") & Space(3) & Format(AdoAgentes.Recordset!preciop, "###,##0.00")
       rs.Close
       cteAnt = AdoAgentes.Recordset!CLIENTE
       AdoAgentes.Recordset.MoveNext
    Wend
    Close #1
    'MsgBox "wordpad " & App.Path & "\F" & Prev & ".TXT"
    Handle = Shell("NOTEPAD " & App.Path & "\F" & Prev & ".TXT", 1)
    cmdRegresar_Click
    stb1.Panels(1).Text = cmen
    stb1.Refresh
    Unload Me
    Exit Sub
Error:
    'If lTrans Then cn.RollbackTrans
    MsgBox Err.Description
End Sub

'Function Actualped(venta As Double) As Boolean
'On Error GoTo Error:
'  Actualped = False
    'cn.Execute "UPDATE ventas_det SET ventas_det.precosto = t.precosto, ventas_det.precostop = t.precosto / t.paquetes, importe = (precio * cantidad) + (preciop * cantidadp)," & _
               "ventas_det.ieps = t.ieps, ventas_det.iva= t.iva,ventas_det.tasaieps = t.tasaieps FROM tfproduc T WHERE consec = cl_producto AND noventa = " & venta
    'cn.Execute "UPDATE ventas_det SET ventas_det.iva = c.iva, ventas_det.ieps = c.ieps FROM cargos C WHERE caprod = cl_producto AND noventa = " & venta
    'cn.Execute "UPDATE VENTAS set montototal = ( SELECT SUM(importe) FROM ventas_det WHERE noventa = " & venta & ") WHERE noventa = " & venta
'  Actualped = True
'  Exit Function
'Error:
'End Function
Private Sub ImpPvtaL()
    cmdg.DialogTitle = "Importar pedido de Agentes de venta"
    cmdg.Filter = "Pedidos de Agentes ( Prevta*.mdb ) | Prevta*.mdb"
    cmdg.CancelError = True
    cmdg.ShowOpen
    cRutArc = cmdg.FileName
    If cRutArc = "" Or IsNull(cRutArc) Then
       MsgBox "DEBE ESPECIFICAR UNA RUTA Y NOMBRE DE ARCHIVO", vbExclamation
       Exit Sub
    End If
    For N = 1 To Len(cRutArc)
       If Mid(cRutArc, N, 1) = "\" Then nPos = N
    Next
    cruta = Mid(cRutArc, 1, nPos)
    cArch = Mid(cRutArc, nPos + 1)
    cArch = Mid(cArch, 1, Len(cArch) - 4) 'Le quito la extension porque si no marca error en la consula SQL.

    AdoAgentes.CommandType = adCmdText
    AdoAgentes.ConnectionString = "DSN=PITICOMDB;DBQ=" & cruta & cArch & ".mdb;DefaultDir=" & cruta & ";DriverId=25;FIL=MS Access;MaxBufferSize=2048;PageTimeout=5;UID=admin;"
    AdoAgentes.RecordSource = "SELECT * FROM ventas"
    AdoAgentes.Refresh
    If AdoAgentes.Recordset.RecordCount > 0 Then
       MsgBox "A CONTINUACION SE PROCESARAN " & AdoAgentes.Recordset.RecordCount & " PEDIDOS ", vbInformation, "Importa pedidos de agentes"
    Else
       MsgBox "EN EL PEDIDO DE MAYOREO NO EXISTEN PRODUCTOS PARA DAR DE ALTA", vbInformation, "Franquicias"
       Exit Sub
    End If
    'Se agrega el detalle de la venta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    cmen = stb1.Panels(1).Text
    
    lTrans = True
    cn.BeginTrans
    folant = 0
    While Not AdoAgentes.Recordset.EOF
        stb1.Panels(1).Text = "Procesando: " & Trim(AdoAgentes.Recordset!Folio) & " " & AdoAgentes.Recordset!CLIENTE
        stb1.Refresh
        AdoAgtedet.CommandType = adCmdText
        AdoAgtedet.ConnectionString = "DSN=PITICOMDB;DBQ=" & cruta & cArch & ".mdb;DefaultDir=" & cruta & ";DriverId=25;FIL=MS Access;MaxBufferSize=2048;PageTimeout=5;UID=admin;"
        AdoAgtedet.RecordSource = "SELECT * FROM pedidosdet WHERE folio = " & AdoAgentes.Recordset!Folio
        AdoAgtedet.Refresh
        If Not (AdoAgtedet.Recordset.BOF And AdoAgtedet.Recordset.EOF) Then
           agente = AdoVentas.Recordset!agente
           chofer = AdoVentas.Recordset!chofer
           Preventa = AdoVentas.Recordset!FOLPREVENTA
           AdoVentas.Recordset.AddNew
           AdoVentas.Recordset!fecha = date & " " & Time
           AdoVentas.Recordset!tipoventa = 2
           AdoVentas.Recordset!CL_TERMINAL = Caja
           AdoVentas.Recordset!CL_operador = Trim(cUsuario)
           AdoVentas.Recordset!clcliente = AdoAgentes.Recordset!CLIENTE
           AdoVentas.Recordset!tienda = Mid(cSucursal, 1, 3)
           AdoVentas.Recordset!chofer = chofer
           AdoVentas.Recordset!agente = agente
           AdoVentas.Recordset!credito = 1
           AdoVentas.Recordset!facrfc = IIf(AdoAgentes.Recordset!Factura, 1, 0)
           AdoVentas.Recordset!situacion = "0"    'Venta Hecha por Mayoreo en tramite
           AdoVentas.Recordset!FOLPREVENTA = Preventa
           AdoVentas.Recordset.Update
        End If
        While Not AdoAgtedet.Recordset.EOF
            rs.Open "SELECT * FROM inventario WHERE inprod = " & AdoAgtedet.Recordset!clave, cn, adOpenForwardOnly, adLockOptimistic, adCmdText
            'Cuando piden solo cajas
             If AdoAgtedet.Recordset!cajas > 0 And AdoAgtedet.Recordset!piezas = 0 Then
                If rs!InCant >= AdoAgtedet.Recordset!cajas Then
                    nSurtido = AdoAgtedet.Recordset!cajas
                ElseIf rs!InCant <= AdoAgtedet.Recordset!cajas And rs!InCant > 0 Then
                    nSurtido = rs!InCant
                End If
                AdoAgtedet.Recordset!VENTAC = nSurtido
                AdoAgtedet.Recordset.Update
                cn.Execute "INSERT INTO ventas_det(noventa,cl_producto,cantidad) VALUES (" & AdoVentas.Recordset!noventa & ",'" & AdoAgtedet.Recordset!clave & "'," & nSurtido & ")"
                cn.Execute "UPDATE inventario SET incant = incant - " & nSurtido & " WHERE inprod = '" & AdoAgtedet.Recordset!clave & "'"
            End If
            AdoAgtedet.Recordset.MoveNext
            rs.Close
        Wend
        AdoAgentes.Recordset.MoveNext
    Wend
    cn.Execute "UPDATE ventas_det SET precio = precio4, preciop = precio1, ventas_det.precosto = t.precosto, ventas_det.precostop = t.precosto / t.paquetes, importe = (precio4 * cantidad) + (precio1 * cantidadp)," & _
                      "ventas_det.ieps =t.ieps,ventas_det.iva= t.iva,ventas_det.tasaieps = t.tasaieps FROM tfproduc T, preprod P WHERE preclave = consec AND consec = cl_producto AND preclave = cl_producto AND noventa = " & AdoVentas.Recordset!noventa
    cn.Execute "UPDATE ventas_det SET ventas_det.iva = c.iva, ventas_det.ieps = c.ieps FROM cargos C WHERE caprod = cl_producto AND noventa = " & AdoVentas.Recordset!noventa
    cn.CommitTrans: lTrans = False
    stb1.Panels(1).Text = "Generando reporte de faltantes...."
    stb1.Refresh
    Open App.Path & "\FAL" & Mid(cArch, 5) & ".TXT" For Output As #1 ' Abre el archivo para operaciones de salida.
    Print #1, Tab(20); "VIVERES Y LICORES S.A DE C.V."
    Print #1, Tab(10); "CARBONERA 1016 COL. TRINIDAD DE LAS HUERTAS"
    Print #1, "FALTANTES DEL PEDIDO DE FRANQUICIA "; cArch; "  "; UCase(Format(date, "long date"))
    Print #1,   ' Imprime una línea en blanco en el archivo.
    Print #1, "========================================================================================="
    Print #1, "    DESCRIPCION                      MEDIDA                SOLC SURC DIFC  SOLC SURC DIFC"
    Print #1, "========================================================================================="
    Dim Valor
    Valor = Space(5)
    AdoAgtedet.RecordSource = "SELECT * FROM pedidosdet,producto WHERE producto.clave = producto.clave"
    AdoAgtedet.Recordset.MoveFirst
    While Not AdoAgtedet.Recordset.EOF
        If AdoAgtedet.Recordset!cajas > AdoAgtedet.Recordset!VENTAC Or AdoAgtedet.Recordset!piezas > AdoAgtedet.Recordset!ventaP Then
           RSet Valor = AdoAgtedet.Recordset!cajas
           Print #1, Mid(AdoAgtedet.Recordset!descripc, 1, 35) & "   " & AdoAgentes.Recordset!medida & Valor;
           RSet Valor = AdoAgtedet.Recordset!VENTAC
           Print #1, Valor;
           RSet Valor = AdoAgtedet.Recordset!piezas - AdoAgtedet.Recordset!VENTAC
           Print #1, Valor;
           
           RSet Valor = AdoAgtedet.Recordset!piezas
           Print #1, " " & Valor;
           RSet Valor = AdoAgtedet.Recordset!ventaP
           Print #1, Valor;
           RSet Valor = AdoAgtedet.Recordset!piezas - AdoAgtedet.Recordset!ventaP
           Print #1, Valor
        End If
        AdoAgtedet.Recordset.MoveNext
    Wend
    Close #1   ' Cierra el archivo de reporte
    Handle = Shell("NOTEPAD " & App.Path & "\FAL" & Mid(cArch, 5) & ".TXT", 1)
    stb1.Panels(1).Text = cmen
    stb1.Refresh
    Exit Sub
Error:
    If lTrans Then cn.RollbackTrans
    MsgBox Err.Description
End Sub

Private Function SELECCIONA_IMPRESORA(cimpresora As String) As Boolean
SELECCIONA_IMPRESORA = False
lImp = False
For Each x In Printers
   If UCase(x.DeviceName) Like cimpresora Then
      lImp = True
      Set Printer = x
      SELECCIONA_IMPRESORA = True
      Exit For
   End If
Next x
End Function

Private Sub compara()
Dim RSMDB As ADODB.Recordset
Dim CNMDB As ADODB.Connection
Dim rs As ADODB.Recordset
Dim Valor As String
Dim lpro As Boolean
'AdoVentas.Recordset!clcliente = 1630 Miahuatlan
If Not (AdoVentas.Recordset!clcliente = 4179 Or AdoVentas.Recordset!clcliente = 12412 Or AdoVentas.Recordset!clcliente = 19173) Then
   Exit Sub
Else
   If MsgBox("La venta corresponde a una franquicia deseas ejecutar el proceso de prorrateo" & Chr(13) & "El proceso tardará dependiendo de la cantidad de productos", vbInformation + vbYesNo, "Prorrateo") = vbNo Then Exit Sub
End If
Set rs = New ADODB.Recordset
Set RSMDB = New ADODB.Recordset
Set CNMDB = New ADODB.Connection
CNMDB.Open "DSN=PITICOMDB;DBQ=P:\PITICO\PITICO10\PITICO10.mdb;DefaultDir=P:\PITICO\PITICO10;DriverId=25;FIL=MS Access;MaxBufferSize=2048;PageTimeout=5;UID=admin;PWD=PORTATIL"
rs.Open "SELECT igualofi,CONSEC,p.precio4 as precio,v.preciop,v.cantidad, v.cantidadp,descripc, LTRIM(STR(T.paquetes)) + ' x ' + LTRIM(STR(CONTENID,10,3)) + space(1) + RTRIM(t.medida) AS MEDIDA, t.paquetes as paquetes FROM preprod p, VENTAS E,ventas_det v, tfproduc T WHERE p.preclave = v.cl_producto and p.preclave = t.consec and E.noventa = v.noventa AND T.consec = cl_producto AND v.cancelado = 0 and E.noventa = " & Me.AdoVentas.Recordset!noventa & " AND E.clcliente = " & AdoVentas.Recordset!clcliente & " ORDER BY descripc, medida", cn, adOpenStatic, adLockOptimistic, adCmdText

Open App.Path & "\" & CStr(AdoVentas.Recordset!noventa) & "s.TXT" For Output As #1 ' Abre el archivo para operaciones de salida.
Print #1, Tab(20); "VIVERES Y LICORES S.A DE C.V."
Print #1, Tab(10); "CARBONERA 1016 COL. TRINIDAD DE LAS HUERTAS"
Print #1, Tab(10); "COMPARATIVO DE VENTAS ENTRE MIGUEL CAB. Y OFICINAS CENTRALES"
Print #1, "PEDIDO DEL CLIENTE " & Me.cmbCliente.Text & " EN EL FOLIO DE PREVENTA " & AdoVentas.Recordset!FOLPREVENTA
Print #1, "===================================================================================================="
Print #1, "                      D E S C R I P C I O N                        CJ.    CJ.OFI   CJ.MCAB     DIF.  "
Print #1, "===================================================================================================="
If AdoVentas.Recordset!clcliente = 1630 Then   'Miahutlan
   PRECIO = "PRECIO3"
Else
   PRECIO = "PRECIO4"
End If
nMasAltomc = 0: nMasAltoO = 0: nTotCjm = 0: NTOTCJO = 0: nVarMc = 0: nVarOfi = 0
nTotpro = 0
cmen = stb1.Panels(1).Text
cn.Execute "DELETE FROM franqtmp"
While Not rs.EOF
    RSMDB.Open "SELECT precio1,precio3,precio4,descripc,LTRIM(STR(paquetes)) + ' x ' + LTRIM(STR(CONTENID)) + space(1) + RTRIM(medida) AS MEDIDA1 FROM preprod,tfproduc WHERE consec = preclave and preclave = '" & rs!CONSEC & "'", CNMDB, adOpenForwardOnly, adLockOptimistic, adCmdText
    stb1.Panels(1).Text = rs!descripc
    stb1.Refresh
    If Not (RSMDB.BOF And RSMDB.EOF) And rs!igualofi Then
       nTotpro = nTotpro + 1
       premc = Round(rs!PRECIO, 2)
       'El precio de Mc es mas alto que carbonera
       If premc > RSMDB.Fields(PRECIO).Value Then
            Print #1, "*";
            nMasAltomc = nMasAltomc + (premc - RSMDB.Fields(PRECIO).Value) * rs!cantidad
            nVarMc = nVarMc + 1
       Else
          Print #1, " ";
          nMasAltoO = nMasAltoO + (RSMDB.Fields(PRECIO).Value - premc) * rs!cantidad
          nVarOfi = nVarOfi + 1
       End If
       Valor = Space(50)
       LSet Valor = Mid(Trim(rs!descripc), 1, 43)
       Print #1, Valor;
       Valor = Space(15)
       RSet Valor = rs!medida
       Print #1, Valor;
       Valor = Space(5)
       RSet Valor = rs!cantidad
       Print #1, Valor;

       Valor = Space(10)
       RSet Valor = Format(RSMDB.Fields(PRECIO).Value * rs!cantidad, "#######0.00")
       Print #1, Valor;
       RSet Valor = Format(premc * rs!cantidad, "#######0.00")
       Print #1, Valor;
       RSet Valor = Format((premc * rs!cantidad) - (RSMDB.Fields(PRECIO).Value * rs!cantidad), "#######0.00")
       Print #1, Valor
       
       NTOTCJO = NTOTCJO + RSMDB.Fields(PRECIO).Value * rs!cantidad
       nTotCjm = nTotCjm + premc * rs!cantidad
       VALDIF = (premc * rs!cantidad) - Round(RSMDB.Fields(PRECIO).Value * rs!cantidad, 2)
       cn.Execute "INSERT INTO franqtmp VALUES ('" & rs!CONSEC & "','" & rs!descripc & "','" & rs!medida & "'," & rs!cantidad & "," & Round(RSMDB.Fields(PRECIO).Value, 2) & "," & premc & "," & VALDIF & ",0,0)"
    Else
       Valor = Space(50)
       LSet Valor = Mid(Trim(rs!descripc), 1, 43)
       Print #1, Valor;
       Valor = Space(16)
       RSet Valor = rs!medida
       Print #1, Valor;
       Valor = Space(5)
       RSet Valor = rs!cantidad
       Print #1, Valor
    End If
    RSMDB.Close
    rs.MoveNext
Wend
stb1.Panels(1).Text = cmen
stb1.Refresh
Print #1, "===================================================================================================="
Print #1, "TOTAL DE VARIEDAD DE PRODUCTOS: " & nTotpro
Print #1, "  PRODUCTOS MAS ALTOS OFICINAS: " & nVarOfi & "     Importe: " & nMasAltoO
Print #1, "  PRODUCTOS MAS ALTOS MCABRERA: " & nVarMc & "     Importe: " & nMasAltomc

Print #1, " "
Print #1, "TOTAL DEL PEDIDO $ OFICINAS: "; Format(NTOTCJO, "#######0.00")
Print #1, "TOTAL DEL PEDIDO $ MCABRERA: "; Format(nTotCjm, "#######0.00")
Print #1, "DIFERENCIA                 : "; Format(nTotCjm - NTOTCJO, "#######0.00")
Close #1   ' Cierra el archivo de reporte
'Handle = Shell("NOTEPAD " & App.Path & "\FRANQ.TXT", 1)
prod = Round(nVarMc * 0.3, 0)
Naum = nMasAltomc / IIf(prod = 0, 1, prod)
rs.Close
rs.Open "SELECT * FROM franqtmp WHERE dif < 0 ORDER BY DIF", cn, adOpenKeyset, adLockOptimistic, adCmdText
N = 0: NtOTaUM = 0
Do While NtOTaUM < nMasAltomc And Not rs.EOF
   Naum = Abs(rs!dif)
   NtOTaUM = NtOTaUM + Naum
   If NtOTaUM > nMasAltomc Then
      'MsgBox "UPDATE franqtmp SET premod = " & Round(RS!MCAB + ((NtOTaUM - nMasAltomc) / RS!CAJAS), 2) & " WHERE consec = '" & RS!CONSEC & "'"
      cn.Execute "UPDATE franqtmp SET premod = " & Round(rs!MCAB + ((NtOTaUM - nMasAltomc) / rs!cajas), 2) & " WHERE consec = '" & rs!CONSEC & "'"
   Else
      cn.Execute "UPDATE franqtmp SET premod = " & rs!Ofi & " WHERE consec = '" & rs!CONSEC & "'"
   End If
   rs.MoveNext
Loop
cn.Execute "UPDATE FRANQTMP SET PREACT = PREMOD WHERE DIF < 0 AND PREMOD > 0"
cn.Execute "UPDATE franqtmp SET PREACT = MCAB WHERE PREMOD = 0"
cn.Execute "UPDATE FRANQTMP SET PREACT = OFI  WHERE DIF >= 0"
cn.Execute "UPDATE ventas_det set precio = PREACT, importe = (PREACT * cantidad) + (preciop * cantidadp) FROM franqtmp WHERE consec = cl_producto and noventa =  " & AdoVentas.Recordset!noventa
cn.Execute "UPDATE VENTAS set montototal = ( select sum(importe) from ventas_det where noventa = " & AdoVentas.Recordset!noventa & ") where noventa = " & AdoVentas.Recordset!noventa
Call COMPARA1
End Sub

'***************************
'Precios prorrateados
'***************************
Private Sub COMPARA1()
Dim RSMDB As ADODB.Recordset
Dim CNMDB As ADODB.Connection
Dim rs As ADODB.Recordset
Dim Valor As String

Set rs = New ADODB.Recordset
Set RSMDB = New ADODB.Recordset
Set CNMDB = New ADODB.Connection
CNMDB.Open "DSN=PITICOMDB;DBQ=P:\PITICO\PITICO10\PITICO10.mdb;DefaultDir=P:\PITICO\PITICO10;DriverId=25;FIL=MS Access;MaxBufferSize=2048;PageTimeout=5;UID=admin;PWD=PORTATIL"
rs.Open "SELECT T.CONSEC,d.precio as precio,d.preciop,d.cantidad, d.cantidadp,T.descripc, LTRIM(STR(T.paquetes)) + ' x ' + LTRIM(STR(T.CONTENID,10,3)) + space(1) + RTRIM(t.medida) AS MEDIDA, t.paquetes as paquetes FROM franqtmp f, VENTAS_det d, tfproduc T WHERE f.consec = D.cl_producto and f.consec = t.consec AND T.consec = cl_producto AND D.cancelado = 0 and D.noventa = " & Me.AdoVentas.Recordset!noventa & " ORDER BY T.descripc, medida", cn, adOpenStatic, adLockOptimistic, adCmdText

Open App.Path & "\" & CStr(AdoVentas.Recordset!noventa) & "p.TXT" For Output As #1  ' Abre el archivo para operaciones de salida.
Print #1, Tab(15); "VIVERES Y LICORES S.A DE C.V."; Space(5); Format(date, "LONG DATE"); "; "; Time()
Print #1, Tab(10); "CARBONERA 1016 COL. TRINIDAD DE LAS HUERTAS"
Print #1, Tab(10); "COMPARATIVO DE VENTAS PRORRATEADO ENTRE MIGUEL CAB. Y OFICINAS CENTRALES"
Print #1, "PEDIDO DEL CLIENTE " & cmbCliente.Text & " EN EL FOLIO DE PREVENTA " & AdoVentas.Recordset!FOLPREVENTA
Print #1, "===================================================================================================="
Print #1, "                      D E S C R I P C I O N                        CJ.    CJ.OFI   CJ.MCAB     DIF.  "
Print #1, "===================================================================================================="
If AdoVentas.Recordset!clcliente = 1630 Then   'Miahutlan
   PRECIO = "PRECIO3"
Else
   PRECIO = "PRECIO4"
End If
nTotCjm = 0: NTOTCJO = 0: nVarMc = 0: nVarOfi = 0: nMasAltomc = 0: nMasAltoO = 0
nTotpro = 0
cmen = stb1.Panels(1).Text
While Not rs.EOF
    RSMDB.Open "SELECT precio1,precio3,precio4,descripc,LTRIM(STR(paquetes)) + ' x ' + LTRIM(STR(CONTENID)) + space(1) + RTRIM(medida) AS MEDIDA1 FROM preprod,tfproduc WHERE consec = preclave and preclave = '" & rs!CONSEC & "'", CNMDB, adOpenForwardOnly, adLockOptimistic, adCmdText
    stb1.Panels(1).Text = rs!descripc
    stb1.Refresh
    If Not (RSMDB.BOF And RSMDB.EOF) Then
       nTotpro = nTotpro + 1
       premc = rs!PRECIO
       'El precio de Mc es mas alto que carbonera
       If premc > RSMDB.Fields(PRECIO).Value Then
          Print #1, "*";
          nMasAltomc = nMasAltomc + (premc - RSMDB.Fields(PRECIO).Value) * rs!cantidad
          nVarMc = nVarMc + 1
       ElseIf premc < RSMDB.Fields(PRECIO).Value Then
          Print #1, "";
          nVarOfi = nVarOfi + 1
          nMasAltoO = nMasAltoO + (RSMDB.Fields(PRECIO).Value - premc) * rs!cantidad
       End If
       Valor = Space(50)
       LSet Valor = Mid(Trim(rs!descripc), 1, 43)
       Print #1, Valor;
       Valor = Space(15)
       RSet Valor = rs!medida
       Print #1, Valor;
       Valor = Space(5)
       RSet Valor = rs!cantidad
       Print #1, Valor;
       'Precio de Oficinas
       Valor = Space(10)
       RSet Valor = Format(RSMDB.Fields(PRECIO).Value * rs!cantidad, "#######0.00")
       Print #1, Valor;
       'Precio de Mc
       RSet Valor = Format(premc * rs!cantidad, "#######0.00")
       Print #1, Valor;
       RSet Valor = Format((premc * rs!cantidad) - (RSMDB.Fields(PRECIO).Value * rs!cantidad), "#######0.00")
       Print #1, Valor
       
       NTOTCJO = NTOTCJO + RSMDB.Fields(PRECIO).Value * rs!cantidad
       nTotCjm = nTotCjm + premc * rs!cantidad
    Else
       Valor = Space(50)
       LSet Valor = Mid(Trim(rs!descripc), 1, 43)
       Print #1, Valor;
       Valor = Space(16)
       RSet Valor = rs!medida
       Print #1, Valor;
       Valor = Space(5)
       RSet Valor = rs!cantidad
       Print #1, Valor
    End If
    RSMDB.Close
    rs.MoveNext
Wend
stb1.Panels(1).Text = cmen
stb1.Refresh
Print #1, "===================================================================================================="
Print #1, "TOTAL DE VARIEDAD DE PRODUCTOS: " & nTotpro
Print #1, "  PRODUCTOS MAS ALTOS OFICINAS: " & nVarOfi & "     Importe: " & nMasAltoO
Print #1, "  PRODUCTOS MAS ALTOS MCABRERA: " & nVarMc & "     Importe: " & nMasAltomc

Print #1, " "
Print #1, "TOTAL DEL PEDIDO $ OFICINAS: "; Format(NTOTCJO, "#######0.00")
Print #1, "TOTAL DEL PEDIDO $ MCABRERA: "; Format(nTotCjm, "#######0.00")
Print #1, "DIFERENCIA                 : "; Format(nTotCjm - NTOTCJO, "#######0.00")
Close #1   ' Cierra el archivo de reporte
Handle = Shell("NOTEPAD " & App.Path & "\" & CStr(AdoVentas.Recordset!noventa) & "p.TXT", vbMinimizedNoFocus)
End Sub

