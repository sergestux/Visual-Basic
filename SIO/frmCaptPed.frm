VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmCaptPed 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Esta forma se utiliza para CAPTURA, CONFIRMACION y RECIBIR Pedido"
   ClientHeight    =   8625
   ClientLeft      =   -45
   ClientTop       =   240
   ClientWidth     =   11895
   Icon            =   "frmCaptPed.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   11895
   WindowState     =   2  'Maximized
   Begin ComctlLib.StatusBar stbmensajes 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   46
      Top             =   8280
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   609
      Style           =   1
      SimpleText      =   "                                                                                           Para salir presione la tecla [Esc]"
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cmdlg 
      Left            =   0
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraAvance 
      Caption         =   "Cargando especificaciones de productos "
      Height          =   855
      Left            =   3720
      TabIndex        =   35
      Top             =   3120
      Visible         =   0   'False
      Width           =   4935
      Begin ComctlLib.ProgressBar pgb 
         Height          =   255
         Left            =   240
         TabIndex        =   47
         Top             =   360
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
   End
   Begin MSAdodcLib.Adodc AdoNotEnt 
      Height          =   330
      Left            =   2520
      Top             =   6360
      Visible         =   0   'False
      Width           =   2280
      _ExtentX        =   4022
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
      Caption         =   "AdoNotEnt"
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
      DataField       =   "ClaveNota"
      DataSource      =   "AdoNotEnt"
      Height          =   285
      Index           =   12
      Left            =   10320
      TabIndex        =   34
      Top             =   5640
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSACAL.Calendar Cal1 
      Height          =   1935
      Left            =   4320
      TabIndex        =   16
      Top             =   960
      Visible         =   0   'False
      Width           =   2655
      _Version        =   524288
      _ExtentX        =   4683
      _ExtentY        =   3413
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2000
      Month           =   6
      Day             =   14
      DayLength       =   1
      MonthLength     =   2
      DayFontColor    =   0
      FirstDay        =   2
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   0   'False
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoFacturas 
      Height          =   330
      Left            =   6840
      Top             =   7320
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
   Begin MSAdodcLib.Adodc AdoCatUsu 
      Height          =   330
      Left            =   6840
      Top             =   7320
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
   Begin VB.TextBox txtCampos 
      Alignment       =   2  'Center
      DataField       =   "p_fecconfirma"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/mm/yy h:nn AM/PM"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   0
      EndProperty
      DataSource      =   "AdoPedidos"
      Height          =   285
      Index           =   7
      Left            =   8280
      TabIndex        =   14
      Top             =   4440
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CheckBox chkFields 
      Caption         =   "Pedido &Confirmado"
      DataField       =   "p_situacion"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   5
      Left            =   8160
      TabIndex        =   11
      Top             =   4080
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.ListBox lstProd 
      Appearance      =   0  'Flat
      Height          =   1200
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   26
      Top             =   5160
      Visible         =   0   'False
      Width           =   7575
   End
   Begin VB.CheckBox chkFields 
      Caption         =   "Pedido r&ecibido"
      DataField       =   "p_recibido"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   6
      Left            =   8160
      TabIndex        =   12
      Top             =   5040
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSAdodcLib.Adodc AdoDetPed 
      Height          =   330
      Left            =   0
      Top             =   6360
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
   Begin MSAdodcLib.Adodc AdoCatPro 
      Height          =   330
      Left            =   0
      Top             =   6720
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
      Caption         =   "AdoCatPro"
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
   Begin MSAdodcLib.Adodc AdoPedidos 
      Height          =   330
      Left            =   5880
      Top             =   0
      Visible         =   0   'False
      Width           =   2520
      _ExtentX        =   4445
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   1
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
      Caption         =   "AdoPedidos"
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
      Left            =   1680
      MaxLength       =   10
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Frame fraGenerales 
      Caption         =   "Datos generales"
      ForeColor       =   &H8000000D&
      Height          =   2535
      Left            =   240
      TabIndex        =   17
      Top             =   480
      Visible         =   0   'False
      Width           =   11295
      Begin VB.CheckBox chksugerido 
         Caption         =   "S&ugerido"
         DataField       =   "p_sugerido"
         DataSource      =   "AdoPedidos"
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
         Left            =   9000
         TabIndex        =   45
         Top             =   360
         Width           =   1335
      End
      Begin VB.ComboBox cmbJefeRec 
         Height          =   315
         Left            =   3120
         TabIndex        =   13
         Top             =   2040
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.TextBox txtCampos 
         DataField       =   "p_receptora"
         DataSource      =   "AdoPedidos"
         Height          =   285
         Index           =   9
         Left            =   2160
         MaxLength       =   3
         TabIndex        =   10
         Top             =   2040
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ComboBox cmbUsuarios 
         Height          =   315
         Left            =   3120
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   1080
         Visible         =   0   'False
         Width           =   3975
      End
      Begin VB.ComboBox cmbSucursal 
         Height          =   315
         Left            =   3120
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   720
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.ComboBox cmbProved 
         Height          =   315
         Left            =   3120
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   360
         Visible         =   0   'False
         Width           =   5415
      End
      Begin VB.TextBox txtCampos 
         DataField       =   "P_Solicita"
         DataSource      =   "AdoPedidos"
         Height          =   285
         Index           =   3
         Left            =   2160
         MaxLength       =   3
         TabIndex        =   5
         Top             =   1080
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtCampos 
         DataField       =   "P_Observaciones"
         DataSource      =   "AdoPedidos"
         Height          =   525
         Index           =   6
         Left            =   2160
         MaxLength       =   50
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   1440
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.TextBox txtCampos 
         Alignment       =   2  'Center
         DataField       =   "P_Fecent"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/mm/yy "
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
         DataSource      =   "AdoPedidos"
         Height          =   285
         Index           =   5
         Left            =   9600
         TabIndex        =   8
         Top             =   1560
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtCampos 
         Alignment       =   2  'Center
         DataField       =   "P_FecPed"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/mm/yy h:nn AM/PM"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
         DataSource      =   "AdoPedidos"
         Height          =   285
         Index           =   4
         Left            =   9120
         TabIndex        =   7
         Top             =   1080
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox txtCampos 
         DataField       =   "P_Sucursal"
         DataSource      =   "AdoPedidos"
         Height          =   285
         Index           =   2
         Left            =   2160
         MaxLength       =   3
         TabIndex        =   3
         Top             =   720
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtCampos 
         DataField       =   "P_Proveedor"
         DataSource      =   "AdoPedidos"
         Height          =   285
         Index           =   1
         Left            =   2160
         MaxLength       =   5
         TabIndex        =   1
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lbletiquetas 
         Caption         =   "Clave del jefe de recibo"
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   31
         Top             =   2040
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lblDesprov 
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   2640
         TabIndex        =   25
         Top             =   360
         Width           =   3975
      End
      Begin VB.Label lbletiquetas 
         Caption         =   "Observaciones"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   24
         Top             =   1440
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lbletiquetas 
         Caption         =   "Clave del jefe de depto."
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   23
         Top             =   1080
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lbletiquetas 
         Caption         =   "Fecha de entrega del producto"
         Height          =   255
         Index           =   5
         Left            =   7200
         TabIndex        =   22
         Top             =   1560
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label lbletiquetas 
         Caption         =   "Fecha de elab. del pedido"
         Height          =   255
         Index           =   4
         Left            =   7200
         TabIndex        =   21
         Top             =   1080
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lbletiquetas 
         Caption         =   "Clave de la sucursal"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   20
         Top             =   720
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lbletiquetas 
         Caption         =   "Clave del proveedor"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   1695
      End
   End
   Begin MSAdodcLib.Adodc AdoCatProved 
      Height          =   330
      Left            =   2520
      Top             =   6720
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
      Caption         =   "adoCatProved"
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
      Align           =   2  'Align Bottom
      BackColor       =   &H00808080&
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   11835
      TabIndex        =   29
      Top             =   7665
      Width           =   11895
      Begin VB.CommandButton Command1 
         Caption         =   "&Nota costo"
         Height          =   450
         Left            =   6480
         Picture         =   "frmCaptPed.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Presentacion preeliminar de Reporte"
         Top             =   60
         Width           =   900
      End
      Begin VB.CommandButton cmdInven 
         Caption         =   "&Invent."
         Height          =   450
         Left            =   4800
         Picture         =   "frmCaptPed.frx":0974
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Ver inventario"
         Top             =   60
         Width           =   795
      End
      Begin VB.CommandButton cmdReporte 
         Caption         =   "&Nota ent."
         Enabled         =   0   'False
         Height          =   450
         Left            =   7440
         Picture         =   "frmCaptPed.frx":0A76
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Presentacion preeliminar de Reporte"
         Top             =   60
         Width           =   885
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Regresar"
         Height          =   450
         Left            =   10920
         Picture         =   "frmCaptPed.frx":0FA8
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Cancela los cambios en los datos generales y regresa a la pantalla principal"
         Top             =   60
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         Height          =   450
         Left            =   10080
         Picture         =   "frmCaptPed.frx":111A
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Graba los cambios y regresa a la pantalla principal"
         Top             =   60
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton cmdDespla 
         Caption         =   "&Desplaz."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   8400
         Picture         =   "frmCaptPed.frx":128C
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Despliega pantalla de desplazamiento de productos"
         Top             =   60
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton cmdBajas 
         Caption         =   "&Bajas"
         Height          =   450
         Left            =   -360
         Picture         =   "frmCaptPed.frx":13FE
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   -120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton CmdRefresh 
         Caption         =   "&Actualizar"
         Height          =   450
         Left            =   9240
         Picture         =   "frmCaptPed.frx":1570
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Actualizar el detalle del pedido"
         Top             =   60
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton cmdExporta 
         Caption         =   "&Exportar"
         Height          =   450
         Left            =   5640
         Picture         =   "frmCaptPed.frx":1672
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Exportar el pedido para Abastecimiento a formato Foxpro"
         Top             =   60
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   360
         TabIndex        =   44
         Top             =   120
         Width           =   3975
      End
   End
   Begin VB.TextBox txtCampos 
      Alignment       =   2  'Center
      DataField       =   "p_fecentreal"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/mm/yy h:nn AM/PM"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   0
      EndProperty
      DataSource      =   "AdoPedidos"
      Height          =   285
      Index           =   8
      Left            =   8280
      TabIndex        =   15
      Top             =   5640
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.PictureBox cRpt 
      Height          =   480
      Left            =   4680
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   50
      Top             =   7320
      Width           =   1200
   End
   Begin MSAdodcLib.Adodc AdoDbf 
      Height          =   330
      Left            =   2880
      Top             =   0
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
      Connect         =   "DSN=PITICODBF"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "PITICODBF"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "AdoDbf Pedido"
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
   Begin MSDataGridLib.DataGrid dbgrdfactu 
      Bindings        =   "frmCaptPed.frx":1774
      Height          =   735
      Left            =   240
      TabIndex        =   36
      Top             =   3120
      Visible         =   0   'False
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   1296
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
            Alignment       =   2
            ColumnWidth     =   1110.047
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column06 
            Alignment       =   2
            ColumnWidth     =   1124.787
         EndProperty
         BeginProperty Column07 
            Alignment       =   2
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column08 
            Alignment       =   2
            ColumnWidth     =   1110.047
         EndProperty
         BeginProperty Column09 
            Alignment       =   2
            ColumnWidth     =   1124.787
         EndProperty
         BeginProperty Column10 
            Alignment       =   2
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column11 
            Alignment       =   2
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column12 
            Alignment       =   2
            ColumnWidth     =   1035.213
         EndProperty
         BeginProperty Column13 
            Alignment       =   2
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column14 
            Alignment       =   2
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column15 
            Alignment       =   2
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column16 
            Alignment       =   2
            ColumnWidth     =   1035.213
         EndProperty
         BeginProperty Column17 
            Alignment       =   2
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column18 
            Alignment       =   2
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column19 
            Alignment       =   2
            ColumnWidth     =   975.118
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid dbgrdDetPed 
      Bindings        =   "frmCaptPed.frx":178E
      Height          =   3495
      Left            =   240
      TabIndex        =   28
      Top             =   3960
      Visible         =   0   'False
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   6165
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      BorderStyle     =   0
      ColumnHeaders   =   -1  'True
      HeadLines       =   1.5
      RowHeight       =   17
      TabAction       =   2
      RowDividerStyle =   6
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
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
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "df_prod"
         Caption         =   "    CLAVE "
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
         DataField       =   "df_pedido"
         Caption         =   "df_pedido"
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
         DataField       =   "df_cantidad"
         Caption         =   "CANT.SUG. CAJAS"
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
         DataField       =   "df_cantsol"
         Caption         =   "CANT.SOL.CAJAS"
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
         DataField       =   "df_cantreal"
         Caption         =   "CANT. REC. CAJA"
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
      BeginProperty Column05 
         DataField       =   "df_cantsolp"
         Caption         =   "CANT.SOL.PZAS."
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
      BeginProperty Column06 
         DataField       =   "df_cantrealp"
         Caption         =   "CANT. REC. PZAS."
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
         DataField       =   "df_costo"
         Caption         =   "   COSTO"
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
            ColumnAllowSizing=   0   'False
            Object.Visible         =   -1  'True
            ColumnWidth     =   1110.047
         EndProperty
         BeginProperty Column01 
            ColumnAllowSizing=   0   'False
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnAllowSizing=   0   'False
            Object.Visible         =   0   'False
            ColumnWidth     =   1409.953
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   1395.213
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            Object.Visible         =   -1  'True
            ColumnWidth     =   1440
         EndProperty
         BeginProperty Column07 
            Object.Visible         =   0   'False
            ColumnWidth     =   1094.74
         EndProperty
      EndProperty
   End
   Begin VB.Label lbletiquetas 
      Alignment       =   2  'Center
      Caption         =   "Nota de entrada"
      Height          =   255
      Index           =   12
      Left            =   10320
      TabIndex        =   33
      Top             =   5400
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lbletiquetas 
      Alignment       =   2  'Center
      Caption         =   "Fecha de recepcion"
      Height          =   255
      Index           =   8
      Left            =   8280
      TabIndex        =   32
      Top             =   5400
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label lbletiquetas 
      Alignment       =   2  'Center
      Caption         =   "Fecha de confirmacion"
      Height          =   255
      Index           =   7
      Left            =   9600
      TabIndex        =   30
      Top             =   3600
      Visible         =   0   'False
      Width           =   1935
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
      ForeColor       =   &H80000006&
      Height          =   1335
      Left            =   8040
      TabIndex        =   27
      Top             =   6240
      Width           =   3615
   End
   Begin VB.Label lbletiquetas 
      Caption         =   "Clave del pedido"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   18
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmCaptPed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const ColIndex_lstProd = 0       'Número de columna que tendra una lista desplegable
Private ntext As Integer         'Indice de los cuadros de texto
Private pConfirmado As Boolean   'El pedido ya se confirmo
Private PedRecibido As Boolean   'El Pedido ya se recibio
Private nFrecuencia              'Dias que transcurren para la entrega del pedido
Private rsttemp As ADODB.Recordset
Private cEnca As String

Private Sub AdoDetPed_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
cmdExporta.Enabled = Not (AdoDetPed.Recordset.BOF And AdoDetPed.Recordset.EOF)
End Sub

Private Sub Cal1_DblClick()
txtcampos(ntext).Text = Cal1.Value
Cal1.Visible = False
txtcampos(ntext).SetFocus
'SendKeys "{TAB}"
keybd_event &H9, 0, 0, 0:   keybd_event &H9, 0, &H2, 0
End Sub

Private Sub Cal1_LostFocus()
txtcampos(ntext).Text = Cal1.Value
Cal1.Visible = False
End Sub

Private Sub chkFields_Click(Index As Integer)
Select Case Index
Case 5  'Confirmado
        'Para que no marque error al no existir ningun registro
        If Not (AdoPedidos.Recordset.BOF And AdoPedidos.Recordset.BOF) Then
           If Not AdoPedidos.Recordset!p_situacion Then
                lbletiquetas(7).Visible = chkFields(Index).Value = 1
                txtcampos(7).Visible = chkFields(Index).Value = 1
                txtcampos(7).Text = date + Time
                txtcampos(7).Enabled = False
           End If
        End If
Case 6  'Recibido
        If Not (AdoPedidos.Recordset.BOF And AdoPedidos.Recordset.BOF) Then
          If Not AdoPedidos.Recordset!P_recibido And lProvAbto <> cInd Then
            Me.dbgrdfactu.Visible = (chkFields(Index).Value = 1)
            lbletiquetas(8).Visible = chkFields(Index).Value = 1
            txtcampos(8).Visible = chkFields(Index).Value = 1
            txtcampos(8).Text = date + Time
            txtcampos(8).Enabled = False
          End If
        End If
End Select
End Sub

Private Sub chkFields_GotFocus(Index As Integer)
  Cal1.Visible = False
End Sub

Private Sub chksugerido_Click()
If chksugerido.Value = 1 Then chkFields(6).Visible = False
End Sub

Private Sub cmbJefeRec_DblClick()
  'SendKeys "{TAB}"
  keybd_event &H9, 0, 0, 0:   keybd_event &H9, 0, &H2, 0
End Sub

Private Sub cmbJefeRec_KeyPress(KeyAscii As Integer)
If KeyAscii Then
   KeyAscii = 0
   'SendKeys "{Tab}"
   keybd_event &H9, 0, 0, 0:   keybd_event &H9, 0, &H2, 0
End If

End Sub

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
   'SendKeys "{TAB}"
   keybd_event &H9, 0, 0, 0:   keybd_event &H9, 0, &H2, 0
End If

End Sub

Private Sub cmbProved_DblClick()
  'SendKeys "{TAB}"
  keybd_event &H9, 0, 0, 0:   keybd_event &H9, 0, &H2, 0
End Sub

Private Sub cmbProved_GotFocus()
RESP = SendMessageLong(cmbProved.hwnd, &H14F, True, 1)
End Sub

Private Sub cmbproved_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
If KeyAscii = 13 Then
   KeyAscii = 0
   'SendKeys "{Tab}"
   keybd_event &H9, 0, 0, 0:   keybd_event &H9, 0, &H2, 0
End If

End Sub

Private Sub cmbProved_LostFocus()
RESP = SendMessageLong(cmbProved.hwnd, &H14F, False, 1)
End Sub

Private Sub cmbProved_Validate(Cancel As Boolean)
Dim N As Integer
If nOp <> 1 Then
   Exit Sub
End If
If cmbProved.Text = "" Or IsNull(cmbProved.Text) Then
   MsgBox "Debe seleccionar un proveedor de la lista desplegable", vbExclamation
   cmbProved.SetFocus
   Cancel = True
Else
   AdoCatProved.Recordset.MoveFirst
   AdoCatProved.Recordset.Find "NomProve = '" & cmbProved.Text & "'"
   If AdoCatProved.Recordset.EOF = True Then
      MsgBox "Debe seleccionar un proveedor de la lista desplegable", vbExclamation
      cmbProved.SetFocus
      Cancel = True
   Else
   txtcampos(1).Text = AdoCatProved.Recordset!prove
   txtcampos(1).SetFocus
   End If
End If
End Sub

Private Sub cmbSucursal_DblClick()
  'SendKeys "{TAB}"
  keybd_event &H9, 0, 0, 0:   keybd_event &H9, 0, &H2, 0
End Sub

Private Sub cmbSucursal_KeyPress(KeyAscii As Integer)
If KeyAscii Then
   KeyAscii = 0
   'SendKeys "{Tab}"
   keybd_event &H9, 0, 0, 0:   keybd_event &H9, 0, &H2, 0
End If

End Sub

Private Sub cmbsucursal_LostFocus()
Dim RST As ADODB.Recordset
Set RST = New ADODB.Recordset
RST.Open "SELECT * FROM cattienda WHERE tidescrip  = '" & Trim(cmbSucursal.Text) & "'", cn, adOpenForwardOnly, adLockOptimistic, adCmdText
If RST.EOF = True Then
   MsgBox "EXISTEN ALGUN ERROR EN LA BASE DE SUCURSALES O AREA", vbCritical
   Cancel = True
Else
  tclave = RST!ticlave
  txtcampos(2).Text = tclave
End If

End Sub

Private Sub cmbUsuarios_DblClick()
 'SendKeys "{TAB}"
 keybd_event &H9, 0, 0, 0:   keybd_event &H9, 0, &H2, 0
End Sub

Private Sub cmbUsuarios_KeyPress(KeyAscii As Integer)
If KeyAscii Then
   KeyAscii = 0
   'SendKeys "{Tab}"
   keybd_event &H9, 0, 0, 0:   keybd_event &H9, 0, &H2, 0
End If

End Sub

Private Sub cmbUsuarios_Validate(Cancel As Boolean)
AdoCatUsu.Recordset.MoveFirst
AdoCatUsu.Recordset.Find "Name = '" & cmbUsuarios.Text & "'"
If AdoCatUsu.Recordset.EOF = True Then
   MsgBox "Debe seleccionar un usuario de la lista desplegable", vbExclamation
   cmbUsuarios.SetFocus
   Cancel = True
Else
   txtcampos(3).Text = AdoCatUsu.Recordset!clave
   'txtCampos(3).SetFocus
End If

End Sub

Private Sub cmdBajas_Click()
 'BAJAS
 AdoPedidos.Recordset.Find "P_pedido = '" & Trim(txtcampos(0).Text) & "'"
 RESP = MsgBox("Deseas borrar este pedido", vbInformation + vbYesNo, "Confirma eliminación")
 If RESP = vbYes Then AdoPedidos.Recordset.Delete
 Unload Me
End Sub

Private Sub cmdCancelar_Click()
  If Not PedRecibido Then AdoPedidos.Refresh
  Unload Me
End Sub

Private Sub cmdDespla_Click()
Dim FecPed
Dim Pedido
Dim CNP As ADODB.Connection
Dim cejecuta As String
Dim mensAnt As String
  If Trim(txtcampos(1).Text) = "ABA" Then
     StbMensajes.SimpleText = Space(40) & "Espere un momento obteniendo productos con existencia en Bodega Miguel cabrera"
  Else
     StbMensajes.SimpleText = Space(40) + "Espere un momento obteniendo desplazamientos de productos en los últimos tres períodos"
  End If
  StbMensajes.Refresh
 
  Set CNP = New Connection
  CNP.ConnectionString = cCadConex
  CNP.ConnectionTimeout = 0
  CNP.CommandTimeout = 0
  CNP.Open
  FecPed = Str(Day(Trim(txtcampos(4)))) + "/" + Str(Month(Trim(txtcampos(4)))) + "/" + Str(Year(Trim(txtcampos(4))))
  If tipotienda = 1 Then
     cejecuta = "semanaMU '" & Trim(txtcampos(1).Text) & "','" & FecPed & "','" & Trim(Mid(cCveDesUsu, 1, 3)) & "'"
  Else
     cejecuta = "semanaMU '" & Trim(txtcampos(1).Text) & "','" & FecPed & "','" & Trim(Mid(cCveDesUsu, 1, 3)) & "',7"
  End If
  CNP.Execute (cejecuta)
  CNP.Close
  Set CNP = Nothing
  StbMensajes.SimpleText = mensAnt
  StbMensajes.Refresh
  If Trim(txtcampos(1).Text) = "ABA" Then
     lpprov = False
     frmdespla.dbgrdTend.Caption = "PRODUCTOS CON EXISTENCIA EN BODEGA MIGUEL CABRERA"
     'frmdespla.dbgrdTend.Splits(1).Columns(9).Width = 0
     frmdespla.dbgrdTend.Splits(1).Columns(10).Width = 0
     frmdespla.dbgrdTend.Splits(1).Columns(12).Width = 0
     frmdespla.dbgrdTend.Splits(1).Columns(13).Width = 0
     frmdespla.dbgrdTend.Splits(1).Columns(7).Caption = "CAJAS CDM"
     frmdespla.dbgrdTend.Splits(1).Columns(8).Caption = "PZAS CDM"
     frmdespla.Caption = "Productos con existencia en Bodega Miguel Cabrera"
  Else
     nDias = 7
     frmdespla.dbgrdTend.Splits(1).Columns(5).Caption = Format(date - nDias, "dd/mm/yy") & "-" & Chr(13) & Format(date, "dd/mm/yy")
     frmdespla.dbgrdTend.Splits(1).Columns(4).Caption = Format(date - nDias * 2, "dd/mm/yy") & "-" & Chr(13) & Format(date - nDias - 1, "dd/mm/yy")
     frmdespla.dbgrdTend.Splits(1).Columns(3).Caption = Format(date - nDias * 3, "dd/mm/yy") & "-" & Chr(13) & Format(date - nDias * 2 - 1, "dd/mm/yy")
     frmdespla.dbgrdTend.Caption = "VENTAS REALIZADAS EN LOS ULTIMOS TRES PERIODOS CON RANGO DE  " & nDias & " DIAS "
     frmdespla.Caption = "Desplazamiento de ventas de " & Me.cmbProved.Text
  End If
  If frmdespla.AdoTend.RecordSource <> "" Then frmdespla.AdoTend.Refresh
  frmdespla.Show
  Set CNP = Nothing
End Sub

'Exporta pedidos SUGERIDOS y de ABASTECIMIENTO a tablas Dbf de Visual Fox Prod para enviar a Carbonera
Private Sub CmdExporta_Click()
On Error GoTo Error:
Dim rsttemp As ADODB.Recordset
   StbMensajes.SimpleText = Space(45) & "Grabando archivo "
   StbMensajes.Refresh
   Set rsttemp = New ADODB.Recordset
   rsttemp.Open "SELECT * FROM DetalleFactura,tfproduc WHERE df_Pedido = '" & txtcampos(0).Text & "' AND df_sugerido = " & IIf(Trim(txtcampos(1).Text) = "ABA", 0, 1) & " AND consec = df_prod ORDER BY descripc,contenid", cn, adOpenForwardOnly, adLockOptimistic, adCmdText
   Open "P:\BUZON\PED" & Mid(cSucursal, 6, 3) & ".TXT" For Output As #1
   While Not rsttemp.EOF
      StbMensajes.SimpleText = Space(75) & "Exportando producto con la clave: " & CStr(rsttemp!df_prod)
      StbMensajes.Refresh
      Print #1, Trim(rsttemp!df_prod) & "|" & rsttemp!df_cantsol & "|" & rsttemp!df_cantsolp
      rsttemp.MoveNext
   Wend
   Close #1
   rsttemp.Close
   Set rsttemp = Nothing
   MsgBox "SE EXPORTARON: " & CStr(AdoDetPed.Recordset.RecordCount) & " PRODUCTOS", vbInformation
  Exit Sub
Error:
   MsgBox UCase(Err.Description), vbCritical
   StbMensajes.SimpleText = cMenAnt
End Sub

Private Sub cmdGrabar_Click()
Dim rsttemp As ADODB.Recordset
Dim rstInvent As ADODB.Recordset
Dim lTrans As Boolean
On Error GoTo Error:
  cMensaje = StbMensajes.SimpleText
  StbMensajes.SimpleText = Space(90) + "Espere un momento, grabando datos ..."
  StbMensajes.Refresh
  lTrans = False   'Para determinar si se deshace la transaccion
  If nOp = 1 And cModo = "CAPTURARPEDIDO" Then   'Altas
     'En los proveedores Instantaneos o abiertos
     If lProvAbto = Ins Or lProvAbto = Abi Then
        If chkFields(5).Value = 0 Then
            MsgBox "ES NECESARIO ACTIVAR LA CASILLA DE PEDIDO CONFIRMADO", vbExclamation
            Exit Sub
        'Una vez confirmado el pedido deben de recibirlo en los Instantaneos
        ElseIf chkFields(6).Value = 0 And Me.dbgrdDetPed.Visible = True And lProvAbto = Ins Then
            MsgBox "ES NECESARIO ACTIVAR LA CASILLA DE PEDIDO RECIBIDO", vbExclamation
            Exit Sub
        'Solo cuando ha sido confirmado para que me deje grabar
        ElseIf chkFields(5).Value = 1 And chkFields(6).Value = 0 Then
            AdoPedidos.Recordset!P_recibido = 0
        End If
          AdoPedidos.Recordset!p_Pedido = txtcampos(0).Text
          AdoPedidos.Recordset!p_situacion = 1
          AdoPedidos.Recordset!p_fecConfirma = date + Time
          AdoPedidos.Recordset!p_proveedor = txtcampos(1).Text
          AdoPedidos.Recordset!p_sucursal = txtcampos(2).Text
          AdoPedidos.Recordset!p_fecent = IIf(Trim(txtcampos(5).Text) = "", Null, txtcampos(5).Text)
          AdoPedidos.Recordset!p_solicita = txtcampos(3).Text
          'AdoPedidos.Recordset!p_observaciones = txtCampos(6).Text
          AdoPedidos.Recordset!p_fecped = txtcampos(4).Text
          If chkFields(6).Value = 1 Then
             AdoPedidos.Recordset!p_fecentreal = txtcampos(8).Text
             AdoPedidos.Recordset!P_recibido = 1
          End If
     Else 'Si es indirecto
          AdoPedidos.Recordset!p_situacion = False
          AdoPedidos.Recordset!P_recibido = False
          AdoPedidos.Recordset!p_Pedido = txtcampos(0).Text
          AdoPedidos.Recordset!p_proveedor = txtcampos(1).Text
          AdoPedidos.Recordset!p_sucursal = txtcampos(2).Text
          AdoPedidos.Recordset!p_fecent = IIf(Trim(txtcampos(5).Text) = "", Null, txtcampos(5).Text)
          AdoPedidos.Recordset!p_solicita = txtcampos(3).Text
          If Trim(txtcampos(6).Text) <> "" Then AdoPedidos.Recordset!p_observaciones = txtcampos(6).Text
          AdoPedidos.Recordset!p_fecped = txtcampos(4).Text
     End If
     cmdGrabar.Enabled = False
     AdoPedidos.Recordset.Update
     cmdReporte.Enabled = True
     cmdReporte.Caption = "&Ped. sug."
     
     'En los proveedores Abiertos e Instantaneos
     If lProvAbto = Ins Or lProvAbto = Abi Then
       frmCaptPed.dbgrdDetPed.Visible = True
       If lProvAbto = Ins Then
          chkFields(6).Visible = (chksugerido.Value = 0)
          frmCaptPed.cmdGrabar.Enabled = True
          frmCaptPed.dbgrdDetPed.Columns(0).Button = True
          frmCaptPed.dbgrdDetPed.Columns(2).Visible = False
          frmCaptPed.dbgrdDetPed.Columns(4).Visible = True
          frmCaptPed.dbgrdDetPed.Columns(5).Visible = True
          frmCaptPed.dbgrdDetPed.Columns(6).Visible = True
       Else
          frmCaptPed.dbgrdDetPed.Columns(0).Button = True
          frmCaptPed.dbgrdDetPed.Columns(2).Visible = True
          frmCaptPed.dbgrdDetPed.Columns(3).Visible = True
          frmCaptPed.dbgrdDetPed.Columns(4).Visible = False
          frmCaptPed.dbgrdDetPed.Columns(5).Visible = True
       End If
       frmCaptPed.dbgrdDetPed.Refresh
       frmCaptPed.cmdReporte.Enabled = True
       frmCaptPed.CmdRefresh.Enabled = True
     Else
         If Trim(Mid(cSucursal, 1, 3)) <> 3 Then
            cmdDespla.Enabled = True
            CmdRefresh.Enabled = True
            cmdDespla_Click
         End If
         frmCaptPed.cmdGrabar.Enabled = False
     End If
     If Trim(txtcampos(1).Text) = "ABA" Or chksugerido.Value = 1 Then
       cmdDespla_Click
     End If
  ElseIf cModo = "CAPTURARPEDIDO" Then     'Modificaciones
        'cmdDespla.Enabled = Not (lProvAbto = Ins) And nOp <> 1
        cmdDespla.Enabled = True
        CmdRefresh.Enabled = True
        cmdReporte.Enabled = nOp <> 1
        If Not lProvAbto = Ins Then
           frmCaptPed.dbgrdDetPed.Visible = True
           frmCaptPed.dbgrdDetPed.Refresh
           frmCaptPed.dbgrdDetPed.Columns(0).Button = True  'No se muestra en altas
        End If
        AdoPedidos.Recordset.Update
  ElseIf cModo = "CONFIRMARPEDIDO" Then
      If chkFields(5).Value = 0 Then
         MsgBox "No puede grabar si no activa la casilla de pedido confirmado", vbExclamation
         chkFields(5).SetFocus
         Exit Sub
      End If
      AdoPedidos.Recordset.Update
      Unload Me

  ElseIf cModo = "RECIBIRPEDIDO" Then
     If chkFields(6).Value = 0 Then
        MsgBox "No puede grabar si no activa la casilla de Pedido recibido", vbExclamation
        chkFields(6).SetFocus
        Exit Sub
     End If
     
  End If
  
  cmdExporta.Visible = IIf(Trim(txtcampos(1).Text = "ABA"), True, False)
  poninfo  'Pone leyenda con numero de productos, cajas y piezas solicitados en el pedido
  'Si se recibe se afecta inventario
  If Me.chkFields(6).Value = 1 Then
     If chksugerido.Value = 1 Then
        MsgBox "NO SE PUEDEN RECIBIR PEDIDOS SUGERIDOS, DESACTIVE LA MARCA DE SUGERIDO", vbExclamation
        Exit Sub
     End If
     If MsgBox("REALMENTE DESEAS RECIBIR EL PEDIDO E INCREMENTAR EL INVENTARIO", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
     End If
     If AdoDetPed.Recordset.BOF And AdoDetPed.Recordset.EOF = True Then
        MsgBox "NO PUEDE RECIBIR EL PEDIDO SI AUN NO CAPTURA SU DETALLE", vbCritical
        Exit Sub
     End If
     cn.BeginTrans: lTrans = True
     StbMensajes.SimpleText = Space(90) + "Espere un momento, aumentando inventario.."
     StbMensajes.Refresh
     'Afecto Inventario
     Set rstInvent = New ADODB.Recordset
     AdoDetPed.Recordset.MoveFirst
     While Not AdoDetPed.Recordset.EOF
        rstInvent.Open "SELECT * FROM inventario WHERE inprod = '" & AdoDetPed.Recordset!df_prod & "'", cn, adOpenDynamic, adLockOptimistic
        If rstInvent.BOF And rstInvent.EOF Then
           'MsgBox "NO EXISTE EN EL INVENTARIO EL ARTICULO " & Chr(13) & _
           AdoDetPed.Recordset!df_prod & "  " & rsttemp!DESCRIPC & Chr(13) & CStr(rsttemp!PAQUETES) & " X " & CStr(rsttemp!Contenid) & " " & rsttemp!Medida & Chr(13) & _
           "A CONTINUACION SE DARA DE ALTA EN EL INVENTARIO", vbInformation
           CADENA = "INSERT INTO inventario (inprod) VALUES(" & AdoDetPed.Recordset!df_prod & ")"
           cn.Execute CADENA
           nCanAnt = 0
           rstInvent!InCant = AdoDetPed.Recordset!df_cantreal
           rstInvent!InCantPza = AdoDetPed.Recordset!df_cantrealP
        Else
           'Grabo la existencia antes de Grabar
           CAD = "UPDATE detallefactura SET df_existencia = " & rstInvent!InCant & " WHERE df_prod = '" & AdoDetPed.Recordset!df_prod & "' AND df_pedido = '" & txtcampos(0).Text & "'"
           cn.Execute CAD
           nCanAnt = rstInvent!InCant
           rstInvent!InCant = rstInvent!InCant + AdoDetPed.Recordset!df_cantreal
           rstInvent!InCantPza = rstInvent!InCantPza + AdoDetPed.Recordset!df_cantrealP
           ''& ", IncantPza = IncantPza + " & AdoDetPed.Recordset!DF_CANTREalp & " WHERE Inprod = " & AdoDetPed.Recordset!df_prod
        End If
        rstInvent.Update
        rstInvent.Close
        AdoDetPed.Recordset.MoveNext
     Wend
     cn.CommitTrans
     Call frmPedProvB.listaexistencias(Trim(txtcampos(1).Text))
     RESP = MsgBox("SE ha incrementado Correctamente el Inventario ?", vbYesNo, "INVENTARIO")
     If RESP = vbYes Then
            'Agrego Nota de entrada
            StbMensajes.SimpleText = Space(80) + "Espere un momento, generando nota de entrada.."
            StbMensajes.Refresh
            FolNot = "N" + Trim(txtcampos(0).Text)
            'Obtengo importe solicitado Y recibido
            Set rsttemp = New ADODB.Recordset
            rsttemp.Open "SELECT SUM(DetalleFactura.df_cantsol * Tfproduc.Precosto) AS ImptSol, SUM(DetalleFactura.df_cantreal * Tfproduc.Precosto) AS ImptRec FROM DetalleFactura,Tfproduc WHERE DetalleFactura.df_prod = Tfproduc.Consec AND DetalleFactura.df_pedido ='" & txtcampos(0).Text & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
            ntot = 0
            ntot = ntot + IIf(IsNull(AdoFacturas.Recordset!Impfac1), 0, AdoFacturas.Recordset!Impfac1)
            ntot = ntot + IIf(IsNull(AdoFacturas.Recordset!Impfac2), 0, AdoFacturas.Recordset!Impfac2)
            ntot = ntot + IIf(IsNull(AdoFacturas.Recordset!Impfac3), 0, AdoFacturas.Recordset!Impfac3)
            ntot = ntot + IIf(IsNull(AdoFacturas.Recordset!Impfac4), 0, AdoFacturas.Recordset!Impfac4)
            ntot = ntot + IIf(IsNull(AdoFacturas.Recordset!Impfac5), 0, AdoFacturas.Recordset!Impfac5)
            ntot = ntot + IIf(IsNull(AdoFacturas.Recordset!Impfac6), 0, AdoFacturas.Recordset!Impfac6)
            ntot = ntot + IIf(IsNull(AdoFacturas.Recordset!Impfac7), 0, AdoFacturas.Recordset!Impfac7)
            ntot = ntot + IIf(IsNull(AdoFacturas.Recordset!Impfac8), 0, AdoFacturas.Recordset!Impfac8)
            ntot = ntot + IIf(IsNull(AdoFacturas.Recordset!Impfac9), 0, AdoFacturas.Recordset!Impfac9)
            ntot = ntot + IIf(IsNull(AdoFacturas.Recordset!Impfac10), 0, AdoFacturas.Recordset!Impfac10)
     
            StbMensajes.SimpleText = Space(90) + "Espere un momento, grabando facturas y nota de entrada"
            StbMensajes.Refresh
            If IsNull(rsttemp!ImptSol) Then
                ImptSol = 0
            Else
                ImptSol = rsttemp!ImptSol
            End If
     
             If IsNull(rsttemp!ImptRec) Then
                ImptRec = 0
            Else
                ImptRec = rsttemp!ImptRec
            End If
     
            CAD = "UPDATE [NotaEntrada] SET ImporteSol = " & ImptSol & ", ImporteRec = " & ImptRec & ", DifeProd = " & ImptSol - ImptRec & ", DifPrecio = " & ntot - ImptRec & " WHERE pedido = '" & Trim(txtcampos(0).Text) & "'"
             cn.Execute CAD
            cn.Execute "INSERT INTO [DetalleNota](ClaveNota,Producto,cantsol,cantsolp,cantrec,cantrecp,costo) SELECT CveNota = '" & FolNot & "', df_prod, df_cantsol, df_cantsolp,df_cantreal, df_cantrealP,df_costo FROM [DetalleFactura] WHERE df_pedido = '" & txtcampos(0).Text & "'"
     
             txtcampos(12).Visible = True
            txtcampos(12).Text = FolNot
            AdoPedidos.Recordset!p_NotaEntrada = FolNot
     
             AdoPedidos.Recordset!p_Pedido = txtcampos(0).Text
            AdoPedidos.Recordset!p_proveedor = txtcampos(1).Text
            AdoPedidos.Recordset!p_sucursal = txtcampos(2).Text
            AdoPedidos.Recordset!p_fecent = IIf(Trim(txtcampos(5).Text) = "", Null, txtcampos(5).Text)
            AdoPedidos.Recordset!p_solicita = txtcampos(3).Text
            'AdoPedidos.Recordset!p_observaciones = IIf(Trim(txtCampos(6).Text) = "", txtCampos(5).Text)
            AdoPedidos.Recordset!p_fecped = txtcampos(4).Text
            AdoPedidos.Recordset!p_fecentreal = txtcampos(8).Text
            AdoPedidos.Recordset!p_situacion = 1
            AdoPedidos.Recordset!P_recibido = 1
            MsgBox "EL FOLIO DE LA NOTA DE ENTRADA GENERADA ES: " + FolNot, vbInformation
            AdoPedidos.Recordset.Update
            cmdGrabar.Enabled = False
            'cn.CommitTrans
            'MsgBox cSucursal
            If tipotienda <> 2 Then
                    RESP = MsgBox("Desea Enviar todo el Producto Recibido a Piso ? ...", vbYesNo, "DISMINUIR INVENTARIO")
                    If RESP = vbYes Then
                            Call generasalida
                    End If
            End If
       Else
           MsgBox "Vuelva a Ejecutar el proceso para Registrar correctamente el inventario", vbInformation
           Exit Sub
       End If 'DE SI GRABO CORRECTAMENTE EL INVENTARIO
           
   End If
   StbMensajes.SimpleText = cMensaje
   StbMensajes.Refresh

  Exit Sub
Error:
   MsgBox "OCURRIO EL SIGUIENTE ERROR: " + Chr(13) + UCase(Err.Description), vbCritical
   If lTrans Then
      MsgBox "A CONTINUACION SE DESHARAN LAS MODIFICACIONES REALIZADAS AL INVENTARIO", vbCritical
      cn.RollbackTrans
      AdoPedidos.Recordset!P_recibido = 0
      AdoPedidos.Recordset.Update
   End If
End Sub

Private Sub generasalida()
'1 se genera un envio a piso en forma automatica
'2 se dismunuye el inventario de bodega
'3 se aumenta inventario de piso
Set rsttemp = New ADODB.Recordset
rsttemp.ActiveConnection = cCadConex
rsttemp.CursorType = adOpenKeyset
csucu = Mid(cSucursal, 6, 2)
rsttemp.Source = "SELECT MAX (CAST(SUBSTRING(t_clave,4,10) AS INT)) As FolTra FROM [Traslados] WHERE SUBSTRING(t_clave,1,3) = 'T" & Trim(csucu) & "'"
rsttemp.Open
If IsNull(rsttemp!FolTra) Then
    Folio = "T" + Trim(csucu) + "1"
Else
    Folio = "T" + Trim(csucu) + Trim(Str(rsttemp!FolTra + 1))
End If
sucu = Mid(cSucursal, 1, 3)
Set rstemp1 = New ADODB.Recordset
rstemp1.Open "select max(t_foliotie) as folio from traslados where t_sucursalreceptor =  '" & Trim(sucu) & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
nfoltie = 0
If rstemp1.EOF Then
   nfoltie = 1
Else
   nfoltie = rstemp1!Folio + 1
End If
If IsNull(nfoltie) Then
   nfoltie = 0
End If
tipo = 0 ' es traslado abierto
sucu = Trim(Mid(cSucursal, 1, 3))
entrada = 0 ' se trata de una salida
CAD = "insert into traslados(t_recibido,t_clave,t_fecha,t_tipo,t_perenv,t_perrec,t_costo,t_sucursalemisor,t_sucursalreceptor,t_enviado,t_entrada,t_foliotie, t_perfle)  values (" & _
      "0,'" & Folio & "','" & date + Time & "'," & tipo & ",'" & Trim(Mid(cCveDesUsu, 1, 3)) & "','" & Trim(Mid(cCveDesUsu, 1, 3)) & "'," & 0 & "," & sucu & "," & sucu & ",1," & entrada & "," & nfoltie & ",1)"
cn.Execute CAD
AdoDetPed.Recordset.MoveFirst
While Not Me.AdoDetPed.Recordset.EOF
   costop = 0
   costop = AdoDetPed.Recordset!df_Costo / paq
   If IsNull(costop) Then
      costop = 0
   End If
   importe = (AdoDetPed.Recordset!df_Costo * AdoDetPed.Recordset!df_cantreal) + (AdoDetPed.Recordset!df_cantrealP * costop)
   CAD = "insert into detalletraslado(dt_clave,dt_producto,dt_cantidad,dt_cantidadp,dt_pedido, dt_costo,dt_costop, dt_importe) values(" & _
   "'" & Folio & "', '" & Trim(AdoDetPed.Recordset!df_prod) & "'," & AdoDetPed.Recordset!df_cantreal & "," & AdoDetPed.Recordset!df_cantrealP & ",'" & Trim(AdoDetPed.Recordset!df_pedido) & "'," & AdoDetPed.Recordset!df_Costo & "," & costop & "," & importe & ")"
   cn.Execute CAD
   CAD = "update inventario set incant = incant - " & AdoDetPed.Recordset!df_cantreal & " , INCANTPZA = INCANTPZA - " & AdoDetPed.Recordset!df_cantrealP & " where inprod =  '" & Trim(AdoDetPed.Recordset!df_prod) & "'"
   cn.Execute CAD 'DISMINUCION DEL INVENTARIO
   CAD = "update inventariopiso set incant = incant + " & AdoDetPed.Recordset!df_cantreal & " , INCANTPZA = INCANTPZA + " & AdoDetPed.Recordset!df_cantrealP & " where inprod =  '" & Trim(AdoDetPed.Recordset!df_prod) & "'"
   cn.Execute CAD 'AUMENTA EL INVENTARIO DE PISO
   AdoDetPed.Recordset.MoveNext
Wend
MsgBox "Se ha generado la Salida con clave " & Folio & " y se ha diminuido el inventario automaticamente ", vbInformation, "INFORMACION"
End Sub

Private Sub cmdInven_Click()
frmPedProvB.listaexistencias (Trim(txtcampos(1).Text))
End Sub

Private Sub CmdRefresh_Click()
AdoDetPed.Refresh
poninfo
End Sub

Private Sub cmdReporte_Click()
Dim cMensaje As String
On Error GoTo Error:
cMensaje = StbMensajes.SimpleText
StbMensajes.SimpleText = Space(90) + "Espere un momento generando reporte..."
StbMensajes.Refresh
Select Case cModo
Case "CAPTURARPEDIDO"
   crpt.ReportFileName = App.Path & "\PeCapCon.rpt"
   crpt.WindowTitle = "Pedido sugerido numero " & txtcampos(0).Text
   crpt.Formulas(0) = "FORMSELEC = {PEDIDOS.p_pedido} = '" & Trim(txtcampos(0).Text) & "'"
   crpt.Formulas(1) = "PEDIDO = 'PEDIDO " & IIf(Trim(txtcampos(1).Text) = "ABA", " DE ABASTECIMIENTO", " SUGERIDO ") & " NUMERO [ " & txtcampos(0).Text & " ]'"
   'cRpt.SQLQuery = "SELECT PEDIDOS.p_pedido, PEDIDOS.p_proveedor, PEDIDOS.p_fecped, PEDIDOS.p_sucursal, PEDIDOS.p_fecentreal, PEDIDOS.p_fecconfirma, " & _
   '                         " DETALLEFACTURA.df_prod, DETALLEFACTURA.df_cantidad, DETALLEFACTURA.df_cantsol, DETALLEFACTURA.df_cantsolp, " & _
   '                         " CATPROV.NOMPROVE, CATTIENDA.tidescrip, TFPRODUC.DESCRIPC, TFPRODUC.CONTENID, TFPRODUC.MEDIDA, TFPRODUC.PAQUETES " & Chr(13) & _
   '                 "FROM pitico.dbo.PEDIDOS PEDIDOS, " & _
   '                      "pitico.dbo.DETALLEFACTURA DETALLEFACTURA," & _
   '                      "pitico.dbo.CATPROV CATPROV," & _
   '                      "pitico.dbo.CATTIENDA CATTIENDA," & _
   '                      "pitico.dbo.TFPRODUC TFPRODUC " & Chr(13) & _
   '                 "WHERE PEDIDOS.p_pedido = DETALLEFACTURA.df_pedido AND " & _
   '                      "PEDIDOS.p_proveedor = CATPROV.PROVE AND " & _
   '                      "PEDIDOS.p_sucursal = CATTIENDA.ticlave AND " & _
   '                      "DETALLEFACTURA.df_prod = TFPRODUC.CONSEC  AND " & _
   '                      "DETALLEFACTURA.df_sugerido = " & IIf(Trim(txtCampos(1).Text) = "ABA", 0, 1) & " AND " & _
   '                      "PEDIDOS.p_pedido = '" & AdoPedidos.Recordset!p_pedido & "' " & Chr(13) & _
   '                 "ORDER BY " & _
   '                      "TFPRODUC.DESCRIPC ASC, TFPRODUC.CONTENID ASC"
Case "CONFIRMARPEDIDO"
   crpt.ReportFileName = App.Path & "\PeCapCon.rpt"
   crpt.WindowTitle = "Pedido Confirmado numero " & txtcampos(0).Text
   crpt.Formulas(0) = "FORMSELEC = {DETALLEFACTURA.df_pedido} = '" & txtcampos(0).Text & "' AND {DETALLEFACTURA.df_sugerido} = " & IIf(Trim(txtcampos(1).Text) = "ABA", 0, 1)
   crpt.Formulas(1) = "PEDIDO = 'PEDIDO CONFIRMADO NUMERO [ " & txtcampos(0).Text & " ]'"
Case "RECIBIRPEDIDO"
   crpt.ReportFileName = App.Path & "\PeNotEnt.rpt"
   crpt.WindowTitle = "Nota de entrada del pedido " & txtcampos(0).Text
   crpt.Formulas(0) = "FORMSELEC = {DETALLEFACTURA.df_pedido} = '" & txtcampos(0).Text & "' AND {DETALLEFACTURA.df_sugerido} = " & IIf(Trim(txtcampos(1).Text) = "ABA", 0, 1)
   crpt.Formulas(3) = "PEDIDO = 'NOTA DE ENTRADA DEL PEDIDO " & txtcampos(0).Text & "'"
   crpt.Formulas(4) = "PROVED = 'PROVEEDOR [ " & txtcampos(1).Text & "  " & cmbProved.Text & " ]'"
   crpt.Formulas(5) = "PARADONDE = 'PEDIDO PARA  [ " & txtcampos(2).Text & "  " & cmbSucursal.Text & " ]'"
   crpt.Formulas(6) = "FOLNOTENT = 'FOLIO " & Trim(txtcampos(12).Text) & "'"
End Select
'MsgBox cRpt.SQLQuery
crpt.Connect = cCadConex
crpt.Action = 1
StbMensajes.SimpleText = cMensaje
StbMensajes.Refresh
Exit Sub
Error:
  MsgBox Err.Description
  StbMensajes.SimpleText = cMensaje
End Sub

Private Sub Command1_Click()
crpt.WindowTitle = "Nota de entrada con costos del pedido " & txtcampos(0).Text
crpt.ReportFileName = App.Path & "\PEpedco.rpt"
crpt.Formulas(0) = "FORMSELEC = '" & txtcampos(0).Text & "'"
crpt.Action = 1
End Sub

Private Sub dbgrdDetPed_AfterColUpdate(ByVal ColIndex As Integer)
On Error GoTo Error:
AdoDetPed.Recordset!df_pedido = txtcampos(0).Text
If Trim(dbgrdDetPed.Columns(ColIndex).Text) = "" Or Not IsNumeric(dbgrdDetPed.Columns(ColIndex).Text) Then
   dbgrdDetPed.Columns(ColIndex).Text = 0
End If
If lProvAbto = Ins Or lProvAbto = Abi Then
  'Se actualizan campos que no se muestran en el dbGrid
  AdoDetPed.Recordset!df_cantidad = 0  'Cant. sugerida
  If ColIndex = 3 And Val(dbgrdDetPed.Columns(4)) > Val(dbgrdDetPed.Columns(ColIndex)) Then
     dbgrdDetPed.Columns(4).Text = dbgrdDetPed.Columns(ColIndex)
  End If
  If ColIndex = 5 And Val(dbgrdDetPed.Columns(6)) > Val(dbgrdDetPed.Columns(ColIndex)) Then
     dbgrdDetPed.Columns(6).Text = dbgrdDetPed.Columns(ColIndex)
  End If
End If
'Temporalmente se desactiva mientras se define lo de las promociones SAVE
'Se valida que no capturen mas de lo solicitado
If ColIndex = 4 Then     'Cantidad recibida por caja
   'If AdoDetPed.Recordset!df_cantreal > AdoDetPed.Recordset!df_cantsol Then
   '   MsgBox "LA CANTIDAD RECIBIDA EN CAJAS NO PUEDE SER MAYOR A LA SOLICITADA", vbExclamation
   '   AdoDetPed.Recordset!df_cantreal = AdoDetPed.Recordset!df_cantsol
   'End If
ElseIf ColIndex = 5 Then 'Cantidad solicitada por pieza
   If AdoDetPed.Recordset!df_cantsolp >= AdoCatPro.Recordset!PAQUETES Then
      MsgBox "LA CANTIDAD EN PIEZAS NO PUEDE SER MAYOR O IGUAL AL NUMERO DE PAQUETES POR CAJA, " & _
             "AUMENTE EL NUMERO DE CAJAS", vbExclamation
      AdoDetPed.Recordset!df_cantsolp = 0
      Exit Sub
   End If
ElseIf ColIndex = 6 Then 'Cantidad recibida por piezas
   'If AdoDetPed.Recordset!df_cantrealp > AdoDetPed.Recordset!df_cantsolp Then
   '   MsgBox "LA CANTIDAD RECIBIDA EN PIEZAS NO PUEDE SER MAYOR A LA SOLICITADA", vbExclamation
   '   AdoDetPed.Recordset!df_cantrealp = AdoDetPed.Recordset!df_cantsolp
   '   Exit Sub
   'End If
End If
'SendKeys "{DOWN}"
Exit Sub
Error:
  MsgBox Err.Description
End Sub

Private Sub dbgrdDetPed_AfterUpdate()
  poninfo   'Pone leyenda num. de productos,cajas y piezas
End Sub

Private Sub dbgrdDetPed_GotFocus()
  Cal1.Visible = False
End Sub

Private Sub dbgrdDetPed_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   'SendKeys "{TAB}"
   KeyAscii = 0
   keybd_event &H9, 0, 0, 0:   keybd_event &H9, 0, &H2, 0
End If
End Sub

Private Sub dbgrdDetPed_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If Not (AdoCatPro.Recordset.EOF And AdoCatPro.Recordset.EOF) And dbgrdDetPed.VisibleRows <> 1 Then
   On Error GoTo Error:
   If dbgrdDetPed.Columns(ColIndex_lstProd).Text <> "" Then
     AdoCatPro.Recordset.MoveFirst
     AdoCatPro.Recordset.Find "Consec = '" & Trim(dbgrdDetPed.Columns(ColIndex_lstProd)) & "'"
     lblDescrip.Caption = AdoCatPro.Recordset!descripc + Chr(13) + " (" + CStr(AdoCatPro.Recordset!PAQUETES) + " X " + CStr(AdoCatPro.Recordset!CONTENID) + " " + AdoCatPro.Recordset!medida + " ) " + Chr(13) & "PROM: " & AdoCatPro.Recordset!cajas & "/" & AdoCatPro.Recordset!encajas
     paq = AdoCatPro.Recordset!PAQUETES
   End If
  Else
     lblDescrip.Caption = ""
  End If
  Exit Sub
Error:
   'MsgBox Err.Description
End Sub

Private Sub dbgrdfactu_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   'SendKeys vbTab
   keybd_event &H9, 0, 0, 0:   keybd_event &H9, 0, &H2, 0
End If
End Sub

Private Sub Form_Load()
Dim RST As ADODB.Recordset
Set RST = New ADODB.Recordset
  fraGenerales.Visible = False
  cmdGrabar.Visible = False
  cmdCancelar.Visible = False
  
  AdoPedidos.ConnectionString = cCadConex
  AdoPedidos.CommandType = adCmdText
  AdoPedidos.RecordSource = "SELECT * FROM [Pedidos]"
  AdoPedidos.Refresh
 RST.Open "SELECT * FROM [Cattienda]", cn, adOpenForwardOnly, adLockOptimistic, adCmdText
 cmbSucursal.Clear
 Do While Not RST.EOF
     If Not IsNull(RST!tidescrip) Then
            cmbSucursal.AddItem RST!tidescrip
     End If
 RST.MoveNext
 Loop
RST.Close

End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Forma = O Then
    frmpedAba.Show
  ElseIf Forma = 1 Then
    frmpedidos.Show
  Else
    frmpedBod.Show
  End If
End Sub

Private Sub mskMonFac_GotFocus()
  Cal1.Visible = False
End Sub

Private Sub txtCampos_GotFocus(Index As Integer)
On Error GoTo Error:
'Campos de tipo fecha en los que se muestra en control calendario
Select Case Index
   Case 0  'Clave del folio del pedido
        If Forma = 0 Then
           frmpedAba.Hide
        Else
           frmpedidos.Hide
        End If
        Me.dbgrdfactu.Visible = False
        If nOp = 1 Then
           'Obtengo el numero de folio consecutivo
           Set rsttemp = New ADODB.Recordset
           rsttemp.ActiveConnection = cn
           rsttemp.CursorType = adOpenKeyset
           rsttemp.Source = "SELECT MAX (CAST(SUBSTRING(P_PEDIDO,4,7) AS INT)) As FolMay FROM [PEDIDOS] WHERE SUBSTRING(p_pedido,1,3) = '" & Trim(Mid(cSucursal, 3, 6)) & "'"
           rsttemp.Open
           
           AdoPedidos.Recordset.AddNew
           If IsNull(rsttemp!FolMay) Then
              txtcampos(0).Text = Mid(Trim(Mid(cSucursal, 3)), 1, 3) + "1"
           Else
              txtcampos(0).Text = Mid(Trim(Mid(cSucursal, 3)), 1, 3) + Trim(Str(rsttemp!FolMay + 1))
           End If
           rsttemp.Close
           Set rsttemp = Nothing
           SendKeys "{TAB}"
           'keybd_event &H9, 0, 0, 0: keybd_event &H9, 0, &H2, 0
           If Forma = 0 Then  'Si se llama de abastecimiento asigno el proveedor
              chksugerido.Value = 0: chksugerido.Visible = False
              SendKeys "ABA"
              'SendKeys "{TAB}"
              keybd_event &H9, 0, 0, 0: keybd_event &H9, 0, &H2, 0
           End If
        Else
           txtcampos(0).SetFocus
        End If
   Case 1
        'En caso de modificaciones no se permite cambiar de proveedor
        If nOp <> 1 Then
           'SendKeys "{TAB}"
           keybd_event &H9, 0, 0, 0
           txtcampos(Index).Enabled = False
           cmbProved.Enabled = False
           chksugerido.Visible = (Forma = 1)
        End If
   Case Else
        Cal1.Visible = False
End Select
Exit Sub
Error:
  MsgBox Err.Description
End Sub

Private Sub txtCampos_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case KeyAscii
     Case 27
         Unload Me
     Case 13
         KeyAscii = 0
         'SendKeys "{Tab}"
         keybd_event &H9, 0, 0, 0
End Select
End Sub

Private Sub txtCampos_LostFocus(Index As Integer)
Dim N As Integer
Dim lvalor As Boolean
'On Error GoTo Error:

Select Case Index
    Case 0   'Clave del pedido
        txtcampos(Index).Text = Trim(UCase(txtcampos(Index).Text))
        txtcampos(Index).Refresh
        If LTrim(RTrim(txtcampos(0).Text)) = "" Or IsNull(txtcampos(0).Text) Then
            MsgBox "No puede dejar en blanco la clave del pedido", vbCritical
            txtcampos(0).SetFocus
            Exit Sub
        End If
        'Realizó la busqueda para determinar si es Alta o Modificación
        If nOp <> 1 Then  ' Modificaciones
           If Not (AdoPedidos.Recordset.BOF And AdoPedidos.Recordset.EOF) Then AdoPedidos.Recordset.MoveFirst
           AdoPedidos.Recordset.Find "P_pedido = '" & txtcampos(0).Text & "'"
           If AdoPedidos.Recordset.EOF = True Then
              RESP = MsgBox("No existe el pedido con la clave especificada", vbExclamation)
              txtcampos(0).SetFocus
              Exit Sub
           End If
        End If
           
        Me.dbgrdfactu.Visible = False
        pConfirmado = AdoPedidos.Recordset!p_situacion
        PedRecibido = AdoPedidos.Recordset!P_recibido
        'Asigno Captions al formulario
        If cModo = "CAPTURARPEDIDO" Then
           cTipPed = IIf(Forma = 0, "para abastecimiento", "sugerido")
           If nOp = 1 Then
              cEnca = " STATUS: [Captura de pedido " & cTipPed & "]              MOVIMIENTO: [Alta]"
           Else
              cEnca = " STATUS: [Captura de pedido " & cTipPed & "]           MOVIMIENTO: [Modificación]"
           End If
           'If pConfirmado And lProvAbto = Ind Or (PedRecibido And lProvAbto = Ins) Then
           If PedRecibido Then
              'Resp = MsgBox("No es posible modificar el pedido" & _
              '" porque ya ha sido confirmado" & Chr(13) & "Deseas ver el pedido confirmado?", vbExclamation + vbYesNo)
              RESP = vbYes
              If RESP = vbNo Then
                 'txtCampos(0).SetFocus
                 Unload Me
                 Exit Sub
              End If
              'cModo = "CONFIRMARPEDIDO"
              cModo = "RECIBIRPEDIDO"
              cEnca = " STATUS: [Confirmar pedido]"
              'SendKeys "{Enter}"
           End If
        ElseIf cModo = "CONFIRMARPEDIDO" Then
           cEnca = " STATUS: [Confirmar pedido]"
           'SendKeys "{Enter}": SendKeys "{Enter}": SendKeys "{Enter}": SendKeys "{Enter}": SendKeys "{Enter}"
           keybd_event &H9, 0, 0, 0:   keybd_event &H9, 0, &H2, 0: keybd_event &H9, 0, 0, 0:  keybd_event &H9, 0, &H2, 0: keybd_event &H9, 0, 0, 0:  keybd_event &H9, 0, &H2, 0: keybd_event &H9, 0, 0, 0:  keybd_event &H9, 0, &H2, 0: keybd_event &H9, 0, 0, 0:  keybd_event &H9, 0, &H2, 0
        ElseIf cModo = "RECIBIRPEDIDO" Then
           If Not pConfirmado Then
              MsgBox "No puede recibir el pedido si aun no ha sido confirmado", vbExclamation
              'txtCampos(0).SetFocus
              Unload Me
              Exit Sub
           End If
           cEnca = " STATUS: [Recibir pedido]"
           'SendKeys "{Enter}": SendKeys "{Enter}"
        End If
        frmCaptPed.Caption = cEnca
        
        'Cargo la factura
        AdoFacturas.ConnectionString = cCadConex
        AdoFacturas.CommandType = adCmdText
        AdoFacturas.RecordSource = "SELECT * FROM [NOTAENTRADA] WHERE [pedido] = '" & txtcampos(0).Text & "'"
        AdoFacturas.Refresh
        'MsgBox "factura antes " + AdoFacturas.Recordset!factura1
        If AdoFacturas.Recordset.BOF And AdoFacturas.Recordset.EOF = True Then
           AdoFacturas.Recordset.AddNew 'Para que no se borren los datos al ecribir en los campos del control adofacturas
           AdoFacturas.Recordset!Pedido = txtcampos(0).Text
           AdoFacturas.Recordset!Clavenota = "N" & Trim(txtcampos(0).Text)
           AdoFacturas.Recordset.Update
           AdoFacturas.Refresh
        End If
        'MsgBox "factura despues " + AdoFacturas.Recordset!factura1
        txtcampos(0).Enabled = False
        fraGenerales.Visible = True
        txtcampos(1).Visible = True
        cmbProved.Visible = True
        txtcampos(1).SetFocus
        
        'LLeno el combo de proveedores
        AdoCatProved.ConnectionString = cCadConex
        AdoCatProved.CommandType = adCmdText
        If Forma = 0 Then
           AdoCatProved.RecordSource = "SELECT * FROM CatProv WHERE prove = 'ABA'"
        Else
           AdoCatProved.RecordSource = "SELECT * FROM CatProv WHERE prove <> 'ABA'"
        End If
        AdoCatProved.Refresh
        cmbProved.Clear
        Do While Not AdoCatProved.Recordset.EOF
            If Not IsNull(AdoCatProved.Recordset!NOMPROVE) Then
               cmbProved.AddItem AdoCatProved.Recordset!NOMPROVE
            End If
            AdoCatProved.Recordset.MoveNext
        Loop
        
    Case 1  'Clave del proveedor
         txtcampos(Index).Text = Trim(UCase(txtcampos(Index).Text))
         txtcampos(Index).Refresh
         If txtcampos(Index).Text = "" Or IsNull(txtcampos(Index).Text) Then
            cmbProved.SetFocus
            Exit Sub
         Else
            AdoCatProved.Recordset.MoveFirst
            AdoCatProved.Recordset.Find "Prove= '" & Trim(txtcampos(1).Text) & "'"
            If AdoCatProved.Recordset.EOF = True Then
                'MsgBox "No existe la clave del proveedor especificado", vbExclamation
                cmbProved.SetFocus
                Exit Sub
            End If
         End If
         cmbProved.Text = AdoCatProved.Recordset.Fields!NOMPROVE
         nFrecuencia = AdoCatProved.Recordset!frecuencia
         'LLeno el Catálogo de productos exclusivamente con los productos del proveedor especificado
         AdoCatPro.ConnectionString = cCadConex
         AdoCatPro.CommandType = adCmdText
         'En caso de que sea pedido para abastecimiento por tienda es de varios proveedores
         'Por lo tanto cargo todo el catalogo de productos
         If txtcampos(1).Text = "ABA" Then
            AdoCatPro.RecordSource = "SELECT * FROM TfProduc,Inventario WHERE consec = inprod AND (InCantcdc > 0 OR InCantPzacdc > 0)"
            'AdoCatPro.RecordSource = "SELECT * FROM TfProduc,Inventario WHERE consec = 5"
         Else
            AdoCatPro.RecordSource = "SELECT * FROM [TfProduc] WHERE [ClaProve] = '" & txtcampos(1).Text & "' AND activo = 1"
         End If
         AdoCatPro.Refresh
         fraAvance.Visible = True
         fraAvance.Refresh
         PGB.Min = 0: PGB.Max = IIf(AdoCatPro.Recordset.RecordCount = 0, 1, AdoCatPro.Recordset.RecordCount): nreg = 0
         Lstprod.Clear
         Do While Not AdoCatPro.Recordset.EOF
            nreg = nreg + 1
            If (Not IsNull(AdoCatPro.Recordset!descripc) And Not IsNull(AdoCatPro.Recordset!CONTENID) And Not IsNull(AdoCatPro.Recordset!PAQUETES) And Not IsNull(AdoCatPro.Recordset!medida)) Then
               Lstprod.AddItem AdoCatPro.Recordset!descripc + " (" + Str(AdoCatPro.Recordset!PAQUETES) + " X " + Str(AdoCatPro.Recordset!CONTENID) + " " + AdoCatPro.Recordset!medida + " ) " _
               + " [" + Str(AdoCatPro.Recordset!CONSEC) + "]"
            End If
            PGB.Value = nreg
            AdoCatPro.Recordset.MoveNext
         Loop
         fraAvance.Visible = False
         
         If AdoCatPro.Recordset.BOF And AdoCatPro.Recordset.EOF And txtcampos(1).Text <> "ABA" Then
            MsgBox "EN EL CATALOGO DE ARTICULOS NO EXISTEN PRODUCTOS" & Chr(13) & _
            "DE EL PROVEEDOR ESPECIFICADO, POR LO TANTO NO PUEDES" & Chr(13) & _
            "CONTINUAR CON LA CAPTURA DEL PEDIDO", vbCritical
            Unload Me
            Exit Sub
         End If
         If Not (AdoCatPro.Recordset.BOF And AdoCatPro.Recordset.EOF) Then AdoCatPro.Recordset.MoveFirst
         
         lProvAbto = IIf(IsNull(AdoCatProved.Recordset!tipo), "", AdoCatProved.Recordset!tipo)
         If lProvAbto = Ins Then frmCaptPed.Caption = cEnca + Space(25) + "TIPO DE PROVEEDOR: [Instantáneo ]"
         If lProvAbto = Abi Then frmCaptPed.Caption = cEnca + Space(25) + "TIPO DE PROVEEDOR: [Abierto ]"
         If lProvAbto = Ind Then frmCaptPed.Caption = cEnca + Space(25) + "TIPO DE PROVEEDOR: [Indirecto]"
         
         txtcampos(2).Text = Trim(Mid(cSucursal, 1, 3))
         cmbSucursal.Text = Trim(Mid(cSucursal, 4))
         
         lbletiquetas(2).Visible = True
         txtcampos(2).Visible = True
         Me.cmbSucursal.Visible = True
         If txtcampos(2).Enabled = True Then txtcampos(2).SetFocus
         'SendKeys "{Tab}"
         keybd_event &H9, 0, 0, 0:   keybd_event &H9, 0, &H2, 0
         
    Case 2  'Clave de la sucursal
         txtcampos(Index).Text = Trim(UCase(txtcampos(Index).Text))
         txtcampos(Index).Refresh
         
         'txtcampos(Index).Enabled = False
         'cmbsucursal.Enabled = False
         
         'LLeno el Combo de Catalogo de usuarios
         AdoCatUsu.ConnectionString = cCadConex
         AdoCatUsu.CommandType = adCmdText
         AdoCatUsu.RecordSource = "SELECT * FROM [Usuarios] WHERE [Sucursal] = '" & txtcampos(2).Text & "'"
         AdoCatUsu.Refresh
         cmbUsuarios.Clear
         Do While Not AdoCatUsu.Recordset.EOF
             If Not IsNull(AdoCatUsu.Recordset!Name) Then
                 cmbUsuarios.AddItem AdoCatUsu.Recordset!Name
                 If Not PedRecibido Then cmbJefeRec.AddItem AdoCatUsu.Recordset!Name
             End If
             AdoCatUsu.Recordset.MoveNext
         Loop
         
         lbletiquetas(3).Visible = True
         txtcampos(3).Visible = True
         Me.cmbUsuarios.Visible = True
         txtcampos(3).Text = Mid(cCveDesUsu, 1, 3)
         txtcampos(3).SetFocus
         'SendKeys "{TAB}"
         keybd_event &H9, 0, 0, 0:   keybd_event &H9, 0, &H2, 0
    
    Case 3 '  Clave del Jefe del departamento
         txtcampos(Index).Text = Trim(UCase(txtcampos(Index).Text))
         txtcampos(Index).Refresh
         AdoCatUsu.Recordset.MoveFirst
         AdoCatUsu.Recordset.Find "Clave = '" & Trim(txtcampos(3).Text) & "'"
         If AdoCatUsu.Recordset.EOF = True Then
            'MsgBox "No existe la clave del usuario", vbExclamation
            Me.cmbUsuarios.SetFocus
            Exit Sub
         End If
         cmbUsuarios.Text = AdoCatUsu.Recordset!Name
         txtcampos(Index).Enabled = False
         cmbUsuarios.Enabled = False
         
         For N = 4 To 6
             lbletiquetas(N).Visible = True
             txtcampos(N).Visible = True
         Next
         
         dbgrdDetPed.Visible = nOp <> 1 'No se muestra en altas
         dbgrdDetPed.Columns(0).Button = nOp <> 1
         cmdReporte.Enabled = True
         CmdRefresh.Visible = True
         cmdDespla.Visible = True
         cmdGrabar.Visible = True
         cmdCancelar.Visible = True

         If cModo = "CAPTURARPEDIDO" Then
            cmdReporte.Caption = "&Ped. sug."
            cmdReporte.Enabled = nOp <> 1
            cmdDespla.Enabled = nOp <> 1
            CmdRefresh.Enabled = nOp <> 1
            For N = 5 To 5
                chkFields(N).Visible = True
                chkFields(N).Enabled = (lProvAbto = Ins Or lProvAbto = Abi)
            Next
            'Si es Prov. instantaneo O Abierto y altas
            If (lProvAbto = Ins Or lProvAbto = Abi) And nOp = 1 Then
               If lProvAbto = Ins Then
                  lbletiquetas(9).Visible = True  'Jefe de recibo
                  txtcampos(9).Visible = True
                  cmbJefeRec.Visible = True
               End If
            'Si es Proveedor Abierto o Instantaneo y Modificaciones
            ElseIf (lProvAbto = Ins Or lProvAbto = Abi) And chksugerido.Value = 0 Then
                chkFields(5).Visible = True
                chkFields(6).Visible = (chksugerido.Value = 0)
                If pConfirmado Then
                   chkFields(5).Value = 1
                   chkFields(5).Enabled = False
                End If
                If PedRecibido Then
                   chkFields(6).Value = 1
                   chkFields(6).Enabled = False
                End If
                'Muestro cantida solicitada y recibida
                frmCaptPed.dbgrdDetPed.Columns(0).Button = True
                frmCaptPed.dbgrdDetPed.Columns(2).Visible = False
                frmCaptPed.dbgrdDetPed.Columns(4).Visible = True
                frmCaptPed.dbgrdDetPed.Columns(6).Visible = True
                
            End If
            dbgrdDetPed.Visible = (lProvAbto = Ins Or lProvAbto = Abi) And nOp <> 1
            dbgrdDetPed.Columns(0).Button = True And nOp <> 1
            
         ElseIf cModo = "CONFIRMARPEDIDO" Then
            cmbProved.Enabled = False
            cmbSucursal.Enabled = False
            cmbUsuarios.Enabled = False
            cmdDespla.Enabled = False
            cmdReporte.Enabled = True
            CmdRefresh.Visible = True

            cmdReporte.Caption = "&Ped. conf."

            If pConfirmado Then
               txtcampos(7).Visible = True
               lbletiquetas(7).Visible = True
               'quien sabe porque no ponia la fecha
               txtcampos(7).Text = AdoPedidos.Recordset!p_fecConfirma
               If PedRecibido Then
                  'Obtengo el folio de la Nota de entrada
                  AdoNotEnt.ConnectionString = cCadConex
                  AdoNotEnt.CommandType = adCmdText
                  AdoNotEnt.RecordSource = "SELECT * FROM [NotaEntrada] WHERE [Pedido] = '" & txtcampos(0).Text & "'"
                  AdoNotEnt.Refresh
                  For N = 8 To 12
                      lbletiquetas(N).Visible = True
                      txtcampos(N).Visible = True
                      txtcampos(N).Enabled = False
                  Next
                  If lProvAbto = Ins Then
                     dbgrdDetPed.Columns(2).Visible = False
                     dbgrdDetPed.Columns(4).Visible = True
                     dbgrdDetPed.Columns(6).Visible = True
                  End If
                  cmbJefeRec.Visible = True
                  cmbJefeRec.Enabled = False
               End If
               chkFields(5).Visible = True
               chkFields(6).Visible = (chksugerido.Value = 0)
               chkFields(5).Value = IIf(pConfirmado, 1, 0)
               chkFields(6).Value = IIf(PedRecibido, 1, 0)
               chkFields(5).Enabled = False
               chkFields(6).Enabled = False

               cmdGrabar.Enabled = False
               txtcampos(7).Enabled = False
            Else
               chkFields(5).Visible = True
               chkFields(5).Enabled = (lProvAbto <> Ind)
            End If
            For N = 1 To 6   'Desactivo todos los cuadros de texto
               txtcampos(N).Enabled = False
            Next
            
            dbgrdDetPed.AllowAddNew = False
            dbgrdDetPed.AllowDelete = False
            dbgrdDetPed.AllowUpdate = False
            dbgrdDetPed.Columns(0).Button = False
         ElseIf cModo = "RECIBIRPEDIDO" Then
            cmbProved.Enabled = False
            cmbSucursal.Enabled = False
            cmbUsuarios.Enabled = False
            cmdDespla.Enabled = False
            cmdReporte.Enabled = True

            lbletiquetas(7).Visible = True
            txtcampos(7).Visible = True
            txtcampos(7).Enabled = False
            lbletiquetas(8).Visible = PedRecibido  'Fecha de recepcion
            txtcampos(8).Visible = PedRecibido
            txtcampos(8).Enabled = Not PedRecibido
            
            dbgrdfactu.Visible = PedRecibido
            dbgrdfactu.AllowUpdate = Not PedRecibido
            For N = 5 To 6
                txtcampos(N).Enabled = False
                chkFields(N).Visible = True
            Next
            cmdGrabar.Enabled = Not PedRecibido
            chkFields(5).Value = 1
            chkFields(6).Value = 1
            chkFields(6).Enabled = Not PedRecibido And lProvAbto <> Ind
            chkFields(5).Enabled = Not PedRecibido And lProvAbto <> Ind
            dbgrdDetPed.AllowUpdate = Not PedRecibido
            dbgrdDetPed.Columns(0).Button = False
            dbgrdDetPed.Columns(2).Visible = False
            dbgrdDetPed.Columns(4).Visible = True
            dbgrdDetPed.Columns(6).Visible = True
            dbgrdDetPed.Columns(0).Locked = True
            dbgrdDetPed.Columns(3).Locked = True
            dbgrdDetPed.Columns(4).Locked = PedRecibido
            dbgrdDetPed.Columns(6).Locked = PedRecibido
        
            lbletiquetas(9).Visible = True
            txtcampos(9).Visible = True
            txtcampos(9).Enabled = Not PedRecibido And lProvAbto <> Ind
            cmbJefeRec.Visible = True
            cmbJefeRec.Enabled = Not PedRecibido And lProvAbto <> Ind
            If Not PedRecibido And lProvAbto <> Ind Then  'Si no se ha recibido y es diferente indirecto
               txtcampos(9).SetFocus
            ElseIf Not PedRecibido And lProvAbto = Ind Then
               dbgrdDetPed.Columns(4).Locked = True  'Bloqueo la columna de cantidad recibida
               
            Else
               AdoCatUsu.Recordset.MoveFirst
               AdoCatUsu.Recordset.Find "Clave = '" & Trim(txtcampos(9).Text) & "'"
               If AdoCatUsu.Recordset.EOF = False Then cmbJefeRec.Text = AdoCatUsu.Recordset!Name
            End If
            poninfo   'Muestra informacion de los productos capturados
            
            If Not PedRecibido Then
               If cn.State = 0 Then cn.Open
               cn.Execute "UPDATE [DetalleFactura] SET [df_cantreal] = [df_cantsol] WHERE [df_pedido] = '" & txtcampos(0).Text & "'"
            Else
               'Obtengo el folio de la Nota de entrada
               AdoNotEnt.ConnectionString = cCadConex
               AdoNotEnt.CommandType = adCmdText
               AdoNotEnt.RecordSource = "SELECT * FROM [NotaEntrada] WHERE [Pedido] = '" & txtcampos(0).Text & "'"
               AdoNotEnt.Refresh
               lbletiquetas(12).Visible = True
               txtcampos(12).Visible = True
               txtcampos(12).Enabled = False
               txtcampos(1).Enabled = False
               txtcampos(3).Enabled = False
               MsgBox "PEDIDO SOLO PARA CONSULTA", vbInformation
            End If
            chkFields(6).Value = IIf(PedRecibido, 1, 0)
         End If
         'LLeno el detalle de pedidos
         AdoDetPed.ConnectionString = cn
         AdoDetPed.CommandType = adCmdText
         AdoDetPed.RecordSource = "SELECT * FROM DetalleFactura  WHERE [df_Pedido] = '" & txtcampos(0).Text & "' AND df_sugerido = " & IIf(Trim(txtcampos(1).Text) = "ABA", 0, 1)
         AdoDetPed.Refresh
         'MsgBox AdoDetPed.RecordSource
         
         If cModo = "CAPTURARPEDIDO" Then
            If nOp = 1 Then
              txtcampos(4).Text = date + Time
              If Not IsNull(nFrecuencia) Then
                 Cal1.Value = txtcampos(4).Text
                 txtcampos(5).Text = Cal1.Value + nFrecuencia
              End If
            End If
            txtcampos(6).SetFocus
         End If
         txtcampos(4).Enabled = False
         txtcampos(5).Enabled = False
    Case 9  'Clave del Jefe de recibo
         txtcampos(Index).Text = Trim(UCase(txtcampos(Index).Text))
         txtcampos(Index).Refresh
         If txtcampos(Index).Text = "" Or IsNull(txtcampos(Index).Text) Then
            cmbJefeRec.SetFocus
            Exit Sub
         Else
            AdoCatUsu.Recordset.MoveFirst
            AdoCatUsu.Recordset.Find "Clave = '" & Trim(txtcampos(Index).Text) & "'"
            If AdoCatUsu.Recordset.EOF = True Then
                'MsgBox "No existe la clave del usuario", vbExclamation
                Me.cmbJefeRec.SetFocus
                Exit Sub
            End If
         End If
         cmbJefeRec.Text = AdoCatUsu.Recordset!Name
         'SendKeys "{TAB}"
         keybd_event &H9, 0, 0, 0:   keybd_event &H9, 0, &H2, 0
    Case 10   'Oculto el calendario al pasar a las casillas de verificación
         cresp = MsgBox("ESTE NUMERO DE FACTURA SERA UTILIZADO PARA PAGAR SU IMPORTE" & Chr(13) & _
               Space(40) & "Y DARLE SEGUIMIENTO" & Chr(13) & Chr(13) & Space(20) & "ES CORRECTO EL NUMERO DE FACTURA?", vbQuestion + vbYesNo)
        If Not cresp = vbYes Then txtcampos(Index).SetFocus
End Select
Exit Sub
Error:
   MsgBox Err.Description
End Sub

Private Sub dbgrdDetPed_ButtonClick(ByVal ColIndex As Integer)
On Error GoTo Error:
    Dim L As ListBox
    If AdoDetPed.Recordset.EOF = True Then AdoDetPed.Recordset.AddNew
    Select Case ColIndex
       Case ColIndex_lstProd
            Set L = Lstprod
    End Select
    If ColIndex = -1 Then Exit Sub
      With L
          'Abajo (3):
          .Left = dbgrdDetPed.Left + dbgrdDetPed.Columns(ColIndex).Left
          .Top = dbgrdDetPed.Top + dbgrdDetPed.RowTop(dbgrdDetPed.Row) + dbgrdDetPed.RowHeight
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

Private Sub lstProd_KeyPress(KeyAscii As Integer)
Dim cveprod As String
Dim N As Integer
    'Habilita el uso de teclado para asignar a la celda o Escape...
    Select Case KeyAscii
        Case vbKeyReturn
             'Asigna la clave del producto seleccionado a la celda
             N = InStr(1, Lstprod.List(Lstprod.ListIndex), "[")
             cveprod = Mid(Lstprod.List(Lstprod.ListIndex), N + 1, Len(Lstprod.List(Lstprod.ListIndex)) - N - 1)
             AdoCatPro.Recordset.MoveFirst
             AdoCatPro.Recordset.Find "Consec = '" & Trim(cveprod) & "'"
             dbgrdDetPed.Columns(ColIndex_lstProd).Text = Trim(cveprod)
             lblDescrip.Caption = AdoCatPro.Recordset!descripc + Chr(13) + " (" + Str(AdoCatPro.Recordset!PAQUETES) + " X " + Str(AdoCatPro.Recordset!CONTENID) + " " + AdoCatPro.Recordset!medida + " ) "
             Lstprod.Visible = False
             AdoDetPed.Recordset!df_cantidad = 0
             AdoDetPed.Recordset!df_cantsol = 0
             AdoDetPed.Recordset!df_cantsolp = 0
             AdoDetPed.Recordset!df_cantreal = 0
             AdoDetPed.Recordset!df_cantrealP = 0
             AdoDetPed.Recordset!df_Costo = AdoCatPro.Recordset!PRECOSTO
             AdoDetPed.Recordset!df_sugerido = IIf(txtcampos(1).Text = "ABA", 0, 1)
             'SendKeys vbTab
             keybd_event &H9, 0, 0, 0:   keybd_event &H9, 0, &H2, 0
        Case vbKeyEscape
             Lstprod.Visible = False
    End Select
End Sub

Private Sub dbgrdDetPed_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
   ' Obliga al usuario a usar un ítem de la lista.
   ' En caso de dar al usuario libertad de escribir, elimine las siguientes líneas (If-End If),
   ' o precedales con un comentario
   If ColIndex = ColIndex_lstProd Then
       Cancel = True
       dbgrdDetPed_ButtonClick (ColIndex)
   End If
End Sub

Private Sub LstProd_LostFocus()
    '//Oculta la lista si pierde el enfoque
    Lstprod.Visible = False
End Sub

Private Sub LstProd_DblClick()
    lstProd_KeyPress vbKeyReturn
End Sub


