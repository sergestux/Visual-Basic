VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "mscal.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Begin VB.Form frmpedBod 
   Caption         =   "Menu de pedidos por proveedor..."
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   555
   ClientWidth     =   11880
   Icon            =   "frmpedBod.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc AdoSugDbf 
      Height          =   330
      Left            =   3120
      Top             =   4080
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
      Caption         =   "AdoSugDbf"
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
   Begin MSComDlg.CommonDialog Cmdlg 
      Left            =   0
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodbf 
      Height          =   330
      Left            =   720
      Top             =   4080
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
      Caption         =   "Adodbf"
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
   Begin MSACAL.Calendar Cal1 
      Height          =   1935
      Left            =   8880
      TabIndex        =   16
      Top             =   1680
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
   Begin MSDataGridLib.DataGrid dbgrdPed 
      Bindings        =   "frmpedBod.frx":0442
      Height          =   6465
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   11404
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1.5
      RowHeight       =   15
      RowDividerStyle =   3
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
         DataField       =   "pp_pedido"
         Caption         =   "CVE. PEDIDO"
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
         DataField       =   "pp_proveedor"
         Caption         =   "  CVE. PROV."
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
         DataField       =   "pp_fechagen"
         Caption         =   "      FECHA ELAB."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/mm/yyyy hh:mm AM/PM"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "pp_fecconfirma"
         Caption         =   "  FEC. CONFIRMACION"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/mm/yyyy hh:mm AM/PM"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "pp_recibe"
         Caption         =   " RECIBIDO"
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
      BeginProperty Column05 
         DataField       =   "pp_fecrecibe"
         Caption         =   "   FEC. RECEPCION"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/mm/yyyy hh:mm AM/PM"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "pp_notent"
         Caption         =   "      NOTA ENT."
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
         Locked          =   -1  'True
         BeginProperty Column00 
            Alignment       =   2
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   1950.236
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   1844.787
         EndProperty
         BeginProperty Column06 
            Alignment       =   2
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoPedidos 
      Height          =   375
      Left            =   0
      Top             =   6480
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
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
   Begin VB.Frame fradescripcion 
      Height          =   735
      Left            =   120
      TabIndex        =   14
      Top             =   720
      Width           =   11655
      Begin VB.CheckBox chkFecGen 
         Caption         =   "Fec. de &Elaboracion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6840
         TabIndex        =   20
         Top             =   240
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.ComboBox cmbProved 
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   240
         Sorted          =   -1  'True
         TabIndex        =   19
         Top             =   360
         Width           =   6375
      End
      Begin VB.TextBox txtFecha 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   0
         Left            =   8280
         TabIndex        =   8
         Top             =   300
         Width           =   1455
      End
      Begin VB.TextBox txtFecha 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   1
         Left            =   9960
         TabIndex        =   9
         Top             =   300
         Width           =   1455
      End
      Begin VB.Label LBLETIQUETAS 
         Alignment       =   2  'Center
         Caption         =   "Final"
         Height          =   255
         Index           =   1
         Left            =   9960
         TabIndex        =   18
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label LBLETIQUETAS 
         Alignment       =   2  'Center
         Caption         =   "Inicial"
         Height          =   255
         Index           =   0
         Left            =   8280
         TabIndex        =   17
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label lblProv 
         Alignment       =   2  'Center
         Caption         =   "Proveedor"
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
         Left            =   240
         TabIndex        =   15
         Top             =   120
         Width           =   6375
      End
   End
   Begin VB.Frame cmdDespla 
      Height          =   732
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   11655
      Begin VB.CommandButton CmdMoverse 
         Height          =   400
         Index           =   8
         Left            =   6120
         Picture         =   "frmpedBod.frx":045B
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Pedidos Recibidos Con Costos"
         Top             =   240
         Width           =   500
      End
      Begin VB.CommandButton CmdMoverse 
         Height          =   400
         Index           =   5
         Left            =   5520
         Picture         =   "frmpedBod.frx":098D
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Reporte global de pedidos por proveedor en el rango especificado"
         Top             =   255
         Width           =   500
      End
      Begin VB.CommandButton CmdMoverse 
         Height          =   400
         Index           =   4
         Left            =   3720
         Picture         =   "frmpedBod.frx":0EBF
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Buscar clave del pedido en el rango seleccionado"
         Top             =   240
         Width           =   500
      End
      Begin VB.CommandButton CmdMoverse 
         Height          =   400
         Index           =   3
         Left            =   3120
         Picture         =   "frmpedBod.frx":1031
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Ir al ultimo"
         Top             =   240
         Width           =   500
      End
      Begin VB.CommandButton CmdMoverse 
         Enabled         =   0   'False
         Height          =   400
         Index           =   2
         Left            =   3000
         Picture         =   "frmpedBod.frx":11A3
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Ir al siguiente"
         Top             =   0
         Visible         =   0   'False
         Width           =   500
      End
      Begin VB.CommandButton CmdMoverse 
         Enabled         =   0   'False
         Height          =   400
         Index           =   1
         Left            =   3000
         Picture         =   "frmpedBod.frx":1315
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Ir al anterior"
         Top             =   0
         Visible         =   0   'False
         Width           =   500
      End
      Begin VB.CommandButton CmdMoverse 
         Height          =   400
         Index           =   0
         Left            =   2520
         Picture         =   "frmpedBod.frx":1487
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Ir al primero"
         Top             =   240
         Width           =   500
      End
      Begin VB.CommandButton CmdMoverse 
         Height          =   400
         Index           =   6
         Left            =   4320
         Picture         =   "frmpedBod.frx":15F9
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Eficiencia por Pedido"
         Top             =   240
         Width           =   500
      End
      Begin VB.CommandButton CmdMoverse 
         Height          =   400
         Index           =   7
         Left            =   4920
         Picture         =   "frmpedBod.frx":1B2B
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Llegadas por Periodo con Eficiencia"
         Top             =   240
         Width           =   500
      End
      Begin VB.CommandButton cmdOpcion 
         Height          =   375
         Index           =   7
         Left            =   1320
         Picture         =   "frmpedBod.frx":205D
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Exportar llegadas de pedidos para enviarlas a Oficinas centrales"
         Top             =   240
         Width           =   500
      End
      Begin VB.CommandButton cmdOpcion 
         Height          =   400
         Index           =   3
         Left            =   1920
         Picture         =   "frmpedBod.frx":2367
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
         Width           =   500
      End
      Begin VB.CommandButton cmdOpcion 
         Height          =   400
         Index           =   2
         Left            =   120
         Picture         =   "frmpedBod.frx":2671
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   240
         Width           =   500
      End
      Begin VB.CommandButton cmdOpcion 
         Height          =   400
         Index           =   4
         Left            =   720
         Picture         =   "frmpedBod.frx":297B
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Importar pedidos por proveedor  enviado por Oficinas centrales"
         Top             =   240
         Width           =   500
      End
      Begin VB.ComboBox cmbOrden 
         Height          =   315
         Left            =   9000
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton cmdOpcion 
         Height          =   400
         Index           =   5
         Left            =   11040
         Picture         =   "frmpedBod.frx":2C85
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Regresar a la pantalla principal"
         Top             =   240
         Width           =   500
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "&Ver Ped. Conf."
         Height          =   372
         Index           =   1
         Left            =   10440
         TabIndex        =   4
         ToolTipText     =   "Modificar pedido capturado"
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "&Nuevo"
         Height          =   372
         Index           =   0
         Left            =   10440
         TabIndex        =   12
         ToolTipText     =   "Consulta o generar pedido por proveedor"
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "&Codigo Barras"
         Height          =   375
         Index           =   6
         Left            =   10440
         TabIndex        =   5
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Numero de pedidos: 999,999"
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
         Left            =   6960
         TabIndex        =   30
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblFiltro 
         Alignment       =   2  'Center
         Caption         =   "Filtro"
         Height          =   255
         Left            =   9480
         TabIndex        =   13
         Top             =   120
         Width           =   735
      End
   End
   Begin ComctlLib.StatusBar Stb1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   21
      Top             =   7980
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   582
      Style           =   1
      SimpleText      =   "                                                  Click en el encabezado ordena los datos en base a la columna   "
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin Crystal.CrystalReport CR1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowLeft      =   0
      WindowTop       =   0
      WindowState     =   2
   End
   Begin VB.Menu mped 
      Caption         =   "&Pedidos"
      Begin VB.Menu mrec 
         Caption         =   "&Recibir"
      End
      Begin VB.Menu mimp 
         Caption         =   "&Importar Pedidos de Oficinas"
      End
      Begin VB.Menu mgen 
         Caption         =   "&Generar Llegadas"
      End
   End
   Begin VB.Menu mrep 
      Caption         =   "R&eportes"
      Begin VB.Menu mefiprov 
         Caption         =   "Eficiencia por Proveedor"
      End
      Begin VB.Menu mefiperi 
         Caption         =   "Eficiencia por Periodo"
      End
      Begin VB.Menu mperiodo 
         Caption         =   "Pedidos por Periodo"
      End
      Begin VB.Menu mllegadas 
         Caption         =   "Llegadas"
      End
   End
   Begin VB.Menu mbusca 
      Caption         =   "&Buscar"
      Begin VB.Menu mclave 
         Caption         =   "Por Clave"
      End
      Begin VB.Menu mprov 
         Caption         =   "Por Proveedor"
      End
   End
End
Attribute VB_Name = "frmpedBod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private prin As Boolean
Private cCond As String     'Condicion del filtro del grid
Private ccondrpt As String  'Condicion del filtro del rpt
Private cFecha As String
Private ntext As Integer

Private Sub AdoPedidos_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.ERROR, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error Resume Next
 cmbProved.Text = AdoPedidos.Recordset!nomprove & "  " & AdoPedidos.Recordset!PROVE
 
End Sub

Private Sub Cal1_DblClick()
On Error GoTo ERROR:
 txtFecha(ntext).Text = Cal1.Value
 Cal1.Visible = False
 txtFecha(ntext).SetFocus
 SendKeys "{TAB}"
 Exit Sub
ERROR:
  MsgBox Err.Description

End Sub

Private Sub chkFecGen_Click()
chkFecGen.Caption = IIf(chkFecGen.Value = 1, "Fecha de &Elaboracion", "Fecha de &Recepcion")
End Sub

Private Sub cmborden_DblClick()
  SendKeys "{TAB}"
End Sub

Private Sub cmborden_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   SendKeys "{TAB}"
End If
End Sub

Private Sub cmborden_LostFocus()
Select Case cmbOrden.ListIndex
Case 0 'Todos
    cCond = "Pp_pedido <> ''"
    ccondrpt = "{PEDPROVE.pp_pedido} <> ''"
Case 1 'pendientes por recibir
    cCond = "Pp_recibe = 0"
    ccondrpt = "{PEDPROVE.Pp_recibe} = 0"
Case 2 'Recibidos
    cCond = "Pp_recibe = 1"
    ccondrpt = "{PEDPROVE.Pp_recibe} = 1"
End Select
'cOper = IIf(Month(txtFecha(0).Text) = Month(txtFecha(1).Text), " AND ", " OR ")
'cFecha = " AND (month(pp_fechagen) >= " & Month(txtFecha(0).Text) & " and (day(pp_fechagen) > = " & Day(txtFecha(0).Text) & cOper & " (day(pp_fechagen)<= " & Day(txtFecha(1).Text) & " and month(pp_fechagen)<= " & Month(txtFecha(1).Text) & ")) and year(pp_fechagen)>= " & Year(txtFecha(0).Text) & " and year(pp_fechagen)<= " & Year(txtFecha(1).Text) & ")"
cCamFec = IIf(chkFecGen.Value = 1, "pp_fechaGen", "pp_fecrecibe")
cFecha = " AND " & cCamFec & " >= '" & txtFecha(0).Text & "' AND " & cCamFec & " <= DATEADD(day,1,'" & txtFecha(1).Text & "')"

AdoPedidos.RecordSource = "select * from pedprove,catprov where pp_proveedor = catprov.prove and " & cCond & cFecha & " ORDER BY pp_proveedor"
AdoPedidos.Refresh
For n = 0 To 6
   CmdMoverse(n).Enabled = Not (AdoPedidos.Recordset.BOF And AdoPedidos.Recordset.EOF)
Next
cmdOpcion(1).Enabled = Not (AdoPedidos.Recordset.BOF And AdoPedidos.Recordset.EOF)
'cmdOpcion(2).Enabled = Not (AdoPedidos.Recordset.BOF And AdoPedidos.Recordset.EOF)
cmdOpcion(3).Enabled = Not (AdoPedidos.Recordset.BOF And AdoPedidos.Recordset.EOF)
lblInfo.Caption = Str(AdoPedidos.Recordset.RecordCount)
End Sub

Private Sub cmbproved_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmbProved_LostFocus()
On Error GoTo ERROR:
If Trim(cmbProved.Text) <> "" Then
   'cOper = IIf(Month(txtFecha(0).Text) = Month(txtFecha(1).Text), " AND ", " OR ")
   'cFecha = " AND (month(pp_fechagen) >= " & Month(txtFecha(0).Text) & " and (day(pp_fechagen) > = " & Day(txtFecha(0).Text) & cOper & "(day(pp_fechagen)<= " & Day(txtFecha(1).Text) & " and month(pp_fechagen)<= " & Month(txtFecha(1).Text) & ")) and year(pp_fechagen)>= " & Year(txtFecha(0).Text) & " and year(pp_fechagen)<= " & Year(txtFecha(1).Text) & ")"
   cCamFec = IIf(chkFecGen.Value = 1, "pp_fechaGen", "pp_fecrecibe")
   cFecha = " AND " & cCamFec & " >= '" & txtFecha(0).Text & "' AND " & cCamFec & " <= DATEADD(day,1,'" & txtFecha(1).Text & "')"
   cadena = "SELECT * FROM pedprove,catprov WHERE pp_proveedor = catprov.prove AND pp_proveedor = '" & Trim(Mid(cmbProved.Text, Len(cmbProved.Text) - 5)) & "' AND " & cCond & cFecha & " ORDER BY pp_fechagen DESC"
   AdoPedidos.RecordSource = cadena
Else
   'cOper = IIf(Month(txtFecha(0).Text) = Month(txtFecha(1).Text), " AND ", " OR ")
   'cFecha = " AND (month(pp_fechagen) >= " & Month(txtFecha(0).Text) & " and (day(pp_fechagen) > = " & Day(txtFecha(0).Text) & cOper & "(day(pp_fechagen)<= " & Day(txtFecha(1).Text) & " and month(pp_fechagen)<= " & Month(txtFecha(1).Text) & ")) and year(pp_fechagen)>= " & Year(txtFecha(0).Text) & " and year(pp_fechagen)<= " & Year(txtFecha(1).Text) & ")"
   cCamFec = IIf(chkFecGen.Value = 1, "pp_fechaGen", "pp_fecrecibe")
   cFecha = " AND " & cCamFec & " >= '" & txtFecha(0).Text & "' AND " & cCamFec & " <= DATEADD(day,1,'" & txtFecha(1).Text & "')"
   cadena = "SELECT * FROM pedprove,catprov WHERE pp_proveedor = catprov.prove AND " & cCond & cFecha
   AdoPedidos.RecordSource = cadena
End If
AdoPedidos.Refresh
dbgrdPed.SetFocus
Exit Sub
ERROR:
   Me.cmbProved.SetFocus
   Exit Sub
End Sub

Private Sub cmdMoverse_Click(Index As Integer)
'On Error Resume Next
Dim rs As ADODB.Recordset
If AdoPedidos.Recordset.BOF And AdoPedidos.Recordset.EOF Then
   For n = 0 To 4
       CmdMoverse(n).Enabled = False
   Next
   Exit Sub
End If
Select Case Index
Case 0  'Primer registro
    AdoPedidos.Recordset.MoveFirst
Case 1  ' Anterior
    AdoPedidos.Recordset.MovePrevious
    If AdoPedidos.Recordset.BOF Then AdoPedidos.Recordset.MoveFirst
Case 2  ' Siguiente
    AdoPedidos.Recordset.MoveNext
    If AdoPedidos.Recordset.EOF Then AdoPedidos.Recordset.MoveLast
Case 3  'Ultimo
    AdoPedidos.Recordset.MoveLast
Case 4
    cCve = InputBox("Introduzca la clave del pedido a buscar", "Introducir clave")
    If IsNull(Trim(cCve)) Or Trim(cCve) = "" Then Exit Sub
    Antes = dbgrdPed.Bookmark
    AdoPedidos.Recordset.MoveFirst
    AdoPedidos.Recordset.Find "pp_pedido = '" & Trim(cCve) & "'"
    If AdoPedidos.Recordset.EOF Then
        MsgBox "LA CLAVE " & cCve & " NO SE ENCUENTRA EN LOS PEDIDOS " + IIf(cmbOrden.Text = "TODOS", "", cmbOrden.Text), vbExclamation
        dbgrdPed.Bookmark = Antes
    End If
Case 5  'Reporte global de pedidos seleccionados
    cMensaje = Stb1.SimpleText
    Stb1.SimpleText = Space(90) + "Espere un momento generando reporte..."
    Stb1.Refresh
    cFecharpt = " AND {PEDPROVE.Pp_fechagen} >= Date(" & CStr(Year(txtFecha(0).Text)) & "," & CStr(Month(txtFecha(0).Text)) & "," & CStr(Day(txtFecha(0).Text)) & ") AND {PEDPROVE.Pp_fechagen} <= Date(" & CStr(Year(txtFecha(1).Text)) & "," & CStr(Month(txtFecha(1).Text)) & "," & CStr(Day(txtFecha(1).Text)) & ")"
    CR1.Connect = cn
    CR1.ReportFileName = App.Path & "\PedProv.rpt"
    'frmpedBod.CR1.SQLQuery = "SELECT CATTIENDA.ticlave, CATTIENDA.tidescrip, CATTIENDA.direccion, CATTIENDA.foliotie From  pitico.dbo.CATTIENDA CATTIENDA Where CATTIENDA.ticlave = '5'"
    'CR1.ReportFileName = App.Path & "\borrar.rpt"
    CR1.WindowTitle = "Reporte de pedidos por proveedor"
    CR1.Formulas(0) = "FORMSELEC = " & ccondrpt & cFecharpt
    CR1.Formulas(1) = "PEDIDO = 'LISTADO DE PEDIDOS POR PROVEEDOR " & IIf(cmbOrden.Text = "TODOS", "", cmbOrden.Text) & " '"
    frmpedBod.CR1.Action = 1
    Stb1.SimpleText = cMensaje
    Stb1.Refresh
Case 6 ' por pedido
    Llegadas (2)
Case 7 'por pedido
    Llegadas (1)
Case 8
    cn.Execute "DELETE FROM LLEGADASTMP"
    'Pedidos por proveedor
    cn.Execute "INSERT INTO LLEGADASTMP(claprove,pedido,fecha,cantsol,cantrec,promsol,promrec,impterec,factura1,impfac1,factura2,impfac2,factura3,impfac3,factura4,impfac4,factura5,impfac5,factura6,impfac6,factura7,impfac7,factura8,impfac8,factura9,impfac9,factura10,impfac10) SELECT MAX(PP_PROVEEDOR), PP_pedido, max(pp_fecrecibe), sum(dg_cantsol), sum(dg_cantreal),sum(dg_promocion), sum(dg_promocionr), sum(dg_cantreal * dg_costo), max(factura1), max(impfac1), max(factura2), max(impfac2), max(factura3), " & _
            " max(impfac3), max(factura4), max(impfac4), max(factura5), max(impfac5), max(factura6), max(impfac6), max(factura7), max(impfac7), max(factura8), max(impfac8), max(factura9), max(impfac9), max(factura10), max(impfac10) from  detalleglobal,pedprove,notaentrada " & _
            " WHERE dg_cantreal > 0 and pp_fecrecibe >= '" & txtFecha(0).Text & "'  and pp_fecrecibe <= '" & DateAdd("d", 1, txtFecha(1).Text) & "' and pp_pedback is null and pp_pedido = dg_pedido and pedido = pp_pedido  group by pp_pedido "
    'Pedidos directos o instantáneos
    cn.Execute "INSERT INTO LLEGADASTMP(claprove,pedido,fecha,cantsol,cantrec,promsol,promrec,impterec,factura1,impfac1,factura2,impfac2,factura3,impfac3,factura4,impfac4,factura5,impfac5,factura6,impfac6,factura7,impfac7,factura8,impfac8,factura9,impfac9,factura10,impfac10,pedprove) SELECT MAX(p_proveedor), p_pedido, max(p_fecentreal), sum(df_cantsol), sum(df_cantreal), prom = 0, promr = 0 , sum(df_cantreal * df_costo), max(factura1), max(impfac1), max(factura2), max(impfac2), max(factura3), " & _
            " max(impfac3), max(factura4), max(impfac4), max(factura5), max(impfac5), max(factura6), max(impfac6), max(factura7), max(impfac7), max(factura8), max(impfac8), max(factura9), max(impfac9), max(factura10), max(impfac10), tipo = 0 FROM detallefactura,pedidos,notaentrada " & _
            " WHERE df_cantreal > 0 and p_fecentreal >= '" & txtFecha(0).Text & "' and p_fecentreal <= '" & DateAdd("d", 1, txtFecha(1).Text) & "' AND p_pedido = df_pedido AND pedido = p_pedido AND p_cancelado = 0 GROUP BY p_pedido "

    cMensaje = Stb1.SimpleText
    Stb1.SimpleText = Space(90) + "Espere un momento generando reporte..."
    Stb1.Refresh
    'cFecharpt = " AND {PEDPROVE.Pp_fechagen} >= Date(" & CStr(Year(txtFecha(0).Text)) & "," & CStr(Month(txtFecha(0).Text)) & "," & CStr(Day(txtFecha(0).Text)) & ") AND {PEDPROVE.Pp_fechagen} <= Date(" & CStr(Year(txtFecha(1).Text)) & "," & CStr(Month(txtFecha(1).Text)) & "," & CStr(Day(txtFecha(1).Text)) & ")"
    CR1.Connect = cn
    CR1.ReportFileName = App.Path & "\Prefipro.rpt"
    CR1.WindowTitle = "Reporte de pedidos por proveedor"
    CR1.Formulas(0) = "ENCAB = ' PEDIDOS RECIBIDOS DEL " & txtFecha(0).Text & " AL " & txtFecha(1).Text & " '"
    'CR1.Formulas(1) = "PEDIDO = 'LISTADO DE PEDIDOS POR PROVEEDOR " & IIf(cmbOrden.Text = "TODOS", "", cmbOrden.Text) & " '"
    frmpedBod.CR1.Action = 1
    Stb1.SimpleText = cMensaje
    Stb1.Refresh
End Select
Exit Sub
ERROR:
  MsgBox Err.Description
End Sub

Private Sub Llegadas(tipo As Integer)
CR1.ReportFileName = App.Path & "\PrEntPed.rpt"
CR1.WindowTitle = "Nota de entrada de las llegadas del " & Me.txtFecha(0).Text & " al " & Me.txtFecha(1).Text
CR1.Connect = cn
If tipo = 1 Then
'POR RANGO DE FECHAS
CR1.SQLQuery = "SELECT PEDPROVE.pp_proveedor, PEDPROVE.pp_pedido, PEDPROVE.pp_fechagen, PEDPROVE.pp_fecconfirma, PEDPROVE.pp_fecrecibe, " & _
                            "DETALLEGLOBAL.dg_producto, DETALLEGLOBAL.dg_cantsol, DETALLEGLOBAL.dg_cantsolp, DETALLEGLOBAL.dg_cantreal, DETALLEGLOBAL.dg_cantrealp, DETALLEGLOBAL.dg_promocion, DETALLEGLOBAL.dg_costo, DETALLEGLOBAL.dg_promocionr, " & _
                            "NOTAENTRADA.factura1, NOTAENTRADA.impfac1, NOTAENTRADA.factura2, NOTAENTRADA.impfac2, NOTAENTRADA.factura3, NOTAENTRADA.impfac3, NOTAENTRADA.factura4, NOTAENTRADA.impfac4, NOTAENTRADA.factura5, NOTAENTRADA.impfac5, NOTAENTRADA.factura6, NOTAENTRADA.impfac6, NOTAENTRADA.factura7, NOTAENTRADA.impfac7, NOTAENTRADA.factura8, NOTAENTRADA.impfac8, NOTAENTRADA.factura9, NOTAENTRADA.impfac9, NOTAENTRADA.factura10, NOTAENTRADA.impfac10, " & _
                            "CATPROV.NOMPROVE, " & _
                            "TFPRODUC.DESCRIPC, TFPRODUC.CONTENID, TFPRODUC.MEDIDA, TFPRODUC.PAQUETES " & Chr(13) & _
                       "FROM pitico.dbo.TFPRODUC TFPRODUC, " & _
                            "pitico.dbo.DETALLEGLOBAL DETALLEGLOBAL, " & _
                            "pitico.dbo.PEDPROVE PEDPROVE, " & _
                            "pitico.dbo.NOTAENTRADA NOTAENTRADA , " & _
                            "pitico.dbo.CATPROV CATPROV " & Chr(13) & _
                       "WHERE PEDPROVE.pp_pedido = DETALLEGLOBAL.dg_pedido AND " & _
                            "PEDPROVE.pp_pedido = NOTAENTRADA.pedido AND " & _
                            "PEDPROVE.pp_proveedor = CATPROV.PROVE AND " & _
                            "DETALLEGLOBAL.dg_producto = TFPRODUC.CONSEC AND " & _
                            "(DETALLEGLOBAL.dg_cantsol > 0 OR DETALLEGLOBAL.dg_cantsolP > 0) AND " & _
                            "PEDPROVE.pp_fecrecibe >= '" & Format(txtFecha(0).Text, "yyyy-dd-mm") & "' AND PEDPROVE.pp_fecrecibe <= '" & Format(txtFecha(1).Text, "yyyy-dd-mm") & "' " & Chr(13) & _
                       "ORDER BY PEDPROVE.pp_pedido ASC, " & _
                            "TFPRODUC.DESCRIPC ASC, " & _
                            "TFPRODUC.CONTENID ASC"
Else

'POR PEDIDO
CR1.WindowTitle = "Nota de entrada del pedido " & AdoPedidos.Recordset!pp_pedido
'CR1.Connect = cn
CR1.SQLQuery = "SELECT PEDPROVE.pp_proveedor, PEDPROVE.pp_pedido, PEDPROVE.pp_fechagen, PEDPROVE.pp_fecconfirma, PEDPROVE.pp_fecrecibe, " & _
                            "DETALLEGLOBAL.dg_producto, DETALLEGLOBAL.dg_cantsol, DETALLEGLOBAL.dg_cantsolp, DETALLEGLOBAL.dg_cantreal, DETALLEGLOBAL.dg_cantrealp, DETALLEGLOBAL.dg_promocion, DETALLEGLOBAL.dg_costo, DETALLEGLOBAL.dg_promocionr, " & _
                            "NOTAENTRADA.factura1, NOTAENTRADA.impfac1, NOTAENTRADA.factura2, NOTAENTRADA.impfac2, NOTAENTRADA.factura3, NOTAENTRADA.impfac3, NOTAENTRADA.factura4, NOTAENTRADA.impfac4, NOTAENTRADA.factura5, NOTAENTRADA.impfac5, NOTAENTRADA.factura6, NOTAENTRADA.impfac6, NOTAENTRADA.factura7, NOTAENTRADA.impfac7, NOTAENTRADA.factura8, NOTAENTRADA.impfac8, NOTAENTRADA.factura9, NOTAENTRADA.impfac9, NOTAENTRADA.factura10, NOTAENTRADA.impfac10, " & _
                            "CATPROV.NOMPROVE, " & _
                            "TFPRODUC.DESCRIPC, TFPRODUC.CONTENID, TFPRODUC.MEDIDA, TFPRODUC.PAQUETES " & Chr(13) & _
                       "FROM pitico.dbo.TFPRODUC TFPRODUC, " & _
                            "pitico.dbo.DETALLEGLOBAL DETALLEGLOBAL, " & _
                            "pitico.dbo.PEDPROVE PEDPROVE, " & _
                            "pitico.dbo.NOTAENTRADA NOTAENTRADA , " & _
                            "pitico.dbo.CATPROV CATPROV " & Chr(13) & _
                       "WHERE PEDPROVE.pp_pedido = DETALLEGLOBAL.dg_pedido AND " & _
                            "PEDPROVE.pp_pedido = NOTAENTRADA.pedido AND " & _
                            "PEDPROVE.pp_proveedor = CATPROV.PROVE AND " & _
                            "DETALLEGLOBAL.dg_producto = TFPRODUC.CONSEC AND " & _
                            "(DETALLEGLOBAL.dg_cantsol > 0 OR DETALLEGLOBAL.dg_cantsolP > 0) AND " & _
                            "PEDPROVE.pp_pedido = '" & AdoPedidos.Recordset!pp_pedido & "' " & Chr(13) & _
                       "ORDER BY PEDPROVE.pp_pedido ASC, " & _
                            "TFPRODUC.DESCRIPC ASC, " & _
                            "TFPRODUC.CONTENID ASC"
End If
CR1.Action = 1
Stb1.SimpleText = cMensaje
Stb1.Refresh
End Sub
Private Sub cmdopcion_Click(Index As Integer)
'On Error GoTo ERROR:
Select Case Index
  Case 0  'Nuevo pedido
       cModo = ""
       nOp = 1
       frmPedProv.Caption = "Consultar pedidos por tienda para formar uno por proveedor"
       frmPedProv.Show
  Case 1  'Modificar pedido
       nOp = 0
       cModo = "VERCONF"
       frmPedProv.Caption = "Consultar pedido por proveedor confirmado"
       frmPedProv.Show
       SendKeys "{TAB}"
  Case 2  'Recibir pedido
       cModo = "RECIBIR"
       nOp = 0
       frmPedProv.Caption = "Recibir pedido por proveedor"
       frmPedProv.Show
       SendKeys "{TAB}"
  Case 3
       frmBackDet.Show
  Case 4 'Importar Pedidos por Proveedor enviados por oficinas centrales
       ImpPedOfi
       'IMPORTA2
  Case 5  'Salir del modulo de pedidos
       Unload Me
  Case 6
       frmCodBarra.Show
       prin = False
  Case 7
       ExpPedCar  ' Exporta llegadas de pedidos para Oficinas centrales
End Select
Exit Sub
ERROR:
  MsgBox Err.Description
End Sub

Private Sub cmdopcion_GotFocus(Index As Integer)
If Index = 0 Then Unload frmAreaRecibo
End Sub

Private Sub dbgrdPed_DblClick()
  If Not (AdoPedidos.Recordset.BOF And AdoPedidos.Recordset.EOF) Then cmdopcion_Click 1
End Sub

Private Sub dbgrdPed_GotFocus()
  Cal1.Visible = False
End Sub

Private Sub DbGrdped_HeadClick(ByVal ColIndex As Integer)
  Stb1.SimpleText = Space(65) + "Espere un momento ordenando Pedidos por " & dbgrdPed.Columns(ColIndex).Caption
  cCamFec = IIf(chkFecGen.Value = 1, "pp_fechaGen", "pp_fecrecibe")
  cFecha = " AND " & cCamFec & " >= '" & txtFecha(0).Text & "' AND " & cCamFec & " <= DATEADD(day,1,'" & txtFecha(1).Text & "')"
  AdoPedidos.RecordSource = "select * from pedprove,catprov where pp_proveedor = catprov.prove and " & cCond & cFecha & " ORDER BY " & dbgrdPed.Columns(ColIndex).DataField & " DESC"
  AdoPedidos.Refresh
  Stb1.SimpleText = Space(85) + "Pedidos ordenandos por " & dbgrdPed.Columns(ColIndex).Caption
End Sub

Private Sub dbgrdPed_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
 If dbgrdPed.SelBookmarks.Count > 0 Then dbgrdPed.SelBookmarks.Remove 0
 dbgrdPed.SelBookmarks.Add dbgrdPed.RowBookmark(dbgrdPed.Row)
End Sub

Private Sub Form_Activate()
 Forma = 1
 If txtFecha(0).Text = "" Then txtFecha(0).Text = "01/01/00"
 If txtFecha(1).Text = "" Then txtFecha(1).Text = Date
 cOper = IIf(Month(txtFecha(0).Text) = Month(txtFecha(1).Text), " AND ", " OR ")
 cFecha = " AND (month(pp_fechagen) >= " & Month(txtFecha(0).Text) & " and (day(pp_fechagen) > = " & Day(txtFecha(0).Text) & cOper & " (day(pp_fechagen)<= " & Day(txtFecha(1).Text) & " and month(pp_fechagen)<= " & Month(txtFecha(1).Text) & ")) and year(pp_fechagen)>= " & Year(txtFecha(0).Text) & " and year(pp_fechagen)<= " & Year(txtFecha(1).Text) & ")"
  'Cargo todos los pedidos
  AdoPedidos.ConnectionString = cCadConex
  AdoPedidos.CommandType = adCmdText
  Dim fini As Date
  Dim ffin As Date
  fini = txtFecha(0).Text
  ffin = txtFecha(1).Text
  'cad = "select * from pedprove,catprov where pp_proveedor = catprov.prove and " & cCond & cFecha & " ORDER BY pp_proveedor"
  'PARA LA COMPATIBILIDAD CON ACCESS
  cFecha = " and pp_fechagen >  " & fini - 1 & "  " & " and pp_fechagen < " & ffin + 1 & " "
  cad = "select * from pedprove,catprov where pp_proveedor = catprov.prove and " & cCond & cFecha & " ORDER BY pp_proveedor"
  AdoPedidos.RecordSource = cad
  'AdoPedidos.RecordSource = "SELECT * FROM [PedProve] WHERE " & cCond & cFecha & " ORDER BY pp_proveedor"
  AdoPedidos.Refresh
  'If Not (AdoPedidos.Recordset.BOF And AdoPedidos.Recordset.EOF) Then AdoPedidos.Recordset.MoveLast
  'For n = 0 To 6
  '   CmdMoverse(n).Enabled = Not (AdoPedidos.Recordset.BOF And AdoPedidos.Recordset.EOF)
  'Next
  cmdOpcion(1).Enabled = Not (AdoPedidos.Recordset.BOF And AdoPedidos.Recordset.EOF)
  'cmdOpcion(2).Enabled = Not (AdoPedidos.Recordset.BOF And AdoPedidos.Recordset.EOF)
  cmdOpcion(3).Enabled = Not (AdoPedidos.Recordset.BOF And AdoPedidos.Recordset.EOF)
  'cmdOpcion(4).Enabled = Not (AdoPedidos.Recordset.BOF And AdoPedidos.Recordset.EOF)
  lblInfo.Caption = Str(AdoPedidos.Recordset.RecordCount)
  Unload frmAreaRecibo
  'cmbProved.SetFocus
  prin = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 119 Then frmCalc.Show  'F8
End Sub

Private Sub Form_Load()
Dim rs As ADODB.Recordset
  'Obtengo datos para el nombre del proveedor y la sucursal
  cmbOrden.AddItem "TODOS"
  cmbOrden.AddItem "PENDIENTES DE RECIBIR"
  cmbOrden.AddItem "RECIBIDOS"
  cmbOrden.ListIndex = 0
  Set rs = New ADODB.Recordset
  rs.Open "CATPROV", cn, adOpenKeyset, adLockOptimistic, adCmdTable
  While Not rs.EOF
     cmbProved.AddItem rs!nomprove & "   " & rs!PROVE
     rs.MoveNext
  Wend
  cCond = "pp_pedido <> ''"                ' Filtro por default todos los pedidos
  ccondrpt = "{PEDPROVE.pp_pedido} <> ''"  ' Filtro por default del RPT
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If prin Then frmAreaRecibo.Show
End Sub

Private Sub mbus_Click()
cmdMoverse_Click 4
End Sub

Private Sub mclave_Click()
cmdMoverse_Click 4
End Sub

Private Sub mefiperi_Click()
cmdMoverse_Click 7
End Sub

Private Sub mefiprov_Click()
cmdMoverse_Click 6
End Sub

Private Sub mgen_Click()
cmdopcion_Click 7
End Sub

Private Sub mimp_Click()
cmdopcion_Click 4
End Sub

Private Sub mperiodo_Click()
cmdMoverse_Click 5
End Sub

Private Sub mprov_Click()
resp = SendMessageLong(cmbProved.hwnd, &H14F, True, 1)
cmbProved.SetFocus
End Sub

Private Sub mrec_Click()
cmdopcion_Click 2
End Sub

Private Sub txtFecha_GotFocus(Index As Integer)
  Cal1.Visible = True
  Cal1.Left = txtFecha(Index).Left - 350
  If txtFecha(Index).Text = "" Then Cal1.Value = Date
  Cal1.Value = txtFecha(Index).Text
  ntext = Index
End Sub

Private Sub txtFecha_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

'Importar pedidos por proveedor de Oficinas centrales
Private Sub ImpPedOfi()
'On Error GoTo ERROR:
Dim nDias As Integer
Dim FolPed As String
Dim rsttemp As ADODB.Recordset
Dim rstBus As ADODB.Recordset
Dim cArch  As String
Dim cProv As String
Dim SUC As String
Dim cnFoxPro As ADODB.Connection
Dim lExiste  As Boolean  'Ver si existe el pedido
Dim lrecibido As Boolean 'si ya se recibio el pedido
Dim rs As ADODB.Recordset
   MenAnt = Stb1.SimpleText
   cmdlg.DialogTitle = "Abrir archivo enviado por Oficinas centrales"
   cmdlg.FileName = ""
   cmdlg.CancelError = True   'Para que se genere error al hacer click en el boton cancelar
   cmdlg.Filter = "Archivos Visual Foxpro (*.dbf) | *.dbf"
   cmdlg.ShowOpen
   cRutArc = cmdlg.FileName
   If cRutArc = "" Or IsNull(cRutArc) Then
      MsgBox "DEBE ESPECIFICAR UN NOMBRE DE ARCHIVO", vbExclamation
      Exit Sub
   End If
   For n = 1 To Len(cRutArc)
      If Mid(cRutArc, n, 1) = "\" Then nPos = n
   Next
   cRuta = Mid(cRutArc, 1, nPos)
   cArch = Mid(cRutArc, nPos + 1)
   cArch = Mid(cArch, 1, Len(cArch) - 4) 'Le quito la extension porque si no marca error en la consula SQL.

   Adodbf.CursorType = adOpenKeyset
   Adodbf.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=PITICODBF;SourceDB=" & cRuta
   Adodbf.RecordSource = "SELECT * FROM " & cArch & " WHERE not EMPTY( Proveed) AND (EMPTY(PEDPROVE) OR PEDPROVE IS NULL) ORDER BY Pedido"
   Adodbf.Refresh
   Set rs = New ADODB.Recordset
   If Adodbf.Recordset.BOF And Adodbf.Recordset.EOF Then
      MsgBox "EL ARCHIVO SELECCIONADO ESTA VACIO, NO EXISTEN DATOS PARA IMPORTAR", vbExclamation
      Exit Sub
   ElseIf Adodbf.Recordset!Importado Then
      MsgBox "EL ARCHIVO SELECCIONADO YA FUE IMPORTADO", vbInformation
      Exit Sub
   End If
   'Genero el archivo de reporte
   Open App.Path & "\PEDPROVE.TXT" For Output As #1   ' Abre el archivo para operaciones de salida.
   Print #1, Tab(20); "VIVERES Y LICORES S.A DE C.V."
   Print #1, Tab(10); "CARBONERA 1016 COL. TRINIDAD DE LAS HUERTAS"
   Print #1, "PEDIDOS POR PROVEEDOR RECIBIDOS EL "; UCase(Format(Date, "long date"))  ' Escribe texto en el archivo.
   Print #1,   ' Imprime una línea en blanco en el archivo.
   Print #1, "=========================================================================================="
   Print #1, "PEDIDO        FEC. ELAB.        FEC.CONF              PROVEEDOR"
   Print #1, "=========================================================================================="
   Set rstBus = New ADODB.Recordset  'Sirve para saber si existe o no el pedido
   Set rsttemp = New ADODB.Recordset
   cProAnt = "": nPed = 0
   Adodbf.Recordset.MoveFirst
   While Not Adodbf.Recordset.EOF
         'Agrego Pedido por proveedor
         If Adodbf.Recordset!Pedido <> cProAnt Then
              FolPed = Adodbf.Recordset!Pedido
                 rstBus.Open "SELECT * FROM pedprove WHERE pp_pedido = '" & Trim(FolPed) & "'", cn, adOpenDynamic, adLockOptimistic, adCmdText
                 'If rstBus.RecordCount > 0 Then
                 If Not (rstBus.BOF Or rstBus.EOF) Then
                    lrecibido = rstBus!pp_recibe
                 Else
                    'se unen los campos de observaciones
                    vobserva = " "
                    'vobserva = Adodbf.Recordset!Observa + Adodbf.Recordset!Observa1 + Adodbf.Recordset!Observa2
                    'On Error Resume Next
                    
                    vobserva = Replace(Adodbf.Recordset!OBSERVA, "'", " ") + Replace(Adodbf.Recordset!Observa1, "'", " ") + Replace(Adodbf.Recordset!Observa2, "'", " ")
                    'MsgBox vobserva
                    cn.Execute "INSERT INTO PEDPROVE(pp_proveedor,pp_pedido,pp_fechagen,pp_fecConfirma,pp_observa) VALUES ('" & Adodbf.Recordset!proveed & "','" & FolPed & "','" & Adodbf.Recordset!FecPed & "','" & Adodbf.Recordset!FecConf & "','" & vobserva & "')"
                    'MsgBox "SE GENERO EL PEDIDO POR PROVEEDOR CON FOLIO: " & FolPed & Chr(13) & " DEL PROVEEDOR CON CLAVE: " & AdoDbf.Recordset!proveed & Space(5) & "CON FECHA: " & AdoDbf.Recordset!FecPed, vbInformation
                    lrecibido = False
                 End If
                 nPed = nPed + 1
                 rstBus.Close
                 rstBus.Open "SELECT * FROM catprov WHERE prove = '" & Adodbf.Recordset!proveed & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
                 cadena = Adodbf.Recordset!Pedido & Adodbf.Recordset!FecPed & "   " & Adodbf.Recordset!FecConf & "    " & Adodbf.Recordset!proveed & IIf(rstBus.RecordCount = 0, "", Mid(rstBus!nomprove, 1, 38))
                ' MsgBox cadena
                 Print #1, cadena
                 Print #1, "SUGERIDOS:"
                 rstBus.Close
                 AdoSugDbf.CursorType = adOpenKeyset
                 AdoSugDbf.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=PITICODBF;SourceDB=" & cRuta
                 AdoSugDbf.RecordSource = "SELECT * FROM " & cArch & " WHERE Proveed <> '' AND PEDPROVE = '" & Adodbf.Recordset!Pedido & " ' ORDER BY Pedido"
                 AdoSugDbf.Refresh
                 
                 cProAntSug = ""
                 While Not AdoSugDbf.Recordset.EOF
                     'Se importan los sugeridos del pedido por proveedor
                     If AdoSugDbf.Recordset!Pedido <> cProAntSug Then
                        rstBus.Open "SELECT * FROM PEDIDOS WHERE p_pedido = '" & Trim(AdoSugDbf.Recordset!Pedido) & "' AND p_proveedor = '" & Adodbf.Recordset!proveed & "'", cn, adOpenDynamic, adLockOptimistic, adCmdText
                        'If rstBus.RecordCount > 0 Then
                        If Not (rstBus.BOF Or rstBus.EOF) Then
                           lrecibido = rstBus!P_recibido
                        Else
                           cad = "INSERT INTO pedidos(p_pedido,p_proveedor,p_fecped,p_sucursal,p_observaciones,p_pedproveedor,p_fecconfirma,p_situacion) VALUES ('" & Trim(AdoSugDbf.Recordset!Pedido) & "','" & AdoSugDbf.Recordset!proveed & "','" & AdoSugDbf.Recordset!FecPed & "','" & AdoSugDbf.Recordset!sucursal & "','" & AdoSugDbf.Recordset!OBSERVA & "','" & AdoSugDbf.Recordset!PEDPROVE & "','" & AdoSugDbf.Recordset!FecConf & "',1)"
                           'MsgBox cad
                           cn.Execute cad
                           lrecibido = False
                        End If
                        Print #1, Space(13) & AdoSugDbf.Recordset!Pedido
                        rstBus.Close
                     End If
                     cadena = "SELECT * FROM detallefactura WHERE  Df_prod = '" & Trim(AdoSugDbf.Recordset!PRODUCTO) & "' AND df_pedido = '" & AdoSugDbf.Recordset!Pedido & "' "
                     'MsgBox CADENA
                     rstBus.Open cadena, cn, adOpenDynamic, adLockOptimistic, adCmdText
                     If Not (rstBus.BOF Or rstBus.EOF) Then
                     'If rstBus.RecordCount > 0 Then
                        Stb1.SimpleText = Space(50) & "Actualizando producto: " & AdoSugDbf.Recordset!PRODUCTO & " del pedido sugerido " & AdoSugDbf.Recordset!Pedido
                        cn.Execute "UPDATE detallefactura SET df_cantsol = " & AdoSugDbf.Recordset!CantSolc & ", df_cantsolP = " & IIf(IsNull(AdoSugDbf.Recordset!cantsolp), 0, AdoSugDbf.Recordset!cantsolp) & ", df_promocion = " & IIf(IsNull(AdoSugDbf.Recordset!promsol), 0, AdoSugDbf.Recordset!promsol) & " FROM Pedidos WHERE df_pedido = p_pedido AND df_prod = '" & AdoSugDbf.Recordset!PRODUCTO & "' AND df_pedido = '" & AdoSugDbf.Recordset!Pedido & "' AND p_proveedor = '" & AdoSugDbf.Recordset!proveed & "'"
                     Else
                        Stb1.SimpleText = Space(50) & "Agregando producto: " & AdoSugDbf.Recordset!PRODUCTO & " del pedido sugerido " & AdoSugDbf.Recordset!Pedido
                        cadena = "INSERT INTO Detallefactura(df_pedido, df_prod,df_cantsol,df_cantsolp,df_promocion,df_costo) VALUES ('" & AdoSugDbf.Recordset!Pedido & "','" & AdoSugDbf.Recordset!PRODUCTO & "'," & AdoSugDbf.Recordset!CantSolc & "," & IIf(IsNull(AdoSugDbf.Recordset!cantsolp), 0, AdoSugDbf.Recordset!cantsolp) & "," & IIf(IsNull(AdoSugDbf.Recordset!promsol), 0, AdoSugDbf.Recordset!promsol) & "," & AdoSugDbf.Recordset!costo & ")"
                        cn.Execute cadena
                     End If
                     rstBus.Close
                     cProAntSug = AdoSugDbf.Recordset!Pedido
                     AdoSugDbf.Recordset.MoveNext
                 Wend
         End If
         rsttemp.Open "SELECT * FROM TFPRODUC WHERE CONSEC = '" & CStr(Adodbf.Recordset!PRODUCTO) & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
         If rsttemp.BOF And rsttemp.EOF Then
            MsgBox "EL PRODUCTO NO SE ENCUENTRA EN EL CATALOGO, ANOTE LOS DATOS E INFORME AL ADMINISTRADOR DEL SISTEMA" & Chr(13) & "PEDIDO:" & FolPed & Chr(13) & "PROVEEDOR: " & Adodbf.Recordset!proveed & Chr(13) & "CLAVE PRODUCTO: " & CStr(Adodbf.Recordset!PRODUCTO) & Chr(13) & "CAJAS: " & CStr(Adodbf.Recordset!CantSolc) & Chr(13) & "PIEZAS: " & CStr(Adodbf.Recordset!cantsolp), vbCritical
         End If
         'SE MODIFICAN SOLOS PEDIDOS PENDIENTES DE RECIBIR
         If lrecibido = False Then
                rstBus.Open "SELECT * FROM detalleglobal WHERE DG_PRODUCTO = '" & Trim(Adodbf.Recordset!PRODUCTO) & "' AND DG_pedido = '" & FolPed & "'", cn, adOpenDynamic, adLockOptimistic, adCmdText
                'If rstBus.RecordCount > 0 Then
                If Not (rstBus.BOF Or rstBus.EOF) Then
                    Stb1.SimpleText = Space(50) & "Actualizando producto: " & Adodbf.Recordset!PRODUCTO & " del pedido por proveedor " & Adodbf.Recordset!Pedido
                    cn.Execute "UPDATE detalleglobal SET dg_cantsol = " & Adodbf.Recordset!CantSolc & ", dg_cantsolP = " & IIf(IsNull(Adodbf.Recordset!cantsolp), 0, Adodbf.Recordset!cantsolp) & ", dg_promocion = " & IIf(IsNull(Adodbf.Recordset!promsol), 0, Adodbf.Recordset!promsol) & " WHERE dg_producto = '" & Adodbf.Recordset!PRODUCTO & "' AND dg_pedido = '" & FolPed & "'"
                Else
                    Stb1.SimpleText = Space(50) & "Agregando producto: " & Adodbf.Recordset!PRODUCTO & " del pedido por proveedor " & Adodbf.Recordset!Pedido
                    cadena = "INSERT INTO DetalleGlobal(dg_pedido,dg_producto,dg_cantsol,dg_cantsolp,dg_promocion,dg_costo) VALUES ('" & FolPed & "','" & Adodbf.Recordset!PRODUCTO & "'," & Adodbf.Recordset!CantSolc & "," & IIf(IsNull(Adodbf.Recordset!cantsolp), 0, Adodbf.Recordset!cantsolp) & "," & IIf(IsNull(Adodbf.Recordset!promsol), 0, Adodbf.Recordset!promsol) & "," & Adodbf.Recordset!costo & ")"
                    cn.Execute cadena
                End If
            rstBus.Close
         End If
         Stb1.Refresh
         cProAnt = Adodbf.Recordset!Pedido
         Adodbf.Recordset.MoveNext
         rsttemp.Close
   Wend
   Adodbf.Recordset.MoveFirst
   MsgBox "SE GENERARON UN TOTAL DE " & CStr(nPed) & " PEDIDOS", vbInformation
   Print #1, "=========================================================================================="
   Print #1, "TOTAL DE PEDIDOS: "; nPed
   
   Close #1   ' Cierra el archivo de reporte
   Handle = Shell("NOTEPAD " & App.Path & "\PEDPROVE.TXT", 1)
   Set cnFoxPro = New ADODB.Connection
   cnFoxPro.ConnectionString = "Provider=MSDASQL.1;DSN=PITICODBF;SourceDB=" & cRuta & ";SourceType=DBF;Exclusive=No;Initial Catalog= " & cRuta
   cnFoxPro.Open
   cnFoxPro.Execute "UPDATE " & cArch & " SET Importado = 1 "
   AdoPedidos.Refresh
   Stb1.SimpleText = MenAnt
   Stb1.Refresh
  Exit Sub
ERROR:
  If Err.Number <> 32755 Then  'Numero de error al presionar el boton cancelar
     MsgBox Err.Description
  End If
  Close #1   ' Cierra el archivo de reporte
End Sub

Private Sub IMPORTA2()
'On Error GoTo ERROR:
Dim nDias As Integer
Dim FolPed As String
Dim rsttemp As ADODB.Recordset
Dim rstBus As ADODB.Recordset
Dim cArch  As String
Dim cProv As String
Dim SUC As String
Dim cnFoxPro As ADODB.Connection
Dim lExiste  As Boolean  'Ver si existe el pedido
Dim lrecibido As Boolean 'si ya se recibio el pedido
Dim rs As ADODB.Recordset
   MenAnt = Stb1.SimpleText
   cmdlg.DialogTitle = "Abrir archivo enviado por Oficinas centrales"
   cmdlg.FileName = ""
   cmdlg.CancelError = True   'Para que se genere error al hacer click en el boton cancelar
   cmdlg.Filter = "Archivos Visual Foxpro (*.dbf) | *.dbf"
   cmdlg.ShowOpen
   cRutArc = cmdlg.FileName
   If cRutArc = "" Or IsNull(cRutArc) Then
      MsgBox "DEBE ESPECIFICAR UN NOMBRE DE ARCHIVO", vbExclamation
      Exit Sub
   End If
   For n = 1 To Len(cRutArc)
      If Mid(cRutArc, n, 1) = "\" Then nPos = n
   Next
   cRuta = Mid(cRutArc, 1, nPos)
   cArch = Mid(cRutArc, nPos + 1)
   cArch = Mid(cArch, 1, Len(cArch) - 4) 'Le quito la extension porque si no marca error en la consula SQL.
   AdoSugDbf.CursorType = adOpenKeyset
   AdoSugDbf.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=PITICODBF;SourceDB=" & cRuta
   AdoSugDbf.RecordSource = "SELECT * FROM " & cArch
   AdoSugDbf.Refresh
   Set rs = New ADODB.Recordset
   If AdoSugDbf.Recordset.BOF And AdoSugDbf.Recordset.EOF Then
      MsgBox "EL ARCHIVO SELECCIONADO ESTA VACIO, NO EXISTEN DATOS PARA IMPORTAR", vbExclamation
      Exit Sub
   ElseIf AdoSugDbf.Recordset!Importado Then
      MsgBox "EL ARCHIVO SELECCIONADO YA FUE IMPORTADO", vbInformation
      Exit Sub
   End If
   Set rstBus = New ADODB.Recordset  'Sirve para saber si existe o no el pedido
   Set rsttemp = New ADODB.Recordset
   cProAnt = "": nPed = 0
   AdoSugDbf.Recordset.MoveFirst
   cn.Execute "DELETE PEDOFI"
   
   While Not AdoSugDbf.Recordset.EOF
             'MsgBox AdoSugDbf.Recordset!Pedido
            If AdoSugDbf.Recordset!Importado = False Then
               Importado = 0
            Else
               Importado = 1
           End If
           If AdoSugDbf.Recordset!pedsug = False Then
               sugerido = 0
            ElseIf AdoSugDbf.Recordset!pedsug = True Then
               sugerido = 1
            End If
             cadena = "INSERT INTO PEDOFI(pedido,producto,cantsolc,cantsolp,promsol,costo,proveed,fecped,fecconf,observa,importado,sucursal,pedsug,pedprove)" & _
             "VALUES (" & "'" & AdoSugDbf.Recordset!Pedido & "','" & AdoSugDbf.Recordset!PRODUCTO & "'," & AdoSugDbf.Recordset!CantSolc & "," & AdoSugDbf.Recordset!cantsolp & "," & _
              AdoSugDbf.Recordset!promsol & "," & AdoSugDbf.Recordset!costo & ",'" & AdoSugDbf.Recordset!proveed & "','" & _
             AdoSugDbf.Recordset!FecPed & "','" & AdoSugDbf.Recordset!FecConf & "', '" & AdoSugDbf.Recordset!OBSERVA & "','" & Importado & "','" & AdoSugDbf.Recordset!sucursal & "','" & sugerido & "','" & AdoSugDbf.Recordset!PEDPROVE & "')"
'             MsgBox cadena
            cn.Execute cadena
            AdoSugDbf.Recordset.MoveNext
   Wend
 Set cnFoxPro = New ADODB.Connection
   cnFoxPro.ConnectionString = "Provider=MSDASQL.1;DSN=PITICODBF;SourceDB=" & cRuta & ";SourceType=DBF;Exclusive=No;Initial Catalog= " & cRuta
   cnFoxPro.Open
   cnFoxPro.Execute "UPDATE " & cArch & " SET Importado = 1 "
MsgBox "A continuacion se realizara el proceso de importacion de Pedidos por Proveedor y Sugeridos  de Tiendas..., Este proceso tarda Aprox 20 mins."
'SE EJECUTA EL PROCEDIMIENTO ALMACENADO
cn.Execute "exec importapedidos"
End Sub

Private Sub txtFecha_LostFocus(Index As Integer)
On Error Resume Next
 'cOper = IIf(Month(txtFecha(0).Text) = Month(txtFecha(1).Text), " AND ", " OR ")
 'cFecha = " AND (month(pp_fechagen) >= " & Month(txtFecha(0).Text) & " and (day(pp_fechagen) > = " & Day(txtFecha(0).Text) & cOper & " (day(pp_fechagen)<= " & Day(txtFecha(1).Text) & " and month(pp_fechagen)<= " & Month(txtFecha(1).Text) & ")) and year(pp_fechagen)>= " & Year(txtFecha(0).Text) & " and year(pp_fechagen)<= " & Year(txtFecha(1).Text) & ")"
 cCamFec = IIf(chkFecGen.Value = 1, "pp_fechaGen", "pp_fecrecibe")
 cFecha = " AND " & cCamFec & " >= '" & txtFecha(0).Text & "' AND " & cCamFec & " <= DATEADD(day,1,'" & txtFecha(1).Text & "')"

 AdoPedidos.RecordSource = "select * from pedprove,catprov where pp_proveedor = catprov.prove and " & cCond & cFecha & " ORDER BY pp_proveedor"
 AdoPedidos.Refresh
 lblInfo.Caption = Str(AdoPedidos.Recordset.RecordCount)
 For n = 0 To 6   'Si esta vacio el recordset desactivo las opciones
   CmdMoverse(n).Enabled = Not (AdoPedidos.Recordset.BOF And AdoPedidos.Recordset.EOF)
 Next
 cmdOpcion(1).Enabled = Not (AdoPedidos.Recordset.BOF And AdoPedidos.Recordset.EOF)
 'cmdOpcion(2).Enabled = Not (AdoPedidos.Recordset.BOF And AdoPedidos.Recordset.EOF)
 cmdOpcion(3).Enabled = Not (AdoPedidos.Recordset.BOF And AdoPedidos.Recordset.EOF)

End Sub

'Exportar llegadas de pedidos para enviarlas a Oficinas centrales
Private Sub ExpPedCar()
On Error GoTo ERROR:
Dim nDias As Integer
Dim FolPed As String
Dim rsttemp As ADODB.Recordset
Dim cnFoxPro As ADODB.Connection
Dim rs As ADODB.Recordset
   cMenAnt = Stb1.SimpleText
   cmdlg.DialogTitle = "Grabar archivo para enviar llegadas de pedidos a Oficinas centrales"
   cmdlg.Filter = "Archivos Visual Fox pro (*.dbf) | *.dbf"
   cmdlg.CancelError = True
   cmdlg.ShowSave
   cRutArc = cmdlg.FileName
   If cRutArc = "" Or IsNull(cRutArc) Then
      MsgBox "DEBE ESPECIFICAR UNA RUTA Y NOMBRE DE ARCHIVO", vbExclamation
      Exit Sub
   End If
   Stb1.SimpleText = Space(45) & "Grabando archivo " & cRutArc
   Stb1.Refresh
   
   For n = 1 To Len(cRutArc)
      If Mid(cRutArc, n, 1) = "\" Then nPos = n
   Next
   cRuta = Mid(cRutArc, 1, nPos)
   cArch = Mid(cRutArc, nPos + 1)
   cArch = Mid(cArch, 1, Len(cArch) - 4) 'Le quito la extension porque marca error en la consula SQL.

   Stb1.SimpleText = Space(65) & "Limpiando archivo " & cArch
   Stb1.Refresh
   Set fs = CreateObject("Scripting.FileSystemObject")
   Set f = fs.GetFile("\\Server_carbo\Programas\EsPedCar.dbf")
   'Set F = fs.GetFile("c:\paso\EsPedCar.dbf")
   f.Copy cRutArc, True

   Set rsttemp = New ADODB.Recordset
   Adodbf.CommandType = adCmdText
   Adodbf.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=PITICODBF;SourceDB=" & cRuta
   Adodbf.RecordSource = "SELECT * FROM " & cArch
   Adodbf.Refresh
     
  Set rs = New ADODB.Recordset
  cOper = IIf(Month(txtFecha(0).Text) = Month(txtFecha(1).Text), " AND ", " OR ")
  cFecha = " AND (month(pp_fecrecibe) >= " & Month(txtFecha(0).Text) & " and (day(pp_fecrecibe) > = " & Day(txtFecha(0).Text) & cOper & " (day(pp_fecrecibe)<= " & Day(txtFecha(1).Text) & " and month(pp_fecrecibe)<= " & Month(txtFecha(1).Text) & ")) and year(pp_fecrecibe)>= " & Year(txtFecha(0).Text) & " and year(pp_fecrecibe)<= " & Year(txtFecha(1).Text) & ")"
  rs.Open "SELECT * FROM pedprove WHERE pp_recibe = 1 " & cFecha, cn, adOpenKeyset, adLockOptimistic, adCmdText
  nPed = 0
  While Not rs.EOF
       rsttemp.Open "SELECT * FROM Pedprove,DetalleGlobal,NotaEntrada WHERE pp_pedido = '" & rs!pp_pedido & "' AND pp_pedido = dg_pedido AND pp_pedido = Pedido", cn, adOpenStatic, adLockOptimistic, adCmdText
       While Not rsttemp.EOF
            Stb1.SimpleText = Space(15) & "Exportando producto con la clave: " & CStr(rsttemp!DG_PRODUCTO) & " del pedido: " & rs!pp_pedido & " de fecha " & rs!pp_fecrecibe
            Stb1.Refresh
            Adodbf.Recordset.AddNew
            Adodbf.Recordset!Clavenota = Trim(rsttemp!Clavenota)
            Adodbf.Recordset!PRODUCTO = rsttemp!DG_PRODUCTO
            Adodbf.Recordset!CantSolc = Val(rsttemp!dg_cantsol)
            Adodbf.Recordset!cantsolp = Val(rsttemp!dg_cantsolp)
            Adodbf.Recordset!CantRecC = Val(rsttemp!dg_cantreal)
            Adodbf.Recordset!cantrecp = Val(rsttemp!dg_cantrealP)
            Adodbf.Recordset!promsol = Val(rsttemp!dg_promocion)
            Adodbf.Recordset!promrec = Val(rsttemp!dg_promocionr)
            Adodbf.Recordset!costo = rsttemp!DG_COSTO
            Adodbf.Recordset!sucursal = Trim(Mid(cSucursal, 1, 3))
            Adodbf.Recordset!FecRec = Mid(rsttemp!pp_fecrecibe, 1, 8)
            Adodbf.Recordset!factura1 = rsttemp!factura1
            Adodbf.Recordset!Impfac1 = rsttemp!Impfac1
            Adodbf.Recordset!factura2 = rsttemp!factura2
            Adodbf.Recordset!Impfac2 = rsttemp!Impfac2
            Adodbf.Recordset!Factura3 = rsttemp!Factura3
            Adodbf.Recordset!Impfac3 = rsttemp!Impfac3
            Adodbf.Recordset!factura4 = rsttemp!factura4
            Adodbf.Recordset!Impfac4 = rsttemp!Impfac4
            Adodbf.Recordset!factura5 = rsttemp!factura5
            Adodbf.Recordset!Impfac5 = rsttemp!Impfac5
            Adodbf.Recordset!factura6 = rsttemp!factura6
            Adodbf.Recordset!Impfac6 = rsttemp!Impfac6
            Adodbf.Recordset!factura7 = rsttemp!factura7
            Adodbf.Recordset!Impfac7 = rsttemp!Impfac7
            Adodbf.Recordset!factura8 = rsttemp!factura8
            Adodbf.Recordset!Impfac8 = rsttemp!Impfac8
            Adodbf.Recordset!factura9 = rsttemp!factura9
            Adodbf.Recordset!Impfac9 = rsttemp!Impfac9
            Adodbf.Recordset!factura10 = rsttemp!factura10
            Adodbf.Recordset!Impfac10 = rsttemp!Impfac10
            Adodbf.Recordset!Importado = False
            Adodbf.Recordset.Update
            rsttemp.MoveNext
      Wend
      rsttemp.Close
      nPed = nPed + 1
      rs.MoveNext
  Wend
  Adodbf.Recordset.Close
  Stb1.SimpleText = cMenAnt
  MsgBox "SE ENVIARON " & CStr(nPed) & " PEDIDOS RECIBIDOS DEL " & txtFecha(0).Text & " AL " & txtFecha(1).Text, vbInformation
  Exit Sub
ERROR:
   If Err.Number <> 32755 Then  'Numero de error al presionar el boton cancelar
      MsgBox Err.Description
   End If
   Stb1.SimpleText = cMenAnt
End Sub

