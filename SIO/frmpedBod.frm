VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmpedBod 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Menu de pedidos por proveedor..."
   ClientHeight    =   8535
   ClientLeft      =   45
   ClientTop       =   255
   ClientWidth     =   11880
   Icon            =   "frmpedBod.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8535
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin ComctlLib.StatusBar stb1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   29
      Top             =   8160
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   661
      Style           =   1
      SimpleText      =   "                                                                  Click en el encabezado ordena los datos en base a la columna"
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoDbf 
      Height          =   330
      Left            =   3000
      Top             =   7080
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
   Begin VB.PictureBox cRpt 
      Height          =   480
      Left            =   720
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   37
      Top             =   3960
      Width           =   1200
   End
   Begin VB.PictureBox PicInf 
      Align           =   2  'Align Bottom
      BackColor       =   &H00808080&
      Height          =   650
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   11820
      TabIndex        =   15
      Top             =   7515
      Width           =   11880
      Begin VB.CommandButton Cmdmoverse 
         Height          =   450
         Index           =   10
         Left            =   3480
         Picture         =   "frmpedBod.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Pedido con costos por producto"
         Top             =   70
         Width           =   500
      End
      Begin VB.CommandButton Cmdmoverse 
         Height          =   450
         Index           =   9
         Left            =   4080
         Picture         =   "frmpedBod.frx":0974
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Pedidos pendientes de recibir"
         Top             =   70
         Width           =   500
      End
      Begin VB.CommandButton Cmdmoverse 
         Height          =   450
         Index           =   8
         Left            =   4680
         Picture         =   "frmpedBod.frx":0EA6
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Vista preeliminar de pedidos comprometidos $"
         Top             =   70
         Width           =   500
      End
      Begin VB.CommandButton Cmdmoverse 
         Height          =   450
         Index           =   7
         Left            =   5880
         Picture         =   "frmpedBod.frx":0FA8
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Vista preeliminar de la nota de entrada del pedido seleccionado"
         Top             =   70
         Width           =   500
      End
      Begin VB.CommandButton Cmdmoverse 
         Height          =   450
         Index           =   6
         Left            =   2880
         Picture         =   "frmpedBod.frx":14DA
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Actualizar (Obtiene nuevamente los datos) "
         Top             =   70
         Width           =   500
      End
      Begin VB.CommandButton Cmdmoverse 
         Height          =   450
         Index           =   0
         Left            =   480
         Picture         =   "frmpedBod.frx":15DC
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Ir al primero"
         Top             =   70
         Width           =   500
      End
      Begin VB.CommandButton Cmdmoverse 
         Height          =   450
         Index           =   1
         Left            =   1680
         Picture         =   "frmpedBod.frx":174E
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Búsqueda por clave del proveedor"
         Top             =   70
         Width           =   500
      End
      Begin VB.CommandButton Cmdmoverse 
         Height          =   450
         Index           =   2
         Left            =   5280
         Picture         =   "frmpedBod.frx":1848
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Vista preeliminar de la nota de entrada de los pedidos en el rango especificado"
         Top             =   70
         Width           =   500
      End
      Begin VB.CommandButton Cmdmoverse 
         Height          =   450
         Index           =   3
         Left            =   1080
         Picture         =   "frmpedBod.frx":197E
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Ir al último"
         Top             =   70
         Width           =   500
      End
      Begin VB.CommandButton Cmdmoverse 
         Height          =   450
         Index           =   4
         Left            =   2280
         Picture         =   "frmpedBod.frx":1AF0
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Búsqueda por clave del pedido"
         Top             =   70
         Width           =   500
      End
      Begin VB.CommandButton Cmdmoverse 
         Height          =   450
         Index           =   5
         Left            =   6480
         Picture         =   "frmpedBod.frx":1BEA
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Vista preeliminar de los pedidos enviados a Carbonera"
         Top             =   70
         Width           =   500
      End
      Begin MSComDlg.CommonDialog cmdlg 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
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
         Left            =   7680
         TabIndex        =   16
         Top             =   165
         Width           =   3255
      End
   End
   Begin MSAdodcLib.Adodc AdoPedidos 
      Height          =   330
      Left            =   120
      Top             =   7080
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
      TabIndex        =   13
      Top             =   1080
      Width           =   11655
      Begin VB.CheckBox chkFecGen 
         Caption         =   "Por fecha de"
         Height          =   255
         Left            =   6840
         TabIndex        =   26
         ToolTipText     =   "Cambia el criterio del filtro de fecha (Elaboración/Recepción)"
         Top             =   120
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.ComboBox cmbProved 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   27
         Top             =   360
         Width           =   6615
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   330
         Index           =   0
         Left            =   8400
         TabIndex        =   35
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         CustomFormat    =   "dd/MM/yyyy "
         Format          =   61341699
         CurrentDate     =   37257
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   330
         Index           =   1
         Left            =   10080
         TabIndex        =   36
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         CustomFormat    =   "dd/MM/yyyy "
         Format          =   61341699
         CurrentDate     =   37257
      End
      Begin VB.Label lbltipfec 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Elaboración"
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
         Left            =   6840
         TabIndex        =   30
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label LBLETIQUETAS 
         Alignment       =   2  'Center
         Caption         =   "Final"
         Height          =   255
         Index           =   1
         Left            =   9960
         TabIndex        =   11
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label LBLETIQUETAS 
         Alignment       =   2  'Center
         Caption         =   "Inicial"
         Height          =   255
         Index           =   0
         Left            =   8280
         TabIndex        =   17
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label lblProv 
         Alignment       =   2  'Center
         Caption         =   "Proveedor:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   6615
      End
   End
   Begin VB.Frame cmdDespla 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11655
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "&Recibir pedido"
         Height          =   495
         Index           =   2
         Left            =   1080
         TabIndex        =   33
         Top             =   240
         Width           =   950
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Desmarcar por fecha"
         Height          =   495
         Index           =   9
         Left            =   5880
         TabIndex        =   28
         Top             =   240
         Width           =   950
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Nuevo ped Indirecto"
         Height          =   495
         Index           =   8
         Left            =   4920
         TabIndex        =   5
         Top             =   240
         Width           =   950
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "&Cerrar"
         Height          =   495
         Index           =   5
         Left            =   10680
         TabIndex        =   6
         ToolTipText     =   "Regresar a la pantalla principal"
         Top             =   240
         Width           =   950
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "&Enviar Ped Carbonera"
         Height          =   495
         Index           =   7
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Exportar pedidos seleccionados en el rango especificado"
         Top             =   240
         Width           =   950
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "&Llegadas Carbonera"
         Height          =   495
         Index           =   6
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Importar llegadas de pedidos recibidos en Bodega Carbonera"
         Top             =   240
         Width           =   950
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "&Importar Pedofi"
         Height          =   495
         Index           =   3
         Left            =   6840
         TabIndex        =   9
         ToolTipText     =   "Importa pedido de oficinas centrales"
         Top             =   240
         Width           =   950
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "&Exportar llegadas"
         Enabled         =   0   'False
         Height          =   495
         Index           =   4
         Left            =   7800
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Actualiza el costo de los pedidos que forman el pedido por proveedor"
         Top             =   240
         Width           =   950
      End
      Begin VB.ComboBox cmbOrden 
         Height          =   315
         Left            =   8880
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "&Ver Pedido Confirmado"
         Height          =   495
         Index           =   1
         Left            =   2040
         TabIndex        =   2
         ToolTipText     =   "Consultar pedido"
         Top             =   240
         Width           =   950
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "&Nuevo Ped     Mixto"
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "Generar pedido por proveedor"
         Top             =   240
         Width           =   950
      End
      Begin VB.Label lblFiltro 
         Alignment       =   2  'Center
         Caption         =   "Filtro"
         Height          =   255
         Left            =   8880
         TabIndex        =   12
         Top             =   120
         Width           =   1695
      End
   End
   Begin MSDataGridLib.DataGrid dbgrdPed 
      Bindings        =   "frmpedBod.frx":211C
      Height          =   5145
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   9075
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   14737632
      ForeColor       =   8388608
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
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "pp_proveedor"
         Caption         =   "CLAVE PROV."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "pp_pedido"
         Caption         =   "  CLAVE PED."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "PP_SUCURSAL"
         Caption         =   "TDA."
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
         DataField       =   "pp_fechaGen"
         Caption         =   "      FECHA ELAB."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/mm/yyyy hh:mm AM/PM"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "pp_fecConfirma"
         Caption         =   "         FEC. CONFIRMA"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/mm/yyyy hh:mm AM/PM"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "pp_fecrecibe"
         Caption         =   "      FEC. RECEPCION"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/mm/yyyy hh:mm AM/PM"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "pp_notent"
         Caption         =   " NOTA ENTRADA"
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
      BeginProperty Column07 
         DataField       =   "pp_pedind"
         Caption         =   "  IND."
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         BeginProperty Column00 
            Alignment       =   2
            ColumnWidth     =   1154.835
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   1154.835
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   540.284
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   1904.882
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   2055.118
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   2250.142
         EndProperty
         BeginProperty Column06 
            Alignment       =   2
            ColumnWidth     =   1379.906
         EndProperty
         BeginProperty Column07 
            Alignment       =   2
            ColumnWidth     =   555.024
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoSugDbf 
      Height          =   330
      Left            =   5520
      Top             =   7080
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
End
Attribute VB_Name = "frmpedBod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cCond As String     'Condicion del filtro del grid
Private ccondrpt As String  'Condicion del filtro del rpt
Private rstProv As ADODB.Recordset
Private ntext As Integer
Private cFecha As String

Private Sub AdoPedidos_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
 On Error Resume Next
 cmbProved.Text = AdoPedidos.Recordset!NOMPROVE & "  " & AdoPedidos.Recordset!prove
End Sub

Private Sub dtpFecha_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys vbTab
End Sub

Private Sub dtpFecha_LostFocus(Index As Integer)
On Error GoTo Error:
 cCamFec = IIf(chkFecGen.Value = 1, "pp_fechaGen", "pp_fecrecibe")
 If chkFecGen.Value = 1 Then
    cFecha = " AND " & cCamFec & " >= '" & dtpFecha(0).Value & "' AND " & cCamFec & " <= DATEADD(day,1,'" & dtpFecha(1).Value & "')"
 Else
    cFecha = " AND " & cCamFec & " >= '" & dtpFecha(0).Value & "' AND " & cCamFec & " <= '" & dtpFecha(1).Value & "'"
 End If
 CADENA = "SELECT * from pedprove,catprov where pp_proveedor = catprov.prove and " & cCond & cFecha
 AdoPedidos.RecordSource = CADENA & "ORDER BY pp_fechagen DESC"
 AdoPedidos.Refresh
 lblInfo.Caption = "Numero de Pedidos:" + Str(AdoPedidos.Recordset.RecordCount)
 For N = 0 To 6  'Si esta vacio el recordset desactivo las opciones
   Cmdmoverse(N).Enabled = Not (AdoPedidos.Recordset.BOF And AdoPedidos.Recordset.EOF)
 Next
  cmdOpcion(1).Enabled = Not (AdoPedidos.Recordset.BOF And AdoPedidos.Recordset.EOF)
  cmdOpcion(7).Enabled = Not (AdoPedidos.Recordset.BOF And AdoPedidos.Recordset.EOF)
  cmdOpcion(4).Enabled = Not (AdoPedidos.Recordset.BOF And AdoPedidos.Recordset.EOF)
  Me.cmdOpcion(2).Enabled = (tipotienda <> 1)
  Me.cmdOpcion(3).Enabled = (tipotienda <> 1)
  Me.cmdOpcion(4).Enabled = (tipotienda <> 1)
  Me.cmdOpcion(0).Enabled = (Nivel <> "P" And tipotienda = 1)
  Me.cmdOpcion(6).Enabled = (Nivel <> "P" And tipotienda = 1)
  Me.cmdOpcion(7).Enabled = (Nivel <> "P" And tipotienda = 1)
  'Me.cmdopcion(8).Enabled = (Nivel <> "P" And tipotienda = 1)
  Me.cmdOpcion(9).Enabled = (Nivel <> "P" And tipotienda = 1)
 Exit Sub
Error:
  MsgBox Err.Description

End Sub

Private Sub chkFecGen_Click()
lbltipfec.Caption = IIf(chkFecGen.Value = 1, "Elaboración", "Recepción")
End Sub

Private Sub cmborden_DblClick()
  SendKeys "{TAB}"
End Sub

Private Sub cmbOrden_GotFocus()
RESP = SendMessageLong(cmbOrden.hwnd, &H14F, True, 1)
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
    cCond = "pp_cancelado = 0"
    ccondrpt = "{PEDPROVE.pp_cancelado} = 0 AND {PEDPROVE.pp_confirma} = 1"
Case 1 'pendientes por recibir
    cCond = "Pp_confirma = 1  AND Pp_recibe = 0 AND pp_cancelado = 0"
    ccondrpt = "{PEDPROVE.Pp_confirma} = 1  AND {PEDPROVE.Pp_recibe} = 0 AND {PEDPROVE.Pp_cancelado} = 0 "
Case 2 'Recibidos
    cCond = "Pp_recibe = 1 AND pp_cancelado = 1"
    ccondrpt = "{PEDPROVE.Pp_recibe} = 1 AND {PEDPROVE.Pp_cancelado} = 0 "
End Select
cCamFec = IIf(chkFecGen.Value = 1, "pp_fechaGen", "pp_fecrecibe")
cFecha = " AND " & cCamFec & " >= '" & dtpFecha(0).Value & "' AND " & cCamFec & " <= DATEADD(day,1,'" & dtpFecha(1).Value & "')"
CADENA = "SELECT * from pedprove,catprov where pp_proveedor = catprov.prove and " & cCond & cFecha
AdoPedidos.RecordSource = CADENA & " ORDER BY pp_proveedor"
AdoPedidos.Refresh
lblInfo.Caption = "Numero de pedidos:" + Str(AdoPedidos.Recordset.RecordCount)
For N = 0 To 6
   Cmdmoverse(N).Enabled = Not (AdoPedidos.Recordset.BOF And AdoPedidos.Recordset.EOF)
Next
Me.cmdOpcion(2).Enabled = (tipotienda <> 1)
Me.cmdOpcion(3).Enabled = (tipotienda = 2)
Me.cmdOpcion(4).Enabled = (tipotienda = 2)
Me.cmdOpcion(0).Enabled = (Nivel <> "P" And tipotienda = 1)
Me.cmdOpcion(6).Enabled = (Nivel <> "P" And tipotienda = 1)
Me.cmdOpcion(7).Enabled = (Nivel <> "P" And tipotienda = 1)
Me.cmdOpcion(9).Enabled = (Nivel <> "P" And tipotienda = 1)
End Sub

Private Sub cmbproved_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    SendKeys vbTab
End If
End Sub

Private Sub cmbProved_LostFocus()
If Trim(cmbProved.Text) <> "" Then
   cCamFec = IIf(chkFecGen.Value = 1, "pp_fechaGen", "pp_fecrecibe")
   cFecha = " AND " & cCamFec & " >= '" & dtpFecha(0).Value & "' AND " & cCamFec & " <= DATEADD(day,1,'" & dtpFecha(1).Value & "')"
   CADENA = "SELECT * FROM pedprove,catprov WHERE pp_proveedor = catprov.prove AND pp_proveedor = '" & Trim(Mid(cmbProved.Text, Len(cmbProved.Text) - 5)) & "' AND " & cCond & cFecha
   AdoPedidos.RecordSource = CADENA
Else
   cCamFec = IIf(chkFecGen.Value = 1, "pp_fechaGen", "pp_fecrecibe")
   cFecha = " AND " & cCamFec & " >= '" & dtpFecha(0).Value & "' AND " & cCamFec & " <= DATEADD(day,1,'" & dtpFecha(1).Value & "')"
   CADENA = "SELECT * FROM pedprove,catprov WHERE pp_proveedor = catprov.prove AND " & cCond & cFecha
   AdoPedidos.RecordSource = CADENA
End If
AdoPedidos.Refresh
lblInfo.Caption = "Numero de Pedidos:" + Str(AdoPedidos.Recordset.RecordCount)
End Sub

Private Sub cmdMoverse_Click(Index As Integer)
'On Error GoTo Error:
Select Case Index
Case 0  'Primer registro
    AdoPedidos.Recordset.MoveFirst
Case 1  'Busqueda por clave de proveedor
    cCve = InputBox("Introduzca la clave del proveedor a buscar", "Introducir clave del proveedor")
    If IsNull(Trim(cCve)) Or Trim(cCve) = "" Then Exit Sub
    Antes = dbgrdPed.Bookmark
    AdoPedidos.Recordset.MoveFirst
    AdoPedidos.Recordset.Find "pp_proveedor = '" & Trim(cCve) & "'"
    If AdoPedidos.Recordset.EOF Then
        MsgBox "LA CLAVE " & cCve & " NO SE ENCUENTRA EN LOS PEDIDOS " + IIf(cmbOrden.Text = "TODOS", "", cmbOrden.Text), vbExclamation
        dbgrdPed.Bookmark = Antes
    End If
Case 2  'Presentacion preeelimar de la nota de entrada
       crpt.ReportFileName = App.Path & "\PrEntPed.rpt"
       crpt.WindowTitle = "Nota de entrada de las llegadas del " & Me.dtpFecha(0).Value & " al " & Me.dtpFecha(1).Value
       crpt.Connect = cCadConex
       crpt.SQLQuery = "SELECT PEDPROVE.pp_proveedor, PEDPROVE.pp_pedido, PEDPROVE.pp_fechagen, PEDPROVE.pp_fecconfirma, PEDPROVE.pp_fecrecibe, " & _
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
                            "PEDPROVE.pp_fecrecibe >= '" & Format(dtpFecha(0).Value, "yyyy-dd-mm") & "' AND PEDPROVE.pp_fecrecibe <= '" & Format(dtpFecha(1).Value, "yyyy-dd-mm") & "' " & Chr(13) & _
                       "ORDER BY PEDPROVE.pp_pedido ASC, " & _
                            "TFPRODUC.DESCRIPC ASC, " & _
                            "TFPRODUC.CONTENID ASC"
       'MsgBox Mid(cRpt.SQLQuery, 450)
       crpt.Action = 1
Case 3  'Ultimo
    AdoPedidos.Recordset.MoveLast
Case 4  'Busqueda por folio de pedido
    cCve = InputBox("Introduzca la clave del pedido a buscar", "Introducir clave del pedido")
    If IsNull(Trim(cCve)) Or Trim(cCve) = "" Then Exit Sub
    Antes = dbgrdPed.Bookmark
    AdoPedidos.Recordset.MoveFirst
    AdoPedidos.Recordset.Find "pp_pedido = '" & Trim(cCve) & "'"
    If AdoPedidos.Recordset.EOF Then
        MsgBox "LA CLAVE " & cCve & " NO SE ENCUENTRA EN LOS PEDIDOS " + IIf(cmbOrden.Text = "TODOS", "", cmbOrden.Text), vbExclamation
        dbgrdPed.Bookmark = Antes
    End If
Case 5
       cMensaje = stb1.SimpleText
       stb1.SimpleText = Space(90) + "Espere un momento generando reporte..."
       stb1.Refresh
       cFecharpt = " AND {PEDPROVE.pp_fechagen} >= Date(" & CStr(Year(dtpFecha(0).Value)) & "," & CStr(Month(dtpFecha(0).Value)) & "," & CStr(Day(dtpFecha(0).Value)) & ") AND {PEDPROVE.pp_fechagen} <= Date(" & CStr(Year(dtpFecha(1).Value)) & "," & CStr(Month(dtpFecha(1).Value)) & "," & CStr(Day(dtpFecha(1).Value)) & ") "
       crpt.Connect = cCadConex
       crpt.ReportFileName = App.Path & "\PedProv.rpt"
       crpt.WindowTitle = "Reporte de pedidos por proveedor"
       crpt.Formulas(0) = "PEDIDO = 'PEDIDOS POR PROVEEDOR GENERADOS EN OFICINAS " & IIf(cmbOrden.Text = "TODOS", "", cmbOrden.Text) & " DEL " & Trim(dtpFecha(0).Value) & " AL " & Trim(dtpFecha(1).Value) & " '"
       crpt.Formulas(1) = "FORMSELEC = " & ccondrpt & cFecharpt
       'MsgBox cRpt.Formulas(1)
       crpt.Action = 1
       stb1.SimpleText = cMensaje
       stb1.Refresh
Case 6
       Me.AdoPedidos.Refresh
       lblInfo.Caption = "Numero de Pedidos:" + Str(AdoPedidos.Recordset.RecordCount)
Case 7
       crpt.ReportFileName = App.Path & "\PrEntPed.rpt"
       crpt.WindowTitle = "Nota de entrada del pedido " & AdoPedidos.Recordset!pp_pedido
       crpt.Connect = cCadConex
       crpt.SQLQuery = "SELECT PEDPROVE.pp_proveedor, PEDPROVE.pp_pedido, PEDPROVE.pp_fechagen, PEDPROVE.pp_fecconfirma, PEDPROVE.pp_fecrecibe, " & _
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
       'MsgBox Mid(crpt.SQLQuery, 400)
       crpt.Action = 1
Case 8
      Enca = "Pedidos por proveedor del " & Me.dtpFecha(0).Value & " AL " & Me.dtpFecha(1).Value
      crpt.ReportFileName = App.Path & "\PrvMonto.rpt"
      crpt.WindowTitle = Enca
      crpt.SQLQuery = "SELECT PEDPROVE.pp_pedido, PEDPROVE.pp_fechagen, PEDPROVE.pp_fecent, PEDPROVE.pp_pedback, " & _
                              "CATPROV.PROVE, CATPROV.NOMPROVE, CATPROV.visita, " & _
                              "DETALLEGLOBAL.dg_producto, DETALLEGLOBAL.dg_cantsol, DETALLEGLOBAL.dg_cantsolp, DETALLEGLOBAL.dg_costo " & _
                       "FROM pitico.dbo.PEDPROVE PEDPROVE, " & _
                              "pitico.dbo.CATPROV CATPROV, " & _
                              "pitico.dbo.DETALLEGLOBAL DETALLEGLOBAL " & _
                       "WHERE PEDPROVE.pp_proveedor = CATPROV.PROVE AND " & _
                              "PEDPROVE.pp_pedido = DETALLEGLOBAL.dg_pedido AND " & _
                              "PEDPROVE.pp_fechagen >= '" & Format(dtpFecha(0).Value, "yyyy-dd-mm") & "' AND PEDPROVE.pp_fechagen <= '" & Format(DateAdd("d", 1, dtpFecha(1).Value), "yyyy-dd-mm") & "' " & _
                              "AND PEDPROVE.pp_pedback IS NULL"
      crpt.Formulas(0) = "encab = '" & UCase(Enca) & "'"
      crpt.Formulas(1) = ""
      If MsgBox("DESEAS VER EL REPORTE EN RESUMEN", vbQuestion + vbYesNo, "Modo de ver reporte") = vbYes Then
         crpt.SectionFormat(1) = "GH2;F;F;F;X;X;X;X"
      Else
         crpt.SectionFormat(1) = "GH2;T;F;F;X;X;X;X"
      End If
      crpt.Action = 1
Case 9
      If MsgBox("MOSTRAR SOLO EL PROVEEDOR SELECCIONADO", vbYesNo + vbQuestion + vbDefaultButton2, "Filtro") = vbYes Then
         ccondrpt = "{PEDPROVE.Pp_confirma} = 1  AND {PEDPROVE.Pp_recibe} = 0 AND {PEDPROVE.Pp_cancelado} = 0 AND {PEDPROVE.pp_fechagen} >= date(" & Format(dtpFecha(0).Value, "YYYY,MM,DD") & ") AND {PEDPROVE.pp_fechagen} <= date(" & Format(dtpFecha(1).Value, "YYYY,MM,DD") & ") and {PEDPROVE.pp_proveedor} = '" & AdoPedidos.Recordset!pp_proveedor & "'"
      Else
         ccondrpt = "{PEDPROVE.Pp_confirma} = 1  AND {PEDPROVE.Pp_recibe} = 0 AND {PEDPROVE.Pp_cancelado} = 0 AND {PEDPROVE.pp_fechagen} >= date(" & Format(dtpFecha(0).Value, "YYYY,MM,DD") & ") AND {PEDPROVE.pp_fechagen} <= date(" & Format(dtpFecha(1).Value, "YYYY,MM,DD") & ")"
      End If
      cEnca = "PEDIDOS PENDIENTES DE RECIBIR POR PROVEEDOR"
      Me.crpt.WindowTitle = cEnca
      crpt.ReportFileName = App.Path & "\pedprov.rpt"
      crpt.Formulas(0) = "FORMSELEC = " & ccondrpt
      crpt.Formulas(1) = "PEDIDO= '" & cEnca & "'"
      crpt.Action = 1
Case 10
      crpt.ReportFileName = App.Path & "\pedproco.rpt"
      crpt.WindowTitle = "Reporte de costos por producto"
      crpt.SQLQuery = "SELECT PEDPROVE.pp_proveedor, PEDPROVE.pp_pedido, PEDPROVE.pp_fechagen, PEDPROVE.pp_fecent, PEDPROVE.pp_fecconfirma, PEDPROVE.pp_fecrecibe, PEDPROVE.pp_observa, " & _
                             "DETALLEGLOBAL.dg_producto, DETALLEGLOBAL.dg_cantsol, DETALLEGLOBAL.dg_cantsolp, DETALLEGLOBAL.dg_cantreal, DETALLEGLOBAL.dg_cantrealp, DETALLEGLOBAL.dg_costo, " & _
                             "NOTAENTRADA.factura1, NOTAENTRADA.impfac1, NOTAENTRADA.factura2, NOTAENTRADA.impfac2, NOTAENTRADA.factura3, NOTAENTRADA.impfac3, NOTAENTRADA.factura4, NOTAENTRADA.impfac4, NOTAENTRADA.factura5, NOTAENTRADA.impfac5, NOTAENTRADA.factura6, NOTAENTRADA.impfac6, NOTAENTRADA.factura7, NOTAENTRADA.impfac7, NOTAENTRADA.factura8, NOTAENTRADA.impfac8, NOTAENTRADA.factura9, NOTAENTRADA.impfac9, NOTAENTRADA.factura10, NOTAENTRADA.impfac10, NOTAENTRADA.factura11, NOTAENTRADA.impfac11, NOTAENTRADA.factura12, NOTAENTRADA.impfac12, NOTAENTRADA.factura13, NOTAENTRADA.impfac13, NOTAENTRADA.factura14, NOTAENTRADA.impfac14, NOTAENTRADA.factura15, NOTAENTRADA.impfac15, NOTAENTRADA.factura16, NOTAENTRADA.impfac16, NOTAENTRADA.factura17, NOTAENTRADA.impfac17, NOTAENTRADA.factura18, NOTAENTRADA.impfac18, NOTAENTRADA.factura19, NOTAENTRADA.impfac19, NOTAENTRADA.factura20, NOTAENTRADA.impfac20, " & _
                             "CATPROV.NOMPROVE, CATPROV.frecuencia, TFPRODUC.DESCRIPC, TFPRODUC.CONTENID, TFPRODUC.MEDIDA, TFPRODUC.PAQUETES, TFPRODUC.precosto " & Chr(13) & _
                      "FROM pitico.dbo.TFPRODUC TFPRODUC, " & Chr(13) & _
                             "pitico.dbo.DETALLEGLOBAL DETALLEGLOBAL, " & _
                             "pitico.dbo.PEDPROVE PEDPROVE, " & _
                             "pitico.dbo.NOTAENTRADA NOTAENTRADA , " & _
                             "pitico.dbo.CATPROV CATPROV " & Chr(13) & _
                      "WHERE PEDPROVE.pp_pedido = DETALLEGLOBAL.dg_pedido AND " & _
                             "PEDPROVE.pp_pedido = NOTAENTRADA.pedido AND " & _
                             "PEDPROVE.pp_proveedor = CATPROV.PROVE AND " & _
                             "DETALLEGLOBAL.dg_producto = TFPRODUC.CONSEC AND " & _
                             "PEDPROVE.pp_pedido = '" & AdoPedidos.Recordset!pp_pedido & "' AND " & _
                             "DETALLEGLOBAL.dg_cantreal > 0 " & Chr(13) & _
                      "ORDER BY PEDPROVE.pp_pedido ASC, TFPRODUC.DESCRIPC ASC, TFPRODUC.CONTENID ASC "
          
      crpt.Action = 1
End Select
Exit Sub
Error:
   MsgBox Err.Description
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
Dim rsttemp As ADODB.Recordset
Dim cArch As String
Dim nPed As Integer
Dim SUC As String
Dim cnFoxPro As ADODB.Connection
'On Error GoTo error
lNvoPedprove = False
Select Case Index
  Case 0  'Nuevo pedido solo oficinas
       If tipotienda = 1 Then
          lNvoPedprove = True
          cModo = "": lpprov = False
          nOp = 1
          frmPedProv.Caption = "Consultar pedidos por tienda para formar uno por proveedor"
          frmPedProv.Show
       End If
  Case 1  'Modificar pedido
       nOp = 0
       cModo = "VERCONF"
       frmPedProv.Caption = "Consultar pedido por proveedor confirmado"
       frmPedProv.txtcampos(9).Text = Me.AdoPedidos.Recordset!PP_SUCURSAL
       frmPedProv.Show
       SendKeys "{TAB}": SendKeys "{TAB}"
  Case 2  'Recibir pedido solo carbonera
       cModo = "RECIBIR"
       nOp = 0
       frmPedProvB.Caption = "Recibir pedido por proveedor"
       frmPedProvB.Show
       SendKeys "{TAB}"
  Case 3  'Importa pedido de oficinas centrales
       ImpPedOfi
  Case 4
       ExpPedCar
  Case 5  'Salir del modulo de pedidos
       Unload Me
  Case 6 'Importar llegadas de pedidos recibidos en bodega Carbonera
       ImpPedCar
  Case 7
        MsgBox "CON ESTE PROCESO YA NO PODRA MODIFICAR LOS PEDIDOS Y SOLO SE ENVIARAN A LA BODEGA, LOS PEDIDOS QUE CORRESPONDAN AL FILTRO DE FECHAS SELECCIONADAS EN LA PANTALLA !!!", vbInformation
       If MsgBox("DESEA GENERAR Y PREPARAR EL ARCHIVO DE PEDIDOS A CARBONERA ", vbYesNo + vbQuestion) = vbYes Then
           MsgBox "Se Sugiere la Siguiente Ruta : p:\paso , para el archivo pedofi", vbInformation
           ExpPedOfi
       End If
  Case 8  'Nuevo pedido indirecto no se necesitan pedidos de tienda para formar uno por proveedor
       lpprov = True
       cModo = ""
       nOp = 1
       lpprov = True  'Para saber cuando es un pedido Indirecto
       frmPedProv.Caption = "Generar nuevo pedido Indirecto por proveedor"
       frmPedProv.Show
  Case 9
       'se toman las fechas inicial que tiene por default, es solo un dia
       If MsgBox("ESTE PROCESO , SOLO SE PUEDE EJECUTAR UNA VEZ POR DIA, ", vbYesNo + vbQuestion) = vbYes Then
           fecha1 = CDate(dtpFecha(0).Value)
           DiaP = Day(fecha1)
           mesP = Month(fecha1)
           anioP = Year(fecha1)
           CADENA = "update pedprove set pp_enviado = 0 where day(pp_fechagen) = " & DiaP & " and month(pp_fechagen) = " & mesP & " and year(pp_fechagen)  = " & anioP
           cn.Execute CADENA
           AdoPedidos.Refresh
           MsgBox "Proceso terminado...", vbInformation
       End If
End Select
Exit Sub
Error:
   MsgBox Err.Description
End Sub

Private Sub cmdopcion_GotFocus(Index As Integer)
If Index = 0 Then Unload frmAreaRecibo
End Sub


Private Sub dbgrdPed_DblClick()
  If Not (AdoPedidos.Recordset.BOF And AdoPedidos.Recordset.EOF) Then cmdOpcion_Click 1
End Sub

Private Sub DbGrdped_HeadClick(ByVal ColIndex As Integer)
  stb1.SimpleText = Space(65) + "Espere un momento ordenando Pedidos por " & dbgrdPed.Columns(ColIndex).Caption
  cCamFec = IIf(chkFecGen.Value = 1, "pp_fechaGen", "pp_fecrecibe")
  cFecha = " AND " & cCamFec & " >= '" & dtpFecha(0).Value & "' AND " & cCamFec & " <= DATEADD(day,1,'" & dtpFecha(1).Value & "')"
  CADENA = "SELECT * from pedprove,catprov where pp_proveedor = catprov.prove and " & cCond & cFecha
  AdoPedidos.RecordSource = CADENA & "ORDER BY " & dbgrdPed.Columns(ColIndex).DataField
  AdoPedidos.Refresh
  lblInfo.Caption = "Numero de Pedidos:" + Str(AdoPedidos.Recordset.RecordCount)

  stb1.SimpleText = Space(85) + "Pedidos ordenandos por " & dbgrdPed.Columns(ColIndex).Caption
End Sub

Private Sub dbgrdPed_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If Not (AdoPedidos.Recordset.BOF And AdoPedidos.Recordset.EOF) Then cmdOpcion_Click 1
End If
End Sub



Private Sub Form_Activate()
On Error Resume Next
  Unload frmAreaRecibo
  Me.dbgrdPed.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 119 Then frmCalc.Show
End Sub

Private Sub Form_Load()
Dim rs As ADODB.Recordset
  lpprov = False   'Por default es pedido Mixto (por tiendas)
  cmbOrden.AddItem "TODOS"
  cmbOrden.AddItem "PENDIENTES DE RECIBIR"
  cmbOrden.AddItem "RECIBIDOS"
  cmbOrden.ListIndex = 0
  Set rs = New ADODB.Recordset
  rs.Open "CATPROV", cn, adOpenKeyset, adLockOptimistic, adCmdTable
  While Not rs.EOF
     cmbProved.AddItem rs!NOMPROVE & "   " & rs!prove
     rs.MoveNext
  Wend
  cCond = "pp_cancelado = 0"                ' Filtro por default todos los pedidos
  ccondrpt = "{PEDPROVE.pp_cancelado} = 0 AND {PEDPROVE.pp_confirma} = 1 "  ' Filtro por default del RPT
  If dtpFecha(0).Value = "01/01/02" Then dtpFecha(0).Value = Format(date, "DD/mm/yyyy")
  If dtpFecha(1).Value = "01/01/02" Then dtpFecha(1).Value = Format(date, "DD/mm/yyyy")
  cOper = IIf(Month(dtpFecha(0).Value) = Month(dtpFecha(1).Value), " AND ", " OR ")
  cFecha = " AND (month(pp_fechagen) >= " & Month(dtpFecha(0).Value) & " and (day(pp_fechagen) > = " & Day(dtpFecha(0).Value) & cOper & "(day(pp_fechagen)<= " & Day(dtpFecha(1).Value) & " and month(pp_fechagen)<= " & Month(dtpFecha(1).Value) & ")) and year(pp_fechagen)>= " & Year(dtpFecha(0).Value) & " and year(pp_fechagen)<= " & Year(dtpFecha(1).Value) & ")"
  'Cargo todos los pedidos
  AdoPedidos.ConnectionString = cCadConex
  AdoPedidos.CommandType = adCmdText
  AdoPedidos.LockType = adLockOptimistic
  CADENA = "select * from pedprove,catprov where pp_proveedor = catprov.prove and " & cCond & cFecha
  AdoPedidos.RecordSource = CADENA & " ORDER BY pp_fechagen DESC"
  AdoPedidos.Refresh
  For N = 0 To 6
     Cmdmoverse(N).Enabled = Not (AdoPedidos.Recordset.BOF And AdoPedidos.Recordset.EOF)
  Next
  cmdOpcion(1).Enabled = Not (AdoPedidos.Recordset.BOF And AdoPedidos.Recordset.EOF)
  cmdOpcion(7).Enabled = Not (AdoPedidos.Recordset.BOF And AdoPedidos.Recordset.EOF)
  cmdOpcion(4).Enabled = Not (AdoPedidos.Recordset.BOF And AdoPedidos.Recordset.EOF)
  lblInfo.Caption = "Numero de pedidos:" + Str(AdoPedidos.Recordset.RecordCount)
  Me.cmdOpcion(2).Enabled = (tipotienda <> 1)
  Me.cmdOpcion(3).Enabled = (tipotienda = 2)
  Me.cmdOpcion(4).Enabled = (tipotienda = 2)
  Me.cmdOpcion(0).Enabled = (Nivel <> "P" And tipotienda = 1)
  Me.cmdOpcion(6).Enabled = (Nivel <> "P" And tipotienda <> 2)
  Me.cmdOpcion(7).Enabled = (Nivel <> "P" And tipotienda = 1)
  'Me.cmdopcion(8).Enabled = (Nivel <> "P" And tipotienda = 1)
  Me.cmdOpcion(9).Enabled = (Nivel <> "P" And tipotienda = 1)
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmAreaRecibo.Show
End Sub

'Importar llegadas de pedidos por proveedor de Bodega Carbonera
Private Sub ImpPedCar()
Dim nDias As Integer
Dim FolPed As String
Dim rsttemp As ADODB.Recordset
Dim rstBus As ADODB.Recordset
Dim cArch  As String
Dim cProv As String
Dim SUC As String
Dim cnFoxPro As ADODB.Connection
Dim lExiste  As Boolean
Dim EsBack As Boolean

'On Error GoTo error:
   MenAnt = stb1.SimpleText
   Cmdlg.DialogTitle = "Abrir archivo enviado por Bodega Carbonera"
   Cmdlg.FileName = ""
   Cmdlg.CancelError = True   'Para que se genere error al hacer click en el boton cancelar
   Cmdlg.Filter = "Archivos Visual Foxpro (*.dbf) | *.dbf"
   Cmdlg.ShowOpen
   cRutArc = Cmdlg.FileName
   If cRutArc = "" Or IsNull(cRutArc) Then
      MsgBox "DEBE ESPECIFICAR UN NOMBRE DE ARCHIVO", vbExclamation
      Exit Sub
   End If
   
   For N = 1 To Len(cRutArc)
      If Mid(cRutArc, N, 1) = "\" Then nPos = N
   Next
   cruta = Mid(cRutArc, 1, nPos)
   cArch = Mid(cRutArc, nPos + 1)
   cArch = Mid(cArch, 1, Len(cArch) - 4) 'Le quito la extension porque si no marca error en la consula SQL.

   AdoDbf.CursorType = adOpenKeyset
   AdoDbf.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=PITICODBF;SourceDB=" & cruta
   AdoDbf.RecordSource = "SELECT * FROM " & cArch
   AdoDbf.Refresh
   
   If AdoDbf.Recordset.BOF And AdoDbf.Recordset.EOF Then
      MsgBox "EL ARCHIVO SELECCIONADO ESTA VACIO, NO EXISTEN DATOS PARA IMPORTAR", vbExclamation
      Exit Sub
   'CUANDO SE IMPORTO PREVIEMENTE, SE MARCA EL PRIMER REGISTRO COMO IMPORTADO
   ElseIf AdoDbf.Recordset!Importado Then
      MsgBox "EL ARCHIVO SELECCIONADO YA FUE IMPORTADO", vbInformation
      Exit Sub
   End If
   
   Open App.Path & "\LLEGACAR.TXT" For Output As #1   ' Abre el archivo para operaciones de salida.
   Print #1, Tab(20); "VIVERES Y LICORES S.A DE C.V."
   Print #1, Tab(10); "CARBONERA 1016 COL. TRINIDAD DE LAS HUERTAS"
   Print #1, "LLEGADAS DE PEDIDOS DE CARBONERA RECIBIDOS EL "; UCase(Format(date, "long date"))
   Print #1,   ' Imprime una línea en blanco en el archivo.
   Print #1, "=========================================================================================="
   Print #1, "PEDIDO        FEC. ELAB.        FEC.RECEPCION         PROVEEDOR"
   Print #1, "=========================================================================================="
   
   Set rstBus = New ADODB.Recordset  'Sirve para saber si existe o no el pedido
   Set rsttemp = New ADODB.Recordset
   Set rstins = New ADODB.Recordset
   rstBus.ActiveConnection = cn
   cNotAnt = "": nPed = 0
   AdoDbf.Recordset.MoveFirst
   While Not AdoDbf.Recordset.EOF
              'se debe buscar el maestro de la nota
              CADENA = "SELECT * FROM notaentrada WHERE clavenota =  '" & AdoDbf.Recordset!Clavenota & "'"
              rstBus.Open CADENA
              If rstBus.BOF And rstBus.EOF Then
                    'ES NECESARIO INSERTAR LA NOTA PARA QUE SE ACCESE AL DETALLE
                    'CHECAR QUE NO TENGA NULOS
                    FAC1 = IIf(Not IsNull(AdoDbf.Recordset!factura1), AdoDbf.Recordset!factura1, "0")
                    FAC2 = IIf(Not IsNull(AdoDbf.Recordset!factura2), AdoDbf.Recordset!factura2, "0")
                    fac3 = IIf(Not IsNull(AdoDbf.Recordset!Factura3), AdoDbf.Recordset!Factura3, "0")
                    fac4 = IIf(Not IsNull(AdoDbf.Recordset!factura4), AdoDbf.Recordset!factura4, "0")
                    fac5 = IIf(Not IsNull(AdoDbf.Recordset!factura5), AdoDbf.Recordset!factura5, "0")
                    fac6 = IIf(Not IsNull(AdoDbf.Recordset!factura6), AdoDbf.Recordset!factura6, "0")
                    fac7 = IIf(Not IsNull(AdoDbf.Recordset!factura7), AdoDbf.Recordset!factura7, "0")
                    fac8 = IIf(Not IsNull(AdoDbf.Recordset!factura8), AdoDbf.Recordset!factura8, "0")
                    fac9 = IIf(Not IsNull(AdoDbf.Recordset!factura9), AdoDbf.Recordset!factura9, "0")
                    fac10 = IIf(Not IsNull(AdoDbf.Recordset!factura10), AdoDbf.Recordset!factura10, "0")
                    'IMPORTES
                    IMP1 = IIf(Not IsNull(AdoDbf.Recordset!Impfac1), AdoDbf.Recordset!Impfac1, "0")
                    imp2 = IIf(Not IsNull(AdoDbf.Recordset!Impfac2), AdoDbf.Recordset!Impfac2, "0")
                    imp3 = IIf(Not IsNull(AdoDbf.Recordset!Impfac3), AdoDbf.Recordset!Impfac3, "0")
                    imp4 = IIf(Not IsNull(AdoDbf.Recordset!Impfac4), AdoDbf.Recordset!Impfac4, "0")
                    Imp5 = IIf(Not IsNull(AdoDbf.Recordset!Impfac5), AdoDbf.Recordset!Impfac5, "0")
                    imp6 = IIf(Not IsNull(AdoDbf.Recordset!Impfac6), AdoDbf.Recordset!Impfac6, "0")
                    imp7 = IIf(Not IsNull(AdoDbf.Recordset!Impfac7), AdoDbf.Recordset!Impfac7, "0")
                    imp8 = IIf(Not IsNull(AdoDbf.Recordset!Impfac8), AdoDbf.Recordset!Impfac8, "0")
                    imp9 = IIf(Not IsNull(AdoDbf.Recordset!Impfac9), AdoDbf.Recordset!Impfac9, "0")
                    imp10 = IIf(Not IsNull(AdoDbf.Recordset!Impfac10), AdoDbf.Recordset!Impfac10, "0")
                    CADENA = "INSERT INTO NOTAENTRADA(PEDIDO,CLAVENOTA,factura1,impfac1,factura2,impfac2,factura3,impfac3,factura4,impfac4,factura5,impfac5,factura6,impfac6,factura7,impfac7,factura8,impfac8,factura9,impfac9,factura10,impfac10) VALUES('" & _
                    Mid(AdoDbf.Recordset!Clavenota, 2) & "','" & AdoDbf.Recordset!Clavenota & "','" & Trim(FAC1) & "'," & IMP1 & ",'" & Trim(FAC2) & "'," & imp2 & ",'" & fac3 & "'," & imp3 & _
                    ",'" & fac4 & "'," & imp4 & ",'" & fac5 & "'," & Imp5 & ",'" & fac6 & "'," & imp6 & ",'" & fac7 & "'," & imp7 & _
                    ",'" & fac8 & "'," & imp8 & ",'" & fac9 & "'," & imp9 & ",'" & fac10 & "'," & imp10 & ")"
                    'MsgBox cadena
                    cn.Execute CADENA
              Else
                  'CRITERIO PARA QUE HACER EN EL CASO DE QUE YA EXISTA, SUPUESTAMENTE
                  'NO SE DEBE REPETIR EL ENVIO
              End If
              rstBus.Close
           
           'Verifico si existe el detalle de la nota
           If cNotAnt <> AdoDbf.Recordset!Clavenota Then
              'SE TIENE QUE CHECAR QUE EL PEDIDO GLOBAL EXISTE, TALVEZ SE GENERO EN LA BODEGA
              rstBus.Open "SELECT * FROM pedprove WHERE pp_pedido  = '" & Mid(AdoDbf.Recordset!Clavenota, 2) & "'"
              If Not rstBus.EOF Then
                 pr = InStr(AdoDbf.Recordset!Clavenota, "-")
                 PRO = Mid(AdoDbf.Recordset!Clavenota, pr - 3, 3)
                 cn.Execute "UPDATE PEDPROVE SET PP_recibe = 1, Pp_FecRecibe = '" & AdoDbf.Recordset!FecRec & "',PP_notent = '" & AdoDbf.Recordset!Clavenota & "' WHERE PP_PEDIDO = '" & Mid(AdoDbf.Recordset!Clavenota, 2) & "'"
                 cFecPed = rstBus!PP_FECHAGEN: cFecConf = rstBus!pp_fecconfirma
              Else
                 pr = InStr(AdoDbf.Recordset!Clavenota, "-")
                 PRO = Mid(AdoDbf.Recordset!Clavenota, pr - 3, 3)
                 EsBack = Len(Mid(AdoDbf.Recordset!Clavenota, 1, pr)) >= 6
                 If EsBack Then
                    CADENA = "INSERT INTO PEDPROVE(pp_proveedor,pp_pedido,pp_fechagen,pp_fecrecibe,pp_notent,pp_observa, pp_pedback, pp_recibe) values(" & _
                             "'" & PRO & "','" & Mid(AdoDbf.Recordset!Clavenota, 2) & "','" & AdoDbf.Recordset!FecRec & "','" & AdoDbf.Recordset!FecRec & "','" & AdoDbf.Recordset!Clavenota & "','PEDIDO GENERADO EN CARBONERA','" & Mid(AdoDbf.Recordset!Clavenota, pr - 3) & "',1)"
                 Else
                    CADENA = "INSERT INTO PEDPROVE(pp_proveedor,pp_pedido,pp_fechagen,pp_fecrecibe,pp_notent,pp_observa) values(" & _
                            "'" & PRO & "','" & Mid(AdoDbf.Recordset!Clavenota, 2) & "','" & AdoDbf.Recordset!FecRec & "','" & AdoDbf.Recordset!FecRec & "','" & AdoDbf.Recordset!Clavenota & "','PEDIDO GENERADO EN CARBONERA')"
                 End If
                 cn.Execute CADENA
                 cFecPed = Space(8): cFecConf = Space(8)
                 'Verifico si es backorder
                 If EsBack Then
                    'Agrego las facturas del backorder al pedido al que pertenecen
                    rsttemp.Open "SELECT * FROM NOTAENTRADA WHERE CLAVENOTA = 'N" & Mid(AdoDbf.Recordset!Clavenota, pr - 3) & "'", cn, adOpenDynamic, adLockOptimistic, adCmdText
                    For f = 1 To 10
                        Factura = "Factura" & Trim(Str(f))
                        impfact = "ImpFac" & Trim(Str(f))
                        If Not IsNull(AdoDbf.Recordset.Fields(Factura).Value) Then
                           lInd = False
                           For p = 1 To 20
                               FACNOTENT = "Factura" & Trim(Str(p))
                               IMPNOTENT = "ImpFac" & Trim(Str(p))
                               If rsttemp.BOF And rsttemp.EOF Then
                                  rsttemp.AddNew
                                  rsttemp!Clavenota = Trim(Mid(AdoDbf.Recordset!Clavenota, pr - 3))
                               End If
                               If rsttemp.Fields(IMPNOTENT).Value = 0 Or Trim(rsttemp.Fields(FACNOTENT).Value) = "" Or IsNull(rsttemp.Fields(FACNOTENT).Value) Then
                                  rsttemp.Fields(FACNOTENT).Value = Trim(AdoDbf.Recordset(Factura).Value)
                                  rsttemp.Fields(IMPNOTENT).Value = AdoDbf.Recordset(impfact).Value
                                  rsttemp.Update
                                  lInd = True
                                  Exit For
                               End If
                           Next
                           If Not lInd Then
                              'MsgBox "ERROR CRITICO; EL NUMERO DE FACTURAS DE TODOS LOS BACKORDERS Y EL PEDIDO EXCEDEN DE 15 FACTURAS, INFORME AL ADMINISTRADOR DEL SISTEMA PARA QUE SE INCREMENTEN CAMPOS YA QUE NO SE GUARDARAN LAS FACTURAS CON SUS IMPORTES CORRESPONDIENTES", vbCritical
                              'MsgBox "FOLIO DEL PEDIDO " & AdoDbf.Recordset!Clavenota
                           End If
                        End If
                    Next
                    rsttemp.Close
                 End If
              End If
              rstBus.Close
              
              rstBus.Open "SELECT * FROM catprov WHERE prove = '" & PRO & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
              Print #1, Mid(AdoDbf.Recordset!Clavenota, 2); cFecPed; Space(22 - Len(cFecPed)); AdoDbf.Recordset!FecRec; Space(22 - Len(AdoDbf.Recordset!FecRec)); PRO; IIf(rstBus.RecordCount = 0, "", "  " & Mid(rstBus!NOMPROVE, 1, 38))
              rstBus.Close
              
              nPed = nPed + 1
           End If
           'EN EL CASO DE QUE EL PEDIDO NO HAYA EXISTIDO , TAMBIEN SE DEBEN AGREGAR
           ' LOS PRODUCTO EN EL DETALLEGLOBAL
           CADENA = "SELECT * FROM DETALLEGLOBAL WHERE DG_PEDIDO = '" & Mid(AdoDbf.Recordset!Clavenota, 2) & "' AND DG_PRODUCTO = '" & Trim(AdoDbf.Recordset!producto) & "'"
           rstBus.Open CADENA
           If Not rstBus.EOF Then
              CADENA = "UPDATE DETALLEglobal SET dg_cantreal = " & AdoDbf.Recordset!CantRecC & ", dg_cantsol = " & AdoDbf.Recordset!CantSolc & ",dg_promocionr = " & AdoDbf.Recordset!promrec & ",dg_costo = " & AdoDbf.Recordset!costo & " where DG_PEDIDO = '" & Mid(AdoDbf.Recordset!Clavenota, 2) & "' and dg_producto = '" & Trim(AdoDbf.Recordset!producto) & "'"
              cn.Execute CADENA
           Else
              'Si es BackOrder se acumula al pedido que corresponde
              If EsBack Then
                 cn.Execute "UPDATE detalleglobal SET dg_cantreal = dg_cantreal + " & AdoDbf.Recordset!CantRecC & ", dg_promocionr = dg_promocionr + " & AdoDbf.Recordset!promrec & " WHERE dg_pedido = '" & Mid(AdoDbf.Recordset!Clavenota, pr - 3) & "' AND dg_producto = '" & Trim(AdoDbf.Recordset!producto) & "'"
              End If
              'Ahora agrego el producto al detalleglobal
              CADENA = "INSERT INTO detalleglobal (dg_cantreal,dg_cantsol,dg_promocion,dg_producto,dg_pedido,dg_promocionr,dg_costo) values(" & _
                        AdoDbf.Recordset!CantRecC & "," & AdoDbf.Recordset!CantSolc & "," & AdoDbf.Recordset!promrec & ",'" & Trim(AdoDbf.Recordset!producto) & "','" & Mid(AdoDbf.Recordset!Clavenota, 2) & "'," & AdoDbf.Recordset!promrec & "," & AdoDbf.Recordset!costo & ")"
              cn.Execute CADENA
           End If
           rstBus.Close
           CADENA = "SELECT * FROM detallenota WHERE producto = '" & AdoDbf.Recordset!producto & "' AND ClaveNota = '" & AdoDbf.Recordset!Clavenota & "'"
           rstBus.Open CADENA
           lExiste = rstBus.RecordCount > 0
           'SE ACTUALIZA LA NOTA DE ENTRADA
           If Not rstBus.EOF Then
              stb1.SimpleText = Space(75) & "Actualizando producto a la nota de entrada: " & AdoDbf.Recordset!producto
              cn.Execute "UPDATE DetalleNota SET Cantrec = " & AdoDbf.Recordset!CantRecC & ",costo = " & AdoDbf.Recordset!costo & " WHERE producto = '" & AdoDbf.Recordset!producto & "' AND Clavenota = '" & AdoDbf.Recordset!Clavenota & "'"
           Else
              stb1.SimpleText = Space(75) & "Agregando producto a la nota de entrada: " & AdoDbf.Recordset!producto
              'SE DEBEN INCLUIR LOS QUE SE SOLICITAR, PORQUE SE CAPTURARON EN LA BODEGA
              CADENA = "INSERT INTO DetalleNota(ClaveNota,producto,cantsol,cantsolp,cantrec,cantrecp,costo) VALUES ('" & AdoDbf.Recordset!Clavenota & "','" & AdoDbf.Recordset!producto & "'," & AdoDbf.Recordset!CantSolc & "," & AdoDbf.Recordset!cantsolp & "," & AdoDbf.Recordset!CantRecC & "," & AdoDbf.Recordset!cantrecp & "," & AdoDbf.Recordset!costo & ")"
             'MsgBox cadena
              cn.Execute CADENA
           End If
           stb1.Refresh
           cNotAnt = AdoDbf.Recordset!Clavenota
           rstBus.Close
           AdoDbf.Recordset.MoveNext
       Wend

   AdoDbf.Recordset.MoveFirst
   MsgBox "SE GENERARON UN TOTAL DE " & CStr(nPed) & " PEDIDOS", vbInformation
   Print #1, "=========================================================================================="
   Print #1, "TOTAL DE PEDIDOS: "; nPed
   Close #1   ' Cierra el archivo de reporte
   Handle = Shell("NOTEPAD " & App.Path & "\LLEGACAR.TXT", 1)

   'MsgBox "SE ACTUALIZARON UN TOTAL DE " & CStr(nPed) & " PEDIDOS", vbInformation
   Set cnFoxPro = New ADODB.Connection
   cnFoxPro.ConnectionString = "Provider=MSDASQL.1;DSN=PITICODBF;SourceDB=" & cruta & ";SourceType=DBF;Exclusive=No;Initial Catalog= " & cruta
   cnFoxPro.Open
   cnFoxPro.Execute "UPDATE " & cArch & " SET Importado = 1 "
   
   AdoPedidos.Refresh
   stb1.SimpleText = MenAnt
   stb1.Refresh
  Exit Sub
Error:
   If Err.Number <> 32755 Then  'Numero de error al presionar el boton cancelar
      MsgBox Err.Description
   End If
End Sub

'Exportar pedidos de Oficinas centrales para enviarlos a Bodega carbonera
Private Sub ExpPedOfi()
'On Error GoTo error:
Dim nDias As Integer
Dim FolPed As String
Dim rsttemp As ADODB.Recordset
Dim cnFoxPro As ADODB.Connection
   cMenAnt = stb1.SimpleText
   Cmdlg.DialogTitle = "Grabar archivo para enviar pedidos por proveedor a carbonera"
   Cmdlg.Filter = "Archivos PedOfi (*.dbf) | PedOfi.dbf"
   'Cmdlg.CancelError = True
   Cmdlg.ShowSave
   cRutArc = Cmdlg.FileName
   If cRutArc = "" Or IsNull(cRutArc) Then
      MsgBox "DEBE ESPECIFICAR UNA RUTA Y NOMBRE DE ARCHIVO", vbExclamation
      Exit Sub
   End If
   stb1.SimpleText = Space(45) & "Grabando archivo " & cRutArc
   stb1.Refresh
   
   For N = 1 To Len(cRutArc)
      If Mid(cRutArc, N, 1) = "\" Then nPos = N
   Next
   cruta = Mid(cRutArc, 1, nPos)
   cArch = Mid(cRutArc, nPos + 1)
   cArch = Mid(cArch, 1, Len(cArch) - 4) 'Le quito la extension porque si no marca error en la consula SQL.

   stb1.SimpleText = Space(65) & "Limpiando archivo " & cArch
   stb1.Refresh
   Set fs = CreateObject("Scripting.FileSystemObject")
   'Set F = fs.GetFile("C:\PASO\ESTPEDOF.DBF")   'save
   Set f = fs.GetFile("p:\ESPEDOF1.DBF")
   f.Copy cRutArc, True

  Set rsttemp = New ADODB.Recordset
  AdoDbf.CommandType = adCmdText
  AdoDbf.CursorType = adOpenKeyset
  AdoDbf.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=PITICODBF;SourceDB=" & cruta
  AdoDbf.RecordSource = "SELECT * FROM " & cArch & " ORDER BY Proveed"
  AdoDbf.Refresh
     
  Set rstDesPro = New ADODB.Recordset
  AdoPedidos.Recordset.MoveFirst
     'Genero el archivo de reporte
   'Open "\\SERVIDOR_OAXACA\PROGRAMAS\PASO\PEDOFIEN.TXT" For Output As #1   ' Abre el archivo para operaciones de salida.
   Open "P:\PASO\PEDOFIEN.TXT" For Output As #1   ' Abre el archivo para operaciones de salida.
   Print #1, Tab(20); "VIVERES Y LICORES S.A DE C.V."
   Print #1, Tab(10); "COLON 1016  OAXACA DE JUAREZ, OAX"
   Print #1, "PEDIDOS POR PROVEEDOR ENVIADOS EL "; UCase(Format(date, "long date"))  ' Escribe texto en el archivo.
   Print #1,   ' Imprime una línea en blanco en el archivo.
   Print #1, "=========================================================================================="
   Print #1, "PEDIDO        FEC. ELAB.        FEC.CONF              PROVEEDOR"
   Print #1, "=========================================================================================="
  nPed = 0
  While Not AdoPedidos.Recordset.EOF
        If AdoPedidos.Recordset!pp_recibe = True Then
            rec = 1
        Else
            rec = 0
        End If
       If AdoPedidos.Recordset!pp_cancelado = True Then
            can = 1
      Else
          can = 0
       End If
       If AdoPedidos.Recordset!pp_enviado = "1" Then
         env = 1
      Else
       env = 0
    End If
   If rec = 0 And can = 0 And env = 0 And AdoPedidos.Recordset!pp_confirma Then   ' SAVE
      If Not AdoPedidos.Recordset!pp_pedind Then
         rsttemp.Open "SELECT DG_PRODUCTO, DG_CANTSOL AS SOLCAJA, DG_CANTSOLP AS SOLPZA, DG_COSTO AS COSTO FROM detalleglobal, TFPRODUC WHERE dg_pedido = '" & AdoPedidos.Recordset!pp_pedido & "' AND DG_PRODUCTO = CONSEC", cn, adOpenKeyset, adLockOptimistic, adCmdText
         While Not rsttemp.EOF
            stb1.SimpleText = Space(50) & "Exportando pedido " & AdoPedidos.Recordset!pp_pedido & " y producto con la clave: " & CStr(rsttemp!DG_PRODUCTO)
            stb1.Refresh
            AdoDbf.Recordset.AddNew
            AdoDbf.Recordset!Pedido = AdoPedidos.Recordset!pp_pedido
            AdoDbf.Recordset!producto = Trim(rsttemp!DG_PRODUCTO)
            AdoDbf.Recordset!CantSolc = rsttemp!solcaja
            AdoDbf.Recordset!cantsolp = rsttemp!solpza
            'Se deben grabar tambien los descuentos como historicos
            'se graban tambien los descuentos y cargos
            CADENA = "UPDATE detalleglobal SET  dg_decto1  =  descprod.decto1, dg_decto2 =  descprod.decto2 " & _
                ",dg_decto3 =  descprod.decto3 ,dg_decto4 =  descprod.decto4 ,dg_decto5 =  descprod.decto5  " & _
                ",dg_decto6 = descprod.financiero ,dg_cargo1 = descprod.cargo1 ,dg_cargo2 = descprod.cargo2 " & _
                ",dg_iva =   descprod.cargo3 ,dg_ieps = descprod.cargo4 ,dg_cargo5 =  descprod.cargo5 " & _
                ",dg_maniobras = descprod.maniobras ,dg_flete = descprod.flete ,dg_efectivo =  descprod.efectivo " & _
                ",dg_prelista =  descprod.preciolista ,dg_cajas = descprod.Cajas ,dg_encajas =  descprod.Encajas " & _
                "  FROM descprod WHERE producto = dg_producto AND producto = '" & Trim(rsttemp!DG_PRODUCTO) & "'"
            'MsgBox CADENA
            'cn.Execute cadena
            AdoDbf.Recordset!costo = rsttemp!costo
            AdoDbf.Recordset!proveed = AdoPedidos.Recordset!pp_proveedor
            AdoDbf.Recordset!FecPed = Trim(AdoPedidos.Recordset!PP_FECHAGEN)
            AdoDbf.Recordset!FecConf = Trim(AdoPedidos.Recordset!pp_fecconfirma)
            AdoDbf.Recordset!sucursal = Trim(Mid(cSucursal, 1, 3))
            AdoDbf.Recordset!Importado = False
            AdoDbf.Recordset!Pedsug = False
            AdoDbf.Recordset!PROTEC = AdoPedidos.Recordset!pp_perprotec
            If IsNull(AdoDbf.Recordset!OBSERVA) Then
               AdoDbf.Recordset!OBSERVA = "."
            Else
               AdoDbf.Recordset!OBSERVA = IIf(Len(Mid(AdoPedidos.Recordset!pp_observa, 1, 250)) > 1, Replace(Mid(AdoPedidos.Recordset!pp_observa, 1, 250), "'", " "), ".")
            End If
            o1 = Mid(AdoPedidos.Recordset!pp_observa, 251, 250)
            If Len(Trim(o1)) > 1 Then
               AdoDbf.Recordset!Observa1 = Replace(o1, "'", " ")
            Else
               AdoDbf.Recordset!Observa1 = "."
            End If
            o2 = Mid(AdoPedidos.Recordset!pp_observa, 501, 250)
            If Len(Trim(o2)) > 1 Then
               AdoDbf.Recordset!Observa2 = Replace(o2, "'", " ")
            Else
               AdoDbf.Recordset!Observa2 = "."
            End If
            rstDesPro.Open "SELECT * FROM DescProd WHERE PRODUCTO = '" & rsttemp!DG_PRODUCTO & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
            If rstDesPro.RecordCount > 0 Then
                If rstDesPro!cajas > 0 Then
                   AdoDbf.Recordset!promsol = Int(rsttemp!solcaja / rstDesPro!encajas) * rstDesPro!cajas
                Else
                   AdoDbf.Recordset!promsol = 0
                End If
            End If
            AdoDbf.Recordset!promrec = 0
            rstDesPro.Close
            AdoDbf.Recordset.Update
            rsttemp.MoveNext
         Wend
         If Len(AdoPedidos.Recordset!pp_fecconfirma) > 0 Then
             Print #1, Me.AdoPedidos.Recordset!pp_pedido; AdoPedidos.Recordset!PP_FECHAGEN; Space(22 - Len(AdoPedidos.Recordset!PP_FECHAGEN)); AdoPedidos.Recordset!pp_fecconfirma; Space(22 - Len(AdoPedidos.Recordset!pp_fecconfirma)); Trim(cmbProved.Text)
        Else
            Print #1, Me.AdoPedidos.Recordset!pp_pedido; AdoPedidos.Recordset!PP_FECHAGEN; Space(22 - Len(AdoPedidos.Recordset!PP_FECHAGEN)); Trim(cmbProved.Text)
         End If
         nPed = nPed + 1

         'Ahora envio los pedidos sugeridos de tiendas que formaron el pedido por proveedor
         Print #1, "SUGERIDOS:"
         PEDANT = ""
         rsttemp.Close
         rsttemp.Open "SELECT * FROM DetalleFactura,pedidos WHERE df_pedido = p_pedido AND p_pedproveedor = '" & Trim(AdoPedidos.Recordset!pp_pedido) & "' AND p_cancelado = 0 AND p_surtbodega = 1", cn, adOpenKeyset, adLockOptimistic, adCmdText
         While Not rsttemp.EOF
            stb1.SimpleText = Space(50) & "Exportando pedido sugerido " & rsttemp!p_Pedido & " y producto con la clave: " & rsttemp!df_prod
            stb1.Refresh
            If rsttemp!df_pedido <> PEDANT Then
               Print #1, Space(13) & rsttemp!df_pedido
            End If
            AdoDbf.Recordset.AddNew
            AdoDbf.Recordset!Pedido = rsttemp!df_pedido
            AdoDbf.Recordset!producto = Trim(rsttemp!df_prod)
            AdoDbf.Recordset!CantSolc = rsttemp!df_cantsol
            AdoDbf.Recordset!cantsolp = rsttemp!df_cantsolp
            AdoDbf.Recordset!costo = rsttemp!df_Costo
            AdoDbf.Recordset!proveed = rsttemp!p_proveedor
            AdoDbf.Recordset!FecPed = Trim(rsttemp!p_fecped)
            AdoDbf.Recordset!FecConf = Trim(AdoPedidos.Recordset!pp_fecconfirma)
            AdoDbf.Recordset!OBSERVA = Replace(Mid(rsttemp!p_observaciones, 1, 250), "'", " ")
            AdoDbf.Recordset!sucursal = rsttemp!p_sucursal
            AdoDbf.Recordset!Importado = False
            AdoDbf.Recordset!Pedsug = True
            AdoDbf.Recordset!PEDPROVE = AdoPedidos.Recordset!pp_pedido
            AdoDbf.Recordset!promsol = 0
            AdoDbf.Recordset!promrec = 0
            'observaciones adicionales
            o1 = Mid(rsttemp!p_observaciones, 251, 250)
            If Len(Trim(o1)) > 1 Then
               AdoDbf.Recordset!Observa1 = Replace(o1, "'", " ")
            Else
               AdoDbf.Recordset!Observa1 = "."
            End If
            o2 = Mid(rsttemp!p_observaciones, 501, 250)
            If Len(Trim(o2)) > 1 Then
               AdoDbf.Recordset!Observa2 = Replace(o2, "'", " ")
            Else
               AdoDbf.Recordset!Observa2 = "."
            End If
            AdoDbf.Recordset.Update
            PEDANT = rsttemp!df_pedido
            rsttemp.MoveNext
         Wend
      Else 'Si es indirecto se envia todo el pedido global
         rsttemp.Open "SELECT * FROM detalleglobal WHERE dg_pedido = '" & Trim(AdoPedidos.Recordset!pp_pedido) & "'", cn, adOpenDynamic, adLockOptimistic, adCmdText
         While Not rsttemp.EOF
            stb1.SimpleText = Space(50) & "Exportando producto con la clave: " & CStr(rsttemp!DG_PRODUCTO) & " del pedido " & AdoPedidos.Recordset!pp_pedido
            stb1.Refresh
            AdoDbf.Recordset.AddNew
            AdoDbf.Recordset!Pedido = AdoPedidos.Recordset!pp_pedido
            AdoDbf.Recordset!producto = Trim(rsttemp!DG_PRODUCTO)
            AdoDbf.Recordset!CantSolc = rsttemp!dg_cantsol
            AdoDbf.Recordset!cantsolp = rsttemp!dg_cantsolp
            AdoDbf.Recordset!promsol = rsttemp!dg_promocion
            AdoDbf.Recordset!costo = rsttemp!dg_costo
            AdoDbf.Recordset!proveed = AdoPedidos.Recordset!pp_proveedor
            AdoDbf.Recordset!FecPed = Trim(AdoPedidos.Recordset!PP_FECHAGEN)
            AdoDbf.Recordset!FecConf = Trim(AdoPedidos.Recordset!pp_fecconfirma)
            AdoDbf.Recordset!sucursal = Trim(Mid(cSucursal, 1, 3))
            AdoDbf.Recordset!Importado = False
            AdoDbf.Recordset!OBSERVA = Mid(AdoPedidos.Recordset!pp_observa, 1, 250)
            'observaciones adicionales
            o1 = Mid(AdoPedidos.Recordset!pp_observa, 251, 250)
            If Len(o1) > 1 Then
               AdoDbf.Recordset!Observa1 = o1
            Else
               AdoDbf.Recordset!Observa1 = "."
            End If
            o2 = Mid(AdoPedidos.Recordset!pp_observa, 501, 250)
            If Len(o2) > 1 Then
               AdoDbf.Recordset!Observa2 = o2
            Else
               AdoDbf.Recordset!Observa2 = "."
            End If
            
            AdoDbf.Recordset.Update
            rsttemp.MoveNext
         Wend
         Print #1, Me.AdoPedidos.Recordset!pp_pedido; AdoPedidos.Recordset!PP_FECHAGEN; Space(22 - Len(AdoPedidos.Recordset!PP_FECHAGEN)); AdoPedidos.Recordset!pp_fecconfirma; Space(22 - Len(AdoPedidos.Recordset!pp_fecconfirma)); AdoPedidos.Recordset!pp_proveedor; Trim(cmbProved.Text)
         nPed = nPed + 1
      End If
      rsttemp.Close
      'PARA QUE YA NO LO MODIFIQUEN Y ENVIEN
      'AdoPedidos.Recordset!pp_enviado = "1"     'SAVE
      'AdoPedidos.Recordset.Update               'SAVE
   Else
      stb1.SimpleText = Space(50) & "PEDIDO ENVIADO A CARBONERA, IMPOSIBLE GENERAR DE NUEVO EL ENVIO..."
      stb1.Refresh
   End If
   AdoPedidos.Recordset.MoveNext
  Wend
  MsgBox "SE GENERARON UN TOTAL DE " & CStr(nPed) & " PEDIDOS", vbInformation
   Print #1, "=========================================================================================="
   Print #1, "TOTAL DE PEDIDOS: "; nPed
   Close #1   ' Cierra el archivo de reporte
   'Handle = Shell("NOTEPAD " & "\\SERVIDOR_OAXACA\PROGRAMAS\PASO\PEDOFIEN.TXT", 1)
   Handle = Shell("NOTEPAD " & "P:\PASO\PEDOFIEN.TXT", 1)
   
   AdoDbf.Recordset.Close
   stb1.SimpleText = cMenAnt
   'Actualizo como enviados los pedidos dentro del rango especificado
   cn.Execute "UPDATE pedprove SET pp_enviado = 1 WHERE pp_fechagen >= '" & dtpFecha(0).Value & "' AND pp_fechagen <= DATEADD(day, 1, '" & dtpFecha(1).Value & "') AND pp_cancelado = 0"
   MsgBox "PROCESO TERMINADO...!!!", vbInformation
   AdoPedidos.Recordset.MoveFirst
  Exit Sub
Error:
   If Err.Number <> 32755 Then  'Numero de error al presionar el boton cancelar
      MsgBox Err.Description
   End If
   stb1.SimpleText = cMenAnt
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
   MenAnt = stb1.SimpleText
   Cmdlg.DialogTitle = "Abrir archivo enviado por Oficinas centrales"
   Cmdlg.FileName = ""
   Cmdlg.CancelError = True   'Para que se genere error al hacer click en el boton cancelar
   Cmdlg.Filter = "Archivos Visual Foxpro (*.dbf) | *.dbf"
   Cmdlg.ShowOpen
   cRutArc = Cmdlg.FileName
   If cRutArc = "" Or IsNull(cRutArc) Then
      MsgBox "DEBE ESPECIFICAR UN NOMBRE DE ARCHIVO", vbExclamation
      Exit Sub
   End If
   For N = 1 To Len(cRutArc)
      If Mid(cRutArc, N, 1) = "\" Then nPos = N
   Next
   cruta = Mid(cRutArc, 1, nPos)
   cArch = Mid(cRutArc, nPos + 1)
   cArch = Mid(cArch, 1, Len(cArch) - 4) 'Le quito la extension porque si no marca error en la consula SQL.

   AdoDbf.CursorType = adOpenKeyset
   AdoDbf.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=PITICODBF;SourceDB=" & cruta
   AdoDbf.RecordSource = "SELECT * FROM " & cArch & " WHERE not EMPTY( Proveed) AND (EMPTY(PEDPROVE) OR PEDPROVE IS NULL) ORDER BY Pedido"
   AdoDbf.Refresh
   Set rs = New ADODB.Recordset
   If AdoDbf.Recordset.BOF And AdoDbf.Recordset.EOF Then
      MsgBox "EL ARCHIVO SELECCIONADO ESTA VACIO, NO EXISTEN DATOS PARA IMPORTAR", vbExclamation
      Exit Sub
   ElseIf AdoDbf.Recordset!Importado Then
      MsgBox "EL ARCHIVO SELECCIONADO YA FUE IMPORTADO", vbInformation
      Exit Sub
   End If
   'Genero el archivo de reporte
   Open App.Path & "\PEDPROVE.TXT" For Output As #1   ' Abre el archivo para operaciones de salida.
   Print #1, Tab(20); "VIVERES Y LICORES S.A DE C.V."
   Print #1, Tab(10); "CARBONERA 1016 COL. TRINIDAD DE LAS HUERTAS"
   Print #1, "PEDIDOS POR PROVEEDOR RECIBIDOS EL "; UCase(Format(date, "long date"))  ' Escribe texto en el archivo.
   Print #1,   ' Imprime una línea en blanco en el archivo.
   Print #1, "=========================================================================================="
   Print #1, "PEDIDO        FEC. ELAB.        FEC.CONF              PROVEEDOR"
   Print #1, "=========================================================================================="
   Set rstBus = New ADODB.Recordset  'Sirve para saber si existe o no el pedido
   Set rsttemp = New ADODB.Recordset
   cProAnt = "": nPed = 0
   AdoDbf.Recordset.MoveFirst
   While Not AdoDbf.Recordset.EOF
         'Agrego Pedido por proveedor
         If AdoDbf.Recordset!Pedido <> cProAnt Then
              FolPed = AdoDbf.Recordset!Pedido
                 rstBus.Open "SELECT * FROM pedprove WHERE pp_pedido = '" & Trim(FolPed) & "'", cn, adOpenDynamic, adLockOptimistic, adCmdText
                 'If rstBus.RecordCount > 0 Then
                 If Not (rstBus.BOF Or rstBus.EOF) Then
                    lrecibido = rstBus!pp_recibe
                 Else
                    'se unen los campos de observaciones
                    vobserva = " "
                    'vobserva = Adodbf.Recordset!Observa + Adodbf.Recordset!Observa1 + Adodbf.Recordset!Observa2
                    'On Error Resume Next
                    
                    vobserva = Replace(AdoDbf.Recordset!OBSERVA, "'", " ") + Replace(AdoDbf.Recordset!Observa1, "'", " ") + Replace(AdoDbf.Recordset!Observa2, "'", " ")
                    'MsgBox vobserva
                    cn.Execute "INSERT INTO PEDPROVE(pp_proveedor,pp_pedido,pp_fechagen,pp_fecConfirma,pp_observa) VALUES ('" & AdoDbf.Recordset!proveed & "','" & FolPed & "','" & AdoDbf.Recordset!FecPed & "','" & AdoDbf.Recordset!FecConf & "','" & vobserva & "')"
                    'MsgBox "SE GENERO EL PEDIDO POR PROVEEDOR CON FOLIO: " & FolPed & Chr(13) & " DEL PROVEEDOR CON CLAVE: " & AdoDbf.Recordset!proveed & Space(5) & "CON FECHA: " & AdoDbf.Recordset!FecPed, vbInformation
                    lrecibido = False
                 End If
                 nPed = nPed + 1
                 rstBus.Close
                 rstBus.Open "SELECT * FROM catprov WHERE prove = '" & AdoDbf.Recordset!proveed & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
                 CADENA = AdoDbf.Recordset!Pedido & AdoDbf.Recordset!FecPed & "   " & AdoDbf.Recordset!FecConf & "    " & AdoDbf.Recordset!proveed & IIf(rstBus.RecordCount = 0, "", Mid(rstBus!NOMPROVE, 1, 38))
                ' MsgBox cadena
                 Print #1, CADENA
                 Print #1, "SUGERIDOS:"
                 rstBus.Close
                 AdoSugDbf.CursorType = adOpenKeyset
                 AdoSugDbf.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=PITICODBF;SourceDB=" & cruta
                 AdoSugDbf.RecordSource = "SELECT * FROM " & cArch & " WHERE Proveed <> '' AND PEDPROVE = '" & AdoDbf.Recordset!Pedido & " ' ORDER BY Pedido"
                 AdoSugDbf.Refresh
                 
                 cProAntSug = ""
                 While Not AdoSugDbf.Recordset.EOF
                     'Se importan los sugeridos del pedido por proveedor
                     If AdoSugDbf.Recordset!Pedido <> cProAntSug Then
                        rstBus.Open "SELECT * FROM PEDIDOS WHERE p_pedido = '" & Trim(AdoSugDbf.Recordset!Pedido) & "' AND p_proveedor = '" & AdoDbf.Recordset!proveed & "'", cn, adOpenDynamic, adLockOptimistic, adCmdText
                        'If rstBus.RecordCount > 0 Then
                        If Not (rstBus.BOF Or rstBus.EOF) Then
                           lrecibido = rstBus!P_recibido
                        Else
                           CAD = "INSERT INTO pedidos(p_pedido,p_proveedor,p_fecped,p_sucursal,p_observaciones,p_pedproveedor,p_fecconfirma,p_situacion) VALUES ('" & Trim(AdoSugDbf.Recordset!Pedido) & "','" & AdoSugDbf.Recordset!proveed & "','" & AdoSugDbf.Recordset!FecPed & "','" & AdoSugDbf.Recordset!sucursal & "','" & AdoSugDbf.Recordset!OBSERVA & "','" & AdoSugDbf.Recordset!PEDPROVE & "','" & AdoSugDbf.Recordset!FecConf & "',1)"
                           'MsgBox cad
                           cn.Execute CAD
                           lrecibido = False
                        End If
                        Print #1, Space(13) & AdoSugDbf.Recordset!Pedido
                        rstBus.Close
                     End If
                     CADENA = "SELECT * FROM detallefactura WHERE  Df_prod = '" & Trim(AdoSugDbf.Recordset!producto) & "' AND df_pedido = '" & AdoSugDbf.Recordset!Pedido & "' "
                     'MsgBox CADENA
                     rstBus.Open CADENA, cn, adOpenDynamic, adLockOptimistic, adCmdText
                     If Not (rstBus.BOF Or rstBus.EOF) Then
                     'If rstBus.RecordCount > 0 Then
                        stb1.SimpleText = Space(50) & "Actualizando producto: " & AdoSugDbf.Recordset!producto & " del pedido sugerido " & AdoSugDbf.Recordset!Pedido
                        cn.Execute "UPDATE detallefactura SET df_cantsol = " & AdoSugDbf.Recordset!CantSolc & ", df_cantsolP = " & IIf(IsNull(AdoSugDbf.Recordset!cantsolp), 0, AdoSugDbf.Recordset!cantsolp) & ", df_promocion = " & IIf(IsNull(AdoSugDbf.Recordset!promsol), 0, AdoSugDbf.Recordset!promsol) & " FROM Pedidos WHERE df_pedido = p_pedido AND df_prod = '" & AdoSugDbf.Recordset!producto & "' AND df_pedido = '" & AdoSugDbf.Recordset!Pedido & "' AND p_proveedor = '" & AdoSugDbf.Recordset!proveed & "'"
                     Else
                        stb1.SimpleText = Space(50) & "Agregando producto: " & AdoSugDbf.Recordset!producto & " del pedido sugerido " & AdoSugDbf.Recordset!Pedido
                        CADENA = "INSERT INTO Detallefactura(df_pedido, df_prod,df_cantsol,df_cantsolp,df_promocion,df_costo) VALUES ('" & AdoSugDbf.Recordset!Pedido & "','" & AdoSugDbf.Recordset!producto & "'," & AdoSugDbf.Recordset!CantSolc & "," & IIf(IsNull(AdoSugDbf.Recordset!cantsolp), 0, AdoSugDbf.Recordset!cantsolp) & "," & IIf(IsNull(AdoSugDbf.Recordset!promsol), 0, AdoSugDbf.Recordset!promsol) & "," & AdoSugDbf.Recordset!costo & ")"
                        cn.Execute CADENA
                     End If
                     rstBus.Close
                     cProAntSug = AdoSugDbf.Recordset!Pedido
                     AdoSugDbf.Recordset.MoveNext
                 Wend
         End If
         rsttemp.Open "SELECT * FROM TFPRODUC WHERE CONSEC = '" & CStr(AdoDbf.Recordset!producto) & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
         If rsttemp.BOF And rsttemp.EOF Then
            MsgBox "EL PRODUCTO NO SE ENCUENTRA EN EL CATALOGO, ANOTE LOS DATOS E INFORME AL ADMINISTRADOR DEL SISTEMA" & Chr(13) & "PEDIDO:" & FolPed & Chr(13) & "PROVEEDOR: " & AdoDbf.Recordset!proveed & Chr(13) & "CLAVE PRODUCTO: " & CStr(AdoDbf.Recordset!producto) & Chr(13) & "CAJAS: " & CStr(AdoDbf.Recordset!CantSolc) & Chr(13) & "PIEZAS: " & CStr(AdoDbf.Recordset!cantsolp), vbCritical
         End If
         'SE MODIFICAN SOLOS PEDIDOS PENDIENTES DE RECIBIR
         If lrecibido = False Then
                rstBus.Open "SELECT * FROM detalleglobal WHERE DG_PRODUCTO = '" & Trim(AdoDbf.Recordset!producto) & "' AND DG_pedido = '" & FolPed & "'", cn, adOpenDynamic, adLockOptimistic, adCmdText
                'If rstBus.RecordCount > 0 Then
                If Not (rstBus.BOF Or rstBus.EOF) Then
                    stb1.SimpleText = Space(50) & "Actualizando producto: " & AdoDbf.Recordset!producto & " del pedido por proveedor " & AdoDbf.Recordset!Pedido
                    cn.Execute "UPDATE detalleglobal SET dg_cantsol = " & AdoDbf.Recordset!CantSolc & ", dg_cantsolP = " & IIf(IsNull(AdoDbf.Recordset!cantsolp), 0, AdoDbf.Recordset!cantsolp) & ", dg_promocion = " & IIf(IsNull(AdoDbf.Recordset!promsol), 0, AdoDbf.Recordset!promsol) & " WHERE dg_producto = '" & AdoDbf.Recordset!producto & "' AND dg_pedido = '" & FolPed & "'"
                Else
                    stb1.SimpleText = Space(50) & "Agregando producto: " & AdoDbf.Recordset!producto & " del pedido por proveedor " & AdoDbf.Recordset!Pedido
                    CADENA = "INSERT INTO DetalleGlobal(dg_pedido,dg_producto,dg_cantsol,dg_cantsolp,dg_promocion,dg_costo) VALUES ('" & FolPed & "','" & AdoDbf.Recordset!producto & "'," & AdoDbf.Recordset!CantSolc & "," & IIf(IsNull(AdoDbf.Recordset!cantsolp), 0, AdoDbf.Recordset!cantsolp) & "," & IIf(IsNull(AdoDbf.Recordset!promsol), 0, AdoDbf.Recordset!promsol) & "," & AdoDbf.Recordset!costo & ")"
                    cn.Execute CADENA
                End If
            rstBus.Close
         End If
         stb1.Refresh
         cProAnt = AdoDbf.Recordset!Pedido
         AdoDbf.Recordset.MoveNext
         rsttemp.Close
   Wend
   AdoDbf.Recordset.MoveFirst
   MsgBox "SE GENERARON UN TOTAL DE " & CStr(nPed) & " PEDIDOS", vbInformation
   Print #1, "=========================================================================================="
   Print #1, "TOTAL DE PEDIDOS: "; nPed
   
   Close #1   ' Cierra el archivo de reporte
   Handle = Shell("NOTEPAD " & App.Path & "\PEDPROVE.TXT", 1)
   Set cnFoxPro = New ADODB.Connection
   cnFoxPro.ConnectionString = "Provider=MSDASQL.1;DSN=PITICODBF;SourceDB=" & cruta & ";SourceType=DBF;Exclusive=No;Initial Catalog= " & cruta
   cnFoxPro.Open
   cnFoxPro.Execute "UPDATE " & cArch & " SET Importado = 1 "
   AdoPedidos.Refresh
   stb1.SimpleText = MenAnt
   stb1.Refresh
  Exit Sub
Error:
  If Err.Number <> 32755 Then  'Numero de error al presionar el boton cancelar
     MsgBox Err.Description
  End If
  Close #1   ' Cierra el archivo de reporte
End Sub


'Exportar llegadas de pedidos para enviarlas a Oficinas centrales
Private Sub ExpPedCar()
On Error GoTo Error:
Dim nDias As Integer
Dim FolPed As String
Dim rsttemp As ADODB.Recordset
Dim cnFoxPro As ADODB.Connection
Dim rs As ADODB.Recordset
   cMenAnt = stb1.SimpleText
   Cmdlg.DialogTitle = "Grabar archivo para enviar llegadas de pedidos a Oficinas centrales"
   Cmdlg.Filter = "Archivos Visual Fox pro (*.dbf) | *.dbf"
   Cmdlg.CancelError = True
   Cmdlg.ShowSave
   cRutArc = Cmdlg.FileName
   If cRutArc = "" Or IsNull(cRutArc) Then
      MsgBox "DEBE ESPECIFICAR UNA RUTA Y NOMBRE DE ARCHIVO", vbExclamation
      Exit Sub
   End If
   stb1.SimpleText = Space(45) & "Grabando archivo " & cRutArc
   stb1.Refresh
   
   For N = 1 To Len(cRutArc)
      If Mid(cRutArc, N, 1) = "\" Then nPos = N
   Next
   cruta = Mid(cRutArc, 1, nPos)
   cArch = Mid(cRutArc, nPos + 1)
   cArch = Mid(cArch, 1, Len(cArch) - 4) 'Le quito la extension porque marca error en la consula SQL.

   stb1.SimpleText = Space(65) & "Limpiando archivo " & cArch
   stb1.Refresh
   Set fs = CreateObject("Scripting.FileSystemObject")
   Set f = fs.GetFile("\\Server_carbo\Programas\EsPedCar.dbf")
   'Set F = fs.GetFile("c:\paso\EsPedCar.dbf")
   f.Copy cRutArc, True

   Set rsttemp = New ADODB.Recordset
   AdoDbf.CommandType = adCmdText
   AdoDbf.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=PITICODBF;SourceDB=" & cruta
   AdoDbf.RecordSource = "SELECT * FROM " & cArch
   AdoDbf.Refresh
     
  Set rs = New ADODB.Recordset
  cOper = IIf(Month(dtpFecha(0).Value) = Month(dtpFecha(1).Value), " AND ", " OR ")
  cFecha = " AND (month(pp_fecrecibe) >= " & Month(dtpFecha(0).Value) & " and (day(pp_fecrecibe) > = " & Day(dtpFecha(0).Value) & cOper & " (day(pp_fecrecibe)<= " & Day(dtpFecha(1).Value) & " and month(pp_fecrecibe)<= " & Month(dtpFecha(1).Value) & ")) and year(pp_fecrecibe)>= " & Year(dtpFecha(0).Value) & " and year(pp_fecrecibe)<= " & Year(dtpFecha(1).Value) & ")"
  rs.Open "SELECT * FROM pedprove WHERE pp_recibe = 1 " & cFecha, cn, adOpenKeyset, adLockOptimistic, adCmdText
  nPed = 0
  While Not rs.EOF
       rsttemp.Open "SELECT * FROM Pedprove,DetalleGlobal,NotaEntrada WHERE pp_pedido = '" & rs!pp_pedido & "' AND pp_pedido = dg_pedido AND pp_pedido = Pedido", cn, adOpenStatic, adLockOptimistic, adCmdText
       While Not rsttemp.EOF
            stb1.SimpleText = Space(15) & "Exportando producto con la clave: " & CStr(rsttemp!DG_PRODUCTO) & " del pedido: " & rs!pp_pedido & " de fecha " & rs!pp_fecrecibe
            stb1.Refresh
            AdoDbf.Recordset.AddNew
            AdoDbf.Recordset!Clavenota = Trim(rsttemp!Clavenota)
            AdoDbf.Recordset!producto = rsttemp!DG_PRODUCTO
            AdoDbf.Recordset!CantSolc = Val(rsttemp!dg_cantsol)
            AdoDbf.Recordset!cantsolp = Val(rsttemp!dg_cantsolp)
            AdoDbf.Recordset!CantRecC = Val(rsttemp!dg_cantreal)
            AdoDbf.Recordset!cantrecp = Val(rsttemp!dg_cantrealp)
            AdoDbf.Recordset!promsol = Val(rsttemp!dg_promocion)
            AdoDbf.Recordset!promrec = Val(rsttemp!DG_PROMOCIONR)
            AdoDbf.Recordset!costo = rsttemp!dg_costo
            AdoDbf.Recordset!sucursal = Trim(Mid(cSucursal, 1, 3))
            AdoDbf.Recordset!FecRec = Mid(rsttemp!pp_fecrecibe, 1, 8)
            AdoDbf.Recordset!factura1 = rsttemp!factura1
            AdoDbf.Recordset!Impfac1 = rsttemp!Impfac1
            AdoDbf.Recordset!factura2 = rsttemp!factura2
            AdoDbf.Recordset!Impfac2 = rsttemp!Impfac2
            AdoDbf.Recordset!Factura3 = rsttemp!Factura3
            AdoDbf.Recordset!Impfac3 = rsttemp!Impfac3
            AdoDbf.Recordset!factura4 = rsttemp!factura4
            AdoDbf.Recordset!Impfac4 = rsttemp!Impfac4
            AdoDbf.Recordset!factura5 = rsttemp!factura5
            AdoDbf.Recordset!Impfac5 = rsttemp!Impfac5
            AdoDbf.Recordset!factura6 = rsttemp!factura6
            AdoDbf.Recordset!Impfac6 = rsttemp!Impfac6
            AdoDbf.Recordset!factura7 = rsttemp!factura7
            AdoDbf.Recordset!Impfac7 = rsttemp!Impfac7
            AdoDbf.Recordset!factura8 = rsttemp!factura8
            AdoDbf.Recordset!Impfac8 = rsttemp!Impfac8
            AdoDbf.Recordset!factura9 = rsttemp!factura9
            AdoDbf.Recordset!Impfac9 = rsttemp!Impfac9
            AdoDbf.Recordset!factura10 = rsttemp!factura10
            AdoDbf.Recordset!Impfac10 = rsttemp!Impfac10
            AdoDbf.Recordset!Importado = False
            AdoDbf.Recordset.Update
            rsttemp.MoveNext
      Wend
      rsttemp.Close
      nPed = nPed + 1
      rs.MoveNext
  Wend
  AdoDbf.Recordset.Close
  stb1.SimpleText = cMenAnt
  MsgBox "SE ENVIARON " & CStr(nPed) & " PEDIDOS RECIBIDOS DEL " & dtpFecha(0).Value & " AL " & dtpFecha(1).Value, vbInformation
  Exit Sub
Error:
   If Err.Number <> 32755 Then  'Numero de error al presionar el boton cancelar
      MsgBox Err.Description
   End If
   stb1.SimpleText = cMenAnt
End Sub


