VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmModInvbod 
   Caption         =   "Modificar Inventario"
   ClientHeight    =   6645
   ClientLeft      =   60
   ClientTop       =   180
   ClientWidth     =   11880
   Icon            =   "frmModInvbod.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6645
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame fratda 
      BackColor       =   &H80000001&
      Caption         =   "Seleccion de Inventario  de Tienda:"
      Enabled         =   0   'False
      Height          =   2415
      Left            =   3360
      TabIndex        =   23
      Top             =   2280
      Visible         =   0   'False
      Width           =   5055
      Begin VB.OptionButton opt 
         Caption         =   "Piso"
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   28
         Top             =   1560
         Width           =   1095
      End
      Begin VB.OptionButton opt 
         Caption         =   "Bodega"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   27
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CheckBox CHKTODAS 
         Caption         =   "TODAS LAS TIENDAS"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   1080
         Width           =   2655
      End
      Begin VB.CommandButton btncambia 
         Caption         =   "Ok"
         Height          =   375
         Left            =   3960
         TabIndex        =   25
         Top             =   1920
         Width           =   975
      End
      Begin VB.ComboBox cmbtiendas 
         Height          =   315
         Left            =   240
         TabIndex        =   24
         Top             =   480
         Width           =   3495
      End
   End
   Begin ComctlLib.StatusBar STB1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   21
      Top             =   6300
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   609
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   5
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   5292
            MinWidth        =   5292
            Text            =   "Click en el encabezado ordena los datos "
            TextSave        =   "Click en el encabezado ordena los datos "
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "F4 = Historial"
            TextSave        =   "F4 = Historial"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   3951
            MinWidth        =   3951
            Text            =   "F5 = Pendientes de recibir"
            TextSave        =   "F5 = Pendientes de recibir"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   3881
            MinWidth        =   3881
            Text            =   "F6  = Seleccion de Inventario"
            TextSave        =   "F6  = Seleccion de Inventario"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   "F7 = Totales"
            TextSave        =   "F7 = Totales"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraCon 
      BackColor       =   &H80000001&
      Caption         =   "PROPORCIONE CONTRASEÑA"
      Height          =   1695
      Left            =   4080
      TabIndex        =   14
      Top             =   2640
      Visible         =   0   'False
      Width           =   4215
      Begin VB.TextBox txtContra 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1920
         PasswordChar    =   "*"
         TabIndex        =   17
         Top             =   480
         Width           =   1815
      End
      Begin VB.CommandButton cmdConAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   840
         TabIndex        =   16
         Top             =   1200
         Width           =   1050
      End
      Begin VB.CommandButton cmdConCance 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   2520
         TabIndex        =   15
         Top             =   1200
         Width           =   1050
      End
      Begin VB.Label lblContra 
         BackColor       =   &H80000001&
         Caption         =   "Contraseña"
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   480
         Width           =   975
      End
   End
   Begin MSAdodcLib.Adodc Adodbf 
      Height          =   330
      Left            =   0
      Top             =   3480
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
   Begin MSComDlg.CommonDialog cmdlg 
      Left            =   8280
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AdoDetAju 
      Height          =   330
      Left            =   0
      Top             =   3120
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
      Caption         =   "AdoDetAju"
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
      Height          =   570
      Left            =   0
      ScaleHeight     =   510
      ScaleWidth      =   11820
      TabIndex        =   12
      Top             =   5730
      Width           =   11880
      Begin VB.CheckBox chktodos 
         Caption         =   "Todos"
         Height          =   220
         Left            =   7920
         TabIndex        =   20
         Top             =   270
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CheckBox chktienda 
         Caption         =   "Tienda"
         Height          =   220
         Left            =   7920
         TabIndex        =   19
         Top             =   30
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdBuscaBarra 
         Caption         =   "&Barras"
         Height          =   375
         Left            =   120
         Picture         =   "frmModInvbod.frx":0E42
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Busqueda por codigo de barras"
         Top             =   80
         Width           =   735
      End
      Begin VB.CommandButton cmdrefresh 
         Caption         =   "&Actualizar"
         Height          =   375
         Left            =   3600
         Picture         =   "frmModInvbod.frx":0F78
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Actualizar datos para reflejar cambios realizados por otros usuarios"
         Top             =   80
         Width           =   800
      End
      Begin VB.CommandButton CmdExporta 
         Caption         =   "&Exportar"
         Height          =   375
         Left            =   4560
         Picture         =   "frmModInvbod.frx":107A
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Exportar existencias a archivo DBF"
         Top             =   80
         Width           =   800
      End
      Begin VB.CommandButton cmdBuscaDesc 
         Caption         =   "&Descripc"
         Height          =   375
         Left            =   1800
         Picture         =   "frmModInvbod.frx":117C
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Busqueda por descripcion"
         Top             =   80
         Width           =   735
      End
      Begin VB.CommandButton cmdBuscaCve 
         Caption         =   "&Clave"
         Height          =   375
         Left            =   960
         Picture         =   "frmModInvbod.frx":1276
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Busqueda por clave"
         Top             =   80
         Width           =   735
      End
      Begin VB.CommandButton CmdActualizar 
         Caption         =   "&Modificar"
         Height          =   375
         Left            =   2640
         Picture         =   "frmModInvbod.frx":1370
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Ajustar inventario"
         Top             =   80
         Width           =   800
      End
      Begin VB.CommandButton cmdUltimo 
         Enabled         =   0   'False
         Height          =   375
         Left            =   720
         Picture         =   "frmModInvbod.frx":14E2
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Ultimo"
         Top             =   80
         Visible         =   0   'False
         Width           =   400
      End
      Begin VB.CommandButton cmdAnterior 
         Height          =   375
         Left            =   600
         Picture         =   "frmModInvbod.frx":1630
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Anterior"
         Top             =   -120
         Visible         =   0   'False
         Width           =   400
      End
      Begin VB.CommandButton cmdSiguiente 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1320
         Picture         =   "frmModInvbod.frx":177E
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Siguiente"
         Top             =   -120
         Width           =   400
      End
      Begin VB.CommandButton cmdPrimero 
         Height          =   375
         Left            =   240
         Picture         =   "frmModInvbod.frx":18CC
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Primero"
         Top             =   80
         Width           =   400
      End
      Begin VB.CommandButton cmdRegresar 
         Caption         =   "&Regresar"
         Height          =   375
         Left            =   5400
         Picture         =   "frmModInvbod.frx":1A1A
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Regresar al menu principal"
         Top             =   80
         Width           =   800
      End
      Begin VB.Label txtactivo 
         Caption         =   "ACTIVO"
         Height          =   255
         Left            =   6360
         TabIndex        =   22
         Top             =   120
         Width           =   2295
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Numero de productos XX"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9120
         TabIndex        =   13
         Top             =   120
         Width           =   2535
      End
   End
   Begin MSAdodcLib.Adodc AdoModInv 
      Height          =   330
      Left            =   0
      Top             =   3840
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
      Caption         =   "AdoModInv"
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
   Begin MSDataGridLib.DataGrid dbgrdModInv 
      Bindings        =   "frmModInvbod.frx":1B8C
      Height          =   7095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   12515
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      BackColor       =   16777215
      ForeColor       =   -2147483642
      HeadLines       =   1.5
      RowHeight       =   16
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
      ColumnCount     =   13
      BeginProperty Column00 
         DataField       =   "inprod"
         Caption         =   "  CLAVE"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Descripc"
         Caption         =   "                              DESCRIPCION"
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
         DataField       =   "medida"
         Caption         =   "         MEDIDA"
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
         DataField       =   "InInicialP"
         Caption         =   "INV.INI.PZAS."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "barraspza"
         Caption         =   "   BARRAS PZA."
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
         DataField       =   "InCant"
         Caption         =   "EXIST. CAJAS"
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
         DataField       =   "InCantPza"
         Caption         =   "EXIST.PZAS."
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
         DataField       =   "activo"
         Caption         =   "ACTIVO"
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
      BeginProperty Column08 
         DataField       =   "InInicial"
         Caption         =   "INV.INI. CAJAS"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "minimo"
         Caption         =   "MINIMO"
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
      BeginProperty Column10 
         DataField       =   "maximo"
         Caption         =   "MAXIMO"
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
      BeginProperty Column11 
         DataField       =   "incantcdc"
         Caption         =   "CAJA CDC"
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
      BeginProperty Column12 
         DataField       =   "incantpzacdc"
         Caption         =   "PZA CDC"
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
         Size            =   508
         BeginProperty Column00 
            Locked          =   -1  'True
            ColumnWidth     =   810.142
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   4844.977
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   1635.024
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            Locked          =   -1  'True
            Object.Visible         =   0   'False
            ColumnWidth     =   1110.047
         EndProperty
         BeginProperty Column04 
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            Object.Visible         =   -1  'True
            ColumnWidth     =   1184.882
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            Object.Visible         =   -1  'True
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column07 
            Alignment       =   2
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   689.953
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   1260.284
         EndProperty
         BeginProperty Column09 
            Alignment       =   1
            Object.Visible         =   -1  'True
            ColumnWidth     =   945.071
         EndProperty
         BeginProperty Column10 
            Alignment       =   1
            Object.Visible         =   -1  'True
            ColumnWidth     =   975.118
         EndProperty
         BeginProperty Column11 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1035.213
         EndProperty
         BeginProperty Column12 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   989.858
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmModInvbod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private nCantAnt
Private lCon

Private Sub AdoModInv_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.ERROR, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error GoTo ERROR:
If AdoModInv.Recordset!activo Then
      txtactivo.Caption = "ACTIVO"
      txtactivo.ForeColor = QBColor(1)
   Else
      txtactivo.Caption = "BAJA"
      txtactivo.ForeColor = QBColor(4)
End If
txtactivo.Refresh
Exit Sub
ERROR:
  MsgBox "Intente De nuevo...", vbInformation, "INVENTARIO"
End Sub

Private Sub btncambia_Click()
On Error GoTo ERROR:
fratda.Enabled = False
fratda.Visible = False
If CHKTODAS.Value = 0 Then
    'VERIFICANDO QUE EXISTA LA TABLA
    Me.Caption = "INVENTARIO BODEGA    --" & cmbtiendas.Text
    CLAVEINVENTARIO = Mid(cmbtiendas.Text, 2, 2)
    If CLAVEINVENTARIO = 10 Or CLAVEINVENTARIO = 3 Then
       cad = "SELECT TFPRODUC.activo, TFPRODUC.BarrasPza, INVENTARIO.inprod, TFPRODUC.Descripc,  LTRIM(STR(TFPRODUC.paquetes)) + ' X ' + lTrim(str(TFPRODUC.contenid,10,3)) + space(2) + TFPRODUC.medida  AS MEDIDA, INVENTARIO.InInicial, INVENTARIO.incant, INVENTARIO.minimo, INVENTARIO.maximo,INVENTARIO.InCantPza, incantcdc, incantpzacdc  FROM INVENTARIO , TFPRODUC WHERE INVENTARIO.inprod = TFPRODUC.consec AND TFPRODUC.Paquetes > 0 ORDER BY DESCRIPC,CONTENID"
    Else
       cad = "SELECT TFPRODUC.activo, TFPRODUC.BarrasPza, INVENTARIO.inprod, TFPRODUC.Descripc,  LTRIM(STR(TFPRODUC.paquetes)) + ' X ' + lTrim(str(TFPRODUC.contenid,10,3)) + space(2) + TFPRODUC.medida  AS MEDIDA, INVENTARIO.InInicial, INVENTARIO.incant, INVENTARIO.minimo, INVENTARIO.maximo,INVENTARIO.InCantPza, incantcdc, incantpzacdc  FROM INVENTARIO" & Trim(CLAVEINVENTARIO) & "  as inventario , TFPRODUC WHERE INVENTARIO.inprod = TFPRODUC.consec AND TFPRODUC.Paquetes > 0 ORDER BY DESCRIPC,CONTENID"
    End If
    AdoModInv.RecordSource = cad
    AdoModInv.Refresh
    lblInfo.Caption = "Num. de productos " + Str(AdoModInv.Recordset.RecordCount)
    Exit Sub
Else
    frminvtodo.Show
    Exit Sub
End If
ERROR:
     MsgBox Err.Description
     Stb1.Panels(1).Text = "No se encontro el Inventario de esta Bodega ..."
     Unload Me
End Sub

Private Sub cmdBuscaBarra_Click()
Dim cCve As String
Dim Antes
cCve = InputBox("Introduzca el codigo de barras a buscar", "Introducir codigo de barras")
If IsNull(Trim(cCve)) Or Trim(cCve) = "" Then Exit Sub
Antes = dbgrdModInv.Bookmark
AdoModInv.Recordset.MoveFirst
AdoModInv.Recordset.Find "Barraspza =" & Trim(cCve)
If AdoModInv.Recordset.EOF Then
   MsgBox "El codigo de barras " & cCve & " no se encuentra en el inventario", vbExclamation
   dbgrdModInv.Bookmark = Antes
End If
dbgrdModInv.SetFocus
End Sub

Private Sub cmdBuscaCve_Click()
Dim cCve As String
Dim Antes
cCve = InputBox("Introduzca la clave a buscar", "Introducir clave")
If IsNull(Trim(cCve)) Or Trim(cCve) = "" Then Exit Sub
Antes = dbgrdModInv.Bookmark
AdoModInv.Recordset.MoveFirst
AdoModInv.Recordset.Find "Inprod LIKE '" & Trim(cCve) & "*'"
If AdoModInv.Recordset.EOF Then
   MsgBox "La clave " & cCve & " no se encuentra en el inventario", vbExclamation
   dbgrdModInv.Bookmark = Antes
End If
dbgrdModInv.SetFocus
End Sub

Private Sub cmdBuscaDesc_Click()
Dim cCve As String
Dim Antes

cCve = InputBox("Introduzca la descripcion del producto a buscar", "Introducir descripcion")
If IsNull(Trim(cCve)) Or Trim(cCve) = "" Then Exit Sub
Antes = dbgrdModInv.Bookmark
AdoModInv.Recordset.MoveFirst
AdoModInv.Recordset.Find "DESCRIPC LIKE '" & Trim(cCve) & "*'"
If AdoModInv.Recordset.EOF Then
   MsgBox "La descripcion " & cCve & " no se encuentra en el inventario", vbExclamation
   dbgrdModInv.Bookmark = Antes
End If
dbgrdModInv.SetFocus
End Sub

Private Sub cmdConAceptar_Click()
If txtContra.Text <> "P4567" Then
   MsgBox "LA CONTRASEÑA DE ACCESO ES INCORRECTA", vbExclamation
   txtContra.SetFocus
   SendKeys "+{HOME}"
   Exit Sub
Else
   lCon = True
   Me.dbgrdModInv.AllowUpdate = True
   fraCon.Visible = False
End If

End Sub

Private Sub cmdConCance_Click()
  fraCon.Visible = False
End Sub

Private Sub CmdExporta_Click()
On Error GoTo ERROR:
Dim nDias As Integer
Dim FolPed As String
Dim rsttemp As ADODB.Recordset
Dim cnFoxPro As ADODB.Connection
   cMenAnt = Stb1.Panels(1).Text
   Cmdlg.DialogTitle = "Seleccionar archivo para exportar existencias"
   Cmdlg.Filter = "Archivos Visual Fox pro (*.dbf) | *.dbf"
   Cmdlg.CancelError = True
   Cmdlg.ShowSave
   cRutArc = Cmdlg.FileName
   If cRutArc = "" Or IsNull(cRutArc) Then
      MsgBox "DEBE ESPECIFICAR UNA RUTA Y NOMBRE DE ARCHIVO", vbExclamation
      Exit Sub
   End If
   Stb1.Panels(1).Text = Space(45) & "Grabando archivo " & cRutArc
   Stb1.Refresh
   
   For n = 1 To Len(cRutArc)
      If Mid(cRutArc, n, 1) = "\" Then nPos = n
   Next
   cRuta = Mid(cRutArc, 1, nPos)
   cArch = Mid(cRutArc, nPos + 1)
   cArch = Mid(cArch, 1, Len(cArch) - 4) 'Le quito la extension porque si no marca error en la consula SQL.

   Stb1.Panels(1).Text = Space(65) & "Limpiando archivo " & cArch
   Stb1.Refresh
   Set fs = CreateObject("Scripting.FileSystemObject")
   Set f = fs.GetFile("P:\ESTINVEN.DBF")
   f.Copy cRutArc, True

   Set rsttemp = New ADODB.Recordset
   AdoDbf.CommandType = adCmdText
   AdoDbf.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=PITICODBF;SourceDB=" & cRuta
   AdoDbf.RecordSource = "SELECT * FROM " & cArch
   AdoDbf.Refresh
   
   'rsttemp.Open "SELECT * FROM Inventario, Tfproduc WHERE Inprod = Consec AND InCant > 0", cn, adOpenStatic, adLockOptimistic, adCmdText
   If CLAVEINVENTARIO = 10 Or CLAVEINVENTARIO = 3 Then
       rsttemp.Open "SELECT * FROM Inventario , Tfproduc WHERE Inprod = Consec AND InCant > 0", cn, adOpenStatic, adLockOptimistic, adCmdText
    Else
       rsttemp.Open "SELECT * FROM Inventario" & Trim(CLAVEINVENTARIO) & " as inventario , Tfproduc WHERE Inprod = Consec AND InCant > 0", cn, adOpenStatic, adLockOptimistic, adCmdText
    End If
   'rsttemp.Open "SELECT * FROM Inventario, Tfproduc WHERE Inprod = Consec", cn, adOpenStatic, adLockOptimistic, adCmdText
   MsgBox "A CONTINUACION SE EXPORTARAN: " & CStr(rsttemp.RecordCount) & " REGISTROS", vbInformation
   
   nReg = 1: nTotReg = CStr(rsttemp.RecordCount)
   While Not rsttemp.EOF
      Stb1.Panels(1).Text = Space(45) & "Exportando producto con la clave: " & CStr(rsttemp!Inprod) & Space(5) & "Producto: " & CStr(nReg) & " de " & nTotReg
      Stb1.Refresh
      AdoDbf.Recordset.AddNew
      AdoDbf.Recordset!CLAFAMIL = ""
      AdoDbf.Recordset!claprove = IIf(IsNull(rsttemp!claprove), "", Trim(rsttemp!claprove))
      AdoDbf.Recordset!CONSEC = Val(rsttemp!CONSEC)
      AdoDbf.Recordset!descripc = rsttemp!descripc
      AdoDbf.Recordset!Paquetes = IIf(IsNull(rsttemp!Paquetes), 0, Trim(rsttemp!Paquetes))
      AdoDbf.Recordset!Contenid = IIf(IsNull(rsttemp!Contenid), 0, Val(rsttemp!Contenid))
      AdoDbf.Recordset!Medida = IIf(IsNull(rsttemp!Medida), "", Trim(rsttemp!Medida))
      'AdoDbf.Recordset!ExiCaja = Val(rsttemp!Incant / rsttemp!PAQUETES)
      AdoDbf.Recordset!Exicaja = Val(rsttemp!Incant)
      AdoDbf.Recordset!ExiPza = Val(rsttemp!InCantPza)
      'desactivar las dos lineas siguientes cuando se verifique las pruebas
      'AdoDbf.Recordset!Minimo = Val(rsttemp!Minimo)
      'AdoDbf.Recordset!Minimo = Val(rsttemp!Maximo)
      AdoDbf.Recordset.Update
      rsttemp.MoveNext
      nReg = nReg + 1
   Wend
   AdoDbf.Recordset.Close
   Stb1.Panels(1).Text = cMenAnt
   MsgBox "Proceso Terminado...", vbInformation
   Exit Sub
ERROR:
  If Err.Number <> 32755 Then  'Numero de error al presionar el boton cancelar
     MsgBox Err.Description
  End If
   Stb1.Panels(1).Text = cMenAnt
End Sub

Private Sub CmdRefresh_Click()
Dim MenAnt
MenAnt = Stb1.Panels(1).Text
Stb1.SimpleText = Space(55) & "Espere un momento actualizando inventario..."
Stb1.Refresh
AdoModInv.Refresh
Stb1.Panels(1).Text = MenAnt
Stb1.Refresh
End Sub

Private Sub Command1_Click()

End Sub

Private Sub dbgrdModInv_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
  nCantAnt = AdoModInv.Recordset!Incant
End Sub

Private Sub dbgrdModInv_KeyDown(KeyCode As Integer, Shift As Integer)
Dim clave
Dim rs As ADODB.Recordset
'On Error GoTo ERROR:
clave = frmModInv.AdoModInv.Recordset!Inprod
If KeyCode = 118 Then ' ver los totales del inventario
    cad = " select sum(incant) as cajas, sum(incantpza) as piezas, count(consec) as variedad, sum(precosto * incant) as costocaj , sum((precosto / paquetes) * incantpza) as costopza , sum(precio1 * paquetes * incant) as precaja , sum(precio1 * incantpza ) as prepza from tfproduc, preprod, inventario where consec = preclave and consec = inprod and (incant >  0 or incantpza > 0 )"
    Set rs = New ADODB.Recordset
    rs.Open cad, cn, adOpenKeyset, adLockOptimistic, adCmdText
    cadena = ""
    If Not rs.EOF Then
       cadena = " " & vbCrLf
       cadena = cadena & " INVENTARIO DE BODEGA" & vbCrLf & vbCrLf
       cadena = cadena & " VARIEDADES  : " & vbTab & vbTab & rs!VARIEDAD & vbCrLf & vbCrLf
       cadena = cadena & " CAJAS       : " & vbTab & vbTab & rs!cajas & vbCrLf & vbCrLf
       cadena = cadena & " PIEZAS      : " & vbTab & vbTab & rs!Piezas & vbCrLf & vbCrLf
       cadena = cadena & " COSTO       : " & vbTab & vbTab & Format(rs!costocaj + rs!COSToPZA, "$###,##0.00") & vbCrLf & vbCrLf
       cadena = cadena & " PRECIO VENTA: " & vbTab & vbTab & Format(rs!PRECAJA + rs!prepza, "$###,##0.00") & vbCrLf & vbCrLf
       cadena = cadena & vbCrLf & vbCrLf
    End If
    MsgBox cadena, vbInformation, "TOTALES"
    rs.Close
End If

If KeyCode = 115 Then      'Tecla de funcion F4
   cn.Execute "DELETE FROM HistEnt"
   'Pedidos por proveedor
   cn.Execute "INSERT INTO histent (folio, fechaelab, cantsol, cantrec, fechaconf,fecharec,facturas,importe,tipo,existencia) SELECT pp_pedido, pp_fechagen, dg_cantsol, dg_cantreal + dg_promocionr , pp_fecconfirma, pp_fecrecibe, factura1, impfac1, tipo = 'PEDPROVE', dg_existencia FROM Pedprove, detalleglobal, Notaentrada WHERE dg_pedido = pp_pedido AND pedido  = pp_pedido AND pp_recibe = 1 AND DG_CANTREAL > 0 AND DG_PRODUCTO = '" & clave & "'"
   'Pedidos sugeridos instantaneos
   cn.Execute "INSERT INTO histent (folio, fechaelab, cantsol, cantrec, fechaconf,fecharec,facturas,importe,tipo,existencia) SELECT p_pedido, P_FECPED, df_cantsol, df_cantreal, p_fecconfirma, p_fecentreal, factura1, impfac1, tipo = 'PEDINST',df_existencia FROM Pedidos, detalleFactura, notaentrada WHERE df_pedido = p_pedido AND pedido = p_pedido AND p_recibido = 1 AND df_CANTREAL > 0 AND df_prod = '" & clave & "'"
   'Recibo de mercancia por traslados (Devoluciones, envasadora)
   cn.Execute "INSERT INTO histent (folio, fechaelab, cantsol, cantrec, fechaconf,fecharec,facturas,importe,tipo) SELECT t_clave, null, 0, dt_cantidad, null, t_fecha, t_foliotie, t_costo, tipo = 'TRASLREC' FROM Traslados, DetalleTraslado WHERE t_clave = dt_clave AND t_motivocancela is null AND Dt_cantidad > 0 AND Dt_producto = '" & clave & "' AND t_entrada = 1 AND t_enviado = 1"
   'Ajuste de inventario (Incrementado o disminuido)
   cn.Execute "INSERT INTO histent (folio, fecharec, cantrec,tipo) SELECT a_clave, a_fecha, da_cantidad, tipo = 'AJUSTE' FROM Ajustes, DetalleAjustes WHERE a_clave = da_clave AND Da_producto = '" & clave & "'"
   'Entrada de mercancia a traves de BackOrder
   cn.Execute "INSERT INTO histent (folio, fecharec, cantsol,cantrec,tipo) SELECT pedidog, fecha, cantasurtir,cantrecibida, tipo = 'BACKORDER' FROM detalleback WHERE producto = '" & clave & "' AND cantrecibida > 0"

   frmHisto.AdoEntradas.ConnectionString = cCadConex
   frmHisto.AdoEntradas.CommandType = adCmdText
   frmHisto.AdoEntradas.RecordSource = "SELECT * FROM histent ORDER BY fecharec DESC"
   frmHisto.AdoEntradas.Refresh

   Set rs = New ADODB.Recordset
   rs.Open "SELECT SUM(Cantrec) AS CAJREC FROM histent", cn, adOpenKeyset, adLockOptimistic, adCmdText
   frmHisto.txtEntrada.Text = IIf(IsNull(rs!Cajrec), 0, rs!Cajrec)
   rs.Close

   frmHisto.AdoEnvios.ConnectionString = cCadConex
   frmHisto.AdoEnvios.CommandType = adCmdText
   frmHisto.AdoEnvios.RecordSource = "SELECT * FROM Traslados, DetalleTraslado, Cattienda WHERE t_clave = dt_clave AND T_sucursalReceptor = ticlave AND t_motivocancela is null AND (Dt_cantidad > 0 or Dt_cantidadP > 0) AND Dt_producto = '" & frmModInv.AdoModInv.Recordset!Inprod & "' AND t_entrada = 0 AND T_ENVIADO = 1 ORDER BY t_fecha DESC"
   frmHisto.AdoEnvios.Refresh
   rs.Open "SELECT SUM(dt_cantidad) AS CAJSAL FROM Traslados, DetalleTraslado, Cattienda WHERE t_clave = dt_clave AND T_sucursalReceptor = ticlave AND t_motivocancela is null AND Dt_cantidad > 0 AND Dt_producto = '" & frmModInv.AdoModInv.Recordset!Inprod & "' AND t_entrada = 0 AND T_ENVIADO = 1", cn, adOpenKeyset, adLockOptimistic, adCmdText
   frmHisto.txtsalida.Text = IIf(IsNull(rs!Cajsal), 0, rs!Cajsal)
   frmHisto.Show 1

ElseIf KeyCode = 116 Then  'Tecla de funcion F5
   frmHisto.AdoEntPend.ConnectionString = cCadConex
   frmHisto.AdoEntPend.CommandType = adCmdText
   frmHisto.AdoEntPend.RecordSource = "SELECT * FROM Pedprove, detalleglobal WHERE dg_pedido = pp_pedido AND pp_recibe = 0 AND DG_PRODUCTO = '" & frmModInv.AdoModInv.Recordset!Inprod & "' ORDER BY pp_fecConfirma DESC"
   frmHisto.AdoEntPend.Refresh
   frmHisto.AdoEntSur.ConnectionString = cCadConex
   frmHisto.AdoEntSur.CommandType = adCmdText
   frmHisto.AdoEntSur.RecordSource = "SELECT * FROM Pedprove, detalleglobal WHERE dg_pedido = pp_pedido AND pp_recibe = 1 AND DG_PRODUCTO = '" & frmModInv.AdoModInv.Recordset!Inprod & "' AND dg_cantsol <> dg_cantreal ORDER BY pp_fecrecibe DESC"
   frmHisto.AdoEntSur.Refresh
   frmHisto.DatGSal.Visible = False
   frmHisto.DatgEnt.Visible = False
   frmHisto.FraPend.Visible = True
   frmHisto.Show 1
ElseIf KeyCode = 123 Then  'Tecla de funcion F5
   Set rs = New ADODB.Recordset
   'rs.Open "SELECT SUM(InCant * Precosto) as ImpInv, COUNT(consec) AS NumPro, SUM(InCant) NumCaj FROM Inventario,Tfproduc WHERE Inprod = consec AND inCant > 0 ", cn, adOpenKeyset, adLockOptimistic, adCmdText
   If CLAVEINVENTARIO = 10 Or CLAVEINVENTARIO = 3 Then
        rs.Open "SELECT SUM(InCant * Precosto) as ImpInv, COUNT(consec) AS NumPro, SUM(InCant) NumCaj FROM Inventario ,Tfproduc WHERE Inprod = consec AND inCant > 0 ", cn, adOpenKeyset, adLockOptimistic, adCmdText
   Else
       rs.Open "SELECT SUM(InCant * Precosto) as ImpInv, COUNT(consec) AS NumPro, SUM(InCant) NumCaj FROM Inventario" & Trim(CLAVEINVENTARIO) & " as inventario ,Tfproduc WHERE Inprod = consec AND inCant > 0 ", cn, adOpenKeyset, adLockOptimistic, adCmdText
   End If
   MsgBox "INFORMACION DEL INVENTARIO " & Chr(13) & Chr(13) & "PRODUCTOS   :   " & Format(rs!Numpro, "###,###,###") & Chr(13) & "CAJAS               :   " & Format(rs!NumCaj, "###,###,###.00") & Chr(13) & "IMPORTE          :   " & Format(rs!ImpInv, "$###,###,###,###.00"), vbInformation
ElseIf KeyCode = 117 Then ' CAMBIO DEL INVENTARIO BASE
   Me.fratda.Enabled = True
   fratda.Visible = True
   Me.cmbtiendas.SetFocus
   
End If
Exit Sub
ERROR:
    MsgBox Err.Description
End Sub

Private Sub dbgrdModInv_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
 If dbgrdModInv.SelBookmarks.Count > 0 Then dbgrdModInv.SelBookmarks.Remove 0
 dbgrdModInv.SelBookmarks.Add dbgrdModInv.RowBookmark(Me.dbgrdModInv.Row)
End Sub

Private Sub Form_Activate()
 Unload frmAreaRecibo
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 119 Then frmCalc.Show   'F8
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
   Unload Me
End If
End Sub

Private Sub Form_Load()
'SE CARGA EL CATALOGO DE TIENDAS
Set rsttempt = New ADODB.Recordset
rsttempt.Open "SELECT TICLAVE,TIDESCRIP FROM CATTIENDA order by tidescrip ", cn, adOpenStatic, adLockOptimistic, adCmdText
While Not rsttempt.EOF
    If Val(rsttempt!Ticlave) < 10 Then
       cmbtiendas.AddItem "[0" & Trim(rsttempt!Ticlave) & "]" & "  " & Trim(rsttempt!TIDESCRIP) & "  "
    Else
       cmbtiendas.AddItem "[" & Trim(rsttempt!Ticlave) & "]" & "  " & Trim(rsttempt!TIDESCRIP) & "  "
    End If
    rsttempt.MoveNext
Wend
rsttempt.Close

CLAVEINVENTARIO = Mid(cSucursal, 1, 3)
CLAVEINVENTARIO = 10
Me.Caption = "INVENTARIO BODEGA    --" & cSucursal

 lDatAJu = False: lCon = False
 AdoModInv.CursorType = adOpenKeyset
 AdoModInv.ConnectionString = cCadConex
 If CLAVEINVENTARIO = 10 Or CLAVEINVENTARIO = 3 Or CLAVEINVENTARIO = 3 Then
    cad = "SELECT TFPRODUC.activo, TFPRODUC.BarrasPza, INVENTARIO.inprod, TFPRODUC.Descripc,  LTRIM(STR(TFPRODUC.paquetes)) + ' X ' + lTrim(str(TFPRODUC.contenid,10,3)) + space(2) + TFPRODUC.medida  AS MEDIDA, INVENTARIO.InInicial, INVENTARIO.incant, INVENTARIO.minimo, INVENTARIO.maximo,INVENTARIO.InCantPza, incantcdc, incantpzacdc  FROM INVENTARIO , TFPRODUC WHERE INVENTARIO.inprod = TFPRODUC.consec AND TFPRODUC.Paquetes > 0 ORDER BY DESCRIPC,CONTENID"
 Else
    cad = "SELECT TFPRODUC.activo, TFPRODUC.BarrasPza, INVENTARIO.inprod, TFPRODUC.Descripc,  LTRIM(STR(TFPRODUC.paquetes)) + ' X ' + lTrim(str(TFPRODUC.contenid,10,3)) + space(2) + TFPRODUC.medida  AS MEDIDA, INVENTARIO.InInicial, INVENTARIO.incant, INVENTARIO.minimo, INVENTARIO.maximo,INVENTARIO.InCantPza, incantcdc, incantpzacdc  FROM INVENTARIO" & Trim(CLAVEINVENTARIO) & "  as inventario , TFPRODUC WHERE INVENTARIO.inprod = TFPRODUC.consec AND TFPRODUC.Paquetes > 0 ORDER BY DESCRIPC,CONTENID"
 End If
 
 AdoModInv.RecordSource = cad
 'AdoModInv.RecordSource = "SELECT TFPRODUC.activo, TFPRODUC.BarrasPza, INVENTARIO.inprod, TFPRODUC.Descripc,  LTRIM(STR(TFPRODUC.paquetes)) + ' X ' + lTrim(str(TFPRODUC.contenid,10,3)) + space(2) + TFPRODUC.medida  AS MEDIDA, INVENTARIO.InInicial, INVENTARIO.incant, INVENTARIO.minimo, INVENTARIO.maximo,INVENTARIO.InCantPza, incantcdc, incantpzacdc  FROM INVENTARIO , TFPRODUC WHERE INVENTARIO.inprod = TFPRODUC.consec AND TFPRODUC.Paquetes > 0 ORDER BY DESCRIPC,CONTENID"
 AdoModInv.Refresh
 lblInfo.Caption = "Num. de productos " + Str(AdoModInv.Recordset.RecordCount)
End Sub

Private Sub CmdActualizar_Click()
If lCon = False Then
  fraCon.Visible = True
  txtContra.Text = ""
  txtContra.SetFocus
End If
End Sub

Private Sub cmdAnterior_Click()
AdoModInv.Recordset.MovePrevious
If AdoModInv.Recordset.BOF Then AdoModInv.Recordset.MoveFirst
End Sub

Private Sub cmdPrimero_Click()
  AdoModInv.Recordset.MoveFirst
End Sub

Private Sub cmdRegresar_Click()
  Unload Me
End Sub

Private Sub cmdSiguiente_Click()
AdoModInv.Recordset.MoveNext
If AdoModInv.Recordset.EOF Then AdoModInv.Recordset.MoveLast
End Sub

Private Sub cmdUltimo_Click()
AdoModInv.Recordset.MoveLast
End Sub

Private Sub dbgrdModInv_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
On Error GoTo ERROR:
nCantAnt = OldValue
If Not lDatAJu Then frmModInvAjus.Show 1   'Si NO capturaron los datos correspondientes al ajuste
If Not lDatAJu Then
   MsgBox "Al no guardar los datos del ajuste no puede modificar el inventario", vbExclamation
   Cancel = True
Else
   If UCase(dbgrdModInv.Columns(ColIndex).DataField) = "INCANT" Then   'Columna correspondiente a cantidad por caja
      CmdActualizar.Enabled = True
      AdoDetAju.Recordset.AddNew
      AdoDetAju.Recordset!da_clave = nOp
      AdoDetAju.Recordset!da_producto = AdoModInv.Recordset!Inprod
      AdoDetAju.Recordset!da_cantidadAnt = nCantAnt
      AdoDetAju.Recordset!da_cantidad = dbgrdModInv.Columns(ColIndex).Text - AdoModInv.Recordset!Incant
      AdoDetAju.Recordset.Update
   ElseIf UCase(dbgrdModInv.Columns(ColIndex).DataField) = "INCANTPZA" Then   'Columna correspondiente a existencia por pieza
      CmdActualizar.Enabled = True
      AdoDetAju.Recordset.AddNew
      AdoDetAju.Recordset!da_clave = nOp
      AdoDetAju.Recordset!da_producto = AdoModInv.Recordset!Inprod
      AdoDetAju.Recordset!da_cantidadp = dbgrdModInv.Columns(ColIndex).Text - AdoModInv.Recordset!InCantPza
      AdoDetAju.Recordset.Update
   End If
   SendKeys "{DOWN}"
End If
Exit Sub
ERROR:
  MsgBox Err.Description
End Sub

Private Sub dbgrdModInv_HeadClick(ByVal ColIndex As Integer)
Dim cCampo As String
Select Case UCase(dbgrdModInv.Columns(ColIndex).DataField)
    'Case 1, 0, 2
    Case "INPROD", "DESCRIPC", "BARRASPZA", "INCANT", "INCANTPZA"
      Stb1.Panels(1).Text = Space(30) & "Espere un momento ordenando datos por la columna  " & Trim(dbgrdModInv.Columns(ColIndex).Caption)
      'AdoModInv.RecordSource = "SELECT TFPRODUC.activo, TFPRODUC.BarrasPza, INVENTARIO.inprod, TFPRODUC.Descripc,  LTRIM(STR(TFPRODUC.paquetes)) + ' X ' + lTrim(str(TFPRODUC.contenid,10,3)) + space(2) + TFPRODUC.medida  AS MEDIDA, INVENTARIO.InInicial, INVENTARIO.incant, INVENTARIO.minimo, INVENTARIO.maximo,INVENTARIO.InCantPza, incantcdc, incantpzacdc  FROM INVENTARIO, TFPRODUC WHERE INVENTARIO.inprod = TFPRODUC.consec AND TFPRODUC.Paquetes > 0 " & _
      '                         "ORDER BY " & Trim(dbgrdModInv.Columns(ColIndex).DataField)
      If CLAVEINVENTARIO = 10 Or CLAVEINVENTARIO = 3 Then
        AdoModInv.RecordSource = "SELECT TFPRODUC.activo, TFPRODUC.BarrasPza, INVENTARIO.inprod, TFPRODUC.Descripc,  LTRIM(STR(TFPRODUC.paquetes)) + ' X ' + lTrim(str(TFPRODUC.contenid,10,3)) + space(2) + TFPRODUC.medida  AS MEDIDA, INVENTARIO.InInicial, INVENTARIO.incant, INVENTARIO.minimo, INVENTARIO.maximo,INVENTARIO.InCantPza, incantcdc, incantpzacdc  FROM INVENTARIO  , TFPRODUC WHERE INVENTARIO.inprod = TFPRODUC.consec AND TFPRODUC.Paquetes > 0 " & _
                               "ORDER BY " & Trim(dbgrdModInv.Columns(ColIndex).DataField)
      Else
       AdoModInv.RecordSource = "SELECT TFPRODUC.activo, TFPRODUC.BarrasPza, INVENTARIO.inprod, TFPRODUC.Descripc,  LTRIM(STR(TFPRODUC.paquetes)) + ' X ' + lTrim(str(TFPRODUC.contenid,10,3)) + space(2) + TFPRODUC.medida  AS MEDIDA, INVENTARIO.InInicial, INVENTARIO.incant, INVENTARIO.minimo, INVENTARIO.maximo,INVENTARIO.InCantPza, incantcdc, incantpzacdc  FROM INVENTARIO" & Trim(CLAVEINVENTARIO) & " AS INVENTARIO , TFPRODUC WHERE INVENTARIO.inprod = TFPRODUC.consec AND TFPRODUC.Paquetes > 0 " & _
                               "ORDER BY " & Trim(dbgrdModInv.Columns(ColIndex).DataField)
      End If
     AdoModInv.Refresh
     Stb1.Panels(1).Text = Space(70) & "Datos ordenados por la columna " & Trim(dbgrdModInv.Columns(ColIndex).Caption)
     SendKeys "{ENTER}"
End Select
End Sub

Private Sub Form_Resize()
On Error GoTo ERROR:
 dbgrdModInv.Width = frmModInv.ScaleWidth - 400
 dbgrdModInv.Height = frmModInv.ScaleHeight - 1100
ERROR:
End Sub

Private Sub Form_Unload(Cancel As Integer)
If nOp <> 31 Then frmAreaRecibo.Show
End Sub

Private Sub Option2_Click()

End Sub

Private Sub txtContra_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    cmdConAceptar_Click
 ElseIf KeyAscii = 27 Then
    cmdConCance_Click
 End If
End Sub
